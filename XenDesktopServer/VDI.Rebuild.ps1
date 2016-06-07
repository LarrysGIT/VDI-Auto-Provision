
<#
    Version 1.00 [Author: Larry Song; Time: 2014-06-26]
        First build, bases on "VDI.Build.ps1"
    Version 1.01 [Author: Larry Song; Time: 2014-07-04]
        Fix bug, when writing instance logs with wrong file name.
        Fix bug, move VM shutdown before remove desktop from citrix
    Version 1.02 [Author: Larry Song; Time: 2014-11-05]
        Add $Enable variable for the script
    Version 1.03 [Author: Larry Song; Time: 2014-12-30]
        Change property 'VDIName' to 'VDI Name' due to sharepoint script reconstructure
    Version 1.04 [Author: Larry Song; Time: 2015-01-07]
        Update codes for failed jobs
#>

Set-Location (Get-Item ($MyInvocation.MyCommand.Definition)).DirectoryName
. '.\_Configuration.ps1'
Define-CommonVariables
Define-VDIBuildVariables
Define-VDIRebuildVariables

. '.\VDI.Common.Function.ps1'

$ExitCode = 0

function Quit{
    PARAM(
        [int]$ExitCode,
        [int[]]$AdditionalJobs
    )
    switch($AdditionalJobs){
        1 {
            @($objRebuildJob) | Export-Csv -Path $VMReportLeftFile -NoTypeInformation -Delimiter "`t"
        }
        default{}
    }
    exit($ExitCode)
}

New-Item -Path "$LocalDes\$strDate" -ItemType 'Directory' -Confirm:$false -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
if(!(Test-Path -Path "$LocalDes\$strDate" -PathType Container)){
    Add-Log -Path $strLogFile -Value "Local folder creation failed, cause:" -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0001 # 0000 0000 0000 0001
    Add-Log -Path $strLogFile -Value "Script exit with code $ExitCode"
    Quit -ExitCode $ExitCode -AdditionalJobs 1
}

Add-Log -Path $strLogFile -Value 'VDI rebuild script start'
if(!$Enable)
{
    Add-Log -Path $strLogFile -Value 'VDI rebuild script disabled'
    Quit -ExitCode $ExitCode
}

Add-Log -Path $strLogFile -Value "Check tag file $RawFile"

while($true){
    if(Test-Path -Path $RawFile){
        break
    }
    Add-Log -Path $strLogFile -Value "Tag file not found, sleep 5 mins."
    Start-Sleep -Seconds 300
}

$RawData = Import-Csv -Path $RawFile -ErrorAction:SilentlyContinue
if(!$?){
    Add-Log -Path $strLogFile -Value 'Read raw data file failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0002 # 0000 0000 0000 0010
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    Quit -ExitCode $ExitCode
}
$RawData = @($RawData)

Add-Log -Path $strLogFile -Value "Read raw data succeed, count: $($RawData.Count)"
if($RawData.Count -eq 1){
    Add-Log -Path $strLogFile -Value "Only 1 item imported, check whether it's a blank."
    if($RawData[0].'VDI Name'){
        Add-Log -Path $strLogFile -Value 'The item is not blank, move on.'
        Add-Log -Path $strLogFile -Value 'VDI Rebuild request will use existing templates in pool.'
    }else{
        Add-Log -Path $strLogFile -Value "The item is blank, no need to process. exit with code $ExitCode"
        Quit -ExitCode $ExitCode
    }
}

Add-Log -Path $strLogFile -Value 'Start adding VMware snapin.'
Add-PSSnapin VMware* -ErrorAction:SilentlyContinue
if(!$?){
    Add-Log -Path $strLogFile -Value 'Add VMware snapin failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    Quit -ExitCode $ExitCode -AdditionalJobs 1
}
Add-Log -Path $strLogFile -Value 'Add VMware snapin succeed.'

Add-Log -Path $strLogFile -Value 'Start adding citrix snappin.'
Add-PSSnapin *Citrix* -ErrorAction:SilentlyContinue
if(!$?){
    Add-Log -Path $strLogFile -Value 'Add Citrix snapin failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    Quit -ExitCode $ExitCode -AdditionalJobs 1
}
Add-Log -Path $strLogFile -Value 'Add Citrix snapin succeed.'
<#
Add-Log -Path $strLogFile -Value 'Start importing ActiveDirectory module.'

Import-Module -Name 'ActiveDirectory' -ErrorAction:SilentlyContinue
if(!$?){
    Add-Log -Path $strLogFile -Value 'Import ActiveDirectory module failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0010 # 0000 0000 0001 0000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    exit($ExitCode)
}

Add-Log -Path $strLogFile -Value 'Import ActiveDirectory module succeed.'
#>

$ClearVIServer = $false
if($DefaultVIServer -and $vCenter.Contains(($DefaultVIServer.Name).ToUpper())){
    Add-Log -Path $LogFile -Value "This script already connect to VI server, will not connect again."
}else{
    Connect-VIServer -Server $vCenter
    if(!$?){
        Add-Log -Path $strLogFile -Value "Connect to vCenter $vCenter failed, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    }
    else
    {
        $ClearVIServer = $true
    }
}

$objRebuildJob = New-Object PSObject -Property @{ID = $null; POD = $POD; VDIUsers = $null; 'VDI Name' = $null; VDINameTmp = $null; Exception = $null; CreatedBy = $null; VDIDesktopGroup = $null; VDICatalog = $null; VMCluster = $null; VMFolder = $null; VMTemplateFolder = $null; VMCPU = $null; VMMEM = $null; VMNetWorks = $null; VMUUID = $null}
$JobsLeft = @()
$Jobs = @()

foreach($Item in $RawData)
{
    $VDIName = $Item.'VDI Name'.ToUpper()
    Add-Log -Path $strLogFile -Value "start processing: [$VDIName]"
    if($VDIName -notmatch '^\s*$')
    {
        if($VDIName -notmatch "^$Prefix")
        {
            Add-Log -Path $strLogFile -Value "VDI name not match RegExp: [^$Prefix]"
            <#
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'VDI name invalid'
            #>
            continue
        }
        Add-Log -Path $strLogFile -Value 'Collecting necessary information.'
        Add-Log -Path $strLogFile -Value 'Get VM from vCenter.'
        $objVM = $null
        $objVM = @(Get-VM -Name $VDIName -ErrorAction:SilentlyContinue)
        if(!$?)
        {
            Add-Log -Path $strLogFile -Value 'Unable to get VM from vCenter, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            Add-Log -Path $strLogFile -Value 'Skipped, continue script.'
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'Failed capture in vCenter'
            continue
        }
        if($objVM.Count -ge 2)
        {
            Add-Log -Path $strLogFile -Value "Capture number of VM greater or equal than 2: [$($objVM.Count)] - [$(($objVM | %{$_.Name}) -join '], [')]"
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'Multiple VDIs matched'
            continue
        }
        $objVM = $objVM[0]
        Add-Log -Path $strLogFile -Value 'Properties of VMware,'
        Add-Log -Path $strLogFile -Value "VM CPU: [$($objVM.NumCpu)]"
        Add-Log -Path $strLogFile -Value "VM Mem: [$($objVM.MemoryMB)]"
        Add-Log -Path $strLogFile -Value "VM Folder: [$($objVM.Folder)]; Folder ID: [$($objVM.FolderId)]"
        #Add-Log -Path $strLogFile -Value "VM Host: [$($objVM.VMHost)]; VMHost ID: [$($objVM.VMHostId)]"
        Add-Log -Path $strLogFile -Value "VM Cluster: [$($objVM.VMHost.Parent.Name)]"
        Add-Log -Path $strLogFile -Value "VM Network: [$(($objVM | Get-NetworkAdapter | %{$_.NetworkName}) -join '], [')]"
        Add-Log -Path $strLogFile -Value "VM UUID: [$($objVM.ExtensionData.Config.Uuid)]"
        Add-Log -Path $strLogFile -Value "VM Notes: [$($objVM.Notes)]"
        Add-Log -Path $strLogFile -Value "VM Guest Name: [$($objVM.ExtensionData.Guest.HostName)]"
        Add-Log -Path $strLogFile -Value "VM Guest IP: [$($objVM.ExtensionData.Guest.IpAddress -join '], [')]"

        Add-Log -Path $strLogFile -Value 'Get VM from XenDesktop.'
        $objVDI = $null
        $objVDI = Get-BrokerDesktop -HostedMachineId $objVM.ExtensionData.Config.Uuid -ErrorAction:SilentlyContinue
        if(!$objVDI)
        {
            Add-Log -Path $strLogFile -Value 'Unable to get VDI from XenDesktop by VM UUID, VDI object is null.' -Type Warning
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'Failed capture in Citrix'
            continue
        }
        Add-Log -Path $strLogFile -Value 'Properties of Citrix,'
        Add-Log -Path $strLogFile -Value "VDI Associated Users: [$($objVDI.AssociatedUserNames -join '], [')]"
        if(!$objVDI.AssociatedUserNames)
        {
            Add-Log -Path $strLogFile -Value 'No user associated to this VDI.' -Type Warning
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'No associated user'
            continue
        }
        Add-Log -Path $strLogFile -Value "VDI Catalog Name/UID: [$($objVDI.CatalogName)/$($objVDI.CatalogUid)]"
        Add-Log -Path $strLogFile -Value "VDI Desktop group Name/UID: [$($objVDI.DesktopGroupName)/$($objVDI.DesktopGroupUid)]"
        if(!$objVDI.DesktopGroupName)
        {
            Add-Log -Path $strLogFile -Value 'No desktop group to this VDI.' -Type Warning
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'No desktop group to this VDI'
            continue
        }
        Add-Log -Path $strLogFile -Value "VDI Hypervisor Name/ID: [$($objVDI.HypervisorConnectionName)/$($objVDI.HypervisorConnectionUid)]"
        Add-Log -Path $strLogFile -Value "VDI VC Machine Name: [$($objVDI.HostedMachineName)]"
        Add-Log -Path $strLogFile -Value "VDI Machine Name/UID: [$($objVDI.MachineName)/$($objVDI.MachineUid)]"
        Add-Log -Path $strLogFile -Value "VDI IPAddress: [$($objVDI.IPAddress)]"

        if($objVM.Name -ne $objVDI.HostedMachineName)
        {
            Add-Log -Path $strLogFile -Value 'VDI name in vCenter and Citrix not matched!' -Type Error
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'VDI name in VC and XD not matched'
            continue
        }

        if($objVDI.MachineName -notmatch "\\$VDIName$")
        {
            Add-Log -Path $strLogFile -Value 'VDI name not matched AD SAMAccountName' -Type Error
            Add-Log -Path $strLogFile -Value 'For protection, this request skipped' -Type Warning
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].'VDI Name' = $VDIName
            $JobsLeft[-1].CreatedBy = $Item.'Created By'
            $JobsLeft[-1].Exception = 'VDI name in XD and AD not matched'
            continue
        }

        # Start VDI processing
        $Jobs += $objRebuildJob.PSObject.Copy()
        $Jobs[-1].VDIUsers = @($objVDI.AssociatedUserNames)
        $Jobs[-1].'VDI Name' = $VDIName
        $Jobs[-1].VDINameTmp = "${VDIName}_Rebuild_$(Get-Random)"
        $Jobs[-1].CreatedBy = $Item.'Created By'
        $Jobs[-1].VDICatalog = $objVDI.CatalogName
        $Jobs[-1].VDIDesktopGroup = $objVDI.DesktopGroupName
        $Jobs[-1].VMCluster = $objVM.VMHost.Parent.Name
        $Jobs[-1].VMFolder = $objVM.Folder.Name
        $Jobs[-1].VMTemplateFolder = $Clusters[$($objVM.VMHost.Parent.Name)][$Type]['POOL']
        $Jobs[-1].VMCPU = $objVM.NumCpu
        $Jobs[-1].VMMEM = [int]($objVM.MemoryMB/1024)
        $Jobs[-1].VMNetWorks = @($objVM.NetworkAdapters | %{$_.NetworkName})
        $Jobs[-1].VMUUID = $objVM.ExtensionData.Config.Uuid
    }
    else
    {
        Add-Log -Path $strLogFile -Value 'Blank VDI name detected, skipped.' -Type Warning
        $JobsLeft += $objRebuildJob.PSObject.Copy()
        $JobsLeft[-1].'VDI Name' = $VDIName
        $JobsLeft[-1].CreatedBy = $Item.'Created By'
        $JobsLeft[-1].Exception = 'Blank VDI Name'
        continue
    }
}

Add-Log -Path $strLogFile -Value "Jobs in stack count: $($Jobs.Count)"
Add-Log -Path $strLogFile -Value "Left jobs in stack count: $($JobsLeft.Count)"

Add-Log -Path $strLogFile -Value 'Start process jobs'
$PowershellInstances = @()
foreach($Item in $Jobs)
{
    Add-Log -Path $strLogFile -Value "Start processing: [$($Item.'VDI Name')]"
    Add-Log -Path $strLogFile -Value 'Set brokerdesktop to maintenance mode'
    Get-BrokerDesktop -HostedMachineId $Item.VMUUID | Set-BrokerPrivateDesktop -InMaintenanceMode $true
    if(!$?)
    {
        Add-Log -Path $strLogFile -Value 'Failed to set brokerdesktop to maintenance mode, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
    }
    Add-Log -Path $strLogFile -Value "Rename VM to: [$($Item.VDINameTmp)]"
    try
    {
        Set-VM -VM $Item.'VDI Name' -Name $Item.VDINameTmp -ErrorAction:SilentlyContinue -Confirm:$false
    }
    catch
    {
        Get-BrokerDesktop -HostedMachineId $Item.VMUUID | Set-BrokerPrivateDesktop -InMaintenanceMode $false
        Add-Log -Path $strLogFile -Value 'Failed to rename VDI, cause:' -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        $JobsLeft += $objRebuildJob.PSObject.Copy()
        $JobsLeft[-1].'VDI Name' = $VDIName
        $JobsLeft[-1].CreatedBy = $Item.'Created By'
        $JobsLeft[-1].Exception = 'Failed to rename VDI'
        continue
    }
    Add-Log -Path $strLogFile -Value "Shutdown VM: [$($Item.VDINameTmp)]"
    try
    {
        Shutdown-VMGuest -VM $Item.VDINameTmp -Confirm:$false
    }
    catch
    {
        Add-Log -Path $strLogFile -Value 'Failed to send shutdown signal, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
        Stop-VM -VM $Item.VDINameTmp -Confirm:$false
    }
    Start-Sleep -Seconds 60
    if((Get-VM -Name $Item.VDINameTmp).PowerState -ne 'PoweredOff')
    {
        Add-Log -Path $strLogFile -Value 'VM still on after 60 seconds, force poweroff.'
        try
        {
            Stop-VM -VM $Item.VDINameTmp -Confirm:$false
        }
        catch
        {
            Add-Log -Path $strLogFile -Value 'Failed to force poweroff VM, cause:' -Type Warning
            Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
        }
    }
    Add-Log -Path $strLogFile -Value "Remove VDI from DesktopGroup: [$($Item.VDIDesktopGroup)]"
    Get-BrokerMachine -HostedMachineId $Item.VMUUID | Remove-BrokerMachine -DesktopGroup $Item.VDIDesktopGroup
    if(!$?)
    {
        Add-Log -Path $strLogFile -Value 'Failed to remove VDI from desktopgroup, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
    }
    Add-Log -Path $strLogFile -Value 'Remove VDI from XenDesktop'
    Get-BrokerMachine -HostedMachineId $Item.VMUUID | Remove-BrokerMachine
    if(!$?)
    {
        Add-Log -Path $strLogFile -Value 'Failed to remove VDI from Citrix, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
    }
    Add-Log -Path $strLogFile -Value 'Set VM network adapter to disconnect'
    try
    {
        Get-NetworkAdapter -VM $Item.VDINameTmp | Set-NetworkAdapter -StartConnected $false -Confirm:$false
    }
    catch
    {
        Add-Log -Path $strLogFile -Value 'Failed to disconnect network, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
    }
    $PowershellInstances += Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @($(
        '-File', 'VDI.Build.Core.ps1',
        '-VIServer', "$vCenter",
        '-User', "$(if($Item.VDIUsers[0].Contains('\')){$Item.VDIUsers[0].Split('\')[1]}else{$Item.VDIUsers[0]})",
        '-Prefix', "$Prefix",
        '-DesktopGroup', "$($Item.VDIDesktopGroup)",
        '-DDCList', "`"$DDCList`"",
        '-SupportGroup', "$SupportGroup",
        '-Environment', "$Type",
        '-ClusterToGo', "$($Item.VMCluster)",
        '-Folder', "$($Item.VMFolder)",
        '-TemplateFolder', "$($Item.VMTemplateFolder)",
        '-VDIName', "$($Item.'VDI Name')", '-IgnoreVDINameAvailable',
        '-NumCPU', $Item.VMCPU,
        '-MemoryGB', $Item.VMMEM,
        '-NetworkPositive', $($Item.VMNetWorks[0].Split('_')[2]),
        '-ExitCodeWriteTo', "$LocalDes\$strDate\ExitCodes.Core.txt"
    ) | ?{$_}) -PassThru -RedirectStandardOutput "$strDate\Console.Rebuild_$($Item.'VDI Name').log" | Select-Object -Property Id, StartTime, @{N = "Args"; E = {(Get-WmiObject -Class Win32_Process -Filter "ProcessId = $($_.Id)").CommandLine}}
    if(!$?){
        Add-Log -Path $strLogFile -Value "New instance failed to start, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        $JobsLeft += $objRebuildJob.PSObject.Copy()
        $JobsLeft[-1].'VDI Name' = $VDIName
        $JobsLeft[-1].CreatedBy = $Item.'Created By'
        $JobsLeft[-1].Exception = 'Failed to launch VDI build process'
    }else{
        Add-Log -Path $strLogFile -Value 'New instance up.'
    }
    Add-Log -Path $strLogFile -Value 'Sleep 5 mins before process next.'
    Start-Sleep -Seconds 300
}

Add-Log -Path $strLogFile -Value 'Start to check pending jobs.'
$Threshold = 8
Add-Log -Path $strLogFile -Value "Jobs loop threshold: $Threshold"
$LoopCount = 0

while($true){
    $LoopCount++
    Add-Log -Path $strLogFile -Value "Current loop count: $LoopCount"
    Add-Log -Path $strLogFile -Value "Powershell instances count: $($PowershellInstances.Count)"
    $PowershellInstances = @($PowershellInstances | %{
        if(Get-Process -Id $_.Id -ErrorAction:SilentlyContinue){$_}else{
            Add-Log -Path $strLogFile -Value "Instance completed, ID: $($_.Id), Start time: $($_.StartTime), CommandLine: $($_.Args)"
            Add-Log -Path $strLogFile -Value "Try to capture exit code for process $($_.Id)."
            if($ExitCodeCore = @(Select-String -Path "$LocalDes\$strDate\ExitCodes.Core.txt" -Pattern "^$($_.Id)\t(\d+)$")){
                Add-Log -Path $strLogFile -Value "Exit code: $($ExitCodeCore[0].Matches[0].Groups[1].Value)."
                if([int]($ExitCodeCore[0].Matches[0].Groups[1].Value)){
                    Add-Log -Path $strLogFile -Value "Abnormal exit code for powershell instance." -Type Warning
                    $ExitCode = $ExitCode -bor 0x0080 # 0000 0000 1000 0000
                }
            }else{
                Add-Log -Path $strLogFile -Value "No exit code captured for process $($_.Id)." -Type Warning
                $ExitCode = $ExitCode -bor 0x0100 # 0000 0001 0000 0000
            }
        }
    })
    if(-not $PowershellInstances){
        Add-Log -Path $strLogFile -Value "All instances completed."
        break
    }
    if($LoopCount -gt $Threshold){
        Add-Log -Path $strLogFile -Value "Threshold hit, this is abnormal." -Type Warning
        $ExitCode = $ExitCode -bor 0x0400 # 0000 0100 0000 0000
        $PowershellInstances | %{
            Add-Log -Path $strLogFile -Value "Pending instance: ProcessId = $($_.Id), StartTime = $($_.StartTime), CommandLine: $($_.Args)" -Type Warning
        }
        break
    }
    Add-Log -Path $strLogFile -Value 'Sleep 15 mins to next loop.'
    Start-Sleep -Seconds (15*60)
}

Add-Log -Path $strLogFile -Value 'Start post processing'
Add-Log -Path $strLogFile -Value 'Collecting new VDI information and verification'
$JobsPost = @()
foreach($Item in $Jobs)
{
    Add-Log -Path $strLogFile -Value "Verify: [$($Item.'VDI Name')]"
    $JobsPost += $objRebuildJob.PSObject.Copy()
    $JobsPost[-1].'VDI Name' = $Item.'VDI Name'
    $JobsPost[-1].VMUUID = (Get-VM -Name $JobsPost[-1].'VDI Name').ExtensionData.Config.Uuid
    Add-Log -Path $strLogFile -Value "VM UUID: [$($JobsPost[-1].VMUUID)]"
    $JobsPost[-1].VDICatalog = (Get-BrokerMachine -HostedMachineId $JobsPost[-1].VMUUID).CatalogName
    Add-Log -Path $strLogFile -Value "VDI Catalog: [$($JobsPost[-1].VDICatalog)]"
    $JobsPost[-1].VDIDesktopGroup = (Get-BrokerDesktop -HostedMachineId $JobsPost[-1].VMUUID).DesktopGroupName
    Add-Log -Path $strLogFile -Value "VDI DesktopGroup: [$($JobsPost[-1].VDIDesktopGroup)]"
    if($JobsPost[-1].VMUUID)
    {
        Add-Log -Path $strLogFile -Value 'VM UUID has value'
        if($JobsPost[-1].VMUUID -eq $Item.VMUUID)
        {
            Add-Log -Path $strLogFile -Value 'New VDI has the same uuid like the old, this should not happen' -Type Error
            $JobsLeft += $objRebuildJob.PSObject.Copy()
            $JobsLeft[-1].CreatedBy = $Item.CreatedBy
            $JobsLeft[-1].Exception = 'VDI rebuild, but captured the same uuid'
            continue
        }
    }
    else
    {
        Add-Log -Path $strLogFile -Value 'Failed to capture uuid of new VDI' -Type Error
        $JobsLeft += $objRebuildJob.PSObject.Copy()
        $JobsLeft[-1].CreatedBy = $Item.CreatedBy
        $JobsLeft[-1].Exception = 'VDI rebuild, but failed to get uuid'
        continue
    }
    if($JobsPost[-1].VDIDesktopGroup)
    {
        Add-Log -Path $strLogFile -Value 'DesktopGroup has value'
        if($Item.VDIUsers.Count -gt 1)
        {
            $Item.VDIUsers[1..$($Item.VDIUsers.Count - 1)] | %{
                Add-BrokerUser -Name $_ -Machine (Get-BrokerMachine -HostedMachineId $JobsPost[-1].VMUUID)
            }
        }
    }
    else
    {
        Add-Log -Path $strLogFile -Value 'DesktopGroup is blank, failed to add to Citrix?' -Type Error
        $JobsLeft += $objRebuildJob.PSObject.Copy()
        $JobsLeft[-1].CreatedBy = $Item.CreatedBy
        $JobsLeft[-1].Exception = 'VDI rebuild, but failed to get desktopgroup'
        continue
    }
    ## Seems rebuild succeed.
    Add-Log -Path $strLogFile -Value 'VDI rebuild succeed, remove the old VDI'
    $objVMTMP = $null
    $objVMTMP = @(Get-VM -Name $Item.VDINameTmp | ?{$_.Name -imatch "^${Prefix}[a-z]\d{6}_Rebuild_\d+$" -and $_.ExtensionData.Config.Uuid -eq $Item.VMUUID})
    if($objVMTMP.Count -eq 1)
    {
        Add-Log -Path $strLogFile -Value "Successfully captured old VM: [$($objVMTMP[0].Name)]"
        try
        {
            # Remove-VM -VM $objVMTMP[0] -DeletePermanently -Confirm:$false
        }
        catch
        {
            Add-Log -Path $strLogFile -Value 'Failed to remove old VDI, cause:' -Type Warning
            Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
        }
    }
    else
    {
        Add-Log -Path $strLogFile -Value "Incorrect number of old VDI captured: [$(($objVMTMP | %{$_.Name}) -join '], [')]" -Type Warning
    }
}

if(!$JobsLeft.Count)
{
    $JobsLeft += $objRebuildJob.PSObject.Copy()
}
$JobsLeft | Export-Csv -Path $VMReportLeftFile -NoTypeInformation -Delimiter "`t"

if(!$JobsPost)
{
    $JobsPost += $objRebuildJob.PSObject.Copy()
}
$JobsPost | Export-Csv -Path $VMReportProcessedFile -NoTypeInformation -Delimiter "`t"

if($ClearVIServer){
    Disconnect-VIServer -Server $vCenter -Confirm:$false -Force
    Add-Log -Path $strLogFile -Value 'Disconnected from vCenter.'
}

Remove-PSSnapin VMware* -Confirm:$false
Remove-PSSnapin *Citrix* -Confirm:$false

Add-Log -Path $strLogFile -Value "Script completed, exit with code $ExitCode."
Quit -ExitCode $ExitCode
