
<#
    Version 1.00 [Author: Matthew Pound; Time: 2015-02-02]
        Matt created the space reclaim script.
    Version 1.01 [Author: Larry Song; Time: 2015-02-09]
        Update script to combine with SharePoint scripts for A environment use
    Version 1.02 [Author: Larry Song; Time: 2015-04-02]
        Add VM before and after committed size
    Version 1.03 [Author: Larry Song; Time: 2015-06-27]
        Add MoveBack switch
    Version 1.04 [Author: Larry Song; Time: 2015-07-02]
        Use random datastore
    Version 1.05 [Author: Larry Song; Time: 2015-07-06]
        Add a variable $ReclaimCount to limit unnecessary vMotion
    Version 1.06 [Author: Larry Song; Time: 2015-07-07]
        Add a variable $StillMoveWhensDeleteFailed
        Use inequality vMontion instead of equality
        Add a new DeployDate attribute to VM
            previous old VDIs of course will have null value for the attribute,
            so, a date '2015-01-01' will be assigned to all old VDIs as the value
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means connect with vCenter failed.
    second   bit set 1, means failed to copy sdelete files
    third    bit set 1, means failed to trigger sdelete script
    fourth   bit set 1, means failed to move VDI to another datastore
    fifth    bit set 1, means failed to move VDI back to original datastore
    sixth    bit set 1, means 
    seventh  bit set 1, means 
    eigthth  bit set 1, means 
    ninth    bit set 1, means 
    tenth    bit set 1, means 
    eleventh bit set 1, means 
#>

. '.\VDI.Common.Function.ps1'
. '.\_Configuration.ps1'
Define-CommonVariables
Define-VDIBuildVariables
Define-ReclaimSpaceVariables

$ExitCode = 0
if(!$Enable)
{
    Add-Log -Path $strLogFile -Value 'Space reclaim script disabled'
    exit($ExitCode)
}

Add-Log -Path $strLogFile -Value 'Space reclaim script started'

if($StartAfter)
{
    Add-Log -Path $strLogFile -Value "StartAfter variable defined: [$StartAfter]"
    $Now = [datetime]::Now
    Add-Log -Path $strLogFile -Value "Current time: [$($Now.ToString('HH:mm'))]"
    $Sleep = [int]([datetime]::Parse($StartAfter) - $Now).TotalSeconds
    if($Sleep -gt 0)
    {
        Add-Log -Path $strLogFile -Value "Script should sleep: [$Sleep]"
        Start-Sleep -Seconds $Sleep
    }
}

#Add-PSSnapin Citrix* -ErrorAction SilentlyContinue
Add-PSSnapin VMware* -ErrorAction SilentlyContinue

$ClearVIServer = $false
if($DefaultVIServer -and $vCenter.Contains(($DefaultVIServer.Name).ToUpper())){
    Add-Log -Path $strLogFile -Value 'This script already connect to VI server, will not connect again'
}else{
    Connect-VIServer -Server $vCenter -ErrorAction:SilentlyContinue
    if(!$?){
        Add-Log -Path $strLogFile -Value "Connect to vCenter $vCenter failed, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        $ExitCode = $ExitCode -bor 0x0001 # 0000 0000 0000 0001
        Add-Log -Path $strLogFile -Value "Script exit with code $ExitCode"
        exit($ExitCode)
    }
}

$Clusters.Keys | %{$j = [int]::MinValue}{
    $Clusters.$_.$Type.Add('VMs', @(Get-VM -Location $Clusters.$_.$Type.Folder `
                                    | ?{$_.PowerState -eq 'PoweredOn' -and !(Get-SnapShot -VM $_) `
                                    -and [int]($_.CustomFields['Reclaim Status']) -lt $ReclaimCount `
                                    -and ($Date - [Datetime]($_.CustomFields['DeployDate'] -ireplace '^\s*$', '2015-01-01')).TotalDays -gt 60} `
                                    | Select-Object -Property @{N='Name'; E={$_.Name}}, `
                                    @{N='ReclaimStatus';E={$_.CustomFields['Reclaim Status']}} `
                                    | Sort-Object 'ReclaimStatus')
                          )
    if($Clusters.$_.$Type.VMs.Count -gt $j)
    {
        $j = $Clusters.$_.$Type.VMs.Count
    }
}

if($j -eq 0)
{
    Add-Log -Path $strLogFile -Value '[NOTE] There is no more VDIs could be migrated'
    Add-Log -Path $strLogFile -Value '[NOTE] This limitation here is to prevent unnecessary infinite vMotion'
    Add-Log -Path $strLogFile -Value '[NOTE] All VDIs migrated once should reclaimed enough space'
    Add-Log -Path $strLogFile -Value '[NOTE] If after a certain of time, space needs to be reclaim again'
    Add-Log -Path $strLogFile -Value '[NOTE] Please update _Configurations.ps1 increase 1 for variable "[ReclaimCount]"'
    Add-Log -Path $strLogFile -Value '[NOTE] Of course you still can trigger auto vMotion forever without manual work'
    Add-Log -Path $strLogFile -Value '[NOTE] For this, set variable [ReclaimCount] to a big value like 99999'
}

$i = $Processed = 0; $ReportArray = @()
while($i -lt $j)
{
    if(((Get-Date $StopBefore) - (Get-Date)).TotalSeconds -le 600)
    {
        Add-Log -Path $strLogFile -Value 'Timeline about to hit, main job quit'
        break
    }

    foreach($Cluster in $Clusters.Keys)
    {
        $VDI = $null
        $VDI = $Clusters.$Cluster.$Type.VMs[$i]
        if($VDI)
        {
            if(!(Get-VM -Name $VDI.Name).CustomFields['DeployDate'])
            {
                Set-Annotation -Entity $VDI.Name -CustomAttribute 'DeployDate' -Value '2015-01-01'
            }
            Add-Log -Path $strLogFile -Value "Processing VDI: [$($VDI.Name)][$($VDI.ReclaimStatus)]"
            (Get-VM -Name $VDI.Name).ExtensionData.RefreshStorageInfo()
            $Committed_Before = 0
            $Committed_Before = '{0:F2}' -f ((Get-VM -Name $VDI.Name).ExtensionData.Storage.PerDatastoreUsage[0].Committed/1GB)

            Add-Log -Path $strLogFile -Value 'Cleanup user temp folders from VDI'
            Get-ChildItem -LiteralPath "\\$($VDI.Name)\C$\Users\" | %{
                if(Test-Path -LiteralPath "$($_.FullName)\AppData\Local\Temp")
                {
                    Remove-Item -Path "$($_.FullName)\AppData\Local\Temp\*" -Recurse -Force -ErrorAction:SilentlyContinue
                }
            }
            Add-Log -Path $strLogFile -Value 'Cleanup Windows temp folder from VDI'
            if(Test-Path -LiteralPath "\\$($VDI.Name)\C$\Windows\Temp")
            {
                Remove-Item -Path "\\$($VDI.Name)\C$\Windows\Temp\*" -Recurse -Force -ErrorAction:SilentlyContinue
            }
            Add-Log -Path $strLogFile -Value 'Cleanup job completed'

            Copy-Item -Path 'VDI.Creation.Packages\sdelete.*' -Destination "\\$($VDI.Name)\C$\company\cmdtools" -Force
            if(!$?)
            {
                Add-Log -Path $strLogFile -Value "Failed to copy sdelete tool into VDI: [$($VDI.Name)]" -Type Error
                $ExitCode = $ExitCode -bor 0x0002 # 0000 0000 0000 0010
                continue
            }
            Add-Log -Path $strLogFile -Value 'Finished copy sdelete tools to VDI'
            # wmic /NODE:$VDI PROCESS CALL Create 'C:\company\cmdtools\sdelete.cmd'
            Invoke-Command -ComputerName $VDI.Name -ScriptBlock {& 'C:\company\cmdtools\sdelete.cmd'}
            if(!$?)
            {
                Add-Log -Path $strLogFile -Value "Failed to trigger sdelete script on VDI: [$($VDI.Name)]" -Type Error
                Add-Log -Path $strLogFile -Value $Error[0] -Type Error
                $ExitCode = $ExitCode -bor 0x0004 # 0000 0000 0000 0100
                Start-Sleep -Seconds 60
                if(!$StillMoveWhensDeleteFailed)
                {
                    continue
                }
            }
            else
            {
                Add-Log -Path $strLogFile -Value 'Finished run sdelete tools on VDI'
            }

            $Datastore = $null
            $Datastore = @(Get-VM $VDI.Name | Get-Datastore | %{$_.Name})[0]
            $NewDatastore = $Clusters.Keys | %{$Clusters.$_.$Type.Datastore} | ?{$_ -ne $Datastore} | Get-Random
            Add-Log -Path $strLogFile -Value "New datastore set to: [$NewDatastore]"
            
            Move-VM -VM $VDI.Name -Datastore $NewDatastore -DiskStorageFormat Thin
            if(!$?)
            {
                Add-Log -Path $strLogFile 'Failed to move VDI to new datastore, cause:' -Type Error
                Add-Log -Path $strLogFile $Error[0] -Type Error
                $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
                continue
            }
            Add-Log -Path $strLogFile -Value 'Finished move VM to new datastore'
            (Get-VM -Name $VDI.Name).ExtensionData.RefreshStorageInfo()
            $Committed_After = 0
            $Committed_After = '{0:F2}' -f ((Get-VM -Name $VDI.Name).ExtensionData.Storage.PerDatastoreUsage[0].Committed/1GB)
            Add-Log -Path $strLogFile -Value "[$($VDI.Name)][$Committed_Before][$Committed_After][$($Committed_After - $Committed_Before)]"
            $Processed++
            $ReportArray += $null
            $ReportArray[-1] = @($VDI.Name, $Committed_Before, $Committed_After)
            Set-Annotation -Entity $VDI.Name -CustomAttribute 'Reclaim Status' -Value ([int]($VDI.ReclaimStatus) + 1)
            Set-Annotation -Entity $VDI.Name -CustomAttribute 'Reclaim Last Date' -Value $strDate
        }
        else
        {
            Start-Sleep -Seconds 30
            continue
        }
    }
    $i++
}

if($ClearVIServer){
    Disconnect-VIServer -Server $vCenter -Confirm:$false -Force
}

Remove-PSSnapin VMware* -Confirm:$false

Add-Log -Path $strLogFile -Value "Total processed: [$Processed], report array followed:`r`n$(@($ReportArray | %{$_ -join "`t"}) -join "`r`n")"

exit($ExitCode)
