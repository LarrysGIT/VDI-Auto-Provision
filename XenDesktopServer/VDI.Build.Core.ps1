
<#
    Version 1.02 [Author: Larry Song]
        Add version control.
        Re-write UK's script to VDI.Build.Core.ps1
    Version 1.03 [Author: Larry Song; Time: 2014-02-11]
        Adjust parameters "VDIName" and "VDINameToGo", "VDIName" has high priority than "VDINameToGo"
    Version 1.04 [Author: Larry Song; Time: 2014-02-13]
        Adjust parameters "Cluster" and "ClusterToGo", "Cluster" has high priority than "ClusterToGo"
    Version 1.05 [Author: Larry Song; Time: 2014-03-05]
        Add new parameters "-WorkingDirectory", "-Test", "-NumCPU" and "-MemoryGB"
    Version 1.06 [Author: Larry Song; Time: 2014-03-29]
        Fix bug -MemoryGB, since x86 VMware snapin supports only -MemoryMB.
    Version 1.07 [Author: Larry Song; Time: 2014-03-30]
        Fix bug unable to add Type2 VDI to Type2 AD group.
    Version 1.08 [Author: Larry Song; Time: 2014-09-18]
        Biz visitor pool requirement, add script to add Biz new users into g-XX-XenDesktop_Users
    Version 1.09 [Author: Larry Song; Time: 2014-12-11]
        Cancel mandatory parameter $User, so script can assign nobody to machine.
        If user is not provided, machine name must provided
        Fix the bug which can't capture vmrun exit code
        Add exit code ignore option for Exec-VMRUN function
    Version 1.10 [Author: Larry Song; Time: 2014-12-17]
        Fix the bug can't join VDI to catalog and desktop group
    Version 1.11 [Author: Larry Song; Time: 2015-04-07]
        Add final reboot at end of the script to get GUID for SCCM
    Version 1.12 [Author: Larry Song; Time: 2015-07-03]
        Fix a bug, template sometimes has no disk, unable to boot up, so script will re-pick another template
    Version 1.13 [Author: Larry Song; Time: 2015-07-07]
        Set new value annotation "DeployDate"
#>

PARAM(
    [parameter(Mandatory=$true)]
    [string]$VIServer,

    [parameter()]
    [int]$VIPort = 443,

    [parameter()]
    [string]$User,

    [parameter(Mandatory=$true)]
    [string]$Prefix,

    [parameter(Mandatory=$true)]
    [string]$DesktopGroup,

    [parameter(Mandatory=$true)]
    [string]$DDCList,

    [parameter(Mandatory=$true)]
    [string]$SupportGroup,

    [parameter(Mandatory=$true)]
    [ValidatePattern('(?i)FIL|IM')]
    [string]$Environment,

    [parameter()]
    [string]$Cluster,

    [parameter(Mandatory=$true)]
    [string]$ClusterToGo,

    [parameter()]
    [string]$IMGroup,

    [parameter()]
    [string]$BizGroup,

    [parameter()]
    [string]$Folder,

    [parameter()]
    [string]$VDIName,

    [parameter()]
    [string]$VDINameToGo,

    [parameter()]
    [string]$NetworkPositive,

    [parameter()]
    [string]$NetworkNegative,

    [parameter()]
    [string]$TemplateFolder,

    [parameter()]
    [string]$Template,

    [parameter()]
    [ValidatePattern('^[1-9]$|^10$')]
    [int]$NumCPU,

    [parameter()]
    [ValidatePattern('^[1-9]$|^1[0-9]$')]
    [int]$MemoryGB,

    [parameter()]
    [switch]$IgnoreVDINameAvailable,

    [parameter()]
    [string]$WorkingDirectory,

    [parameter()]
    [string]$PackageFolder = "VDI.Creation.Packages",

    [parameter()]
    [string]$LogFile,

    [parameter()]
    [string]$ExitCodeWriteTo,

    [parameter()]
    [switch]$AutoCreateXDGroups = $true,

    [parameter()]
    [switch]$Test
)

$NetworkNegative += " 145 48"

function Quit{
    PARAM(
        [int]$ExitCode = 0
    )
    if($WorkingDirectory){
         Write-Host 'Move out from working directory.' -ForegroundColor Yellow
         Pop-Location
    }
    if($ClearVIServer){
        Disconnect-VIServer -Server $VIServer -Confirm:$false -Force
        $ClearVIServer = $false
    }
    if($ExitCodeWriteTo){
        Add-Content -Path $ExitCodeWriteTo -Value "$PID`t$ExitCode"
    }
    Exit($ExitCode)
}

$ExitCode = 0

if($WorkingDirectory){
     Write-Host 'Working directory provided, move in.' -ForegroundColor Yellow
     Push-Location -Path $WorkingDirectory
}

$PackageFolder = "$PWD\$PackageFolder"

if(Test-Path -Path '.\VDI.Common.Function.ps1' -PathType Leaf){
    Write-Host 'Common functions loaded.' -ForegroundColor Yellow
    . '.\VDI.Common.Function.ps1'
}else{
    Write-Host 'Common functions file not found, script can not continue.' -ForegroundColor Red
    Quit -ExitCode 1
}

function Check-Prerequisite{
    if(!$User)
    {
        if($VDIName -imatch '^$' -and $VDINameToGo -imatch '^$')
        {
            Quit -ExitCode 8
        }
    }
    if([IntPtr]::Size -ne 4){
        Add-Log -Path $LogFile -Value 'This script must be invoked on 32bit powershell.exe!' -Type Error
        Add-Log -Path $LogFile -Value 'Please use powershell.exe in SysWOW64 folder.' -Type Error
        Add-Log -Path $LogFile -Value 'Script quit.'
        Quit -ExitCode 2
    }
    if(Test-Path -Path "$PackageFolder\vmrun.exe" -PathType Leaf){}else{
        Add-Log -Path $LogFile -Value 'vmrun.exe not found, can not continue.'
        Quit -ExitCode 3
    }
    if(Test-Path -Path "$PackageFolder\W7Support.exe" -PathType Leaf){}else{
        Add-Log -Path $LogFile -Value 'W7Support.exe not found, can not continue.'
        Quit -ExitCode 4
    }
    if(Test-Path -Path "$PackageFolder\UnPackPayload.vbs" -PathType Leaf){}else{
        Add-Log -Path $LogFile -Value 'UnPackPayload.vbs not found, can not continue.'
        Quit -ExitCode 5
    }
    if(Test-Path -Path "$PackageFolder\ccmdelcert.exe" -PathType Leaf){}else{
        Add-Log -Path $LogFile -Value 'ccmdelcert.exe not found, can not continue.'
        Quit -ExitCode 6
    }
    Add-Log -Path $LogFile -Value 'Add vmware and citrix snapin, import ActiveDirectory module anyway.'
    Add-PSSnapin *vmware* -ErrorAction:SilentlyContinue
    Add-PSSnapin *citrix* -ErrorAction:SilentlyContinue
    Import-Module ActiveDirectory
}

# Add-Log -Path $LogFile -Value "Script initialized, check prerequisites."
Check-Prerequisite

$DomainSuffix = $env:USERDNSDOMAIN
if($VIServer -notmatch '^\w+\..+'){
    Add-Log -Path $LogFile -Value 'The VIServer name provided is not a FQDN, add current domain suffix.'
    $VIServer += ".$DomainSuffix"
    Add-Log -Path $LogFile -Value "VIServer name set to $VIServer"
}
$VIServer = $VIServer.ToUpper()

if($User)
{
    $objUser = Get-ADExisting -SAMAccountName $User -Type user
    if(!$objUser){
        Add-Log -Path $LogFile -Value "$User not found in AD." -Type "Error"
        Quit -ExitCode 7
    }
    $UserToGo = $User
    Add-Log -Path $LogFile -Value 'User found in AD, details:'
    Add-Log -Path $LogFile -Value "SAMAccountName: $($objUser.Properties['samaccountname'][0])"
    Add-Log -Path $LogFile -Value "Display Name  : $($objUser.Properties['displayname'][0])"
    Add-Log -Path $LogFile -Value "Title         : $($objUser.Properties['title'][0])"
    Add-Log -Path $LogFile -Value "User DN       : $($objUser.Properties['distinguishedname'][0])"
}
else
{
    Add-Log -Path $LogFile -Value 'User parameter is null'
}

############# VDI pick up
if($VDIName){
    Add-Log -Path $LogFile -Value "VDI name provided: $VDIName."
    if($VDINameToGo){
        Add-Log -Path $LogFile -Value "VDI name abandoned: $VDINameToGo."
    }
    $VDINameToGo = $VDIName
}

if($UserToGo)
{
    $VDIName_ = @(Get-AvailableVDIName -User $UserToGo -Prefix $Prefix)
    Add-Log -Path $LogFile -Value "Retrieved available VDI from AD completed, count: $($VDIName_.Count)"
    Add-Log -Path $LogFile -Value "Available VDI names: $($VDIName_ -join ' ')"
    if($VDIName_.Count -eq 0)
    {
        Add-Log -Path $LogFile -Value 'No VDI name available in AD.' -Type Warning
    }
}
else
{
    $VDIName_ = @()
}

if($VDINameToGo)
{
    Add-Log -Path $LogFile -Value "VDI name provided: $VDINameToGo."
    if($UserToGo)
    {
        if($VDIName_ -notcontains $VDINameToGo)
        {
            Add-Log -Path $LogFile -Value "$VDINameToGo not in available list, continue check IgnoreVDINameAvailable parameter."
            if($IgnoreVDINameAvailable){
                Add-Log -Path $LogFile -Value "IgnoreVDINameAvailable parameter detected, force use $VDINameToGo."
            }else{
                Add-Log -Path $LogFile -Value 'IgnoreVDINameAvailable parameter not set, VDI name provided can not be used, quit.'
                Quit -ExitCode 9
            }
        }
    }
    else
    {
        if(Get-ADExisting -SAMAccountName $VDINameToGo -Type computer)
        {
            Add-Log -Path $LogFile -Value "$VDINameToGo already exists in AD, continue check IgnoreVDINameAvailable parameter."
            if($IgnoreVDINameAvailable){
                Add-Log -Path $LogFile -Value "IgnoreVDINameAvailable parameter detected, force use $VDINameToGo."
            }else{
                Add-Log -Path $LogFile -Value 'IgnoreVDINameAvailable parameter not set, VDI name provided can not be used, quit.'
                Quit -ExitCode 9
            }
        }
    }
}
else
{
    if($VDIName_.Count -ne 0)
    {
        Add-Log -Path $LogFile -Value 'VDI name not provided, script will auto pick one.'
        $VDINameToGo = $VDIName_[0]
    }
    else
    {
        Add-Log -Path $LogFile -Value 'VDI name not provided, also no available VDI names.'
        Quit -ExitCode 8
    }
}
Add-Log -Path $LogFile -Value "VDI name is going to use: $VDINameToGo"

############# VDI pick up
$ClearVIServer = $false
if($DefaultVIServer -and $VIServer.Contains(($DefaultVIServer.Name).ToUpper())){
    Add-Log -Path $LogFile -Value 'This script already connect to VI server, will not connect again.'
}else{
    Connect-VIServer -Server $VIServer -Port $VIPort -ErrorAction:Stop | Out-Null
    if(!$?){
        Add-Log -Path $LogFile -Value "Connect to vCenter server $VIServer on port $VIPort failed, cause:" -Type Error
        Add-Log -Path $LogFile -Value $Error[0] -Type Error
        Add-Log -Path $LogFile -Value 'Script quit.'
        Remove-PSSnapin *vmware* -ErrorAction:SilentlyContinue
        Quit -ExitCode 10
    }
    $ClearVIServer = $true
}
Add-Log -Path $LogFile -Value "Connected to vCenter $VIServer on port $VIPort."
$VDI_ALL = @(Get-VM -Name "${Prefix}*" | ?{$_.PowerState -ne 'PoweredOff'})
Add-Log -Path $LogFile -Value "Collected all VDIs for future using: $($VDI_ALL.Count)."

############# Cluster pick up
if($Cluster){
    Add-Log -Path $LogFile -Value "Cluster name provided: $Cluster."
    if($ClusterToGo){
        Add-Log -Path $LogFile -Value "VDI name abandoned: $ClusterToGo."
    }
    $ClusterToGo = $Cluster
}
$Cluster_ = @(Get-Cluster -Name $ClusterToGo -ErrorAction:SilentlyContinue)
if($Cluster_.Count -ne 1){
    Add-Log -Path $LogFile -Value "After all processed, cluster should be only 1 option, result: $($Cluster_.Count)" -Type Warning
    Add-Log -Path $LogFile -Value "Script can't determine cluster, quit." -Type Warning
    Quit -ExitCode 11
}
Remove-Variable -Name ClusterToGo -Confirm:$false
$ClusterToGo = @($Cluster_[0].Name, $Cluster_[0].Id)
Add-Log -Path $LogFile -Value "Cluster chose: $($Cluster_[0])."
############# Cluster pick up
############# Network pick up
function NetworkAutoPick{
    PARAM(
        [string[]]$Networks,
        [string[]]$Yes,
        [string[]]$No
    )
    $obj = New-Object PSObject -Property @{Name = $null; Number = $null}
    $objNetworks = @()
    $Networks | %{
        if($Yes -and ($Yes -notcontains $_)){
            Add-Log -Path $LogFile -Value "vLan $_ not found in prefer list, discarded."
            return
        }
        if($No -and ($No -icontains $_)){
            Add-Log -Path $LogFile -Value "vLan $_ found in avoid list, discarded."
            return
        }
        $objNetworks += $obj.PSObject.Copy()
        $objNetworks[-1].Name = $_
        $objNetworks[-1].Number = @((Get-NetworkAdapter -VM $VDI_ALL | ?{$_.NetworkName -imatch "_$($objNetworks[-1].Name)_"})).Count
    }
    $objNetworks = @($objNetworks | Sort-Object Number)
    return $objNetworks[0].Name
}

$NetworkToGo = [regex]::Matches($NetworkPositive, '\d+') | %{$_.Value}
$NetworkToAvoid = [regex]::Matches($NetworkNegative, '\d+') | %{$_.Value}
$Network_ = Get-VMHost -Location (Get-Cluster -Id $ClusterToGo[1]) | Get-VirtualPortGroup | %{if($_.Name -imatch '_(\d+)_'){$Matches[1]}}
Add-Log -Path $LogFile -Value "Retrieved networks from cluster $($ClusterToGo[0]), count: $($Network_.Count)"
Add-Log -Path $LogFile -Value "Available network names: $($Network_ -join " ")"
if($NetworkToGo){
    Add-Log -Path $LogFile -Value "Prefer network provided, name: $($NetworkToGo -join " ")"
}
if($NetworkToAvoid){
    Add-Log -Path $LogFile -Value "Avoid network provided, name: $($NetworkToAvoid -join " ")"
}
$NetworkToGo = NetworkAutoPick -Networks $Network_ -Yes $NetworkToGo -No $NetworkToAvoid
$NetworkToGo = @(Get-VMHost -Location (Get-Cluster -Id $ClusterToGo[1]) | Get-VirtualPortGroup -Name "*_${NetworkToGo}_*")
if($NetworkToGo.Count -ne 1){
    Add-Log -Path $LogFile -Value "After all processed, network should be only 1 option, result: $($NetworkToGo.Count)" -Type Warning
    Add-Log -Path $LogFile -Value "Script can't determine network, quit." -Type Warning
    Quit -ExitCode 12
}
$NetworkToGo = @($NetworkToGo[0].Name, $NetworkToGo[0].Id)
Add-Log -Path $LogFile -Value "Network chose: $($NetworkToGo[0])."
############# Network pick up
############# Template folder pick up
if(!($TemplateFolder -or $Template)){
    Add-Log -Path $LogFile -Value "Template or template folder not specified." -Type Error
    Add-Log -Path $LogFile -Value "Unable to determine template. Script quit." -Type Error
    Quit -ExitCode 13
}

function TemplatePickup{
    $TemplateFolderToGo = $TemplateFolder -ireplace '^$', '*'
    $TemplateToGo = $Template -ireplace '^$', '*'
    Add-Log -Path $LogFile -Value "Template folder: $TemplateFolderToGo."
    Add-Log -Path $LogFile -Value "Template: $TemplateToGo."

    $Template_ = @(Get-VM -Name $TemplateToGo -Location (Get-Folder -Name $TemplateFolderToGo))
    Add-Log -Path $LogFile -Value "Get VM template in folder count: $($Template_.Count)."

    $Template_ = @($Template_ | ?{$_.PowerState -eq 'PoweredOff' -and $_.VMHost.Parent.Id -eq $ClusterToGo[1]} | Sort-Object Name)
    Add-Log -Path $LogFile -Value "Remove powered on VMs and choose only on $($ClusterToGo[0]), left: $($Template_.Count)."
    if($Template_.Count -eq 0){
        Add-Log -Path $LogFile -Value "No template available. Script quit." -Type Warning
        Quit -ExitCode 14
    }
    return @($Template_[0].Name, $Template_[0].Id)
}
$TemplateToGo = TemplatePickup
Add-Log -Path $LogFile -Value "Set template to: [$($TemplateToGo -join '][')]"

############# Template folder pick up
$FolderToGo = $Folder

############# Start build
$DesktopGroupToGo = $DesktopGroup

Add-Log -Path $LogFile -Value ' ******** Summary start.'
Add-Log -Path $LogFile -Value " * User    : $UserToGo"
Add-Log -Path $LogFile -Value " * VDI Name: $VDINameToGo"
Add-Log -Path $LogFile -Value " * Cluster : $($ClusterToGo[0])"
Add-Log -Path $LogFile -Value " * Network : $($NetworkToGo[0])"
Add-Log -Path $LogFile -Value " * Template: $($TemplateToGo[0])"
Add-Log -Path $LogFile -Value " * Citrix  : $DesktopGroupToGo"
Add-Log -Path $LogFile -Value " * Folder  : $FolderToGo"
if($NumCPU){
Add-Log -Path $LogFile -Value " * CPU Num : $NumCPU"
}
if($MemoryGB){
Add-Log -Path $LogFile -Value " * Memory  : ${MemoryGB} GB"
}
Add-Log -Path $LogFile -Value ' ******** Summary end.'

if($Test){
    Add-Log -Path $LogFile -Value 'This instance is a TEST, script will not continue.'
    Write-Host 'This intance marked as a TEST, script stopped.' -ForegroundColor Yellow
    Write-Host 'Press any key to quit...' -ForegroundColor Yellow
    $x = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Quit 
}

### Get template and move to folder.
function MoveTemplate
{
    if($FolderToGo){
        $objTheVM = Get-VM -Name $VDINameToGo -ErrorAction:SilentlyContinue
        if($objTheVM){
            Add-Log -Path $LogFile -Value "$VDINameToGo found in folder $FolderToGo, VM id: $($objTheVM.Id)." -Type Warning
            $objTheVM | Set-VM -Name "${VDINameToGo}_OLD_$((Get-Date).ToString('yyyyMMddHHmmss'))" -Confirm:$false
        }
        Get-VM -Id $TemplateToGo[1] | Move-VM -Destination (Get-Folder -Name $FolderToGo) -ErrorAction:SilentlyContinue
        if(!$?){
            Add-Log -Path $LogFile -Value "Failed to move $($TemplateToGo[0]) to $($FolderToGo[0]), cause:" -Type Warning
            Add-Log -Path $LogFile -Value $Error[0] -Type Warning
        }else{
            Add-Log -Path $LogFile -Value 'VM moved to target folder.'
        }
    }
}
MoveTemplate

function SetVM
{
    Get-VM -Id $($TemplateToGo[1]) | Set-VM -Name $VDINameToGo -Confirm:$false
    if($NumCPU -gt 0){
        Get-VM -Id $($TemplateToGo[1]) | Set-VM -NumCPU $NumCPU -Confirm:$false
    }
    if($MemoryGB -gt 0){
        Get-VM -Id $($TemplateToGo[1]) | Set-VM -MemoryMB ($MemoryGB * 1024) -Confirm:$false
    }
}
SetVM

### Set network.
function SetNetwork
{
    Get-VM -Id $TemplateToGo[1] | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName (Get-VirtualPortGroup -Id $NetworkToGo[1]) -Confirm:$false
    if(!$?){
        Add-Log -Path $LogFile -Value 'Failed to set virtual network group to VM, cause:' -Type Error
        Add-Log -Path $LogFile -Value $Error[0] -Type Error
        Quit -ExitCode 15
    }else{
        Add-Log -Path $LogFile -Value 'Set virtual network group for VM.'
    }
    Get-VM -Id $TemplateToGo[1] | Get-NetworkAdapter | Set-NetworkAdapter -StartConnected:$false -Confirm:$false | Out-Null
    if(!$?){
        Add-Log -Path $LogFile -Value 'Failed to set network adapter as disconnected.' -Type Error
        Add-Log -Path $LogFile -Value $Error[0] -Type Error
        Quit -ExitCode 16
    }else{
        Add-Log -Path $LogFile -Value 'Set network adapter for VM as disconnected.'
    }
}
SetNetwork

### Booting VM now.
Add-Log -Path $LogFile -Value 'Power on VM and waiting for VM tools running.'
do
{
    Get-VM -Id $TemplateToGo[1] | Start-VM -Confirm:$False | Wait-Tools -ErrorAction:SilentlyContinue
    if(!$?)
    {
        Add-Log -Path $strLogFile -Value 'Wait-Tools command failed, this is abnormal, error:' -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        Add-Log -Path $strLogFile -Value 'Shutdown and remove the template and going to pickup a new'
        Stop-VM -VM (Get-VM -Id $TemplateToGo[1]) -Confirm:$false
        Remove-VM -VM (Get-VM -Id $TemplateToGo[1]) -DeletePermanently -Confirm:$false
        if(!$?)
        {
            Add-Log -Path $strLogFile -Value 'Remove bad template failed, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            Add-Log -Path $strLogFile -Value "At least rename the VM to another name: [$($TemplateToGo[0])_BADTemplate_$strDate]"
            Set-VM -VM (Get-VM -Id $TemplateToGo[1]) -Name "$($TemplateToGo[0])_BADTemplate_$strDate" -Confirm:$false
        }
        Add-Log -Path $strLogFile -Value 'Trigger new template pickup and redo previous VM jobs'
        $TemplateToGo = TemplatePickup
        MoveTemplate
        SetVM
        SetNetwork
        Add-Log -Path $strLogFile -Value 'Sleep 60 seconds to continue'
        Start-Sleep -Seconds 60
    }
    else
    {
        break
    }
}
while($true)

function Shutdown-VMAndWait{
    PARAM(
        [string]$VMId
    )
    $CommandTimeout = 5; $Count = 5
    while((Get-VM -Id $VMId).'PowerState' -ne 'PoweredOff'){
        $Count++
        Start-Sleep -Seconds 5
        if($Count -gt $CommandTimeout){
            Add-Log -Path $LogFile -Value 'Send shutdown signal to VM again.'
            Get-VM -Id $VMId | Shutdown-VMGuest -Confirm:$false | Out-Null
            $Count = 0
        }
    }
}

Add-Log -Path $LogFile -Value 'Send shutdown signal to VM, and waiting for power off.'
Shutdown-VMAndWait -VMId $TemplateToGo[1]

$objVMConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$objOptionValue = New-Object VMware.Vim.optionvalue
$objOptionValue.Key = 'devices.hotplug'
$objOptionValue.Value = 'FALSE'
$objVMConfigSpec.extraconfig = $objOptionValue
$objVMConfigSpec.BootOptions = New-Object VMware.Vim.VirtualMachineBootOptions
$objVMConfigSpec.BootOptions.BootDelay = "0"

(Get-VM -Id $TemplateToGo[1] | Get-View).ReconfigVM_Task($objVMConfigSpec)

### Booing VM now.
Get-VM -Id $TemplateToGo[1] | Start-VM -Confirm:$false | Wait-Tools

### this is secert
$Hash1 = HashConv -strHash (Get-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash1').Hash1
$Hash2 = HashConv -strHash (Get-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash2').Hash2
$Hash5 = HashConv -strHash (Get-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash5').Hash5
$Hash6 = HashConv -strHash (Get-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash6').Hash6

### Copy package to VM, and remote execution.
$strVMPath = (Get-VM -Id $TemplateToGo[1]).ExtensionData.Config.Files.VmPathName
Add-Log -Path $LogFile -Value "Got VM path $strVMPath"
$vCenterSDK = "https://$VIServer/sdk"
Add-Log -Path $LogFile -Value "vCenter SDK set to: $vCenterSDK"

function Exec-VMRUN{
    PARAM(
        [string]$vmRunPath = "$PackageFolder\vmrun.exe",
        [string]$Actions,
        [int]$Timeout = 120,
        [int]$ProcessTTL = 30,
        [switch]$IgnoreExitCode
    )
    Add-Log -Path $LogFile -Value "VMRUN function triggered, timeout: $Timeout, process TTL: $ProcessTTL."
    Add-Log -Path $LogFile -Value "Try vmrun to run [$Actions]."
    $TimeStart = Get-Date
    $pStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $pStartInfo.FileName = (Get-Item -Path $vmRunPath -ErrorAction:SilentlyContinue).FullName
    $pStartInfo.RedirectStandardError = $true
    $pStartInfo.RedirectStandardOutput = $true
    $pStartInfo.UseShellExecute = $false
    $pStartInfo.WindowStyle = 'Hidden'
    $pStartInfo.Arguments = "-T vc -h `"$vCenterSDK`" -u $Hash1 -p $Hash2 -gu $Hash5 -gp $Hash6 $Actions"
    $pProcess = New-Object System.Diagnostics.Process
    $pProcess.StartInfo = $pStartInfo
    do{
        # $objProcess = Start-Process -FilePath $vmRunPath -ArgumentList @('-T vc', "-h $vCenterSDK", "-u $Hash1", "-p $Hash2", "-gu $Hash5", "-gp $Hash6", $Actions) -PassThru
        $pProcess.Start()
        if(!$?){
            Add-Log -Path $LogFile -Value 'Try to start vmrun failed, cause:' -Type Error
            Add-Log -Path $LogFile -Value $Error[0] -Type Error
            Add-Log -Path $LogFile -Value 'This problem need to be fixed, script quit.' -Type Error
            Quit -ExitCode 17
        }
        Add-Log -Path $LogFile -Value "Process started, PID: $($pProcess.Id)"
        if($pProcess.WaitForExit($ProcessTTL * 1000)){
            Add-Log -Path $LogFile -Value "Process completed, exit code: $($pProcess.ExitCode)"
            if($pProcess.ExitCode -imatch '^$|0'){
                Add-Log -Path $LogFile -Value 'Exit code looks good, no need to run again.'
                break
            }else{
                Add-Log -Path $LogFile -Value 'Bad exit code, try to run again.' -Type Warning
                Add-Log -Path $LogFile -Value $pProcess.StandardOutput.ReadToEnd() -Type Error
                Add-Log -Path $LogFile -Value $pStartInfo.Arguments -Type Debug
                if($IgnoreExitCode)
                {
                    Add-Log -Path $LogFile -Value 'Ignore exit code enabled, continue'
                    break
                }
            }
        }else{
            Add-Log -Path $LogFile -Value 'TTL failed, kill process and try run again.' -Type Warning
            $pProcess.Kill()
        }
        if(((Get-Date) - $TimeStart).TotalSeconds -ge $Timeout){
            Add-Log -Path $LogFile -Value 'Timed out, no need to run.' -Type Warning
            break
        }
    }
    while($true)
}

$WinRoot = "$($env:SystemRoot)\System32"
Add-Log -Path $LogFile -Value 'Copy W7Support package to VM.'
Exec-VMRUN -Actions "copyFileFromHostToGuest `"$strVMPath`" `"$PackageFolder\W7Support.exe`" `"C:\company\W7Support.exe`""
Add-Log -Path $LogFile -Value 'Copy unpack VBS script to VM.'
Exec-VMRUN -Actions "copyFileFromHostToGuest `"$strVMPath`" `"$PackageFolder\UnPackPayload.vbs`" `"C:\company\UnPackPayload.vbs`""
Add-Log -Path $LogFile -Value 'Run unpack script on VM.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\cscript.exe`" `"C:\company\UnPackPayload.vbs`""
Add-Log -Path $LogFile -Value 'Set ccmexec service to manual start.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\sc.exe`" config ccmexec start= demand"
Add-Log -Path $LogFile -Value 'Set powershell execution policy.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\cscript.exe`" `"C:\company\W7Support\SetupPowershell.vbs`""
Add-Log -Path $LogFile -Value "Run powershell script to change computer name to $($VDINameToGo)."
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\WindowsPowerShell\v1.0\powershell.exe`" -File `"C:\company\W7Support\RenameComputer.ps1`" $VDINameToGo"

Add-Log -Path $LogFile -Value 'Shutdown VM for name changing.'
Shutdown-VMAndWait -VMId $TemplateToGo[1]

Add-Log -Path $LogFile -Value 'Set network adapter to connect.'
Get-VM -Id $TemplateToGo[1] | Get-NetworkAdapter | Set-NetworkAdapter -StartConnected:$true -Confirm:$false

Add-Log -Path $LogFile -Value 'Start VM for addressing ip from DHCP.'
Get-VM -Id $TemplateToGo[1] | Start-VM -Confirm:$false | Wait-Tools

$vLanID = (Get-VirtualPortGroup -Id $NetworkToGo[1]).ExtensionData.Config.DefaultPortConfig.Vlan.VlanId
Add-Log -Path $LogFile -Value 'Capturing IPv4 address for the VM.'
while(!($IPv4Addr = (Get-VM -Id $TemplateToGo[1] | Get-VMGuest).IPAddress -imatch "^10\.(?:161|162)\.$vLanID\.\d+$")){
    Start-Sleep -Seconds 10
}
Add-Log -Path $LogFile -Value "Valid IPv4 address captured: [$($IPv4Addr -join "][")]."

Add-Log -Path $LogFile -Value 'Trigger domain join script in guest.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\WindowsPowerShell\v1.0\powershell.exe`" -File `"C:\company\W7Support\AddToDomain.ps1`""

Add-Log -Path $LogFile -Value "Shutdown and start VM for AD joining."
Shutdown-VMAndWait -VMId $TemplateToGo[1]
Get-VM -Id $TemplateToGo[1] | Start-VM -Confirm:$false | Wait-Tools

Add-Log -Path $LogFile -Value 'Waiting computer object sync in AD.'
while(!(Get-ADExisting -SAMAccountName $VDINameToGo -Type computer)){
    Start-Sleep -Seconds 15
}

Add-Log -Path $LogFile -Value "$VDINameToGo found in AD, continue config SCCM."
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"C:\company\W7Support\ccmdelcert.exe`""
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\cmd.exe`" `"/c del /q /f C:\company\W7Support\AddToDomain.ps1`""

Add-Log -Path $LogFile -Value 'Set custom fields.'
Get-VM -Id $TemplateToGo[1] | Set-Annotation -CustomAttribute 'DeployDate' -Value $strDate
Get-VM -Id $TemplateToGo[1] | Set-Annotation -CustomAttribute 'Description' -Value $DesktopGroupToGo
Get-VM -Id $TemplateToGo[1] | Set-Annotation -CustomAttribute 'Owner' -Value $User
Get-VM -Id $TemplateToGo[1] | Set-Annotation -CustomAttribute 'Real Name' -Value $($objUser.Properties['displayname'])
Get-VM -Id $TemplateToGo[1] | Set-Annotation -CustomAttribute 'XenDesktop Group' -Value $DesktopGroupToGo

Add-Log -Path $LogFile -Value 'Set notes field for VM.'
$strDescription =   $DesktopGroup + " : " +  $User + " : Created on " + $(Get-Date) + " : By " + $env:USERNAME + " (" + $((Get-ADExisting -SAMAccountName $($env:USERNAME) -Type user).Properties['displayname']) + ")"
Get-VM -Id $TemplateToGo[1] | Set-VM -Description $strDescription -Confirm:$false

Add-Log -Path $LogFile -Value 'Run activation executions.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\cscript.exe`" $WinRoot\slmgr.vbs /ato"
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\cscript.exe`" `"C:\company\W7Support\OSPP.VBS`" /act"

Add-Log -Path $LogFile -Value 'Set ccmexec service to auto.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\sc.exe`" config ccmexec start= auto"

Add-Log -Path $LogFile -Value 'Shutdown and start VM to apply changes.'
Shutdown-VMAndWait -VMId $TemplateToGo[1]
Get-VM -Id $TemplateToGo[1] | Start-VM -Confirm:$false | Wait-Tools

$DDCList_ = $DDCList.Split(' ')
$DDCList_ = @($DDCList_ | %{if($_){if($_ -notmatch '^\w+\..+'){"${_}.${DomainSuffix}"}else{$_}}})
Add-Log -Path $LogFile -Value 'Add DDC server list to registry.'
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\reg.exe`" add `"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Citrix\VirtualDesktopAgent`" /v ListOfDDCs /d `"$($DDCList_ -join ' ')`" /f"

Add-Log -Path $LogFile -Value "Add remote management user $SupportGroup."
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"C:\company\W7Support\ConfigRemoteMgmt.exe`" /configwinrmuser intl\$SupportGroup" -Timeout 300 -ProcessTTL 150 -IgnoreExitCode

Add-Log -Path $LogFile -Value "Add user to remote group on VM."
Exec-VMRUN -Actions "runProgramInGuest `"$strVMPath`" `"$WinRoot\net.exe`" localgroup `"Remote Desktop Users`" intl\$User /add"

Add-Log -Path $LogFile -Value 'Trigger a last reboot for VDI to get SCCM new GUID'
Restart-VMGuest -VM $VDINameToGo -Confirm:$false
if(!$?)
{
    Add-Log -Path $LogFile -Value 'Restart VM guest failed, cause:' -Type Error
    Add-Log -Path $LogFile -Value $Error[0] -Type Error
}

$VMuuid = (Get-VM -Id $TemplateToGo[1]).ExtensionData.Config.Uuid

Add-Log -Path $LogFile -Value 'Release session from vCenter.'
if($ClearVIServer){
    Disconnect-VIServer -Server $VIServer -Confirm:$false -Force
    $ClearVIServer = $false
}

Add-Log -Path $LogFile -Value 'Start adding the VM to Citrix DDC.'

$strAdminDDC = $DDCList_[0]
$Catalog = $DesktopGroupToGo
Add-Log -Path $LogFile -Value "DDC server picked: $strAdminDDC."
Add-Log -Path $LogFile -Value "Catalog container: $Catalog."
Add-Log -Path $LogFile -Value "Desktop container: $DesktopGroupToGo."

if(!(Get-Brokercatalog -Name $Catalog)){
    if($AutoCreateXDGroups)
    {
        Add-Log -Path $LogFile -Value "Catalog $Catalog not found, auto create."
        New-BrokerCatalog -AllocationType 'Permanent' -CatalogKind PowerManaged -Description "Catalog1 Description" -Name $Catalog -AdminAddress $strAdminDDC
    }
}
$strCatalogUid = (Get-BrokerCatalog -Name $Catalog -AdminAddress $strAdminDDC).Uid

if(!(Get-BrokerDesktopGroup -Name $DesktopGroupToGo)){
    if($AutoCreateXDGroups)
    {
        Add-Log -Path $LogFile -Value "Desktop $DesktopGroupToGo not found, auto create."
        New-BrokerDesktopGroup $DesktopGroup -PublishedName $DesktopGroupToGo -DesktopKind 'Private' -OffPeakBufferSizePercent 10 -PeakBufferSizePercent 10 -ShutdownDesktopsAfterUse $False -TimeZone 'GMT Standard Time' -AdminAddress $strAdminDDC
	    $strAGAccess = "${DesktopGroupToGo}_AG"
	    $strDirectAccess = "${DesktopGroupToGo}_Direct"
        New-BrokerAccessPolicyRule -AllowedConnections 'NotViaAG' -AllowedProtocols @('RDP','HDX') -AllowedUsers 'AnyAuthenticated' -AllowRestart $True -Enabled $True -IncludedDesktopGroupFilterEnabled $True -IncludedDesktopGroups @($DesktopGroupToGo) -IncludedSmartAccessFilterEnabled $True -IncludedUserFilterEnabled $True -Name $strDirectAccess -AdminAddress $strAdminDDC
        New-BrokerAccessPolicyRule -AllowedConnections 'ViaAG' -AllowedProtocols @('RDP','HDX') -AllowedUsers 'AnyAuthenticated' -AllowRestart $True -Enabled $True -IncludedDesktopGroupFilterEnabled $True -IncludedDesktopGroups @($DesktopGroupToGo) -IncludedSmartAccessFilterEnabled $True -IncludedSmartAccessTags @() -IncludedUserFilterEnabled $True -Name $strAGAccess -AdminAddress $strAdminDDC
        $strDesktopGroupUid = (Get-BrokerDesktopGroup -Name $DesktopGroupToGo).Uid
	    $strWeekdays = "UID${strDesktopGroupUid}Weekdays"
	    $strWeekend = "UID${strDesktopGroupUid}Weekend"
        New-BrokerPowerTimeScheme -DaysOfWeek 'Weekdays' -DesktopGroupUid $strDesktopGroupUid -DisplayName 'Weekdays' -Name $strWeekdays -PeakHours @($False,$False,$False,$False,$False,$False,$False,$True,$True,$True,$True,$True,$True,$True,$True,$True,$True,$True,$False,$False,$False,$False,$False,$False) -PoolSize @(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0) -AdminAddress $strAdminDDC
        New-BrokerPowerTimeScheme -DaysOfWeek 'Weekend' -DesktopGroupUid $strDesktopGroupUid -DisplayName 'Weekend' -Name $strWeekend -PeakHours @($False,$False,$False,$False,$False,$False,$False,$True,$True,$True,$True,$True,$True,$True,$True,$True,$True,$True,$False,$False,$False,$False,$False,$False) -PoolSize @(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0) -AdminAddress $strAdminDDC
    }
}
$strDesktopGroupUid = (Get-BrokerDesktopGroup -Name $DesktopGroupToGo -AdminAddress $strAdminDDC).Uid

$DomainNetBIOS = $env:USERDOMAIN
if($strCatalogUid)
{
    Add-Log -Path $LogFile -Value 'Added VM to catalog.'
    New-BrokerMachine -CatalogUid $strCatalogUid -MachineName "$DomainNetBIOS\$VDINameToGo" -HostedMachineId "$VMuuid" -HypervisorConnectionUid 1 -AdminAddress $strAdminDDC
}
else
{
    Add-Log -Path $LogFile -Value 'Catalog not found, can not add VM into machine catalog' -Type Warning
}

if($strDesktopGroupUid)
{
    Add-Log -Path $LogFile -Value 'Assigned VM to desktop group.'
    Add-BrokerMachine -MachineName "$DomainNetBIOS\$VDINameToGo" -DesktopGroup "$DesktopGroupToGo" -AdminAddress $strAdminDDC
    if($User)
    {
        Add-Log -Path $LogFile -Value 'Assigned user to the VM.'
        Add-BrokerUser "$DomainNetBIOS\$User" -PrivateDesktop "$DomainNetBIOS\$VDINameToGo" -AdminAddress $strAdminDDC
    }
}
else
{
    Add-Log -Path $LogFile -Value 'Desktop group not found, can not publish VDI' -Type Warning
}

Add-Log -Path $LogFile -Value "Environment set to $Environment."
if($Environment -ieq 'IM')
{
	Add-Log -Path $LogFile -Value "Add $VDINameToGo into group $IMGroup"
    while(!(Get-ADComputer $VDINameToGo))
    {
        Start-Sleep -Seconds 15
    }
	Get-ADComputer $VDINameToGo | Add-ADPrincipalGroupMembership -MemberOf $IMGroup
}
else
{
    if(!($BizGroup -imatch '^\s*$') -and $User){
        Add-Log -Path $LogFile -Value 'Biz group specified.'
        Add-Log -Path $LogFile -Value "Add $User into group $BizGroup"
        Add-ADGroupMember -Identity $BizGroup -Members $User -ErrorAction:SilentlyContinue
        if(!$?)
        {
            Add-Log -Path $LogFile -Value "Failed to add $User to $BizGroup, cause:" -Type Error
            Add-Log -Path $LogFile -Value $Error[0] -Type Error
        }
    }
}

Add-Log -Path $LogFile -Value 'All job done, about to quit.'
Quit
