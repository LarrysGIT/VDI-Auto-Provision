
<#
    Version 1.00 [Author: Larry Song]
        Create this new script for VDI rapid clone.
    Version 1.01 [Author: Larry Song; Time: 2014-02-25]
        Bug fixed - Previous templates will be hosted randomly in the whole datacenter, fix to specified cluster.
    Version 1.02 [Author: Larry Song; Time: 2014-12-31]
        A rapid clone bug will set CPU to 1 and memory to 1GB, add codes to change spec when templates created.
        Add 2 parameters -NumCPU and -MemoryMB
#>

PARAM(
    [parameter(Mandatory=$true)]
    [string]$VIServer,

    [parameter()]
    [int]$VIPort = 8143,

    [parameter()]
    [System.Management.Automation.PSCredential]$VICredential,

    [parameter()]
    [System.Management.Automation.PSCredential]$ControllerCredential,

    [parameter(Mandatory=$true)]
    [string]$Controller,

    [parameter()]
    [int]$ControllerPort = 443,

    [parameter(Mandatory=$true)]
    [string]$MasterTemplate,

    [parameter(Mandatory=$true)]
    [string]$Cluster,

    [parameter(Mandatory=$true)]
    [string]$Datastore,

    [parameter(Mandatory=$true)]
    [string]$Container,

    [parameter(Mandatory=$true)]
    [string[]]$TemplateNames,

    [parameter(Mandatory=$true)]
    [int]$NumCPU,

    [parameter(Mandatory=$true)]
    [int]$MemoryMB,

    [parameter()]
    [switch]$PassThru,

    [parameter()]
    [string]$LogFile
)

function Quit{
    PARAM(
        [int]$ExitCode = 0
    )
    if($ExitCodeWriteTo){
        Add-Content -Path $ExitCodeWriteTo -Value "$PID`t$ExitCode"
    }
    Exit($ExitCode)
}

$ExitCode = 0

if(Test-Path -Path '.\VDI.Common.Function.ps1' -PathType Leaf){
    Write-Host 'Common functions loaded.' -ForegroundColor Yellow
    . '.\VDI.Common.Function.ps1'
}else{
    Write-Host "Common functions file not found, script can't continue." -ForegroundColor Red
    Quit -ExitCode 1
}

function Check-Prerequisite{
    if(!$VICredential){
        $Hash1 = HashConv -strHash (Get-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash1').Hash1
        $Hash2 = ConvertTo-SecureString -String (Get-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash2').Hash2
        $Script:VICredential = New-Object System.Management.Automation.PSCredential($Hash1, $Hash2)
    }
    if(!$VICredential){
        Add-Log -Path $LogFile -Value "Failed to create credentital for VI server."
        Quit
    }
    if(!$ControllerCredential){
        $HashSP1 = HashConv -strHash (Get-ItemProperty -Path 'HKCU:\Software\company\SharePoint' -Name 'Hash1').Hash1
        $HashSP2 = ConvertTo-SecureString -String (Get-ItemProperty -Path 'HKCU:\Software\company\SharePoint' -Name 'Hash2').Hash2
        $Script:ControllerCredential = New-Object System.Management.Automation.PSCredential($HashSP1, $HashSP2)
    }
    if(!$ControllerCredential){
        Add-Log -Path $LogFile -Value "Failed to create credentital for controller."
        Quit
    }
}

Check-Prerequisite

# $DomainSuffix = $env:USERDNSDOMAIN
$DomainSuffix = (Get-WmiObject -Class Win32_ComputerSystem).Domain
if($VIServer -notmatch '^\w+\..+'){
    Add-Log -Path $LogFile -Value "The VIServer name provided is not a FQDN, add current domain suffix."
    $VIServer += ".$DomainSuffix"
    Add-Log -Path $LogFile -Value "VIServer name set to $VIServer"
}

### Avoid cert check.
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$True}

### Define web API connection string.
$strVIConnection = "https://${VIServer}:${VIPort}/kamino/public/api?wsdl"
Add-Log -Path $LogFile -Value "Define web API as $strVIConnection."

### Connect to web API.
$objConnection = New-WebServiceProxy -Uri $strVIConnection -Credential $VICredential
if(!$?){
    Add-Log -Path $LogFile -Value "Connect to web API failed, cause:" -Type Error
    Add-Log -Path $LogFile -Value $Error[0] -Type Error
    Add-Log -Path $LogFile -Value "No need to continue, script quit."
    Quit
}

$Namespace = $objConnection.GetType().Namespace
Add-Log -Path $LogFile -Value "Namespace: [$Namespace]"

### Create request specification.
Add-Log -Path $LogFile -Value "Create and fill request specification object."
$requestSpec = New-Object "${Namespace}.requestSpec"
if(!$?){
    Add-Log -Path $LogFile -Value "request spec failed to create, cause:" -Type Error
    Add-Log -Path $LogFile -Value $Error[0] -Type Error
    Add-Log -Path $LogFile -Value "No need to continue, script quit."
}
Add-Log -Path $LogFile -Value "Request specification object created, for following other specifications, will not check again."
$requestSpec.serviceUrl = "https://$VIServer/sdk"
$requestSpec.vcUser = $VICredential.UserName
$requestSpec.vcPassword = $VICredential.GetNetworkCredential().Password

### Get id of mutiple objects.
$MasterTemplateId = $objConnection.getMoref($MasterTemplate, "VirtualMachine", $requestSpec)
$ClusterId = $objConnection.getMoref($Cluster, "ClusterComputeResource", $requestSpec)
#$DatacenterId = $objConnection.getMoref($Datacenter, "Datacenter", $requestSpec)
$ContainerId = $objConnection.getMoref($Container, "Folder", $requestSpec)
$DatastoreId = $objConnection.getMoref($Datastore, "Datastore", $requestSpec)
Add-Log -Path $LogFile -Value "Template Id: [$MasterTemplateId]; Cluster id: [$ClusterId]; Folder id: [$ContainerId]; DS id: [$DatastoreId]"
if(!($MasterTemplateId -and $ClusterId -and $ContainerId -and $DatastoreId)){
    Add-Log -Path $LogFile -Value "One or more IDs are blank, can't continue." -Type Error
    Quit
}

### Create controller specification.
Add-Log -Path $LogFile -Value "Create and fill controller specification object."
$controllerSpec = New-Object "${Namespace}.controllerSpec"
$controllerSpec.username = $ControllerCredential.UserName
$controllerSpec.password = $ControllerCredential.GetNetworkCredential().Password
$controllerSpec.ipAddress = $Controller
$controllerSpec.ssl = $True
$controllerSpec.port = $ControllerPort

### Create VM specification.
$vmSpec = New-Object "${namespace}.vmSpec"
$vmSpec.powerOn = $false

$clones = @()
### Create clone spec entry.
Add-Log -Path $LogFile -Value "Create and fill clones array with cloneSpecEntry objects."
$TemplateNames | %{
    $cloneSpecEntry = New-Object "${namespace}.cloneSpecEntry"
    $cloneSpecEntry.key = $_
    $cloneSpecEntry.Value = $vmSpec
    $clones += $cloneSpecEntry
}

### Set template files destination controller and datastore.
Add-Log -Path $LogFile -Value "Set controller spec to template files."
$objFiles = $objConnection.getVmFiles($MasterTemplateId, $requestSpec)
if(!$objFiles){
    Add-Log -Path $LogFile -Value "Nothing captured with template ID $MasterTemplateId, can't continue." -Type Error
    Quit
}
$objFiles | %{
    $_.destDatastoreSpec.controller = $controllerSpec
    $_.destDatastoreSpec.mor = $DatastoreId
    $_.destDatastoreSpec.thinProvision = $true
}

### Create and fill clonespec object.
$cloneSpec = New-Object "${namespace}.cloneSpec"
$cloneSpec.files = $objFiles
$cloneSpec.clones = $clones
$cloneSpec.templateMoref = $MasterTemplateId
$cloneSpec.containerMoref = $ClusterId
$cloneSpec.destVmFolderMoref = $ContainerId
$cloneSpec.memMB = $MemoryMB
$cloneSpec.memMBSpecified = $true
$cloneSpec.numberCPU = $NumCPU
$cloneSpec.numberCPUSpecified = $true

### Pass clonespec to requestSpec object.
$requestSpec.CloneSpec = $cloneSpec
$objTask = $objConnection.CreateClones($requestSpec)
Add-Log -Path $LogFile -Value "Submited requests: $objTask"
if($PassThru){
    return $objTask -ireplace ':', '-'
}
