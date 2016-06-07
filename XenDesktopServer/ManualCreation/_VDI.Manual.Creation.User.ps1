
$Users = 'larrysong'
$Environment = "Prod"
$Cluster = "VMwareCluster2"

$AdvanceOptions = "-VDIName XXVBlarrysong -IgnoreVDINameAvailable"

######################################
Set-Location (Get-Item ($MyInvocation.MyCommand.Definition)).DirectoryName
if(Test-Path -Path '_Configuration.ps1')
{
    . '.\_Configuration.ps1'
}
else
{
    . '..\_Configuration.ps1'
}
Define-CommonVariables
Define-VDIBuildVariables

foreach($User in $Users){
    $PowershellInstances += Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @($(
        '-File', '..\VDI.Build.Core.ps1',
        "-VIServer", "$vCenter",
        "-User", "$User",
        "-Prefix", "$Prefix",
        "-DesktopGroup", "$($Clusters[$Cluster][$Environment]["Group"])",
        "-DDCList", "`"$DDCList`"",
        "-SupportGroup", "$SupportGroup",
        "-Environment", "$Environment",
        "-ClusterToGo", "$Cluster",
        "-IMGroup", "$IMADGroup",
        "-Folder", "$($Clusters[$Cluster][$Environment]["Folder"])",
        "-TemplateFolder", "$($Clusters[$Cluster][$Environment]['POOL'])",
        "-WorkingDirectory", "..\",
        "$AdvanceOptions"
    ) | ?{$_}) -Wait
}
