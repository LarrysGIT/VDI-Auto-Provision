
$VDIs = 'XXVA001', 'XXVA002'

$Users = ''
$Environment = 'Test'
$Cluster = 'VMwareCluster1'

$AdvanceOptions = '-NumCPU 2 -IgnoreVDINameAvailable'

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

foreach($VDI in $VDIs){
    $PowershellInstances += Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @($(
        '-File', '..\VDI.Build.Core.ps1',
        "-VIServer", "$vCenter",
        "-User", "`"$User`"",
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
        "-VDIName $VDI",
        "$AdvanceOptions"
    ) | ?{$_}) -Wait -RedirectStandardOutput 'x.txt'
}
