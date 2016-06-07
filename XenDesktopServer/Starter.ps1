
<#
    Version 1.01
        Add version control.
    Version 1.02
        Combined Data.Report.ps1 with VDI.Build.ps1, removed Data.Report.ps1 and configurations from _Configuration.ps1
    Version 1.03 [Author: Larry Song; Time: 2014-03-13]
        Add log folder compression job.
    Version 1.04 [Author: Larry Song; Time: 2014-06-30]
        Invoke VDI.Rebuild.ps1 script for VDI rebuild job.
    Version 1.05 [Author: Larry Song; Time: 2014-11-03]
        Update exit code handler for VDI.Data.Report.ps1 capacity exit code
    Version 1.06 [Author: Larry Song; Time: 2014-11-10]
        Add variable $CapacityMonitoringEnable to control capacity monitoring
    Version 1.07 [Author: Larry Song; Time: 2015-01-16]
        Adjust new codes to compact for new VDI.Data.Report.ps1
    Version 1.08 [Author: Larry Song; Time: 2015-01-30]
        Remove function 'Add-Log' since already contains in common functions
    Version 1.09 [Author: Larry Song; Time: 2015-02-12]
        Invoke VDI.ReclaimStorage.ps1 asynchronously
    Version 1.10 [Author: Larry Song; Time: 2015-02-25]
        Cancel exit code for VDI.Data.Report.ps1
    Version 1.11 [Author: Larry Song; Time: 2015-07-08]
        Launch VDI.MoveBack.ps1 asynchronously
#>

Set-Location (Get-Item $MyInvocation.MyCommand.Definition).Directory

. '.\VDI.Common.Function.ps1'
. '.\_Configuration.ps1'
Define-CommonVariables
Define-StarterVariables

New-Item -Path "$LocalDes\$strDate" -ItemType 'Directory' -Confirm:$false -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

$ExitCodes = @{}

Add-Log -Path $strLogFile -Value 'Start space reclaim script asynchronously'
Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @('-File', 'VDI.ReclaimStorage.ps1')

Add-Log -Path $strLogFile -Value 'Start move back script asynchronously'
Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @('-File', 'VDI.MoveBack.ps1')

Add-Log -Path $strLogFile -Value 'Start VDI build script.'
$Main_Instance = Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @('-File', 'VDI.Build.ps1') -Wait -PassThru
Add-Log -Path $strLogFile -Value "VDI build script completed! exit code: $($Main_Instance.ExitCode)"
$ExitCodes.Add('VDI.Build.ps1', $Main_Instance.ExitCode)

Add-Log -Path $strLogFile -Value 'Start VDI rebuild script.'
$Main_Instance = Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @('-File', 'VDI.Rebuild.ps1') -Wait -PassThru
Add-Log -Path $strLogFile -Value "VDI build script completed! exit code: $($Main_Instance.ExitCode)"
$ExitCodes.Add('VDI.Rebuild.ps1', $Main_Instance.ExitCode)

Add-Log -Path $strLogFile -Value 'Start data report script.'
& '.\VDI.Data.Report.ps1'
Add-Log -Path $strLogFile -Value "Data report completed! exit code: $LASTEXITCODE"
# $ExitCodes.Add('VDI.Data.Report.ps1', $LASTEXITCODE)

$send = $false
$ExitCodes.Keys | %{
    if($ExitCodes.$_){
        $EmailSubject += "$($ExitCodes.$_) "
        $send = $true
        return
    }
}

if($send){
    $EmailContent = $EmailContent -ireplace '%PathToLog', $((Get-Item $strLogFile).DirectoryName -ireplace '(.)\:\\', "\\$($env:computername)\`$1`$\")
    Send-MailMessage -To $EmailTo -From $EmailFrom -SmtpServer $EmailSMTPServer -Subject $EmailSubject -BodyAsHtml $EmailContent
    if(!$?){
        Add-Log -Path $strLogFile -Value 'Email sent failed, cause:' -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    }
}

Add-Log -Path $strLogFile -Value 'Main job completed, start expension script.'
& '.\_Expension.Script.ps1'
Add-Log -Path $strLogFile -Value "Expension script completed! exit code: $LASTEXITCODE"

Add-Log -Path $strLogFile -Value 'Start log folder compression'
& '.\VDI.Compress.Logs.ps1'
