
<#
    Version 1.01
        Add version control.
        Export VDI available names to "Exchanges" folder.
    Version 1.02 [Author: Larry Song; Time: 2014-03-14]
        Invoke log compression script.
    Version 1.03 [Author: Larry Song; Time: 2014-05-23]
        Add template creation report to post VDI creation report.
    Version 1.04 [Author: Larry Song; Time: 2014-07-02]
        Add codes apply to new list "VDI Rebuild".
    Version 1.05 [Author: Larry Song; Time: 2014-07-07]
        Bug fixed, notification email sending without any items.
    Version 1.06 [Author: Larry Song; Time: 2014-11-25]
        Add maker for capacity js file
    Version 1.07 [Author: Larry Song; Time: 2014-12-02]
        Capacity JS file put in the worng code, fix it.
    Version 1.08 [Author: Larry Song; Time: 2014-12-11]
        Capacity JS file put in the worng code, fix it again.
    Version 1.09 [Author: Larry Song; Time: 2014-12-24]
        Reconstruct all sharepoint scripts due to bad desgin.
#>

Set-Location (Get-Item $MyInvocation.MyCommand.Definition).Directory

. '.\VDI.Common.Function.ps1'
. '.\_Configuration.ps1'

New-Item -Path "$LocalDes\$strDate\Exports" -ItemType 'Directory' -Confirm:$false -Force -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
New-Item -Path "$LocalDes\$strDate\Imports" -ItemType 'Directory' -Confirm:$false -Force -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
New-Item -Path "$LocalDes\$strDate\Exchanges" -ItemType 'Directory' -Confirm:$false -Force -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

$ExitCodes = @{}

if((Test-Path -Path @("$LocalDes\$strDate\Exports","$LocalDes\$strDate\Imports", "$LocalDes\$strDate\Exchanges") -PathType Container) -contains $false)
{
    # Add-Log -Path $strLogFile -Value 'Folders creation failed' -Type Error
    Write-Host 'Folders creation failed' -ForegroundColor Red
    exit
}
Add-Log -Path $strLogFile -Value '============== Starter script start'
Add-Log -Path $strLogFile -Value 'Start VDI pre-report script'
& '.\VDI.Prereport.ps1'
$ExitCodes.Add('VDI.Prereport.ps1', $LASTEXITCODE)

Add-Log -Path $strLogFile -Value 'Sleep 30 mins to start create VDI'
Start-Sleep -Seconds 1800

Add-Log -Path $strLogFile -Value 'First time actual export VDI build requests'
& '.\VDI.Build.Export.ps1' -ListName $VDI_Lists_Export.VDI_Build.List -KeyProperty $VDI_Lists_Export.VDI_Build.KeyProperty -LeftKeyProperty $VDI_Lists_Export.VDI_Build.LeftKeyProperty -Suffix 0 -RemoveListItemsAlso
$ExitCodes.Add('VDI.Build.Export.ps1', $LASTEXITCODE)

Add-Log -Path $strLogFile -Value 'First time actual export VDI rebuild requests'
& '.\VDI.Rebuild.Export.ps1' -ListName $VDI_Lists_Export.VDI_Rebuild.List -KeyProperty $VDI_Lists_Export.VDI_Rebuild.KeyProperty -LeftKeyProperty $VDI_Lists_Export.VDI_Rebuild.LeftKeyProperty -Suffix 0 -RemoveListItemsAlso
$ExitCodes.Add('VDI.Rebuild.Export.ps1', $LASTEXITCODE)

Add-Log -Path $strLogFile -Value 'Import script start'
& '.\VDI.Import.ps1'
$ExitCodes.Add('VDI.Import.ps1', $LASTEXITCODE)

Add-Log -Path $strLogFile -Value 'Report script start'
& '.\VDI.Email.Report.ps1'
$ExitCodes.Add('VDI.Report.ps1', $LASTEXITCODE)

Add-Log -Path $strLogFile -Value 'Start logs compression'
& '.\VDI.Compress.Logs.ps1'
$ExitCodes.Add('VDI.Compress.Logs.ps1', $LASTEXITCODE)

$ExitCodes.Values | %{$ExitCode = 0}{$ExitCode = $ExitCode -bor $_}
Add-Log -Path $strLogFile -Value "All scripts exit with combined exit codes: [$ExitCode]"
# $SPSite = Get-SPWeb $VDI_WebUrl
# $SPSite.RecycleBin.DeleteAll()

exit($ExitCode)
