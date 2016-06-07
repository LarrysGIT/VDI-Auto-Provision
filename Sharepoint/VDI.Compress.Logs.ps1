
<#
    Version 1.01 [Author: Larry Song; Time: 2014-02-27]
        First build.
    Version 1.02 [Author: Larry Song; Time: 2014-07-06]
        Auto change folder name if there is exists
    Version 1.03 [Author: Larry Song; Time: 2014-07-08]
        Auto archive to "ArchivedLogs" folder instead
#>

Set-Location (Get-Item ($MyInvocation.MyCommand.Definition)).DirectoryName

$OldLogsFolder = "$LocalDes\ArchivedLogs"
if(!(Test-Path -Path $OldLogsFolder))
{
    New-Item -Name $OldLogsFolder -ItemType Directory
}

Get-ChildItem -Filter '*.zip' | Move-Item -Destination "$OldLogsFolder\" -Force

$strMonth = $Date.ToString('yyyy-MM')
$SecondsSince = [int64]($Date.ToUniversalTime() - [datetime]::Parse('1970-01-01')).TotalSeconds

$objShell = New-Object -ComObject Shell.Application
$arrZIPFile = [byte[]](80,75,5,6,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
Get-ChildItem | %{if($_.Mode -imatch '^d' -and $_.Name -imatch "^(20\d{2}-(?:0[1-9]|1[0-2]))-\d{2}"){$Matches[1]}} | Sort -Unique | %{
    if($_ -eq $strMonth){return}
    Write-Host $_ -ForegroundColor Yellow
    $strZIPFile = "$OldLogsFolder\${_}.ZIP"
    if(!(Test-Path -Path $strZIPFile -PathType Leaf)){
        Set-Content -Path $strZIPFile -Value $arrZIPFile -Encoding Byte
    }
    $strZIPFile = (Get-Item -Path $strZIPFile).FullName
    $objZIPPackage = $null
    $objZIPPackage = $objShell.NameSpace($strZIPFile)
    $strTemp = $_
    Get-ChildItem | ?{$_.Mode -imatch '^d' -and $_.Name -imatch "^${strTemp}-\d{2}.*"} | %{
        $_.Name
        if($objZIPPackage.ParseName($_.Name))
        {
            $_.MoveTo("$PWD\$($_.BaseName)_AUTO_RENAME_$SecondsSince$($_.Extension)")
            $objZIPPackage.MoveHere($_.FullName, 1024)
        }
        else
        {
            $objZIPPackage.MoveHere($_.FullName, 1024)
        }
        do{
            Start-Sleep -Seconds 2
        }while(Test-Path -Path $_.FullName)
    }
}
