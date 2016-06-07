
<#
    Version 1.00 [Author: Larry Song; Time: 2014-11-12]
        Add version control
    Version 1.01 [Author: Larry Song; Time: 2014-12-09]
        Upload FIL_Capacity.js to site
    Version 1.02 [Author: Larry Song; Time: 2014-12-24]
        Reconstruct all sharepoint scripts due to bad desgin.
    Version 1.02 [Author: Larry Song; Time: 2015-01-05]
        Invoke rebuild export function suffix should be 1, fixed.
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means read result file failed.
    second   bit set 1, means SharePoint snapin failed adding.
    third    bit set 1, means retrieve web site failed.
    fourth   bit set 1, means retrieve list failed.
    fifth    bit set 1, means delete old item failed.
    sixth    bit set 1, means add new item failed.
#>

PARAM(
    [switch]$RemoveListItemsOnly
)

$strLogFile = "$LocalDes\$strDate\Import_Vesbose.log"

while($true)
{
    if((Test-Path -Path @($VDI_ImportTags.Values) -PathType Leaf) -notcontains $false)
    {
        Add-Log -Path $strLogFile -Value 'Found enough tag files, export request after timeline again'
        & '.\VDI.Build.Export.ps1' -ListName $VDI_Lists_Export.VDI_Build.List -KeyProperty $VDI_Lists_Export.VDI_Build.KeyProperty -LeftKeyProperty $VDI_Lists_Export.VDI_Build.LeftKeyProperty -Suffix 1 -RemoveListItemsAlso
        if($LASTEXITCODE)
        {
            Add-Log -Path $strLogFile -Value "Abnormal build export exit code captured: [$LASTEXITCODE]"
        }
        & '.\VDI.Rebuild.Export.ps1' -ListName $VDI_Lists_Export.VDI_Rebuild.List -KeyProperty $VDI_Lists_Export.VDI_Rebuild.KeyProperty -LeftKeyProperty $VDI_Lists_Export.VDI_Rebuild.LeftKeyProperty -Suffix 1 -RemoveListItemsAlso
        if($LASTEXITCODE)
        {
            Add-Log -Path $strLogFile -Value "Abnormal rebuild export exit code captured: [$LASTEXITCODE]"
        }
        break
    }
    Add-Log -Path $strLogFile -Value 'Did not get enough job tags, sleep 0.5 hour to detect again'
    Start-Sleep -Seconds 1800
}

$ExitCode = 0

Add-Log -Path $strLogFile -Value 'Start adding SharePoint snapin and get SP web.'

do{
    Add-PSSnapin 'Microsoft.SharePoint.PowerShell' -ErrorAction:SilentlyContinue
    $objWeb = Get-SPWeb $VDI_WebUrl -ErrorAction:SilentlyContinue
    if(!$? -or !$objWeb){
        Add-Log -Path $strLogFile -Value 'Get specified web failed, cause:' -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    }else{
        break
    }
    Add-Log -Path $strLogFile -Value 'Sleep 10 minutes to try again'
    Remove-PSSnapin 'Microsoft.SharePoint.PowerShell'
    Start-Sleep -Seconds 600
}while($true)

<#
$objWeb = Get-SPWeb $VDI_WebUrl
if(!$? -or $objWeb -eq $null){
    Add-Log -Path $strLogFile -Value 'Get specified web failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0004 # 0000 0000 0000 0100
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    exit($ExitCode)
}
#>

Add-Log -Path $strLogFile -Value 'Get specified web succeed.'

# generate JS script for capacity prediction and upload to SP
$VDI_ImportTags.Keys | %{
    $Capacity_JS_File = "$LocalDes\$strDate\Imports\FIL_Capacity.js"
}{
    $Capacity = $Null
    $Capacity = Get-Content -Path "$LocalDes\$strDate\Imports\${_}_FIL_Capacity.txt" -ErrorAction:SilentlyContinue
    Add-Log -Path $strLogFile -Value "Read capacity for [$_]: $Capacity" -Type Info
    Add-Content -Path $Capacity_JS_File -Value "var $($_ -ireplace '\s', $Null)_Capacity = $Capacity;"
}

if($JS_Upload_Lib)
{
    Add-Log -Path $strLogFile -Value 'Start upload Capacity JS script'
    $objList = $objWeb.Lists[$JS_Upload_Lib]
    $objList.RootFolder.Files.Add("$($objList.ParentWebUrl)/$JS_Upload_Lib/FIL_Capacity.js", [System.IO.File]::ReadAllBytes("$PWD\$strDate\Imports\FIL_Capacity.js"), $true)
    if(!$?)
    {
        Add-Log -Path $strLogFile -Value 'JS file upload failed' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    }
}

# generate user&vdi report and import to SP site
foreach($ListName in $VDI_Lists_Import.Keys)
{
    $ResultFile = "$LocalDes\$strDate\Imports\${ListName}_Report.CSV"

    Add-Log -Path $strLogFile -Value 'Script start'
    Add-Log -Path $strLogFile -Value "Log file: $strLogFile"
    Add-Log -Path $strLogFile -Value "Result file: $ResultFile"

    Add-Log -Path $strLogFile -Value "Start generating report: [$ListName]"
    & '.\VDI.Generate.Results.ps1' -ListName $ListName -ResultFile $ResultFile -OU $VDI_Lists_Import.$ListName
    if($LASTEXITCODE)
    {
        Add-Log -Path $strLogFile -Value 'There is some errors during report generation, report will not be imported' -Type Warning
    }
    else
    {
        Add-Log -Path $strLogFile -Value 'Report generated succeed'
        Add-Log -Path $strLogFile -Value 'Start to import to web'
        Add-Log -Path $strLogFile -Value 'Try to get list'
        $objList = $objWeb.Lists[$ListName]
        if(!$? -or $objList -eq $null){
            Add-Log -Path $strLogFile -Value 'Get list from web failed, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
            Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
            exit($ExitCode)
        }
        Add-Log -Path $strLogFile -Value 'Get list from web succeed'
        $iTotal = $objList.ItemCount
        Add-Log -Path $strLogFile -Value "Items in list count: $iTotal"
        Add-Log -Path $strLogFile -Value 'Start to clean all items'
        while($iTotal){
            $Item = $null
            $Item = $objList.Items[--$iTotal]
            $Item.Delete()
            if(!$?)
            {
                Add-Log -Path $strLogFile -Value "Error occurred when delete old item $($Item['Name']) $($Item['Alias']), cause:" -Type Warning
                Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
                $ExitCode = $ExitCode -bor 0x0010 # 0000 0000 0001 0000
            }
        }

        Add-Log -Path $strLogFile -Value 'Sleep 60 seconds to refresh list items from web'
        Start-Sleep -Seconds 60
        $objList = $objWeb.Lists[$ListName]
        $iTotal = $objList.ItemCount
        Add-Log -Path $strLogFile -Value "Items refreshed, count: $iTotal"
        Add-Log -Path $strLogFile -Value 'List cleaned, now to import new data'

        Add-Log -Path $strLogFile -Value "Load CSV data from: [$ResultFile]"
        $RawData = $null
        $RawData = Import-Csv -Path $ResultFile -ErrorAction:SilentlyContinue
        if(!$?){
            Add-Log -Path $strLogFile -Value 'Read raw data file failed, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            $ExitCode = $ExitCode -bor 0x0001 # 0000 0000 0000 0001
            Add-Log -Path $strLogFile -Value "Script with code $ExitCode"
        }else{
            $RawData = @($RawData)
            Add-Log -Path $strLogFile -Value "Read raw data succeed, count: [$($RawData.Count)]"
        }

        $RawData | %{
            $Item = $null
            $Item = $objList.AddItem()
            $Item['Name'] = $_.'Name'
            $Item['Alias'] = $_.'Alias'
            $Item['Disabled?'] = $_.'Disabled?'
            $Item['VIP?'] = $_.'VIP?'
            $Item['OU'] = $_.'OU'
            $Item['XX POD1'] = $_.'XX POD1'
            $Item['XXPOD1'] = $_.'XX POD1'

            $Item['XX POD2'] = $_.'XX POD2'
            $Item['XXPOD2'] = $_.'XX POD2'

            $Item['YY POD1'] = $_.'YY POD1'
            $Item['YYPOD1'] = $_.'YY POD1'

            $Item['YY POD2'] = $_.'YY POD2'
            $Item['YYPOD2'] = $_.'YY POD2'
            $Item['Last Check'] = $_.'Last Check'
            $Item.Update()
            if(!$?){
                Add-Log -Path $strLogFile -Value "Error occurred when update item $($_.'Name') $($_.'Alias'), cause:" -Type Warning
                Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
                $ExitCode = $ExitCode -bor 0x0020 # 0000 0000 0010 0000
            }
        }
    }
}

Add-Log -Path $strLogFile -Value "Exit code is: [$ExitCode]"

exit($ExitCode)
