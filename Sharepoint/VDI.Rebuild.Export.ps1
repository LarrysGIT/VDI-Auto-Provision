
<#
    Version 1.01
    Version 1.02 [Author: Larry Song; Time: 2014-04-01]
        Add code to identity user's account not exists in AD.
    Version 1.03 [Author: Larry Song; Time: 2014-12-24]
        Reconstruct all sharepoint scripts due to bad desgin.
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means local folder created failed.
    second   bit set 1, means SharePoint snapin adding failed.
    third    bit set 1, means web geting failed.
    fourth   bit set 1, means list geting failed.
    fifth    bit set 1, means export CSV failed.
    sixth    bit set 1, means set flag file failed.
    seventh  bit set 1, means unknown argument.
#>

PARAM(
    [string]$ListName,
    [string]$KeyProperty,
    [string]$LeftKeyProperty,
    [int]$Suffix,
    [switch]$RemoveListItemsAlso
)

$strLogFile = "$LocalDes\$strDate\${ListName}_Export_Vesbose_${Suffix}.log"
$RawFile = "$LocalDes\$strDate\Exports\${ListName}_Export_Raw_$Suffix.CSV"
$JobsLeftFile = "$LocalDes\$strDate\Exports\${ListName}_Export_Left_${Suffix}.CSV"

if(!$ListName)
{
    Add-Log -Path $strLogFile -Value "List name is empty, function exit: [Suffix:$Suffix]"
    exit
}

$ExitCode = 0
Add-Log -Path $strLogFile -Value 'VDI Build Export Script start'

Remove-Item -Path $RawFile -Confirm:$false -Force -ErrorAction:SilentlyContinue
Add-Log -Path $strLogFile -Value 'Start adding SharePoint snapin and get SP web'

do{
    Add-PSSnapin 'Microsoft.SharePoint.PowerShell' -ErrorAction:SilentlyContinue
    $objWeb = Get-SPWeb $VDI_WebUrl -ErrorAction:SilentlyContinue
    if(!$? -or !$objWeb)
    {
        Add-Log -Path $strLogFile -Value 'Get specified web failed, cause:' -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    }
    else
    {
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

Add-Log -Path $strLogFile -Value 'Get specified web succeeded'
Add-Log -Path $strLogFile -Value 'Try to get list'
$objList = $objWeb.Lists[$ListName]
if(!$? -or $objList -eq $null){
    Add-Log -Path $strLogFile -Value 'Get list from web failed' -Type Error
    $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    exit($ExitCode)
}

Add-Log -Path $strLogFile -Value 'Get list from web succeeded'
Add-Log -Path $strLogFile -Value "Items in list count: $($objList.ItemCount)"
Add-Log -Path $strLogFile -Value 'Start to export items'

$objRaw = New-Object PSObject
$KeyProperty, 'Modified By', 'Modified', 'Created By', 'Created' | %{
    $objRaw | Add-Member -MemberType NoteProperty -Name $_ -Value $null -Force
}

$JobsLeft = @()
$objJob = New-Object PSObject -Property @{ID = $null; User = $null; $KeyProperty = $null; Type = $null; Exception = $null; CreatedBy = $null}

$RawData = @()
foreach($Item in $objList.Items){
    if($Item[$KeyProperty]){
        $RawData += $objRaw.PSObject.Copy()
        $RawData[-1].$KeyProperty = $Item[$KeyProperty]
        $RawData[-1].'Modified By' = $Item['Modified By'] -ireplace '.*?#', ''
        $RawData[-1].'Modified' = ($Item['Modified']).ToString()
        $RawData[-1].'Created By' = $Item['Created By'] -ireplace '.*?#', ''
        $RawData[-1].'Created' = ($Item['Created']).ToString()
    }else{
        $JobsLeft += $objJob.PSObject.Copy()
        $JobsLeft[-1].'Exception' = "[$KeyProperty] is null"
        $JobsLeft[-1].'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
    }
}

if($RemoveListItemsAlso){
    $iTotal = $objList.ItemCount
    Add-Log -Path $strLogFile -Value "Start remove items in [$ListName], count: [$iTotal]"
    while($iTotal){
        $iTotal--
        $Item = $objList.Items[$iTotal]
        $Item.Delete()
        if(!$?){
            Add-Log -Path $strLogFile -Value 'Error occurred when removing item, cause:' -Type Warning
            Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
        }
    }
}

Add-Log -Path $strLogFile -Value 'All items exported'

if($RawData.Count -eq 0){
    Add-Log -Path $strLogFile -Value 'No items with key property, add a blank'
    $RawData = $objRaw.PSObject.Copy()
}

Add-Log -Path $strLogFile -Value 'Start exporting to CSV'
$RawData | Export-Csv -Path $RawFile -Encoding Unicode -NoTypeInformation
if(!$?){
    Add-Log -Path $strLogFile -Value 'Exporting CSV failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0010 # 0000 0000 0001 0000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    exit($ExitCode)
}

if($JobsLeft.Count -eq 0){
    Add-Log -Path $strLogFile -Value 'No left items, add a blank'
    $JobsLeft = $objJob.PSObject.Copy()
}
$JobsLeft | Export-Csv -Path $JobsLeftFile -Encoding Unicode -NoTypeInformation

Add-Log -Path $strLogFile -Value "Exit code is: [$ExitCode]"

exit($ExitCode)
