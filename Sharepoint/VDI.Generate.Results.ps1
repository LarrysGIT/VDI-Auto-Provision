
<#
    Version 1.00 [Author: Larry Song; Time: 2014-09-09]
        Add version control
    Version 1.01 [Author: Larry Song; Time: 2014-09-09]
        Switch over to native AD query commands due to no 2008+ DC in YY, slow to query object from XX DC.
    Version 1.02 [Author: Larry Song; Time: 2014-11-12]
        Add codes to generate JavaScript contains capacitry variables
    Version 1.03 [Author: Larry Song; Time: 2014-11-25]
        Move capacity js maker to 'VDI starter.ps1'
    Version 1.04 [Author: Larry Song; Time: 2014-12-24]
        Reconstruct all sharepoint scripts due to bad desgin.
    Version 1.05 [Author: Larry Song; Time: 2014-12-31]
        Add codes due to VDI inventory requirement from Frank, report saved as CSV
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means local folder created failed.
    second   bit set 1, means read raw file failed.
    third    bit set 1, means VMware snapin failed adding.
    fourth   bit set 1, means AD module import failed.
    fifth    bit set 1, means vCenter connection failed.
    sixth    bit set 1, means VM names geting failed.
    seventh  bit set 1, means retrieve user objects from AD failed.
    eigthth  bit set 1, means can NOT get VDI names from AD for user.
    ninth    bit set 1, means trigger scheduled task on sharepoint server failed.
#>

<# Default properties already exists.
    DistinguishedName
    Enabled
    GivenName
    Name
    ObjectClass
    ObjectGUID
    samaccountname
    SID
    Surname
    UserPrincipalName
#>

PARAM(
    [string]$ListName,
    [string]$OU,
    [string]$ResultFile
)

# Define-VDIGenerateReport

$strLogFile = "$LocalDes\$strDate\${ListName}_Report_Log_Verbose.log"
$VDIInventoryReport = "$LocalDes\$strDate\${ListName}_VDI_Inventory.CSV"

$LDAPProperties_User = 'name', 'samaccountname', 'useraccountcontrol', 'extensionattribute3', 'canonicalname'
$LDAPProperties_Computer = 'name'

$ExitCode = 0
Add-Log -Path $strLogFile -Value 'VDI result generation script start'
Add-Log -Path $strLogFile -Value "OU of users: [$OU]"

$objRaw = New-Object PSObject
'Alias', 'Name', 'Disabled?', 'VIP?', 'OU', 'XX POD1', 'YY POD2', 'ZZ POD1', 'UU POD2', 'Last Check' | %{
    $objRaw | Add-Member -MemberType NoteProperty -Name $_ -Value $null -Force
}
$objReport = New-Object PSObject
'Alias', 'Name', 'Disabled?', 'VIP?', 'OU', 'Type', 'VDI(s)' | %{
    $objReport | Add-Member -MemberType NoteProperty -Name $_ -Value $null -Force
}

# Define for users
$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$OU")
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.SearchScope = 'Subtree'

Add-Log -Path $strLogFile -Value 'Read VM names from tag files'
$VDI_ImportTags.Keys | %{
    $VMs = New-Object System.Collections.Specialized.OrderedDictionary
    $RetrySleep = 1
}{
    Add-Log -Path $strLogFile -Value "Reading VDIs in: [$_]"
    $VMs.Add($_, $Null)
    $VMNames = Get-Content -Path $VDI_ImportTags.$_ -ErrorAction:SilentlyContinue
    if(!$?){
        Add-Log -Path $strLogFile -Value 'Get VMs names failed, cause:' -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        Add-Log -Path $strLogFile -Value "Try again in $RetrySleep seconds"
        Start-Sleep -Seconds $RetrySleep
        $VMNames = Get-Content -Path $VDI_ImportTags.$_ -ErrorAction:SilentlyContinue
        if(!$?){
            Add-Log -Path $strLogFile -Value 'Get VM  names failed again, cause:' -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            Add-Log -Path $strLogFile -Value "Set $_ VMs as null" -Type Error
            $VMs.$_ = $Null
            $ExitCode = $ExitCode -bor 0x0010 # 0000 0000 0001 0000
            return
        }
    }
    $VMs.$_ = $VMNames
    Add-Log -Path $strLogFile -Value "Get VM names succeed, count: $(($VMs.$_).Count)"
}

Add-Log -Path $strLogFile -Value 'Start to retrieve user objects from AD'

$LDAPProperties_User | %{
    $objSearcher.PropertiesToLoad.Add($_) | Out-Null
}

$objSearcher.Filter = '(objectCategory=User)'
$objUsers = $objSearcher.FindAll()
if(!$?){
    Add-Log -Path $strLogFile -Value 'Retrieve user objects failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0040 # 0000 0000 0100 0000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    exit($ExitCode)
}

$objUsers = $objUsers | ?{$_.Properties['samaccountname'][0] -imatch '^A\d{6}$'}
$objUsers = @($objUsers)

Add-Log -Path $strLogFile -Value 'User objects retrieving succeeded'
Add-Log -Path $strLogFile -Value "User objects count: $($objUsers.Count)"
Add-Log -Path $strLogFile -Value 'Start to processing user objects'

# Define for computers
$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$VDI_Computer_OU")
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.SearchScope = 'Subtree'

$LDAPProperties_Computer | %{
    $objSearcher.PropertiesToLoad.Add($_) | Out-Null
}

$objUsers | %{
    $UserCount = 0
    $RetrySleep = 5
    $ResultData = @()
    $ReportData = @()
}{
    $UserCount++
    $UserVDIsAD = $Null
    Add-Log -Path $strLogFile -Value ("****** Start process $UserCount/$($objUsers.Count) <$($_.Properties['samaccountname'][0])><$($_.Properties['name'][0])>")
    $objSearcher.Filter = "(&(objectCategory=Computer)(samaccountname=*V*$($_.Properties['samaccountname'][0] -ireplace '^A','')`$))"
    $UserVDIsAD = $objSearcher.FindAll() | %{$_.Properties['name'][0]}
    if(!$?){
        Add-Log -Path $strLogFile -Value 'Query AD computer failed, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
        Add-Log -Path $strLogFile -Value "Try to query again in $RetrySleep seconds."
        Start-Sleep -Seconds $RetrySleep
        $UserVDIsAD = $objSearcher.FindAll() | %{$_.Properties['name'][0]}
        if(!$?){
            Add-Log -Path $strLogFile -Value 'Query AD computer failed again, cause:' -Type Warning
            Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
            $ExitCode = $ExitCode -bor 0x0080 # 0000 0000 1000 0000
            Add-Log -Path $strLogFile -Value 'Move next'
            return
        }
    }
    $UserVDIsAD = @($UserVDIsAD)
    Add-Log -Path $strLogFile -Value "Get VDI from AD succeed, count: $($UserVDIsAD.Count)"
    $UserVDIsAD = @($UserVDIsAD | ?{$_ -imatch $ADComputerNameFilter})
    Add-Log -Path $strLogFile -Value "VDIs in AD name(s)(filtered): [$($UserVDIsAD -join '], [')]"
    $UserVDIsVCSum = @()
    $UserVDIsVC = $Null

    $ResultData += $objRaw.PSObject.Copy()
    $ReportData += $objReport.PSObject.Copy()
    $ResultData[-1].'Name' = $_.Properties['name'][0]
    $ReportData[-1].'Name' = $ResultData[-1].'Name'
    $ResultData[-1].'Alias' = $_.Properties['samaccountname'][0]
    $ReportData[-1].'Alias' = $ResultData[-1].'Alias'
    $ResultData[-1].'Disabled?' = if($_.Properties['useraccountcontrol'][0] -band 0x2){'True'}else{'False'}
    $ReportData[-1].'Disabled?' = $ResultData[-1].'Disabled?'
    $ResultData[-1].'VIP?' = if($_.Properties['extensionattribute3'] -imatch '^VIP$'){"True"}else{"False"}
    $ReportData[-1].'VIP?' = $ResultData[-1].'VIP?'
    $_.Properties['canonicalname'][0] -imatch 'FIL_Users/(.*?)/[^\$]*$' | Out-Null
    $ResultData[-1].'OU' = $matches[1]
    $ReportData[-1].'OU' = $ResultData[-1].'OU'
    foreach($Key in $VMs.Keys){
        if($VMs.$Key -ne $null){
            if($UserVDIsVC = $VMs.$Key -imatch "$($_.Properties['samaccountname'][0] -ireplace '^A', '')"){
                $ResultData[-1].$Key = 'True'
            }else{
                $ResultData[-1].$Key = 'False'
            }
        }else{
            Add-Log -Path $strLogFile -Value "VMs in $Key is null, keep old." -Type Warning
        }
        $UserVDIsVCSum += $UserVDIsVC.PSObject.Copy()
        $UserVDIsVC = $null
    }
    Add-Log -Path $strLogFile -Value "VDIs in POD name(s): [$($UserVDIsVCSum -join '], [')]"
    Add-Log -Path $strLogFile -Value "Verify user's VDIs in AD and POD"
    $UserVDIs = $UserVDIsVCSum + $UserVDIsAD
    $UserVDIs = $UserVDIs | ?{$_}
    $UserVDIs_Valid = @($UserVDIs | Group-Object | ?{$_.Count -eq 2} | %{$_.Name})
    $UserVDIs = @($UserVDIs | Group-Object | ?{$_.Count -eq 1} | %{$_.Name})
    $ReportData[-1].'VDI(s)' = $UserVDIs_Valid -join "`n"
    Add-Log -Path $strLogFile -Value "Combined VDIs in AD and POD, count: [$($UserVDIs.Count)]"
    if($UserVDIs.Count){
        Add-Log -Path $strLogFile -Value "Removed duplicate VDI names, left: [$($UserVDIs -join '][')]" -Type Warning
        Add-Log -Path $strLogFile -Value "Set user $($ResultData[-1].'Alias') last check to against!" -Type Warning
        $ResultData[-1].'Last Check' = "* VDIs in AD and POD against.`n"
        $ResultData[-1].'Last Check' += $UserVDIs -join ', '
    }else{
        $ResultData[-1].'Last Check' = $null
    }

    ### determine if user is IM or biz
    if($_.Properties['canonicalname'][0] -imatch 'IMS/')
    {
        $ReportData[-1].'Type' = 'IMS'
        $ResultData[-1] = $Null
    }
    else
    {
        $ReportData[-1].'Type' = 'Biz'
    }
}

Add-Log -Path $strLogFile -Value 'Start export results and report to CSV'
$Global:ResultData = $ResultData
$ResultData | ?{$_} | Export-Csv -Path $ResultFile -NoTypeInformation -Encoding Unicode
$ReportData | ?{$_} | Export-Csv -Path $VDIInventoryReport -Delimiter "`t" -NoTypeInformation -Encoding Unicode

Add-Log -Path $strLogFile -Value "Exit code is: [$ExitCode]"

exit($ExitCode)
