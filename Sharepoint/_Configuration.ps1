
<#
    Version 1.01
        Add version control.
        Start creating VDI Rebuild codes.
        Optimize variable definitions and codes.
    Version 1.02 [Author: Larry Song; Time: 2014-09-09]
        Add a variable $VDI_Computer_OU, corresponding with new version of VDI.Generate.Results.ps1
    Version 1.03 [Author: Larry Song; Time: 2014-12-30]
        Reconstruct all sharepoint scripts due to bad desgin.
#>

$LocalDes = '.'
$Date = Get-Date
$strDate = $Date.ToString('yyyy-MM-dd')

$VDI_Prefix = 'YYV', 'XXV'
$AD_ComputerNameFilter = '^YYV|^XXV'
$VDI_AliasFormat = ''
$VDI_Rebuild_NamePattern = ''

$strLogFile = "$LocalDes\$strDate\_Starter.Verbose.log"

$VDI_Computer_OU = 'OU=company_Computers,DC=company,DC=com'
$VDI_Lists_Import = @{
        'XX' = 'OU=XX,OU=company_Users,DC=company,DC=com';
        'YY' = 'OU=YY,OU=company_Users,DC=company,DC=com';
        'ZZ' = 'OU=ZZ,OU=company_Users,DC=company,DC=com';
}

$VDI_Lists_Export = @{
    'VDI_Build'   = @{List = 'New Starter VDI'; KeyProperty = 'POD Changes'; LeftKeyProperty = 'User'};
    'VDI_Rebuild' = @{List = 'VDI Rebuild'    ; KeyProperty = 'VDI Name'   ; LeftKeyProperty = 'VDI Name'};
}

$VDI_ImportTags = @{
    'XX POD1' = "$LocalDes\$strDate\Imports\XX POD1.txt";
    'YY POD2' = "$LocalDes\$strDate\Imports\YY POD1.txt";
}

$VDI_WebUrl = 'http://sharepoint/sites/VDI'
$JS_Upload_Lib = 'Shared Documents'

$Email_CC = 'LarrySong@company.com'
#$Email_CC = 'LarrySong@company.com'    # Debugger
$Email_To = ''
#$Email_To = 'LarrySong@company.com'    # Debugger
$Email_From = "$($env:COMPUTERNAME)@company.com"   # No need to change, current computer name plus @company.com.
$Email_SMTPServer = 'smtp.company.com'    # SMTP server to contact with.
$Email_Subject = "[$strDate] VDI creation report."
$Email_Content = ''
