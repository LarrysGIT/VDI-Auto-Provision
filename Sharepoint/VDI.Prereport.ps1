
<#
    Version 1.01 [Author: Larry Song; Time: 2014-03-14]
        First build of pre-report.
    Version 1.02 [Author: Larry Song; Time: 2014-04-01]
        Add code to identity user's account not exists in AD.
    Version 1.03 [Author: Larry Song; Time: 2014-07-04]
        Bug fixed for invalid job email without creator.
    Version 1.04 [Author: Larry Song; Time: 2014-07-07]
        Bug fixed for notification email sending without any items in SP lists.
#>

$strLogFile = "$LocalDes\$strDate\Pre-report.log"

$Email_Subject = "VDI request summary before take action - $VDI_Prefix"

Add-Log -Path $strLogFile -Value 'Pre report script start'
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
    Add-Log -Path $strLogFile -Value "Sleep 1 minutes to try again."
    Remove-PSSnapin 'Microsoft.SharePoint.PowerShell'
    Start-Sleep -Seconds 60
}while($true)

Add-Log -Path $strLogFile -Value 'Get specified web succeeded.'
Add-Log -Path $strLogFile -Value 'Try to get list.'
#########################################
$objVDIBuildJob = New-Object PSObject -Property @{User = $null; Changes = $null; CreatedBy = $null; Comment = $null}
$objVDIRebuildJob = New-Object PSObject -Property @{VDIName = $null; CreatedBy = $null; Comment = $null}

$VDI_BuildJobsValid = @{}
$VDI_BuildJobsInvalid = @{}

$VDI_RebuildJobsValid = @{}
$VDI_RebuildJobsInvalid = @{}
$VDI_RequesterNames = @()

$EmailToGo = $false
$VDI_Lists_Export.Keys | %{
    $Key = $_
    $KeyProperty = $VDI_Lists_Export[$Key]['KeyProperty']
    $List = $VDI_Lists_Export[$Key]['List']
    Add-Log -Path $strLogFile -Value "Processing list: [$List]"
    $objList = $null
    $objList = $objWeb.Lists[$List]
    if(!$? -or $objList -eq $null)
    {
        Add-Log -Path $strLogFile -Value 'Get list from web failed.' -Type Error
        $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
        Add-Log -Path $strLogFile -Value "Script add with code $ExitCode"
        return
    }
    Add-Log -Path $strLogFile -Value 'Get list from web succeeded.'
    Add-Log -Path $strLogFile -Value "Items in list count: $($objList.ItemCount)"
    if($objList.ItemCount -eq 0)
    {
        Add-Log -Path $strLogFile -Value 'No items.'
        return
    }
    $EmailToGo = $true
    switch($Key)
    {
        'VDI_Build'
        {
            foreach($Item in $objList.Items){
                if($Item[$KeyProperty]){
                    $strAlias = $Item['Alias']
                    Add-Log -Path $strLogFile -Value "Processing [$strAlias]"
                    $strAlias = $strAlias.Split(',') | %{if($_){$_.Trim()}} | %{
                        if($_ -imatch $VDI_AliasFormat){$_}else{
                            $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                            $VDI_BuildJobsInvalid.$_.'User' = $_
                            $VDI_BuildJobsInvalid.$_.'Changes' = $Item[$KeyProperty]
                            $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                            $VDI_BuildJobsInvalid.$_.'Comment' = 'Invalid alias'
                        }
                    }
                    $strAlias = $strAlias | ?{$_}
                    Add-Log -Path $strLogFile -Value "Valid alias captured: [$($strAlias -join '], [')]"
                    $strAlias | %{
                        if(!(Get-ADExisting -SAMAccountName $_ -Type user)){
                            Add-Log -Path $strLogFile -Value "[$_] not found in AD." -Type Warning
                            $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                            $VDI_BuildJobsInvalid.$_.'User' = $_
                            $VDI_BuildJobsInvalid.$_.'Changes' = $Item[$KeyProperty]
                            $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                            $VDI_BuildJobsInvalid.$_.'Comment' = 'User not found in AD'
                            return
                        }
                        if($VDI_BuildJobsValid.Keys -icontains $_){
                            foreach($PODName in $VDI_ImportTags.Keys){
                                $strChangeExs = ([regex]::Match($VDI_BuildJobsValid.$_.'Changes', "[\S ]+${PODName}[\S ]*")).Value
                                $strChangeDup = ([regex]::Match($Item[$KeyProperty], "[\S ]+${PODName}[\S ]*")).Value
                                if($strChangeExs -or $strChangeDup){
                                    if($strChangeExs -and $strChangeDup){
                                    Add-Log -Path $strLogFile -Value "Duplicate POD change found: [$PODName]"
                                    Add-Log -Path $strLogFile -Value "Existing POD change: [$strChangeExs]"
                                    Add-Log -Path $strLogFile -Value "New POD change: [$strChangeDup]"
                                    $strChangeExsArray = $strChangeExs.Split(':')
                                    $strChangeDupArray = $strChangeDup.Split(':')
                                    $strFinalPODChange = @($strChangeExsArray[0])
                                    if($strChangeExsArray[1] -eq 'Force'){
                                        if($strChangeDupArray[1] -eq 'Force'){
                                            $strFinalPODChange += 'Force'
                                            if([int]$strChangeDupArray[2] -lt [int]$strChangeExsArray[2]){
                                                $strFinalPODChange += $strChangeDupArray[2]
                                                if($VDI_BuildJobsInvalid.Keys -icontains $_){
                                                    $VDI_BuildJobsInvalid.$_.'Changes' += "`n$strChangeExs"
                                                    $VDI_BuildJobsInvalid.$_.'CreatedBy' += "`n$($Item['Created By'] -ireplace '.*?#', '')"
                                                    $VDI_BuildJobsInvalid.$_.'Comment' += "`nVDI number against"
                                                }else{
                                                    $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                                                    $VDI_BuildJobsInvalid.$_.'User' = $_
                                                    $VDI_BuildJobsInvalid.$_.'Changes' = $strChangeExs
                                                    $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                                                    $VDI_BuildJobsInvalid.$_.'Comment' = 'VDI number against'
                                                }
                                            }else{
                                                if($VDI_BuildJobsInvalid.Keys -icontains $_){
                                                    $VDI_BuildJobsInvalid.$_.'Changes' += "`n$strChangeDup"
                                                    $VDI_BuildJobsInvalid.$_.'CreatedBy' += "`n$($Item['Created By'] -ireplace '.*?#', '')"
                                                    $VDI_BuildJobsInvalid.$_.'Comment' += "`nVDI number against"
                                                }else{
                                                    $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                                                    $VDI_BuildJobsInvalid.$_.'User' = $_
                                                    $VDI_BuildJobsInvalid.$_.'Changes' = $strChangeDup
                                                    $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                                                    $VDI_BuildJobsInvalid.$_.'Comment' = 'VDI number against'
                                                }
                                            }
                                        }else{
                                            if($VDI_BuildJobsInvalid.Keys -icontains $_){
                                                $VDI_BuildJobsInvalid.$_.'Changes' += "`n$strChangeExs"
                                                $VDI_BuildJobsInvalid.$_.'CreatedBy' += "`n$($Item['Created By'] -ireplace '.*?#', '')"
                                                $VDI_BuildJobsInvalid.$_.'Comment' += "`nForce tag against"
                                            }else{
                                                $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                                                $VDI_BuildJobsInvalid.$_.'User' = $_
                                                $VDI_BuildJobsInvalid.$_.'Changes' = $strChangeExs
                                                $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                                                $VDI_BuildJobsInvalid.$_.'Comment' = 'Force tag against'
                                            }
                                        }
                                    }else{
                                        if($strChangeDupArray[1] -eq 'Force'){
                                            if($VDI_BuildJobsInvalid.Keys -icontains $_){
                                                $VDI_BuildJobsInvalid.$_.'Changes' += "`n$strChangeDup"
                                                $VDI_BuildJobsInvalid.$_.'CreatedBy' += "`n$($Item['Created By'] -ireplace '.*?#', '')"
                                                $VDI_BuildJobsInvalid.$_.'Comment' += "`nForce tag against"
                                            }else{
                                                $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                                                $VDI_BuildJobsInvalid.$_.'User' = $_
                                                $VDI_BuildJobsInvalid.$_.'Changes' = $strChangeDup
                                                $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                                                $VDI_BuildJobsInvalid.$_.'Comment' = 'Force tag against'
                                            }
                                        }else{
                                            if($VDI_BuildJobsInvalid.Keys -icontains $_){
                                                $VDI_BuildJobsInvalid.$_.'Changes' += "`n$strChangeDup"
                                                $VDI_BuildJobsInvalid.$_.'CreatedBy' += "`n$($Item['Created By'] -ireplace '.*?#', '')"
                                                $VDI_BuildJobsInvalid.$_.'Comment' += "`nNormal request against"
                                            }else{
                                                $VDI_BuildJobsInvalid.Add($_, $objVDIBuildJob.PSObject.Copy())
                                                $VDI_BuildJobsInvalid.$_.'User' = $_
                                                $VDI_BuildJobsInvalid.$_.'Changes' = $strChangeDup
                                                $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                                                $VDI_BuildJobsInvalid.$_.'Comment' = 'Normal request against'
                                            }
                                        }
                                    }
                                    $strFinalPODChange = $strFinalPODChange -join ":"
                                    $VDI_BuildJobsValid.$_.'Changes' = $VDI_BuildJobsValid.$_.'Changes'.Replace($strChangeExs, $strFinalPODChange)
                                }else{
                                        if($strChangeDup){
                                            $VDI_BuildJobsValid.$_.'Changes' += "`n$strChangeDup"
                                        }
                                    }
                                }
                            }
                        }else{
                            $VDI_BuildJobsValid.Add($_, $objVDIBuildJob.PSObject.Copy())
                            $VDI_BuildJobsValid.$_.'User' = $_
                            $VDI_BuildJobsValid.$_.'Changes' = $Item[$KeyProperty]
                            $VDI_BuildJobsValid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                        }
                    }
                }else{
                    if($VDI_BuildJobsInvalid.Keys -icontains $Item['Alias']){
                        $VDI_BuildJobsInvalid.$_.'CreatedBy' += "`n$($Item['Created By'] -ireplace '.*?#', '')"
                        $VDI_BuildJobsInvalid.$_.'Comment' += "`nPOD Change is null"
                    }else{
                        $VDI_BuildJobsInvalid.Add($Item['Alias'], $objVDIBuildJob.PSObject.Copy())
                        $VDI_BuildJobsInvalid.$_.'User' = $Item['Alias']
                        $VDI_BuildJobsInvalid.$_.'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                        $VDI_BuildJobsInvalid.$_.'Comment' = 'POD Change is null'
                    }
                }
            }
        }
        'VDI_Rebuild'
        {
            foreach($Item in $objList.Items)
            {
                if($Item[$KeyProperty])
                {
                    Add-Log -Path $strLogFile -Value "Processing VDI Name: [$($Item[$KeyProperty])]"
                    if($Item[$KeyProperty] -imatch $VDI_Rebuild_NamePattern)
                    {
                        Add-Log -Path $strLogFile -Value "[$($Item[$KeyProperty])] matches pattern [$VDI_Rebuild_NamePattern]"
                        $VDI_RebuildJobsValid.Add($Item[$KeyProperty], $objVDIRebuildJob.PSObject.Copy())
                        $VDI_RebuildJobsValid.$($Item[$KeyProperty]).'VDIName' = $Item[$KeyProperty]
                        $VDI_RebuildJobsValid.$($Item[$KeyProperty]).'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                    }
                    else
                    {
                        Add-Log -Path $strLogFile -Value "[$($Item[$KeyProperty])] NOT matches pattern [$VDI_Rebuild_NamePattern]"
                        $VDI_RebuildJobsInvalid.Add($Item[$KeyProperty], $objVDIRebuildJob.PSObject.Copy())
                        $VDI_RebuildJobsInvalid.$($Item[$KeyProperty]).'VDIName' = $Item[$KeyProperty]
                        $VDI_RebuildJobsInvalid.$($Item[$KeyProperty]).'CreatedBy' = $Item['Created By'] -ireplace '.*?#', ''
                    }
                }
            }
        }
        Default
        {
            Add-Log -Path $strLogFile -Value "No any codes for list: [$List]"
        }
    }
}

#########################################
$HTMLHeader = @'
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 14 (filtered medium)"><style><!--
/* Font Definitions */
@font-face
	{font-family:SimSun;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:SimSun;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:SimSun;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"Lucida Console";
	panose-1:2 11 6 9 4 5 4 2 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:blue;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:purple;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal-compose;
	font-family:"Calibri","sans-serif";
	color:windowtext;}
.MsoChpDefault
	{mso-style-type:export-only;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;}
div.WordSection1
	{page:WordSection1;}
--></style><!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]--><!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]-->
</head>
	<body lang=EN-US link=blue vlink=purple><div class=WordSection1>

'@
$HTMLTableHeader = @'
        <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>

'@
$HTMLRow = @'
			<tr style='height:10.75pt'>
				<td valign=top style='border:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:10.75pt'>
					<p class=MsoNormal><span style='font-size:12.0pt;font-family:"Lucida Console"'>%Cell_0%<o:p></o:p></span></p>
				</td>
				<td valign=top style='border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt;height:10.75pt'>
					<p class=MsoNormal><span style='font-size:12.0pt;font-family:"Lucida Console"'>%Cell_1%<o:p></o:p></span></p>
				</td>
				<td valign=top style='border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt;height:10.75pt'>
					<p class=MsoNormal><span style='font-size:12.0pt;font-family:"Lucida Console"'>%Cell_2%<o:p></o:p></span></p>
				</td>
				<td valign=top style='border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt;height:10.75pt'>
					<p class=MsoNormal><span style='font-size:12.0pt;font-family:"Lucida Console"'>%Cell_3%<o:p></o:p></span></p>
				</td>
			</tr>

'@
$HTMLTableTail = @'
        </table>

'@
$HTMLTail = @'
		<p class=MsoNormal><span style='font-size:12.0pt;font-family:"Lucida Console"'><o:p>&nbsp;</o:p></span></p></div>
	</body></html>

'@

$HTMLBody = $HTMLHeader
######################################### VDI accepted
if($VDI_BuildJobsValid.Count)
{
    $HTMLBody += "<b><span style='color:red;font-size:20.0pt'>Requests accepted:<o:p></o:p></span></b>"
    $HTMLBody += $HTMLTableHeader
    $HTMLBody += $HTMLRow.PSObject.Copy()
    $HTMLBody = $HTMLBody -ireplace '%Cell_0%', "<span style='color:red'>User<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_1%', "<span style='color:red'>Changes<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_2%', "<span style='color:red'>Created by<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_3%', "<span style='color:red'>Comment<o:p></o:p></span>"
    $VDI_BuildJobsValid.Keys | Sort | %{
        $VDI_RequesterNames += $VDI_BuildJobsValid.$_.'Createdby'
        $HTMLBody += $HTMLRow.PSObject.Copy()
        $HTMLBody = $HTMLBody -ireplace '%Cell_0%', $VDI_BuildJobsValid.$_.'User'
        $HTMLBody = $HTMLBody -ireplace '%Cell_1%', $($VDI_BuildJobsValid.$_.'Changes' -replace "`n", "<br>")
        $HTMLBody = $HTMLBody -ireplace '%Cell_2%', $($VDI_BuildJobsValid.$_.'Createdby' -replace "`n", "<br>")
        $HTMLBody = $HTMLBody -ireplace '%Cell_3%', $($VDI_BuildJobsValid.$_.'Comment' -replace "`n", "<br>")
    }
    $HTMLBody += $HTMLTableTail
}
######################################### VDI NOT accepted
if($VDI_BuildJobsInvalid.Count)
{
    $HTMLBody += "<br>"
    $HTMLBody += "<b><span style='color:red;font-size:20.0pt'>Requests NOT accepted:<o:p></o:p></span></b>"
    $HTMLBody += $HTMLTableHeader
    $HTMLBody += $HTMLRow.PSObject.Copy()
    $HTMLBody = $HTMLBody -ireplace '%Cell_0%', "<span style='color:red'>User<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_1%', "<span style='color:red'>Changes<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_2%', "<span style='color:red'>Created by<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_3%', "<span style='color:red'>Comment<o:p></o:p></span>"
    $VDI_BuildJobsInvalid.Keys | Sort | %{
        $VDI_RequesterNames += $VDI_BuildJobsInvalid.$_.'Createdby'
        $HTMLBody += $HTMLRow.PSObject.Copy()
        $HTMLBody = $HTMLBody -ireplace '%Cell_0%', $VDI_BuildJobsInvalid.$_.'User'
        $HTMLBody = $HTMLBody -ireplace '%Cell_1%', $($VDI_BuildJobsInvalid.$_.'Changes' -replace "`n", "<br>")
        $HTMLBody = $HTMLBody -ireplace '%Cell_2%', $($VDI_BuildJobsInvalid.$_.'Createdby' -replace "`n", "<br>")
        $HTMLBody = $HTMLBody -ireplace '%Cell_3%', $($VDI_BuildJobsInvalid.$_.'Comment' -replace "`n", "<br>")
    }
    $HTMLBody += $HTMLTableTail
}
######################################### VDI rebuild accepted
if($VDI_RebuildJobsValid.Count)
{
    $HTMLBody += "<br>"
    $HTMLBody += "<b><span style='color:red;font-size:20.0pt'>Rebuild requests accepted:<o:p></o:p></span></b>"
    $HTMLBody += $HTMLTableHeader
    $HTMLBody += $HTMLRow.PSObject.Copy()
    $HTMLBody = $HTMLBody -ireplace '%Cell_0%', "<span style='color:red'>VDI Name<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_1%', "<span style='color:red'>Created by<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_2%', "<span style='color:red'>Comment<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_3%', $null
    $VDI_RebuildJobsValid.Keys | Sort | %{
        $VDI_RequesterNames += $VDI_RebuildJobsValid.$_.'Createdby'
        $HTMLBody += $HTMLRow.PSObject.Copy()
        $HTMLBody = $HTMLBody -ireplace '%Cell_0%', $VDI_RebuildJobsValid.$_.'VDIName'
        $HTMLBody = $HTMLBody -ireplace '%Cell_1%', $VDI_RebuildJobsValid.$_.'Createdby'
        $HTMLBody = $HTMLBody -ireplace '%Cell_2%', $VDI_RebuildJobsValid.$_.'Comment'
        $HTMLBody = $HTMLBody -ireplace '%Cell_3%', $null
    }
    $HTMLBody += $HTMLTableTail
}
######################################### VDI rebuild not accepted
if($VDI_RebuildJobsInvalid.Count)
{
    $HTMLBody += "<br>"
    $HTMLBody += "<b><span style='color:red;font-size:20.0pt'>Rebuild requests NOT accepted:<o:p></o:p></span></b>"
    $HTMLBody += $HTMLTableHeader
    $HTMLBody += $HTMLRow.PSObject.Copy()
    $HTMLBody = $HTMLBody -ireplace '%Cell_0%', "<span style='color:red'>VDI Name<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_1%', "<span style='color:red'>Created by<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_2%', "<span style='color:red'>Comment<o:p></o:p></span>"
    $HTMLBody = $HTMLBody -ireplace '%Cell_3%', $null
    $VDI_RebuildJobsInvalid.Keys | Sort | %{
        $VDI_RequesterNames += $VDI_RebuildJobsInvalid.$_.'Createdby'
        $HTMLBody += $HTMLRow.PSObject.Copy()
        $HTMLBody = $HTMLBody -ireplace '%Cell_0%', $VDI_RebuildJobsInvalid.$_.'VDIName'
        $HTMLBody = $HTMLBody -ireplace '%Cell_1%', $VDI_RebuildJobsInvalid.$_.'Createdby'
        $HTMLBody = $HTMLBody -ireplace '%Cell_2%', $VDI_RebuildJobsInvalid.$_.'Comment'
        $HTMLBody = $HTMLBody -ireplace '%Cell_3%', $null
    }
    $HTMLBody += $HTMLTableTail
}
#########################################
$HTMLBody += $HTMLTail

Add-Log -Path $strLogFile -Value "RequesterNames: [$($VDI_RequesterNames -join '], [')]"

$VDI_RequesterNames = @($VDI_RequesterNames | ?{$_} | Sort-Object -Unique)
Import-Module active*
$Script:Email_To += for($i = 0; $i -lt $VDI_RequesterNames.Count; $i++){
    Add-Log -Path $strLogFile -Value "Capture requester from AD: $($VDI_RequesterNames[$i])"
    (Get-ADUser -Filter "DisplayName -eq '$($VDI_RequesterNames[$i])'" | %{
        if($_.SamAccountName -imatch '(a\d{6})'){return $Matches[1]}
    }) | Sort-Object -Unique | %{"${_}@fil.com"}
}

if(!$Email_To){
    $Email_To = $Email_Cc
}

Add-Log -Path $strLogFile -Value 'About to send email'
Add-Log -Path $strLogFile -Value "Email sending decision is: [$EmailToGo]"
if($EmailToGo)
{
    Send-MailMessage -From $Email_From -To $Email_To -Cc $Email_Cc -BodyAsHtml $HTMLBody -Subject $Email_Subject -SmtpServer $Email_SMTPServer
    if(!$?){
        Add-Log -Path $strLogFile -Value "Failed to send email, cause:"
        Add-Log -Path $strLogFile -Value $Error[0]
    }
}

Add-Log -Path $strLogFile -Value "Exit code is: [$ExitCode]"
