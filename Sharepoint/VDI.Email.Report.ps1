
<#
    Version 1.01 [Author: Larry Song; Time: 2014-12-24]
        First build for VDI.Report.ps1
        Reconstruct all sharepoint scripts due to bad desgin
    Version 1.02 [Author: Larry Song; Time: 2015-01-07]
        Trying to fix report of rebuild
    Version 1.03 [Author: Larry Song; Time: 2015-01-15]
        Trying to fix report of rebuild
    Version 1.04 [Author: Larry Song; Time: 2015-01-30]
        Fix report bug for rebuilt VDIs
#>

$strLogFile = "$LocalDes\$strDate\Email_Report.log"

Add-Log -Path $strLogFile -Value 'Report script start'
Import-Module *ActiveDirectory*
$RequesterNames = @()
$ExitCode = 0

#$objVDIBuildJob = New-Object PSObject -Property @{User = $null; Changes = $null; CreatedBy = $null; Comment = $null}
#$objVDIRebuildJob = New-Object PSObject -Property @{VDIName = $null; CreatedBy = $null; Comment = $null}

# Left jobs add into variable
Add-Log -Path $strLogFile -Value 'Read jobs left by PODs'
$JobsLeft = @{}
$VDI_ImportTags.Keys | %{
    $PODName = $_
    $JobsLeft.Add($PODName, @{})
    $VDI_Lists_Export.Keys | %{
        $Type = $_
        $JobsLeft.$PODName.Add($Type, @($(
            if(Test-Path -Path "$LocalDes\$strDate\Imports\${PODName}_${Type}_Left.csv")
            {
                Import-Csv -Path "$LocalDes\$strDate\Imports\${PODName}_${Type}_Left.csv" -Delimiter "`t" -ErrorAction:SilentlyContinue | ?{$_.$($VDI_Lists_Export.$Type.LeftKeyProperty)}
            }
        )))
    }
}

# Valid and invalid jobs import into variable
Add-Log -Path $strLogFile -Value 'Read valid and invalid jobs'
$JobsValid = @{}
$JobsInvalid = @{}
$VDI_Lists_Export.Keys | %{
    $JobsValid.Add($_, @($(
        if(Test-Path -Path "$LocalDes\$strDate\Exports\$($VDI_Lists_Export.$_.List)_Export_Raw_0.CSV")
        {
            Import-Csv -Path "$LocalDes\$strDate\Exports\$($VDI_Lists_Export.$_.List)_Export_Raw_0.CSV"
        }
    )))
    $JobsInvalid.Add($_, @($(
        if(Test-Path -Path "$LocalDes\$strDate\$($VDI_Lists_Export.$_.List)_Export_Raw_1.CSV")
        {
            Import-Csv -Path "$LocalDes\$strDate\$($VDI_Lists_Export.$_.List)_Export_Raw_1.CSV"
        }
    )))
}

# template creation logs import into variable
Add-Log -Path $strLogFile -Value 'Read logs for template creation'
$TemplateCreated = @()
$VDI_ImportTags.Keys | %{
    $TemplateCreated += @($(
        if(Test-Path -Path "$LocalDes\$strDate\Imports\${_}_TemplateCreationReport.csv")
        {
            Import-Csv -Path "$LocalDes\$strDate\Imports\${_}_TemplateCreationReport.csv" -Delimiter "`t" -ErrorAction:SilentlyContinue | ?{$_.Number}
        }
    ))
}

# make HTML report
$strHTMLHeader = @'
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 14 (filtered medium)"><style><!--
/* Font Definitions */
@font-face
	{font-family:\5B8B\4F53;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:\5B8B\4F53;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"\@\5B8B\4F53";
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"Lucida Console";
	panose-1:2 11 6 9 4 5 4 2 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Lucida Console","sans-serif";}
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
	font-family:"Lucida Console","sans-serif";
	color:windowtext;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-family:"Lucida Console","sans-serif";}
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
</o:shapelayout></xml><![endif]--></head><body lang=EN-US link=blue vlink=purple><div class=WordSection1>
'@
$strHTMLTableTitle = @'
    <p class=MsoNormal>%Title%<o:p></o:p></p>
'@
$strHTMLTableHead = @'
    <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
'@
$strHTMLTableRow = @'
			<tr>
				<td valign=top style='border:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt'>
					<p class=MsoNormal>%Col1%<o:p></o:p></p>
				</td>
				<td valign=top style='border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt'>
					<p class=MsoNormal>%Col2%<o:p></o:p></p>
				</td>
				<td valign=top style='border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt'>
					<p class=MsoNormal>%Col3%<o:p></o:p></p>
				</td>
				<td valign=top style='border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt'>
					<p class=MsoNormal>%Col4%<o:p></o:p></p>
				</td>
			</tr>
'@
$strHTMLTableTail = @'
    </table>
'@
$strHTMLTail = @'
                <p class=MsoNormal><o:p>&nbsp;</o:p></p>
            </div>
        </body>
    </html>
'@

$strHTMLTables = $null
$ReportFlag = $false

# make report for valid requests
Add-Log -Path $strLogFile -Value 'Start making valid requests requests'
$VDI_Lists_Export.Keys | %{
    $Key = $_
    if($JobsValid.$Key.Count -eq 1)
    {
        Add-Log -Path $strLogFile -Value "Only one item detected in valid list for [$Key], continue check if it's a blank"
        if($JobsValid.$Key[0].$($VDI_Lists_Export.$_.KeyProperty))
        {
            Add-Log -Path $strLogFile -Value 'The item is not blank'
        }
        else
        {
            Add-Log -Path $strLogFile -Value 'The item is blank'
            $JobsValid.$Key = @()
        }
    }
    if($JobsValid.$Key)
    {
        switch($Key)
        {
            'VDI_Build' {
                $strHTMLTables += $strHTMLTableTitle -ireplace '%Title%', 'New VDI build valid requests:'
                $strHTMLTables += $strHTMLTableHead
                $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "<b><span style='color:red'>User</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "<b><span style='color:red'>$($VDI_Lists_Export.$Key.KeyProperty)</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col3%', "<b><span style='color:red'>Submitted by</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col4%', "<b><span style='color:red'>VDI Names</b>"
                for($i = 0; $i -lt $JobsValid.$Key.Count; $i++){
                    $RequesterNames += $JobsValid.$Key[$i].'Created By'
                    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', $JobsValid.$Key[$i].'Alias'
                    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', $($JobsValid.$Key[$i].$($VDI_Lists_Export.$Key.KeyProperty) -ireplace '[\r\n]+', '<br>')
                    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $JobsValid.$Key[$i].'Created By'
                    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $(
                        Add-Log -Path $strLogFile -Value 'Parsing POD info'
                        [regex]::Matches($JobsValid.$Key[$i].$($VDI_Lists_Export.$Key.KeyProperty), '(?i)\* Create VDI in ([\w ]+)[\:\r\n]*') | %{
                            $PODName = $_.Groups[1].Value
                            "$((Get-Content -Path "$LocalDes\$strDate\Exchanges\$($JobsValid.$Key[$i].'Alias')_${PODName}_*_Picked.txt" | %{
                                if(@($JobsLeft.$Key | %{$_.VDIName}) -notcontains $_ -and (Get-Content -Path $VDI_ImportTags[$PODName]) -icontains $_){
                                    "${_}[Y]"
                                }else{
                                    "${_}[N]"
                                }
                            }) -join ' ')<br>"
                        }
                    )
                }
                $strHTMLTables += $strHTMLTableTail
                $strHTMLTables += '<p class=MsoNormal><o:p>&nbsp;</o:p></p>'
            }
            'VDI_Rebuild' {
                ### import VDI rebuild processed requests
                $RebuildJobsProcessed = @{}
                $VDI_ImportTags.Keys | %{
                    $RebuildJobsProcessed.Add($_, @(
                        @(Import-Csv -Path "$LocalDes\$strDate\Imports\${_}_${Key}_Processed.csv" -Delimiter "`t" -ErrorAction:SilentlyContinue | ?{$_.$($VDI_Lists_Export.$Key.LeftKeyProperty)})
                    ))
                }

                $strHTMLTables += $strHTMLTableTitle -ireplace '%Title%', 'Old VDI rebuild valid requests:'
                $strHTMLTables += $strHTMLTableHead
                $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "<b><span style='color:red'>$($VDI_Lists_Export.$Key.LeftKeyProperty)</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "<b><span style='color:red'>Submitted by</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col3%', "<b><span style='color:red'>POD</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $null
                for($i = 0; $i -lt $JobsValid.$Key.Count; $i++){
                    $RequesterNames += $JobsValid.$Key[$i].'Created By'
                    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', $JobsValid.$_[$i].$($VDI_Lists_Export.$_.KeyProperty)
                    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', $JobsValid.$_[$i].'Created By'
                    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $(
                        $tmp = $null
                        $tmp = @($VDI_ImportTags.Keys | %{
                            $PODName = $_
                            $RebuildJobsProcessed.$PODName | %{
                                # if(@($JobsValid.$Key | %{$_.$($VDI_Lists_Export.$Key.KeyProperty)}) -icontains $_.$($VDI_Lists_Export.$Key.KeyProperty)){$PODName}
                                # if($JobsValid.$Key[$i].$($VDI_Lists_Export.$Key.KeyProperty) -icontains $_.$($VDI_Lists_Export.$Key.KeyProperty)){$PODName}
                                if($_.$($VDI_Lists_Export.$Key.KeyProperty) -imatch $JobsValid.$Key[$i].$($VDI_Lists_Export.$Key.KeyProperty)){$PODName}
                            }
                        })
                        if($tmp){$tmp -join '<br>'}else{'No POD found the VDI'}
                    )
                    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $null
                }
                $strHTMLTables += $strHTMLTableTail
                $strHTMLTables += '<p class=MsoNormal><o:p>&nbsp;</o:p></p>'
            }
        }
    }
}

# make report for invalid requests
Add-Log -Path $strLogFile -Value 'Start making invalid requests requests'
$VDI_Lists_Export.Keys | %{
    $Key = $_
    if($JobsInvalid.$Key.Count -eq 1)
    {
        Add-Log -Path $strLogFile -Value "Only one item detected in invalid list for [$Key], continue check if it's a blank"
        if($JobsInvalid.$Key[0].$($VDI_Lists_Export.$_.KeyProperty)){
            Add-Log -Path $strLogFile -Value 'The item is not blank'
        }
        else
        {
            Add-Log -Path $strLogFile -Value 'The item is blank'
            $JobsInvalid.$Key = @()
        }
    }
    if($JobsInvalid.$Key)
    {
        switch($Key)
        {
            'VDI_Build' {
                $strHTMLTables += $strHTMLTableTitle -ireplace '%Title%', 'New VDI build requests after timeline:'
                $strHTMLTables += $strHTMLTableHead
                $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "<b><span style='color:red'>User</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "<b><span style='color:red'>$($VDI_Lists_Export.$Key.KeyProperty)</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col3%', "<b><span style='color:red'>Submitted by</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $null
                for($i = 0; $i -lt $JobsInvalid.$Key.Count; $i++){
                    $RequesterNames += $JobsInvalid.$Key[$i].'Created By'
                    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', $JobsInvalid.$Key[$i].'Alias'
                    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', $($JobsInvalid.$Key[$i].$($VDI_Lists_Export.$Key.KeyProperty) -ireplace '[\r\n]+', '<br>')
                    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $JobsInvalid.$Key[$i].'Created By'
                    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $null
                }
                $strHTMLTables += $strHTMLTableTail
                $strHTMLTables += '<p class=MsoNormal><o:p>&nbsp;</o:p></p>'
            }
            'VDI_Rebuild' {
                $strHTMLTables += $strHTMLTableTitle -ireplace '%Title%', 'Old VDI rebuild requests after timeline:'
                $strHTMLTables += $strHTMLTableHead
                $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "<b><span style='color:red'>$($VDI_Lists_Export.$Key.LeftKeyProperty)</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "<b><span style='color:red'>Submitted by</b>"
                $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $null
                $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $null
                for($i = 0; $i -lt $JobsInvalid.$Key.Count; $i++){
                    $RequesterNames += $JobsInvalid.$Key[$i].'Created By'
                    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', $JobsInvalid.$_[$i].$($VDI_Lists_Export.$_.KeyProperty)
                    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', $JobsInvalid.$_[$i].'Created By'
                    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $null
                    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $null
                }
                $strHTMLTables += $strHTMLTableTail
                $strHTMLTables += '<p class=MsoNormal><o:p>&nbsp;</o:p></p>'
            }
        }
    }
}

# make report for requests left from PODs
Add-Log -Path $strLogFile -Value 'Start making requests before timeline but not processed by PODs'
$VDI_ImportTags.Keys | %{
    $Processed = @{}
    $VDI_Lists_Export.Keys | %{$Processed.Add($_, @{})}
}{
    $PODName = $_
    $VDI_Lists_Export.Keys | %{
        $Key = $_
        switch($Key)
        {
            'VDI_Build' {
                $JobsLeft[$PODName][$Key] | %{
                    if($Processed.$Key.Keys -notcontains $_.$($VDI_Lists_Export.$Key.LeftKeyProperty))
                    {
                        $Processed.$Key.Add($_.$($VDI_Lists_Export.$Key.LeftKeyProperty), @("$(if($_.VDIName){"$($_.VDIName) in "})$PODName", $_.Exception, $_.CreatedBy))
                    }
                    else
                    {
                        $Processed.$Key[$_.$($VDI_Lists_Export.$Key.LeftKeyProperty)][0] += "<br>$(if($_.VDIName){"$($_.VDIName) in "})$PODName"
                        $Processed.$Key[$_.$($VDI_Lists_Export.$Key.LeftKeyProperty)][1] += "<br>$($_.Exception)"
                    }
                }
            }
            'VDI_Rebuild' {
                <#
                $JobsLeft[$PODName][$Key] | %{
                    if($Processed.$Key.Keys -notcontains $_.$($VDI_Lists_Export.$Key.LeftKeyProperty))
                    {
                        $Processed.$Key.Add($_.$($VDI_Lists_Export.$Key.LeftKeyProperty), @($PODName, $_.Exception, $_.CreatedBy))
                    }
                    else
                    {
                        $Processed.$Key[$_.$($VDI_Lists_Export.$Key.LeftKeyProperty)][0] += "<br>$PODName"
                        $Processed.$Key[$_.$($VDI_Lists_Export.$Key.LeftKeyProperty)][1] += "<br>$($_.Exception)"
                    }
                }
                #>
            }
        }
    }
}
<#
$VDI_Lists_Export.Keys | %{
    $Type = $_
    $JobsLeft.Keys | %{
        if($Processed.$Key.Keys -notcontains $JobsLeft.$_.$($VDI_Lists_Export.$Key.LeftKeyProperty))
        {
            $Processed.$Key.Add($JobsLeft.$_.$($VDI_Lists_Export.$Key.LeftKeyProperty), @($null, $_.Exception, $_.CreatedBy))
        }
    }
}
#>

if(@($Processed.Values | %{$_.Values} | %{$_.Values}).Count){
    $strHTMLTables += $strHTMLTableTitle -ireplace '%Title%', 'Requests failed to process:'
    $strHTMLTables += $strHTMLTableHead
    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "<b><span style='color:red'>User|VDI</b>"
    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "<b><span style='color:red'>VDI&POD Name</b>"
    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', "<b><span style='color:red'>Exception</b>"
    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', "<b><span style='color:red'>Submitted by</b>"
    $VDI_Lists_Export.Keys | Sort-Object | %{
        $Key = $_
        switch($Key)
        {
            'VDI_Build' {
                $Processed.$Key.Keys | %{
                    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', $_
                    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', $Processed.$Key.$_[0]
                    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $Processed.$Key.$_[1]
                    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $Processed.$Key.$_[2]
                }
            }
            'VDI_Rebuild' {
                $Processed.$Key.Keys | %{
                    if(($RebuildJobsProcessed.Values) -notcontains $_)
                    {
                        $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', $_
                        $strHTMLTables = $strHTMLTables -ireplace '%Col2%', $Processed.$Key.$_[0]
                        $strHTMLTables = $strHTMLTables -ireplace '%Col3%', $Processed.$Key.$_[1]
                        $strHTMLTables = $strHTMLTables -ireplace '%Col4%', $Processed.$Key.$_[2]
                    }
                }
            }
        }
    }
    $strHTMLTables += $strHTMLTableTail
}

# make report for template creation
if($TemplateCreated){
    $strHTMLTables += '<p class=MsoNormal><o:p>&nbsp;</o:p></p>'
    $strHTMLTables += '<p class=MsoNormal><o:p>&nbsp;</o:p></p>'
    $strHTMLTables += $strHTMLTableTitle -ireplace '%Title%', 'For storage team - Template creation logs:'
    <#
    $strHTMLTables += @'
            <p class=MsoNormal><span style='font-size:14.0pt;color:#7F7F7F;mso-style-textfill-fill-color:#7F7F7F;mso-style-textfill-fill-alpha:100.0%'>&lt;
                    For storage team - Template creation logs
                &gt;<o:p></o:p></span>
            </p>
'@
    #>
    $strHTMLTables += $strHTMLTableHead
    $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "<b><span style='color:red'>vCenter</b>"
    $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "<b><span style='color:red'>Datastore</b>"
    $strHTMLTables = $strHTMLTables -ireplace '%Col3%', "<b><span style='color:red'>Number</b>"
    $strHTMLTables = $strHTMLTables -ireplace '%Col4%', "<b><span style='color:red'>Controller</b>"
    $TemplateCreated | %{
        $strHTMLTables += $strHTMLTableRow -ireplace '%Col1%', "$($_.VIServer)"
        $strHTMLTables = $strHTMLTables -ireplace '%Col2%', "$($_.Datastore)"
        $strHTMLTables = $strHTMLTables -ireplace '%Col3%', "$($_.Number)"
        $strHTMLTables = $strHTMLTables -ireplace '%Col4%', "$($_.Controller)"
    }
    $strHTMLTables += $strHTMLTableTail
}

$RequesterNames = @($RequesterNames | ?{$_} | Sort-Object -Unique)
[array]$Email_To += for($i = 0; $i -lt $RequesterNames.Count; $i++){
    (Get-ADUser -Filter "DisplayName -eq '$($RequesterNames[$i])'" | %{
        if($_.SamAccountName -imatch '(a\d{6})'){return $Matches[1]}
    }) | Sort-Object -Unique | %{"${_}@fil.com"}
}

$Email_To = $Email_To | ?{$_}
if(!$Email_To)
{
    $Email_To = $Email_Cc
}

# Final report
if($strHTMLTables)
{

    $strHTMLBody = $strHTMLHeader + $strHTMLTables + $strHTMLTail
    Send-MailMessage -From $Email_From -To $Email_To -Cc $Email_Cc -SmtpServer $Email_SMTPServer -Subject $Email_Subject -BodyAsHtml $strHTMLBody
    if(!$?)
    {
        Add-Log -Path $strLogFile -Value 'Notification email sent failed, cause:' -Type Warning
        Add-Log -Path $strLogFile -Value $Error[0] -Type Warning
        Add-Log -Path $strLogFile -Value "To: [$($Email_To -join '][')]" -Type Info
        Add-Log -Path $strLogFile -Value "Cc: [$($Email_Cc -join '][')]" -Type Info
    }
    else
    {
        Add-Log -Path $strLogFile -Value 'Notification email sent'
    }
}

Add-Log -Path $strLogFile -Value "Exit code is: [$ExitCode]"
