
<#
    Version 1.02
        Add function for waiting report.
    Version 1.01
        Add version control.
    Version 1.02 [Author: Larry Song; Time: 2014-11-03]
        Add POD capacity report
    Version 1.03 [Author: Larry Song; Time: 2014-11-20]
        Update algorithm for $DSCapacity counting
    Version 1.04 [Author: Larry Song; Time: 2014-12-08]
        Add logs about datastore usage
    Version 1.05 [Author: Larry Song; Time: 2014-12-09]
        Fix bug on capacity warning/alert
    Version 1.06 [Author: Larry Song; Time: 2015-01-16]
        Update codes according to Jones's request for alerting to EUS and wintel
    Version 1.07 [Author: Larry Song; Time: 2015-03-19]
        When capacity lower than 0, set capacity to 0
    Version 1.08 [Author: Larry Song; Time: 2015-04-23]
        Add some figures to report
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means connect with vCenter failed.
    second   bit set 1, means capacity warning
    third    bit set 1, means capacity alert
    fourth   bit set 1, means 
    fifth    bit set 1, means 
    sixth    bit set 1, means 
    seventh  bit set 1, means 
    eigthth  bit set 1, means 
    ninth    bit set 1, means 
    tenth    bit set 1, means 
    eleventh bit set 1, means 
#>

# This script is invoked directly by Starter.ps1, variables inherited.
Define-VDIBuildVariables
Define-VDIDataReportVariables

$Current_Time = Get-Date
$Expect_Time = Get-Date "$strDate $WaitUntil"
$Diff_Time = [int]($Expect_Time - $Current_Time).TotalSeconds
if($Diff_Time -gt 0){
    Add-Log -Path $strLogFile -Value "About to sleep seconds: $Diff_Time"
    if(!$Wait){
        Add-Log -Path $strLogFile -Value 'Wait function is disabled, no need to wait'
    }else{
        Start-Sleep -Seconds $Diff_Time
    }
}

$ExitCode = 0
Add-PSSnapin VMware* -ErrorAction:SilentlyContinue

Add-Log -Path $strLogFile -Value "Start connecting to vCenter $vCenter."
$ClearVIServer = $false
if($DefaultVIServer -and $vCenter.Contains(($DefaultVIServer.Name).ToUpper())){
    Add-Log -Path $LogFile -Value 'This script already connect to VI server, will not connect again'
}else{
    Connect-VIServer -Server $vCenter -ErrorAction:SilentlyContinue
    if(!$?){
        Add-Log -Path $strLogFile -Value "Connect to vCenter $vCenter failed, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        $ExitCode = $ExitCode -bor 0x0002 # 0000 0000 0000 0001
        Add-Log -Path $strLogFile -value "Report null to '$VMReportFile'"
        Set-Content -Path $VMReportFile -Value $null -Force
        Add-Log -Path $strLogFile -Value "Script exit with code $ExitCode"
        exit($ExitCode)
    }
}

Get-VM | %{$_.Name} | Set-Content -Path $VMReportFile -Force

############################
Add-Log -Path $strLogFile -Value 'Start caculate capacity'
$VDICapacityRemaining_Cluster = $GolbalVDILimite
$VDICapacityRemaining_Datastore = 0
$Clusters.Keys | %{$DSCapacity = 0; $DSUsed = 0}{
    $VDICapacityRemaining_Cluster -= (Get-VM -Location $_ | ?{$_.Name -imatch $Clusters_VMFilter}).Count
    $objDatastore_ = Get-Datastore -Name $Clusters[$_][$CapacityReportType]['Datastore'] -Refresh
    Add-Log -Path $strLogFile -Value "Volume control state for: [$($Clusters[$_][$CapacityReportType]['Datastore'])][$($Clusters[$_][$CapacityReportType]['VolumeControl'])]"
    if($Clusters[$_][$CapacityReportType]['VolumeControl'])
    {
        $DSCapacity += ($objDatastore_.ExtensionData.Summary.Capacity * $Clusters[$_][$CapacityReportType]['CapacityMultiple'])/1GB
    }
    else
    {
        $DSCapacity += ($objDatastore_.ExtensionData.Summary.Capacity)/1GB
    }
    $DSUsed += ($objDatastore_.ExtensionData.Summary.Capacity - $objDatastore_.ExtensionData.Summary.FreeSpace)/1GB
}
$DSCapacity = [int]$DSCapacity
$DSUsed = [int]$DSUsed
Add-Log -Path $strLogFile -Value "Datastore total capacity: [$DSCapacity GB]"
Add-Log -Path $strLogFile -Value "Datastore total capacity: [$($DSCapacity * $CapacityWarning) GB]"
Add-Log -Path $strLogFile -Value "Datastore total used: [$DSUsed GB]"
$DSCapacity_ = [int](($DSCapacity * $CapacityWarning - $DSUsed)/6)
if($DSCapacity_ -lt 0)
{
    $DSCapacity_ = 0
}
Add-Log -Path $strLogFile -Value "VDI datastore capacity remaining: [$DSCapacity_]"
Add-Log -Path $strLogFile -Value "VDI cluster capacity: [$GolbalVDILimite]"
Add-Log -Path $strLogFile -Value "VDI cluster capacity remaining: [$VDICapacityRemaining_Cluster]"
Add-Log -Path $strLogFile -Value "Warning/Alert: [$CapacityWarning/$CapacityAlert]"
$VDIRate = 1 - $VDICapacityRemaining_Cluster/$GolbalVDILimite
$DSRate = $DSUsed/$DSCapacity
Add-Log -Path $strLogFile -Value "VDI cluster usage: [$VDIRate]"
Add-Log -Path $strLogFile -Value "Datastore usage: [$DSRate]"
$MailTo_ = $Subject_ = @()
if($CapacityMonitoringEnable)
{
    if($VDIRate -ge $CapacityAlert)
    {
        $ExitCode = $ExitCode -bor 0x0008 # 0000 1000 0000 1000
        $MailTo_ += $CapacityAlertMailTo
        $Subject_ += $CapacityAlertMailSubject
    }
    elseif($VDIRate -ge $CapacityWarning)
    {
        $ExitCode = $ExitCode  -bor 0x0004 # 0000 1000 0000 0100
        $MailTo_ += $CapacityWarningMailTo
        $Subject_ += $CapacityWarningMailSubject
    }

    if($DSRate -ge $CapacityAlert)
    {
        $ExitCode = $ExitCode  -bor 0x0001 # 0000 1000 0000 0001
        $MailTo_ += $CapacityAlertMailTo
        $Subject_ += $CapacityAlertMailSubject
    }
    elseif($DSRate -ge $CapacityWarning)
    {
        $ExitCode = $ExitCode  -bor 0x0002 # 0000 0000 0000 0010
        $MailTo_ += $CapacityWarningMailTo
        $Subject_ += $CapacityWarningMailSubject
    }
}

$MailTo_ = $MailTo_ | Sort-Object -Unique
$Subject_ = $Subject_ | Sort-Object -Unique
$Body_ = (Get-Content $strLogFile) -join "`r`n"

Send-MailMessage -From $EmailFrom -To $MailTo_ -Subject $Subject_ -Body $Body_ -SmtpServer $EmailSMTPServer

Set-Content -Path $CapacityReport -Value ([math]::Min($VDICapacityRemaining_Cluster, $DSCapacity_))
############################

if($ClearVIServer){
    Disconnect-VIServer -Server $vCenter -Confirm:$false -Force
}

Remove-PSSnapin VMware* -Confirm:$false

exit($ExitCode)
