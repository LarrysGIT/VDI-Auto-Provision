
<#
    Version 1.00 [Author: Matthew Pound; Time: 2015-07-06]
        First build, based on VDI.ReclaimStorage.ps1
        For old VDIs, there is no 'Reclaim Last Date' set,
            so default set as '2015-01-01', just a old enough date, no much meanings.
            Script will judge current datastore and desired datastore, if found mismatch, move back
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means connect with vCenter failed.
    second   bit set 1, means failed to copy sdelete files
    third    bit set 1, means failed to trigger sdelete script
    fourth   bit set 1, means failed to move VDI to another datastore
    fifth    bit set 1, means failed to move VDI back to original datastore
    sixth    bit set 1, means 
    seventh  bit set 1, means 
    eigthth  bit set 1, means 
    ninth    bit set 1, means 
    tenth    bit set 1, means 
    eleventh bit set 1, means 
#>

. '.\VDI.Common.Function.ps1'
. '.\_Configuration.ps1'
Define-CommonVariables
Define-VDIBuildVariables
Define-MoveBackVariables

$ExitCode = 0
if(!$Enable)
{
    Add-Log -Path $strLogFile -Value 'Move back script disabled'
    exit($ExitCode)
}

Add-Log -Path $strLogFile -Value 'Move back script started'

if($StartAfter)
{
    Add-Log -Path $strLogFile -Value "StartAfter variable defined: [$StartAfter]"
    $Now = [datetime]::Now
    Add-Log -Path $strLogFile -Value "Current time: [$($Now.ToString('HH:mm'))]"
    $Sleep = [int]([datetime]::Parse($StartAfter) - $Now).TotalSeconds
    if($Sleep -gt 0)
    {
        Add-Log -Path $strLogFile -Value "Script should sleep: [$Sleep]"
        Start-Sleep -Seconds $Sleep
    }
}

#Add-PSSnapin Citrix* -ErrorAction SilentlyContinue
Add-PSSnapin VMware* -ErrorAction SilentlyContinue

$ClearVIServer = $false
if($DefaultVIServer -and $vCenter.Contains(($DefaultVIServer.Name).ToUpper())){
    Add-Log -Path $strLogFile -Value 'This script already connect to VI server, will not connect again'
}else{
    Connect-VIServer -Server $vCenter -ErrorAction:SilentlyContinue
    if(!$?){
        Add-Log -Path $strLogFile -Value "Connect to vCenter $vCenter failed, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        $ExitCode = $ExitCode -bor 0x0001 # 0000 0000 0000 0001
        Add-Log -Path $strLogFile -Value "Script exit with code $ExitCode"
        exit($ExitCode)
    }
}

$Clusters.Keys | %{$j = [int]::MinValue}{
    $Clusters.$_.$Type.Add('VMs', @(Get-VM -Location $Clusters.$_.$Type.Folder `
                                    | ?{!(Get-SnapShot -VM $_) `
                                    -and ($Date - [datetime]($_.CustomFields['Reclaim Last Date'] -ireplace '^\s*$', '2015-01-01')).TotalDays -ge $MoveBackAfter} `
                                    | Select-Object -Property @{N='Name'; E={$_.Name}}, `
                                    @{N='ReclaimLastDate';E={$_.CustomFields['Reclaim Last Date']}} `
                                    | Sort-Object 'ReclaimLastDate'))
    if($Clusters.$_.$Type.VMs.Count -gt $j)
    {
        $j = $Clusters.$_.$Type.VMs.Count
    }
}

$i = $Processed = 0; $ReportArray = @()
while($i -lt $j)
{
    if(((Get-Date $StopBefore) - (Get-Date)).TotalSeconds -le 600)
    {
        Add-Log -Path $strLogFile -Value 'Timeline about to hit, main job quit'
        break
    }

    foreach($Cluster in $Clusters.Keys)
    {
        $VDI = $null
        $VDI = $Clusters.$Cluster.$Type.VMs[$i]
        if($VDI)
        {
            Add-Log -Path $strLogFile -Value "Processing VDI: [$($VDI.Name)][$($VDI.ReclaimLastDate)]"
            (Get-VM -Name $VDI.Name).ExtensionData.RefreshStorageInfo()
            $Committed_Before = 0
            $Committed_Before = '{0:F2}' -f ((Get-VM -Name $VDI.Name).ExtensionData.Storage.PerDatastoreUsage[0].Committed/1GB)

            $Datastore = $null
            $Datastore = @(Get-VM $VDI.Name | Get-Datastore | %{$_.Name})[0]
            $NewDatastore = $Clusters.$Cluster.$Type.Datastore
            Add-Log -Path $strLogFile -Value "The VDI [$($VDI.Name)] is on [$Cluster], datastore should be [$NewDatastore]"
            Add-Log -Path $strLogFile -Value "Current datastore: [$Datastore]"
            if($NewDatastore -eq $Datastore)
            {
                Add-Log -Path $strLogFile -Value 'VDI already locates on desired datastore, no need to move back'
                continue
            }
            
            Add-Log -Path $strLogFile -Value "New datastore set to: [$NewDatastore]"
            Move-VM -VM $VDI.Name -Datastore $NewDatastore -DiskStorageFormat Thin
            if(!$?)
            {
                Add-Log -Path $strLogFile 'Failed to move VDI back, cause:' -Type Error
                Add-Log -Path $strLogFile $Error[0] -Type Error
                $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
                continue
            }
            Add-Log -Path $strLogFile -Value 'Finished move VM back'
            (Get-VM -Name $VDI.Name).ExtensionData.RefreshStorageInfo()
            $Committed_After = 0
            $Committed_After = '{0:F2}' -f ((Get-VM -Name $VDI.Name).ExtensionData.Storage.PerDatastoreUsage[0].Committed/1GB)
            Add-Log -Path $strLogFile -Value "[$($VDI.Name)][$Committed_Before][$Committed_After][$($Committed_After - $Committed_Before)]"
            $Processed++
            $ReportArray += $null
            $ReportArray[-1] = @($VDI.Name, $Committed_Before, $Committed_After)
        }
        else
        {
            Start-Sleep -Seconds 30            
            continue
        }
    }
    $i++
}

if($ClearVIServer){
    Disconnect-VIServer -Server $vCenter -Confirm:$false -Force
}

Remove-PSSnapin VMware* -Confirm:$false

Add-Log -Path $strLogFile -Value "Total processed: [$Processed], report array followed:`r`n$(@($ReportArray | %{$_ -join "`t"}) -join "`r`n")"

exit($ExitCode)
