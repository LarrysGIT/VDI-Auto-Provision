
<#
    Version 1.01 [Author: Larry Song]
        Add version control.
        Update 'POD Changes' paring support like [XX POD1:Force:10].
    Version 1.02 [Author: Larry Song]
        Re-design this script to combine with VDI.Build.Core.ps1 script.
    Version 1.03 [Author: Larry Song]
        Add 23:00 time control, combined Data.Report.ps1.
    Version 1.04 [Author: Larry Song]
        Template auto creation according to number of VDIs in this POD.
    Version 1.05 [Author: Larry Song]
        Add disk control and template automatic creation switch, switch defines in _Configuration.ps1.
    Version 1.06 [Author: Larry Song; Time: 2014-02-11]
        Adjust code to apply 1.03 version of VDI.Build.Core.ps1
        Add timeout for template creation task.
        Auto add "-" at tail when the last letter of template prefix matched [a-z\d].
    Version 1.07 [Author: Larry Song; Time: 2014-02-11]
        Adjust code to apply 1.04 version of VDI.Build.Core.ps1
    Version 1.08 [Author: Larry Song; Time: 2014-02-25]
        Adjust code to apply 1.01 version of VDI.Template.Creation.ps1
    Version 1.09 [Author: Larry Song; Time: 2014-02-27]
        Adjust code to identity parameter '-VDIName' which shall provided in advanced options.
    Version 1.10 [Author: Larry Song; Time: 2014-03-20]
        Add variable $PickClusterByDS, used to pick up cluster by datastore.
    Version 2.00 [Author: Larry Song; Time: 2014-03-29]
        Rebuild the part of cluster automatic pickup.
        Optimized some algorithm.
    Version 2.01 [Author: Larry Song; Time: 2014-05-23]
        Add template creation report, so sharepoint script can send.
    Version 2.02 [Author: Larry Song; Time: 2014-09-18]
        Pass -BizGroup parameter to core script.
    Version 2.03 [Author: Larry Song; Time: 2014-10-24]
        Add global VDI number control, variable is $GolbalVDILimite
    Version 2.04 [Author: Larry Song; Time: 2014-11-03]
        Align exit codes, remove some unused codes
    Version 2.05 [Author: Larry Song; Time: 2014-12-31]
        Update to apply new version of VDI.Template.Creation.ps1
    Version 2.06 [Author: Larry Song; Time: 2015-01-16]
        Update to avoid "??VA" VDI for non-POD1 POD
    Version 2.07 [Author: Larry Song; Time: 2015-07-03]
        Kill pending powershell instance after timeout
    Version 2.08 [Author: Larry Song; Time: 2015-07-07]
        Setup global custom attributes after connected to vCenter
#>

<#
    $ExitCode details,
    XXXX XXXX XXXX XXXX
    From right to left,
    first    bit set 1, means local folder created failed.
    second   bit set 1, means read raw file failed.
    third    bit set 1, means when hit 23:00, stop continue creation.
    fourth   bit set 1, means VMware snapin failed adding.
    fifth    bit set 1, means AD module import failed.
    sixth    bit set 1, means there is user not found in AD.
    seventh  bit set 1, means no POD info matched.
    eigthth  bit set 1, means there is abnormal exit code for powershell core script.
    ninth    bit set 1, means no exit code captured for powershell core script.
    tenth    bit set 1, means 1 or more datastores encounter insufficient space.
    eleventh bit set 1, means hit threshold.
#>

Set-Location (Get-Item ($MyInvocation.MyCommand.Definition)).DirectoryName
. '.\_Configuration.ps1'
Define-CommonVariables
Define-VDIBuildVariables

$AdditionalProperties = 'CanonicalName'
$objClusterJob = New-Object PSObject -Property @{ID = $null; User = $null; VDIName = $null; Type = $null; Exception = $null; AdvanceOptions = $null; CreatedBy = $null}
. '.\VDI.Common.Function.ps1'

$ExitCode = 0

function Quit{
    PARAM(
        [int]$ExitCode,
        [int[]]$AdditionalJobs
    )
    switch($AdditionalJobs){
        1 {
            @($objClusterJob) | Export-Csv -Path $VMReportLeftFile -NoTypeInformation -Delimiter "`t"
        }
        default{}
    }
    exit($ExitCode)
}

New-Item -Path "$LocalDes\$strDate" -ItemType 'Directory' -Confirm:$false -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
if(!(Test-Path -Path "$LocalDes\$strDate" -PathType Container)){
    Add-Log -Path $strLogFile -Value "Local folder creation failed, cause:" -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0001 # 0000 0000 0000 0001
    Add-Log -Path $strLogFile -Value "Script exit with code $ExitCode"
    Quit -ExitCode $ExitCode -AdditionalJobs 1
}

Add-Log -Path $strLogFile -Value 'Script start'
Add-Log -Path $strLogFile -Value "Check tag file $RawFile"

while($true){
    if(Test-Path -Path $RawFile){
        break
    }
    Add-Log -Path $strLogFile -Value "Tag file not found, sleep 5 mins."
    Start-Sleep -Seconds 300
}

$RawData = Import-Csv -Path $RawFile
if(!$?){
    Add-Log -Path $strLogFile -Value 'Read raw data file failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0002 # 0000 0000 0000 0010
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    Quit -ExitCode $ExitCode -AdditionalJobs 1
}
$RawData = @($RawData)

Add-Log -Path $strLogFile -Value "Read raw data succeed, count: $($RawData.Count)"
if($RawData.Count -eq 1){
    Add-Log -Path $strLogFile -Value "Only 1 item imported, check whether it's a blank."
    if($RawData[0].'POD Changes'){
        Add-Log -Path $strLogFile -Value "The item is not blank, move on."
    }else{
        Add-Log -Path $strLogFile -Value "The item is blank, no need to process. exit with code $ExitCode"
        Quit -ExitCode $ExitCode -AdditionalJobs 1
    }
}

Add-Log -Path $strLogFile -Value 'Start adding VMware snapin.'
Add-PSSnapin VMware* -ErrorAction:SilentlyContinue
if(!$?){
    Add-Log -Path $strLogFile -Value 'Add VMware snapin failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0008 # 0000 0000 0000 1000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    Quit -ExitCode $ExitCode -AdditionalJobs 1
}

Add-Log -Path $strLogFile -Value 'Add VMware snapin succeed.'
<#
Add-Log -Path $strLogFile -Value 'Start importing ActiveDirectory module.'

Import-Module -Name 'ActiveDirectory' -ErrorAction:SilentlyContinue
if(!$?){
    Add-Log -Path $strLogFile -Value 'Import ActiveDirectory module failed, cause:' -Type Error
    Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    $ExitCode = $ExitCode -bor 0x0010 # 0000 0000 0001 0000
    Add-Log -Path $strLogFile -Value "Script quit with code $ExitCode"
    exit($ExitCode)
}

Add-Log -Path $strLogFile -Value 'Import ActiveDirectory module succeed.'
#>

$ClearVIServer = $false
if($DefaultVIServer -and $vCenter.Contains(($DefaultVIServer.Name).ToUpper())){
    Add-Log -Path $LogFile -Value "This script already connect to VI server, will not connect again."
}else{
    Connect-VIServer -Server $vCenter
    if(!$?){
        Add-Log -Path $strLogFile -Value "Connect to vCenter $vCenter failed, cause:" -Type Error
        Add-Log -Path $strLogFile -Value $Error[0] -Type Error
    }
    Add-Log -Path $strLogFile -Value "Start counting VMs in clusters."
}

@(
    'DeployDate',
    'Description',
    'Owner',
    'Real Name',
    'Reclaim Status',
    'Reclaim Last Date',
    'XdConfig',
    'XenDesktop Group'
) | %{
    if(!(Get-CustomAttribute -Name $_ -TargetType VirtualMachine -ErrorAction:SilentlyContinue))
    {
        New-CustomAttribute -Name $_ -TargetType VirtualMachine
    }
}

$Clusters.Keys | Sort-Object | %{
    ($Clusters[$_]).Add('Jobs', @())
    ($Clusters[$_]).Add('Number', 0)
    $Clusters[$_]['Number'] = (Get-VM -Location $_ | ?{$_.Name -imatch $Clusters_VMFilter}).Count
    Add-Log -Path $strLogFile -Value "There are $($Clusters[$_]['Number']) VMs in $_."
    Add-Log -Path $strLogFile -Value "Pickup cluster by datastore status enabled, collecting DS usage data."
    foreach($Type in $Clusters[$_].Keys){
        if($Clusters[$_][$Type].Datastore -imatch '^\s*$')
        {
            continue
        }
        $objDatastore_ = $null
        $objDatastore_ = Get-Datastore -Name $Clusters[$_][$Type]['Datastore'] -Refresh
        $Clusters[$_][$Type].Add('DatastoreUsage', @{
            'Used' = [int](($objDatastore_.ExtensionData.Summary.Capacity - $objDatastore_.ExtensionData.Summary.FreeSpace)/1GB);
            'Capacity' = [int]($objDatastore_.ExtensionData.Summary.Capacity/1GB);
            }
        )
        Add-Log -Path $strLogFile -Value "[$_] [$Type] datastore usage:"
        Add-Log -Path $strLogFile -Value "[$($Clusters[$_][$Type]['Datastore'])]: $($Clusters[$_][$Type]['DatastoreUsage']['Used']) Used(GB)."
        Add-Log -Path $strLogFile -Value "[$($Clusters[$_][$Type]['Datastore'])]: $($Clusters[$_][$Type]['DatastoreUsage']['Capacity']) Capacity(GB)."
        Add-Log -Path $strLogFile -Value "[$($Clusters[$_][$Type]['Datastore'])]: $($Clusters[$_][$Type]['CapacityMultiple']) capacity multiple."
        $Clusters[$_][$Type]['DatastoreUsage']['Capacity'] = $Clusters[$_][$Type]['DatastoreUsage']['Capacity'] * $Clusters[$_][$Type]['CapacityMultiple']
        Add-Log -Path $strLogFile -Value "[$($Clusters[$_][$Type]['Datastore'])]: $($Clusters[$_][$Type]['DatastoreUsage']['Capacity']) Capacity(GB) final."
        if(($Clusters[$_][$Type]['DatastoreUsage']['Used'] + 6) -gt $Clusters[$_][$Type]['DatastoreUsage']['Capacity']){
            Add-Log -Path $strLogFile -Value "Set discision to FALSE."
            $Clusters[$_][$Type].Add('Allow', $false)
        }else{
            # DS can host at least 1 VDI more.
            Add-Log -Path $strLogFile -Value "Set discision to TRUE."
            $Clusters[$_][$Type].Add('Allow', $true)
        }
    }
}

$JobsTotalCount = 0
Add-Log -Path $strLogFile -Value "Start pre-processing."
$JobsLeft = @()

$Clusters.Keys | %{$TotalCake = 0; $TotalVDI = 0}{$TotalCake += $Clusters.$_.Proportion; $TotalVDI += $Clusters[$_]['Number']}

for($i = 0; $i -lt $RawData.Count; $i++){
    Add-Log -Path $strLogFile -Value "****** Start process $($RawData[$i].'Alias') $(($i + 1))/$($RawData.Count)"
    Add-Log -Path $strLogFile -Value "Parsing 'POD Changes' info, string:$($RawData[$i].'POD Changes')"
    $RegResults = [regex]::Matches($RawData[$i].'POD Changes', '(?i)Create VDI in ([\S ]*)') | %{$_.Groups[1].Value}
    Add-Log -Path $strLogFile -Value "Matched results: [$($RegResults -join "], [")]"
    if($RegResults.Count -eq 0){
        Add-Log -Path $strLogFile -Value "Nothing matched, skip!" -Type Warning
        $ExitCode = $ExitCode -bor 0x0040 # 0000 0000 0100 0000
        continue
    }
    foreach($tPOD in $RegResults){
        Add-Log -Path $strLogFile -Value "Parsing [$tPOD]"
        $SubResults = $null
        $SubResults = $tPOD.Split(':')
        $Tag_PODName = $SubResults[0]
        $Tag_Force = $SubResults[1]
        $Tag_Number = $SubResults[2]
        Add-Log -Path $strLogFile -Value "TAG Original  - POD Name: [$Tag_PODName]; Force: [$Tag_Force]; Number: [$Tag_Number]"
        if($Tag_Force -ine 'Force'){
            $Tag_Force = $null
            Add-Log -Path $strLogFile -Value "Force tag not equal to 'Force', set to null."
        }
        if($Tag_Number -imatch '^$'){
            $Tag_Number = 1
            Add-Log -Path $strLogFile -Value "Number tag is blank, set to default 1."
        }
        if($Tag_Number -notmatch '^\d+$'){
            Add-Log -Path $strLogFile -Value "$Tag_Number is not a valid number, skipped current item." -Type Warning
            continue
        }
        Add-Log -Path $strLogFile -Value "TAG Processed - POD Name: [$Tag_PODName]; Force: [$Tag_Force]; Number: [$Tag_Number]"
        $Tag_Number = [int]$Tag_Number
        switch ($Tag_PODName){
        "$POD" {
            if($objUser = Get-ADExisting -SAMAccountName $RawData[$i].'Alias' -Type user -Properties @('name', 'useraccountcontrol', 'canonicalname', 'distinguishedname')){
                Add-Log -Path $strLogFile -Value "User found in AD."
                Add-Log -Path $strLogFile -Value "Name: $($objUser.Properties['name'][0])"
                Add-Log -Path $strLogFile -Value "Status: $(if([int][string]($objUser.Properties['useraccountcontrol'][0]) -band 0x0002){"Disabled"}else{"Enabled"})"
                Add-Log -Path $strLogFile -Value "CanonicalName: $($objuser.Properties['canonicalname'][0])"
                Add-Log -Path $strLogFile -Value "DN: $($objUser.Properties['distinguishedname'][0])"
            }else{
                Add-Log -Path $strLogFile -Value "User not found in AD, this item should not been exported. skip!" -Type Warning
                $ExitCode = $ExitCode -bor 0x0020 # 0000 0000 0010 0000
                continue
            }
<#
            if($objUser = Get-ADUser -Filter "SAMAccountName -eq '$($RawData[$i].'Alias')'" -Properties $AdditionalProperties){
                Add-Log -Path $strLogFile -Value "User found in AD."
                Add-Log -Path $strLogFile -Value "Name: $($objUser.Name)"
                Add-Log -Path $strLogFile -Value "Status: $(if($objUser.Enabled){"Enabled"}else{"Disabled"})"
                Add-Log -Path $strLogFile -Value "CanonicalName: $($objuser.CanonicalName)"
                Add-Log -Path $strLogFile -Value "DN: $($objUser.DistinguishedName)"
            }else{
                Add-Log -Path $strLogFile -Value "User not found in AD, this item should not been exported. skip!" -Type Warning
                # $ExitCode = $ExitCode -bor 0x0020 # 0000 0000 0010 0000
                continue
            }
#>
            $FlagCreate = $false, $false, $false
            $aNumber = $RawData[$i].'Alias' -ireplace '^a', ''
            $objVM = Get-VM "*$aNumber*" | %{$_.Name}
            if($objVM){
                Add-Log -Path $strLogFile -Value "Found VDI for user exists in $Tag_PODName, name: [$($objVM -join "], [")]. About to skip." -Type Warning
                if($Tag_Force -ieq 'Force'){
                    Add-Log -Path $strLogFile -Value "'Force' tag found, continue to create VDI, set decision to TRUE."
                    $FlagCreate[0] = $true
                }else{
                    Add-Log -Path $strLogFile -Value "':Force' tag not found, skipped, set decision to FALSE."
                    $FlagCreate[0] = $false
                    $JobsLeft += $objClusterJob.PSObject.Copy()
                    $JobsLeft[-1].User = $RawData[$i].'Alias'
                    $JobsLeft[-1].Exception = "VDI exists but no FORCE tag found."
                    $JobsLeft[-1].CreatedBy = $RawData[$i].'Created By'
                }
            }else{
                Add-Log -Path $strLogFile -Value "VDI not found in $Tag_PODName, good to go, set decision TRUE."
                $FlagCreate[0] = $true
            }
            $AvailableVDIName_File = "$RemoteDes\$strDate\Exchanges\$($($RawData[$i].'Alias'))_${Prefix}_Available.txt"
            $AvailableVDINameAll = @(Get-Content -Path $AvailableVDIName_File)
            Add-Log -Path $strLogFile -Value "Available VDI names: $AvailableVDIName_File, count: $($AvailableVDINameAll.Count)"

            if($RawData[$i].'AdvanceOptions' -imatch '-VDIName (?<Close>"?)\b(\w+)\b\k<Close>'){
                Add-Log -Path $strLogFile -Value "Advance option: [$($RawData[$i].'AdvanceOptions')]"
                Add-Log -Path $strLogFile -Value "VDIName parameter provided: [$($Matches[1])], chose."
                $AvailableVDINameAll = @($Matches[1])
            }elseif($WaitForPick){
                Add-Log -Path $strLogFile -Value "VDIName parameter not provided, remove '${Prefix}A' from available list."
                $AvailableVDINameAll = $AvailableVDINameAll | ?{$_ -notmatch '^${Prefix}A'}
            }
            $Picked_Set = "$RemoteDes\$strDate\Exchanges\$($($RawData[$i].'Alias'))_${POD}_${Prefix}_Picked.txt"
            $Picked_Get = $null
            if($WaitForPick -and ($RegResults -imatch $WaitForPick)){
                Add-Log -Path $strLogFile -Value "Waiting for primary $WaitForPick POD to pick VDI name first."
                $Picked_File_Wait = "$RemoteDes\$strDate\Exchanges\$($($RawData[$i].'Alias'))_${WaitForPick}_${Prefix}_Picked.txt"
                while(!(Test-Path -Path $Picked_File_Wait)){
                    Start-Sleep -Seconds 60
                }
                $Picked_Get = @(Get-Content -Path $Picked_File_Wait)
                Add-Log -Path $strLogFile -Value "VDI name picked count: $($Picked_Get.Count), remove VDI names already picked by other PODs."
                $AvailableVDI = @($AvailableVDINameAll | ?{$Picked_Get -notcontains $_})
            }else{
                $AvailableVDI = $AvailableVDINameAll
            }
            if($WaitForPick)
            {
                Add-Log -Path $strLogFile -Value 'Reserve "A" VDI for primary POD'
                $AvailableVDI = @($AvailableVDI -notmatch '^..VA')
            }

            Add-Log -Path $strLogFile -Value "VDI names for [$POD]: $($AvailableVDI -join ' ')"
            if($AvailableVDI.Count -lt $Tag_Number){
                Add-Log -Path $strLogFile -Value "Request says to create $Tag_Number VDIs for this user, but there is no enough available names in AD." -Type Warning
                Add-Log -Path $strLogFile -Value "All available name will be used, the exceed will be discarded." -Type Warning
                $VDINames = $AvailableVDI
            }else{
                $VDINames = $AvailableVDI[0..$($Tag_Number - 1)]
            }
            $VDINames = @($VDINames | ?{$_ -notmatch '^\s*$'})
            Add-Log -Path $strLogFile -Value "VDI name choose: [$($VDINames -join "], [")]."
            $VDINames | Set-Content -Path $Picked_Set
            if($VDINames){
                $FlagCreate[1] = $true
            }
            if($TotalVDI -lt $GolbalVDILimite)
            {
                $FlagCreate[2] = $true
            }
            else
            {
                Add-Log -Path $strLogFile -Value 'VDI number will over global VDI limite' -Type Warning
            }
            $FlagCreate = $FlagCreate[0] -and $FlagCreate[1] -and $FlagCreate[2]
            if(!$FlagCreate){
                Add-Log -Path $strLogFile -Value "Final decision is [$FlagCreate]. skiped."
                continue
            }else{
                Add-Log -Path $strLogFile -Value "Final decision is [$FlagCreate]. move on."
            }
            $ClusterToGo = $null
            if($RawData[$i].'AdvanceOptions' -imatch '-Cluster (?<Close>"?)\b(\w+)\b\k<Close>'){
                Add-Log -Path $strLogFile -Value "Advance option: [$($RawData[$i].'AdvanceOptions')]"
                Add-Log -Path $strLogFile -Value "Cluster parameter provided: [$($Matches[1])]"
                if($Clusters.Keys -icontains $Matches[1]){
                    Add-Log -Path $strLogFile -Value "$($Matches[1]) found in clusters, pick it."
                    $ClusterToGo = $Matches[1]
                }else{
                    Add-Log -Path $strLogFile -Value "$($Matches[1]) not found in clusters, skip this user."
                    continue
                }
                Add-Log -Path $strLogFile -Value "Cluster chose: [$ClusterToGo], VM count: $($Clusters[$ClusterToGo]['Number'])"
            }
            switch -Regex ($RawData[$i].'OU'){
                'IMS'{$Type = 'IM'}
                default{$Type = 'FIL'}
            }
            for($iVDI = 0; $iVDI -lt $VDINames.Count; $iVDI++){
                Add-Log -Path $strLogFile -Value "Current progress: $($iVDI + 1)/$Tag_Number"
                if(!$ClusterToGo){
                    # Pickup cluster automatically
                    $objClusterUsage = New-Object PSObject -Property @{'Name' = $null; Free = 0}
                    $Clusters.Keys | %{
                        if($Clusters[$_]['Number']/$TotalVDI -le $Clusters[$_]['Proportion']/$TotalCake){
                            $ClusterToGo = $_
                        }
                    }
                    Add-Log -Path $strLogFile -Value "Picked up [$ClusterToGo] by cluster balancing."
                    if($PickClusterByDS){
                        Add-Log -Path $strLogFile -Value "Pickup cluster by datastore usage enabled."
                        $ClusterToGo = $null
                        Add-Log -Path $strLogFile -Value "Try to pickup cluster with available [$Type] datastore."
                        $CandidateClusters = @()
                        $Clusters.Keys | %{
                            if($Clusters[$_][$Type]['Allow']){
                                $CandidateClusters += $objClusterUsage.PSObject.Copy()
                                $CandidateClusters[-1].Name = $_
                                $CandidateClusters[-1].Free = $Clusters[$_][$Type]['DatastoreUsage']['Capacity'] - $Clusters[$_][$Type]['DatastoreUsage']['Used']
                            }
                        }
                        if($CandidateClusters){
                            Add-Log -Path $strLogFile -Value "Cluster with [$Type] datastore availble: [$(($CandidateClusters | %{"Name: $($_.Name), Free: $($_.Free)"}) -join '], [')]"
                            $ClusterToGo = @($CandidateClusters | Sort-Object Free -Descending)[0].Name
                            Add-Log -Path $strLogFile -Value "Cluster picked [$ClusterToGo]"
                        }else{
                            Add-Log -Path $strLogFile -Value "No cluster with [$Type] datastore found." -Type Warning
                            Add-Log -Path $strLogFile -Value "Try to get [$Type] datastore with disk control off."
                            $CandidateClusters = @()
                            $Clusters.Keys | %{
                                if(!$Clusters[$_][$Type]['VolumeControl']){
                                    $CandidateClusters += $objClusterUsage.PSObject.Copy()
                                    $CandidateClusters[-1].Name = $_
                                    $CandidateClusters[-1].Free = $Clusters[$_][$Type]['DatastoreUsage']['Capacity'] - $Clusters[$_][$Type]['DatastoreUsage']['Used']
                                }
                            }
                            if($CandidateClusters){
                                Add-Log -Path $strLogFile -Value "Cluster with [$Type] datastore availble: [$(($CandidateClusters | %{"Name: $($_.Name), Free: $($_.Free)"}) -join '], [')]"
                                $ClusterToGo = @($CandidateClusters | Sort-Object Free -Descending)[0].Name
                                Add-Log -Path $strLogFile -Value "Cluster picked [$ClusterToGo]"
                            }else{
                                Add-Log -Path $strLogFile -Value "No cluster with [$Type] datastore disk control off found." -Type Warning
                            }
                        }
                    }
                }else{
                    # Cluster provided by parameter
                    if(!$Clusters[$ClusterToGo][$Type]['Allow'] -and $Clusters[$ClusterToGo][$Type]['VolumeControl']){
                        Add-Log -Path $strLogFile -Value "There is no enough space on [$ClusterToGo] for [$Type] datastore." -Type Warning
                        $ClusterToGo = $null
                    }
                }
                if($ClusterToGo){
                    Add-Log -Path $strLogFile -Value "Update static datastore usage data."
                    $Clusters[$ClusterToGo][$Type]['DatastoreUsage']['Used'] += 6
                    if(($Clusters[$ClusterToGo][$Type]['DatastoreUsage']['Used'] + 6) -ge $Clusters[$ClusterToGo][$Type]['DatastoreUsage']['Capacity']){
                        $Clusters[$ClusterToGo][$Type]['Allow'] = $false
                    }
                }else{
                    Add-Log -Path $strLogFile -Value "No cluster available, this job marked as discarded." -Type Warning
                    $JobsLeft += $objClusterJob.PSObject.Copy()
                    $JobsLeft[-1].User = $RawData[$i].'Alias'
                    $JobsLeft[-1].Exception = "Insufficient space"
                    $JobsLeft[-1].CreatedBy = $RawData[$i].'Created By'
                    continue
                }
                Add-Log -Path $strLogFile -Value "Push job in [$Type] stack with ID: $JobsTotalCount."
                $JobsTotalCount++
                $Clusters[$ClusterToGo]['Jobs'] += $objClusterJob.PSObject.Copy()
                $Clusters[$ClusterToGo]['Jobs'][-1].User= $RawData[$i].'Alias'
                $Clusters[$ClusterToGo]['Jobs'][-1].ID = $JobsTotalCount
                $Clusters[$ClusterToGo]['Jobs'][-1].VDIName = $VDINames[$iVDI]
                $Clusters[$ClusterToGo]['Jobs'][-1].Type = $Type
                $Clusters[$ClusterToGo]['Jobs'][-1].AdvanceOptions = $RawData[$i].'AdvanceOptions'
                $Clusters[$ClusterToGo]['Jobs'][-1].CreatedBy = $RawData[$i].'Created By'
                $Clusters[$ClusterToGo]['Number']++
                $TotalVDI++
            }
            break
        }
        default {
            Add-Log -Path $strLogFile -Value "[$Tag_PODName] VDI request handled by [$Tag_PODName] build script. Skipped."
            # Add-Log -Path $strLogFile -Value "Unknown POD name: $tPOD" -Type Warning
        }
        }
    }
}

Add-Log -Path $strLogFile -Value "Continue VDI template creation."
$TemplateJobs = @()
$objTemplateJob = New-Object PSObject -Property @{TaskID = $null; VIServer = $null; Controller = $null; MasterTemplate = $null; Cluster = $null; Datastore = $null; Container = $null; TemplateNames = $null; Number = $null; Exception = $null}
$Clusters.Keys | Sort-Object | %{
    $VIServer = $vCenter
}{
    $TemplateToCreate = @{IM = 0; FIL = 0}
    $Cluster = $_
    $Controller = $Clusters[$Cluster]['Controller']
    Add-Log -Path $strLogFile -Value "Name: [$_]; Jobs count: $($Clusters[$Cluster]['Jobs'].Count)"
    $Clusters[$Cluster]['Jobs'] | %{
        switch($_.Type){
        "IM" {$TemplateToCreate['IM']++}
        "FIL" {$TemplateToCreate['FIL']++}
        }
    }
    $TemplateToCreate.Keys | %{
        $Type = $_

        Add-Log -Path $strLogFile -Value "Type: $Type; Preparing to create: $($TemplateToCreate[$Type])."
        $Container = $Clusters[$Cluster][$Type]['POOL']
        $TemplateInPool = @(Get-VM -Location $Container -ErrorAction:SilentlyContinue | ?{$_.PowerState -eq 'PoweredOff'} | Sort-Object Name | %{$_.Name})
        Add-Log -Path $strLogFile -Value "Templates in pool count: $($TemplateInPool.Count)."
        $MoreOrLess = $TemplateInPool.Count - $TemplateToCreate[$Type]
        if($MoreOrLess -ge $Clusters[$Cluster][$Type]['TemplatePreserve']){
            # $Clusters[$Cluster][$Type]['SizeEstimate'] = $TemplateToCreate[$_] * 6
            Add-Log -Path $strLogFile -Value "${Type}: Current template is enough, no need to create."
        }
        if($MoreOrLess -ge 0 -and $MoreOrLess -lt $Clusters[$Cluster][$Type]['TemplatePreserve']){
            # $Clusters[$Cluster][$Type]['SizeEstimate'] = $TemplateToCreate[$_] * 6
            Add-Log -Path $strLogFile -Value "${Type}: Current template is enough, preserve $($Clusters[$Cluster][$Type]['TemplatePreserve']) for future use."
        }elseif($MoreOrLess -lt 0){
            # $Clusters[$Cluster][$Type]['SizeEstimate'] = $TemplateToCreate[$_] * 6
            Add-Log -Path $strLogFile -Value "${Type}: Current template is not enough."
        }
        $MoreOrLess = $Clusters[$Cluster][$Type]['TemplatePreserve'] - $MoreOrLess

        if($MoreOrLess -le 0){
            Add-Log -Path $strLogFile -Value "Template doesn't needed on $Cluster for $Type. Continue next."
            return
        }
        Add-Log -Path $strLogFile -Value "Template needed on $Cluster for $Type, number: $MoreOrLess."

        if(!$Clusters[$Cluster][$Type]['TemplateCreation']){
            Add-Log -Path $strLogFile -Value "Template automatic creation is disabled, will not create any template."
            return
        }
        $MasterTemplate = $Clusters[$Cluster][$Type]['Template']
        $NumCPU = $Clusters[$Cluster][$Type]['NumCPU']
        $MemoryMB = $Clusters[$Cluster][$Type]['MemoryMB']
        $TemplateToGo = TemplatePick -Cluster $Cluster -Template $MasterTemplate -CurrentMax $TemplateInPool[-1] -Auto
        $Datastore = $Clusters[$Cluster][$Type]['Datastore']
        $Controller = $Clusters[$Cluster]['Controller']
        $Container = $Clusters[$Cluster][$Type]['POOL']
        $TemplatePrefix = $TemplateToGo[0] -ireplace '([a-z\d])$', '$1-'
        $StartNumber = [int]$TemplateToGo[1]
        $NumberAlign = [int]$TemplateToGo[2]
        $Number = $MoreOrLess
        $TemplateNames = @()
        for($VDINameCount = 0; $VDINameCount -lt $Number; $VDINameCount++){
            $strTemplateName = "${TemplatePrefix}$(FillZero -InStr ($StartNumber + $VDINameCount) -Len $NumberAlign)"
            if(Get-VM -Name $strTemplateName -Location $Container -ErrorAction:SilentlyContinue){
                Add-Log -Path $strLogFile -Value "$strTemplateName already exists." -Type Warning
                $Number++
            }else{
                Add-Log -Path $strLogFile -Value "$strTemplateName not exists, good to create."
                $TemplateNames += $strTemplateName
            }
        }
        $Task = & ".\VDI.Template.Creation.ps1" -VIServer $VIServer -Controller $Controller -MasterTemplate $MasterTemplate `
                                                -Cluster $Cluster -Datastore $Datastore `
                                                -Container $Container -TemplateNames $TemplateNames `
                                                -NumCPU $NumCPU -MemoryMB $MemoryMB -PassThru
        $TemplateJobs += $objTemplateJob.PSObject.Copy()
        $TemplateJobs[-1].'VIServer' = $VIServer
        $TemplateJobs[-1].'Container' = $Container
        $TemplateJobs[-1].'Cluster' = $Cluster
        $TemplateJobs[-1].'TaskID' = $Task
        $TemplateJobs[-1].'Controller' = $Controller
        $TemplateJobs[-1].'TemplateNames' = $TemplateNames -join "$([char]13 + [char]10)"
        $TemplateJobs[-1].'Number' = @($TemplateNames).Count
        $TemplateJobs[-1].'MasterTemplate' = $MasterTemplate
        $TemplateJobs[-1].'Datastore' = $Datastore
        Add-Log -Path $strLogFile -Value "Wait for $Task to complete."
        $TaskLoopTimeout = 10
        $TaskLoopCount = 0
        $TaskCaptured = $false
        do{
            Start-Sleep -Seconds 60
            # $TaskStart = Get-Task -Id $Task -ErrorAction:SilentlyContinue
            $TaskStart = Get-Task | ?{
                $_.Id -ieq $Task
            }
            if($TaskStart){$TaskCaptured = $true}
            if(++$TaskLoopCount -ge $TaskLoopTimeout -and !$TaskCaptured){
                Add-Log -Path $strLogFile -Value 'Cannot capture task until timed out, break the loop.' -Type Warning
                $TemplateJobs[-1].'Exception' = 'Not sure succeed'
                break
            }
        }while($TaskStart.State -eq 'Running' -or !$TaskCaptured)
        if($TaskStart.State -ne 'Success' -and $TaskCaptured){
            Add-Log -Path $strLogFile -Value "Task state is not success: $($TaskStart.State)." -Type Warning
            Add-Log -Path $strLogFile -Value "Try to get other tasks with abnormal state and messages."
            Get-Task | ?{$_.StartTime -ge $TaskStart.StartTime -and $TaskStart.State -ne 'Success'} | %{
                Add-Log -Path $strLogFile -Value "Found task: $($_.Id); Name: $($_.Name); State: $($_.State); Message: $($_.ExtensionData.Info.Error.LocalizedMessage)"
            }
            $TemplateJobs[-1].'Exception' = 'Failed'
        }else{
            $TemplateJobs[-1].'Exception' = 'Succeed'
        }
        Start-Sleep -Seconds 60
    }
}
if(!$TemplateJobs.Count){
    $TemplateJobs += $objTemplateJob.PSObject.Copy()
}
$TemplateJobs | Export-Csv -Path $TemplateCreationReport -NoTypeInformation -Delimiter "`t"

Add-Log -Path $strLogFile -Value "Start converting VDI."
$Clusters.Keys | Sort-Object | %{
    $PowershellInstances = @()
}{
    $Cluster = $_
    $JobsCompleteCount = 0
    $JobsInCluster = $Clusters[$Cluster]['Jobs'].Count
    Add-Log -Path $strLogFile -Value "Cluster: $Cluster; Jobs count: $JobsInCluster."
    foreach($Job in $Clusters[$Cluster]['Jobs']){
        Add-Log -Path $strLogFile -Value "Processing: $($JobsCompleteCount + 1)."
        if([int]((Get-Date).Hour) -ge 23){
            Add-Log -Path $strLogFile -Value "Current time: [$((Get-Date).ToString())] over 23:00, stop submit requests." -Type Warning
            $JobsLeft += $Clusters[$Cluster]['Jobs'][$JobsCompleteCount..$JobsInCluster] | %{
                $_['Exception'] = "Timeline hit"
            }
            $ExitCode = $ExitCode -bor 0x0004 # 0000 0000 0000 0100
            continue
        }

        $PowershellInstances += Start-Process -FilePath 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' -ArgumentList @($(
            '-File', 'VDI.Build.Core.ps1',
            "-VIServer", "$vCenter",
            "-User", "$($Job.User)",
            "-Prefix", "$Prefix",
            "-DesktopGroup", "$($Clusters[$Cluster][$($Job.Type)]['Group'])",
            "-DDCList", "`"$DDCList`"",
            "-SupportGroup", "$SupportGroup",
            "-Environment", "$($Job.Type)",
            "-ClusterToGo", "$Cluster",
            "-BizGroup", "'$BizADGroup'",
            "-IMGroup", "$IMADGroup",
            "-Folder", "$($Clusters[$Cluster][$($Job.Type)]['Folder'])",
            "-TemplateFolder", "$($Clusters[$Cluster][$($Job.Type)]['POOL'])",
            "-VDINameToGo", "$($Job.VDIName)",
            "-ExitCodeWriteTo", "$LocalDes\$strDate\ExitCodes.Core.txt", 
            "$($Job.AdvanceOptions)"
            ) | ?{$_}) -PassThru -RedirectStandardOutput "$strDate\Console.$($Job.User)_$($Job.VDIName).log" | Select-Object -Property Id, StartTime, @{N = "Args"; E = {(Get-WmiObject -Class Win32_Process -Filter "ProcessId = $($_.Id)").CommandLine}}
        if(!$?){
            Add-Log -Path $strLogFile -Value "New instance failed to start, cause:" -Type Error
            Add-Log -Path $strLogFile -Value $Error[0] -Type Error
        }else{
            Add-Log -Path $strLogFile -Value "New instance up."
        }
        Add-Log -Path $strLogFile -Value "Sleep 5 mins before process next."
        Start-Sleep -Seconds 300
    }
}

if(!$JobsLeft.Count){
    $JobsLeft += $objClusterJob.PSObject.Copy()
}
$JobsLeft | Export-Csv -Path $VMReportLeftFile -NoTypeInformation -Delimiter "`t"

if($ClearVIServer){
    Disconnect-VIServer -Server $vCenter -Confirm:$false -Force
}
Add-Log -Path $strLogFile -Value "Start to check pending jobs."

$Threshold = 8
Add-Log -Path $strLogFile -Value "Jobs loop threshold: $Threshold"
$LoopCount = 0
while($true){
    $LoopCount++
    Add-Log -Path $strLogFile -Value "Current loop count: $LoopCount"
    Add-Log -Path $strLogFile -Value "Powershell instances count: $($PowershellInstances.Count)"
    $PowershellInstances = @($PowershellInstances | %{
        if(Get-Process -Id $_.Id -ErrorAction:SilentlyContinue){$_}else{
            Add-Log -Path $strLogFile -Value "Instance completed, ID: $($_.Id), Start time: $($_.StartTime), CommandLine: $($_.Args)"
            Add-Log -Path $strLogFile -Value "Try to capture exit code for process $($_.Id)."
            if($ExitCodeCore = @(Select-String -Path "$LocalDes\$strDate\ExitCodes.Core.txt" -Pattern "^$($_.Id)\t(\d+)$")){
                Add-Log -Path $strLogFile -Value "Exit code: $($ExitCodeCore[0].Matches[0].Groups[1].Value)."
                if([int]($ExitCodeCore[0].Matches[0].Groups[1].Value)){
                    Add-Log -Path $strLogFile -Value "Abnormal exit code for powershell instance." -Type Warning
                    $ExitCode = $ExitCode -bor 0x0080 # 0000 0000 1000 0000
                }
            }else{
                Add-Log -Path $strLogFile -Value "No exit code captured for process $($_.Id)." -Type Warning
                $ExitCode = $ExitCode -bor 0x0100 # 0000 0001 0000 0000
            }
        }
    })
    if(-not $PowershellInstances){
        Add-Log -Path $strLogFile -Value "All instances completed."
        break
    }
    if($LoopCount -gt $Threshold){
        Add-Log -Path $strLogFile -Value "Threshold hit, this is abnormal." -Type Warning
        $ExitCode = $ExitCode -bor 0x0400 # 0000 0100 0000 0000
        $PowershellInstances | %{
            Add-Log -Path $strLogFile -Value "Pending instance: ProcessId = $($_.Id), StartTime = $($_.StartTime), CommandLine: $($_.Args)" -Type Warning
            Add-Log -Path $strLogFile -Value 'Send a kill singal to the pending instance'
            Stop-Process -Id $_.Id -Force -Confirm:$false
            if(!$?)
            {
                Add-Log -Path $strLogFile -Value 'Kill process failed, cause:' -Type Error
                Add-Log -Path $strLogFile -Value $Error[0] -Type Error
            }
        }
        break
    }
    Add-Log -Path $strLogFile -Value "Sleep 15 mins to next loop."
    Start-Sleep -Seconds (15*60)
}

Remove-PSSnapin VMware* -Confirm:$false
Add-Log -Path $strLogFile -Value "Script completed, exit with code $ExitCode."
Quit -ExitCode $ExitCode
