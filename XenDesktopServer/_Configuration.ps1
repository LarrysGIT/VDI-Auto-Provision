
## Variables for Starter.ps1 and VDI.Build.ps1
## Both script will invoke this script to setup their variables.
## Run the functions to make variables define comes true.

# All
function Define-CommonVariables{
    $Script:LocalDes = '.'    # Where log folder to create, "." means create with script.
    $Script:RemoteDes = "\\sharepoint\harePoint_Scripts\XX VDI Creation"    # Where to read data, also, to upload data.
    $Script:Date = Get-Date    # No need to change, get the time when script start to run.
    $Script:strDate = $Date.ToString('yyyy-MM-dd')    # No need to change, get date like 2013-10-17
    $Script:vCenter = 'XXvCenter'    # vCenter in $POD
    $Script:POD = 'XX POD1'    # Tell script which POD it belong to.
}

# Starter.ps1
function Define-StarterVariables{
    $Script:strLogFile = "$LocalDes\$strDate\Starter.Verbose.log"
    $Script:EmailTo = "larrysong@company.com"    # Send warning email to.
    #$Script:EmailTo = "larrysong@company.com"   # Debugger
    $Script:EmailFrom = "$($env:COMPUTERNAME)@company.com"   # No need to change, current computer name plus @company.com.
    $Script:EmailSMTPServer = "smtp.company.com"    # SMTP server to contact with.
    $Script:EmailSubject = "[Warning] VDI automatic build with abnormal exit code: "    # Email Subject, exit code will auto plus to the behind.
    $Script:EmailContent = (Get-Content -Path 'Template.html.txt') -join ""    # No need to change, it's html format template.
}

# VDI.Build.ps1
function Define-VDIBuildVariables{
    $Script:strLogFile = "$LocalDes\$strDate\Build.Log.Verbose.log"
    $Script:RawFile = "$RemoteDes\$strDate\Exports\New VDI Requests_Export_Raw_0.CSV"
    $Script:GolbalVDILimite = 500
    $Script:Prefix = "XXV"
    $Script:WaitForPick = ""
    $Script:DDCList = "XXServer01.company.com XXServer01.company.com"
    $Script:SupportGroup = "g-XX-Company_Support"
    $Script:IMADGroup = "g-XX-Company_Support"
    $Script:BizADGroup = "g-XX-Company_Support"
    $Script:PickClusterByDS = $true

    $Script:Clusters_VMFilter = '^XXV[A-Z]\d{6}'
    $Script:Clusters = @{
        'XXVMwareCluster_1'  = @{
            'Proportion'           = 300;
            'Controller'           = 'NetApp01.company.com';
            'Type1' = @{
                'VolumeControl'    = $false;
                'CapacityMultiple' = 0.6;
                'TemplateCreation' = $true;
                'Template'         = 'XXTemplate01';
                'POOL'             = 'XX-POOL1';
                'TemplatePreserve' = 10;
                'Folder'           = 'XX-Folder1';
                'Group'            = 'XX-Group1';
                'Datastore'        = 'XX-Datastore1';
                'NumCPU'           = 2;
                'MemoryMB'         = 2048;
            };
            'Type2' = @{
                'VolumeControl'    = $false;
                'CapacityMultiple' = 0.6;
                'TemplateCreation' = $true;
                'Template'         = 'XXTemplate02';
                'POOL'             = 'XX-POOL2';
                'TemplatePreserve' = 10;
                'Folder'           = 'XX-Folder2';
                'Group'            = 'XX-Group2';
                'Datastore'        = 'XX-Datastore2';
                'NumCPU'           = 2;
                'MemoryMB'         = 4096;
            };
        };
        'XXVMwareCluster_2'  = @{
            'Proportion'           = 300;
            'Controller'           = 'NetApp01.company.com';
            'Type1' = @{
                'VolumeControl'    = $false;
                'CapacityMultiple' = 0.6;
                'TemplateCreation' = $true;
                'Template'         = 'XXTemplate01';
                'POOL'             = 'XX-POOL1';
                'TemplatePreserve' = 10;
                'Folder'           = 'XX-Folder1';
                'Group'            = 'XX-Group1';
                'Datastore'        = 'XX-Datastore1';
                'NumCPU'           = 2;
                'MemoryMB'         = 2048;
            };
            'Type2' = @{
                'VolumeControl'    = $true;
                'CapacityMultiple' = 0.6;
                'TemplateCreation' = $true;
                'Template'         = 'XXTemplate02';
                'POOL'             = 'XX-POOL2';
                'TemplatePreserve' = 10;
                'Folder'           = 'XX-Folder2';
                'Group'            = 'XX-Group2';
                'Datastore'        = 'XX-Datastore2';
                'NumCPU'           = 2;
                'MemoryMB'         = 4096;
            };
        };
    }

    $Script:VMReportLeftFile = "$RemoteDes\$strDate\Imports\${POD}_VDI_Build_Left.csv"
    $Script:TemplateCreationReport = "$RemoteDes\$strDate\Imports\${POD}_TemplateCreationReport.csv"
}

# VDI.Rebuild.ps1
function Define-VDIRebuildVariables{
    $Script:Enable = $true
    $Script:strLogFile = "$LocalDes\$strDate\Rebuild.Log.Verbose.log"
    $Script:RawFile = "$RemoteDes\$strDate\Exports\VDI Rebuild_Export_Raw_0.CSV"
    $Script:VMReportLeftFile = "$RemoteDes\$strDate\Imports\${POD}_VDI_Rebuild_Left.csv"
    $Script:VMReportProcessedFile = "$RemoteDes\$strDate\Imports\${POD}_VDI_Rebuild_Processed.csv"
    $Script:Type = 'Type1'
}

# VDI.Data.Report.ps1
function Define-VDIDataReportVariables{
    $Script:Wait = $true
    $Script:strLogFile = "$LocalDes\$strDate\DataReport.Log.Verbose.log"
    $Script:WaitUntil = "23:05"
    $Script:VMReportFile = "$RemoteDes\$strDate\Imports\$POD.txt"

    $Script:CapacityReportType = 'Type1'
    $Script:CapacityReport = "$RemoteDes\$strDate\Imports\${POD}_${CapacityReportType}_Capacity.txt"
    $Script:CapacityWarning = 0.85
    $Script:CapacityAlert = 0.9

    $Script:CapacityMonitoringEnable = $false
    
    $Script:CapacityWarningMailSubject = '[Attention] VDI POD capacity warning!'
    $Script:CapacityWarningMailTo = $EmailTo

    $Script:CapacityAlertMailSubject = '[Attention] VDI POD capacity alert!'
    $Script:CapacityAlertMailTo = $EmailTo
}

# VDI.ReclaimStorage.ps1
function Define-ReclaimSpaceVariables{
    $Script:Enable = $false
    $Script:strLogFile = "$LocalDes\$strDate\Reclaim.Log.Verbose.log"
    $Script:Type = 'Type1'
    $Script:StillMoveWhensDeleteFailed = $true
    $Script:ReclaimCount = 1
    $Script:StartAfter = '21:00'
    $Script:StopBefore = '23:00'
}

# VDI.MoveBack.ps1
function Define-MoveBackVariables{
    $Script:Enable = $false
    $Script:strLogFile = "$LocalDes\$strDate\MoveBack.Log.Verbose.log"
    $Script:Type = 'Type1'
    $Script:StartAfter = '21:00'
    $Script:StopBefore = '23:00'
    $Script:MoveBackAfter = 1 # Days / No guarantee because time probably not enough
}
