<#    
    This script is designed to implement a faster workflow for bulk printing .pdf docs.
    It is also designed to handle .zip archives, provided they contain .pdf docs.

    I wrote it out of anger at Adobe and Microsoft who broke Outlook's Quick Print in September 2023,
    and have not yet fixed it for my end users, even though they say they absolutely have. Liars.

    Drop .pdf files into $pdfFolder location, and .zip files into $zipFolder location.

    Unknown filetypes will be discarded, as will subfolders. So, be careful.
#>

<#
    Required modules:
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
        Install-Module -Name Recycle -MinimumVersion 1.5.0 -Force
#>

# GLOBAL VARIABLES
$adobe="C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

# Folder and logfile paths
$rootDirectory = "$env:SYSTEMDRIVE\Temp\AutoPrintOnDrop"
$pdfFolder = "$rootDirectory\PrintPDFs"
$zipFolder = "$rootDirectory\PrintZIPs"
$logFolder = "$rootDirectory\Logs"

$logfile = "$logFolder\infolog.log"
$errfile = "$logFolder\errlog.log"

#####################
# Logging Functions #
#####################
Function Get-TimeStamp
{
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

Function Write-InfoLog
{
    Param ([string]$logstring)
    Add-Content $logfile -value "$(Get-TimeStamp) $logstring"
}

Function Write-ErrorLog
{
    Param ([string]$errstring)
    Add-Content $errfile -value "$(Get-TimeStamp) $errstring"
}


#########################################
# PRINTER INFORMATION - Globally scoped #
#########################################

# Get default printer using gwmi - made this dynamic
$defaultPrinter = $($(Get-WmiObject win32_printer -computername $env:COMPUTERNAME) | Where-Object {$_.Default -eq $true}).DeviceID

# gets driver name dynamically
$drivername = $(Get-Printer -Name $defaultPrinter).DriverName
# gets port name dynamically
$portname = $(Get-Printer -Name $defaultPrinter).PortName

######################################
# Test for Setup and Setup If Needed #
######################################

function New-AutoprintDirectory
{
    # Do folders exist? If not, create them
    # Simple structure, could probably be cleaned up
    if (!(Test-Path -Path $rootDirectory)) {
        New-Item -ItemType directory -Path $rootDirectory
    }
    if (!(Test-Path -Path $pdfFolder)) {
        New-Item -ItemType directory -Path $pdfFolder
    } 
    if (!(Test-Path -Path $zipFolder)) {
        New-Item -ItemType directory -Path $zipFolder
    }
    if (!(Test-Path -Path $logFolder)) {
        New-Item -ItemType directory -Path $logFolder

        New-Item -ItemType file -Path $logfile
        New-Item -ItemType file -Path $errfile
    }
}

##############################################
# Actions for the FileSystemWatchers to Take #
##############################################

# Action for $acctPrintWatcher object to perform when file is dropped
$processPDF = 
{
    # Pulls filepath from the event that triggered processing, then pulls the filename since full path may change in processing
    $inboundFile = Split-Path $event.SourceEventArgs.FullPath -leaf
    # Pulls file extension for decision tree
    $fileExt = [System.IO.Path]::GetExtension($inboundFile)
    
    # Decision structure based on filetype, moves unknown file types to $Recycle.Bin
    # This is NOT the same structure as $processZIP
    if (Test-Path -Path $pdfFolder\$inboundFile -PathType Leaf) {
        if ($fileExt -eq ".pdf") {
            # print file
            Write-InfoLog "Printing $pdfFolder\$inboundFile"
            $arglist = ' /S /T "' + "$pdfFolder\$inboundFile" + '" "' + $defaultPrinter + '" "' + $drivername + " " + $portname
            Start-Process $adobe -ArgumentList $arglist -WindowStyle Minimized
            # allow Adobe time to process print - can't track files by subprocess
            Start-Sleep -Seconds 10
            
            Write-InfoLog "Printed $inboundFile to $defaultPrinter on $portName"

            # BUT the removal doesn't work if the file is locked, so attempt to remove the file until it unlocks
            # remove file
            while (Test-Path -Path $pdfFolder\$inboundFile) {
                Write-InfoLog "Attempting removal of $inboundFile..."
                Remove-Item -Path $pdfFolder\$inboundFile
                Start-Sleep -seconds 1 # delay 1 second each iteration to reduce processing load
            }
            Write-InfoLog "Deleted $inboundFile"
        } elseif ($fileExt -eq ".zip") {
            # move to ZIP folder for processing
            Move-Item $pdfFolder\$inboundFile -Destination $zipFolder\$inboundFile
            Write-InfoLog "Moved $inboundFile to $zipFolder\$inboundFile"
        } else {
            Remove-ItemSafely $pdfFolder\$inboundFile
            Write-ErrorLog "$pdfFolder\$inboundFile is not a recognized filetype. Moved to Recycle.Bin."
        }
     
     } else {
        # Move to $Recycle.Bin - can't handle subfolders right now
        Remove-ItemSafely -Path $pdfFolder\$inboundFile -Recurse -Force
        Write-ErrorLog "Moved $inboundFile to Recycle.Bin. Cannot handle subfolders."
     }

}

$processZIP = 
{
    # Pulls filepath from the event that triggered processing, then pulls the filename since full path may change in processing
    $inboundFile = Split-Path $event.SourceEventArgs.FullPath -leaf
    # Pulls file extension for decision tree
    $fileExt = [System.IO.Path]::GetExtension($inboundFile)

    # Decision structure based on filetype, handles subfolders and unknown filetypes by moving to hidden Trash directory
    # This is NOT the same structure as $processPDF
    if (Test-Path -Path $zipFolder\$inboundFile -PathType Leaf) {
        if ($fileExt -eq ".pdf") {
            # move file to PDF folder
            Move-Item $zipFolder\$inboundFile -Destination $pdfFolder\$inboundFile
            Write-InfoLog "Moved $inboundFile to $pdfFolder\$inboundFile"

        } elseif ($fileExt -eq ".zip") {
            # unzip and send files to $pdfFolder
            Expand-Archive $zipFolder\$inboundFile -DestinationPath $pdfFolder
            Write-InfoLog "Extracted files from $inboundFile to $pdfFolder"

            Start-Job -ScriptBlock {Expand-Archive $zipFolder\$inboundFile -DestinationPath $pdfFolder} -name "Unzip"
            Wait-Job -Name "Unzip"
        
            # delete .zip after waiting for unzip
            Remove-Item $zipFolder\$inboundFile
        } else {
            Remove-ItemSafely $zipFolder\$inboundFile -Force
            Write-ErrorLog "$zipFolder\$inboundFile is not a recognized filetype. Moved to Recycle.Bin."
        }
     } else {
        # Move to $Recycle.Bin - can't handle subfolders right now
        Remove-ItemSafely $zipFolder\$inboundFile -Recurse -Force
        Write-InfoLog "Moved $inboundFile to Recycle Bin. Cannot handle subfolders."
     }
}


######################
# Main Script Driver #
######################
# call function to spin up folders and logfiles
New-AutoprintDirectory

# test Adobe path
if (!(Test-Path -Path $adobe -PathType Leaf)) {
    # Adobe not found
    Write-ErrorLog "Supplied Adobe path not found: $adobe"
} else {
    # Adobe is fine
    # create a filesystem watcher to report on new files in $pdfFolder
    $acctPDFWatcher = New-Object System.IO.FileSystemWatcher
    $acctPDFWatcher.Path = $pdfFolder
    $acctPDFWatcher.EnableRaisingEvents = $true
    # create a filesystem watcher to report on new files in $zipFolder
    $acctZipWatcher = New-Object System.IO.FileSystemWatcher
    $acctZipWatcher.Path = $zipFolder
    $acctZipWatcher.EnableRaisingEvents = $true

    # Register both watchers
    Register-ObjectEvent $acctPDFWatcher 'Created' -Action $processPDF
    Write-InfoLog "PDF Watcher registered"

    Register-ObjectEvent $acctZipWatcher 'Created' -Action $processZIP
    Write-InfoLog "ZIP Watcher registered."
}
