# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks flash version agaist the EXE file and installs new version if out-of-date

# Setup installer function
function Install-Exe ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /qb /I " -Wait -Passthru
    
    # Exit code handler

    if ($installer.ExitCode -eq 1641) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "AutoDesk Revit 2016 - The installer returned - The installer has initiated a restart. This message is indicative of a success."
    }

    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "AutoDesk Revit 2016 - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "AutoDesk Revit 2016 - Unknown Error ocured"
    }
}

# Unistall
function Uninstall-exe ($EXE) {

   # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /uninstall" -Wait -Passthru
    
}

# Get arct
$arch = [Environment]::Is64BitProcess

# If not 64 then quit with reason
if ($arch -eq $false) {
   # Exit with a reason if we are not in 64bit
   Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "AutoDesk Revit 2016 - Unable to install as installer is not x32 compatible"
   exit
}

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}


# Installer EXE
$EXE = ""

# Get Current Version from registry
if ($arch -eq $true) {
    # Grab registry for office version (x64) 
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"{7346B4A0-1600-0510-0000-705C0D862004}" -ErrorAction SilentlyContinue
}

# Setup Auth
# Get Auth hive
$authreg = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS -ErrorAction SilentlyContinue

# Get Office Key
$gotimeauthreg = $authreg.AutoDeskRevit2016

# Find
if (!$gotimeauthreg) {
    # Not Allowed
    if ($gotimeauthreg -eq 0) {
        # Check if installed
        if ($Currentversion) {
            # Uninstall if found to be installed
            Function-Uninstall($exe)
            exit
        }

     
        else {
            # Else exit  
            exit
        }
        
    }
}
else {
   # No Auth
   Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "AutoDesk Revit 2016 - No registry key auth exists"
   exit
}

# Installed if not installed
if (!$Currentversionreg -and $gotimeauthreg -eq 1) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "AutoDesk Revit 2016 - Currently not installed"
    # Run installer
    Install-Exe ($EXE)
    exit
}