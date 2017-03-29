# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks flash version agaist the EXE file and installs new version if out-of-date

# Setup installer function
function Install-Flash ($FlashEXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $FlashEXE -ArgumentList " -install" -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 1003) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - The installer returned - Invalid argument passed to installer"
    }

    if ($installer.ExitCode -eq 1011) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - The installer returned - Allready an installer active"
    }
 
    if ($installer.ExitCode -eq 1013) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - The installer returned - Trying to install older revision"
    }

    elseif ($installer.ExitCode -eq 1022) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - The installer returned - Does not have admin permissions"
    }

    elseif ($installer.ExitCode -eq 1025) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - The installer returned - Existing Player in use"
    }

    elseif ($installer.ExitCode -eq 1032) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - The installer returned - ActiveX registration failed"
    }
    
    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - Unknown Error ocured"
    }
}

# Get windows version
$os = [Environment]::OSVersion.Version.Major

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}

# Check for Windows 10
if ($os -ge "10") {
Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "This script will not support Windows 10 or above as flash is bundled"
exit
}

# Flash Installer EXE
$flashEXE = ""

# Get EXE version
$flashversion = (Get-Item $flashEXE).VersionInfo

# Get Current Flash Version from registry
$CurrentFlashversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Macromedia\FlashPlayerActiveX -ErrorAction SilentlyContinue

# Change the version txt to match format of how it comes from registry
$exeflashversion = ($flashversion.FileVersion -replace ',',".")

# Set current flash version ready for checking later
$CurrentFlashversion = ($CurrentFlashversionreg.Version -replace ',',".")

# Installed if not installed
if (!$CurrentFlashversion) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - Currently not installed"
    # Run installer
    Install-Flash ($FlashEXE)
    exit
}

# Do comparison agaist the EXE file version if found to be is
if ([version]$exeflashversion -gt [Version]$CurrentFlashversion) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - This computer is out-of-date, current version is $exeflashversion. The computer is currently running $CurrentFlashversion. Invoking installer!"
    # Inovke installer
    Install-Flash ($flashEXE)
    exit
}
elseif ([Version]$CurrentFlashversion -gt [version]$exeflashversion ) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Flash - This computer is running a higher version than the exe file $exeflashversion. The computer is currently running $CurrentFlashversion."
    # Inovke installer
    Install-Flash ($flashEXE)
    exit
}
else {
    # 
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Flash - Up-to-date"
    exit
}