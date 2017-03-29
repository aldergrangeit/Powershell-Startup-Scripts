# Alder Grange School - Deployment System
# Script Version - V1
# Date Created - 16/01/2017
# Primary Task - Checks if Adobe CS4 is installed and if not installs or vice versa if needed to be uninstalled.

# Setup installer/uninstaller function
function Invoke-Run ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $EXE -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 9999) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Catastrophic error"
    }

    if ($installer.ExitCode -eq 1) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unable to parse command line"
    }
 
    if ($installer.ExitCode -eq 2) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unknown user interface mode specified"
    }

    elseif ($installer.ExitCode -eq 3) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unable to initialize ExtendScript"
    }

    elseif ($installer.ExitCode -eq 4) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - User interface workflow failed"
    }

    elseif ($installer.ExitCode -eq 5) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unable to initialize user interface workflow"
    }

    elseif ($installer.ExitCode -eq 6) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Silent workflow completed with errors"
    }

    elseif ($installer.ExitCode -eq 7) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unable to complete the Silent workflow"
    }
    
    elseif ($installer.ExitCode -eq 8) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Exit and restart"
    }

    elseif ($installer.ExitCode -eq 9) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unsupported operating system version"
    }

    elseif ($installer.ExitCode -eq 10) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Unsupported file system"
    }

    elseif ($installer.ExitCode -eq 11) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Another instance running"
    }

    elseif ($installer.ExitCode -eq 12) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - CAPS integrity error"
    }

    elseif ($installer.ExitCode -eq 13) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Media optimization failed"
    }

    elseif ($installer.ExitCode -eq 14) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - Failed due to insufficient privileges"
    }

    elseif ($installer.ExitCode -eq -1) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - The installer returned - The AdobeUberinstaller failed (before launching the installer)"
    }

    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Adobe Creative Suite 4 - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - Unknown Error ocured"
    }
}


# Get arct
$arch = [Environment]::Is64BitProcess

# Get windows version
$os = [Environment]::OSVersion.Version.Major

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}


# Installer EXE
$InstallEXE = ""

# Uninstaller EXE
$UninstallerEXE = ""

# Get Current Version from registry
if ($arch -eq $true) {
    # Grab registry for office version (x64) 
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"{A2881E09-38DB-4F79-9135-00FDA01768A7}" -ErrorAction SilentlyContinue
}
else {
    # Grab current version (x32)
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"{A2881E09-38DB-4F79-9135-00FDA01768A7}" -ErrorAction SilentlyContinue
}

# Set current flash version ready for checking later
$Currentversion = $Currentversionreg.DisplayVersion

# Setup Auth
# Get Auth hive
$authreg = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS -ErrorAction SilentlyContinue

# Get Office Key
$adobecs4reg = $authreg.adobecs4

# Find
if ($adobecs4reg) {
    # Not Allowed
    if ($adobecs4reg -eq 0) {
        # Check if installed
        if ($Currentversion) {
            # Uninstall if found to be installed
            Invoke-Run($UninstallerEXE)
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
   Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Adobe Creative Suite 4 - No registry key auth exists"
   exit
}

# Installed if not installed
if (!$Currentversion -and $adobecs4reg -eq 1) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Adobe Creative Suite 4 - Currently not installed"
    # Run installer
    Invoke-Run ($InstallEXE)
    exit
}
else {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Adobe Creative Suite 4 - Currently installed"
}