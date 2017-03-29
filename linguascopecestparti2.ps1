# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks flash version agaist the EXE file and installs new version if out-of-date

# Setup installer function
function Install-Exe ($airapp)
{
    #Set Air Installer Path
    if ($arch = $true) {
        $airinstaller = "C:\Program Files (x86)\Common Files\Adobe AIR\Versions\1.0\Adobe AIR Application Installer.exe"
    }
    else {
        $airinstaller = "C:\Program Files\Common Files\Adobe AIR\Versions\1.0\Adobe AIR Application Installer.exe"
    }
    
    # Run the EXE
    $installer = Start-Process -FilePath $airinstaller -ArgumentList "-silent -eulaAccepted -programMenu -desktopShortcut ""$airapp""" -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 1) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Warning -EventId 3001 -Message "Cest Parti 2 - The installer returned - Successful, but restart required for completion"
    }

    if ($installer.ExitCode -eq 2) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Usage error (incorrect arguments)"
    }
 
    if ($installer.ExitCode -eq 3) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Runtime not found"
    }

    elseif ($installer.ExitCode -eq 4) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Loading runtime failed"
    }

    elseif ($installer.ExitCode -eq 5) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Unknown error"
    }

    elseif ($installer.ExitCode -eq 6) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Installation canceled"
    }

    elseif ($installer.ExitCode -eq 7) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Installation failed"
    }

    elseif ($installer.ExitCode -eq 8) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Installation failed; update already in progress"
    }
    
    elseif ($installer.ExitCode -eq 9) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - The installer returned - Installation failed; application already installed"
    }

    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Cest Parti 2 - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - Unknown Error ocured"
    }
}

# Unistall
function Uninstall-exe ($EXE) {

   # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /uninstall" -Wait -Passthru
    
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
$airapp = ""

# This app dosnt throw the string into the unistall hive, it seems to go into another hive that I dont have trust in being the same in the wild.
# Get Current Version from registry
if ($arch -eq $true) {
    # Grab registry for office version (x64) 
    $Currentversionfile = Get-ItemProperty -Path "C:\Program Files (x86)\C'est Parti 2\C'est Parti 2.exe" -ErrorAction SilentlyContinue
}
else {
    # Grab current version (x32)
    $Currentversionfile = Get-ItemProperty -Path "C:\Program Files\C'est Parti 2\C'est Parti 2.exe" -ErrorAction SilentlyContinue
}



# Setup Auth
# Get Auth hive
$authreg = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS -ErrorAction SilentlyContinue

# Get Office Key
$cestpartireg = $authreg.CestParti2

# Find
if (!$cestpartireg) {
    # Not Allowed
    if ($cestpartireg -eq 0) {
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
   Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Cest Parti 2 - No registry key auth exists"
   exit
}

# Installed if not installed
if (!$Currentversionfile -and $cestpartireg -eq 1) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Cest Parti 2 - Currently not installed"
    # Run installer
    Install-Exe ($airapp)
    exit
}