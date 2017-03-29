# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks flash version agaist the EXE file and installs new version if out-of-date

# Setup installer function
function Install-Exe ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " -ouser -opwd -mng yes -s -xp " -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - Installation was successful."
    }

    if ($installer.ExitCode -eq 1) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - A command line parameter value is missing or an unrecognized parameter was specified."
    }
 
    if ($installer.ExitCode -eq 2) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - Verification of the AutoUpdate package failed. The package files did not match the manifest."
    }

    elseif ($installer.ExitCode -eq 3) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - AutoUpdate was already installed."
    }

    elseif ($installer.ExitCode -eq 4) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - AutoUpdate does not support this operating system."
    }

    elseif ($installer.ExitCode -eq 5) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate- The installer returned - AutoUpdate requires Internet Explorer 5.0 or above; the system does not have this version of IE."
    }

    elseif ($installer.ExitCode -eq 6) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - Installation of AutoUpdate failed."
    }
	
	elseif ($installer.ExitCode -eq 7) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - Some file that was required could not be found e.g. an RMS configuration file or Sophos AutoUpdate.msi"
    }
	
	elseif ($installer.ExitCode -eq 99) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - The installer returned - Some other error occurred."
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Sophos AutoUpdate - Unknown Error ocured"
    }
}

# Unistall
function Uninstall-exe ($EXE) {

   # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /uninstall" -Wait -Passthru
    
}

# Get arct
$arch = [Environment]::Is64BitProcess

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}


# Sophos path
$sophospath = ""

# Installer EXE
$EXE = "setup.exe"

# Autoupdate Version
$autoupdateexe = "sau\SophosAlert.exe"

# Set install path
$installer = $sophospath + $exe

# Get EXE version
$version = (Get-Item $sophospath$autoupdateexe).VersionInfo

# Get Current Version from registry
if ($arch -eq $true) {
    # Grab registry for office version (x64) 
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Sophos\AutoUpdate -ErrorAction SilentlyContinue
}
else {
    # Grab current version (x32)
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Sophos\AutoUpdate -ErrorAction SilentlyContinue
}

# Get file version
$exeversion = $version.FileVersion

# Set current version
$Currentversion = $Currentversionreg.ProductVersion

# Installed if not installed
if (!$Currentversion) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Sophos AutoUpdate $exeversion - Currently not installed"
    # Run installer
    Install-Exe ($installer)
    exit
}

# Do comparison agaist the EXE file version if found to be is
if ([version]$exeversion -gt [Version]$Currentversion) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Sophos AutoUpdate - This computer is out-of-date, current version is $exeversion. The computer is currently running $Currentversion. Invoking installer!"
    # InVoke installer
    Install-Exe ($installer)
    exit
}
else {
    # 
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Sophos AutoUpdate $exeversion - Up-to-date"
    exit
}