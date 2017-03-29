# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks for authorisation to run installer, upon sucessfully being allowed checks if installed else installs. 

# Setup installer function
function Install-Exe ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /adminfile " -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 17301) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: General Detection error"
    }

    if ($installer.ExitCode -eq 17302) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: Applying patch"
    }
 
    if ($installer.ExitCode -eq 17303) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: Extracting file"
    }

    elseif ($installer.ExitCode -eq 17201) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: Creating temp folder"
    }

    elseif ($installer.ExitCode -eq 17022) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Success: Reboot flag set"
    }

    elseif ($installer.ExitCode -eq 17023) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: User cancelled installation"
    }

    elseif ($installer.ExitCode -eq 17024) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: Creating folder failed"
    }

    elseif ($installer.ExitCode -eq 17025) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Patch already installed"
    }
    
    elseif ($installer.ExitCode -eq 17026) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Patch already installed to admin installation"
    }

    elseif ($installer.ExitCode -eq 17027) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Installation source requires full file update"
    }

    elseif ($installer.ExitCode -eq 17028) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - No product installed for contained patch"
    }

    elseif ($installer.ExitCode -eq 17029) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Patch failed to install"
    }

    elseif ($installer.ExitCode -eq 17030) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Detection: Invalid CIF format"
    }

    elseif ($installer.ExitCode -eq 17031) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Detection: Invalid baseline"
    }

    elseif ($installer.ExitCode -eq 17034) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Error: Required patch does not apply to the machine"
    }

    elseif ($installer.ExitCode -eq 17038) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - You do not have sufficient privileges to complete this installation for all users of the machine. Log on as administrator and then retry this installation."
    }
    
    elseif ($installer.ExitCode -eq 17044) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - Installer was unable to run detection for this package."
    }

    elseif ($installer.ExitCode -eq 17048) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - The installer returned - This installation requires Windows Installer 3.1 or greater."
    }

    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Office 2013 Lync - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - Unknown Error ocured"
    }
}

# Unistall
function Uninstall-Exe ($EXE) {

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
$EXE = ""

# Get EXE version
$version = (Get-Item $EXE).VersionInfo

# Get Current Version from registry
if ($arch -eq $true) {
    # Grab registry for office version (x64) 
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Lync -ErrorAction SilentlyContinue
}
else {
    # Grab current version (x32)
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Office\15.0\Lync -ErrorAction SilentlyContinue
}

# Change the version txt to match format of how it comes from registry
$exeversion = $version.FileVersion

# Set current flash version ready for checking later
$Currentversion = $Currentversionreg.Path

# Setup Auth
# Get Auth hive
$authreg = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS -ErrorAction SilentlyContinue

# Get Office Key
$office2013lyncreg = $authreg.office2013lync

# Find
if ($officereg) {
    # Not Allowed
    if ($office2013lyncreg -eq 0) {
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
   Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Office 2013 Lync - No registry key auth exists"
   exit
}

# Installed if not installed
if (!$Currentversion -and $officereg -eq 1) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Office 2013 Lync - Currently not installed"
    # Run installer
    Install-Exe ($EXE)
    exit
}