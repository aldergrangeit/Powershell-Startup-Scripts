# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Check the version set and installs the exe as required.

# Setup installer function
function Install-Exe ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath msiexec -ArgumentList "/i $exe" -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 1011) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 - The installer returned - Allready an installer active"
    }
   
    elseif ($installer.ExitCode -eq 1024) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 - The installer returned - Unable to write files to directory"
    }
    
    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 - Unknown Error ocured"
    }
}

# Get Architecture
$arch = [Environment]::Is64BitProcess

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}

# Root Path to files
$rootpath = ""

# 32bit file
$32file = $rootpath + "vcredist.msi"

# 64bit filel
$64file = $rootpath + "x64\vcredist.msi"

# Get Current Flash Version from registry
if ($arch -eq $true) {
    # Get  registry
    $Currentversionregx32 = Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"{710f4c1c-cc18-4c49-8cbf-51240c89a1a2}" -ErrorAction SilentlyContinue
    $Currentversionregx64 = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"{ad8a2fa1-06e7-4b0d-927d-6e54b3d31028}" -ErrorAction SilentlyContinue
}
else
{
    # More modern versions of Air change their GUID, so check for that
    $Currentversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"{ad8a2fa1-06e7-4b0d-927d-6e54b3d31028}" -ErrorAction SilentlyContinue
}

# Set current flash version ready for checking later
$Currentversionx32 = ($Currentversionregx32.DisplayVersion)
$Currentversionx64 = ($Currentversionregx64.DisplayVersion)

# Installed if not installed
if ($arch -eq $true) {
    # Check for x64 existance
    if (!$Currentversionx64) {
        # Log
        Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 (x32) - Currently not installed"
        # Run Installer x64
        Install-Exe ($64file)
        exit
        }
        
     elseif (!$Currentversionx32) {
        #Log
        Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 (x64) - Currently not installed"
        # Run installer for x32 on x64 machine
        Install-Exe ($32file)
        exit
        }
}
else {
        #Log
        Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Visual Studio C++ Runtime 2005 - Currently not installed"
        # Run installer for x32 on x64 machine
        Install-Exe ($32file)
        exit
        }