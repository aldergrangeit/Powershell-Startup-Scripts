# Alder Grange School - Deployment System
# Script Version - V1
# Date Created - 16/01/2017
# Primary Task - Checks for the presence of DirectX 9 Runtimes and installs them if required.

# Setup installer function
function Install-Exe ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /silent" -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "DirectX 9 Runtimes - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "DirectX 9 Runtimes - Unknown Error ocured"
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
$EXE = ""

# Get Direct version to expect
$version = "9.29.952.3111"

# DLL Location
$dll = $env:windir + "\System32\D3DX9_43.dll"

# Get properties of the exists dll file
$dllproperties = Get-ItemProperty -Path $dll -ErrorAction SilentlyContinue

# Change the version txt to match format of how it comes from registry
$installedversion = $dllproperties.VersionInfo

# Get ProductVersion
$productversion = $installedversion.FileVersion

# Setup Auth
# Get Auth hive
$authreg = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS -ErrorAction SilentlyContinue

# Get Office Key
$directxreg = $authreg.DirectX9Runtimes

# Find
if ($directxreg) {
    # Not Allowed
    if ($directxreg -eq 1) {

        # Installed if not installed
        if (!$productversion -and $directxreg -eq 1) {
            # Log
            Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "DirectX 9 Runtimes - Currently not installed"
            # Run installer
            Install-Exe ($EXE)
            exit
        }
}

    else {
        # No Auth
        Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "DirectX 9 Runtimes - No registry key auth exists"
        exit
    }
}