# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks air version agaist the EXE file and installs new version if out-of-date

# Setup installer function
function Install-Air ($AirEXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $AirEXE -ArgumentList " -silent -eulaAccepted" -Wait -Passthru
    
    # Exit code handler
    if ($installer.ExitCode -eq 1011) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - The installer returned - Allready an installer active"
    }
   
    elseif ($installer.ExitCode -eq 1024) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - The installer returned - Unable to write files to directory"
    }
    
    elseif ($installer.ExitCode -eq 0) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - Successfully installed"
    exit
    }

    else {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - Unknown Error ocured"
    }
}

# Get windows version
$os = [Environment]::OSVersion.Version.Major

# Get Architecture
$arch = [Environment]::Is64BitProcess

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}

# Air Installer EXE
$AirEXE = ""

# Get EXE version
$Airversion = (Get-Item $AirEXE).VersionInfo

# Get Current Flash Version from registry
if ($arch) {
    # More modern versions of Air change their GUID, so check for that
    $CurrentAirversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"{63B5DA5A-477B-438D-A6A0-118787A4C71B}" -ErrorAction SilentlyContinue
    # Check for older version
    if (!$CurrentAirversionreg) {
        $CurrentAirversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"{0274D240-4D1D-4FDA-9A36-09F0BECD288F}" -ErrorAction SilentlyContinue
    }
}
else
{
    # More modern versions of Air change their GUID, so check for that
    $CurrentAirversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"{63B5DA5A-477B-438D-A6A0-118787A4C71B}" -ErrorAction SilentlyContinue
    # Check for older version
    if (!$CurrentAirversionreg) {
        $CurrentAirversionreg = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"{0274D240-4D1D-4FDA-9A36-09F0BECD288F}" -ErrorAction SilentlyContinue
    }
}

# Change the version txt to match format of how it comes from registry
$exeairversion = ($airversion.FileVersion -replace ',',".")

# Set current flash version ready for checking later
$CurrentAirversion = ($CurrentAirversionreg.DisplayVersion -replace ',',".")

# Installed if not installed
if (!$CurrentAirversion) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - Currently not installed"
    # Run installer
    Install-Air ($AirEXE)
    exit
}

# Do comparison agaist the EXE file version
if ([version]$exeairversion -gt [Version]$Currentairversion) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - This computer is out-of-date, current version is $exeairversion. The computer is currently running $Currentairversion. Invoking installer!"
    # Run Installer
    Install-Air ($AirEXE)
    exit
}
elseif ([Version]$Currentairversion -gt [version]$exeairversion) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Air - Exe is older than the Air version on this computer"
    exit
}

else {
Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Air - Up-to-date"
exit
}