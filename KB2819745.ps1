#
#

function Install-Hotfix($exefile) {

    # Run the EXE
    Start-Process -FilePath $exefile -ArgumentList " /quiet" -Wait -Passthru
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "KB2819745 - Invoked the installer, which will require this computer to restart"

}

function Install-Framework($frameworkexe) {
    # Run the EXE
    $installer = Start-Process -FilePath $frameworkexe -ArgumentList " /q /norestart /ChainingPackage ADMINDEPLOYMENT" -Wait -Passthru
    # Installer Log
    if ($installer.ExitCode -eq 1602) {
        Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Framework - The user canceled installation"
    }
    if ($installer.ExitCode -eq 1603) {
        Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Framework - The user canceled installation"
    }
    if ($installer.ExitCode -eq 5100) {
        Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Framework - The user's computer does not meet system requirements."
    }
    if ($installer.ExitCode -eq 0 -or $installer.ExitCode -eq 1641 -or $installer.ExitCode -eq 3010) {
        Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Framework - Installed Successfully."
    }
    else {
        Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Framework - Unknown Error has occurred."
    }

}

# Initialise Event log
$event = Get-EventLog -Source AGSDS -LogName Application -Newest 1 -ErrorAction SilentlyContinue

# Check if the event source exists
if ($event.count -lt 1 -or !$event) {
    # Create the event source
    New-EventLog -Source AGSDS -LogName Application
}

# Framework Exe
$frameworkexe = ""

# Framework Version
$newframeworkversion = (Get-Item $frameworkexe).FileVersion

# Windows Management Framework V4 Folder
$wmfpath = ""

# X64 file
$64file = "Windows6.1-KB2819745-x64-MultiPkg.msu"

# x32 file
$32file = "Windows6.1-KB2819745-x86-MultiPkg.msu"

# KB id
$updateid = "KB2819745"

# Set the arch 
$arch = [IntPtr]::Size * 8

# Get windows version
$os = [Environment]::OSVersion.Version.Major

# Windows 10 allready comes with V5
if ($os -ge "10") {
    exit
}

# Get powershell version
$powershellv = $PSVersionTable.PSVersion.Major

# Framework version
$frameworkversion = $PSVersionTable.CLRVersion.Major

# Check for registry that determins clr version
if ($arch -eq 64 -and $newframeworkversion -lt 3) {
    # 64bit check
    $x64reg = Get-ItemProperty -Path HKLM:SOFTWARE\Microsoft\.NETFramework -ErrorAction SilentlyContinue
    $x64uselastest = ($x64reg.OnlyUseLatestCLR)
    $x32reg = Get-ItemProperty -Path HKLM:SOFTWARE\Microsoft\.NETFramework -ErrorAction SilentlyContinue
    $x32uselastest = ($x32reg.OnlyUseLatestCLR)
    if (!$x64uselastest) {
        New-ItemProperty -Path HKLM:SOFTWARE\Microsoft\.NETFramework -name OnlyUseLatestCLR -Value 1 -PropertyType DWORD -Force
    }
    if (!$x32uselastest) {
        New-ItemProperty -Path HKLM:SOFTWARE\Wow6432Node\Microsoft\.NETFramework -name OnlyUseLatestCLR -Value 1 -PropertyType DWORD -Force
    }
}
else {
    #32bit check
    $x32reg = Get-ItemProperty -Path HKLM:SOFTWARE\Microsoft\.NETFramework -ErrorAction SilentlyContinue
    $x32uselastest = ($x32reg.OnlyUseLatestCLR)
        if (!$x32uselastest) {
        New-ItemProperty -Path HKLM:SOFTWARE\Microsoft\.NETFramework -name OnlyUseLatestCLR -Value 1 -PropertyType DWORD -Force
    }
}

# This is if the above dosnt work, we will try and install the latest version
if ($frameworkversion -lt 3) {
    # Display error and exit if not
    Write-EventLog -LogName Application -Source AGSDS -EntryType Warning -EventId 3001 -Message "$updateid - Unable to install due to wrong framework version which requires version 3.5, attempting install of version $newframeworkversion with the current being version $frameworkversion)"
    # Invoke Framework Installer
    Install-Framework($frameworkexe)
    exit
}

# Run commands depending on arch
if ($arch -eq "64") {
    if ($powershellv -lt "4" -and $frameworkversion -ge 3) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "$updateid - x64 not found! Invoking Installer"
    # Set exe file to correct one
    $exefile = $wmfpath+$64file
    # Install Hotfix
    Install-Hotfix ($exefile)
    exit
   }
}
else {
   if ($powershellv -lt "4" -and $frameworkversion -ge 3) {
   Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "$updateid - x32 not found! Invoking Installer"
    # Set exe file to correct one
    $exefile = $wmfpath+$32file
    # Install Hotfix
    Install-Hotfix ($exefile)
    exit
   }
}