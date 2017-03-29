#
#

function Install-Hotfix($exefile) {

 # Run the EXE
    Start-Process -FilePath $exefile -ArgumentList " /quiet" -Wait -Passthru
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "KB3134760 - Invoked installer, this computer will restart to complete this update"

}

# Framework version
$frameworkversion = $PSVersionTable.CLRVersion

# Windows Management Framework V4 Folder
$wmfpath = ""

# X64 file
$64file = "Win7AndW2K8R2-KB3134760-x64.msu"

# x32 file
$32file = "Win7-KB3134760-x86.msu"

# KB id
$updateid = "KB3134760"

# Set the arch 
$arch = [IntPtr]::Size * 8

# WIndows 10 allready comes with V5
if ($os -ge "10") {
    exit
}

# Get windows version
$os = [Environment]::OSVersion.Version.Major

# Get powershell version
$powershellv = $PSVersionTable.PSVersion.Major

# Framework check
if ($frameworkversion -lt 4 -and $powershellv -lt 5) {
    # Inform user that we are unable to install due to .NET framework date
    Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "$updateid - Unable to install update due to missing dependency (Current - $frameworkversion)"
    # exit
    exit
}

# Run commands depending on arch
if ($arch -eq "64" -and $frameworkversion -ge 4) {
    if ($powershellv -lt 5) {
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "$updateid - x64 not found! Invoking Installer"
    # Set exe file to correct one
    $exefile = $wmfpath+$64file
    # Install Hotfix
    Install-Hotfix ($exefile)
    exit
   }
}
else {
   if ($powershellv -lt 5 -and $frameworkversion -ge 4) {
   Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "$updateid - x32 not found! Invoking Installer"
    # Set exe file to correct one
    $exefile = $wmfpath+$32file
    # Install Hotfix
    Install-Hotfix ($exefile)
    exit
   }
}
