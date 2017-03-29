# Alder Grange School - Deployment System
# Script Version - V2
# Date Created - 16/01/2017
# Primary Task - Checks flash version agaist the EXE file and installs new version if out-of-date

# Setup installer function
function Install-Exe ($EXE)
{
    # Run the EXE
    $installer = Start-Process -FilePath $EXE -ArgumentList " /quiet /norestart" -Wait -Passthru

    if ($installer.ExitCode -eq 0) {
        Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Win32Help - The installer returned - The operation completed successfully."
        
        # Get update status
        $updateinstalled = Get-HotFix -id KB917607 -ErrorAction SilentlyContinue

        # If we have found it but the registry isnt up-to-date with it then lets stick it into the registry
            if (!$updateinstalled) {
                # Define registry
                $winhelpinstalledreg = HKLM:\SOFTWARE\AGSDS\WinHelp32
                # Lets add the registry hive
                New-Item -Path $winhelpinstalledreg
                # Add the item
                Get-Item -Path $winhelpinstalledreg | New-ItemProperty -Name Installed -Value 1
        }
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

# If os isnt Windows 7 then exit
if ($os -ne 7) {
    exit
}

# Get current registry install status
$installstatushive = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS\WinHelp32 -ErrorAction SilentlyContinue

$installstatus = $installstatushive.Installed

# Queryig the above costs process time, in turn slows the login. Once installed we will use the registry to check if it is installed.
if (!$installstatus) {
    # Allready Installed so exit
    exit
}
else {
    # Get update status
    $updateinstalled = Get-HotFix -id KB917607 -ErrorAction SilentlyContinue

    # If we have found it but the registry isnt up-to-date with it then lets stick it into the registry
    if (!$updateinstalled) {
        # Define registry
        $winhelpinstalledreg = HKLM:\SOFTWARE\AGSDS\WinHelp32
        # Lets add the registry hive
        New-Item -Path $winhelpinstalledreg
        # Add the item
        Get-Item -Path $winhelpinstalledreg | New-ItemProperty -Name Installed -Value 1
    }
}

# Setup Auth
# Get Auth hive
$authreg = Get-ItemProperty -Path HKLM:\SOFTWARE\AGSDS -ErrorAction SilentlyContinue

# Get Office Key
$Win32helpreg = $authreg.Win32help

# Find
if ($Win32helpreg) {
   # No Auth
   Write-EventLog -LogName Application -Source AGSDS -EntryType Error -EventId 3001 -Message "Win32Help - No registry key auth exists"
   exit
}

# Installed if not installed
if (!$Currentversion -and $Win32helpreg -eq 1) {
    # Log
    Write-EventLog -LogName Application -Source AGSDS -EntryType Information -EventId 3001 -Message "Win32Help - Currently not installed"
    # Run installer
    Install-Exe ($EXE)
    exit
}

# SIG # Begin signature block
# MIIH8AYJKoZIhvcNAQcCoIIH4TCCB90CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJDb4LX16b3xeC7VSvm7RoYKn
# Q3KgggWoMIIFpDCCBIygAwIBAgITSQAAAETbaFqwoxWV4wAAAAAARDANBgkqhkiG
# 9w0BAQUFADB6MRIwEAYKCZImiZPyLGQBGRYCdWsxEjAQBgoJkiaJk/IsZAEZFgJh
# YzEZMBcGCgmSJomT8ixkARkWCWxhbmNzbmdmbDEVMBMGCgmSJomT8ixkARkWBWFs
# ZGVyMR4wHAYDVQQDExVhbGRlci1BTERFUi1EQzAzLUNBLTEwHhcNMTcwMTE2MTQw
# MzQ0WhcNMTgwMTE2MTQwMzQ0WjCBmTESMBAGCgmSJomT8ixkARkWAnVrMRIwEAYK
# CZImiZPyLGQBGRYCYWMxGTAXBgoJkiaJk/IsZAEZFglsYW5jc25nZmwxFTATBgoJ
# kiaJk/IsZAEZFgVhbGRlcjEOMAwGA1UECxMFU3RhZmYxFDASBgNVBAsTC1RlY2hu
# aWNpYW5zMRcwFQYDVQQDEw5TdHVhcnQgUGVhcnNvbjCBnzANBgkqhkiG9w0BAQEF
# AAOBjQAwgYkCgYEA0/M6jovLvqZKiZ6jC/cHC6IdA3M4b4qQasdsJ+i2f78XB54e
# 5VSN1voZtZPqeEW/S/x71BS7Q/9E/2OycBFt7/A4UPz5725aUSykEKdRMg7L8g02
# XacUoF3SD2PrI2fvx9Lzb3Cwmc7STxMwgbo4dUb4ptRpGtqYMiUS3I476GECAwEA
# AaOCAoUwggKBMCUGCSsGAQQBgjcUAgQYHhYAQwBvAGQAZQBTAGkAZwBuAGkAbgBn
# MBMGA1UdJQQMMAoGCCsGAQUFBwMDMAsGA1UdDwQEAwIHgDAdBgNVHQ4EFgQU2Fol
# 7fgl5hRyxJyRgscNpKHF9ZcwHwYDVR0jBBgwFoAU4jjx2XFhMzVcSBb/FONywC/Y
# pGswgeYGA1UdHwSB3jCB2zCB2KCB1aCB0oaBz2xkYXA6Ly8vQ049YWxkZXItQUxE
# RVItREMwMy1DQS0xLENOPWFsZGVyLWRjMDMsQ049Q0RQLENOPVB1YmxpYyUyMEtl
# eSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9YWxk
# ZXIsREM9bGFuY3NuZ2ZsLERDPWFjLERDPXVrP2NlcnRpZmljYXRlUmV2b2NhdGlv
# bkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludDCB1wYI
# KwYBBQUHAQEEgcowgccwgcQGCCsGAQUFBzAChoG3bGRhcDovLy9DTj1hbGRlci1B
# TERFUi1EQzAzLUNBLTEsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2Vz
# LENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9YWxkZXIsREM9bGFuY3Nu
# Z2ZsLERDPWFjLERDPXVrP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1j
# ZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MDMGA1UdEQQsMCqgKAYKKwYBBAGCNxQCA6Aa
# DBhzcGVhcnNvbkBhbGRlcmdyYW5nZS5jb20wDQYJKoZIhvcNAQEFBQADggEBAEek
# Y77Cj78VdgsfWfDLCYknk9FNJwAbVsFt34up6AEowp9jsVmN7NgmxCx+FVWGDF7a
# OKDPcZd6XDdCjHY/gFxlN3Oz2NXcIfhuUs1CF82CfOa4kIKWATHd5FYe59G95Zr8
# Mma/SdeUxVOqhbmXMBCfifAhEXLWsh5jkxjJkpJUjnVCqsSQq9Sgw7Rokke6+YvU
# nJXZvChl1DZzzJ6g6Z/mOBOqTaAe5LFS7AyhfU8uqa2V2wyrc7fr5/mkBHC0u5oh
# NWaF/biTB3AXA+SXxe0cqfhMnZzAuZ32Onl9Ddp95Wzz5ZZrwSi+ct+ro6vL28IS
# /ZQG3BXYGFSQnaUNVZ4xggGyMIIBrgIBATCBkTB6MRIwEAYKCZImiZPyLGQBGRYC
# dWsxEjAQBgoJkiaJk/IsZAEZFgJhYzEZMBcGCgmSJomT8ixkARkWCWxhbmNzbmdm
# bDEVMBMGCgmSJomT8ixkARkWBWFsZGVyMR4wHAYDVQQDExVhbGRlci1BTERFUi1E
# QzAzLUNBLTECE0kAAABE22hasKMVleMAAAAAAEQwCQYFKw4DAhoFAKB4MBgGCisG
# AQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFFC0
# +AOJglVARD1k83UlplANp9w4MA0GCSqGSIb3DQEBAQUABIGAshuUYam2N5v54Qpi
# JqWojx6Bjl7AvZ0os3WWYHvuexAUvM/wiDHx28+R4dRkP2MCWLqUzkr3lUeEsFWL
# 6XzH5waagkxuKkLsf88xrZJePoO5AtDHqf/sXqYlyb0be1pF0+7tmTHde2ErQ9En
# IPXa0Edxm1CMptzeN/NmWMMtSqk=
# SIG # End signature block
