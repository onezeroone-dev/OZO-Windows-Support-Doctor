#Requires -Modules OZOLogger,PSWindowsUpdate -Version 5.1 -RunAsAdministrator

<#PSScriptInfo
    .VERSION 1.0.0
    .GUID 33232476-5644-44d8-ae63-e3739af537a6
    .AUTHOR Andy Lievertz <alievertz@onezeroone.dev>
    .COMPANYNAME One Zero One
    .COPYRIGHT This script is released under the terms of the GNU General Public License ("GPL") version 2.0.
    .TAGS
    .LICENSEURI https://github.com/onezeroone-dev/OZO-Windows-Support-Doctor/blob/main/LICENSE
    .PROJECTURI https://github.com/onezeroone-dev/OZO-Windows-Support-Doctor
    .ICONURI
    .EXTERNALMODULEDEPENDENCIES
    .REQUIREDSCRIPTS
    .EXTERNALSCRIPTDEPENDENCIES
    .RELEASENOTES https://github.com/onezeroone-dev/OZO-Windows-Support-Doctor/blob/main/CHANGELOG.md
#>

<# 
    .SYNOPSIS
    See description.
    .DESCRIPTION 
    Performs a series of simple Windows maintenance tasks and [optionally] reboots the computer.
    .PARAMETER Reboot
    Reboots the computer after performing the maintenance tasks.
    .LINK
    https://github.com/onezeroone-dev/OZO-Windows-Support-Doctor/blob/main/README.md
    .NOTES
    Run this script in an Administrator PowerShell. When the One Zero One Windows event log provider is available, messages are written to Applications and Services > One Zero One > Operational. Otherwise, messages are written to the Microsoft > Windows > PowerShell > Operational provider with event ID 4100.
#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess = $true)] Param (
    [Parameter(Mandatory=$false,HelpMessage="Invokes a reboot")][Switch]$Reboot
)

# CLASSES
Class OZOWSDMain {
    # PROPERTIES: Strings
    [String] $updateURL = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $ozoLogger = $null
    # METHODS: Constructor method
    OZOWSDMain($Reboot) {
        # Set properties
        $this.updateURL = "catalog.update.microsoft.com"
        # Create a logger object
        $this.ozoLogger = (New-OZOLogger)
        # Log a process start message
        $this.ozoLogger.Write("Starting OZO Windows Support Doctor.","Information")
        # Determine if the configuation is valid
        If ($this.ValidateConfiguration() -And $this.ValidateEnvironment()) {
            # Call the support methods
            $this.SFCScanNow()
            $this.DismRestoreHealth()
            $this.FlushDNS()
            $this.CycleNICs()
            $this.RefreshGroupPolicy()
            $this.WindowsUpdates()
        }
        # Log a process end message
        $this.ozoLogger.Write("Finished OZO Windows Support Doctor.","Information")
        # Reboot
        $this.RebootComputer($Reboot)
    }
    # JSON validation method
    Hidden [Boolean] ValidateConfiguration() {
        # Control variable
        [Boolean] $Return = $true
        # Return
        return $Return
    }
    # Environment validation method
    Hidden [Boolean] ValidateEnvironment() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if the session is not user-interactive
        If ([Environment]::UserInteractive -eq $false) {
            # Session is not user-interactive
            $this.ozoLogger.Write("Please run this script in an interactive session.","Error")
            $Return = $false
        }
        # Determine if the OS is 64-bit but PowerShell is 32-bit
        If ([Environment]::Is64BitOperatingSystem -eq $true -And ([Environment]::Is64BitProcess) -eq $false) {
            # OS is 64-bit but PowerShell is 32-bit; restart as 64-bit process
            $this.ozoLogger.Write("OS is 64-bit but PowerShell is 32-bit; restarting.","Warning")
            Start-Process -Wait -NoNewWindow -FilePath (Join-Path $Env:SystemRoot -ChildPath "sysnative\WindowsPowerShell\v1.0\powershell.exe") -ArgumentList ('-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -NoProfile -File "' + $MyInvocation.MyCommand.Definition + '"')
            $Return = $false
        }
        # Return
        return $Return
    }
    # SFC Scannow repair method
    Hidden [Void] SFCScanNow() {
        # Log
        $this.ozoLogger.Write("Attempting to repair Windows (SFC).","Information")
        # Try to repair Windows (SFC)
        Try {
            Start-Process -FilePath (Join-Path -Path $Env:SystemRoot -ChildPath "System32\sfc.exe") -ArgumentList "/scannow" -Wait -NoNewWindow -ErrorAction Stop
            # Success
            $this.ozoLogger.Write("Repaired Windows (SFC).","Information")
        } Catch {
            # Failure
            $this.ozoLogger.Write("Unable to repair Windows (SFC).","Warning")
        }
    }
    # DISM repair method
    Hidden [Void] DismRestoreHealth() {
        # Log
        $this.ozoLogger.Write("Attempting to repair Windows (DISM).","Information")
        # Try to repair Windows (DISM)
        Try {
            Repair-WindowsImage -Online -RestoreHealth -ErrorAction Stop
            # Success
            $this.ozoLogger.Write("Repaired Windows (DISM).","Information")
        } Catch {
            # Failure
            $this.ozoLogger.Write("Unable to repair Windows (DISM).","Warning")
        }
    }
    # Flush DNS cache method
    Hidden [Void] FlushDNS() {
        # Log
        $this.ozoLogger.Write("Flushing DNS cache.","Information")
        # Try to flush the DNS cache
        Try {
            Clear-DnsClientCache -ErrorAction Stop
            # Success
            $this.ozoLogger.Write("Flushed DNS cache.","Information")
        } Catch {
            # Failure
            $this.ozoLogger.Write("Unable to flush the DNS cache.","Warning")
        }
    }
    # Disable/Enable NICs method
    Hidden [Void] CycleNICs() {
        # Log
        $this.ozoLogger.Write("Managing all active network controllers (you may briefly lose connectivity).","Information")
        # Iterate through a list of active network adapters
        ForEach ($netAdapter in (Get-NetAdapter | Where-Object {$_.Status -eq "Up"})) {
            # Get the WMI object
            [System.Management.ManagementBaseObject] $wmiInterface = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.Index -eq $netAdapter.ifIndex})
            # Determine if DHCP is enabled
            If ($wmiInterface.DHCPEnabled -eq $true) {
                # Interface has DHCP enabled; renew the address
                $this.ozoLogger.Write(("Interface " + $netAdapter.Name + " is configured by DHCP. Attempting to renew the IP address."),"Information")
                # Renew the address (forces release)
                $wmiInterface.RenewDHCPLease()
            }
            # Try to disable the adapter
            Try {
                Disable-NetAdapter -Confirm:$false -Name $netAdapter.Name -ErrorAction Stop
                # Success; try to enable
                $this.ozoLogger.Write(("Disabled " + $netAdapter.Name + "; attempting to re-enable."),"Information")
                Try {
                    Enable-NetAdapter -Confirm:$false -Name $netAdapter.Name  -ErrorAction Stop
                    # Success
                    $this.ozoLogger.Write(("Re-enabled " + $netAdapter.name + "."),"Information")
                } Catch {
                    # Failure
                    $this.ozoLogger.Write(("Failed to re-enable " + $netAdapter.name + "."),"Warning")
                }
            } Catch {
                # Failure
                $this.ozoLogger.Write(("Failed to disable " + $netAdapter.Name + ". Error message is: " + $_),"Warning")
            }
        }
    }
    # Refresh group policy method
    Hidden [Void] RefreshGroupPolicy() {
        # Determine if system is part of a domain
        If ((Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain -eq $true) {
            # System is part of a domain; log
            $this.ozoLogger.Write("Computer is part of a domain. Refreshing computer policy.","Information")
            # Try to refresh group policy
            Try {
                Start-Process -FilePath (Join-Path -Path $Env:SystemRoot -ChildPath "System32\gpupdate.exe") -ArgumentList "/force" -Wait -NoNewWindow -ErrorAction Stop
                # Success
                $this.ozoLogger.Write("Refreshed computer policy.","Information")
            } Catch {
                # Failure
                $this.ozoLogger.Write("Unable to refresh computer policy.","Warning")
            }
        } Else {
            # System is not part of a domain; log
            $this.ozoLogger.Write("Computer is not part of a domain; skipping group policy update.","Information")
        }
    }
    # Windows updates method
    Hidden [Void] WindowsUpdates() {
        # Determine if the Windows updates catalog is available by HTTP
        If ([Boolean](Test-NetConnection -ComputerName $this.updateURL -CommonTCPPort HTTP -ErrorAction SilentlyContinue) -eq $true) {
            # Catalog is available by http; try to get windows updates
            Try {
                # Determine if there are any pending updates
                If ([Boolean](Get-WindowsUpdate -ErrorAction Stop) -eq $true) {
                    # Found pending updates; try to install updates
                    Try {
                        Install-WindowsUpdate -AcceptAll -IgnoreReboot -ErrorAction Stop
                        # Success
                        $this.ozoLogger.Write("Installed Windows updates","Information")
                    } Catch {
                        # Failure
                        $this.ozoLogger.Write("Unable to install Windows updates","Warning")
                    }
                } Else {
                    # No pending updates
                    $this.ozoLogger.Write("Found no pending Windows updates.","Information")
                }
            } Catch {
                $this.ozoLogger.Write("Unable to query for Windows updates.","Warning")
            }
        } Else {
            $this.ozoLogger.Write("Unable to reach the Windows Updates catalog.","Warning")
        }
    }
    # Rebooot method
    Hidden [Void] RebootComputer($Reboot) {
        # Determine if operator requested reboot
        If ($Reboot -eq $true) {
            # Operator requested reboot
            $this.ozoLogger.Write("Rebooting.","Warning")
            Restart-Computer -Force
        } Else {
            # Operator did not request reboot
            $this.ozoLogger.Write("Operator did not provide the Reboot parameter. Recommend rebooting at the next opportunity.","Information")
        }
    }
}

# Create a Main object
[OZOWSDMain]::new($Reboot) | Out-Null
