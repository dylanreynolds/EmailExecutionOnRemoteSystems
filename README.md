# EmailExecutionOnRemoteSystems
Uses Outlook rules to run a macro and execute a powershell script.

# Important
This makes your remote execution (RPC) using powershell possible, considered to be very unsecure when enabling on devices.

# Template
See template for further information

Password Setup
$credential = Get-Credential
$credential.Password | ConvertFrom-SecureString | Set-Content c:\scripts\encryptedPassword.txt


# Computer Setup
powershell.exe start-process PowerShell -verb runas
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
Get-NetConnectionProfile
Set-NetConnectionProfile -InterfaceIndex 6 -NetworkCategory Private
Enable-PSRemoting -force
netsh advfirewall firewall set rule group="Windows Management Instrumentation (WMI)" new enable=yes


# Running Setup
$UserName = "admin"
$ComputerName = "[insert computer name]"
$Password = Get-Content c:\scripts\encrypted_password.txt | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PsCredential($UserName, $Password)
Restart-Computer -ComputerName $ComputerName -Credential $Credential -Force

