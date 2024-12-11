# Check if the script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    # Re-launch the script with elevated privileges
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Import the Active Directory module
Import-Module ActiveDirectory

# Define the GPO name
$GPOName = "DisableNewOutlook"

# Check if the GPO already exists
$GPO = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue

if ($GPO -eq $null) {
    # Create a new GPO
    New-GPO -Name $GPOName -Comment "GPO to set DoNewOutlookAutoMigration registry key"

    # Get the newly created GPO
    $GPO = Get-GPO -Name $GPOName

 # Define the registry key paths and values
 $RegistryKeyPath1 = "Software\Microsoft\Office\16.0\Outlook\Options\General"
 $RegistryKeyPath2 = "Software\Policies\Microsoft\Office\16.0\Outlook\Options\General"
 $RegistryKeyPath3 = "Software\Policies\Microsoft\Office\16.0\Outlook\Preferences"
 $RegistryValueName1 = "DoNewOutlookAutoMigration"
 $RegistryValueName2 = "NewOutlookMigrationUserSetting"
 $RegistryValueType = "Dword"
 $RegistryValueData = 0

 # Create the registry keys in the GPO
 Set-GPRegistryValue -Name $GPOName -Key "HKCU\$RegistryKeyPath1" -ValueName $RegistryValueName1 -Type $RegistryValueType -Value $RegistryValueData
 Set-GPRegistryValue -Name $GPOName -Key "HKCU\$RegistryKeyPath2" -ValueName $RegistryValueName1 -Type $RegistryValueType -Value $RegistryValueData
 Set-GPRegistryValue -Name $GPOName -Key "HKCU\$RegistryKeyPath3" -ValueName $RegistryValueName2 -Type $RegistryValueType -Value $RegistryValueData

    # Link the GPO to a specific OU (replace 'OU=YourOU,DC=domain,DC=com' with your actual OU path)
    $OUPath = (Get-ADDomain).DistinguishedName
    New-GPLink -Name $GPOName -Target $OUPath

    Write-Output "GPO '$GPOName' created and linked to '$OUPath'."
} else {
    Write-Output "GPO '$GPOName' already exists."
}
