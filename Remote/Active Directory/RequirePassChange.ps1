#Execute
#######################################
# Require user to change password at next login
#######################################

Function chgPassword {
    Write-Output "Enter the SAMAccountName of the user who needs to change their password at the next logon"
    $adminName = $env:adminName
    if ($cred) {} else { $cred = Get-Credential $adminName }
    $user = Read-Host 'SAMAccountName'
    Import-Module ActiveDirectory
    Set-ADUser -Identity $user -ChangePasswordAtLogon $true -Enabled $true -Credential $cred # This sets "User must change password at next logon" to TRUE
}
chgPassword