#Execute
#######################################
# Reset User Password
#######################################

Function ResetPassword {
    Write-Output "Enter the SAMAccountName of the user who needs their password reset"
    $user = Read-Host 'SAMAccountName'    
    $adminName = $env:adminName
    if ($cred) {} else { $cred = Get-Credential $adminName }
    Import-Module ActiveDirectory
    Set-ADAccountPassword -Credential $cred -Identity $user -Reset -NewPassword (Read-Host -AsSecureString "New Password") # Prompt for new password
    $reset = Read-Host 'Reset at next login? [Y/n]' # Set "User must change password at next logon" to TRUE
    if ($reset.Substring(0,1).ToLower() -ne "n"){
        Set-ADUser -Identity $user -ChangePasswordAtLogon $true -Enabled $true -Credential $cred
    }
}
ResetPassword