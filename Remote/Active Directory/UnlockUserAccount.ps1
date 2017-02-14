#Execute
##################################################################################
# This unlocks a user's account
##################################################################################
Import-Module ActiveDirectory
$adminName = $env:adminName # Domain admin
if ($cred) {} else { $cred = Get-Credential $adminName }
## Search for locked accounts:
Search-ADAccount -LockedOut
## Auto Unlock:
# Search-ADAccount -LockedOut | Unlock-ADAccount -Credential $cred
$UserName = Read-Host 'Username to unlock '
$status = net user $UserName /DOMAIN | Find /I "Account Active"

if ([string]$status -like '*Yes*'){
    Write-Host "  Account is not locked... nothing to do" -ForegroundColor Green
} else {
    Write-Host "  Account is locked" -ForegroundColor Yellow
    Write-Host "   Enabling account.." -ForegroundColor Cyan
    #net user $UserName /DOMAIN /Active:Yes
    Set-ADUser -Credential $cred -Identity $UserName -Enabled $true
    # Get-ADUser -Identity $UserName | Unlock-ADAccount -Credential $cred
    $status = net user $UserName /DOMAIN | Find /I "Account Active"
    Write-Host "    $status" -ForegroundColor White
}
pause