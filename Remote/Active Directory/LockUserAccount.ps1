#Execute
##################################################################################
# This locks a user's account
##################################################################################
Import-Module ActiveDirectory
$adminName = $env:adminName # Domain admin
if ($cred) {} else { $cred = Get-Credential $adminName }
$UserName = Read-Host 'Username to lock '
$status = net user $UserName /DOMAIN | Find /I "Account Active"

if ([string]$status -like '*No*'){
    Write-Host "  Account is already locked... nothing to do" -ForegroundColor Green
} else {
    Write-Host "  Account is not locked" -ForegroundColor Yellow
    Write-Host "   Locking account.." -ForegroundColor Cyan
    #net user $UserName /DOMAIN /Active:Yes
    Set-ADUser -Credential $cred -Identity $UserName -Enabled $false
    $status = net user $UserName /DOMAIN | Find /I "Account Active"
    Write-Host "    $status" -ForegroundColor White
}
pause