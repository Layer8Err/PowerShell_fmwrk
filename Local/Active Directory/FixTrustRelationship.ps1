#Execute
#############################################################
# Fix trust between client and domain controller
#
# This should be run as administrator from the affected PC
#############################################################
$domainName = "contoso.local"
$adminName = $domainName + "\" + "domainAdmin"
if ($cred) {} else { $cred = Get-Credential $adminName }

$pcname = [string]((Get-WmiObject Win32_ComputerSystem).Name)
Write-Host "Fixing $pcname Trust Relationship with $domainName..." -ForegroundColor Yellow

Reset-ComputerMachinePassword -Server $domainName -Credential $cred
