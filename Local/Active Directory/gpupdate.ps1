#Execute
###############################################################################################
# Group Policy Update
###############################################################################################

$pcname = [string]((Get-WmiObject Win32_ComputerSystem).Name)
Write-Host "Updating Group Policy for $pcname..." -ForegroundColor Yellow

gpupdate /force

$gpoapplied = gpresult /Scope Computer /v
Write-Output $gpoapplied

PAUSE
