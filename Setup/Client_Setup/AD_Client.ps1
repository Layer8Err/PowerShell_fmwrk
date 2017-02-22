#Execute
###############################################################################################
# Enable WinRM in an Active Directory environment
###############################################################################################
$host.ui.RawUI.WindowTitle = 'Enabling WinRM...'
###############################################################################################

Write-Host "Enabling WinRM..." -ForegroundColor Cyan
winrm enable