#Execute
###############################################################################################
# Enable WinRM in an Active Directory environment
###############################################################################################
$host.ui.RawUI.WindowTitle = 'Enabling WinRM...'
###############################################################################################

Write-Host "Enabling WinRM..." -ForegroundColor Cyan
Enable-PSRemoting -Force ; Write-Output "y" | winrm qc
