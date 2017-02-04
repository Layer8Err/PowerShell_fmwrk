#Execute
###############################################################################################
# Flushdns
###############################################################################################

Write-Host "Releasing IP..." -ForegroundColor Yellow
cmd /c "ipconfig /release"
Write-Host "Flushing DNS..." -ForegroundColor Yellow
cmd /c "ipconfig /flushdns"
Write-Host "Renewing IP..." -ForegroundColor Yellow
cmd /c "ipconfig /renew"
