# Note: This script requires elevation
$currentRSAT = Get-WindowsCapability -Name RSAT* -Online

$currentRSAT | Where-Object State -eq NotPresent | ForEach-Object {
    $thisRSATpkg = $_.Name
    Write-Host "Installing $thisRSATpkg ..." -ForegroundColor Cyan
    Add-WindowsCapability -Name $thisRSATpkg -Online
}