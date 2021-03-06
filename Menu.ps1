﻿###############################################################################################
#  __  __      _          __  __                  
# |  \/  |__ _(_)_ __    |  \/  |___ _ __  _   _ 
# | |\/| / _` | | '_ \   | |\/| / _ \ '_ \| | | |
# | |  || (_| | | | | |  | |  ||  __/ | | | |_| |
# |_|  |_\__,_|_|_| |_|  |_|  |_\___|_| |_|\__,_|                                                  
#
# Setup Envoronment Variables and check for pre-requisites
# This script has been tested with PowerShell v3.0 and up
###############################################################################################
# Additional scripts should begin with the following lines in order to be read by the menu:
# #ExecuteOut      # Run the selected script
# #ExecuteOutOpen  # or anything else simply opens the script in ISE
###############################################################################################
# Setup stage of the menu script. Modify to set global variables if necessary
###############################################################################################
# Environment variables

Write-Host "Setting environment variable for PSFmwrkRoot..." -ForegroundColor Yellow
Set-Location $PSScriptRoot
$basePath = $PWD.Path
[Environment]::SetEnvironmentVariable("PSFmwrkRoot", $basePath) # environment variable for root directory
Write-Host '$env:PSFmwrkRoot = ', $env:PSFmwrkRoot
Set-Location $env:PSFmwrkRoot

Write-Host "Reading environment_settings.xml from Config folder..." -ForegroundColor Yellow
$settingsXMLFile = $env:PSFmwrkRoot + '\Setup\Config\environment_settings.xml'
if (!(Test-Path $settingsXMLFile)){
    & ($env:PSFmwrkRoot + '\Setup\Config\Configure.ps1') # Launch xml Configuration script for initial setup
    Set-Location $env:PSFmwrkRoot
}
$xml = [xml](Get-Content $settingsXMLFile)
$domainName = $xml.environment.domain
[Environment]::SetEnvironmentVariable("domainName", $domainName) # environment variable for domain name
Write-Host '$env:domainName = ', $env:domainName -ForegroundColor Cyan

$domainAdmin = $xml.environment.domainadmin
$adminName = $env:domainName + "\" + $domainAdmin
Write-Host "Setting environment variable for domain admin..." -ForegroundColor Yellow
[Environment]::SetEnvironmentVariable("adminName", $adminName) # environment variable for domain admin

# Computer/Servers lists
$ListOfComputers = $basePath + "\CSVLists\userPCs.csv"
Write-Host "Setting environment variable for list of computers:" -ForegroundColor Yellow
[Environment]::SetEnvironmentVariable("COMPUTERSLIST", $ListOfComputers)
Write-Host '$env:COMPUTERSLIST = ', $env:COMPUTERSLIST -ForegroundColor Cyan

$ListOfServers = $basePath + "\CSVLists\servers.csv"
Write-Host "Setting environment variable for list of servers:" -ForegroundColor Yellow
[Environment]::SetEnvironmentVariable("SERVERSLIST", $ListOfServers)
Write-Host '$env:SERVERSLIST = ', $env:SERVERSLIST -ForegroundColor Cyan

if (!(Test-Path $ListOfServers) -or !(Test-Path $ListOfComputers)){
    # Automatically create dummy CSV files
    & ($env:PSFmwrkRoot + '\CSVLists\CreateEmptyLists.ps1')
}

###############################################################################################

function menu {
    Get-PSSession | Remove-PSSession
    $host.ui.RawUI.WindowTitle = "PS Management Framework"
    [console]::ResetColor()
    #$psISE.Options.RestoreDefaultTokenColors()
    Where-Object (Test-Path variable:global:menu ) { Remove-Variable -Name menu -Force }
    Where-Object (Test-Path variable:global:dirMenu ) {Remove-Variable -Name dirMenu -Force }
    Where-Object (Test-Path variable:global:cwdStruct ) {Remove-Variable -Name cwdStruct -Force }
    Where-Object (Test-Path variable:global:chosen ) {Remove-Variable -Name chosen -Force }
    Where-Object (Test-Path variable:global:color ) {Remove-Variable -Name color -Force }
    Where-Object (Test-Path variable:global:cwd ) {Remove-Variable -Name cwd -Force }
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble (thanks Write-Host)
    $index = 1
    $menu= @()
    $cwd = (Get-Item -Path ".\" -Verbose).FullName
    $cwd | Get-ChildItem -Attributes Directory | ForEach-Object {
        if($_.Extension -eq ".ps1"){
            $color = "Cyan"
        } else {
            $color = "Green"
        }
        $properties = @{
            'Selection'=$index;
            'Path'=$_.FullName;
            'Name'=$_.Name;
            'Color'=$color
        }
        $dirMenu = New-Object PSObject -Property $properties
        $menu += $dirMenu
        ++$index
    }
    $cwd | Get-ChildItem | Where-Object Extension -eq ".ps1" | ForEach-Object {
        if($_.Extension -eq ".ps1"){
            $color = "Cyan"
        } else {
            $color = "Green"
        }
        $properties = @{
            'Selection'=$index;
            'Path'=$_.FullName;
            'Name'=$_.Name;
            'Color'=$color
        }
        $dirMenu = New-Object PSObject -Property $properties
        $menu += $dirMenu
        ++$index
    }
    if ($cwd -eq $env:PSFmwrkRoot){
        $properties = @{
            'Selection'=$index;
            'Path'="..";
            'Name'="Exit";
            'Color'="Red"
        }
    } else { 
        $properties = @{
            'Selection'=$index;
            'Path'="..";
            'Name'="Back";
            'Color'="Yellow"
        }
    }
    $dirMenu = New-Object PSObject -Property $properties
    $menu += $dirMenu
    Clear-Host
    $rootName = (Get-Item -Path ".\" -Verbose).Name
    Write-Host "===========$rootName==========" -ForegroundColor White
    $menu | ForEach-Object {
        [String]$color = $_.Color
        $selection = $_.Selection
        $name = $_.Name
        Write-Host "$selection`t$name" -ForegroundColor $color
    }
    Write-Host "`n"
    [INT]$choice = Read-Host ">"
    try { 
        $chosen = $menu[$choice - 1]
    } 
    catch { 
        Write-Host "Selection error" -ForegroundColor Red 
        pause
    }
    if (($chosen.Color -eq "Green") -or ($chosen.Color -eq "Yellow")){
        Set-Location $chosen.Path
        menu
    }
    if ($chosen.Color -eq "Cyan") {
        [String]$script = $chosen.Path
        $exeLvl = Get-Content $script -First 1 # get first line of PowerShell script
        if ($exeLvl -eq "#Execute"){ # run the PowerShell script and return to the menu
            $execute = $true
        } else { $execute = $false }
        Clear-Host
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Write-Host $chosen.Name -ForegroundColor White
        Write-Host "`n"
        if ($execute) {
            & $script
            menu
        } else {
            if ($exeLvl -eq "#ExecuteOut") { # run the script but don't return to the menu
                & $script
            } elseif ($exeLvl -eq "#ExecuteOutOpen") { # open the script in ISE and run the script
                $psISE.CurrentPowerShellTab.Files.Add($script)
                & $script
            } else { # open the script in ISE
                $psISE.CurrentPowerShellTab.Files.Add($script)
            }
        }
    }
    if ($chosen.Color -eq "Red"){
        exit
    }
}

# Get the creds for other scripts to use
if ($cred) {} else { $cred = Get-Credential $adminName }

menu # Launch the menu