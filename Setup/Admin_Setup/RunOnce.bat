@ECHO OFF
:: Run as Admin ::
TITLE Setting Powershell Execution Policy
ECHO Setting Powershell Execution Policy to "RemoteSigned"
powershell -command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -Confirm:$false"
CD "%~dp0"
TITLE Installing PowerShell Modules
ECHO Installing Binary PowerShell Modules
powershell -command .\InstallModules.ps1
ECHO Installing Azure Module
powershell -command "Install-Module -Name NuGet -Force -Confirm:$false; Install-Module -Name Azure -Force -Confirm:$false"
ECHO Setup Complete. Enjoy PowerShell :)
PAUSE