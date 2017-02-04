@ECHO OFF
TITLE PowerShell Management Framework
COLOR 1F
ECHO Starting PowerShell...
ECHO Checking privilege level...
 
:checkPrivileges 
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges ) 
 
:getPrivileges 
if '%1'=='ELEV' (shift & goto gotPrivileges)  
ECHO Invoking UAC for Privilege Escalation...

setlocal DisableDelayedExpansion
set "batchPath=%~0"
setlocal EnableDelayedExpansion
ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs" 
ECHO UAC.ShellExecute "!batchPath!", "ELEV", "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs"
"%temp%\OEgetPrivileges.vbs" 
exit

:gotPrivileges

:START
setlocal & pushd .
CLS
cd %~dp0
ECHO Starting PowerShell framework...
ECHO Running as administrator...
ECHO Setting things up...
powershell -command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -Confirm:$false"
powershell.exe "%~dp0\Menu.ps1

CLS
ECHO Press any key to exit
PAUSE
