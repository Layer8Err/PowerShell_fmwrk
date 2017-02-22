@ECHO OFF
TITLE Configuring Client...
color 0a
cd %~dp0
powershell -ExecutionPolicy RemoteSigned -Command .\AD_Client.ps1