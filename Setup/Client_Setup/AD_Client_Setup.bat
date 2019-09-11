@ECHO OFF
TITLE Configuring Client...
COLOR 0a
CD "%~dp0"
powershell -ExecutionPolicy RemoteSigned -Command .\AD_Client.ps1