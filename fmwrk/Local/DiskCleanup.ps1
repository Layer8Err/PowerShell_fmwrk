#Execute
###############################################################################################
# Run Cleanmgr.exe after setting automated options
#
# RegKeys for options are set in:
# HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*
###############################################################################################

#Set-Location 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders'
#Set-ItemProperty.newproperty '"StateFlags0001"=dword:00000002'

Write-Host "Beginning Disk Cleanup..." -ForegroundColor Yellow

# Set reg keys for StateFlags0001 cleanmgr option set
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\BranchCache' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Downloaded Program Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Offline Pages Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Old ChkDsk Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\RetailDemo Offline Content' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Service Pack Cleanup' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Setup Log Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Upgrade Discarded Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\User file versions' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Defender' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Archive Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Queue Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Archive Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Queue Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Temp Files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows ESD installation files' -Name StateFlags0001 -Value 2
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files' -Name StateFlags0001 -Value 2


## Run Cleanmgr.exe
# Use task scheduler elevation:
schtasks /create /tn "CleanUp" /sc once /st 00:00:00 /sd 01/01/2020 /tr "cleanmgr.exe /sagerun:1 /d C:"
schtasks /run /tn "CleanUp"
schtasks /delete /tn "CleanUp" /f


## TODO:
#  Automate cleanup of:
# "Temporary Windows installation files"
# "Windws Update Cleanup"
#  Clean up System Restore and Shadow Copies

Write-Host "Disk cleanup should be complete" -ForegroundColor Green
Pause
