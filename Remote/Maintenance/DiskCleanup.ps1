#Execute
###############################################################################################
# Run Disk Cleanup
#
# Runs Cleanmgr.exe after setting automated options
# RegKeys for options are set in:
# HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*
###############################################################################################
$adminName = $env:adminName
###############################################################################################
$pcname = Read-Host 'PC-Name to run Disk Cleanup on'

#Set Registry Keys for Disk Cleanup Options
$scriptBlock = [ScriptBlock]::Create({
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\BranchCache' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Content Indexer Cleaner' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Downloaded Program Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files' -Name StateFlags0001 -Value 2
    If ( -Not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files')){
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files'
    }
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Offline Pages Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Old ChkDsk Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin' -Name StateFlags0001 -Value 2
    If ( -Not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Remote Desktop Cache Files')){
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Remote Desktop Cache Files'
    }
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Remote Desktop Cache Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\RetailDemo Offline Content' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Service Pack Cleanup' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Setup Log Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files' -Name StateFlags0001 -Value 2
    If ( -Not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Offline Files')){
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Offline Files'
    }
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Offline Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files' -Name StateFlags0001 -Value 2
    If ( -Not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Sync Files')){
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Sync Files'
    }
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Sync Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Upgrade Discarded Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\User file versions' -Name StateFlags0001 -Value 2
    If ( -Not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\WebClient and WebPublisher Cache')){
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\WebClient and WebPublisher Cache'
    }
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\WebClient and WebPublisher Cache' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Defender' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Archive Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Queue Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Archive Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Queue Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Temp Files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows ESD installation files' -Name StateFlags0001 -Value 2
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files' -Name StateFlags0001 -Value 2
})

#Schedule task to run Disk Cleanup
$scriptBlock2 = [ScriptBlock]::Create({
    schtasks /create /tn "CleanUp" /sc once /st 00:00:00 /sd 01/01/2020 /tr "cleanmgr.exe /sagerun:2"
    schtasks /run /tn "CleanUp"
    schtasks /delete /tn "CleanUp" /f
})

if ((Get-WmiObject Win32_ComputerSystem).Name -ne $pcname){
    if ($cred) {} else { $cred = Get-Credential $adminName }
    ## Execute the scriptblock on the remote PC
    Write-Host "Running disk cleanup on $pcname Please wait..." -ForegroundColor Yellow
    Invoke-Command -ComputerName $pcname -Credential $cred -ScriptBlock $ScriptBlock # Set reg keys for sagerun 2
    Invoke-Command -ComputerName $pcname -Credential $cred -ScriptBlock $ScriptBlock2 # Execute sagerun 2
} else {
    Write-Host "Running disk cleanup on your machine Please wait..." -ForegroundColor Yellow
    Invoke-Command -ScriptBlock $ScriptBlock # Set reg keys for sagerun 2
    Invoke-Command -ScriptBlock $ScriptBlock2 # Execute sagerun 2e
}

## TODO:
#  Automate cleanup of:
# "Temporary Windows installation files"
# "Windws Update Cleanup"
#  Clean up System Restore and Shadow Copies