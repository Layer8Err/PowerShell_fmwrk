#requires -RunAsAdministrator

Function Install-ADModule {
    [CmdletBinding()]
    Param(
        [switch]$Test = $false
    )

    If ((Get-CimInstance Win32_OperatingSystem).Caption -like "*Windows 10*") {
        Write-Verbose '---This system is running Windows 10'
    } Else {
        Write-Warning '---This system is not running Windows 10'
        break
    }

    If (Get-HotFix -Id KB2693643 -ErrorAction SilentlyContinue) {

        Write-Verbose '---RSAT for Windows 10 is already installed'

    } Else {

        Write-Verbose '---Downloading RSAT for Windows 10'

        If ((Get-CimInstance Win32_ComputerSystem).SystemType -like "x64*") {
            $dl = 'WindowsTH-KB2693643-x64.msu'
        } Else {
            $dl = 'WindowsTH-KB2693643-x86.msu'
        }
        Write-Verbose "---Hotfix file is $dl"

        Write-Verbose "---$(Get-Date)"
        #Download file sample
        #https://gallery.technet.microsoft.com/scriptcenter/files-from-websites-4a181ff3
        $BaseURL = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/'
        $URL = $BaseURL + $dl
        $Destination = Join-Path -Path $HOME -ChildPath "Downloads\$dl"
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($URL,$Destination)
        $WebClient.Dispose()

        Write-Verbose '---Installing RSAT for Windows 10'
        Write-Verbose "---$(Get-Date)"
        # http://stackoverflow.com/questions/21112244/apply-service-packs-msu-file-update-using-powershell-scripts-on-local-server
        wusa.exe $Destination /quiet /norestart /log:$home\Documents\RSAT.log

        # wusa.exe returns immediately. Loop until install complete.
        do {
            Write-Host "." -NoNewline
            Start-Sleep -Seconds 3
        } until (Get-HotFix -Id KB2693643 -ErrorAction SilentlyContinue)
        Write-Host "."
        Write-Verbose "---$(Get-Date)"
    }

    # The latest versions of the RSAT automatically enable all RSAT features
    If ((Get-WindowsOptionalFeature -Online -FeatureName `
        RSATClient-Roles-AD-Powershell -ErrorAction SilentlyContinue).State `
        -eq 'Enabled') {

        Write-Verbose '---RSAT AD PowerShell already enabled'

    } Else {

        Write-Verbose '---Enabling RSAT AD PowerShell'
        Enable-WindowsOptionalFeature -Online -FeatureName RSATClient-Roles-AD-Powershell

    }

    Write-Verbose '---Downloading help for AD PowerShell'
    Update-Help -Module ActiveDirectory -Verbose -Force

    Write-Verbose '---ActiveDirectory PowerShell module install complete.'

    # Verify
    If ($Test) {
        Write-Verbose '---Validating AD PowerShell install'
        dir (Join-Path -Path $HOME -ChildPath Downloads\*msu)
        Get-HotFix -Id KB2693643
        Get-Help Get-ADDomain
        Get-ADDomain
    }
}

###Get-Help Install-ADModule -Full

Write-Host "Installing ADModules..." -ForegroundColor Green
Install-ADModule -Verbose

###Install-ADModule -Test -Verbose

#break

<#
# Remove
wusa.exe /uninstall /kb:2693643 /quiet /norestart /log:$home\RSAT.log
#>


Function InstallMSOnline {
    $basePath = $PSScriptRoot
    $client = New-Object System.Net.WebClient
    $installerPath = $basePath + "\msoidcli_64.msi"
    #echo $installerPath
    Write-Host "Downloading Microsoft Online Sign-In Assistant for IT Professionals RTW..." -ForegroundColor Yellow
    $client.DownloadFile("https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi", $installerPath)
    Write-Host "Beginning install..." -ForegroundColor Yellow
    msiexec.exe /package $installerPath /qn
    
    $installerPath = $basePath + "\AdministrationConfig-en.msi"
    Write-Host "Downloading Azure Active Directory Module for Windows PowerShell..." -ForegroundColor Yellow
    $client.DownloadFile("https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi", $installerPath)
    Write-Host "Beginning install..." -ForegroundColor Yellow
    msiexec.exe /package $installerPath /qn ACCEPT-EULA=YES
}

Write-Host "Installing MSOnlineModules..." -ForegroundColor Green
InstallMSOnline