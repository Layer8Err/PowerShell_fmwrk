#Execute
###############################################################################################
# Query All Installed Programs
# 
# Return a list of installed programs as listed in the registry
###############################################################################################

$pcname = [string]((Get-WmiObject Win32_ComputerSystem).Name)

function InstalledPrograms {
    [CmdletBinding()] Param(
    [Parameter(Position = 0, Mandatory = $True)]
    [String]$Computer
    )

    $allprogs = @()
    ## Query registry for uninstall info for installed programs
    if ($Computer -eq $env:COMPUTERNAME) {
        $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { Get-ItemProperty $_.PSPath } | Select-Object DisplayName, Publisher, DisplayVersion, InstallDate
        $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { Get-ItemProperty $_.PSPath } | Select-Object DisplayName, Publisher, DisplayVersion, InstallDate
    } else  {
        $uninstall32 = Invoke-Command -ComputerName $Computer -Credential $cred -ScriptBlock { Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { Get-ItemProperty $_.PSPath } | Select-Object DisplayName, Publisher, DisplayVersion, InstallDate }
        $uninstall64 = Invoke-Command -ComputerName $Computer -Credential $cred -ScriptBlock { Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { Get-ItemProperty $_.PSPath } | Select-Object DisplayName, Publisher, DisplayVersion, InstallDate }
    }
    function progstats {
        [CmdletBinding()] Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [PSObject]$uninstallData,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$architecture
        )
        $uninstallData | ForEach-Object {
            if ($installDate) { Clear-Variable -Name installDate }
            if (($_.InstallDate).Length -gt 1){
                [String]$installDate = ($_.InstallDate).Replace(' ','')
            } else { $installDate = "" }
            if ($installDate.Length -eq 8) {
                $installDate = $installDate.Substring(0,4) + "/" + $installDate.Substring(4,2) + "/" + $installDate.Substring(6,2)
            } else {
                $installDate = ""
            }
            $properties = @{
                    'Computer'=$Computer;
                    'DisplayName'=$_.DisplayName;
                    'Publisher'=$_.Publisher;
                    'Version'=$_.DisplayVersion;
                    'InstallDate'=$installDate;
                    'Architecture'=$architecture}
            if (($_.DisplayName).Length -gt 1) {
                $progstats = New-Object PSObject -Property $properties
                $allprogs += $progstats
            }
        }
        Return $allprogs
    }

    if ($uninstall64) { $allProgs = progstats -uninstallData $uninstall64 -architecture "x64" }
    if ($uninstall32) { $allProgs += progstats -uninstallData $uninstall32 -architecture "x32" }

    Return $allprogs
}


$allPrograms = @()


Write-Host "Checking installed programs on $pcname..." -ForegroundColor Yellow
$allPrograms = InstalledPrograms -Computer $pcname

$allPrograms | Select-Object Computer, DisplayName, Version, InstallDate, Architecture | Out-GridView
