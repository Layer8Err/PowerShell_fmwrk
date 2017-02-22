#Execute
###############################################################################################
# Get a list of all Installed Programs
#
# Run through a list of computers and build an array of installed applications
###############################################################################################
$adminName = $env:adminName
$pcList = $env:COMPUTERSLIST ## Path to list of computer-names and user-names
$querySingle = $false  ## Change to false to check installed programs on all PCs
###############################################################################################
if ($cred) {} else { $cred = Get-Credential $adminName }
if ($querySingle) { $pcname = Read-Host 'PC-Name to query for software' }
$ErrorActionPreference = 'SilentlyContinue' ## Suppress errors
$user = @()
$computer = @()
#$speed = Measure-Command { # Need for speed
Import-Csv $pcList | ForEach-Object {
    $user += $_.user
    $computer += $_.computer
}

function InstalledProgramz {
    [CmdletBinding()] Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [String]$Computer,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$User
    )
    $allprogs = @() ## Query registry for uninstall info for installed programs
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
                    'User'=$User;
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

function InstalledPrograms {
    [CmdletBinding()] Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [String]$pcname,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$uname
    )
    $installedBlockFunction = [ScriptBlock]::Create({
        function InstalledProgramz {
            [CmdletBinding()] Param(
                [Parameter(Position = 0, Mandatory = $True)]
                [String]$Computer,
                [Parameter(Position = 1, Mandatory = $True)]
                [String]$User
            )
            $allprogs = @() ## Query registry for uninstall info for installed programs
            if ($Computer -eq $env:COMPUTERNAME) {
                $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { Get-ItemProperty $_.PSPath } | Select-Object DisplayName, Publisher, DisplayVersion, InstallDate
                $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { Get-ItemProperty $_.PSPath } | Select-Object DisplayName, Publisher, DisplayVersion, InstallDate
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
                            'User'=$User;
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
    })
    $installedBlock = [ScriptBlock]::Create($installedBlockFunction.ToString() + "InstalledProgramz -Computer $pcname -User $uname")
    if (!(Test-Connection -ComputerName $pcname -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { 
        Write-Host "......Unreachable" -ForegroundColor Red
    } else {
        if ($pcname -ne $env:COMPUTERNAME){
            Invoke-Command -ComputerName $pcname -Credential $cred -ScriptBlock $installedBlock -AsJob # Get installed software as a job
            Clear-Host
        } else {
            Start-Job -ScriptBlock $installedBlock # Get local installed software as a job
            Clear-Host
        }
    }

}

$userList = Import-CSV $pcList
function userLookup (){
    param(
    [string]$pcname
    )
    $index = [array]::IndexOf($userList.computer, $pcname)
    return $userList.user[$index]
}

if ($querySingle -eq $false){ 
    Write-Host "Building a list of all installed programs on all user PC's..." -ForegroundColor Cyan
    Write-Host "_____________________________________________________________" -ForegroundColor White
}

for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$icomputer = $computer[$i]
    $iuser = $user[$i]
    $currentOpp = "Checking programs installed on " + $icomputer + " " + $iuser
    Write-Output $currentOpp
    InstalledPrograms -pcname $icomputer -uname $iuser
}
## Wait for jobs to finish
$count = 0
While (((Get-Job -State Running).count -gt 0) -and ($count -lt 60)){
    Clear-Host
    $jobs = (Get-Job -State Running).count
    Write-Host -NoNewLine "Waiting for all jobs ($jobs) to complete" -ForegroundColor White
    for ($i = 0 ; $i -lt $count + 1 ; $i ++){
        Write-Host -NoNewline "." -ForegroundColor White
    }
    Start-Sleep -Seconds 1 # Wait for any running jobs to finish
    $count = $count + 1
    if ((Get-Job State Running).count -gt 0){ $count = 60 }
}
Clear-Host
## Import data from remote and local jobs
$allPrograms = @()
foreach ($job in (Get-Job -State Completed)){
    $pcname = $job.Location
    if ($pcname -eq "localhost"){
        $pcname = $env:COMPUTERNAME
    }
    $allPrograms += ((Get-Job -Name $job.Name) | Receive-Job) # Gather jobs
}
Get-Job | Remove-Job

###############################################################################################
# Display search results
###############################################################################################
if ($querySingle){ $allPrograms | Select-Object Computer, User, DisplayName, Version, InstallDate, Architecture | Out-GridView }

function ProgSearch {
    $progsearch = Read-Host 'Program to list (leave blank for all)'
    $allPrograms | Select-Object Computer, User, DisplayName, Version, InstallDate, Architecture | Where-Object -Property DisplayName -match "$progsearch" | Out-GridView
}
Clear-Host
#}

$speed.TotalSeconds
ProgSearch
function searchAgain {
    $searchAgain = Read-Host 'Search Again? [Y/n]'
    if ( ($searchAgain.ToLower() -eq "y") -or ($searchAgain -eq "" )){
        Clear-Host
        ProgSearch
        searchAgain
    }
}
searchAgain