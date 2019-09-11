#Execute
###############################################################################################
# Query Computer Specs
# Modified to utilize jobs (parallel data gathering)
###############################################################################################
# Path to list of computer-names and user-names
$pcList = $env:COMPUTERSLIST
$adminName = $env:adminName
if ($cred) {} else { $cred = Get-Credential $adminName }
#######################################*Begin Script*##########################################
$host.ui.RawUI.WindowTitle = "Getting PC info..."
$user = @()
$computer = @()

#import list of usernames and computernames from .csv
#save into two arrays
Import-Csv $pcList | ForEach-Object {
        $user += $_.user
        $computer += $_.computer
}

function HWcomputerInfo {
    [CmdletBinding()] Param(
    [Parameter(Position = 0, Mandatory = $True)]
    [String]$pcname,
    [Parameter(Position = 1, Mandatory = $True)]
    [String]$uname
    )
    $runningBlockFunction = [ScriptBlock]::Create({
    function getInfo {
        [CmdletBinding()] Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [String]$Computer,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$User
        )
        $ErrorActionPreference = 'SilentlyContinue' ## Suppress errors
        $PCinfo = Get-WmiObject Win32_ComputerSystem
        $BIOSinfo = Get-WmiObject Win32_Bios
        $OSinfo = Get-WmiObject Win32_OperatingSystem
        $CPUinfo = Get-WmiObject Win32_Processor
        $GPUinfo = Get-WmiObject Win32_VideoController
        $HDDinfo = Get-WmiObject Win32_DiskDrive
        $freeSpace = Get-WmiObject Win32_LogicalDisk -Filter { DeviceID = "C:" }
        [String]$TBLfmt = "MBR"
        if ((Get-WmiObject Win32_DiskPartition | Where-Object BootPartition -eq $true | Select-Object Type).Type -like "*GPT*"){ [String]$TBLfmt = "GPT" }
        $installedMemory = $PCinfo.TotalPhysicalMemory
        $installedMemoryGB = [Math]::Ceiling($installedMemory / [Math]::Pow(1024,3))
        if ($OSinfo.Version -match "6.1") {
            $OS = "Windows 7"
        } 
        if ($OSinfo.Version -match "6.3") {
            $OS = "Windows 8.1"
        }
        if ($OSinfo.Version -match "10.0") {
            #$OS = "Windows 10"
            $OS = [String]($OSinfo.Version)
        }
        $PCmodel = $PCinfo.Model
        if ($PCmodel.Length -ge 14) {
            while ($PCmodel -like "*  *") {
                $PCmodel = $PCmodel -replace "  ", "" #Remove whitespace
            }
            if ($PCmodel.Length -ge 14) {
                $PCmodel = $PCmodel.Substring(0,13) #restrict to first 13 chars
            }
        }
        $CPU = $CPUinfo.Name
        while (($CPU -like "*  *") -or ($CPU -like "*`t*")) {
            $CPU = $CPU -replace "  ", " "
            $CPU = $CPU -replace "`t", " "
        }
        $cpuArch = [String]$OSinfo.OSArchitecture
        $memString = [String]$installedMemoryGB + "GB"
        $hdd = @()
        $HDDinfo | ForEach-Object { $hdd += ([String]([Math]::Ceiling($_.Size / [Math]::Pow(1024,3))) + "GB") }
        $GPU = $GPUinfo.Name
        $BIOS = $BIOSinfo.SMBIOSBIOSVersion
        $pctfree = @()
        $freeSpace | ForEach-Object { if($_.Size -gt 0){ $pctfree += [Math]::Round(($_.FreeSpace / $_.Size)*100) }}
        If ($memString -notlike "0*") {
            $props = @{
                Computer = $Computer
                User = $User
                Model = $PCmodel
                BIOS = $BIOS
                DISK = $TBLfmt
                CPU = $CPU
                RAM = $memString
                HDD = $hdd
                FREE = $pctfree
                OS = $OS
                Arch = $cpuArch
                GPU = $GPU
            }
            $computerInfo = @()
            $computerInfo = New-Object PSObject -Property $props
            return $computerInfo
        }
    }

    })
    $runningBlock = [ScriptBlock]::Create($runningBlockFunction.ToString() + "getInfo -Computer $pcname -User $uname")

    if (!(Test-Connection -ComputerName $pcname -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { 
        Write-Host "......Unreachable" -ForegroundColor Red
    } else {
        if ($pcname -ne $env:COMPUTERNAME){
            Invoke-Command -ComputerName $pcname -Credential $cred -ScriptBlock $runningBlock -AsJob
        } else {
            Start-Job -ScriptBlock $runningBlock
        }
    }

}

##loop through arrays to get computer info
for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$iuser = $user[$i]
	$icomputer = $computer[$i]
    $currentOpp = "Checking PC info on " + $icomputer + " " + $iuser
    Write-Output $currentOpp
    HWcomputerInfo -pcname $icomputer -uname $iuser

}

$count = 0
$njobs = (Get-Job -State Running).count
Clear-Host
Write-Host -NoNewline "Waiting for most jobs ($njobs) to complete" -ForegroundColor White
## Wait for jobs to finish
While (((Get-Job -State Running).count -gt 0) -and ($count -lt 60)){
    $jobs = (Get-Job -State Running).count
    if ($njobs -ne $jobs){
        Clear-Host
        Write-Host -NoNewline "Waiting for most jobs ($jobs) to complete" -ForegroundColor White
        for ($i = 0 ; $i -lt $count + 1 ; $i ++){ Write-Host -NoNewline "." -ForegroundColor White }
        $njobs = $jobs
    } else {
        Write-Host -NoNewline "." -ForegroundColor White
    }
    Start-Sleep -Seconds 1 # Wait for any running jobs to finish
    $count = $count + 1
    if ((Get-Job -State Running).count -eq 0){$count = 60}
}
Clear-Host

## Return computerInfo to an array of objects
$computerInfos = @()
foreach ($job in (Get-Job -State Completed)){
    $pcname = $job.Location
    if ($pcname -eq "localhost"){ $pcname = $env:COMPUTERNAME }
    $computerInfos += ((Get-Job -Name $job.Name) | Receive-Job)
}
Get-Job | Remove-Job

# Array of objects has been stored in $computerInfos
$computerInfos | Select-Object Computer, User, Model, BIOS, DISK, CPU, Arch, RAM, HDD, FREE, OS, GPU | Out-GridView

## Save computer info to XML
#$computerInfos | Export-Clixml -Path infos.xml
## Read computer info from XML
#$computerInfos2 = Import-Clixml -Path infos.xml