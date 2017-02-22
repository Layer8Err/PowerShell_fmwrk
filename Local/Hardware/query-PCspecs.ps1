#Execute
###############################################################################################
# Get System Information
#
# List the local computer's hardware/software
###############################################################################################

#get my computername
$MycompName = [string]((Get-WmiObject Win32_ComputerSystem).Name)
Write-Host "Getting localll hardware info..." -ForegroundColor Cyan
$results = ""

$PCinfo = Get-WmiObject -ComputerName $MycompName Win32_ComputerSystem
$OSinfo = Get-WmiObject -ComputerName $MycompName Win32_OperatingSystem
$CPUinfo = Get-WmiObject -ComputerName $MycompName Win32_Processor
$installedMemory = $PCinfo.TotalPhysicalMemory
$installedMemoryGB = [Math]::Ceiling($installedMemory / [Math]::Pow(1024,3))
if ($OSinfo.Version -match "6.1") {
    $OS = "Windows 7"
} 
if ($OSinfo.Version -match "6.3") {
    $OS = "Windows 8.1"
}
if ($OSinfo.Version -match "10.0") {
    $OS = "Windows " + [String]($OSinfo.Version)
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
$memString = [String]$installedMemoryGB + "GB RAM"
If ($memString -notlike "0*") {
    $props = @{
        Computer = $MycompName
        Model = $PCmodel
        CPU = $CPU
        RAM = $memString
        OS = $OS
        Arch = $cpuArch
    }
    $computerInfo = New-Object PSObject -Property $props
}

$computerInfo | Select-Object Model, CPU, Arch, RAM, OS | Format-Table
pause