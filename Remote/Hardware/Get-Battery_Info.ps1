#Execute
###############################################################################################
# Get Battery Info
###############################################################################################
# Path to list of computer-names and user-names
$pcList = $env:COMPUTERSLIST
$adminName = $env:adminName
if ($cred) {} else { $cred = Get-Credential $adminName }
#######################################*Begin Script*##########################################

$user = @()
$computer = @()
$batteries = @()

#import list of usernames and computernames from .csv
#save into two arrays
Import-Csv $pcList | ForEach-Object {
        $user += $_.user
        $computer += $_.computer
}

#get my computername
$MycompName = [string]((Get-WmiObject Win32_ComputerSystem).Name)

# Helper functions
$scriptBlock = [ScriptBlock]::Create({Get-WmiObject -Class Win32_Battery})
# (Get-Win32_Battery -wmibattery (Get-WmiObject -Class Win32_Battery)) | Format-Table
function Get-Win32_Battery ($wmibattery){
    $DeviceID = [string]($wmibattery.DeviceID)
    $TimeOnBattery = [string]($wmibattery.TimeOnBattery)
    $EstimatedChargeRemaining = [int]($wmibattery.EstimatedChargeRemaining)
    $EstimatedRunTime = [string]($wmibattery.EstimatedRunTime)
    $Availability = Switch ([string]($wmibattery.Availability)){
        1 {"Other";break}
        2 {"Unknown";break}
        3 {"Running or Full Power";break}
        4 {"Warning";break}
        5 {"In Test";break}
        6 {"Not Applicable";break}
        7 {"Power Off";break}
        8 {"Off Line";break}
        9 {"Off Duty";break}
        10 {"Degraded";break}
        11 {"Not Installed";break}
        12 {"Install Error";break}
        13 {"Power Save - Unknown";break}
        14 {"Power Save - Low Power Mode";break}
        15 {"Power Save - Standby";break}
        16 {"Power Cycle";break}
        17 {"Power Save - Warning";break}
    }
    $BatteryStatus = Switch ([string]($wmibattery.Availability)){
        1 {"Discharging";break}
        2 {"On A/C";break}
        3 {"Fully Charged";break}
        4 {"Low";break}
        5 {"Critical";break}
        6 {"Charging";break}
        7 {"Charging High";break}
        8 {"Charging Low";break}
        9 {"Charging Critical";break}
        10 {"Undefined";break}
        11 {"Partially Charged";break}
    }
        
    $properties = @{
        'Computer'=$icomputer;
        'User'=$iuser;
        'DeviceID'=$DeviceID;
        'TimeOnBattery'=[INT]$TimeOnBattery;
        'EstChargeRemaining'=[INT]$EstimatedChargeRemaining;
        'EstRunTime'=[INT]$EstimatedRunTime;
        'Availability'=$Availability;
        'Status'=$BatteryStatus}
    $battery = New-Object PSObject -Property $properties
    return $battery
}
# Main loop
for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$iuser = $user[$i]
	$icomputer = $computer[$i] 
	$currentOpp = "Querying battery info on " + $icomputer + " (" + $iuser + "`'s computer)"
	Write-Output $currentOpp
    if (!(Test-Connection -ComputerName $icomputer -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { echo "Unreachable!" } else {
        if ($icomputer -ne $MycompName){
            $batteryInfo = Invoke-Command -ComputerName $icomputer -Credential $cred -ScriptBlock $scriptBlock
        } else {
            $batteryInfo = Get-WmiObject -Class Win32_Battery
        }
        $batteries += (Get-Win32_Battery -wmibattery $batteryInfo)
    }
}
Get-PSSession | Remove-PSSession

## Show all battery data
#$batteries | Select Computer, User, DeviceID, EstRunTime, EstChargeRemaining, Status | Out-GridView

## Show all battery data where est runtime < 30
#$batteries | Where-Object -Property EstRunTime -lt 30 | Where-Object -Property EstChargeRemaining -gt 0 | Select Computer, User, DeviceID, EstRunTime, EstChargeRemaining, Status | Out-GridView
$batteries | Sort-Object -Property User |Where-Object -Property EstChargeRemaining -gt 0 | Select-Object Computer, User, DeviceID, EstRunTime, EstChargeRemaining, Status | Out-GridView
Pause