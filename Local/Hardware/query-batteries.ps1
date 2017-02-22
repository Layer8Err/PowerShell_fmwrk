#Execute
###############################################################################################
# Query Battery Info
#
# Get Battery level on local PC
###############################################################################################

$battery = New-Object PSObject
$properties = @{}

#get my computername
$MycompName = [string]((Get-WmiObject Win32_ComputerSystem).Name)
#$currentOpp = "Querying battery info on " + $MycompName
#Write-Output $currentOpp
    
$batteryInfo = Get-WmiObject -Class Win32_Battery # Grab the Win32 Battery info

$DeviceID = [string]($batteryInfo.DeviceID)
$TimeOnBattery = [string]($batteryInfo.TimeOnBattery)
$EstimatedChargeRemaining = [int]($batteryInfo.EstimatedChargeRemaining)
$EstimatedRunTime = [string]($batteryInfo.EstimatedRunTime)
$Availability = Switch ([string]($batteryInfo.Availability)){
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
$BatteryStatus = Switch ([string]($batteryInfo.Availability)){
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
    'Computer'=$MycompName;
    'DeviceID'=$DeviceID;
    'TimeOnBattery'=[INT]$TimeOnBattery;
    'EstChargeRemaining'=[INT]$EstimatedChargeRemaining;
    'EstRunTime'=[INT]$EstimatedRunTime;
    'Availability'=$Availability;
    'Status'=$BatteryStatus
}
$battery = New-Object PSObject -Property $properties
$battery | Select-Object Status, EstChargeRemaining, TimeOnBattery, DeviceID | Format-Table
pause