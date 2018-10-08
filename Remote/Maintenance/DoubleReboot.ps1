#Execute
# Schedule reboot
# Reboot

$ListOfComputers = $env:COMPUTERSLIST
$adminName = $env:adminName
if ($cred) {} else { $cred = Get-Credential $adminName }
$pcList = $env:COMPUTERSLIST

$user = @()
$computer = @()
#import list of usernames and computernames from .csv
Import-Csv $pcList | ForEach-Object {
        $user += $_.user
        $computer += $_.computer
}

# Double reboot script-block
$remoteRebootBlock = [scriptblock]::Create({
    function buildTaskXML {
        $tasktext1 = '<?xml version="1.0" encoding="UTF-16"?>
    <Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
    <RegistrationInfo>
        <Date>2018-08-31T09:42:50.4597794</Date>
        <Author>Administrator</Author>
        <URI>\SchedReboot</URI>
    </RegistrationInfo>
    <Triggers>
        <TimeTrigger>
        <StartBoundary>'
        $tasktext2 = '</StartBoundary>
        <EndBoundary>2019-09-02T10:30:00</EndBoundary>
        <Enabled>true</Enabled>
    </TimeTrigger>
    </Triggers>
    <Principals>
        <Principal id="Author">
        <UserId>S-1-5-21-2403106016-3164382704-2769491958-500</UserId>
        <LogonType>Password</LogonType>
        <RunLevel>HighestAvailable</RunLevel>
        </Principal>
    </Principals>
    <Settings>
        <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
        <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
        <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
        <AllowHardTerminate>true</AllowHardTerminate>
        <StartWhenAvailable>false</StartWhenAvailable>
        <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
        <IdleSettings>
        <StopOnIdleEnd>true</StopOnIdleEnd>
        <RestartOnIdle>false</RestartOnIdle>
        </IdleSettings>
        <AllowStartOnDemand>true</AllowStartOnDemand>
        <Enabled>true</Enabled>
        <Hidden>false</Hidden>
        <RunOnlyIfIdle>false</RunOnlyIfIdle>
        <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
        <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
        <WakeToRun>true</WakeToRun>
        <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
        <DeleteExpiredTaskAfter>P30D</DeleteExpiredTaskAfter>
        <Priority>7</Priority>
    </Settings>
    <Actions Context="Author">
        <Exec>
        <Command>cmd.exe</Command>
        <Arguments>-c "shutdown /t 030 /r /f"</Arguments>
        </Exec>
    </Actions>
</Task>'
        $SecondRebootTime = Get-Date -Date ($(Get-Date).AddMinutes(40)) -Format yyyy-MM-ddTHH:mm:00
        # Schedule reboot out 40 minutes
        $tasktext = $tasktext1 + [String]$SecondRebootTime + $tasktext2
        Return $tasktext
    }
    function createRebootTask ($rebootTask, $taskPath) {
        echo "" > $taskPath
        echo $rebootTask > $taskPath
        schtasks /create /tn RebootTwice /np /xml $taskPath
        Remove-Item -Path $taskPath
    }
    if(!(Test-Path 'C:\Windows\Temp\RebootTwice')){
        mkdir 'C:\Windows\Temp\RebootTwice'
    }
    $taskpath = 'C:\Windows\Temp\RebootTwice\RebootTwice.xml'
    $taskTxt = buildTaskXML
    createRebootTask -rebootTask $taskTxt -taskPath $taskpath
    Remove-Item -Recurse -Path 'C:\Windows\Temp\RebootTwice' -Force
    cmd -c "shutdown /t 030 /r /f"
})

### Loop through PCs
for ( $i=0; $i -le ($user.Length - 1); $i++) {
    $iuser = $user[$i]
	$icomputer = $computer[$i]
    if (!(Test-Connection -ComputerName $icomputer -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { echo "......Unreachable" } else {
        Write-Host "Scheduling reboot and rebooting $icomputer, $iuser`'s PC" -ForegroundColor Cyan
        ########################################################################
        # Create remote reboot job and reboot
        Invoke-Command -ComputerName $icomputer -Credential $cred -ScriptBlock $remoteRebootBlock -AsJob
        ########################################################################
    }
}