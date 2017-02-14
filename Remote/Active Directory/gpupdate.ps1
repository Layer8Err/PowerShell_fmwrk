#Execute
###############################################################################################
# Update Group Policy (/force) on a list of machines
###############################################################################################
# Path to list of computer-names and user-names
$pcList = $env:COMPUTERSLIST
$adminName = $env:adminName # Domain admin
###############################################################################################

$user = @()
$computer = @()

Import-Csv $pcList | ForEach-Object {
    $user += $_.user
    $computer += $_.computer
}

if ($cred) {} else { $cred = Get-Credential $adminName }

for ( $i=0; $i -le ($user.Length - 1); $i++) {
    $iuser = $user[$i]
	$icomputer = $computer[$i]
	$currentOpp = "Group Policy Update for " + $icomputer + " (" + $iuser + "`'s computer)"
	Write-Output $currentOpp
    Invoke-Command -ComputerName $icomputer -Credential $cred -ScriptBlock { cmd.exe /c "gpupdate /force" } -AsJob
}
Clear-Host

$count = 0
$njobs = (Get-Job -State Running).count
Write-Host -NoNewLine "Waiting for all jobs ($njobs) to complete" -ForegroundColor White
While ((Get-Job -State Running).count -gt 0){
    $jobs = (Get-Job -State Running).count
    if ($jobs -ne $njobs){ 
        Clear-Host
        Write-Host -NoNewLine "Waiting for all jobs ($jobs) to complete" -ForegroundColor White
        $njobs = $jobs
        for ($i = 0 ; $i -lt $count + 1 ; $i ++){ Write-Host -NoNewline "." -ForegroundColor White } # Add total dots to the end
    }
    Write-Host -NoNewline "." -ForegroundColor White # Only add one more dot
    ++$count
    Start-Sleep -Seconds 1 # Wait for any running jobs to finish
}
Write-Host -NoNewLine "Done" -ForegroundColor White
Start-Sleep -Seconds 2
Get-Job | Remove-Job