#Execute
###############################################################################################
# Flushdns on list of Machines
###############################################################################################
## Admin Name
$adminName = $env:adminName
## Path to list of computer-names and user-names
$pcList = $env:COMPUTERSLIST
###############################################################################################
if ($cred) {} else { $cred = Get-Credential $adminName }

$scriptBlock = [ScriptBlock]::Create("cmd /c `"ipconfig /flushdns`"")

$user = @()
$computer = @()
Import-Csv $pcList | ForEach-Object {
        $user += $_.user
        $computer += $_.computer
}

for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$icomputer = $computer[$i]
    $iuser = $user[$i]
    echo "Flushing DNS on $icomputer ($iuser's PC)"
    if (!(Test-Connection -ComputerName $icomputer -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { echo "...Unreachable" } else {
        Invoke-Command -ComputerName $icomputer -Credential $cred $scriptBlock
    }
}