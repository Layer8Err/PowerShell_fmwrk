#Execute
###############################################################################################
# Clear Outlook Autocomplete cache
# 
# Clear Outlook Autocomplete cache on remote PC
# This resolves issues with autocomplete entries
###############################################################################################
$pcList = $env:COMPUTERSLIST
$adminName = $env:adminName
if ($cred) {} else { $cred = Get-Credential $adminName }
###############################################################################################
#######################################*Begin Script*##########################################
###############################################################################################

$user = @()
$computer = @()
$pcname = Read-Host 'PC-Name to clear Outlook cache'

#import list of usernames and computernames from .csv
#save into two arrays
Import-Csv $pcList | ForEach-Object {
        $user += $_.user
        $computer += $_.computer
}

#loop through arrays to find $userName
for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$iuser = $user[$i]
	$icomputer = $computer[$i]
    if ($icomputer -eq $pcname) { $userName = $iuser }
}
############
$cacheCleanBlock = [scriptblock]::Create({
    function cleanAutocomplete ($username) {
        Kill -Name OUTLOOK -Force
        $user = $username
        $path = "C:\Users\$user\AppData\Local\Micorosft\Outlook\RoamCache\"
        $delString = $path + "Stream_Autocomplete*.dat"
        Remove-Item -Path $delString -Force
    }
})
$cleanUserCache = [ScriptBlock]::Create($cacheCleanBlock.ToString() + "cleanAutocomplete `"$userName`"")

#Execute Outlook Autocomplete Cache Clear
Invoke-Command -Computer $icomputer -Credential $cred -ScriptBlock $cleanUserCache
