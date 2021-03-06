#ExecuteOut
###############################################################################################
# Connect to remote PC
#
# The remote PC must be joined to the domain and the remote PC
# must have WinRM enabled already
###############################################################################################

$pcname = Read-Host 'PC-Name to connect to'
if ($cred) {} else { $cred = Get-Credential $env:adminName } # Get credential for connection
New-PSSession -ComputerName $pcname -Credential $cred | Enter-PSSession # Connect to $pcname