#Execute
###############################################################################################
# Create new user account and assign to the appropriate OUs
# This script sets AD properties for a new user
# You should heavily modify this script before use in a production environment
###############################################################################################
## Constants ##
$adminName = $env:adminName
$emailDomain = "contoso.com"
$dummyPasswd = "Th3P4ssw0rd123!"
$creationOU = 'OU=creationOU,OU=Users,OU=ThisBusiness,DC=contoso,DC=local'
$activeUserOU = 'OU=activeOU,OU=Users,OU=ThisBusiness,DC=contoso,DC=local'
$Company = "Contoso Inc."
$StreetAddress = "101 Main Street"
$City = "Beverly Hills"
$State = "CA"
$PostalCode = "90210"
$Country = "US"
###############################################################################################
if ($cred) {} else { $cred = Get-Credential $adminName }
Import-Module ActiveDirectory

## Create New User ##

$FirstName = Read-Host 'First Name '
$LastName = Read-Host 'Last Name  '
$PhoneNum = Read-Host 'Phone # (XXX-XXX-XXXX)'
$computer = Read-Host 'Primary PC '

## Format Capitalization Properly
$FirstName = $FirstName.ToUpper().Substring(0,1) + $FirstName.ToLower().Substring(1,([INT](($FirstName | Measure-Object -Character).Characters) - 1))
$LastName = $LastName.ToUpper().Substring(0,1) + $LastName.ToLower().Substring(1,([INT](($LastName | Measure-Object -Character).Characters) - 1))
$FullName = $FirstName + " " + $LastName
$LoginName = $FirstName.Substring(0,1) + $LastName
$email = ($FirstName).Substring(0,1) + "." + $LastName + "@" + $emailDomain

## Check Info
Write-Output "FirstName  $FirstName"
Write-Output "LastName   $LastName"
Write-Output "FullName   $FullName"
Write-Output "LoginName  $LoginName"
Write-Output "email      $email"
Write-Output "Primary PC $computer"
Write-Output " "
Write-Output "Hit ENTER to create new user"
PAUSE

## Create New User in creation OU # this line should be modified
New-ADUser -Credential $cred -Name $FullName -GivenName $FirstName -Surname $LastName -SamAccountName $LoginName -EmailAddress $email -UserPrincipalName $email -Path $creationOU -AccountPassword (ConvertTo-SecureString -String $dummyPasswd -AsPlainText -Force)

## Set up User Properties
Set-ADUser -Credential $cred -Identity $LoginName -CannotChangePassword $false
Set-ADUser -Credential $cred -Identity $LoginName -StreetAddress $StreetAddress -City $City -State $State -PostalCode $PostalCode -Company $Company -Country $Country

## Set up Phone Info
if (($PhoneNum.Substring(3,1) -eq "-") -and ($PhoneNum.Substring(7,1) -eq "-")){
    $officePhone = "$PhoneNum"
    $PhoneExt = $PhoneNum.Substring(8,4)
    Set-ADUser -Credential $cred -Identity $LoginName -Add @{ipPhone="$PhoneExt"}
    Set-ADUser -Credential $cred -Identity $LoginName -OfficePhone $officePhone
    Set-AdUser -Credential $cred -Identity $LoginName -Add @{telephoneNumber="$PhoneNum"}
}
Set-ADUser -Credential $cred -Identity $LoginName -DisplayName $FullName -Description $FullName

##Add-ADGroupMember -Credential $cred "Domain Users" $LoginName ## Not needed
Add-ADGroupMember -Credential $cred "All Users" $LoginName
Add-ADGroupMember -Credential $cred "Contoso" $LoginName

## Add User proxyAddresses for Office 365
$proxyAddresses = @()
$proxyAddresses += "SMTP:" + $email
$proxyAddresses += "smtp:" + ($FirstName).Substring(0,1) + $LastName + "@" + $emailDomain

Set-ADUser -Credential $cred -Identity $LoginName -Replace @{proxyAddresses=$proxyAddresses}

## Set up Office 365 mailNickname
$mailNickname = (($FirstName).Substring(0,1) + "." + $LastName).ToLower()
Set-ADUser -Credential $cred -Identity $LoginName -Add @{mailNickname="$mailNickname"}
Set-ADUser -Credential $cred -Identity $LoginName -Enabled $false

# Move user to Active OU
Write-Host "Moving $LoginName to Active OU..." -ForegroundColor Yellow
Get-ADUser -Credential $cred -Identity $LoginName | Move-ADObject -Credential $cred -TargetPath $activeUserOU

$activate = Read-Host "Activate $LoginName now? [y/N]"
if ( $activate.ToLower().Substring(0,1) -eq "y" ){
    Set-ADUser -Credential $cred -Identity $LoginName -Enabled $true
}