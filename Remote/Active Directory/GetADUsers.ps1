#Execute
########################
# Display a list of Active Directory Users
Import-Module ActiveDirectory
$users = Get-ADUser -Filter {(Enabled -eq $true) -and (ObjectClass -eq "user") -and (EmailAddress -like "*.com")} | Select-Object SamAccountName
#$users = Get-ADUser -Filter {(Enabled -eq $true) -and (ObjectClass -eq "user")} | Select-Object SamAccountName
#$users = Get-ADUser -Filter {(Enabled -eq $false)} | Select SamAccountName
$userInfo = @()
$users | ForEach-Object {
    $userInfo += Get-ADUser -Identity $_.SamAccountName -Properties Name, Created, Department, Title, EmailAddress, StreetAddress, City, State, PostalCode, Company, telephoneNumber, OfficePhone, mobile, MobilePhone, ipPhone  | Select Name, SamAccountName, Created, EmailAddress, UserPrincipalName, Title, Department, Company, StreetAddress, City, State, PostalCode, OfficePhone, telephoneNumber, ipPhone, MobilePhone, mobile
}
$userInfo | Out-GridView