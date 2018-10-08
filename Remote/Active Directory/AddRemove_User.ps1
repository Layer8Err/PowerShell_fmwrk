#Execute
###############################################################################################
# Add or remove users from Active Directory
# You will probably want to modify this as needed
###############################################################################################
Import-Module ActiveDirectory # This is required
$adminName = $env:adminName 
if ($cred) {} else { $cred = Get-Credential $adminName }
$server = $env:LOGONSERVER.Substring(2, ($env:LOGONSERVER).Length - 2)

#### Domain user defaults:
$emailDomain = "@acme.com"
$defaultPhone = '513-867-5309'
$CompanyName = 'ACME, Inc.'
$StreetAddress = '1234 Main St.'
$City = 'Cincinnati'
$State = 'OH'
$PostalCode = '45040'
$Country = 'US'
# New User OU
$newUserOU = 'OU=NewUsers,OU=Users,OU=AcmeInc,DC=corp,DC=acme,DC=com'
$newUserPassword = 'Highly1ns3curePasswordOnGithub!'
$activeUsersOU = 'OU=ActiveUsers,OU=Users,OU=AcmeInc,DC=corp,DC=acme,DC=com'
$deactivatedUsersOU = 'OU=ActiveUsers,OU=Users,OU=AcmeInc,DC=corp,DC=acme,DC=com'

## Create new user account
function createUser {
    $FirstName = Read-Host 'First Name '
    $LastName = Read-Host 'Last Name  '
    $PhoneNum = Read-Host 'Phone # (XXX-XXX-XXXX)' # May not be known at the time of setup
    $computer = Read-Host 'Primary PC ' # May not be known at the time of setup
    ## Format Capitalization Properly
    $FirstName = $FirstName.ToUpper().Substring(0,1) + $FirstName.ToLower().Substring(1,([INT](($FirstName | Measure-Object -Character).Characters) - 1))
    $LastName = $LastName.ToUpper().Substring(0,1) + $LastName.ToLower().Substring(1,([INT](($LastName | Measure-Object -Character).Characters) - 1))
    $FullName = $FirstName + " " + $LastName
    $LoginName = ($FirstName.Substring(0,1) + $LastName).ToLower()
    $email = ($FirstName).Substring(0,1) + "." + $LastName + $emailDomain
    $UPNLogin = $email
    Write-Host "--------------------------------------------" -ForegroundColor White
    ## Check Info
    Write-Output "FirstName  $FirstName"
    Write-Output "LastName   $LastName"
    Write-Output "FullName   $FullName"
    Write-Output "LoginName  $LoginName"
    Write-Output "UPNLogin $UPNLogin"
    Write-Output "email      $email"
    Write-Output "Primary PC $computer"
    Write-Output " "
    Write-Output "Hit ENTER to create new user"
    PAUSE
    Write-Host "--------------------------------------------" -ForegroundColor White
    ## Create New User in service-and-terminated
    Write-Host "Creating New Active Directory user account for $LoginName (in service-and-terminated OU)..." -ForegroundColor Cyan
    New-ADUser -Credential $cred -Name $FullName -GivenName $FirstName -Surname $LastName `
        -SamAccountName $LoginName -EmailAddress $email -UserPrincipalName $email -Path $newUserOU `
        -AccountPassword (ConvertTo-SecureString -String $newUserPassword -AsPlainText -Force)
    
    ## Set up User Properties
    Write-Host "Allowing $LoginName to change their password..." -ForegroundColor Cyan
    Set-ADUser -Credential $cred -Identity $LoginName -CannotChangePassword $false
    Write-Host "Setting company address for $LoginName..." -ForegroundColor Cyan
    Set-ADUser -Credential $cred -Identity $LoginName -StreetAddress $StreetAddress -City $City -State $State -PostalCode $PostalCode -Company $CompanyName -Country $Country
    ## Set up Phone Info
    if ($PhoneNum.Length -lt 11){
        Write-Host "Using $defaultPhone as the dummy phone number" -ForegroundColor Yellow
        $PhoneNum = $defaultPhone
    }
    if (($PhoneNum.Substring(3,1) -eq "-") -and ($PhoneNum.Substring(7,1) -eq "-")){
        $officePhone = "$PhoneNum"
        $PhoneExt = $PhoneNum.Substring(8,4)
        Write-Host "Setting up phone number info..." -ForegroundColor Cyan
        Set-ADUser -Credential $cred -Identity $LoginName -Add @{ipPhone="$PhoneExt"}
        Set-ADUser -Credential $cred -Identity $LoginName -OfficePhone $officePhone
        Set-AdUser -Credential $cred -Identity $LoginName -Add @{telephoneNumber="$PhoneNum"}
    }
    Write-Host "Setting DisplayName and Description for $LoginName..." -ForegroundColor Cyan
    Set-ADUser -Credential $cred -Identity $LoginName -DisplayName $FullName -Description $FullName
    ## SBS Stuff for SBS Console (God help you if you actually need this crap)
    # Write-Host "Adding custom attribute `'msSBSCreationState=`"Created`"`'..." -ForegroundColor Cyan
    # Set-ADUser -Credential $cred -Identity $LoginName -Add @{msSBSCreationState="Created"}
    # Write-Host "Adding custom attribute `'msSBSRoleGuid=`"f5aab79f-b306-4d93-b805-36e5597b7654`"`'..." -ForegroundColor Cyan
    # Set-ADUser -Credential $cred -Identity $LoginName -Add @{msSBSRoleGuid="f5aab79f-b306-4d93-b805-36e5597b7654"}
    ## Set Users default computer (where they are an adimin)
    if ($computer.Length -gt 0){
        $computerCN = "S:1:5:" + (Get-ADComputer -Identity $computer).DistinguishedName
        if ($computerCN.Length -gt 0){
            $computers = @()
            $computers += $computerCN
            Write-Host "Setting up $computer for $LoginName..." -ForegroundColor Cyan
            Set-ADUser -Credential $cred -Identity $LoginName -Add @{msSBSComputerUserAccessOverride="$computerCN"}
        }
    } else {
        Write-Host "Not setting the user`'s PC at this time (normal)" -ForegroundColor Yellow
    }
    ## Add AD User to groups
    Write-Host "Adding $LoginName to `'All Users`' group..." -ForegroundColor Cyan
    Add-ADGroupMember -Credential $cred "All Users" $LoginName

    ## Add User proxyAddresses for Office 365
    $proxyAddresses = @()
    $proxyAddresses += "SMTP:" + $email
    $proxyAddresses += "smtp:" + ($FirstName).Substring(0,1) + $LastName + $emailDomain
    Write-Host "Applying proxyAddresses for $LoginName..." -ForegroundColor Cyan
    Set-ADUser -Credential $cred -Identity $LoginName -Replace @{proxyAddresses=$proxyAddresses}

    ## Set up Office 365 mailNickname
    $mailNickname = (($FirstName).Substring(0,1) + "." + $LastName).ToLower()
    Write-Host "Setting custom attribute `'mailNickname=`"$mailNickname`"`'..." -ForegroundColor Cyan
    Set-ADUser -Credential $cred -Identity $LoginName -Add @{mailNickname="$mailNickname"}
    # Do not enable $LoginName at this time
    Write-Host "Setting user login enabled to FALSE..." -ForegroundColor Cyan
    Set-ADUser -Credential $cred -Identity $LoginName -Enabled $false
    # Move user to Regular users OU
    Write-Host "Moving $LoginName to Regular users OU..." -ForegroundColor Yellow
    Get-ADUser -Credential $cred -Identity $LoginName | Move-ADObject -Credential $cred -TargetPath $activeUsersOU
    $activate = Read-Host "Activate $LoginName now? [y/N]"
    if ($activate.Length -gt 0){
        if ( $activate.ToLower() -eq "y" ){
            Set-ADUser -Credential $cred -Identity $LoginName -Enabled $true
        }
    }
    #$setAdmin = Read-Host "Set as $LoginName as admin on $computer now? [y/N]" # Ideally users should never be admins
    if ($setAdmin.Length -gt 0){
        if ($setAdmin.ToLower() -eq "y"){
            Write-Host "Attempting to set $LoginName as an admin on $computer" -ForegroundColor Cyan
            $localAccountsBlockFunction = [ScriptBlock]::Create({
                function addToAdmin ($newUsername){
                    Add-LocalGroupMember -Group Administrators -Member $newUsername
                }
            })
            $localAccountsBlock = [ScriptBlock]::Create($localAccountsBlockFunction.ToString() + "addToAdmin -newUsername '" + "$env:domainName" + '\' + $LoginName + "'")
            if (Test-Connection -ComputerName $computer -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet) { 
                Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock $localAccountsBlock
            }
        }
    }
    Write-Host "New user setup for $FirstName $LastName ($LoginName) is done." -ForegroundColor Green
    Write-Host "You may need to use the UpdateUser script to finish adding user attributes." -ForegroundColor Yellow
    Pause
}

## Deactivate a user account
function deactivateUser {
    Write-Host "This will lock a user's account and move them to the `"Service and Terminated`" OU"
    $termUserSam = Read-Host 'SamAccountName to disable (e.g. jdoe)'
    if ($cred) {} else { $cred = Get-Credential $adminName }
    Write-Output "Hit ENTER to disable $termUserSam"
    PAUSE
    ##Set random password for user
    $set = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@$#%*".ToCharArray()
    $result = ""
    for ($i = 0; $i -lt 17 ; $i++){
        $result += $set | Get-Random
    }
    Write-Host "Setting random password for $termUserSam..." -ForegroundColor Cyan
    Set-ADAccountPassword -Credential $cred -Identity $termUserSam -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $result -Force)
    #Lock user account
    Write-Host "Locking AD account..." -ForegroundColor Yellow
    $status = net user $termUserSam /DOMAIN | Find /I "Account Active"
    if ([string]$status -like '*No*'){
        Write-Host "  Account is already locked... nothing to do" -ForegroundColor Green
    } else {
        Write-Host "  Account is not locked" -ForegroundColor Yellow
        Write-Host "   Locking account.." -ForegroundColor Cyan
        Set-ADUser -Credential $cred -Identity $UserName -Enabled $false
        $rmtCmd = "net user $termUserSam /domain /active:no"
        $scriptBlock = [ScriptBlock]::Create($rmtCmd) ## Prevent logins
        Invoke-Command -ComputerName $server -Credential $cred -ScriptBlock $scriptBlock
        $status = net user $termUserSam /DOMAIN | Find /I "Account Active"
        Write-Host "    $status" -ForegroundColor White
    }
    Write-Host "Moving $termUserSam to `"Deactivated`" OU..." -ForegroundColor Yellow
    #Move user to $deactivatedUsersOU
    Get-ADUser $termUserSam | Move-ADObject -Credential $cred -TargetPath $deactivatedUsersOU
    Write-Host "$termUserSam`'s account is fully locked out." -ForegroundColor Green
    Write-Host "You may need to edit another user`'s proxyAddresses so that they can receive $termUserSam`'s email." -ForegroundColor Yellow
    PAUSE
}

## AddRemove menu
function addremoveMenu {
    Clear-Host
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Write-Host "=======Add or Remove User======" -ForegroundColor White
    Write-Host "1`tAdd new user account" -ForegroundColor Green
    Write-Host "2`tDeactivate user account" -ForegroundColor Red
    [INT]$addremovechoice = Read-Host ">"
    if ($addremovechoice -eq 1){ createUser }
    if ($addremovechoice -eq 2){ deactivateUser }
}

addremoveMenu
