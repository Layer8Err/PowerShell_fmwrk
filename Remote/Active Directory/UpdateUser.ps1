#Execute
###############################################################################################
# Edit Active Directory User Attributes
# You should heavily modify this script before use in a production environment
###############################################################################################
$adminName = $env:adminName
$emailDomain = "contoso.com"
###############################################################################################

Import-Module ActiveDirectory
Clear-Host
if ($cred) {} else { $cred = Get-Credential $adminName }

Write-Host "Edit a user's Active Directory attributes" -ForegroundColor Green
$SamName = Read-Host 'SamAccountName (e.g. dbrenner)'
$error.clear()

function userMenu($SamName){
    $user = @()
    $user = Get-ADUser -Identity $SamName -Properties * 
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    
    $properties = @{
        'GivenName'=$user.GivenName;
        'Surname'=$user.Surname;
        'Email'=$user.EmailAddress;
        'Company'=$user.Company;
        'Department'=$user.Department;
        'Title'=$user.Title;
        'OfficePhone'=$user.OfficePhone;
        'MobilePhone'=$user.MobilePhone;
        #'Computers'=$computers;
    }
    $userItems = New-Object PSObject -Property $properties

    Clear-Host
    Write-Host "=======Active Directory User Edit======" -ForegroundColor White
    Write-Host "1`tFirst Name`t`t$($userItems.GivenName)" -ForegroundColor Cyan
    Write-Host "2`tLast Name`t`t$($userItems.Surname)" -ForegroundColor Cyan
    Write-Host "3`tEmail`t`t`t$($userItems.Email)" -ForegroundColor Cyan
    Write-Host "4`tCompany`t`t`t$($userItems.Company)" -ForegroundColor Cyan
    Write-Host "5`tDepartment`t`t$($userItems.Department)" -ForegroundColor Cyan
    Write-Host "6`tTitle`t`t`t$($user.Title)" -ForegroundColor Cyan
    Write-Host "7`tOffice Phone #`t$($userItems.OfficePhone)" -ForegroundColor Cyan
    Write-Host "8`tMobile Phone #`t$($userItems.MobilePhone)" -ForegroundColor Cyan
    #Write-Host "9`tRemote Web Access Computer(s)" -ForegroundColor Cyan
    #Write-Host "10`tPrinter Group(s)" -ForegroundColor Cyan
    Write-Host "11`tEdit another user" -ForegroundColor Yellow
    Write-Host "`n"
    [INT]$choice = Read-Host ">"

    if ($choice -eq 1){
        [String]$newName = Read-Host "New First Name"
        if ($newName.length -gt 1){
            Write-Host "You should probably update the email address after this" -ForegroundColor White
            [String]$choicen = Read-Host "Are you sure you want to rename $($userItems.GivenName) to $($newName)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Credential $cred -Identity $SamName -GivenName $newName
                $user.GivenName = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties GivenName).GivenName -ne $newName){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    } # Update FirstName (GiveName)
    if ($choice -eq 2){
        [String]$newName = Read-Host "New Last Name"
        if ($newName.length -gt 1){
            Write-Host "You should probably update the email address after this" -ForegroundColor White
            [String]$choicen = Read-Host "Are you sure you want to rename $($userItems.Surname) to $($newName)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Credential $cred -Identity $SamName -Surname $newName
                $user.Surname = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties Surname).Surname -ne $newName){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    } # Update LastName (Surname)
    if ($choice -eq 3){
        $proxyAddresses = @()
        $setEmail = ($user.GivenName).Substring(0,1) + "." + $user.Surname + "@" + $emailDomain
        $proxyAddresses += "SMTP:" + $setEmail
        $proxyAddresses += "smtp:" + ($user.GivenName).Substring(0,1) + $user.Surname + "@" + $emailDomain
            
        Write-Host "This will also change the UPN from $($user.UserPrincipalName) to $setEmail" -ForegroundColor White
        Write-Host "This will change the proxyAddresses to: $proxyAddresses" -ForegroundColor White
        [String]$choicen = Read-Host "Are you sure you want to update $($userItems.Email) to $($setEmail)? [y/N]"
        if ($choicen.Length -lt 1){$choicen = 'n'}
        if ($choicen.ToLower().Substring(0,1) -eq "y"){
            Set-ADUser -Credential $cred -Identity $SamName -EmailAddress $setEmail
            Set-ADUser -Credential $cred -Identity $SamName -Replace @{proxyAddresses=$proxyAddresses}
            Set-ADUser -Credential $cred -Identity $SamName -UserPrincipalName $setEmail
            $user.EmailAddress = $setEmail
            Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
            while ((Get-ADUser -Identity $SamName -Properties EmailAddress).EmailAddress -ne $setEmail){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
        }

        userMenu -SamName $SamName
    } # Automatically update Email (based on FirstName/LastName)
    if ($choice -eq 4){
        $companies = Get-ADUser -Filter { (Enabled -eq $true) -and (ObjectClass -eq "user") } -Properties Company | Select-Object Company | Sort-Object -Property Company -Unique
        $companies | ForEach-Object { 
            $_ | Add-Member Selection (([array]::IndexOf($companies.Company, $_.Company)) + 1)
        } # Add menu numbers (don't start at 0)
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Clear-Host
        Write-Host "=======Edit Company======" -ForegroundColor White
        Write-Host "$($user.Name)`t$($user.Company)" -ForegroundColor White
        Write-Host "Change company to (pick a number)" -ForegroundColor Green
        $companies | Sort-Object -Property Selection | ForEach-Object {
            Write-Host "$($_.Selection)`t$($_.Company)" -ForegroundColor Cyan
        }
        Write-Host "`n"
        [INT]$choice = Read-Host ">"
        if (($choice -lt 1) -or ($choice -gt $companies.Selection.Length)){
            Write-Host "Company not changed" -ForegroundColor Yellow
        } else {
            [String]$choicen = Read-Host "Are you sure you want to change $($userItems.Company) to $($companies[($choice - 1)].Company)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                $newName = $companies[($choice - 1)].Company
                Set-ADUser -Credential $cred -Identity $SamName -Company $newName
                $user.Company = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties Company).Company -ne $newName){ 
                    Start-Sleep 1
                    Write-Host -NoNewline "." -ForegroundColor Yellow
                }
            }
        }
        userMenu -SamName $SamName
    } # Update Company (From list of companies)
    if ($choice -eq 5){
        $departments = Get-ADUser -Filter {(Enabled -eq $true) -and (ObjectClass -eq "user") -and (EmailAddress -like "*.com") -and (SamAccountName -ne "NewtonScheduler")} -Properties Department | Select Department | Sort-Object -Property Department -Unique
        $departments | ForEach-Object { 
            $_ | Add-Member Selection (([array]::IndexOf($departments.Department, $_.Department)) + 1)
        } # Add menu numbers (don't start at 0)
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Clear-Host
        Write-Host "=======Edit Department======" -ForegroundColor White
        Write-Host "$($user.Name)`t$($user.Department)" -ForegroundColor White
        Write-Host "Change Department to (pick a number)" -ForegroundColor Green
        $departments | Sort-Object -Property Selection | ForEach-Object {
            Write-Host "$($_.Selection)`t$($_.Department)" -ForegroundColor Cyan
        }
        Write-Host "`n"
        [INT]$choice = Read-Host ">"
        if (($choice -lt 1) -or ($choice -gt $departments.Selection.Length)){
            Write-Host "Department not changed" -ForegroundColor Yellow
        } else {
            [String]$choicen = Read-Host "Are you sure you want to change $($userItems.Department) to $($departments[($choice - 1)].Department)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                $newName = $departments[($choice - 1)].Department
                Set-ADUser -Credential $cred -Identity $SamName -Department $newName
                $user.Department = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties Department).Department -ne $newName){ 
                    Start-Sleep 1
                    Write-Host -NoNewline "." -ForegroundColor Yellow
                }
            }
        }
        Start-Sleep -Seconds 1
        userMenu -SamName $SamName
    } # Update Department (From list of departments)
    if ($choice -eq 6){
        [String]$newTitle = Read-Host "New Title"
        if (($newTitle.Length -gt 1)){
            [String]$choicen = Read-Host "Are you sure you want to update $($user.Name)'s Title: `"$($user.Title)`" to `"$($newTitle)`"? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Credential $cred -Identity $SamName -Title $newTitle
                $user.Title = $newTitle
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties Title).Title -ne $newTitle){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    } # Update Title
    if ($choice -eq 7){
        [String]$newPhone = Read-Host "New Office Phone # (XXX-XXX-XXXX)"
        if (($newPhone.Length -eq 12) -and ($newPhone.Split("-").Length -eq 3) -and ($newPhone.Split("-")[0].Length -eq 3) -and ($newPhone.Split("-")[1].Length -eq 3) -and ($newPhone.Split("-")[2].Length -eq 4)){
            [String]$choicen = Read-Host "Are you sure you want to update the office Phone # $($userItems.OfficePhone) to $($newPhone)? [y/N]"
            if ($choicen.Length -lt 1){ $choicen = 'n' }
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Credential $cred -Identity $SamName -OfficePhone $newPhone
                $phoneExt = $newPhone.Substring(8,4)
                Set-ADUser -Credential $cred -Identity $SamName -Replace @{ipPhone="$phoneExt"}
                $user.OfficePhone = $newPhone
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties OfficePhone).OfficePhone -ne $newPhone){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        } else {
            Write-Host "Invalid phone number" -ForegroundColor Yellow
        }
        userMenu -SamName $SamName
    } # Update Office Phone #
    if ($choice -eq 8){
        [String]$newPhone = Read-Host "New Mobile Phone # (XXX-XXX-XXXX)"
        if (($newPhone.Length -eq 12) -and ($newPhone.Split("-").Length -eq 3) -and ($newPhone.Split("-")[0].Length -eq 3) -and ($newPhone.Split("-")[1].Length -eq 3) -and ($newPhone.Split("-")[2].Length -eq 4)){
            [String]$choicen = Read-Host "Are you sure you want to update the mobile Phone # $($userItems.MobilePhone) to $($newPhone)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Credential $cred -Identity $SamName -MobilePhone $newPhone
                $user.MobilePhone = $newPhone
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Identity $SamName -Properties MobilePhone).MobilePhone -ne $newPhone){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        } else {
            Write-Host "Invalid phone number" -ForegroundColor Yellow
        }
        userMenu -SamName $SamName
    } # Update Mobile Phone #
    if ($choice -eq 9){
        #$ADComputers = Get-ADComputer -Filter *
        $computers = @()
        $user.msSBSComputerUserAccessOverride | ForEach-Object{ 
            $distinguishedName = ($_.Substring(6,($_.Length - 6))).Trim()
            $ADPC = Get-ADComputer -Filter {DistinguishedName -eq $distinguishedName}
            if ($ADPC.Name.Length -ne 0) {
                $computers += $ADPC
            }
        }
        $computers | ForEach-Object { $_ | Add-Member index (([array]::IndexOf($computers.Name, $_.Name)) + 1) -Force } # Add menu numbers (don't start at 0)
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Clear-Host
        Write-Host "=========Edit Computers========" -ForegroundColor White
        Write-Host "Currently Accessable Computers:" -ForegroundColor Green
        $computers | ForEach-Object {
            Write-Host "$($_.Name)" -ForegroundColor Cyan
        }
        Write-Host "1`tAdd new PC" -ForegroundColor Yellow
        Write-Host "2`tRemove PC" -ForegroundColor Red
        Write-Host "`n"
        [INT]$cchoice = Read-Host ">"
        if ($cchoice -eq 1) {
            $ADComputers = Get-ADComputer -Filter {Enabled -eq $true}
            $potentialPCs = @()
            $potentialPCs = $ADComputers | Where-Object { $computers.Name -notcontains $_.Name } | Sort-Object -Property Name -Unique
            Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
            Clear-Host
            Write-Host "=========Add Computers=========" -ForegroundColor White
            Write-Host "Add computer for $($user.Name)" -ForegroundColor Green
            Write-Host "They will be able to remote in to this PC"
            Write-Host "`n"
            [String]$pcchoice = Read-Host ">"
            $potentialPCs | ForEach-Object { 
                if ($_.Name.Trim() -like $pcchoice) {
                    Write-Host "Adding $pcchoice to $($user.Name)..."
                    $newPC = Get-ADComputer -Filter {Name -eq $pcchoice}
                    [String]$newSBSPC = ($newPC.SID -replace "-", ":").Substring(0,6) + $newPC.DistinguishedName
                    try {
                        Set-ADUser -Credential $cred -Identity $user.SamAccountName -Add @{msSBSComputerUserAccessOverride=$newSBSPC} # Simply add a string to the list
                    } catch { 
                        Write-Host "Oops! Something went wrong." -ForegroundColor Yellow
                        Write-Output $error
                        Start-Sleep -Seconds 3 
                    }
                }
            }
        }
        if ($cchoice -eq 2) {
            Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
            Clear-Host
            Write-Host "=========Remove Computers=========" -ForegroundColor White
            Write-Host "Remove computer from $($user.Name)" -ForegroundColor Green
            $computers | ForEach-Object { Write-Host "$($_.index)`t$($_.Name)" -ForegroundColor Cyan }
            Write-Host "`n"
            [INT]$pcnchoice = Read-Host ">"
            if (($pcnchoice -gt 0) -and ($pcnchoice -le $computers.Count)){
                $pc = $computers[($pcnchoice - 1)]
                $ADPC = Get-ADComputer -Filter {Name -eq $pc.Name}
                ## Get Current List
                [System.Collections.CollectionBase]$sbs = (Get-ADUser -Identity $user.SamAccountName -Properties * ).msSBSComputerUserAccessOverride
                [String]$badEntry = ""
                $sbs | ForEach-Objcet {
                    if ( $_ -like "*$($pc.Name)*"){ 
                        $badEntry = $_
                    }
                } ## Get $badEntry from msSBSComputerUserAccessOverride
                if ($badEntry.Length -gt 1){ $sbs.Remove($badEntry) }
                [String]$choicen = Read-Host "Are you sure that you want to remove $($pc.Name) from $($user.Name)'s Remote Access list? [y/N]"
                if ($choicen.Length -lt 1){$choicen = 'n'}
                if ($choicen.ToLower().Substring(0,1) -eq "y"){
                    try {
                        Set-ADUser -Credential $cred -Identity $user.SamAccountName -Replace @{msSBSComputerUserAccessOverride=$sbs} # For some reason this removes ALL of the msSBSComputerUserAccessOverride entries
                        $sbs | ForEach-Object { 
                            Set-ADUser -Credential $cred -Identity $user.SamAccountName -Add @{msSBSComputerUserAccessOverride=[String]$_}
                        } # Add all of the appropriate entries back in
                    } catch { 
                        Write-Host "Oops! Something went wrong." -ForegroundColor Yellow
                        Write-Output $error
                        Start-Sleep -Seconds 3
                    }
                }
            }
        }
        userMenu -SamName $SamName
    } # Update Remote Computer
    if ($choice -eq 10){
        $allUserGroups = Get-ADPrincipalGroupMembership $user
        $allPrinterGroups = Get-ADGroup -Filter {samAccountName -like "*PrinterUsers"}
        $userPrinterGroups = @()
        $allUserGroups | ForEach-Object {
            $thisGroup = $_
            $allPrinterGroups | ForEach-Object {
                if ($thisGroup -match $_) {
                    $userPrinterGroups += $_
                }
            }
        }
        $potentialPrinterGroups = @()
        if ($userPrinterGroups.Count -gt 0) {
            $allPrinterGroups | ForEach-Object {
                $thisGroup = $_
                $userPrinterGroups | ForEach-Object {
                    if ($thisGroup -notmatch $_){
                        $potentialPrinterGroups += $thisGroup
                    }
                }
            }
        } else { 
            $potentialPrinterGroups = $allPrinterGroups
        }
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Clear-Host
        Write-Host "=========Edit Printer Groups========" -ForegroundColor White
        Write-Host "Current Printer Groups:" -ForegroundColor Green
        $userPrinterGroups | %{ Write-Host "$($_.name)" -ForegroundColor Cyan}
        Write-Host "1`tAdd to group" -ForegroundColor Yellow
        Write-Host "2`tRemove from group" -ForegroundColor Red
        Write-Host "`n"
        [INT]$pchoice = Read-Host ">"
        if ($pchoice -eq 1){
            Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
            Clear-Host
            Write-Host "=========Add to Printer Group=========" -ForegroundColor White
            Write-Host "Add $($user.Name) to a printer group" -ForegroundColor Green
            Write-Host "They will be able to print to the printers in this group"
            $potentialPrinterGroups | ForEach-Object {
                $_.Selection = (([array]::IndexOf($potentialPrinterGroups.Name, $_.Name)) + 1)
            }
            $potentialPrinterGroups | ForEach-Objcet { 
                Write-Host "$($_.Selection)`t$($_.Name)" -ForegroundColor Cyan
            }
            Write-Host "`n"
            [String]$ppchoice = Read-Host ">"
            try {
                $name = ""
                $potentialPrinterGroups | ForEach-Object {
                    if( $_.Selection -match $ppchoice) {
                        $name = $_.Name
                    }
                }
                $targetGroup = Get-ADGroup -Filter {samAccountName -like $name}
                Write-Host "Adding $($user.Name) to $($targetGroup.Name)..." -ForegroundColor White
                Add-ADGroupMember -Credential $cred $targetGroup -Members $user
            } catch {
                Write-Host "Operation failed" -ForegroundColor Red
                Start-Sleep -Seconds 3
            }
        }
        if ($pchoice -eq 2){
            Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
            Clear-Host
            Write-Host "=========Remove from Printer Group=========" -ForegroundColor White
            Write-Host "Remove $($user.Name) from a printer group" -ForegroundColor Green
            Write-Host "They will not be able to print to the printers in this group"
            $userPrinterGroups | ForEach-Object {
                $_.Selection = (([array]::IndexOf($userPrinterGroups.Name, $_.Name)) + 1)
            }
            $userPrinterGroups | ForEach-Object {
                Write-Host "$($_.Selection)`t$($_.Name)" -ForegroundColor Cyan
            }
            Write-Host "`n"
            [String]$ppchoice = Read-Host ">"
            try {
                $name = ""
                $userPrinterGroups | ForEach-Object {
                    if( $_.Selection -match $ppchoice) {
                        $name = $_.Name
                    }
                }
                $targetGroup = Get-ADGroup -Filter {samAccountName -like $name}
                Write-Host "Removing $($user.Name) from $($targetGroup.Name) printer group..." -ForegroundColor White
                Remove-ADGroupMember $targetGroup -Members $user -Credential $cred
            } catch { 
                Write-Host "Operation failed" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
        Start-Sleep -Seconds 1
        userMenu -SamName $SamName
    } # Update Printer Group(s)
    if ($choice -eq 11){
        $origSamName = $SamName
        $SamName = Read-Host 'SamAccountName (e.g. jdoe)'
        try {
            $user = Get-ADUser -Identity $SamName -Properties *
        } catch {
            Write-Host "ERROR! The user `"$SamName`" could not be found" -ForegroundColor Red
            $SamName = $origSamName
            Start-Sleep -Seconds 3
        }
        userMenu -SamName $SamName
    } # Select a different user
}

try {
    $user = Get-ADUser -Identity $SamName -Properties * 
    userMenu -SamName $user.SamAccountName
} catch {
    Write-Host "ERROR! The user `"$SamName`" could not be found" -ForegroundColor Red
}
