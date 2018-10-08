#Execute
###############################################################################################
# Edit Active Directory User Attributes
# You should heavily modify this script before use in a production environment
###############################################################################################
$adminName = $env:adminName
$emailDomain = "acme.com"
$domainName = $env:domainName
$server = $env:LOGONSERVER.Substring(2, ($env:LOGONSERVER).Length - 2)
###############################################################################################
if ($cred) {} else { $cred = Get-Credential $adminName }
Import-Module ActiveDirectory
Clear-Host

function userMenu {
    [CmdletBinding()] Param(
    [Parameter(Position = 0, Mandatory = $False)]
    [String]$SamName = $env:USERNAME
    )
    ########################## User Update Functions #########################
    ### FirstNameEdit # Update FirstName (GiveName)
    function firstNameEdit($user){
        [String]$newName = Read-Host "New First Name"
        if ($newName.length -gt 1){
            Write-Host "You should probably update the email address after this" -ForegroundColor White
            [String]$choicen = Read-Host "Are you sure you want to rename $($userItems.GivenName) to $($newName)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -GivenName $newName
                $user.GivenName = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties GivenName).GivenName -ne $newName){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    }
    
    ### LastNameEdit # Update LastName (Surname)
    function lastNameEdit($user){
        [String]$newName = Read-Host "New Last Name"
        if ($newName.length -gt 1){
            Write-Host "You should probably update the email address after this" -ForegroundColor White
            [String]$choicen = Read-Host "Are you sure you want to rename $($userItems.Surname) to $($newName)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -Surname $newName
                $user.Surname = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties Surname).Surname -ne $newName){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    }

    ### EmailEdit # Automatically update Email (based on FirstName/LastName)
    function emailEdit($user){
        $proxyAddresses = @()
        $setEmail = ($user.GivenName).Substring(0,1) + "." + $user.Surname + $emailDomain
        $proxyAddresses += "SMTP:" + $setEmail
        $proxyAddresses += "smtp:" + ($user.GivenName).Substring(0,1) + $user.Surname + $emailDomain
        Write-Host "This will also change the UPN from $($user.UserPrincipalName) to $setEmail" -ForegroundColor White
        Write-Host "This will change the proxyAddresses to: $proxyAddresses" -ForegroundColor White
        [String]$choicen = Read-Host "Are you sure you want to update $($userItems.Email) to $($setEmail)? [y/N]"
        if ($choicen.Length -lt 1){$choicen = 'n'}
        if ($choicen.ToLower().Substring(0,1) -eq "y"){
            Set-ADUser -Server $server -Credential $cred -Identity $SamName -EmailAddress $setEmail
            Set-ADUser -Server $server -Credential $cred -Identity $SamName -Replace @{proxyAddresses=$proxyAddresses}
            Set-ADUser -Server $server -Credential $cred -Identity $SamName -UserPrincipalName $setEmail
            $user.EmailAddress = $setEmail
            Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
            while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties EmailAddress).EmailAddress -ne $setEmail){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
        }
        userMenu -SamName $SamName
    }

    ### AliasesEdit # Allow selective update of user aliases
    function aliasEdit($user){
        $proxyAddresses = @()
        $user = Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties * 
        $proxyAddresses = $user.proxyAddresses
        $userProxyAddresses = @()
        $index = 0
        $proxyAddresses | %{
            $index = $index + 1
            $proxyProps = @{
                'Selection' = $index;
                'Proxy' = $_
            }
            $thisProxy = New-Object PSObject -Property $proxyProps
            $userProxyAddresses += $thisProxy
        }
        Clear-Host
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Write-Host "=======Edit Proxy Alias======" -ForegroundColor White
        $userProxyAddresses | ForEach-Object{
            Write-Host "$($_.Selection)`t$($_.Proxy)" -ForegroundColor Cyan
        }
        Write-Host "`nSelect `"0`" to add an alias" -ForegroundColor White
        Write-Host "`n"
        [INT]$choicep = Read-Host ">"
        if (($choicep -ge 0) -and ($choicep -le ($userProxyAddresses.Selection).Count)){
            if ($choicep -ne 0){
                $selectedProxy = $userProxyAddresses.Proxy[($choicep - 1)]
                $primaryAlias = $false
                if ($selectedProxy.Substring(0,4) -ceq $selectedProxy.Substring(0,4).ToUpper()){ $primaryAlias = $true }
                Clear-Host
                Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
                Write-Host "=======Edit Proxy Alias======" -ForegroundColor White
                Write-Host -NoNewline "$($selectedProxy)" -ForegroundColor Green
                if ($primaryAlias) { Write-Host -NoNewline "  -Primary Alias`n" -ForegroundColor Cyan } else { Write-Host -NoNewline "`n" }
                Write-Host "1`tToggle Primary Alias" -ForegroundColor Yellow
                Write-Host "2`tAdd alias" -ForegroundColor Green
                Write-Host "3`tRemove alias" -ForegroundColor Red
                Write-Host "`n"
                [INT]$choicea = Read-Host ">"
                if (($choicea -ge 1) -and ($choicea -le 3)){
                    if ($choicea -eq 1){
                        # Toggle Primary Alias
                        $newProxies = @()
                        $AliasMod = $userProxyAddresses[($choicep - 1)].Proxy
                        if ($primaryAlias){
                            # Toggle off primary Alias
                            $AliasMod = [string]$AliasMod.Substring(0,4).ToLower() + $AliasMod.Substring(4,($aliasMod.Length - 4))
                        } else {
                            # Toggle on primary Alias
                            $AliasMod = [string]$AliasMod.Substring(0,4).ToUpper() + $AliasMod.Substring(4,($aliasMod.Length - 4))
                        }
                        $newProxies += $AliasMod
                        $userProxyAddresses | ForEach-object {
                            if (($_.Selection) -ne ($choicep)){
                                $newProxies += [string](($_.Proxy).Substring(0,4).ToLower() + ($_.Proxy).Substring(4,(($_.Proxy).Length - 4)))
                            }
                        }
                        $newProxies | Format-List
                        $updateYN = Read-Host "Is this OK? [y/N]"
                        if (($updateYN.Substring(0,1).ToLower()) -eq "y"){
                            Write-Host "Updating proxies..." -ForegroundColor Magenta
                            Set-ADUser -Server $server -Credential $cred -Identity $SamName -Replace @{proxyAddresses=$newProxies}
                            Start-Sleep -Seconds 3
                        }
                    }
                    if ($choicea -eq 2){
                        # Add alias
                        [String]$newAlias = Read-Host "New alias (without `"smtp:`")"
                        if ($newAlias.Substring(0,5).ToLower() -eq "smtp:"){ $newAlias = $newAlias.Substring(5,($newAlias.Length - 5)) }
                        $newAlias = "smtp:" + $newAlias
                        $newProxies = @()
                        $userProxyAddresses | ForEach-Object {
                            $newProxies += [string]$_.Proxy
                        }
                        $newProxies += $newAlias
                        Write-Host "New Proxies:" -ForegroundColor white
                        $newProxies | Format-List
                        $updateYN = Read-Host "Is this OK? [y/N]"
                        if (($updateYN.Substring(0,1).ToLower()) -eq "y"){
                            Write-Host "Updating proxies..." -ForegroundColor Magenta
                            Set-ADUser -Server $server -Credential $cred -Identity $SamName -Replace @{proxyAddresses=$newProxies}
                            Start-Sleep -Seconds 3
                        }
                    }
                    if ($choicea -eq 3){
                        # Remove alias
                        $newProxies = @()
                        if ($primaryAlias){
                            Write-Host "ERROR! You are trying to delete a primary alias, you must switch primary aliases first." -ForegroundColor Yellow
                            Start-Sleep -Seconds 5
                        } else {
                            $userProxyAddresses | ForEach-Object {
                                if (($_.Selection) -ne ($choicep)){
                                    $newProxies += [string]$_.Proxy
                                }
                            }
                            Write-Host "New Proxies:" -ForegroundColor White
                            $newProxies | Format-List
                            $updateYN = Read-Host "Is this OK? [y/N]"
                            if (($updateYN.Substring(0,1).ToLower()) -eq "y"){
                                Write-Host "Updating proxies..." -ForegroundColor Magenta
                                Set-ADUser -Server $server -Credential $cred -Identity $SamName -Replace @{proxyAddresses=$newProxies}
                                Start-Sleep -Seconds 3
                            }
                        }
                    }
                }
            } else {
                # Add alias
                [String]$newAlias = Read-Host "New alias (without `"smtp:`")"
                if ($newAlias.Substring(0,5).ToLower() -eq "smtp:"){ $newAlias = $newAlias.Substring(5,($newAlias.Length - 5)) }
                $newAlias = "smtp:" + $newAlias
                $newProxies = @()
                $userProxyAddresses | ForEach-Object {
                    $newProxies += [string]$_.Proxy
                }
                $newProxies += [string]$newAlias
                Write-Host "New Proxies:" -ForegroundColor white
                $newProxies | Format-List
                $updateYN = Read-Host "Is this OK? [y/N]"
                if (($updateYN.Substring(0,1).ToLower()) -eq "y"){
                    Write-Host "Updating proxies..." -ForegroundColor Magenta
                    Set-ADUser -Server $server -Credential $cred -Identity $SamName -Replace @{proxyAddresses=$newProxies}
                    Start-Sleep -Seconds 3
                }
            }
        } else {
            userMenu -SamName $SamName
        }
        userMenu -SamName $SamName
    }

    ### CompanyEdit # Update Company (From list of companies)
    function companyEdit($user){
        $companies = Get-ADUser -Filter {(Enabled -eq $true) -and (ObjectClass -eq "user") -and (EmailAddress -like "*.com")} -Properties Company | Select Company | Sort-Object -Property Company -Unique
        $companies | ForEach-Object { $_ | Add-Member Selection (([array]::IndexOf($companies.Company, $_.Company)) + 1) } # Add menu numbers (don't start at 0)
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
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -Company $newName
                $user.Company = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties Company).Company -ne $newName){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    }

    ### DepartmentEdit # Update Department (From list of departments)
    function departmentEdit($user){
        $departments = Get-ADUser -Filter {(Enabled -eq $true) -and (ObjectClass -eq "user") -and (EmailAddress -like "*.com")} -Properties Department | Select Department | Sort-Object -Property Department -Unique
        $departments | ForEach-Object { $_ | Add-Member Selection (([array]::IndexOf($departments.Department, $_.Department)) + 1) } # Add menu numbers (don't start at 0)
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
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -Department $newName
                $user.Department = $newName
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties Department).Department -ne $newName){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        Start-Sleep -Seconds 1
        userMenu -SamName $SamName
    }

    ### TitleEdit # Update Title
    function titleEdit($user){
        [String]$newTitle = Read-Host "New Title"
        if (($newTitle.Length -gt 1)){
            [String]$choicen = Read-Host "Are you sure you want to update $($user.Name)'s Title: `"$($user.Title)`" to `"$($newTitle)`"? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -Title $newTitle
                $user.Title = $newTitle
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties Title).Title -ne $newTitle){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        }
        userMenu -SamName $SamName
    }

    ### OfficePhoneEdit # Update Office Phone #
    function officePhoneEdit($user){
        [String]$newPhone = Read-Host "New Office Phone # (513-867-5309)"
        if (($newPhone.Length -eq 12) -and ($newPhone.Split("-").Length -eq 3) -and ($newPhone.Split("-")[0].Length -eq 3) -and ($newPhone.Split("-")[1].Length -eq 3) -and ($newPhone.Split("-")[2].Length -eq 4)){
            [String]$choicen = Read-Host "Are you sure you want to update the office Phone # $($userItems.OfficePhone) to $($newPhone)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -OfficePhone $newPhone
                $phoneExt = $newPhone.Substring(8,4)
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -Replace @{ipPhone="$phoneExt"}
                $user.OfficePhone = $newPhone
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties OfficePhone).OfficePhone -ne $newPhone){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        } else {
            Write-Host "Invalid phone number" -ForegroundColor Yellow
        }
        userMenu -SamName $SamName
    }

    ### MobilePhoneEdit # Update Mobile Phone #
    function mobilePhoneEdit($user){
        [String]$newPhone = Read-Host "New Mobile Phone # (513-867-5309)"
        if (($newPhone.Length -eq 12) -and ($newPhone.Split("-").Length -eq 3) -and ($newPhone.Split("-")[0].Length -eq 3) -and ($newPhone.Split("-")[1].Length -eq 3) -and ($newPhone.Split("-")[2].Length -eq 4)){
            [String]$choicen = Read-Host "Are you sure you want to update the mobile Phone # $($userItems.MobilePhone) to $($newPhone)? [y/N]"
            if ($choicen.Length -lt 1){$choicen = 'n'}
            if ($choicen.ToLower().Substring(0,1) -eq "y"){
                Set-ADUser -Server $server -Credential $cred -Identity $SamName -MobilePhone $newPhone
                $user.MobilePhone = $newPhone
                Write-Host -NoNewLine "Updating Active Directory record for $($user.Name)" -ForegroundColor Yellow
                while ((Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties MobilePhone).MobilePhone -ne $newPhone){ Start-Sleep 1 ; Write-Host -NoNewline "." -ForegroundColor Yellow }
            }
        } else {
            Write-Host "Invalid phone number" -ForegroundColor Yellow
        }
        userMenu -SamName $SamName
    }

    ### ComputerEdit # Update Remote Computer
    function computerEdit($user){
        $computers = @()
        $user.msSBSComputerUserAccessOverride | ForEach-Object{ 
            $distinguishedName = ($_.Substring(6,($_.Length - 6))).Trim()
            $ADPC = Get-ADComputer -Server $server -Credential $cred -Filter {DistinguishedName -eq $distinguishedName}
            if ($ADPC.Name.Length -ne 0){
                $computers += $ADPC
            }
        }
        $computers | ForEach-Object { 
            $_ | Add-Member index (([array]::IndexOf($computers.Name, $_.Name)) + 1) -Force # Add menu numbers (don't start at 0)
            $localadmin = "Local Permissions Unknown"
            if (Test-Connection -ComputerName $_.Name -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet) {
                if ($_.Name -ne $env:COMPUTERNAME){
                    $testAdmin = (Invoke-Command -ComputerName $_.Name -Credential $cred -ScriptBlock { Get-LocalGroupMember -Group Administrators }) | Where-Object Name -match $user.SamAccountName
                    $testRDP = (Invoke-Command -ComputerName $_.Name -Credential $cred -ScriptBlock { Get-LocalGroupMember -Group "Remote Desktop Users" }) | Where-Object Name -match $user.SamAccountName
                } else {
                    $testAdmin = (Get-LocalGroupMember -Group Administrators) | Where-Object Name -match $user.SamAccountName
                    $testRDP = (Get-LocalGroupMember -Group "Remote Desktop Users") | Where-Object Name -match $user.SamAccountName
                }
                Write-Output $testAdmin
                if ($testAdmin.count -eq 0){
                    $localadmin = "User"
                    if ($testRDP -ne $null){
                        if ($testRDP.ToString().count -gt 0){
                            $localadmin = "RDP User"
                        }
                    }
                } else {
                    $localadmin = "Admin"
                }
            }
            $_ | Add-Member Priv $localadmin -Force
        }
        #PAUSE
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Clear-Host
        Write-Host "=========Edit Computers========" -ForegroundColor White
        Write-Host "Currently Accessable Computers:" -ForegroundColor Green
        $computers | ForEach-Object { Write-Host "$($_.Name) `t $($_.Priv)" -ForegroundColor Cyan }
        Write-Host "1`tAdd new PC" -ForegroundColor Yellow
        Write-Host "2`tRemove PC" -ForegroundColor Red
        Write-Host "3`tChange privileges (User, Admin, RDP User)" -ForegroundColor Magenta
        Write-Host "`n"
        [INT]$cchoice = Read-Host ">"
        switch ($cchoice) {
            1 {
                $ADComputers = Get-ADComputer -Server $server -Credential $cred -Filter {Enabled -eq $true}
                $potentialPCs = @()
                $potentialPCs = $ADComputers | Where-Object { $computers.Name -notcontains $_.Name } | Sort-Object -Property Name -Unique
                Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
                Clear-Host
                Write-Host "=========Add Computers=========" -ForegroundColor White
                Write-Host "Add computer for $($user.Name)" -ForegroundColor Green
                Write-Host "They will be able to remote in to this PC"
                Write-Host "`n"
                [String]$pcchoice = Read-Host ">"
                $potentialPCs | ForEach-Object { if ($_.Name.Trim() -like $pcchoice){
                    Write-Host "Adding $pcchoice to $($user.Name)..."
                    $newPC = Get-ADComputer -Server $server -Credential $cred -Filter {Name -eq $pcchoice}
                    [String]$newSBSPC = ($newPC.SID -replace "-", ":").Substring(0,6) + $newPC.DistinguishedName
                    try {
                        Set-ADUser -Server $server -Credential $cred -Identity $user.SamAccountName -Add @{msSBSComputerUserAccessOverride=$newSBSPC} # Simply add a string to the list
                    } catch { Write-Host "Oops! Something went wrong." -ForegroundColor Yellow ; Write-Output $error; Start-Sleep -Seconds 3 }
                }}
            } # Add new PC
            2 {
                Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
                Clear-Host
                Write-Host "=========Remove Computers=========" -ForegroundColor White
                Write-Host "Remove computer from $($user.Name)" -ForegroundColor Green
                $computers | ForEach-Object { Write-Host "$($_.index)`t$($_.Name)" -ForegroundColor Cyan }
                Write-Host "`n"
                [INT]$pcnchoice = Read-Host ">"
                if (($pcnchoice -gt 0) -and ($pcnchoice -le $computers.Count)){
                    $pc = $computers[($pcnchoice - 1)]
                    $ADPC = Get-ADComputer -Server $server -Credential $cred -Filter {Name -eq $pc.Name}
                    ## Get Current List
                    [System.Collections.CollectionBase]$sbs = (Get-ADUser -Identity $user.SamAccountName -Properties * ).msSBSComputerUserAccessOverride
                    [String]$badEntry = ""
                    $sbs | ForEach-Object{ if ( $_ -like "*$($pc.Name)*"){ $badEntry = $_ }} ## Get $badEntry from msSBSComputerUserAccessOverride
                    if ($badEntry.Length -gt 1){ $sbs.Remove($badEntry) }
                    
                    [String]$choicen = Read-Host "Are you sure that you want to remove $($pc.Name) from $($user.Name)'s Remote Access list? [y/N]"
                    if ($choicen.Length -lt 1){$choicen = 'n'}
                    if ($choicen.ToLower().Substring(0,1) -eq "y"){
                        try {
                            Set-ADUser -Server $server -Credential $cred -Identity $user.SamAccountName -Replace @{msSBSComputerUserAccessOverride=$sbs} # For some reason this removes ALL of the msSBSComputerUserAccessOverride entries
                            $sbs | ForEach-Object { Set-ADUser -Server $server -Credential $cred -Identity $user.SamAccountName -Add @{msSBSComputerUserAccessOverride=[String]$_} } # Add all of the appropriate entries back in
                        } catch { Write-Host "Oops! Something went wrong." -ForegroundColor Yellow ; Write-Output $error; Start-Sleep -Seconds 3 }
                    }
                }
            } # Remove PC
            3 {
                Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
                Clear-Host
                Write-Host "=========Toggle Permissions=========" -ForegroundColor White
                Write-Host "Toggle $($user.Name)'s permissions on:" -ForegroundColor Green
                $computers | ForEach-Object { Write-Host "$($_.index)`t$($_.Name) `t$($_.Priv)" -ForegroundColor Cyan }
                Write-Host "`n"
                [INT]$pcnchoice = Read-Host ">"
                if (($pcnchoice -gt 0) -and ($pcnchoice -le $computers.Count)){
                    $pc = $computers[($pcnchoice - 1)].Name
                    $privChg = "Admin"
                    if (($computers[($pcnchoice - 1)].Priv -eq "Admin")){
                        $privChg = "RDP User"
                    }
                    if (($computers[($pcnchoice - 1)].Priv -eq "RDP User")){
                        $privChg = "User"
                    }
                    if (Test-Connection -ComputerName $pc -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet) {
                        switch ($privChg) {
                            "User" {
                                Write-Host "Removing Admin priviledges for $($user.Name) on $pc..." -ForegroundColor Yellow
                                $localAccountsBlockFunction = [ScriptBlock]::Create({
                                    function removeAdmin ($Username){
                                        Remove-LocalGroupMember -Group Administrators -Member $Username
                                        Remove-LocalGroupMember -Group "Remote Desktop Users" -Member $Username
                                    }
                                })
                                $localAccountsBlock = [ScriptBlock]::Create($localAccountsBlockFunction.ToString() + "removeAdmin -Username '" + "$domainName" + '\' + $SamName + "'") 
                                Invoke-Command -ComputerName $pc -Credential $cred -ScriptBlock $localAccountsBlock
                            }
                            "RDP User" {
                                Write-Host "Granting RDP User priviledges for $($user.Name) on $pc..." -ForegroundColor Yellow
                                $localAccountsBlockFunction = [ScriptBlock]::Create({
                                    function restrictToRDP ($Username){
                                        Remove-LocalGroupMember -Group Administrators -Member $Username
                                        Add-LocalGroupMember -Group "Remote Desktop Users" -Member $Username
                                    }
                                })
                                $localAccountsBlock = [ScriptBlock]::Create($localAccountsBlockFunction.ToString() + "restrictToRDP -Username '" + "$domainName" + '\' + $SamName + "'")
                                Invoke-Command -ComputerName $pc -Credential $cred -ScriptBlock $localAccountsBlock
                            }
                            "Admin" {
                                Write-Host "Granting Admin priviledges for $($user.Name) on $pc..." -ForegroundColor Yellow
                                $localAccountsBlockFunction = [ScriptBlock]::Create({
                                    function addToAdmin ($Username){
                                        Add-LocalGroupMember -Group Administrators -Member $Username
                                    }
                                })
                                $localAccountsBlock = [ScriptBlock]::Create($localAccountsBlockFunction.ToString() + "addToAdmin -Username '" + "$domainName" + '\' + $SamName + "'")
                                Invoke-Command -ComputerName $pc -Credential $cred -ScriptBlock $localAccountsBlock
                            }
                        }
                    } else {
                        Write-Host "Unable to change local permissions on unreachable PC ($pc)" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            } # Change privileges on PC
        }
        Start-Sleep -Seconds 1
        userMenu -SamName $SamName
    }

    ### PrintersEdit # Update Printer Group(s)
    function printersEdit($user){
        $allUserGroups = Get-ADPrincipalGroupMembership -Server $server -Credential $cred $user
        $allPrinterGroups = Get-ADGroup -Server $server -Credential $cred -Filter {samAccountName -like "*PrinterUsers"}
        $userPrinterGroups = @()
        $allUserGroups | ForEach-Object {
            $thisGroup = $_
            $allPrinterGroups | ForEach-Object {
                if ($thisGroup -match $_){
                    $userPrinterGroups += $_
                }
            }
        }
        $potentialPrinterGroups = @()
        if ($userPrinterGroups.Count -gt 0){
            $allPrinterGroups | ForEach-Object {
                $thisPotentialGroup = $_
                $found = $false
                $userPrinterGroups | ForEach-Object {
                    if ($thisPotentialGroup -match $_){
                        $found = $true
                    }
                }
                if (!($found)){
                    $potentialPrinterGroups += $thisPotentialGroup
                }
            }
        } else { 
            $potentialPrinterGroups = $allPrinterGroups
        }
        Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
        Clear-Host
        Write-Host "=========Edit Printer Groups========" -ForegroundColor White
        Write-Host "Current Printer Groups:" -ForegroundColor Green
        $userPrinterGroups | ForEach-Object { Write-Host "$($_.name)" -ForegroundColor Cyan}
        Write-Host "1`tAdd to group" -ForegroundColor Yellow
        Write-Host "2`tRemove from group" -ForegroundColor Red
        Write-Host "`n"
        [INT]$pchoice = Read-Host ">"
        switch ($pchoice) {
            1 {
                Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
                Clear-Host
                Write-Host "=========Add to Printer Group=========" -ForegroundColor White
                Write-Host "Add $($user.Name) to a printer group" -ForegroundColor Green
                Write-Host "They will be able to print to the printers in this group"
                $potentialPrinterGroups | ForEach-Object { $_.Selection = (([array]::IndexOf($potentialPrinterGroups.Name, $_.Name)) + 1) }
                $potentialPrinterGroups | ForEach-Object { Write-Host "$($_.Selection)`t$($_.Name)" -ForegroundColor Cyan }
                Write-Host "`n"
                [String]$ppchoice = Read-Host ">"
                try {
                    $name = ""
                    $potentialPrinterGroups | ForEach-Object { if( $_.Selection -match $ppchoice){ $name = $_.Name }}
                    $targetGroup = Get-ADGroup -Server $server -Filter {samAccountName -like $name}
                    Write-Host "Adding $($user.Name) to $($targetGroup.Name)..." -ForegroundColor White
                    Add-ADGroupMember -Server $server -Credential $cred $targetGroup -Members $user
                } catch { Write-Host "Operation failed" -ForegroundColor Red ; Start-Sleep -Seconds 3 }
            }
            2 {
                Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
                Clear-Host
                Write-Host "=========Remove from Printer Group=========" -ForegroundColor White
                Write-Host "Remove $($user.Name) from a printer group" -ForegroundColor Green
                Write-Host "They will not be able to print to the printers in this group"
                $userPrinterGroups | ForEach-Object { $_.Selection = (([array]::IndexOf($userPrinterGroups.Name, $_.Name)) + 1) }
                $userPrinterGroups | ForEach-Object { Write-Host "$($_.Selection)`t$($_.Name)" -ForegroundColor Cyan}
                Write-Host "`n"
                [String]$ppchoice = Read-Host ">"
                try {
                    $name = ""
                    $userPrinterGroups | ForEach-Object { if( $_.Selection -match $ppchoice){ $name = $_.Name }}
                    $targetGroup = Get-ADGroup -Server $server -Filter {samAccountName -like $name}
                    Write-Host "Removing $($user.Name) from $($targetGroup.Name) printer group..." -ForegroundColor White
                    Remove-ADGroupMember -Server $server $targetGroup -Members $user -Credential $cred
                } catch { Write-Host "Operation failed" -ForegroundColor Red ; Start-Sleep -Seconds 1 }
            }
        }
        Start-Sleep -Seconds 1
        userMenu -SamName $SamName
    }

    ########################## Active Directory Menu #########################
    $user = @()
    $user = Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties * 
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
        'Computers'=$computers;
    }
    $userItems = New-Object PSObject -Property $properties
    Clear-Host
    Write-Host "=======Active Directory User Edit======" -ForegroundColor White
    Write-Host "1`tFirst Name`t`t$($userItems.GivenName)" -ForegroundColor Cyan
    Write-Host "2`tLast Name`t`t$($userItems.Surname)" -ForegroundColor Cyan
    Write-Host "3`tEmail`t`t`t$($userItems.Email)" -ForegroundColor Cyan
    Write-Host "4`tAliases" -ForegroundColor Cyan
    Write-Host "5`tCompany`t`t`t$($userItems.Company)" -ForegroundColor Cyan
    Write-Host "6`tDepartment`t`t$($userItems.Department)" -ForegroundColor Cyan
    Write-Host "7`tTitle`t`t`t$($user.Title)" -ForegroundColor Cyan
    Write-Host "8`tOffice Phone #`t$($userItems.OfficePhone)" -ForegroundColor Cyan
    Write-Host "9`tMobile Phone #`t$($userItems.MobilePhone)" -ForegroundColor Cyan
    Write-Host "10`tComputer Access" -ForegroundColor Cyan
    Write-Host "11`tPrinter Group(s)" -ForegroundColor Cyan
    Write-Host "12`tEdit another user" -ForegroundColor Yellow
    Write-Host "`n"
    [INT]$choice = Read-Host ">"
    switch ( $choice ){
        1 { firstNameEdit -user $user }
        2 { lastNameEdit -user $user }
        3 { emailEdit -user $user }
        4 { aliasEdit -user $user }
        5 { companyEdit -user $user }
        6 { departmentEdit -user $user }
        7 { titleEdit -user $user }
        8 { officePhoneEdit -user $user }
        9 { mobilePhoneEdit -user $user }
        10 { computerEdit -user $user }
        11 { printersEdit -user $user }
        12 { 
            $origSamName = $SamName
            $SamName = Read-Host 'SamAccountName (e.g. jdoe)'
            try {
                $user = Get-ADUser -Server $server -Credential $cred -Identity $SamName -Properties *
            } catch {
                Write-Host "ERROR! The user `"$SamName`" could not be found" -ForegroundColor Red
                $SamName = $origSamName
                Start-Sleep -Seconds 3
            }
            userMenu -SamName $SamName
         }
    }
}

############################### Call userMenu function ###################################
Write-Host "Edit a user's Active Directory attributes" -ForegroundColor Green
$SamName = Read-Host 'SamAccountName (e.g. dbrenner)'
$error.clear()
userMenu -SamName $SamName