#Execute
######################################################################
# Connect to 365 online so that we run
# commands in the 365 environment
# 
# This script wraps some common Office 365 tasks within a menu
# structure
#
# TODO: Error Corrections need massive improvement
# Need to actually check values before attempting change, if a change 
# is needed.
######################################################################
## Install Azure Active Directory Module for Windows PowerShell first
# http://go.microsoft.com/fwlink/p/?linkid=236297
# https://www.microsoft.com/en-us/download/details.aspx?id=28177
######################################################################
$ConnectionUri = "https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS"
$inPlaceHoldMailboxIdentity = "In-Place Hold for archiving" # Mailbox for journaling
######################################################################

if ($365UserCredentials) {} else {
    Write-Host "Please enter your Office 365 e-mail address and password." -ForegroundColor White
    $365UserCredentials = Get-Credential 
}
function Connect365 {
    Import-Module MSOnline
    ## Create 365 Session ##
    $365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $365UserCredentials -Authentication Basic -AllowRedirection
    Import-PSSession $365Session
    Connect-MSolService -Credential $365UserCredentials
    $365connected = $true
}

#Determine if connected in this tab
if ($365connected) {
    if ($365connected -eq $false){
        Connect365
        $365connected = $true
    }
} else {
    Connect365
    $365connected = $true
}
$users = Get-Mailbox -RecipientTypeDetails UserMailbox

Function listUsers {
    Clear-Host
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Write-Host "Getting 365 Users..." -ForegroundColor White
    Write-Host "`n"
    $users = Get-Mailbox -RecipientTypeDetails UserMailbox 
    $users | Select-Object Alias, SamAccountName, Identity, WindowsEmailAddress, IsMailboxEnabled | Out-GridView
}

Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "JPG (*.jpg)| *.jpg"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

Function UploadPicture{
    Clear-Host
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Write-Host "Upload profile picture for the following user" -ForegroundColor White
    $userName = Read-Host "User Identity"
    Write-Host "Select the user's photo (200px X 200px)"
    $userPhoto = Get-FileName $env:PSMgmtRoot
    if ($userPhoto -ne $null){
        Write-Host "Uploading photo ($userPhoto) for $userName..." -ForegroundColor Yellow
        $pictData = ([Byte[]] $(Get-Content -Path $userPhoto -Encoding Byte -ReadCount 0))
        Set-UserPhoto $userName -PictureData $pictData -Preview -Confirm:$false
        Set-UserPhoto $userName -Save -Confirm:$false
        Write-Host "...Finished Uploading" -ForegroundColor White
    }
    PAUSE
}

Function GrantUserRights {
    Clear-Host
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Write-Host "Grant the following user Edit rights on all other user's calendars" -ForegroundColor White
    $userName = Read-Host "User Identity"
    $users | For-EachObject { 
        Try { 
            Write-Host "Attempting to grant $userName rights on $_`'s Calendar" -ForegroundColor Yellow
            Set-MailboxFolderPermission $_":\Calendar" -User $userName -AccessRights Editor -ErrorAction SilentlyContinue
            Add-MailboxFolderPermission $_":\Calendar" -User $userName -AccessRights Editor -ErrorAction SilentlyContinue
        } Catch {}
    }
}

Function ListUserRights {
    Clear-Host
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Write-Host "Mailbox Rights" -ForegroundColor White
    Write-Host "`n"
    $users | ForEach-Object {
        Write-Host "$_" -ForegroundColor Yellow
        $permissions = Get-MailboxFolderPermission $_":\Calendar"
        $permissions | Select-Object User, AccessRights | Format-Table
    }
    Write-Host "`n"
    Pause
}

function MobileSearch ($MailboxSearch) {
    if ( $MailboxSearch -eq "" ) {
        $Mobiles = Get-MobileDevice
    } else {
        $Mobiles = Get-MobileDevice -Mailbox $MailboxSearch
    }
    Return $Mobiles
}

function MobileMenu ($Mobiles) {
    $menu= @()
    $index = 1
    #$Mobiles | Sort-Object -Property UserDisplayName | ForEach-Object {
    $Mobiles | Sort-Object -Property Identity | ForEach-Object {
        $properties = @{
            'Selection'=$index;
            'UserDisplayName'=$_.UserDisplayName
            #'UserDisplayName'=$_.Id
            'DeviceModel'=$_.DeviceModel;
            'FirstSync'=$_.FirstSyncTime;
            'WhenChangedUTC'=$_.WhenChangedUTC
            'Identity'=$_.Guid
        }
        $menuItems = New-Object PSObject -Property $properties
        $menu += $menuItems
        ++$index
    }

    [console]::ResetColor()
    $psISE.Options.RestoreDefaultTokenColors()
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Clear-Host
    Write-Host "=======Mobiles======" -ForegroundColor White
    $menu | ForEach-Object {
        $index = $_.Selection
        $UserDisplayName = $_.UserDisplayName
        #$UserDisplayName = $_.Id
        $DeviceModel = $_.DeviceModel
        $FirstSync = $_.FirstSync
        $LastChange = $_.WhenChangedUTC
        Write-Host "$index`t$UserDisplayName`t$DeviceModel`t$FirstSync`t$LastChange" -ForegroundColor Cyan
    }
    if ($menu.Count -ge 10){
        #$menu | Select-Object Selection, UserDisplayName, DeviceModel, FirstSync, WhenChangedUTC | Out-GridView
    }

    Write-Host "`n"
    $choicem = Read-Host ">"
    if (($choicem -like "") -or ($choicem.ToLower() -like "exit")) { return }

    try { 
        $chosen = $menu[$choicem - 1]
    } catch { }

    ## Phone Identity
    $identity = ($chosen.Identity).Guid

    ## Do Something
    [console]::ResetColor()
    $psISE.Options.RestoreDefaultTokenColors()
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Clear-Host
    Get-MobileDevice -Identity $identity | Select-Object UserDisplayName, Id, FriendlyName, DeviceType, DeviceModel, DeviceOS, ClientVersion, WhenChanged
    Write-Host "========Action========" -ForegroundColor White
    Write-Host "1`tGet ALL Device Info" -ForegroundColor Cyan
    Write-Host "2`tRemove Device from 365" -ForegroundColor Yellow
    Write-Host "3`tWipe Device" -ForegroundColor Red
    Write-Host "`n"
    $ActionChoice = Read-Host ">"
    if ($ActionChoice -eq 1 ) {
        Get-MobileDevice -Identity $identity
        pause
    } elseif ($ActionChoice -eq 2 ) {
        Remove-MobileDevice -Identity $identity
        #$menutemp = @()
        #$menutemp = $Mobiles | Where-Object { $_.Guid.Guid -ne $identity }
        #$Mobiles = $menutemp
        #MobileMenu -Mobiles $Mobiles
    } elseif ($ActionChoice -eq 3 ) {
        Clear-MobileDevice -Identity $identity
        #$menutemp = @()
        #$menutemp = $Mobiles | Where-Object { $_.Guid.Guid -ne $identity }
        #$Mobiles = $menutemp
        #MobileMenu -Mobiles $Mobiles
    } else { return }
}

function MobileMain {
    <#
    if ($Mobiles.Count -le 0){
        Clear-Host
        $MailboxSearch = Read-Host "Mailbox Username (blank for all)"
        $Mobiles = MobileSearch -MailboxSearch $MailboxSearch
    }
    Clear-Host
    MobileMenu -Mobiles $Mobiles
    #>
    $newSearch = Read-Host "New Search [N/y/exit]"
    if ( $newSearch.ToLower() -eq "y" ){
        Clear-Host
        $MailboxSearch = Read-Host "Mailbox Username (blank for all)"
        $Mobiles = MobileSearch -MailboxSearch $MailboxSearch
        MobileMenu -Mobiles $Mobiles
    } elseif ( $newSearch.ToLower() -like "exit" ) {
        return
    } elseif ( ($newSearch.ToLower() -like "n" ) -or ($newSearch.ToLower() -like "") ){
        MobileMenu -Mobiles $Mobiles
    }
    MobileMain
}

function RoomMenu ($choice) {
    function EditRoom {
        $rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox
        $menu= @()
        $index = 1
        $rooms | Sort-Object -Property Name | ForEach-Object {
            $properties = @{
                'Selection'=$index;
                'DisplayName'=$_.DisplayName
                'Alias'=$_.Alias;
                'Identity'=$_.Identity;
                'Created'=$_.WhenMailboxCreated;
                'Capacity'=$_.ResourceCapacity
            }
            $menuItems = New-Object PSObject -Property $properties
            $menu += $menuItems
            ++$index
        }
        Write-Host "=======Edit Room======" -ForegroundColor White
        $menu | ForEach-Object {
            $index = $_.Selection
            $DisplayName = $_.DisplayName
            $Capacity = $_.Capacity
            Write-Host "$index`t$DisplayName`tCapacity: $Capacity" -ForegroundColor Cyan
        }
        Write-Host "`n"
        [INT]$choicem = Read-Host ">"
        try {
            $room = $rooms | Where-Object Identity -eq (($menu[$choicem - 1]).Identity)
            $roomTimes = Get-MailboxCalendarConfiguration -Identity $room.Identity | Select-Object WorkingHoursStartTime, WorkingHoursEndTime
            $startTime = $roomTimes.WorkingHoursStartTime
            $endTime = $roomTimes.WorkingHoursEndTime
            $displayName = $room.DisplayName
            $capacity = $room.Capacity
            Write-Host "$displayName `t $capacity `t $startTime `t $endTime" -ForegroundColor Green
        } catch { return }
        Write-Host "=======Edit Room======" -ForegroundColor White
        Write-Host "1`tAssign to Room List" -ForegroundColor Cyan
        Write-Host "2`tChange Capacity" -ForegroundColor Cyan
        Write-Host "3`tChange Hours" -ForegroundColor Cyan
        Write-Host "4`tChange Name" -ForegroundColor Cyan
        Write-Host "`n"
        [INT]$choicen = Read-Host ">"
        if ($choicn -eq 1){ ## Assign to Room List
            $roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList
            $menuDistGroup = @()
            $index = 1
            $roomLists | Sort-Object -Property Name | ForEach-Object {
                $properties = @{
                    'Selection'=$index;
                    'Name'=$_.Name;
                    'Identity'=$_.Identity;
                }
                $menuDistItems = New-Object PSObject -Property $properties
                $menuDistGroup += $menuDistItems
                ++$index
            }
            Write-Host "$room.DisplayName" -ForegroundColor Green
            Write-Host "=======Assign to Room List======" -ForegroundColor White
            $menuDistGroup | ForEach-Object {
                $index = $_.Selection
                $Name = $_.Name
                Write-Host "$index`t$Name" -ForegroundColor Cyan
            }
            Write-Host "`n"
            [INT]$choiceg = Read-Host ">"
            $group = $roomLists | Where-Object Identity -eq (($menuDistGroup[$choiceg - 1]).Identity)
            Add-DistributionGroupMember -Identity $group.Identity -Member $room.Identity

        }
        elseif ($choicen -eq 2){ ## Change room capacity
            Write-Host "=======Edit Room  Capacity======" -ForegroundColor White
            Write-Host "Current Capacity: $room.ResourceCapacity" -ForegroundColor Green
            [INT]$choiceo = Read-Host "New Capacity"
            if ($choiceo -eq ""){ return } # No Change
            if ($choiceo -eq 0) { Set-MailboxCalendarConfiguration -Identity $room.Identity -ResourceCapacity "" }
            if ($choiceo -gt 0) { Set-MailboxCalendarConfiguration -Identity $room.Identity -ResourceCapacity $choiceo }
            return
        }
        elseif ($choicen -eq 3){ ## Change Hours
            Write-Host "=======Edit Room Hours======" -ForegroundColor White
            Write-Host "Start time:`t$startTime" -ForegroundColor Green
            Write-Host "End Time:`t$endTime" -ForegroundColor Green
            $newStart = Read-Host "New Start Time"
            $newEnd = Read-Host "New End Time"
            Set-MailboxCalendarConfiguration -Identity $room.Identity -WorkingHoursStartTime $newStart -WorkingHoursEndTime $newEnd -WorkingHoursTimeZone "Eastern Standard Time"
        }
        elseif ($choicen -eq 4){ ## Change Room Name
            Write-Host "=======Edit Room Name======" -ForegroundColor White
            $roomIdentity = $room.Identity
            $roomDisplayName = $room.DisplayName
            #$roomOffice = $room.Office
            Write-Host "Name:         $roomIdentity" -ForegroundColor Green
            Write-Host "Display Name: $roomDisplayName" -ForegroundColor Green
            #Write-Host "Location:     $roomOffice" -ForegroundColor Green
            $newName = Read-Host "New Room Name"
            $newDescription = Read-Host "New Room Display Name"
            #$newOffice = Read-Host "New Room Location"
            if (($newName.Length -gt 0) -and ($newDescription.Length -gt 0)){
                Set-Mailbox -Identity $room.Identity -Name "$newName"
                Set-Mailbox -Identity $room.Identity -DisplayName "$newDescription"
                #Set-Mailbox -Identity $room.Identity -office "$newOffice"
            }
        }
        else {
            return
        }
        return
    }
    function CreateRoom {
        Write-Host "=======Create Room======" -ForegroundColor White
        $confIdent = Read-Host "Identity (e.g. ConfRoomA"
        $displayName = Read-Host "Display Name (e.g. Conference Room A)"
        New-Mailbox -Alias $confIdent -Name $confIdent -DisplayName $displayName -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String 'StupidPassword11!' -AsPlainText -Force)
        Write-Output "Please wait while Office 365 processes the new calendar mailbox request..."
        Start-Sleep -s 30
        Set-CalendarProcessing -Identity $confIdent -DeleteSubject $false -DeleteComments $false -AddOrganizerToSubject $false
        Set-CalendarProcessing -Identity $confIdent -TentativePendingApproval $false -AllRequestOutOfPolicy $false
        Set-MailboxCalendarConfiguration -Identity $confIdent -WorkingHoursStartTime 07:30:00 -WorkingHoursEndTime 19:00:00
        return
    }
    function CreateRoomList {
        Write-Host "=======Create Room List======" -ForegroundColor White
        $newRoomList = Read-Host "New Room List Name"
        New-DistributionGroup -Name $newRoomList -RoomList
        return
    }
    
    [console]::ResetColor()
    $psISE.Options.RestoreDefaultTokenColors()
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Clear-Host
    if ($choice -eq 1){ EditRoom }
    if ($choice -eq 2){ CreateRoom }
    if ($choice -eq 3){ CreateRoomList }
    return
}

function InPlaceHold {
    function Set-InPlaceHold($user) {
        Write-Host "Getting current Hold Mailboxes..." -ForegroundColor Yellow
        $InPlaceHoldMailboxes = (Get-MailboxSearch -Identity $inPlaceHoldMailboxIdentity).sourceMailboxes
        $holdMailboxes += $InPlaceHoldMailboxes | ForEach-Object { Get-Mailbox $_ }
        Write-Host "Adding $user to the list..." -ForegroundColor Yellow
        $mailbox = Get-Mailbox -Identity $user
        $holdMailboxes += $mailbox
        $holdMailboxes2 = $holdMailboxes.GUID.GUID
        Write-Host "Getting Soft Deleted Maibox list..." -ForegroundColor Yellow
        $softDeleted = Get-Mailbox -SoftDeletedMailbox # Double-check to avoid duplicates
        $holdMailboxes2 += $softDeleted.ExchangeGuid.GUID # Add Soft Deleted Mailboxes back in
        Write-Host "Setting In-Place Hold to the full list..." -ForegroundColor Cyan
        Set-MailboxSearch -Identity $inPlaceHoldMailboxIdentity -SourceMailboxes $holdMailboxes2 -InPlaceHoldEnabled $true # Re-Do InPlaceHold
    }
    Clear-Host
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Write-Host "Add a user to InPlaceHold for journaling" -ForegroundColor White
    $userName = Read-Host "User Identity"
    Set-InPlaceHold -user $userName
    return
}

function RoomMain {
    Write-Host "========Room Action========" -ForegroundColor White
    Write-Host "1`tEdit Room" -ForegroundColor Cyan
    Write-Host "2`tCreate Room" -ForegroundColor Yellow
    Write-Host "3`tCreate Room List" -ForegroundColor Red
    Write-Host "`n"
    [INT]$ActionChoice = Read-Host ">"
    if (!(($ActionChoice -gt 3) -or ($ActionChoice -lt 1))){
        RoomMenu -choice $ActionChoice
    } else {
        return
    }
    RoomMain
}

function 365menu {
    [console]::ResetColor()
    $psISE.Options.RestoreDefaultTokenColors()
    Start-Sleep -Milliseconds 200 ## Sleep to avoid color jumble
    Clear-Host
    Write-Host "==============365 Admin=============" -ForegroundColor White
    Write-Host "1`tList Users" -ForegroundColor Cyan
    Write-Host "2`tUpload User Photo" -ForegroundColor Cyan
    Write-Host "3`tGive User Edit Rights" -ForegroundColor Cyan
    Write-Host "4`tList User Mailbox rights" -ForegroundColor Cyan
    Write-Host "5`tManage Mobile Devices" -ForegroundColor Cyan
    Write-Host "6`tManage Room Calendars" -ForegroundColor Cyan
    Write-Host "7`tAdd user to Journaling" -ForegroundColor Cyan
    Write-Host "8`tExit" -ForegroundColor Yellow

    Write-Host "`n"
    [INT]$choice = Read-Host ">"

    if ($choice -eq 1){
        listUsers
        365menu
    }
    if ($choice -eq 2){
        UploadPicture
        365menu
    }
    if ($choice -eq 3){
        GrantUserRights
        365menu
    }
    if ($choice -eq 4){
        ListUserRights
        365menu
    }
    if ($choice -eq 5){
        MobileMain
        365menu
    }
    if ($choice -eq 6){
        RoomMain
        365menu
    }
    if ($choice -eq 7){
        InPlaceHold
        365menu
    }
}

365menu