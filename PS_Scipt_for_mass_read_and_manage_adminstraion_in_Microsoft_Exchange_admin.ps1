Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

#define the shared mailbox's
# $users = @("user1@domain.com","user2@domain.com")
# $sharedMailboxes = @(mailbox1@domain.com, mailbox2@domain.com)


#EXCHANGE ONLINE: ADDING OR REMOVING MULTIPLE USERS ACCESS TO MULTIPLE MAILBOXES.
#Import the exchange online powershell module
Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline #-Credential $AdminCredentials
$users =@()
$user = Read-Host "Please enter in the user emails with a comma and space (e.g: user1@domain.com, user2@domain.com)"
Write-Host "-------------------------------------------------------------------------------------------------------"
$user = $user -split ", "
Foreach ($name in $user) {
    $user_temp = "$name"
    $users += $user_temp
}
$sharedMailboxes = @()
$user = Read-Host "please enter the Mailbox's email address with a comma and space (e.g: mailbox1@domain.com, mailbox2@domain.com)"
Write-Host "------------------------------------------------------------------------------------------------------"
$user = $user -split ", "
Foreach ($name in $user) {
    $user_temp = "$name"
    $sharedMailboxes += $user_temp
}

#testing if users and mailboxes are entered in correctly 
Foreach ($mailbox in $sharedMailboxes) {
    Write-host "`nADD/REMOVE TO/FROM -$mailbox-: "
    foreach($user in $users) {
        Write-Host "User: $user"
    }
}
Foreach ($user in $users) {
    $mailbox = Get-Recipient -Filter "EmailAddresses -eq '$user'"
    if ($mailbox) {
        continue
    } else {
        Write-host "User with email $user is not found"
        $users = $users -ne $user
    }
}
Write-Host "------------------------------------------------------------------------------------------------------"
Foreach ($mailbox in $sharedMailboxes) {
    # $mailbox = Get-mailbox -Resultsize Unlimited | Where-Object {$_.PrimarySmtpAddress -eq $user -or $_.EmailAddress -contains $emailAddress}
    $mailbox_found = Get-Mailbox -Filter "PrimarySmtpAddress -eq '$mailbox'" -ErrorAction SilentlyContinue
    if ($mailbox_found) {
        continue
    } else {
        Write-Host "Mailbox with the following email was not found $mailbox"
        $sharedMailboxes = $sharedMailboxes -ne $mailbox
    }
}
$proceed = Read-Host "Would you like to proceed? Yes or no"

if ($proceed.ToLower() -eq "yes") {
    While ($true) {
    $add_or_remove = Read-Host "Add or remove users from group? Type: Add or Remove"
    if ($add_or_remove.ToLower() -eq "add") {
        #Add Delegates to the shared mailboxes
        $add_sendas = Read-Host "Would you like to also grant SendAs permissions? Type: Yes or No"
        foreach ($mailbox in $sharedMailboxes) {
            foreach($user in $users) {
                Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All > $null
                if ($add_sendas.ToLower() -eq "yes") {
                Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false #-Confirm:false removes then need to ask confirmation
                }
                $currentDateTime = Get-Date
                Write-Host "$user has been granted Read and manage (Full Access) to $mailbox at $currentDateTime"
            }
        }
        Break
    } elseif ($add_or_remove.Tolower() -eq "remove") {
        foreach ($mailbox in $sharedMailboxes) {
            foreach($user in $users) {
                Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false
                Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false 
                $currentDateTime = Get-Date
                Write-Host "$user's Read and manage (Full Access) to $mailbox has been removed at $currentDateTime"
            }
        }
        Break
    } else {
        Write-host "Invalid answer to Add or remove users from group"
    }
    }
} else {
    Write-host "Script cancelled"
}