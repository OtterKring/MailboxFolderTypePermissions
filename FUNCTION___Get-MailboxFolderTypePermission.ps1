function Get-MailboxFolderTypePermission {
    param (
        [Parameter(Mandatory,Position=0,ValueFromPipelineByPropertyName)]
        $UserPrincipalName,
        [Parameter(Mandatory,Position=1)]
        [ValidateSet("Calendar","CalendarLogging","Contacts","ConversationActions","DeletedItems","Drafts","ExternalContacts","Files","GalContacts","Inbox","Journal","JunkEmail","Notes","Outbox","QuickContacts","RecipientCache","RecoverableItemsDeletions","RecoverableItemsPurges","RecoverableItemsRoot","RecoverableItemsVersions","Root","SentItems","Tasks","User Created","YammerFeeds","YammerInbound","YammerOutbound","YammerRoot")]
        $Type
    )

    begin {
        # nothing yet
    }

    process {

        $Mbx = Get-Mailbox $UserPrincipalName
        if ($Mbx) {
            $FolderIDs = (Get-MailboxFolderStatistics $Mbx.UserPrincipalName | Where-Object FolderType -eq $Type).FolderID
            foreach ($id in $FolderIDs) {
                Get-MailboxFolderPermission ($Mbx.PrimarySmtpAddress + ":" + $id) `
                | Select-Object @{Name='PrimarySmtpAddress';E={$Mbx.PrimarySmtpAddress}},FolderName,User,AccessRights
            }
        }

    }

    end {
        # nothing yet
    }

}