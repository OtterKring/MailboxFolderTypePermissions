function Remove-MailboxFolderTypePermission {
    param (
        [Parameter(Mandatory,Position=0,ValueFromPipelineByPropertyName)]
        $UserPrincipalName,
        [Parameter(Mandatory,Position=1)]
        $User,
        [Parameter(Mandatory)]
        [ValidateSet("Calendar","CalendarLogging","Contacts","ConversationActions","DeletedItems","Drafts","ExternalContacts","Files","GalContacts","Inbox","Journal","JunkEmail","Notes","Outbox","QuickContacts","RecipientCache","RecoverableItemsDeletions","RecoverableItemsPurges","RecoverableItemsRoot","RecoverableItemsVersions","Root","SentItems","Tasks","User Created","YammerFeeds","YammerInbound","YammerOutbound","YammerRoot")]
        $Type
    )

    begin {

        $UserObj = Get-Recipient $User -ErrorAction SilentlyContinue

    }

    process {

        $Mbx = Get-Mailbox $UserPrincipalName
        if ($Mbx) {
            $FolderIDs = (Get-MailboxFolderStatistics $Mbx.UserPrincipalName | Where-Object FolderType -eq $Type).FolderID
            foreach ($id in $FolderIDs) {
                Remove-MailboxFolderPermission ($Mbx.PrimarySmtpAddress + ":" + $id) -User $UserObj.PrimarySmtpAddress
            }            
        }

    }

    end {
        # nothing yet
    }

}