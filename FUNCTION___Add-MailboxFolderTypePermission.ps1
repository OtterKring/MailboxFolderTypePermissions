function Add-MailboxFolderTypePermission {
    param (
        [Parameter(Mandatory,Position=0,ValueFromPipelineByPropertyName)]
        $UserPrincipalName,
        [Parameter(Mandatory,Position=1)]
        $User,
        [Parameter(Mandatory,Position=2)]
        [ValidateSet('CreateItems','CreateSubfolders','DeleteAllItems','DeleteOwnedItems','EditAllItems','EditOwnedItems','FolderContact','FolderOwner','FolderVisible','ReadItems','Author','Contributor','Editor','None','NonEditingAuthor','Owner','PublishingEditor','PublishingAuthor','Reviewer')]
        $AccessRights,
        [Parameter(Mandatory)]
        [ValidateSet("Calendar","CalendarLogging","Contacts","ConversationActions","DeletedItems","Drafts","ExternalContacts","Files","GalContacts","Inbox","Journal","JunkEmail","Notes","Outbox","QuickContacts","RecipientCache","RecoverableItemsDeletions","RecoverableItemsPurges","RecoverableItemsRoot","RecoverableItemsVersions","Root","SentItems","Tasks","User Created","YammerFeeds","YammerInbound","YammerOutbound","YammerRoot")]
        $Type,
        [Parameter()]
        [switch]$PassThru
    )

    begin {

        $UserObj = Get-Recipient $User -ErrorAction SilentlyContinue
        if ($UserObj) {
            if ($UserObj.RecipientType -like "*Group*") {
                if ($UserObj.RecipientType -ne 'MailUniversalSecurityGroup') {
                    Throw "Group `"$($UserObj.Name)`" must be of type `"MailUniversalSecurityGroup`""
                }
            } elseif ($UserObj.RecipientType -ne 'UserMailbox' -and $UserObj.RecipientType -ne 'MailUser') {
                Throw "`"$($UserObj.Name)`" must be of type `"UserMailbox`",`"MailUser`" (for non-cloud mailbox in an hybrid environment) or `"MailUniversalSecurityGroup`""
            }
        }

    }

    process {

        $Mbx = Get-Mailbox $UserPrincipalName
        if ($Mbx) {
            $FolderIDs = (Get-MailboxFolderStatistics $Mbx.UserPrincipalName | Where-Object FolderType -eq $Type).FolderID
            foreach ($id in $FolderIDs) {
                $null = Add-MailboxFolderPermission ($Mbx.PrimarySmtpAddress + ":" + $id) -User $UserObj.PrimarySmtpAddress -AccessRights $AccessRights
                if ($PassThru) {
                    Get-MailboxFolderPermission ($Mbx.PrimarySmtpAddress + ":" + $id) -User $UserObj.PrimarySmtpAddress `
                    | Select-Object @{Name='PrimarySmtpAddress';E={$Mbx.PrimarySmtpAddress}},FolderName,User,AccessRights
                }
            }            
        }

    }

    end {
        # nothing yet
    }

}