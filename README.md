p0r
===

Powershell replacement for Outlook rules

### Why?

You might wonder - "Why would anyone resign from using Outlook rules wizard?". 

I'm not sure if you stumbled on the same problem with Outlook rules that troubled me, but here it is:
You have very limited space for Outlook rules. This is a quote from microsoft [KB page](http://support.microsoft.com/kb/886616):

> The rules size limit for mailboxes in Exchange Server 2007 (and later) has a default size of *64 KB per mailbox*. The total rules size limit is also customizable limit up to 256 KB per mailbox.

> Mailboxes on Exchange Server 2003
This behavior occurs if the rules that are in your mailbox exceed a size of 32 kilobytes (KB). The total rules size limit for mailboxes on Exchange Server 2003 is 32 KB. The rules limit for Exchange 2003 cannot be changed.


32, 64 kb - that is definitely not enough for me. 256 kb is better but still, I had a requirement for more than hundred of rules and Exchange server was limiting me.

### How?

That is why I've created a powershell script that is able to move the emails in the Inbox to other folders, subfolders and even PST file. On top of that is shows me a specific color depending on to where the email is being moved.

Here's how it will look like in the console:
![powershell outlook rules](https://raw.githubusercontent.com/mnmnc/img/master/powershell_rules.png)

It is 20 characters per item, 4 items a row, plus pipe characters and some spaces, so the script assumes you have a `94 columns` in the console window.

### Features

You can move emails by matching one of their properties to specified value. Email object has many properties:

``` Actions
AlternateRecipientAllowed
Application
Attachments
AutoForwarded
AutoResolvedWinner
BCC
BillingInformation
Body
BodyFormat
Categories
CC
Class
Companies
Conflicts
ConversationIndex
ConversationTopic
CreationTime
DeferredDeliveryTime
DeleteAfterSubmit
DownloadState
EnableSharedAttachments
EntryID
ExpiryTime
FlagDueBy
FlagIcon
FlagRequest
FlagStatus
FormDescription
GetInspector
HasCoverSheet
HTMLBody
Importance
InternetCodepage
IsConflict
IsIPFax
IsMarkedAsTask
ItemProperties
LastModificationTime
Links
MAPIOBJECT
MarkForDownload
MessageClass
Mileage
NoAging
OriginatorDeliveryReportRequested
OutlookInternalVersion
OutlookVersion
Parent
Permission
PermissionService
PropertyAccessor
ReadReceiptRequested
ReceivedByEntryID
ReceivedByName
ReceivedOnBehalfOfEntryID
ReceivedOnBehalfOfName
ReceivedTime
RecipientReassignmentProhibited
Recipients
ReminderOverrideDefault
ReminderPlaySound
ReminderSet
ReminderSoundFile
ReminderTime
RemoteStatus
ReplyRecipientNames
ReplyRecipients
Saved
SaveSentMessageFolder
SenderEmailAddress
SenderEmailType
SenderName
SendUsingAccount
Sensitivity
Sent
SentOn
SentOnBehalfOfName
Session
Size
Subject
Submitted
TaskCompletedDate
TaskDueDate
TaskStartDate
TaskSubject
To
ToDoTaskOrdinal
UnRead
UserProperties
VotingOptions
VotingResponse
```

I've used properties like `To`, `Subject` and `SenderEmailAddress` but you can customize it however you like.

### Customize it to suite your needs

For example let use `Subject` field and move emails that will have subject matching to string `Alert` to Deleted items folder. It can be done by adding following condition to main `for` loop within the scipt:

```
    IF ($Email.Subject -match "Alert" ) {
        $Email.Move($DeletedItems) | out-null
        continue
    }
```
If you would want to see the email subject in the console after the move, you can add an additional function call before `continue` :

```
display ([string]$Email.Subject ) ([string]"Red")
```

