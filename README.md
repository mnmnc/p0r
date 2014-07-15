p0r
===

Powershell replacement for Outlook rules

### Why?

You might wonder - "Why would anyone resign from using Outlook rules creator?". 

Been there, done that. I'm not sure if you stumbled on the same problem with Outlook rules that troubled me but here it is:
You have very limited space for Outlook rules. Here is a quote from microsoft [http://support.microsoft.com/kb/886616](KB page):

"""
The rules size limit for mailboxes in Exchange Server 2007 (and later) has a default size of 64 KB per mailbox. The total rules size limit is also customizable limit up to 256 KB per mailbox.

Mailboxes on Exchange Server 2003

This behavior occurs if the rules that are in your mailbox exceed a size of 32 kilobytes (KB). The total rules size limit for mailboxes on Exchange Server 2003 is 32 KB. The rules limit for Exchange 2003 cannot be changed.
"""

32, 64 kb - that is definitely not enough for me. 256 kb is better but still, I had a requirement for more than hundred of rules and Exchange server was limiting me.

That is why I've created a powershell script that is able to move the emails in the Inbox to other folders, subfolders and even PST file. On top of that is shows me a specific color depending on to where the email is being moved.

### Features

You can move emails virtually by matching one of their properties to specified value. Email object has many properties.
