# OUTLOOK RULES #
#################

# VARIABLES
$index=0;
$pstPath = "D:\path\to\pst\file.pst"     

# DISPLAY INFO
function display( [string]$subject, [string]$color , [string]$out)  {

    # REQUIRED LENGTH OF STRING
    $len = 20

    # STRINGS THAT ARE LONGER WILL BE CUT DOWN,
    # STRINGS THAT ARE TO SHORT WILL BE MADE LONGER
    if ( $subject.length -lt 20 ){
        $toadd=20-$subject.length;
        for ( $i=0; $i -lt $toadd; $i++ ){
            $subject=$subject+" ";
        }
        $len = $subject.length
    }
    else { $len = 20 }
    
    $index=$index+1
    Write-host -ForegroundColor $color -nonewline " |" ((($subject).ToString()).Substring(0,$len)).ToUpper()
}


# CREATING OUTLOOK OBJECT
$outlook = New-Object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")

# GETTING PST FILE THAT WAS SPECIFIED BY THE PSTPATH VARIABLE
$pst = $namespace.Stores | ?{$_.FilePath -eq $pstPath}

# ROOT FOLDER
$pstRoot = $pst.GetRootFolder()

# SUBFOLDERS
$pstFolders = $pstRoot.Folders

# PERSONAL SUBFOLDER
$personal = $pstFolders.Item("Personal")

# INBOX FOLDER 
$DefaultFolder = $namespace.GetDefaultFolder(6)

# INBOX SUBFOLDERS
$InboxFolders = $DefaultFolder.Folders

# DELETED ITEMS
$DeletedItems = $namespace.GetDefaultFolder(3)

# EMAIL ITEMS
$Emails = $DefaultFolder.Items

# PROCESSING EMAILS
Foreach ($Email in $Emails) {
    
    # NEW LINE EVERY 4 EMAILS - FOR READABILITY PURPOSES
    $index=$index+1
    if ( ($index % 4) -eq 0 ) {  Write-host -nonewline -ForegroundColor DarkGray " |`n" }

    # IF EMAILS ARE SENT TO MYSELF -> MOVE TO PERSONAL FOLDER UNDER PST FILE
    # !      DESTINATION FOLDER SPECIFIED BEFOREHAND AS A VARIABLE
    IF ($Email.To -eq "MySurname, MyName") {    
        $Email.Move($personal) | out-null
        display  ([string]$Email.Subject ) ([string]"Cyan")
        continue
    }
 
    # MOVE EMAILS WITH SPECIFIC STRING IN TITLE TO THE SUBFOLDER /RANDOM/ UNDER PST FILE
    # !      DESTINATION FOLDER SPECIFIED INLINE
    IF ($Email.Subject -match "SPECIFIC STRING IN TITLE") {
        $Email.Move($pstFolders.Item("Random")) | out-null
        display  ([string]$Email.Subject ) ([string]"Yellow")
        continue
    }
   
    # MOVING NOT IMPORTANT MESSAGES TO DELETED ITEMS
    # !     MARKING EACH MOVED ITEM AS UNREAD
    IF ($Email.Subject -match "not important" -or $Email.Subject -match "not-important" ) {
        $Email.UnRead = $True
        $Email.Move($DeletedItems) | out-null
        display  ([string]$Email.Subject ) ([string]"Red")
        continue
    }

    # MOVING MESSAGES FROM SPECIFIC AD OBJECT TO DELETED ITEMS
    IF ($Email.SenderEmailAddress -match "/O=COMPANY/OU=AD GROUP/CN=RECIPIENTS/CN=SOME-NAME") {
        $Email.Move($DeletedItems) | out-null
        display  ([string]$Email.Subject ) ([string]"Red")
        continue
    }

    # MOVING MESSAGES FROM SPECIFIC EMAIL ADDRESS TO DELETED ITEMS
    IF ($Email.SenderEmailAddress -match "email@gmail.com") {
        $Email.Move($DeletedItems) | out-null
        display  ([string]$Email.Subject ) ([string]"Red")
        continue
    }   
    
    # DISPLAYING OTHER PROCESSED BUT NOT MOVED EMAILS
    display  ([string]$Email.Subject ) ([string]"DarkGray") ([string]"out")
    
    # FAKE DELAY IF YOU NEED ONE
    # Start-Sleep -s 2
}

Write-host ""