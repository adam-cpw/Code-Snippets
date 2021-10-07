$leaver = Read-Host -Prompt "Name of Leaver"
$leaveDate = Read-Host -Prompt "Leaving Date"
$forwardedTo = Read-Host -Prompt "Emails to be forwarded to"
if ($forwardedTo -ne "" ){
    $forwardedUntil = Read-Host -Prompt "Until"
}

$ol = new-object -ComObject "Outlook.Application"
$mapi = $ol.GetNamespace("mapi")

#CPW Shared Calendar - UID found through $mapi.Folders.Item("Calendar Name")
$SharedCalendar = $mapi.GetFolderFromID("000000003C5A2A86517F774FA33E52D77D84D2830100E49DF6BB9172BE4ABC1743BC3CA41FD50000000001080000")

if ($forwardedTo -ne "") {
    $BodyMessage = "Emails to be forwarded to $forwardedTo"
} else {
    $BodyMessage = "Emails not to be forwarded"
}


$NewItem = $SharedCalendar.Items.Add(1)
$NewItem.Subject = "AD - Disable account for $leaver"
$NewItem.Body = $BodyMessage
$Recipient = $NewItem.Recipients.Add("adam@cpwc.co.uk")
$Recipient.Type = 1
$NewItem.BusyStatus = 0
$NewItem.Start = "$leaveDate 4:30:00 PM"
$NewItem.End = "$leaveDate 5:00:00 PM"
$NewItem.ReminderSet = $true
$NewItem.ReminderMinutesBeforeStart = 0
$NewItem.MeetingStatus = 1
$NewItem.Importance = 1
$NewItem.Save()
$NewItem.Send()

if ($forwardedTo -ne "") {
    $NewItem = $SharedCalendar.Items.Add(1)
    $NewItem.Subject = "AD - Contact $forwardedTo about $leaver forwarder"
    $NewItem.Body = "Forwarder for $leaver has been in place since $leaveDate. Contact to see if forwarder still required."
    $Recipient = $NewItem.Recipients.Add("adam@cpwc.co.uk")
    $Recipient.Type = 1
    $NewItem.BusyStatus = 0
    $NewItem.Start = "$forwardedUntil 9:0:00 AM"
    $NewItem.End = "$forwardedUntil 9:30:00 AM"
    $NewItem.ReminderSet = $true
    $NewItem.ReminderMinutesBeforeStart = 0
    $NewItem.MeetingStatus = 1
    $NewItem.Importance = 1
    $NewItem.Save()
    $NewItem.Send()
}

pause