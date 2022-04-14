# Automations
Various small scripts to automate tedious tasks

* `LeaverCalendar.ps1` - Add leavers to the group calendar. Interactive.
* `QuoteOTron.ps1` - Build a quick equipment quote. Interactive

# Code snippets
Things too small for their own file, but still worth writing down

## Export distribution group members to file
Exports a single email address per line. Need to run `Connect-ExchangeOnline` or run from an on prem Exchange powershell session first:
```powershell
Get-DistributionGroupMember -Identity "<UPN or Identity>" | ft PrimarySmtpAddress > /path/to/output.txt
```

## Add group members from file
Single UPN or email address per line. Need to run `Connect-ExchangeOnline` or run from an on prem Exchange powershell session first:
```powershell
ForEach($line in Get-Content <Path to File>){
    Add-DistributionGroupMember -Identity "<Group Identifier or email>" -Member $line.trim() 
}
```

## Convert hybrid user to shared mailbox
This needs to be ran from the on-prem Exchange powershell:
```powershell
Set-ADUser -Identity ((Get-Recipient <shared mailbox email adddress>).samaccountname) -Replace @{msExchRemoteRecipientType=100;msExchRecipientTypeDetails=34359738368}
```

Followed by a sync on the on-prem DC's Azure AD connector:
```powershell
Start-ADSyncSyncCycle -PolicyType Delta
```

Wait for the sync to finish. Log into Office 365 Exchange, either PS or web portal, and convert the user to a shared mailbox.

Any licenses can then be removed in the admin centre.

## Bulk add members to Azure AD group from file
After running `Connect-AzureAD` to log in, first run Get-AzureADGroup to find the group object ID:
```powershell
Get-AzureADGroup -Filter "DisplayName eq '<Group Display Name>'" | Select ObjectID
```

Then using this ID, add them from file:
```powershell
ForEach ($line in Get-Content <Path to File>) {
    Add-AzureADGroupMember -ObjectId <Group object ID> -RefObjectId $(Get-AzureADUser -ObjectID $line).ObjectID
}
```

## Elevate a powershell session from within Powershell
The RunAs verb here elevates to admin. Replace powershell & arguments with whatever is needed.
```powershell
$startprocessParams = @{
    FilePath = "powershell.exe"
    ArgumentList = '-ExecutionPolicy', 'Bypass"', '-File', '"E:\Path\to\file.ps1"'
    Verb = 'RunAs'
    PassThru = $true
    Wait = $true
}
Start-Process @startprocessParams
```

## Get a list of delegated access permissions for a mailbox
Prints a semicolon seperated list of UPNs. Great for Egress.
```powershell
$list = ""; ForEach ($user in Get-EXOMailboxPermission -Identity "<Mailbox Identifier>"){$list += $user.User + ";"}; Write-Host $List
```