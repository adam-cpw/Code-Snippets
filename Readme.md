# Automations
Various small scripts to automate tedious tasks

* `LeaverCalendar.ps1` - Add leavers to the group calendar. Interactive.

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
Set-RemoteMailbox "mailbox@domain.com" -Type Shared
```

Followed by a sync on the on-prem DC's Azure AD connector:
```powershell
Start-ADSyncSyncCycle -PolicyType Delta
```

From here wait for a little while (15 - 30 mins approx), then 365 license can be downgraded to an E1.

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