<###
Author: Luke Gough
Last Edited: 05/01/2021
###>
# Get current Date and Time
$dateTime = get-date -format "MM-dd-yyyy_HH-mm"

# Set CSV File Path
$csvFilepath = "c:\temp\O365SizeReport_$dateTime.csv"

# Create empty arrays
$exportArray = @()
$mailboxesArray = @()

# Create a searchBase
$searchBase = Get-mailbox -Resultsize Unlimited

# Add searchBase to mailboxesArray
$mailboxesArray += $searchBase

# Foreach mailbox inside mailboxesArray get results 
foreach ($mailbox in $mailboxesArray) {

    # Get mailbox DisplayName,MailboxType attributes
    $results1 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Select-Object DisplayName,RecipientType

    # Get MailboxStatistics totalMailboxSize,mailboxItems,lastLogin FolderSize attributes
    $results2 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxStatistics | Select-Object TotalItemSize,ItemCount,LastLogonTime

    # Get MailboxFolderStatistics inboxSize,sentItemsSize,deletedItemsSize attributes
    $results3 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "Inbox"} | Select-Object FolderSize
    $results4 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "SentItems"} | Select-Object FolderSize
    $results5 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "DeletedItems"} | Select-Object FolderSize

    # Create custom PS object and store all previous results into the object and format
    $finalResults = [pscustomobject] @{
        'User' = $results1.DisplayName
        'Mailbox Type' = $results1.RecipientType
        'Total Mailbox Size (Mb)' = $results2.TotalItemSize
        'Mailbox Items' = $results2.ItemCount
        'Inbox Folder Size (Mb)' = $results3.FolderSize
        'Sent Items Folder Size (Mb)'` = $results4.FolderSize
        'Deleted Items Folder Size (Mb)' = $results5.FolderSize
        'Last Mailbox Logon' = $results2.LastLogonTime
    }

    # Add finalResults custom object into the exportArray
    $exportArray += $finalResults
}

# Export export array to CSV
$exportArray | Export-Csv -path $csvFilepath -NoTypeInformation
