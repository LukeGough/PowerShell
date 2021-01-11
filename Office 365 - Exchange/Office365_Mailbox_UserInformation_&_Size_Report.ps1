<###
Author: Luke Gough
Last Edited: 11/01/2021
###>
# If using MFA you need to install “Exchange Online PowerShell Module” (EXO). Find more information in the below link
    # https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/

# If using Powershell or Powershell ISE you can use the below commented out line to import the MFA Enabled Exchange Online Powershell module into ISE
    # Make sure the installation location is correct if you are going to do this
    <#

    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    . "$MFAExchangeModule" // command is ran on new line not on the same line as the above
        
    #>

try {
    # Create MFA module so script can be used in Powershell/Powershell ISE
    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    . "$MFAExchangeModule"
}
catch {
    Write-Host "Unable create MFA Module" -ForegroundColor Red
}

# Prompt for MFA/Office 365 authentication and credentials
try{
    # Add PresentationFramework to allow MessageBox to be created
    Add-Type -AssemblyName PresentationFramework

    # Create a MessageBox to select MFA or non MFA sign-in
    $msgBox = [System.Windows.MessageBox]::Show('Do you require MFA Authentication?','MFA Authencation','YesNo', 'Information')

    switch ($msgBox) {
        # MFA
        'Yes' {
            Connect-EXOPSSession
        }
        'No' {
            # Non MFA
            $msgBox = [System.Windows.MessageBox]::Show('Do you require Office 365 Authentication?','Office 365 Authencation','YesNo', 'Information')
            switch ($msgBox) {
                'Yes' {
                    try {
                        # Prompt for Oiffce365 Administraitor credentials
                        $exchcred = Get-Credential -Message "Enter Office365 Admin Credentials"

                        # Create a new PSSession using the credenticals provided
                        $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $exchcred -Authentication Basic -AllowRedirection;

                        # Import PSSession
                        Import-PSSession $s;
                    }
                    catch {
                        Write-Host "Unable to sign-in to Office 365" -ForegroundColor Red
                        Exit
                    }
                }
                'No' {
                    Exit
                }
            }
        }
    }
}
catch {
    Write-Host "Unable to sign-in to Office 365" -ForegroundColor Red
    Exit
}


# Get current Date and Time
$dateTime = Get-Date -format "MM-dd-yyyy_HH-mm"

# Check if output path exists and set CSV File Path
$path = "C:\temp"
If(!(test-path $path)) {
    New-Item -ItemType Directory -Force -Path $path
    $csvFilepath = "$path\O365SizeReport_$dateTime.csv"
} elseif (test-path $path) {
    $csvFilepath = "$path\O365SizeReport_$dateTime.csv"
} else {
    Write-Host "$path does not exist and can't be created" -ForegroundColor Red
    Exit
}

# Create empty arrays
$exportArray = @()
$mailboxesArray = @()
$erroredMailboxes = @()

# Create a searchBase
$searchBase = Get-Mailbox -Resultsize Unlimited

# Add searchBase to mailboxesArray
$mailboxesArray += $searchBase

# Foreach mailbox inside mailboxesArray get results 
foreach ($mailbox in $mailboxesArray) {
    try {
        # Output currently mailbox being checked
        Write-Host "Getting information for $mailbox" -ForegroundColor Yellow

        # Create Empty Custom Object
        $finalResults = [pscustomobject] @{}

        # Get mailbox DisplayName,MailboxType attributes
        $results1 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Select-Object DisplayName,RecipientType

        # Get MailboxStatistics totalMailboxSize,mailboxItems,lastLogin FolderSize attributes
        $results2 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxStatistics | Select-Object TotalItemSize,ItemCount,LastLogonTime

        # Get MailboxFolderStatistics inboxSize,sentItemsSize,deletedItemsSize attributes
        $results3 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "Inbox"} | Select-Object FolderSize
        $results4 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "SentItems"} | Select-Object FolderSize
        $results5 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "DeletedItems"} | Select-Object FolderSize
        $results6 = Get-Mailbox -Identity $mailbox.UserPrincipalName | Get-MailboxFolderStatistics | where {$_.FolderType -eq "Calendar"} | Select-Object FolderSize,ItemsInFolder

        # Create custom PS object and store all previous results into the object and format
        $finalResults = [pscustomobject] @{
            'User'                              = $results1.DisplayName
            'Mailbox Type'                      = $results1.RecipientType
            'Total Mailbox Size (Mb)'           = $results2.TotalItemSize
            'Mailbox Items'                     = $results2.ItemCount
            'Inbox Folder Size (Mb)'            = $results3.FolderSize
            'Sent Items Folder Size (Mb)'       = $results4.FolderSize
            'Deleted Items Folder Size (Mb)'    = $results5.FolderSize
            'Calendar Items Folder Size'        = $results6.FolderSize
            'Calendar Items'                    = $results6.ItemsInFolder
            'Last Mailbox Logon'                = $results2.LastLogonTime
        }

        # Add finalResults custom object into the exportArray
        $exportArray += $finalResults
    } catch {
        $erroredMailboxes += $mailbox.UserPrincipalName
    }
    

}

# Print out all errored mailboxes to screen
If($erroredMailboxes.count -gt 0) {
    Write-Host "Unable to got information for the following mailboxes..." -ForegroundColor Yellow
    foreach ($mailbox in $erroredMailboxes) {
        Write-Host $mailbox
    }
}

# Export export array to CSV
$exportArray | Export-Csv -path $csvFilepath -NoTypeInformation

# Script Ended output
Write-Host "Script has finished" -ForegroundColor Green
