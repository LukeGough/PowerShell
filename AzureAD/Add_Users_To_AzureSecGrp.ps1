# Connect to AzureAD Module
# Use Command Connect-AzureAd

# User Group Variables
$secGrpName = "{Azure Security Group Name}" # Should match the group which adds mailbox licensing.
$secGrp     = Get-AzureADGroup -SearchString $secGrpName
$secGrpId   = $secGrp.ObjectId
$users      = Import-CSV "{CSV File Location}" # Update this File for different batches.

# Logging Variables
$dateTime   = Get-Date -Format "ddMMyyyy"
$log_file   = "{Log Directory}\User_SecGrp_$dateTime.txt" # Script doesn't check if folder exists, but does check if folder exists.
ForEach ($user in $users)
{
    Try
    {
        # Get ObjectId for the user
        $userId = Get-AzureADUser -ObjectID $user.mail
        
        # Add user to group
        Add-AzureADGroupMember -ObjectId $secGrpId -RefObjectId $userId.ObjectId
    }
    Catch
    {
        # Check if log file exists
        If(-not(Test-Path -path $log_file))
        {
            Try
            {
                # Create new log file
                New-Item -ItemType File -Path $log_file -Force -ErrorAction Stop
                Write-Host "Log file create: $log_file"
            }
            Catch
            {
                throw $_.Exception.Message
            }
        }
        # Write error to log file
        Write-Host -ForegroundColor Red "Error: Unable to add $user.mail to $secGrpName" 
        $timeStamp = Get-date -format G
        Add-Content "$timeStamp" -path $log_file
        Add-Content $Error -path $log_file
    }
    Finally
    {
        $Error.Clear()
    }
}
