<#####
Descirption: Script will get keyword from CSV column "name" and will search Active Directory groups containing the keyword.
    For each keyword / search word, it will:
        * Search for groups containing that word (case-insensitive)
        * Return details for all matching groups
        * Handle cases where no matches are found
        * Track any errors that occur during the search

Run Requirements:
    * Expects a CSV File with a column of "name".
    * Prepare your input CSV file with a column named "name" which will be the keyword which is searched.
    * Access to load PowerShell Module Active Directory.
    * Read access to domain.

Usage:
    * Run the script and select your CSV file when prompted.
    * Script will export results to a CSV File, output CSV will contain:
        * The search word used
        * All matching group details
        * Status (Found/NoMatchesFound/Error)
        * Error messages if applicable

Author: Luke Gough
Created: 04-11-2024
Last Edited: 05-11-2024
#####>

# Import required modules
Import-Module ActiveDirectory


# FUNC - File Selection
function Select-File {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    
    # Set file restults to show in FileBrowser window
    $fileBrowser.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    
    # Set FileBrowser Title
    $fileBrowser.Title = "Select file"

    # If selection successfull get file
    if ($fileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileBrowser.FileName
    }
    # If selection failure return null
    return $null
}

# FUNC - Search and Results
function Get-ADGroupBySearchWord {
    # Expected function parameters
    param(
        [Parameter(Mandatory=$true)]
        [string]$searchWord
    )

    Try {
        # Search for groups containing the searchWord
        $groups = Get-ADGroup -Filter "Name -like '*$searchWord*'" -Properties * -ErrorAction Stop

        if ($groups) {
            # Convert each group to a custom object
            $groupDetails = ForEach ($group in $groups) {
                [PSCustomObject]@{
                    "SearchWord"        = $searchWord
                    "GroupName"         = $group.Name
                    "CN"                = $group.CN
                    "Description"       = $group.Description
                    "DistinguishedName" = $group.DistinguishedName
                    "Created"           = $group.Created
                    "Status"            = "Found"
                    "Error"             = $null
                }
            }
            return $groupDetails
        }
        else {
            # Return ObjectNotFound
            return [PSCustomObject]@{
                    "SearchWord"        = $searchWord
                    "GroupName"         = $null
                    "CN"                = $null
                    "Description"       = $null
                    "DistinguishedName" = $null
                    "Created"           = $null
                    "Status"            = "NoMatchesFound"
                    "Error"             = "No groups found containing '$searchWord'"
                }
        }
    }
    Catch {
        $errorMsg = "Error searching for groups with keyword '$searchWord': $($_.Exception.Message)"
        Write-Error $errorMsg

        return [PSCustomObject]@{
            "SearchWord"        = $searchWord
            "GroupName"         = $null
            "CN"                = $null
            "Description"       = $null
            "DistinguishedName" = $null
            "Created"           = $null
            "Status"            = "NoMatchesFound"
            "Error"             = "No groups found containing '$searchWord'"
        }
    }
}

# FUNC - Main
function Search-ADGroups {
    # Setup output path
    $scriptPath = $PSScriptRoot
    if (-not $scriptPath) { $scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition}
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputPath = Join-Path $scriptPath "ADGroupSearchResults_$timestamp.csv"


    # Prompt for input file
    $inputPath = Select-File
    if (-not $inputPath) {
        Write-Warning "No file selected. Terminating."
        return
    }

    Try {
        # Read File
        $searchWords = Import-Csv -Path $inputPath
        if (-not $searchWords) {
            throw "No data found in file."
        }

        Write-Host "Processing search words from '$inputPath'"

        # Process each search word
        $results = ForEach ($row in $searchWords) {
            # Assumes the CSV column will be named 'name'
            $word = $row.name
            if (-not $word) {
                Write-Warning "Skipping row - No name found"
                continue
            }

            Write-Host "Searching for groups containing: $word"
            Get-ADGroupBySearchWord -searchWord $word
        }

        # Export results to CSV
        $results | Export-Csv -Path $outputPath -NoTypeInformation
        Write-Host "`n--- Outfile: '$outputPath'"


        # Display summary
        Write-Host "`n--- Summary"
        Write-Host "Total Keywords Searched: " ($searchWords | Measure-Object).Count
        Write-Host "Total Groups Found: " ($results | Where-Object { $_.Status -eq 'Found' } | Measure-Object).Count
        Write-Host "Searches with No Matches: " ($results | Where-Object { $_.Status -eq 'NoMatchesFound' } | Measure-Object).Count
        Write-Host "Searches with Errrors: " ($results | Where-Object { $_.Status -eq 'Error' } | Measure-Object).Count
    }
    Catch {
        Write-Error "Critical errror in script execution: $($_.Exception.Message)"
    }
}

# Run Main
Search-ADGroups