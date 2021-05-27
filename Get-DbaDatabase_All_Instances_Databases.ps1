param(
    [switch]$GetDbaUsers,
    [switch]$GetDbaUserPermissions
)

function Get-DbaUsers {
    # Export-DbaUsers
    Try
    {
        Foreach ($sqlinst in $sqlinsts)
        {
            # Object variables
            $ComputerName = $sqlinst.ComputerName
            $InstanceName = $sqlinst.InstanceName
            $SqlInstance  = $sqlinst.SqlInstance
            $DatabaseName = $sqlinst.Name
            $Comp_Inst    = $ComputerName + "\" + $InstanceName
            $FileName     = $SqlInstance + "_" + $DatabaseName + ".sql"

            # Export DBAUser for each object found in Get-DbaDatabase
            Export-DbaUser -SqlInstance $Comp_Inst -database $DatabaseName -FilePath $exportPath\$FileName
                
            # Print file created message 
            Write-Host -ForegroundColor Green "File created:" $exportPath"\"$FileName
        }
    }
    Catch
    {
        # Print Error message (To be expanded on for actual error reporting)
        Write-Host "Error in Export-DbaUser"
    }
}

function Get-DbaUserPermissions {
    # Get-DbaUserPermissions
    Try
    {
        Foreach ($sqlinst in $sqlinsts)
        {
            # Object variables
            $DateTime      = Get-Date -Format "MM-dd-yyyy_HH-mm"
            $ComputerName  = $sqlinst.ComputerName
            $InstanceName  = $sqlinst.InstanceName
            $SqlInstance   = $sqlinst.SqlInstance
            $DatabaseName  = $sqlinst.Name
            $Comp_Inst     = $ComputerName + "\" + $InstanceName
            $FileName      = $ComputerName + "_" + $InstanceName + "_UserPermissions_" + $DateTime + ".csv"
            $dbUsersExport = @()

            # Each DB get all users
            $dbUsers = Get-DbaUserPermission -SqlInstance $Comp_Inst -database $DatabaseName
                
            # Loop through each user, create custom object, store object in dbUsersExport array
            Foreach ($dbUser in $dbUsers)
            {
                # Create empty Custom PSObject
                $userResult = [pscustomobject] @{}

                # Set custom object with $dbUser properties
                # If property is $null sets to empty string
                    # Fix for error "Export-Csv : Cannot bind argument to parameter 'InputObject' because it is null."
                $userResult = [pscustomobject] @{
                    'ComputerName'       = if($dbUser.ComputerName -eq $null){""}Else{$dbUser.ComputerName};
                    'InstanceName'       = if($dbUser.InstanceName -eq $null){""}Else{$dbUser.InstanceName};
                    'SqlInstance'        = if($dbUser.SqlInstance -eq $null){""}Else{$dbUser.SqlInstance};
                    'Object'             = if($dbUser.Object -eq $null){""}Else{$dbUser.Object};
                    'Type'               = if($dbUser.Type -eq $null){""}Else{$dbUser.Type};
                    'Member'             = if($dbUser.Member -eq $null){""}Else{$dbUser.Member};
                    'RoleSecurableClass' = if($dbUser.RoleSecurableClass -eq $null){""}Else{$dbUser.RoleSecurableClass};
                    'SchemaOwner'        = if($dbUser.SchemaOwner -eq $null){""}Else{$dbUser.SchemaOwner};
                    'Securable'          = if($dbUser.Securable -eq $null){""}Else{$dbUser.Securable};
                    'GranteeType'        = if($dbUser.GranteeType -eq $null){""}Else{$dbUser.GranteeType};
                    'Grantee'            = if($dbUser.Grantee -eq $null){""}Else{$dbUser.Grantee};
                    'Permission'         = if($dbUser.Permission -eq $null){""}Else{$dbUser.Permission};
                    'State'              = if($dbUser.State -eq $null){""}Else{$dbUser.State};
                    'Grantor'            = if($dbUser.Grantor -eq $null){""}Else{$dbUser.Grantor};
                    'GrantorType'        = if($dbUser.GrantorType -eq $null){""}Else{$dbUser.GrantorType};
                    'SourceView'         = if($dbUser.SourceView -eq $null){""}Else{$dbUser.SourceView};
                }
                    
                # Add custom object to dbUsersExport
                $dbUsersExport += $userResult
            }

        }

        # Export dbUsersExport array to CSV
        $dbUsersExport | Export-Csv -Path $exportPath\$FileName -NoTypeInformation

        # Print file created message
        Write-Host -ForegroundColor Green "File created:" $exportPath"\"$FileName
    }
    Catch
    {
        # Print Error message (To be expanded on for actual error reporting)
        Write-Host "Error in Get-DbaUserPermissions"
    }
}

Try
{
    # Import the DBATools Module - Requires Internet
    Get-Module dbatools

    $dbatoolsPathDest = "C:\Program Files\WindowsPowerShell\Modules\dbatools"
    If(!(test-path $dbatoolsPathDest))
    {
        Write-Host "DBATools module missing."
    }
    Else
    {
        # Import DBATools
        If(!(Get-Module -Name dbatools))
        {
            Write-Host -ForegroundColor Yellow "Importing DBATools..."
            Import-Module dbatools
        }

        # Check if c:\temp exists
        $exportPath = "c:\Temp"
        If(!(test-path $exportPath))
        {
            New-Item -ItemType Directory -Force -Path $exportPath
            Write-Host -ForegroundColor Green "Temp folder created - " + $exportPath
        }

        # Get local SQL Instance and add to array
        $sqlInsts = @()
        $sqlInsts += Get-DbaDatabase -SqlInstance localhost | Select-Object ComputerName,InstanceName,SqlInstance,Name

        # Export-DbaUsers
        if($GetDbaUsers){Get-DbaUsers}

        # Get-DbaUserPermissions
        if($GetDbaUserPermissions){Get-DbaUserPermissions}
    }
}
Catch
{
    # Print Error message (To be expanded on for actual error reporting)
    Write-Host -ForegroundColor Red "Error!"
}