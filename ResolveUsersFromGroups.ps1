#requires -Module ActiveDirectory
<#
.Synopsis
   The script will generate CSV with the users for every group in $groupListFile
.DESCRIPTION
   The script will generate CSV with the users for every group in $groupListFile
.EXAMPLE
   . C:\temp\ResolveUsersFromGroups.ps1
.EXAMPLE
    . C:\temp\ResolveUsersFromGroups.ps1 -groupListFile C:\temp\groupList.txt
#>
    [CmdletBinding()]

    Param
    (
        # Param1 help description
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $groupListFile= "c:\temp\groupList.txt"
    )

    Begin
    {
    $ErrorActionPreference = 'Stop'
    
    # Import ActiveDirectory module
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Read groups from file
    $Groups = cat $groupListFile
    }
    Process
    {
    foreach ($Group in $Groups)
        {
            # get users from AD Group
            $GroupMembers = Get-ADGroupMember -Identity $Group -Recursive
            foreach ($Member in $GroupMembers)
                {
                    # Take group list file path for output files
                    $path = Split-Path $groupListFile
                    
                    # Find user and export to file in $path, if there are spaces they will be remove
                    Get-ADUser -Identity $Member -Properties office | `
                    Select SamAccountName,Office | `
                    Export-Csv ("$path\$Group.csv").Replace(" ","") -Append -Force 
                }
        }
    }
    End
    {
    }
