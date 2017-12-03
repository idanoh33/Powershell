#requires -Module ActiveDirectory
<#
.Synopsis
   The script will reslove users from the groups are listed on groupList file
.DESCRIPTION
   The script will reslove users from the groups are listed on groupListFile (by default c:\temp\groupList.txt)
   and will export it to c:\temp\GroupList\$group.csv
.EXAMPLE
   . C:\temp\ResolveUsers.ps1
.EXAMPLE
    . C:\temp\ResolveUsers.ps1 -groupListFile C:\temp\groupList.txt
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
