<#
.Synopsis
   The script will reslove users from the groups are listed on groupList file
.DESCRIPTION
   The script will reslove users from the groups are listed on groupListFile (by default c:\temp\groupList.txt)
   and will export it to c:\temp\GroupList\$group.csv
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
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
    # Read groups from file
    $Groups = cat $groupListFile
    }
    Process
    {
    foreach ($Group in $Groups)
        {
            $GroupMembers = Get-ADGroupMember -Identity $Group -Recursive
            foreach ($Member in $GroupMembers)
                {
                    $path = Split-Path $groupListFile
                    Get-ADUser -Identity $Member -Properties office | `
                    Select SamAccountName,Office | `
                    Export-Csv ("$path\$Group.csv").Replace(" ","") -Append -Force 
                }
        }
    }
    End
    {
    }
