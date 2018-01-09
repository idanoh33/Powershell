<#
.Synopsis
   This script will export destination folder permissions to csv
.DESCRIPTION
   Long description
.EXAMPLE
   .\Export-PermissionsToCsv.ps1 -FolderPath C:\temp -CsvPath c:\temp\1.csv
.EXAMPLE
   Another example of how to use this cmdlet
#>

    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $FolderPath,

        # Param2 help description
        $CsvPath
    )

    Begin
    {
    $i= 1
    $files = Get-childitem $FolderPath -recurse -Filter *
    }
    Process
    {

    $acls= $files | Get-Acl -Exclude *.* 
    workflow relove-inparallel
    {
    param($acls,$CsvPath)
    $i= 0
    foreach -parallel ($acl in $acls)
    {
    inlinescript {$l= $i++}

    Write-Progress -activity "Reslove Permissions" -status "Resloving $($acl.PSChildName)" -PercentComplete (($l/$($acls.count))  * 100)
#        $output =
            $acl.Access | % {
            
                New-Object PSObject -Property @{
                    
                    file = ($acl.PSPath).replace('Microsoft.PowerShell.Core\FileSystem::','')
                    Folder = ($acl.PSParentPath).replace('Microsoft.PowerShell.Core\FileSystem::','')
                    Access = $_.FileSystemRights
                    Control = $_.AccessControlType
                    User = $_.IdentityReference
                    Inheritance = $_.IsInherited
                    } 
                } | select-object -Property file, User, Access, Folder | export-csv -Append $CsvPath -force -NoTypeInformation
    }
    }

    relove-inparallel -acls $acls -CsvPath $CsvPath
    
    }
    End
    {
    #$output | select-object -Property file, User, Access, Folder | export-csv $CsvPath -force -NoTypeInformation

    }
