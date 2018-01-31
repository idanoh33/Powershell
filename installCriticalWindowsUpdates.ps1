<#
.Synopsis
   The script will  Install critical windows updates
.DESCRIPTION
   The script will install critical windows updates'Security Updates', 'Critical Updates'
.EXAMPLE
   Example of how to use this cmdlet
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
        $Param1,

        # Param2 help description
        [int]
        $Param2
    )

    Begin
    {
        # Install PSwindowsUpdate module if not exist
        if (!(get-module pswindowsupdate -listavailable))
            {Install-Module pswindowsupdate -Force}
        
        # Import PSwindowsUpdate module
        Import-Module pswindowsupdate
    }
    Process
    {
        # select categories
        $categories = 'Security Updates', 'Critical Updates'

        # Install windows updates for the categories above
        Install-WindowsUpdate -Category $categories -verbose
    }
    End
    {
    }
