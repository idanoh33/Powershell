# Thanks for Grzegorz Tworek https://github.com/gtworek
# you can find the original script on
# https://github.com/gtworek/PSBits/blob/master/Misc/DumpKeePassDB.ps1
<#
.Synopsis
   The script will import KeepPass database
.DESCRIPTION
   The script will import KeepPass database
   you can change the parameter path of from $dbPath = "C:\Users\IdanOhayon\Documents\test.kdbx"  to your default database location
.EXAMPLE
   Import-KeePassDB -dbPass 'SecretPass'
.EXAMPLE
   Import-KeePassDB -dbPass '64tHEgAr7nst@il' -dbPath c:\temp\database.kdbx
#>
function Import-KeePassDB
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $dbPass,

        # Param2 help description
        $keepassBinary = "C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe",
        $dbPath = "C:\Users\IdanOhayon\Documents\test.kdbx"
    )

    Begin
    {
        
        # Actual dump script
        [Reflection.Assembly]::LoadFile($keepassBinary)
        $ici = [KeepassLib.Serialization.IOConnectionInfo]::new()
        $ici.Path = $dbPath
        $kp = [KeepassLib.Keys.KcpPassword]::new($dbPass)
        $ck = [KeePassLib.Keys.CompositeKey]::new()
        $ck.AddUserKey($kp)
        $db = [KeepassLib.PWDatabase]::new()
        $db.Open($ici, $ck, $null) # if you see an error here, it may mean your password was incorrect
        $entries = $db.RootGroup.GetEntries($true)
        $db.Close()


    }
    Process
    {
        # Done. let's convert $netries to an array.
        $arrExp=@()
        foreach ($entry in $entries)
        {
            $row = New-Object psobject
            $row | Add-Member -Name Title -MemberType NoteProperty -Value ($entry.Strings.ReadSafe('Title'))
            $row | Add-Member -Name UserName -MemberType NoteProperty -Value ($entry.Strings.ReadSafe('UserName'))
            $row | Add-Member -Name Password -MemberType NoteProperty -Value ($entry.Strings.ReadSafe('Password'))
            $row | Add-Member -Name URL -MemberType NoteProperty -Value ($entry.Strings.ReadSafe('URL'))
            $row | Add-Member -Name Notes -MemberType NoteProperty -Value ($entry.Strings.ReadSafe('Notes'))
            $arrExp += $row
        }
    }
    End
    {
        # Let's display the array
        if (Test-Path Variable:PSise)
        {
            $arrExp | Out-GridView
        }
        else
        {
            $arrExp | Format-Table
        }
    }
}
