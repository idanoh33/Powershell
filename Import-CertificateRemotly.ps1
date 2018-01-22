<#
.Synopsis
   This script will import certificate remotly
.DESCRIPTION
   This script will copy certificate from local to remote host
   and import certificate remotly by running invoke-command
   please you will be prompted for credentials
.EXAMPLE
   .\Import-CertificateRemotly.ps1 -certificatePath c:\tmp\server1.cer -destinationServer devtest1
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
        $certificatePath,

        # Param2 help description
        
        $destinationServer
    )

    Begin
    {
    # resolve certificate name from path
    $certificateName= split-path $certificatePath -Leaf

    # copy certificate to remote server
    copy $certificatePath "\\$destinationServer\c$\ProgramData\$certificateName" -Force
    }
    Process
    {
    # run remote command
    Invoke-Command -ScriptBlock {
    
    # Import certificate (on remote server)
    Import-Certificate -FilePath "$env:ProgramData\$($args[0])" -CertStoreLocation Cert:\LocalMachine\my

    
    } -Credential (Get-Credential) -ComputerName $destinationServer -ArgumentList $certificateName

    }
    End
    {
    }
