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
    
# resolve certificate name from path
$certificateName= split-path $certificatePath -Leaf

# copy certificate to remote server
Copy-VMGuestFile -Source $certificatePath -Destination $env:ProgramData -VM VM -GuestToLocal -GuestUser administrator -GuestPassword pass2


# Set script content
$script = @"
Import-Certificate -FilePath "$env:ProgramData\$certificateName" -CertStoreLocation Cert:\LocalMachine\my
"@

# Import certificate on remote guest
Invoke-VMScript -VM VM -ScriptText $script -GuestUser administrator -GuestPassword pass2