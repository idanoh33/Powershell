$SourceFile= "C:\temp\Host's_list.txt"
$OutputFile= "C:\temp\IP's_list.txt"

function Get-HostToIP($hostname) {    
    # Ignore error
    $ErrorActionPreference = 'ignore'
    
    # Resolve hostname to IP
    $result = [system.Net.Dns]::GetHostByName($hostname)

    # if found retun name + IP address
    if ($result -ne $null){"$hostname " + $result.AddressList.IPAddressToString}

    # Else return "Hostname was not found"
    else {Write-Output "$hostname was not found"}
}

# Run the function for every line in the txt file and create an outputfile
Get-Content $SourceFile | ForEach-Object {(Get-HostToIP($_))>> $OutputFile}

# Open foutput file on notepad
notepad $OutputFile