#Based on: Serhat AKINCI - Hyper-V MVP - @serhatakinci 
#region Script Parameters
# -----------------------

[CmdletBinding(SupportsShouldProcess=$True)]

Param (
    
    [parameter(
                Mandatory=$false,
                HelpMessage='Hyper-V Cluster name (like HvCluster1 or hvcluster1.domain.corp')]
               
                [string]$Cluster,

    [parameter(
                Mandatory=$false,
                HelpMessage='Standalone Hyper-V Host name(s) (like Host1, Host2, Host3)')]
               
                [array]$VMHost,

    [parameter(
                Mandatory=$false,
                HelpMessage='Reports that shown only highlighted events and alerts')]
               
                [bool]$HighlightsOnly = $false,
    
    [parameter(
                Mandatory=$false,
                HelpMessage='Disk path for HTML reporting file')]
               
                [string]$ReportFilePath = (Get-Location).path,

    [parameter(
                Mandatory=$false,
                HelpMessage='Adds a prefix to the HTML report file name (The default nameprefix is HyperVReport)')]
               
                [string]$ReportFileNamePrefix = "HyperVReport",

    [parameter(
                Mandatory=$false,
                HelpMessage='Adds Timestamp to HTML report file name (The default is $true)')]
               
                [bool]$ReportFileNameTimeStamp = $true,

    [parameter(
                Mandatory=$false,
                HelpMessage='Activates the e-mail sending feature ($true/$false). The default value is "$false"')]
               
                [bool]$SendMail = $false,

    [parameter(
                Mandatory=$false,
                HelpMessage='SMTP Server Address (Like IP address, hostname or FQDN)')]
            
                [string]$SMTPServer,

    [parameter(
                Mandatory=$false,
                HelpMessage='SMTP Server port number (Default 25)')]
            
                [int]$SMTPPort = "25",

    [parameter(
                Mandatory=$false,
                HelpMessage='Recipient e-mail address')]
               
                [array]$MailTo,

    [parameter(
                Mandatory=$false,
                HelpMessage='Sender e-mail address')]
               
                [string]$MailFrom,

    [parameter(
                Mandatory=$false,
                HelpMessage='Sender e-mail address password for SMTP authentication (If needed)')]
               
                [string]$MailFromPassword,

    [parameter(
                Mandatory=$false,
                HelpMessage='SMTP TLS/SSL option ($true/$false). The default value is "$false"')]
            
                [bool]$SMTPServerTLSorSSL = $false,

    [parameter(
                Mandatory=$false,
                HelpMessage='Test Mode will open report automaticly')]
            
                [switch]$TestMode
)

#endregion Script Parameters
#region Functions
#----------------

# Get WMI data
function sGet-Wmi {

    param (

        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
            
        [Parameter(Mandatory = $true)]
        [string]$Namespace,

        [Parameter(Mandatory = $true)]
        [string]$Class,

        [Parameter(Mandatory = $false)]
        $Property,

        [Parameter(Mandatory = $false)]
        $Filter,

        [Parameter(Mandatory = $false)]
        [switch]$AI

    )
    
    # Base string
    $wmiCommand = "gwmi -ComputerName $ComputerName -Namespace $Namespace -Class $Class -ErrorAction Stop"

    # If available, add Filter parameter
    if ($Filter)
    {
        # $Filter = ($Filter -join ',').ToString()
        $Filter = [char]34 + $Filter + [char]34
        $wmiCommand += " -Filter $Filter"
    }

    # If available, add Property parameter
    if ($Property)
    {
        $Property = ($Property -join ',').ToString()
        $wmiCommand += " -Property $Property"
    }

    # If available, Authentication and Impersonation
    if ($AI)
    {
        $wmiCommand += " -Authentication PacketPrivacy -Impersonation Impersonate"
    }

    # Try to connect
    $ResultCode = "1"
    Try
    {
        # $wmiCommand
        $wmiResult = iex $wmiCommand
    }
    Catch
    {
        $wmiResult = $_.Exception.Message
        $ResultCode = "0"
    }
    
    # If wmiResult is null
    if ($wmiResult -eq $null)
    {
        $wmiResult = "Result is null"
        $ResultCode = "2"
    }

    Return $wmiResult, $ResultCode
}

# Write Log
Function sPrint {

    param( 
        
        [byte]$Type=1,

        [string]$Message,
        
        [bool]$WriteToLogFile
        
    )

    $TimeStamp = Get-Date -Format "dd.MMM.yyyy HH:mm:ss"
    $Time = Get-Date -Format "HH:mm:ss"

    if ($Type -eq 1)
    {
        Write-Host "[INFO]    - $Time - $Message" -ForegroundColor Green

        if (($WriteToLogFile) -and ($Logging))
        {
            Add-Content -Path $LogFile -Value "[INFO]    - $TimeStamp - $Message"
        }
    }
    elseif ($Type -eq 2)
    {
        Write-Host "[WARNING] - $Time - $Message" -ForegroundColor Yellow

        if (($WriteToLogFile) -and ($Logging))
        {
            Add-Content -Path $LogFile -Value "[WARNING] - $TimeStamp - $Message"
        }
    }
    elseif ($Type -eq 5)
    {
        if (($WriteToLogFile) -and ($Logging))
        {
            Add-Content -Path $LogFile -Value "[DEBUG]   - $TimeStamp - $Message"
        }
    }
        elseif ($Type -eq 6)
    {
        if (($WriteToLogFile) -and ($Logging))
        {
            Add-Content -Path $LogFile -Value ""
        }
    }
    elseif ($Type -eq 0)
    {
        Write-Host "[ERROR]   - $Time - $Message" -ForegroundColor Red

        if (($WriteToLogFile) -and ($Logging))
        {
            Add-Content -Path $LogFile -Value "[ERROR]   - $TimeStamp - $Message"
        }
    }
    else
    {
        Write-Host "[UNKNOWN] - $Time - $Message" -ForegroundColor Gray

        if (($WriteToLogFile) -and ($Logging))
        {
            Add-Content -Path $LogFile -Value "[UNKNOWN] - $TimeStamp - $Message"
        }
    }
}

# Convert Volume Size to KB/MB/GB/TB
Function sConvert-Size {

    param (
 
	# Disk or Volume Space
	[Parameter(Mandatory = $false)]
	$DiskVolumeSpace,

	# Disk or Volume Space Input Unit
	[Parameter(Mandatory = $true)]
	[string]$DiskVolumeSpaceUnit
 
    )

    if ($DiskVolumeSpaceUnit -eq "byte") # byte input
    {
        if (($DiskVolumeSpace -ge "1024") -and ($DiskVolumeSpace -lt "1048576"))
        {
            $DiskVolumeSpace =  [math]::round(($DiskVolumeSpace/1024))
            $DiskVolumeSpaceUnit = "KB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1048576") -and ($DiskVolumeSpace -lt "1073741824"))
        {
            $DiskVolumeSpace =  [math]::round(($DiskVolumeSpace/1024/1024))

            $DiskVolumeSpaceUnit = "MB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1073741824") -and ($DiskVolumeSpace -lt "1099511627776"))
        {
            $DiskVolumeSpace =  "{0:N1}" -f ($DiskVolumeSpace/1024/1024/1024)
            $DiskVolumeSpaceUnit = "GB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1099511627776") -and ($DiskVolumeSpace -lt "1125899906842624"))
        {
            $DiskVolumeSpace =  "{0:N2}" -f ($DiskVolumeSpace/1024/1024/1024/1024)
            $DiskVolumeSpaceUnit = "TB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif ($DiskVolumeSpace -eq $null)
        {
            $DiskVolumeSpace =  "N/A"
            $DiskVolumeSpaceUnit = "-"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        else
        {
            $DiskVolumeSpace =  $DiskVolumeSpace
            $DiskVolumeSpaceUnit = "Byte"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }    
    }
    elseif ($DiskVolumeSpaceUnit -eq "kb") # kb input
    {
        if (($DiskVolumeSpace -ge "1") -and ($DiskVolumeSpace -lt "1024"))
        {
            $DiskVolumeSpace =  $DiskVolumeSpace
            $DiskVolumeSpaceUnit = "KB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1024") -and ($DiskVolumeSpace -lt "1048576"))
        {
            $DiskVolumeSpace =  ($DiskVolumeSpace/1024)
            $DiskVolumeSpaceUnit = "MB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1048576") -and ($DiskVolumeSpace -lt "1073741824"))
        {
            $DiskVolumeSpace =  "{0:N1}" -f ($DiskVolumeSpace/1024/1024)
            $DiskVolumeSpaceUnit = "GB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1073741824") -and ($DiskVolumeSpace -lt "1099511627776"))
        {
            $DiskVolumeSpace =  "{0:N2}" -f ($DiskVolumeSpace/1024/1024/1024)
            $DiskVolumeSpaceUnit = "TB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif ($DiskVolumeSpace -eq $null)
        {
            $DiskVolumeSpace =  "N/A"
            $DiskVolumeSpaceUnit = "-"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        else
        {
            $DiskVolumeSpace =  $DiskVolumeSpace
            $DiskVolumeSpaceUnit = "KB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }    
    }
    elseif ($DiskVolumeSpaceUnit -eq "mb") # mb input
    {
        if (($DiskVolumeSpace -ge "1") -and ($DiskVolumeSpace -lt "1024"))
        {
            $DiskVolumeSpace =  $DiskVolumeSpace
            $DiskVolumeSpaceUnit = "MB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1024") -and ($DiskVolumeSpace -lt "1048576"))
        {
            $DiskVolumeSpace =  "{0:N1}" -f ($DiskVolumeSpace/1024)
            $DiskVolumeSpaceUnit = "GB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif (($DiskVolumeSpace -ge "1048576") -and ($DiskVolumeSpace -lt "1073741824"))
        {
            $DiskVolumeSpace =  "{0:N2}" -f ($DiskVolumeSpace/1024/1024)
            $DiskVolumeSpaceUnit = "TB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        elseif ($DiskVolumeSpace -eq $null)
        {
            $DiskVolumeSpace =  "N/A"
            $DiskVolumeSpaceUnit = "-"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }
        else
        {
            $DiskVolumeSpace =  $DiskVolumeSpace
            $DiskVolumeSpaceUnit = "MB"
            return $DiskVolumeSpace, $DiskVolumeSpaceUnit
        }    
    }
    else
    {
        return "Unknown Parameter"
    }
}

# Convert BusType Value to BusType Name
Function sConvert-BusTypeName {
    
    Param ([Byte] $BusTypeValue)
    
    if ($BusTypeValue -eq 1){$Result = "SCSI"}
    elseif ($busTypeValue -eq 2){$Result = "ATAPI"}
    elseif ($busTypeValue -eq 3){$Result = "ATA"}
    elseif ($busTypeValue -eq 4){$Result = "IEEE 1394"}
    elseif ($busTypeValue -eq 5){$Result = "SSA"}
    elseif ($busTypeValue -eq 6){$Result = "FC"}
    elseif ($busTypeValue -eq 7){$Result = "USB"}
    elseif ($busTypeValue -eq 8){$Result = "RAID"}
    elseif ($busTypeValue -eq 9){$Result = "iSCSI"}
    elseif ($busTypeValue -eq 10){$Result = "SAS"}
    elseif ($busTypeValue -eq 11){$Result = "SATA"}
    elseif ($busTypeValue -eq 12){$Result = "SD"}
    elseif ($busTypeValue -eq 13){$Result = "SAS"}
    elseif ($busTypeValue -eq 14){$Result = "Virtual"}
    elseif ($busTypeValue -eq 15){$Result = "FB Virtual"}
    elseif ($busTypeValue -eq 16){$Result = "Storage Spaces"}
    elseif ($busTypeValue -eq 17){$Result = "NVMe"}
    else {$Result = "Unknown"}

    Return $Result
}

# Convert Cluster Disk State Value to Name
Function sConvert-ClusterDiskState {
    
    Param ([Byte] $StateValue)

    if ($StateValue -eq 0){$Result = "Inherited",$stateBgColors[5],$stateWordColors[5]}
    elseif ($StateValue -eq 1){$Result = "Initializing",$stateBgColors[4],$stateWordColors[4]}
    elseif ($StateValue -eq 2){$Result = "Online",$stateBgColors[1],$stateWordColors[1]}
    elseif ($StateValue -eq 3){$Result = "Offline",$stateBgColors[2],$stateWordColors[2]}
    elseif ($StateValue -eq 4){$Result = "Failed",$stateBgColors[3],$stateWordColors[3]}
    elseif ($StateValue -eq 127){$Result = "Offline",$stateBgColors[2],$stateWordColors[2]}
    elseif ($StateValue -eq 128){$Result = "Pending",$stateBgColors[4],$stateWordColors[4]}
    elseif ($StateValue -eq 129){$Result = "Online Pending",$stateBgColors[4],$stateWordColors[4]}
    elseif ($StateValue -eq 130){$Result = "Offline Pending",$stateBgColors[4],$stateWordColors[4]}  
    else {$Result = "Unknown",$stateBgColors[5],$stateWordColors[5]} # Including "-1" state

    Return $Result
}

# Convert BusType Value to BusType Name
Function sConvert-DiskPartitionStyle {
    
    Param ([Byte] $PartitionStyleValue)
    
    if ($PartitionStyleValue -eq 1)
    {
        $Result = "MBR"
    }
    elseif ($PartitionStyleValue -eq 2)
    {
        $Result = "GPT"
    }
    else 
    {
        $Result = "Unknown"
    }

    Return $Result
}

# Generate Volume Size Colors
Function sConvert-VolumeSizeColors {

    Param ([Byte] $FreePercent)

    if (($FreePercent -le 10) -and ($FreePercent -gt 5))
    {
        $Result = $stateBgColors[4],$stateBgColors[4],$stateWordColors[4]
    }
    elseif ($FreePercent -le 5)
    {
        $Result = $stateBgColors[3],$stateBgColors[3],$stateWordColors[3]
    }
    else
    {
        $Result = $stateBgColors[0],$stateBgColors[0],"#BDBDBD"
    }

    Return $Result
}

#endregion Functions
#region Variables
#----------------

# Print MSG
sPrint -Type 1 -Message "Started! Hyper-V Reporting Script (Version 1.5)"
Start-Sleep -Seconds 3

# State Colors
[array]$stateBgColors = "", "#ACFA58","#E6E6E6","#FB7171","#FBD95B","#BDD7EE" #0-Null, 1-Online(green), 2-Offline(grey), 3-Failed/Critical(red), 4-Warning(orange), 5-Other(blue)
[array]$stateWordColors = "", "#298A08","#848484","#A40000","#9C6500","#204F7A","#FFFFFF" #0-Null, 1-Online(green), 2-Offline(grey), 3-Failed/Critical(red), 4-Warning(orange), 5-Other(blue), 6-White

# Date and Time
$Date = Get-Date -Format d/MMM/yyyy
$Time = Get-Date -Format "hh:mm:ss tt"

# Log and report file/folder
$FileTimeSuffix = ((Get-Date -Format dMMMyy).ToString()) + "-" + ((get-date -Format hhmmsstt).ToString())

if ($ReportFileNameTimeStamp)
{
    $ReportFile = $ReportFilePath + "\" + $ReportFileNamePrefix + "-" + $FileTimeSuffix + ".html"
}
else
{
    $ReportFile = $ReportFilePath + "\" + $ReportFileNamePrefix + ".html"
}

$LogFile = $ReportFilePath + "\" + "ScriptLog" + ".txt"

# Logging enabled
[bool]$Logging = $True

# HighlightsOnly Mode String
$hlString = $null
if ($HighlightsOnly)
{
    $hlString = "<center><span style=""padding-top:1px;padding-bottom:1px;font-size:12px;background-color:#FBD95B;color:#FFFFFF"">&nbsp;(HighlightsOnly Mode)&nbsp;</span></center>"
    sPrint -Type 1 -Message "HighlightsOnly mode is enabled." -WriteToLogFile $True
}


#endregion Variables

#region Prerequisities Check
#---------------------------

# Log file check and write subject line
if (!(Test-Path -Path $LogFile)) {

    New-Item -Path $LogFile -ItemType file -Force -ErrorAction SilentlyContinue | Out-Null
    
    if (Test-Path -Path $LogFile)
    {
        sPrint -Type 6 -WriteToLogFile $true
        sPrint -Type 5 -Message "----- Start -----" -WriteToLogFile $true
        sPrint -Type 1 -Message "Logging started: $LogFile" -WriteToLogFile $True
        Start-Sleep -Seconds 3
    }
    else
    {
        $Logging = $false
        sPrint -Type 2 -Message "Unable to create the log file. Script will continue without logging..."
        Start-Sleep -Seconds 3
    }
}
else {

    sPrint -Type 6 -WriteToLogFile $true
    sPrint -Type 5 -Message "----- Start -----" -WriteToLogFile $true
    sPrint -Type 1 -Message "Logging started: $LogFile" -WriteToLogFile $true
    Start-Sleep -Seconds 3
}

# Controls for some important prerequisites
if ((!$VMHost) -and (!$Cluster)) {

    sPrint -Type 0 -Message "Hyper-V target parameter is missing. Use -Cluster or -VMHost parameter to define target." -WriteToLogFile $True
    sPrint -Type 2 -Message "For technical information, type: Get-Help .\Get-HyperVReport.ps1 -examples" -WriteToLogFile $True
    sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
    Break
}
if (($VMHost) -and ($Cluster)) {

    sPrint -Type 0 -Message "-Cluster and -VMHost parameters can not be used together." -WriteToLogFile $True
    sPrint -Type 2 -Message "For technical information, type: Get-Help .\Get-HyperVReport.ps1 -examples" -WriteToLogFile $True
    sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
    Break
}

# Controls for runtime environment operating system version, Hyper-V PowerShell and Clustering PowerShell modules
sPrint -Type 1 -Message "Checking prerequisites to run script on the $($env:COMPUTERNAME.ToUpper())..." -WriteToLogFile $True
$osVersion = $null
$osName = $null
$osVersion = sGet-Wmi -ComputerName $env:COMPUTERNAME -Namespace root\Cimv2 -Class Win32_OperatingSystem -Property Version,Caption
    
    if ($osVersion[1] -eq 1)
    {
        $osName = $osVersion[0].Caption
        $osVersion = $osVersion[0].Version
    }
    else
    {
        sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): $($osVersion[0])" -WriteToLogFile $True
        sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
        Break
    }

    if ($osVersion)
    {
        if (($OsVersion -like "6.2*") -or ($OsVersion -like "6.3*"))
        {
            if ($osName -like "Microsoft Windows 8*")
            {
                sPrint -Type 5 -Message "$($env:COMPUTERNAME.ToUpper()): Operating system is supported as script runtime environment." -WriteToLogFile $True

                # Check Hyper-V PowerShell
                if ((Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-Management-PowerShell -Online).State -eq "Enabled")
                {
                    sPrint -Type 5 -Message "$($env:COMPUTERNAME.ToUpper()): Hyper-V PowerShell Module is OK." -WriteToLogFile $True
                }
                else
                {
                    sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Hyper-V PowerShell Module is not found. Please enable manually and run this script again. You can use `"Turn Windows features on or off`" to enable `"Hyper-V Module for Windows PowerShell`"." -WriteToLogFile $True
                    sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                    Break
                }

                # Check Failover Cluster PowerShell
                if ($Cluster)
                {
                    if (Get-Hotfix -ID KB2693643 -ErrorAction SilentlyContinue)
                    {
                        if ((Get-WindowsOptionalFeature -FeatureName RemoteServerAdministrationTools-Features-Clustering -Online).State -eq "Enabled")
                        {
                            sPrint -Type 5 -Message "$($env:COMPUTERNAME.ToUpper()): Failover Clustering PowerShell Module is OK." -WriteToLogFile $True
                        }
                        else
                        {
                            sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Failover Clustering PowerShell Module is not found. Please enable manually and run this script again. You can use `"Turn Windows features on or off`" to enable `"Failover Clustering Tools`"." -WriteToLogFile $True
                            sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                            Break
                        }
                    }
                    else
                    {
                        sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Remote Server Administration Tools (RSAT) is not found. Please download (KB2693643) and install manually and run this script again." -WriteToLogFile $True
                        sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                        Break
                    }
                }
            }
            else
            {
                sPrint -Type 5 -Message "$($env:COMPUTERNAME.ToUpper()): Operating system is supported as script runtime environment." -WriteToLogFile $True

                # Check Hyper-V PowerShell
                if ((Get-WindowsFeature -ComputerName $env:COMPUTERNAME -Name "Hyper-V-PowerShell").Installed)
                {
                    sPrint -Type 5 -Message "$($env:COMPUTERNAME.ToUpper()): Hyper-V PowerShell Module is OK." -WriteToLogFile $True
                }
                else
                {
                    sPrint -Type 2 -Message "$($env:COMPUTERNAME.ToUpper()): Hyper-V PowerShell Module is not found." -WriteToLogFile $True
                    sPrint -Type 2 -Message "$($env:COMPUTERNAME.ToUpper()): Installing Hyper-V PowerShell Module... " -WriteToLogFile $True
                    Start-Sleep -Seconds 3
                    Add-WindowsFeature -Name "Hyper-V-PowerShell" -ErrorAction SilentlyContinue | Out-Null

                    if ((Get-WindowsFeature -ComputerName $env:COMPUTERNAME -Name "Hyper-V-PowerShell").Installed)
                    {
                        sPrint -Type 1 -Message "$($env:COMPUTERNAME.ToUpper()): Hyper-V PowerShell Module is OK." -WriteToLogFile $True
                    }
                    else
                    {
                        sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Hyper-V PowerShell Module could not be installed. Please install it manually." -WriteToLogFile $True
                        sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                        Break
                    }
                }
            
                # Check Failover Cluster PowerShell
                if ($Cluster)
                {
                    if ((Get-WindowsFeature -ComputerName $env:COMPUTERNAME -Name "RSAT-Clustering-PowerShell").Installed)
                    {
                        sPrint -Type 5 -Message "$($env:COMPUTERNAME.ToUpper()): Failover Clustering PowerShell Module is OK." -WriteToLogFile $True
                    }
                    else
                    {
                        sPrint -Type 2 -Message "$($env:COMPUTERNAME.ToUpper()): Failover Clustering PowerShell Module is not found." -WriteToLogFile $True
                        sPrint -Type 2 -Message "$($env:COMPUTERNAME.ToUpper()): Installing Failover Clustering PowerShell Module..." -WriteToLogFile $True
                        Start-Sleep -Seconds 3
                        Add-WindowsFeature -Name "RSAT-Clustering-PowerShell" | Out-Null

                        if ((Get-WindowsFeature -ComputerName $env:COMPUTERNAME -Name "RSAT-Clustering-PowerShell").Installed)
                        {
                            sPrint -Type 1 -Message "$($env:COMPUTERNAME.ToUpper()): Failover Clustering PowerShell Module is OK." -WriteToLogFile $True
                        }
                        else
                        {
                            sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Failover Clustering PowerShell Module could not be installed. Please install it manually." -WriteToLogFile $True
                            sPrint -Type 0 -Message "Script terminated!"
                            Break
                        }
                    }
                }
            }
        }
        else
        {
            sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Incompatible operating system version detected. Supported operating systems are Windows Server 2012 and Windows Server 2012 R2." -WriteToLogFile $True
            sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
            Break
        }    
    }
    else
    {
        sPrint -Type 0 -Message "$($env:COMPUTERNAME.ToUpper()): Could not detect operating system version." -WriteToLogFile $True
        sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
        Break
    }

$Computers = $null
$ClusterName = $null
[array]$VMHosts = $null

#endregion Prerequisities Check
#region HTML Start
#----------------

# HTML Head
$outHtmlStart = "<!DOCTYPE html>
<html>
<head>
<title>Hyper-V Environment Report</title>
<style>
/*Reset CSS*/
html, body, div, span, applet, object, iframe, h1, h2, h3, h4, h5, h6, p, blockquote, pre, a, abbr, acronym, address, big, cite, code, del, dfn, em, img, ins, kbd, q, s, samp,
small, strike, strong, sub, sup, tt, var, b, u, i, center, dl, dt, dd, ol, ul, li, fieldset, form, label, legend, table, caption, tbody, tfoot, thead, tr, th, td,
article, aside, canvas, details, embed, figure, figcaption, footer, header, hgroup, menu, nav, output, ruby, section, summary, 
time, mark, audio, video {margin: 0;padding: 0;border: 0;font-size: 100%;font: inherit;vertical-align: baseline;}
ol, ul {list-style: none;}
blockquote, q {quotes: none;}
blockquote:before, blockquote:after,
q:before, q:after {content: '';content: none;}
table {border-collapse: collapse;border-spacing: 0;}
/*Reset CSS*/

body{
    width:100%;
    min-width:10px;
    font-family: Verdana, sans-serif;
    font-size:14px;
    /*font-weight:300;*/
    line-height:1.5;
    color:#222222;
    background-color:#fcfcfc;
}

p{
    color:222222;
}

strong{
    font-weight:600;
}

h1{
    font-size:30px;
    font-weight:300;
}

h2{
    font-size:20px;
    font-weight:300;
}

#ReportBody{
    width:95%;
    height:500;
    /*border: 1px solid;*/
    margin: 0 auto;
}

.Overview{
    width:60%;
	min-width:180px;
    margin-bottom:30px;
}

.OverviewFrame{
    background:#F9F9F9;
    border: 1px solid #CCCCCC;
}

table#Overview-Table{
    width:100%;
    border: 0px solid #CCCCCC;
    background:#F9F9F9;
    margin-top:0px;
}

table#Overview-Table td {
    padding:0px;
    border: 0px solid #CCCCCC;
    text-align:center;
    vertical-align:middle;
}

.VMHosts{
    width:100%;
    /*height:200px;*/
    /*border: 1px solid;*/
    float:left;
    margin-bottom:30px;
}

table#VMHosts-Table tr:nth-child(odd){
    background:#F9F9F9;
}

table#Disks-Volumes-Table tr:nth-child(odd){
    background:#F9F9F9;
}

.Disks-Volumes{
    width:100%;
    /*height:400px;*/
    /*border: 1px solid;*/
    float:left;
    margin-bottom:30px;
}

.VMs{
    width:50%;
    table-layout: wrap;
    /*height:200px;*/
    /*border: 1px solid;*/
    float:left;
    margin-bottom:22px;
    line-height:1.5;
}

table{
    width:80%;
    min-width:300px;
    table-layout: wrap;
    /*border-collapse: collapse;*/
    border: 1px solid #CCCCCC;
    /*margin-bottom:15px;*/
}

/*Row*/
tr{
    font-size: 12px;
}

/*Column*/
td {
    padding:10px 8px 10px 8px;
    font-size: 12px;
    border: 1px solid #CCCCCC;
    text-align:center;
    vertical-align:middle;
}

/*Table Heading*/
th {
    background: #f3f3f3;
    border: 1px solid #CCCCCC;
    font-size: 14px;
    font-weight:normal;
    padding:12px;
    text-align:center;
    vertical-align:middle;
}
</style>
</head>
<body>
<br><br>
<center><h1>Hyper-V Environment Report</h1></center>
<center><font face=""Verdana,sans-serif"" size=""3"" color=""#222222"">Generated on $($Date) at $($Time)</font></center>
$($hlString)
<br>
<div id=""ReportBody""><!--Start ReportBody-->"

#endregion

#region Gathering Hyper-V Host Information
#-----------------------------------------
    
if ($Cluster) {
    
    if (($Cluster -eq "localhost") -or ($Cluster -eq "127.0.0.1"))
    {
        $ClusterName = $env:COMPUTERNAME
    }
    else
    {
        $ClusterName = $Cluster
    }
    
    $ClusterNodes = $null

    if (Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue)
    {
        $ClusterName = (Get-Cluster -Name $ClusterName).Name
        $hostTableCaption = "Cluster Nodes"
        $volumeTableCaption = "Clustered Disks/Volumes"

        sPrint -Type 1 -Message "$($ClusterName.ToUpper()) is accessible. Gathering Node information..." -WriteToLogFile $True
        Start-Sleep -Seconds 3

        sPrint -Type 1 -Message "Checking prerequisites for Hyper-V Cluster reporting..." -WriteToLogFile $True
        Start-Sleep -Seconds 3

        $clusterNodesData = Get-ClusterNode -Cluster $ClusterName -ErrorAction SilentlyContinue | select Name,State
        $ClusterNodes = ($clusterNodesData | where{$_.State -ne "Down"}).Name
        $downClusterNodes = ($clusterNodesData | where{$_.State -eq "Down"}).Name
        $ovTotalNode = ($clusterNodesData).Count

        if ($downClusterNodes)
        {
            sPrint -Type 0 "Unavailable or down Hyper-V Cluster Node(s): $downClusterNodes" -WriteToLogFile $True
            Start-Sleep -Seconds 3
        }

        if ($ClusterNodes)
        {
            # Checking Cluster Owner Node OS version and Hyper-V role
            $clusterOwnerHostName = sGet-Wmi -ComputerName $ClusterName -Namespace root\Cimv2 -Class  Win32_ComputerSystem -Property Name
            if ($clusterOwnerHostName[1] -eq 1)
            {
                $clusterOwnerHostName = $clusterOwnerHostName[0].Name
            }
            else
            {
                sPrint -Type 0 -Message "$ClusterName`: $($clusterOwnerHostName[0])" -WriteToLogFile $True
                sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                Break
            }
            
            $osVersion = $null
            $getClusterOwnerNode = Get-ClusterNode -Cluster $ClusterName -Name $clusterOwnerHostName
            $osVersion = ($getClusterOwnerNode.MajorVersion).ToString() + "." + ($getClusterOwnerNode.MinorVersion).ToString()
            if (($osVersion -like "6.2") -or ($osVersion -like "6.3"))
            {
                if ((Get-WindowsFeature -ComputerName $clusterOwnerHostName -Name "Hyper-V").Installed)
                {
                    sPrint -Type 5 -Message "Operating system version and Hyper-V role on the cluster owner node is OK." -WriteToLogFile $True
                    $VMHosts = $ClusterNodes

                    # Clear
                    $offlineVmConfigData = $null
                    $ovTotalVm, $ovOfflineVmConfig = 0

                    # Get ClusterResource Data
                    $clusterResourceData = Get-ClusterResource -Cluster $ClusterName

                    # Detect offline Virtual Machine Configuration resources
                    $offlineVmConfigData = $clusterResourceData | where{($_.ResourceType -eq "Virtual Machine Configuration") -and ($_.State -ne "Online")}

                    # For Cluster Overview
                    $ovTotalVm = ($clusterResourceData | where{$_.ResourceType -eq "Virtual Machine"}).Count
                }
                else
                {
                    sPrint -Type 2 -Message "Hyper-V role is not installed on $clusterOwnerHostName." -WriteToLogFile $True
                    sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                    Break
                }
            }
            else
            {
                sPrint -Type 2 -Message "$($ClusterName.ToUpper()): Incompatible operating system version detected. Supported operating systems are Windows Server 2012 and Windows Server 2012 R2." -WriteToLogFile $True
                sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
                Break
            }
        }
        else
        {
            sPrint -Type 0 -Message "$ClusterName`: $($error[0].Exception.Message)" -WriteToLogFile $True
            sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
            Break
        }
    }
    else
    {
        sPrint -Type 0 -Message "$($ClusterName.ToUpper()): $($error[0].Exception.Message)" -WriteToLogFile $True
        sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
        Break
    }
}
    
if ($VMHost) {

    if (($Cluster -eq "localhost") -or ($Cluster -eq "127.0.0.1"))
    {
        $Computers = $env:COMPUTERNAME
    }
    else
    {
        $Computers = $VMHost | Sort-Object -Unique
    }

    [array]$invalidVmHost = $null
    [array]$invalidVmHostMsg = $null
    $hostTableCaption = "Standalone Host(s)"
    $volumeTableCaption = "Local Disks/Volumes"

    sPrint -Type 1 -Message "Checking prerequisites for standalone Hyper-V host(s) reporting..." -WriteToLogFile $True
    Start-Sleep -Seconds 3

    foreach ($ComputerName in $Computers)
    {
        $osVersion = $null
        $osVersion = sGet-Wmi -ComputerName $ComputerName -Namespace root\Cimv2 -Class Win32_OperatingSystem -Property Version
        
        if ($osVersion[1] -eq 1)
        {
            $osVersion = $osVersion[0].Version
        }
        else
        {
            sPrint -Type 0 -Message "$($ComputerName.ToUpper()): $($osVersion[0])" -WriteToLogFile $True
            Start-Sleep -Seconds 3
            $invalidVmHost += $ComputerName
            $invalidVmHostMsg += $osVersion[0]
            Continue
        }

        if ($OsVersion)
        {
            if (($OsVersion -like "6.2*") -or ($OsVersion -like "6.3*"))
            {
                if ((Get-WindowsFeature -ComputerName $ComputerName -Name "Hyper-V").Installed)
                {
                    $checkClusterMember = sGet-Wmi -ComputerName $ComputerName -Namespace root\MSCluster -Class MSCluster_Cluster -Property Name
                    if ($checkClusterMember[1] -eq 1)
                    {
                        sPrint -Type 0 -Message "$($ComputerName.ToUpper()) is a member of a Hyper-V Cluster and didn't included in the VMHost list. Please use -Cluster parameter to report this node." -WriteToLogFile $True
                        $invalidVmHost += $ComputerName
                        $invalidVmHostMsg += "This Node is a member of a cluster. Please use -Cluster parameter to report this node."
                    }
                    else
                    {
                        sPrint -Type 5 -Message "$($ComputerName.ToUpper()): Operating system version and Hyper-V role is OK." -WriteToLogFile $True
                        $VMHosts += $ComputerName
                    }
                }
                else
                {
                    sPrint -Type 0 -Message "$($ComputerName.ToUpper()): Could not be added to the VMHost list because Hyper-V role is not installed." -WriteToLogFile $True
                    $invalidVmHost += $ComputerName
                    $invalidVmHostMsg += "Could not be added to the VMHost list because Hyper-V role is not installed"
                }
            }
            else
            {
                sPrint -Type 0 -Message "$($ComputerName.ToUpper()): Could not be added to the VMHost list because incompatible operating system version detected." -WriteToLogFile $True
                $invalidVmHost += $ComputerName
                $invalidVmHostMsg += "Could not be added to the VMHost list because incompatible operating system version detected"
            }            
        }
        else
        {
            sPrint -Type 0 -Message "$($ComputerName.ToUpper()): Could not be added to the VMHost list because operating system version could not be detected." -WriteToLogFile $True
            $invalidVmHost += $ComputerName
            $invalidVmHostMsg += "Could not be added to the VMHost list because operating system version could not be detected"
        }
    }
}

if (!$VMHosts) {

    sPrint -Type 2 "No valid server for reporting." -WriteToLogFile $True
    sPrint -Type 0 -Message "Script terminated!" -WriteToLogFile $True
    Break
}
else {
    
    if ($Cluster)
    {
        sPrint -Type 1 "Available Hyper-V Cluster Node(s) for reporting: $VMHosts" -WriteToLogFile $True
        Start-Sleep -Seconds 3
    }
    else
    {
        sPrint -Type 1 "Available Hyper-V server(s) for reporting: $VMHosts" -WriteToLogFile $True
        Start-Sleep -Seconds 3
    }
}

# Print MSG
sPrint -Type 1 "Gathering Hyper-V Host information..." -WriteToLogFile $True

# VMHosts-Table Header
    $outVMHostTableStart ="
    <div class=""VMHosts""><!--Start VMHosts Class-->
        <h2>$($hostTableCaption)</h2><br>
        <table id=""VMHosts-Table"">
        <tbody>
            <tr><!--Header Line-->
                <th><p style=""text-align:left;margin-left:-4px"">Name</p></th>
                <th><p>State</p></th>
                <th><p>Uptime</p></th>
                <th><p>Domain</p></th>
                <th><p style=""line-height:1.2"">Total<br>VM</p></th>
                <th><p style=""line-height:1.2"">Active<br>vProcessor</p></th>
                <th><p style=""line-height:1.2"">Logical<br>Processor</p></th>
                <th><p style=""line-height:1.2"">Used<br>Memory</p></th>
                <th><p style=""line-height:1.2"">Free<br>Memory</p></th>
                <th><p style=""line-height:1.2"">Total<br>Memory</p></th>
            </tr>"

# Generate Data Lines
$outVMHostTable = $null
$ovUpNode, $ovTotalLP, $ovTotalMemory, $ovUsedMemory = 0
$ovTotalVProc, $ovTotalVmMemory, $ovUsedVmMemory, $ovUsedVmVHD, $ovTotalVmVHD = 0

foreach ($vmHostItem in $vmHosts) {

    $highL = $false
    $chargerVMHostTable = $null
    $vmHostData = $null
    $vmHostTotalVProc = 0
    $vmHostVpLpRatio = 0
    $vmHostRunningClusVmCount= 0
    $vmHostGet = Get-VMHost -ComputerName $vmHostItem
    $vmHostVMs = Get-VM -ComputerName $vmHostItem
    $vmHostVmCount = $vmHostVMs.Count + ($offlineVmConfigData | where{$_.OwnerNode -eq "$vmHostItem"}).Count
    $vmHostRunningVmCount = ($vmHostVMs | where{$_.State -eq "Running"}).Count
    $vmHostRunningClusVmCount = ($vmHostVMs | where{($_.IsClustered -eq $true) -and ($_.State -eq "Running")}).Count
    $vmHostRunningNonClusVmCount = $vmHostRunningVmCount - $vmHostRunningClusVmCount 
    $vmHostTotalVProc = (($vmHostVMs | where{(($_.State -eq "Running") -or ($_.State -eq "Paused"))}).ProcessorCount | Measure-Object -Sum).Sum
    $vmHostClusVProc = (($vmHostVMs | where{(($_.State -eq "Running") -and ($_.IsClustered -eq $true)) -or (($_.State -eq "Paused") -and ($_.IsClustered -eq $true))}).ProcessorCount | Measure-Object -Sum).Sum
    $vmHostWmiData = Get-WmiObject -ComputerName $vmHostItem -Class Win32_OperatingSystem

    # For Cluster Overview
    $ovTotalVProc = $ovTotalVProc + $vmHostClusVProc

    # State
    if ($Cluster)
    {
        $vmHostState = (Get-ClusterNode -Cluster $ClusterName -Name $vmHostItem).State
    }
    else
    {
        $vmHostState = "Up"  
    }

    # State Colors
    if ($vmHostState -eq "Up")
    {
        $outVmHostState = "Up",$stateBgColors[1],$stateWordColors[1]
    }
    elseif ($vmHostState -eq "Down")
    {
        $outVmHostState = "Down",$stateBgColors[3],$stateWordColors[3]
        $highL = $true
    }
    elseif ($vmHostState -eq "Paused")
    {
        $outVmHostState = "Paused",$stateBgColors[2],$stateWordColors[2]
    }
    elseif ($vmHostState -eq "Joining")
    {
        $outVmHostState = "Joining",$stateBgColors[4],$stateWordColors[4]
    }
    else
    {
        $outVmHostState = "Unknown",$stateBgColors[5],$stateWordColors[5]
        $highL = $true
    }
    
    # Clear
    $TotalUsedMemory = $null
    $TotalFreeMemory = $null
    $TotalVisibleMemory = $null
    $vmHostUptime = $null
    $TotalFreeMemoryPercentage = $null
     
    # Memory Capacty
    $TotalUsedMemory = sConvert-Size -DiskVolumeSpace ($vmHostWmiData.TotalVisibleMemorySize - $vmHostWmiData.FreePhysicalMemory) -DiskVolumeSpaceUnit kb
    $TotalFreeMemory = sConvert-Size -DiskVolumeSpace $vmHostWmiData.FreePhysicalMemory -DiskVolumeSpaceUnit kb
    $TotalVisibleMemory = sConvert-Size -DiskVolumeSpace $vmHostWmiData.TotalVisibleMemorySize -DiskVolumeSpaceUnit kb
    $TotalFreeMemoryPercentage = [math]::round(($vmHostWmiData.FreePhysicalMemory/$vmHostWmiData.TotalVisibleMemorySize)*100)

    # Free Memory Percentage Colors
    # 0 - $totalFreeMemoryBgColor
    # 1 - $TotalFreeMemoryPercentageBgColor
    # 2 - $TotalFreeMemoryPercentageWordColor
    if (($TotalFreeMemoryPercentage -le 10) -and ($TotalFreeMemoryPercentage -gt 5))
    {
        $outVmHostFreeMemoryState = $stateBgColors[4],$stateBgColors[4],$stateWordColors[4]
        $highL = $true
    }
    elseif ($TotalFreeMemoryPercentage -le 5)
    {
        $outVmHostFreeMemoryState = $stateBgColors[3],$stateBgColors[3],$stateWordColors[3]
        $highL = $true
    }
    else
    {
        $outVmHostFreeMemoryState = $stateBgColors[0],$stateBgColors[0],"#BDBDBD"
    }

    # Hostname
    $outVMHostName = ($vmHostGet.ComputerName).ToUpper()
     
    # Uptime
    $vmHostUptime = ([Management.ManagementDateTimeConverter]::ToDateTime($vmHostWmiData.LocalDateTime)) - ([Management.ManagementDateTimeConverter]::ToDateTime($vmHostWmiData.LastBootUpTime))
        if($vmHostUptime.Days -eq "0"){$vmHostUptimeDays = ""}
        else{$vmHostUptimeDays = ($vmHostUptime.Days).ToString() + " <span style=""font-size:10px;color:#BDBDBD"">Days</span> <br>"}
    $vmHostUptime = ($vmHostUptime.Hours).ToString() + ":" + ($vmHostUptime.Minutes).ToString() + ":" + ($vmHostUptime.Seconds).ToString()

    # OS Version
    $vmHostOsVersion = ($vmHostWmiData.Caption).Replace("Microsoft ","")

    # Processor socket and HT state
    $processorData = sGet-Wmi -ComputerName $vmHostItem -Namespace root\CIMv2 -Class Win32_Processor -Property DeviceID,NumberOfCores,NumberOfLogicalProcessors
    if ($processorData[1] -eq 1)
    {
        $socketCount = ($processorData[0] | ForEach-Object {$_.DeviceID} | select-object -unique).Count
        $coreCount = ($processorData[0].NumberOfCores | Measure-Object -Sum).Sum
        $logicalProcCount = ($processorData[0].NumberOfLogicalProcessors | Measure-Object -Sum).Sum

        if ($logicalProcCount -gt $coreCount)
        {
            $htState = "Active"
        }
        Else
        {
            $htState = "Inactive"
        }
    }
    else
    {
        $socketCount = "-"
        $htState = "Unknown"
    }

    $vmHostLpCount = $vmHostGet.LogicalProcessorCount
    if (!$vmHostLpCount)
    {
        $vmHostLpCount = $logicalProcCount
    }

    # For Cluster Overview
    if(($Cluster) -and ($vmHostState -eq "Up"))
    {
        $ovUpNode = $ovUpNode + 1
        $ovTotalLP = $ovTotalLP + $vmHostLpCount
        $ovUsedMemory = $ovUsedMemory + ($vmHostWmiData.TotalVisibleMemorySize - $vmHostWmiData.FreePhysicalMemory)
        $ovTotalMemory = $ovTotalMemory + $vmHostWmiData.TotalVisibleMemorySize
    }

    # LP:VP Ratio
    $vmHostVpLpRatio = ("{0:N2}" -f ($vmHostTotalVProc / $vmHostLpCount)).Replace(".00","")

    # Computer and Processor Manufacturer/Model Info
    $outVmHostComputerInfo = gwmi -ComputerName $vmHostItem -Class Win32_ComputerSystem -Property Manufacturer,Model
    $outVmHostProcModel = (gwmi -ComputerName $vmHostItem -Class Win32_Processor).Name
    if($outVmHostProcModel.count -gt 1)
    {
        $outVmHostProcModel = $outVmHostProcModel[0]
    }
    $outVmHostProcModel = $outVmHostProcModel.Replace("           "," ")

    # Data Line
    $chargerVMHostTable ="
            <tr><!--Data Line-->
                <td><p style=""text-align:left;""><abbr title=""Manufacturer: $($outVmHostComputerInfo.Manufacturer)&#10;Model: $($outVmHostComputerInfo.Model)"">$($outVMHostName) <span style=""font-size:10px;color:orange"">*</span></abbr><br><span style=""font-size:10px;color:#BDBDBD;text-align:left"">$($vmHostOsVersion)</span></p></td>
                <td bgcolor=""$($outVmHostState[1])""><p style=""color:$($outVmHostState[2])"">$($outVmHostState[0])</p></td>
                <td><p>$($vmHostUptimeDays)$($vmHostUptime)</p></td>
                <td><p>$($vmHostGet.FullyQualifiedDomainName)</p></td>
                <td><p style=""line-height:1.2"">$($vmHostVmCount) <br><span style=""font-size:10px;color:#BDBDBD""><abbr title=""Non-Clustered: $($vmHostRunningNonClusVmCount) | Clustered: $($vmHostRunningClusVmCount)"">$($vmHostRunningVmCount) Running <span style=""font-size:10px;color:orange"">*</span></abbr></span></p></td>
                <td><p style=""line-height:1.2"">$($vmHostTotalVProc)<br><abbr title=""Virtual Processors per Logical Processor ratio (VP:LP)""><span style=""font-size:10px;color:#BDBDBD"">$($vmHostVpLpRatio):1 <span style=""font-size:10px;color:orange"">*</span></span></abbr></p></td>
                <td><p style=""line-height:1.2"">$($vmHostLpCount) <br><span style=""font-size:10px;color:#BDBDBD""><abbr title=""Hyper-Threading: $($htState)&#10;Model: $($outVmHostProcModel)"">$($socketCount) Socket <span style=""font-size:10px;color:orange"">*</span></abbr></span></p></td>
                <td><p style=""line-height:1.2"">$(($TotalUsedMemory)[0])<br><span style=""font-size:10px"">$(($TotalUsedMemory)[1])</span></p></td>
                <td bgcolor=""$($outVmHostFreeMemoryState[0])""><p style=""line-height:1.2"">$(($TotalFreeMemory)[0])<br><span style=""font-size:10px"">$(($TotalFreeMemory)[1])</span></p></td>
                <td><p style=""line-height:1.2"">$(($TotalVisibleMemory)[0]) <span style=""font-size:10px"">$(($TotalVisibleMemory)[1])</span> <br><span style=""font-size:10px;background-color:$($outVmHostFreeMemoryState[1]);color:$($outVmHostFreeMemoryState[2])"">&nbsp;~%$($TotalFreeMemoryPercentage) free&nbsp;</span></p></td>
            </tr>"

    # Add to HTML Table
    if ($HighlightsOnly -eq $false)
    {
        # VMHost Output
        $outVMHostTable += $chargerVMHostTable
    }
    elseif (($HighlightsOnly -eq $true) -and ($highL -eq $true))
    {
        # VMHost Output
        $outVMHostTable += $chargerVMHostTable
    }
    else
    {
        # Blank
    }
}

# Add offline or unsupported standalone hosts
if ($invalidVmHost)
{
    [bytle]$numb = 0
    ForEach ($VMhostIN in $invalidVmHost)
    {

    $outVMHostTable +="
            <tr><!--Data Line-->
                <td><p style=""text-align:left;"">$(($VMhostIN).ToUpper())<br><span style=""font-size:10px;color:#BDBDBD;text-align:left"">Operating System Unknown</span></p></td>
                <td bgcolor=""$($stateBgColors[2])""><p style=""color:$($stateWordColors[2])"">Unaccessible</p></td>
                <td colspan=""8""><p style=""text-align:left; color:#BDBDBD"">$($invalidVmHostMsg[$numb])</p></td>
            </tr>"
    
    $numb = $numb + 1

    }
}

# Add down cluster nodes
if ($downClusterNodes)
{
    ForEach ($downClusterNode in $downClusterNodes)
    {

    $outVMHostTable +="
            <tr><!--Data Line-->
                <td><p style=""text-align:left;"">$($downClusterNode)<br><span style=""font-size:10px;color:#BDBDBD;text-align:left"">Operating System Unknown</span></p></td>
                <td bgcolor=""$($stateBgColors[3])""><p style=""color:$($stateWordColors[3])"">Down</p></td>
                <td colspan=""8""><p style=""text-align:left; color:#BDBDBD"">Hyper-V Cluster node is down or unavailable</p></td>
            </tr>"
    }
}

if (($outVMHostTable -eq $null) -and ($downClusterNodes -eq $null))
{
    if ($Cluster)
    {
        $outVMHostTable +="
            <tr><!--Data Line-->
                <td colspan=""10""><p style=""text-align:center""><span style=""padding-top:1px;padding-bottom:1px;background-color:#ACFA58;color:#298A08"">&nbsp;&nbsp;All Hyper-V Cluster Nodes are healthy&nbsp;&nbsp;</span></p></td>
            </tr>"
    }
    else
    {
        $outVMHostTable +="
            <tr><!--Data Line-->
                <td colspan=""10""><p style=""text-align:center""><span style=""padding-top:1px;padding-bottom:1px;background-color:#ACFA58;color:#298A08"">&nbsp;All Standalone Hyper-V Hosts are healthy&nbsp;&nbsp;</span></p></td>
            </tr>"
    }
}

    # End VMHosts-Table
    $outVMHostTableEnd ="
        </tbody>
        </table>
    </div><!--End VMHosts Class-->"

#endregion

#region Gathering VM Information
#-------------------------------

# Print MSG
sPrint -Type 1 "Gathering Virtual Machine information..." -WriteToLogFile $True

$outVMTableStart = "
    <div class=""VMs""><!--Start VM Class-->
        <h2>Virtual Machines</h2><br>
        <table>
        <tbody>
            <tr><!--Header Line-->
                <th><p style=""text-align:left;margin-left:-4px"">Name</p></th>
                <th><p>State</p></th>
                <th><p>Uptime</p></th>
            </tr>"

# Generate Data Lines
$outVmTable = $null
$cntVM = 0
$vmNoInTable = 0
$ovRunningVm = 0
$ovPausedVm = 0

# Active VHD Array
$activeVhds = @()

ForEach ($VMHostItem in $VMHost) {
    
    $getVMerr = $null
    if ($Cluster)
        {
        $clusterNodes = Get-ClusterNode;
        $VMS= ForEach($item in $clusterNodes) {Get-VM -ComputerName $item.Name; }
        $vNetworkAdapters = $VMS | Get-VMNetworkAdapter -ErrorAction SilentlyContinue
        }
    else
        {
        $VMs = Get-VM -ComputerName $VMHostItem -ErrorVariable getVMerr -ErrorAction SilentlyContinue
        $vNetworkAdapters = Get-VM -ComputerName $VMHostItem | Get-VMNetworkAdapter -ErrorAction SilentlyContinue
        }
    
    

    # Offline Virtual Machine Configuration resources on this node
    if ($Cluster)
    {
        $offlineVmConfigs = $null
        $offlineVmConfigs = $offlineVmConfigData | where{$_.OwnerNode -eq "$VMHostItem"}
    }

    # If Get-VM is success
    if ($VMs)
    {
        $cntVM = $cntVM + 1
        
        foreach ($VM in $VMs)
        {
            $highL = $false
            $chargerVmTable = $null
            $chargerVmMemoryTable = $null
            $outVmReplReplicaServer = $null
            $outVmReplFrequency = $null

            # Table TR Color
            if([bool]!($vmNoInTable%2))
            {
               #Even or Zero
               $vmTableTrBgColor = ""
            }
            else
            {
               #Odd
               $vmTableTrBgColor = "#F9F9F9"
            }

            # Name and Config Path
            $outVmName = $VM.VMName
            $outVmPath = $VM.ConfigurationLocation

            # Generation and Version
            if (!$VM.Generation -and !$VM.Version)
            {
                $outVmGenVer = ""
            }
            else
            {
                $outVmGenVer = ""
            }

            # VM State
            $outVmState = $VM.State

            # IsClustered Yes or No
            if ($VM.IsClustered -eq $True)
            {
                # For Cluster Overview (Total and Used VmMemory)
                if ($VM.State -eq "Running")
                {
                    $ovRunningVm = $ovRunningVm + 1
                }

                if ($VM.State -eq "Paused")
                {
                    $ovPausedVm = $ovPausedVm + 1
                }

                if(($VM.State -eq "Running") -or ($VM.State -eq "Paused"))
                {
                    if(!$VM.DynamicMemoryEnabled)
                    {
                        $ovTotalVmMemory = $ovTotalVmMemory + $VM.MemoryStartup
                    }
                    else
                    {
                        $ovTotalVmMemory = $ovTotalVmMemory + $VM.MemoryMaximum
                    }

                    $ovUsedVmMemory = $ovUsedVmMemory + $VM.MemoryAssigned
                } 

                # Clustered VM State
                $getClusVMerr = $null
                $outVmIsClustered = "Yes"
                $clusVmState = (Get-ClusterResource -Cluster $ClusterName -VMId $VM.VMId -ErrorAction SilentlyContinue -ErrorVariable getClusVMerr).State

                if ($getClusVMerr)
                {
                    $outVmState = "Unknown"
                }
                elseif ($clusVmState -eq "Online")
                {
                    if ($VM.State -eq "Paused")
                    {
                        $outVmState = "Paused"
                    }
                    else
                    {
                        $outVmState = "Running"
                    } 
                }
                elseif ($clusVmState -eq "Offline")
                {
                    if ($VM.State -eq "Saved")
                    {
                        $outVmState = "Saved"
                    }
                    else
                    {
                        $outVmState = "Off"
                    } 
                }
                else
                {
                    $outVmState = $clusVmState
                }
            }
            else
            {
                $outVmIsClustered = "No"
            }

            # VM State Color
            if ($outVmState -eq "Running")
            {
                $vmStateBgColor = $stateBgColors[1]
                $vmStateWordColor = $stateWordColors[1]
            }
            Elseif ($outVmState -eq "Off")
            {
                $vmStateBgColor = $stateBgColors[2]
                $vmStateWordColor = $stateWordColors[2]
            }
            Elseif (($outVmState -match "Critical") -or ($outVmState -match "Failed"))
            {
                $vmStateBgColor = $stateBgColors[3]
                $vmStateWordColor = $stateWordColors[3]
                $highL = $true
            }
            Elseif (($outVmState -eq "Paused") -or ($outVmState -eq "Saved"))
            {
                $vmStateBgColor = $stateBgColors[4]
                $vmStateWordColor = $stateWordColors[4]
                $highL = $true
            }
            else
            {
                $vmStateBgColor = $stateBgColors[5]
                $vmStateWordColor = $stateWordColors[5]
            }
        
            # Uptime
            if ($VM.Uptime -eq "00:00:00")
            {
                $outVmUptimeDays = $null
                $outVmUptime = "Stopped"
            }
            else
            {
                $outVmUptimeDays = (($VM.Uptime).Days).ToString()
                    if ($outVmUptimeDays -eq "0")
                    {
                        $outVmUptimeDays = $null
                    }
                    else
                    {
                        $outVmUptimeDays = $outVmUptimeDays + " <span style=""font-size:10px;color:#BDBDBD"">Days</span> <br>"
                    }
                $outVmUptime = (($VM.Uptime).Hours).ToString() + ":" + (($VM.Uptime).Minutes).ToString() + ":" + (($VM.Uptime).Seconds).ToString()
            }

            # Owner Host
            $outVmHost = ($VM.ComputerName).ToUpper()

            # vCPU
            $outVmCPU = $VM.ProcessorCount
        
            # IS State, Version and Color
            if ($VM.IntegrationServicesState -eq "Up to date")
            {
                $outVmIs = "UpToDate"
                $outVmIsVer = $VM.IntegrationServicesVersion
                $vmIsStateBgColor = ""
                $vmIsStateWordColor = ""
            }
            elseif ($VM.IntegrationServicesState -eq "Update required")
            {
                $outVmIs = "UpdateRequired"
                $outVmIsVer = $VM.IntegrationServicesVersion
                $vmIsStateBgColor = $stateBgColors[4]
                $vmIsStateWordColor = $stateWordColors[4]
                $highL = $true
            }
            else
            {
                if ($vm.State -eq "Running")
                {
                    if ($VM.IntegrationServicesVersion -eq "6.2.9200.16433")
                    {
                        $outVmIs = "UpToDate"
                        $outVmIsVer = $VM.IntegrationServicesVersion
                        $vmIsStateBgColor = ""
                        $vmIsStateWordColor = ""
                    }
                    elseif ($VM.IntegrationServicesVersion -eq $null)
                    {
                        $outVmIs = "NotDetected"
                        $outVmIsVer = "NotDetected"
                        $vmIsStateBgColor = ""
                        $vmIsStateWordColor = ""
                    }
                    else
                    {
                        $outVmIs = "MayBeRequired"
                        $outVmIsVer = $VM.IntegrationServicesVersion
                        $vmIsStateBgColor = ""
                        $vmIsStateWordColor = ""
                    }
                }
                else
                {
                    $outVmIs = "NotDetected"
                    $outVmIsVer = "NotDetected"
                    $vmIsStateBgColor = ""
                    $vmIsStateWordColor = ""
                }
            }
        
            # Checkpoint State and Color
            if ($VM.ParentSnapshotId)
            {
                $outVmChekpoint = "Yes"
                $vmCheckpointBgColor = $stateBgColors[4]
                $vmCheckpointWordColor = $stateWordColors[4]
                $vmChekpointCount = (Get-VMSnapshot -ComputerName $VM.ComputerName -VMName $VM.Name).Count
                $outVmChekpointCount = ($vmChekpointCount).ToString() + " Checkpoint(s)"
                $highL = $true
            }
            else
            {
                $outVmChekpoint = "No"
                $vmCheckpointBgColor = ""
                $vmCheckpointWordColor = ""
                $outVmChekpointCount = $null
            }

            # Replication
            if ($VM.ReplicationState -ne "Disabled")
            {
                $outVmRepl = $null
                $chargerVmRepl1 = $null
                $chargerVmRepl2 = $null
                $getVmReplication = Get-VMReplication -ComputerName $VM.ComputerName -VMName $VM.Name

                foreach ($getVmReplItem in $getVmReplication)
                {
                    if ($getVmReplItem.Mode -eq "Primary") #Primary
                    {
                        $outVmReplType = "Primary"
                        $outVmReplServer = "ReplicaServer: $($getVmReplItem.ReplicaServer) &#10;"
                    }
                    elseif ($getVmReplItem.Mode -eq "Replica" -and $getVmReplItem.RelationshipType -eq "Simple") #Replica
                    {
                        $outVmReplType = "Replica"
                        $outVmReplServer = "PrimaryServer: $($getVmReplItem.PrimaryServer) &#10;"
                    }
                    elseif ($getVmReplItem.Mode -eq "Replica" -and $getVmReplItem.RelationshipType -eq "Extended") #Replica to Extended
                    {
                        $outVmReplType = "Extended"
                        $outVmReplServer = "ReplicaServer: $($getVmReplItem.ReplicaServer) &#10;"
                    }
                    elseif ($getVmReplItem.Mode -eq "ExtendedReplica") #Extended
                    {
                        $outVmReplType = "Extended"
                        $outVmReplServer = "PrimaryServer: $($getVmReplItem.PrimaryServer) &#10;"
                    }
                    elseif ($getVmReplItem.Mode -eq "Replica")
                    {
                        $outVmReplType = "Replica"
                        $outVmReplServer = "PrimaryServer: $($getVmReplItem.PrimaryServer) &#10;"
                    }
                    else
                    {
                        $outVmReplType = $getVmReplItem.Mode
                        $outVmReplServer = "PrimaryServer/ReplicaServer &#10;"
                    }

                    $outVmReplHealth = "Health: $($getVmReplItem.Health) &#10;"
                    $outVmReplMode = "Mode: $($getVmReplItem.Mode) &#10;"
                    $outVmLastReplTime = "LastReplTime: $($getVmReplItem.LastReplicationTime) &#10;"
                    $outVMReplState = "ReplState: $($getVmReplItem.State)"

                    # Repl Frequency
                    if ($getVmReplItem.FrequencySec -gt 30)
                    {
                        $outVmReplFrequency = "Frequency: " + (($getVmReplItem.FrequencySec)/60) + " Min &#10;"
                    }
                    elseif ($getVmReplItem.FrequencySec -le 30 -and $getVmReplItem.FrequencySec -gt 0)
                    {
                        $outVmReplFrequency = "Frequency: " + ($getVmReplItem.FrequencySec) + " Sec &#10;"
                    }
                    elseif($OsVersion -like "6.2*")
                    {
                        $outVmReplFrequency = "Frequency: 5 Min &#10;"
                    }
                    else
                    {
                        $outVmReplFrequency = "Frequency: &#10;"
                    }

                    # Repl Health Colors
                    if ($getVmReplItem.Health -eq "Normal")
                    {
                        $vmReplHealthBgColor = $stateBgColors[1]
                        $vmReplHealthWordColor = $stateWordColors[1]
                    }
                    elseif ($getVmReplItem.Health -eq "Warning")
                    {
                        $vmReplHealthBgColor = $stateBgColors[4]
                        $vmReplHealthWordColor = $stateWordColors[4]
                        $highL = $true
                    }
                    elseif ($getVmReplItem.Health -eq "Critical")
                    {
                        $vmReplHealthBgColor = $stateBgColors[3]
                        $vmReplHealthWordColor = $stateWordColors[3]
                        $highL = $true
                    }
                    else
                    {
                        $vmReplHealthBgColor = $stateBgColors[5]
                        $vmReplHealthWordColor = $stateWordColors[5]
                        $highL = $true
                    }

                    if ($getVmReplItem.Mode -eq "Replica" -and $getVmReplItem.RelationshipType -eq "Extended")
                    {
                        $chargerVmRepl2 = "<p style=""margin-top:8px;background-color:$($vmReplHealthBgColor);color:$($vmReplHealthWordColor)""><abbr title=""$($outVmReplHealth)$($outVmReplMode)$($outVmReplServer)$($outVmReplFrequency)$($outVmLastReplTime)$($outVMReplState)"">$($outVmReplType)</abbr></p>"
                    }
                    else
                    {
                        $chargerVmRepl1 = "<p style=""background-color:$($vmReplHealthBgColor);color:$($vmReplHealthWordColor)""><abbr title=""$($outVmReplHealth)$($outVmReplMode)$($outVmReplServer)$($outVmReplFrequency)$($outVmLastReplTime)$($outVMReplState)"">$($outVmReplType)</abbr></p>"
                    }
                }

                $outVmRepl = $chargerVmRepl1 + $chargerVmRepl2
            }
            else
            {
                $outVmRepl = "<p>N/E</p>"
                $vmReplHealthBgColor = ""
                $vmReplHealthWordColor = ""
            }

            # Network Adapter
            if ($vNetworkAdapters | where{$_.VMId -eq $VM.VMId})
            {
                $vmNetAdapterCount = 1
                $vmNetAdapters = $null
                $outVmNetAdapter = $null
                $vmNetAdapters = $vNetworkAdapters | where{$_.VMId -eq $VM.VMId}

                foreach ($vmNetAdapter in $vmNetAdapters)
                {
                    # Type
                    if (!$vmNetAdapter.IsLegacy)
                    {
                        $outVmNetAdapterName = "Synthetic Network Adapter"
                    }
                    else
                    {
                        $outVmNetAdapterName = "Legacy Network Adapter"
                    }
                    
                    # IP
                    if ($vmNetAdapter.IPAddresses)
                    {
                        if ($vmNetAdapter.IPAddresses.Count -gt 1)
                        {
                            $outVmNetAdapterIP = ($vmNetAdapter.IPAddresses -join ', ').ToString()
                        }
                        else
                        {
                            $outVmNetAdapterIP = $vmNetAdapter.IPAddresses
                        }
                    }
                    else
                    {
                        $outVmNetAdapterIP = "Unable to get ip address information"
                    }

                    # MAC
                    if ($vmNetAdapter.MacAddress)
                    {
                        $outVmNetAdapterMacAddress = "MAC Address: $($vmNetAdapter.MacAddress)"
                    }
                    else
                    {
                        $outVmNetAdapterMacAddress = "MAC Address: Null"
                    }
                    
                    if ($vmNetAdapter.DynamicMacAddressEnabled)
                    {
                        $outVmNetAdapterMacAddressType = "MAC Type: Dynamic"
                    }
                    else
                    {
                        $outVmNetAdapterMacAddressType = "MAC Type: Static"
                    }

                    # Connection
                    if ($vmNetAdapter.Connected)
                    {
                        $outVmNetAdapterConnection = "Connected"
                        $outVmNetAdapterSwitch = "Virtual Switch Name: $($vmNetAdapter.SwitchName)"

                    }
                    else
                    {
                        $outVmNetAdapterConnection = "Not connected"
                        $outVmNetAdapterSwitch = "Not connected to a switch"
                    }

                    # VLAN
                    if (($vmNetAdapter.VlanSetting.AccessVlanId -eq 0) -or ($vmNetAdapter.VlanSetting.AccessVlanId -eq $null))
                    {
                        $outVmNetAdapterVlan = "VLAN: Disabled"
                    }
                    else
                    {
                        $outVmNetAdapterVlan = "VLAN: Enabled, ID $($vmNetAdapter.VlanSetting.AccessVlanId)"
                    }

                    # Other
                    $outVmNetAdapterDhcpGuard = "DHCP Guard: $($vmNetAdapter.DhcpGuard)"
                    $outVmNetAdapterRouterGuard = "Router Guard: $($vmNetAdapter.RouterGuard)"
                    $outVmNetAdapterPortMirroringMode = "Port Mirroring: $($vmNetAdapter.PortMirroringMode)"
                    
                    if ($vmNetAdapter.ClusterMonitored)
                    {
                        $outVmNetAdapterClusterMonitored = "Protected Network: On"
                    }
                    else
                    {
                        if ($OsVersion -like "6.2*")
                        {
                            $outVmNetAdapterClusterMonitored = "Protected Network: N/A"
                        }
                        else
                        {
                            $outVmNetAdapterClusterMonitored = "Protected Network: Off"
                        }
                    }
                    
                    # Write
                    if ($vmNetAdapterCount -eq 1)
                    {
                        $chargerVmNetAdapter = "<p style=""text-align:left""><abbr title=""$($outVmNetAdapterIP)"">$($outVmNetAdapterName)<span style=""font-size:10px;color:orange""> *</span><br><span style=""font-size:10px;color:#BDBDBD"">&#10148; <abbr title=""$($outVmNetAdapterSwitch)"">$($outVmNetAdapterConnection)</abbr> | <abbr title=""$($outVmNetAdapterVlan)"">VLAN</abbr> | <abbr title=""$($outVmNetAdapterMacAddress) &#10;$outVmNetAdapterMacAddressType &#10;$outVmNetAdapterDhcpGuard &#10;$outVmNetAdapterRouterGuard &#10;$outVmNetAdapterPortMirroringMode &#10;$outVmNetAdapterClusterMonitored"">Advanced</abbr></span></p>"
                    }
                    else
                    {
                        $chargerVmNetAdapter = "<p style=""text-align:left;margin-top:6px""><abbr title=""$($outVmNetAdapterIP)"">$($outVmNetAdapterName)<span style=""font-size:10px;color:orange""> *</span><br><span style=""font-size:10px;color:#BDBDBD"">&#10148; <abbr title=""$($outVmNetAdapterSwitch)"">$($outVmNetAdapterConnection)</abbr> | <abbr title=""$($outVmNetAdapterVlan)"">VLAN</abbr> | <abbr title=""$($outVmNetAdapterMacAddress) &#10;$outVmNetAdapterMacAddressType &#10;$outVmNetAdapterDhcpGuard &#10;$outVmNetAdapterRouterGuard &#10;$outVmNetAdapterPortMirroringMode &#10;$outVmNetAdapterClusterMonitored"">Advanced</abbr></span></p>"
                    }

                    $outVmNetAdapter += $chargerVmNetAdapter

                    $vmNetAdapterCount = $vmNetAdapterCount + 1
                }
            }
            else
            {
                        $outVmNetAdapter = "<p style=""text-align:left"">No Network Adapter</p>"
            }

            # Get Disks
            $vmDiskOutput = $null
            $rowSpanCount = 0
            $getVhdErr = $null
            $vmDisks = Get-VHD -ComputerName $VMHostItem -VMId $vm.VMId -ErrorAction SilentlyContinue -ErrorVariable getVhdErr
            
            if ($getVhdErr)
            {
                if ($rowSpanCount -eq 0)
                {
			        $vmDiskOutput +=""
                    $highL = $true
                    $rowSpanCount = $rowSpanCount + 1
                }
                else
                {
                    $vmDiskOutput +="
            <tr style=""background:$($vmTableTrBgColor)"">
                <td Style=""border-top:2px dotted #ccc""><p style=""text-align:left""><span style=""background-color:$($stateBgColors[3]);color:$($stateWordColors[3])"">&nbsp;$($getVhdErr.count) VHD file(s) missing&nbsp;</span> <br><span style=""font-size:10px;color:#BDBDBD"">&#10148; CurrentFileSize N/A (MaximumDiskSize N/A) <br><span style=""font-size:10px;color:#BDBDBD"">&#10148; VhdType N/A | ControllerType N/A | Fragmentation N/A</span></p></td>
            </tr>"
                    $highL = $true
                    $rowSpanCount = $rowSpanCount + 1
                }
            }

            $vmPTDisks = Get-VMHardDiskDrive -ComputerName $VMHostItem -VMname $vm.name | where{$_.Path -like "Disk*"}

            # Pass-through
            $vmPTDiskNo = 1
            if ($vmPTDisks)
            {
                foreach ($vmPTDisk in $vmPTDisks)
                {
                    if ($rowSpanCount -eq 0)
                    {
			            $vmDiskOutput +=""
                        $rowSpanCount = $rowSpanCount + 1
                        $vmPTDiskNo = $vmPTDiskNo + 1
                    }
                    else
                    {
                        $vmDiskOutput=""
                        $rowSpanCount = $rowSpanCount + 1
                        $vmPTDiskNo = $vmPTDiskNo + 1
                    }               
                }
            }

            # VHD
            if ($vmDisks -eq $null)
            {
                 $vmDiskOutput = "
                <td rowspan=""0""><p style=""text-align:left""><span style=""background-color:$($stateBgColors[4]);color:$($stateWordColors[4])"">&nbsp;Does not have a virtual disk&nbsp;</span></p></td>"
                $highL = $true
            }
            else
            {    
                foreach($vmDisk in $vmDisks)
                {
                    [array]$vmDiskData = $null

                    # Name, Path, Type, Size and File Size
                    $vmDiskName = $vmDisk.Path.Split('\')[-1]
                    $vmDiskPath = $vmDisk.Path
                    $vmDiskType = $vmDisk.VhdType
                    $vmDiskMaxSize = sConvert-Size -DiskVolumeSpace $vmDisk.Size -DiskVolumeSpaceUnit byte
                    $vmDiskFileSize = sConvert-Size -DiskVolumeSpace $vmDisk.FileSize -DiskVolumeSpaceUnit byte

                    # For Cluster Overview
                    if ($VM.IsClustered -eq $true -and $VM.State -eq "Running")
                    {
                        $ovUsedVmVHD = $ovUsedVmVHD + $vmDisk.FileSize
                        $ovTotalVmVHD = $ovTotalVmVHD + $vmDisk.Size
                    }

                    # For Active VHDs File Size
                    $activeVhdFileSize = $vmDisk.FileSize
                 
                    # Get Controller Type
                    $vmDiskControllerType = (Get-VMHardDiskDrive -ComputerName $VMHostItem -VMName $vm.VMName | where{$_.Path -eq $vmDiskPath}).ControllerType

                    # VHD Fragmentation and Color
                    if ($vmDisk.FragmentationPercentage -eq $null)
                    {
                       $vmDiskFragmentation = "N/A"
                       $vmDiskFragmentationBgColor = ""
                       $vmDiskFragmentationTextColor = ""
                    }
                    else
                    {
                       $vmDiskFragmentation = "%$($vmDisk.FragmentationPercentage)"
                
                       if (($vmDisk.FragmentationPercentage -ge "30") -and ($vmDisk.FragmentationPercentage -lt "50")) 
                       {
                           $vmDiskFragmentationBgColor = $stateBgColors[4]
                           $vmDiskFragmentationTextColor = $stateWordColors[4]
                           $highL = $true
                       }
                       elseif ($vmDisk.FragmentationPercentage -ge "50") 
                       {
                           $vmDiskFragmentationBgColor = $stateBgColors[3]
                           $vmDiskFragmentationTextColor = $stateWordColors[3]
                           $highL = $true
                       }
                       else
                       {
                           $vmDiskFragmentationBgColor = ""
                           $vmDiskFragmentationTextColor = ""
                       }
                    }

                    # If differencing exist
                    if ($vmDisk.ParentPath)
                    {
                        # Checkpoint label
                        $cpNumber = $null
                        $cpNumber = $vmChekpointCount

                        if ($vmDiskPath.EndsWith(".avhdx",1))
                        {
                            if (($cpNumber -ne 0) -or ($cpNumber -ne $null))
                            {
                                $vmDiskName = "Checkpoint $cpNumber"
                                $cpNumber = $cpNumber - 1
                            }
                        }

                        $vmDiskData += ""
                        $parentPath = $vmDisk.ParentPath

                        # Differencing disk loop
                        Do
                        {
                            $vmDiskName = $null
                            $vmDiskPath = $null
                            $vmDiskType = $null
                            $vmDiskMaxSize = $null
                            $vmDiskFileSize = $null
                            $vmDiffDisk = Get-VHD -ComputerName $VMHostItem -Path $parentPath
                            $vmDiskPath = $vmDiffDisk.Path
                            $vmDiskName = $vmDiffDisk.Path.Split('\')[-1]

                            # Checkpoint label
                            if ($vmDiskPath.EndsWith(".avhdx",1))
                            {
                                if (($cpNumber -ne 0) -or ($cpNumber -ne $null))
                                {
                                    $vmDiskName = "Checkpoint $cpNumber"
                                    $cpNumber = $cpNumber - 1
                                }
                            }

                            $vmDiskType = $vmDiffDisk.VhdType
                            $vmDiskMaxSize = sConvert-Size -DiskVolumeSpace $vmDiffDisk.Size -DiskVolumeSpaceUnit byte
                            $vmDiskFileSize = sConvert-Size -DiskVolumeSpace $vmDiffDisk.FileSize -DiskVolumeSpaceUnit byte

                            # For Active VHD file size
                            $activeVhdFileSize = $activeVhdFileSize + $vmDiffDisk.FileSize

                            # For Cluster Overview
                            if ($VM.IsClustered -eq $true -and $VM.State -eq "Running")
                            {
                                $ovUsedVmVHD = $ovUsedVmVHD + $vmDiffDisk.FileSize
                            }

                            # Disk Fragmentation and Color
                            if ($vmDiffDisk.FragmentationPercentage)
                            {
                                $vmDiskFragmentation = "%$($vmDiffDisk.FragmentationPercentage)"
                
                                if (($vmDiffDisk.FragmentationPercentage -ge "30") -and ($vmDiffDisk.FragmentationPercentage -lt "50")) 
                                {
                                    $vmDiskFragmentationBgColor = $stateBgColors[4]
                                    $vmDiskFragmentationTextColor = $stateWordColors[4]
                                    $highL = $true
                                }
                                elseif ($vmDiffDisk.FragmentationPercentage -ge "50") 
                                {
                                    $vmDiskFragmentationBgColor = $stateBgColors[3]
                                    $vmDiskFragmentationTextColor = $stateWordColors[3]
                                    $highL = $true
                                }
                                else
                                {
                                    $vmDiskFragmentationBgColor = ""
                                    $vmDiskFragmentationTextColor = ""
                                }
                            }

                            $vmDiskData += ""
                            $parentPath = $vmDiffDisk.ParentPath
                        }
                        Until ($parentPath -eq $null)
                    }
                    else
                    {
                        $vmDiskData = ""
                    }

                    # Active VHD Array ($activeVhds)
                    if ($vm.State -eq "Running")
                    {
                        $vhdHash = @{
            
                            Path      = $vmDisk.Path
                            Size      = $vmDisk.Size
                            FileSize  = $activeVhdFileSize
                            Host      = $Vm.ComputerName
                            VhdType   = $vmDisk.VhdType
                            VhdFormat = $vmDisk.VhdFormat
                            Attached  = $vmDisk.Attached
                            VMName    = $Vm.VMName
                            }

                        # Create PSCustom object
                        $customObjVHD = New-Object PSObject -Property $vhdHash

                        # Add to Array
                        $activeVhds += $customObjVHD
                    }

                    # Remove top-margin of last item
                    #$vmDiskData[-1] = $vmDiskData[-1].Replace("margin-top:5px;","")

                    # Add Indents
                    $itemC = 0
                    $indentV = ($vmDiskData.count - 1) * 14

                    Do
                    {
                        $vmDiskData[$itemC] = $vmDiskData[$itemC].Replace("1nd3ntPlaceHolder","$indentV")
                        $indentV = $indentV - 14
                        $itemC = $itemC + 1
                    }
                    Until ($itemC -eq $vmDiskData.Count)


                    # Convert String
                    [array]::Reverse($vmDiskData) 
                    $vmDiskData = ($vmDiskData -join "").ToString()

                    # Write
                    if ($rowSpanCount -eq 0)
                    {
			            $vmDiskOutput =""
                        $rowSpanCount = $rowSpanCount + 1
                    }
                    else
                    {
                        $vmDiskOutput +=""
                        $rowSpanCount = $rowSpanCount + 1
                    }
                }
            }

            #If single VHD, rowSpanCount equal to 0 
            if ($rowSpanCount -eq 1)
            {
                $rowSpanCount = 0
            }
            
            # VM Memory Information
            if ($VM.DynamicMemoryEnabled)
            {
                # Startup Memory
                $outVmMemStartup = sConvert-Size -DiskVolumeSpace $VM.MemoryStartup -DiskVolumeSpaceUnit byte

                # Assigned Memory
                if ($VM.MemoryAssigned -eq 0)
                {
                    $outVmMemAssigned = "-"
                }
                else
                {
                    $outVmMemAssigned = sConvert-Size -DiskVolumeSpace $VM.MemoryAssigned -DiskVolumeSpaceUnit byte
                }

                # Maximum Memory, Minimum Memory
                $outVmMemMax = sConvert-Size -DiskVolumeSpace $VM.MemoryMaximum -DiskVolumeSpaceUnit byte
                $outVmMemMin = sConvert-Size -DiskVolumeSpace $VM.MemoryMinimum -DiskVolumeSpaceUnit byte

                # Charge chargerVmMemoryTable
                $chargerVmMemoryTable =""
            }
            else
            {
                # Startup Memory
                $outVmMemStartup = sConvert-Size -DiskVolumeSpace $VM.MemoryStartup -DiskVolumeSpaceUnit byte

                # Charge chargerVmMemoryTable
                $chargerVmMemoryTable ="
"
            }

            # Data Line
            $chargerVmTable +="
            <tr style=""background:$($vmTableTrBgColor)""><!--Data Line-->
                <td rowspan=""$($rowSpanCount)""><p style=""text-align:left""><abbr title=""$($outVmPath)"">$($outVmName) <span style=""font-size:10px;color:orange"">*</span></abbr> $($outVmGenVer) </span></p></td>
                <td rowspan=""$($rowSpanCount)"" bgcolor=""$vmStateBgColor""><p style=""color:$($vmStateWordColor)"">$($outVmState)</p></td>
                <td rowspan=""$($rowSpanCount)""><p>$($outVmUptimeDays)$($outVmUptime)</p></td>
            "
            $chargerVmTable += $chargerVmMemoryTable
            
            #<td rowspan=""$($rowSpanCount)""><p style=""line-height:1.1"">$($outVmCPU)<br><span style=""font-size:10px"">CPU</span></p></td>"

            $chargerVmTable +=""
		        $chargerVmTable += $vmDiskOutput

            # Output Data
            if ($HighlightsOnly -eq $false)
            {
                $outVMTable += $chargerVmTable
                $vmNoInTable = $vmNoInTable + 1
            }
            elseif (($HighlightsOnly -eq $true) -and ($highL -eq $true))
            {      
                $outVMTable += $chargerVmTable
                $vmNoInTable = $vmNoInTable + 1
            }
            else
            {
                # Blank
            }
        }
    }
    # Error
    elseif ($getVMerr)
    {
        sPrint -Type 0 -Message "$($VMHostItem.ToUpper()): $($getVMerr.exception.message)" -WriteToLogFile $True
        sPrint -Type 2 -Message "Gathering VM Information for '$($VMHostItem.ToUpper())' failed." -WriteToLogFile $True
        Start-Sleep -Seconds 3
        Continue
    }
    else
    # Blank
    {
        sPrint -Type 2 -Message "$($VMHostItem.ToUpper()): Does not have Virtual Machine." -WriteToLogFile $True
        Start-Sleep -Seconds 3
    }

    # If detected clustered VM configuration resource problem
    if ($offlineVmConfigs)
    {
        ForEach ($offlineVmConfig in $offlineVmConfigs)
        {
            # Table TR Color
            if([bool]!($vmNoInTable%2))
            {
               #Even or Zero
               $vmTableTrBgColor = ""
            }
            else
            {
               #Odd
               $vmTableTrBgColor = "#F9F9F9"
            }

            $outVMTable +="
            <tr style=""background:$($vmTableTrBgColor)""><!--Data Line-->
                <td><p style=""text-align:left""><abbr title=""Virtual Machine Configuration resource is $($offlineVmConfig.State)"">$($offlineVmConfig.OwnerGroup) <span style=""font-size:10px;color:orange"">*</span></abbr> <br><span style=""font-size:10px;color:#BDBDBD"">IsClustered:Yes</span></p></td>
                <td bgcolor=""$($stateBgColors[3])""><p style=""color:$($stateWordColors[3])"">$($offlineVmConfig.State)</p></td>
                <td><p>-</p></td>
                <td><p>-</p></td>
                <td colspan=""4""><p>-</p></td>
                <td><p>-</p></td>
                <td><p>-</p></td>
                <td><p>-</p></td>
                <td><p>-</p></td>
                <td><p style=""text-align:left""><span style=""background-color:$($stateBgColors[3]);color:$($stateWordColors[3])"">&nbsp;VM cluster configuration resource is $($offlineVmConfig.State)&nbsp;</span></p></td>
            </tr>"

            $vmNoInTable = $vmNoInTable + 1
        }
    }
}

if (($HighlightsOnly -eq $true) -and ($outVmTable -eq $null) -and ($cntVM -ne 0))
{
    $outVmTable +="
            <tr><!--Data Line-->
                <td colspan=""14""><p style=""text-align:center""><span style=""padding-top:1px;padding-bottom:1px;background-color:#ACFA58;color:#298A08"">&nbsp;&nbsp;All VMs are healthy&nbsp;&nbsp;</span></p></td>
            </tr>"
}

if (($outVmTable -eq $null) -and ($cntVM -eq 0))
{
    $outVmTable +="
            <tr><!--Data Line-->
                <td colspan=""14""><p style=""text-align:center""><span style=""color:#BDBDBD"">No virtual machine for reporting</span></p></td>
            </tr>"
}

# VMs Table - End
$outVMTableEnd ="
        </tbody>
        </table>
    </div><!--End VMs Class-->"

#endregion

#region HTML End
#---------------

$outHtmlEnd ="
</div><!--End ReportBody-->
<br>
</body>
</html>"

# Print MSG
sPrint -Type 1 -Message "Writing output to file $ReportFile" -WriteToLogFile $True

if ($Cluster)
{
    $outFullHTML = $outHtmlStart + $outClusterOverview + $outVMHostTableStart + $outVMHostTable + $outVMHostTableEnd + $outVolumeTableStart + $outVolumeTable + $outVolumeTableEnd + $outVMTableStart + $outVMTable + $outVMTableEnd + $outHtmlEnd
}
else
{
    $outFullHTML = $outHtmlStart + $outVMHostTableStart + $outVMHostTable + $outVMHostTableEnd + $outVolumeTableStart + $outVolumeTable + $outVolumeTableEnd + $outVMTableStart + $outVMTable + $outVMTableEnd + $outHtmlEnd
}

$outFullHTML | Out-File $ReportFile

if (Test-Path -Path $ReportFile)
{
    sPrint -Type 1 -Message "Report created successfully." -WriteToLogFile $True
}
else
{
    sPrint -Type 2 -Message "Reporting file could not be created. Please review the log file." -WriteToLogFile $True
}
if ($TestMode) {iex $ReportFile}

#endregion
