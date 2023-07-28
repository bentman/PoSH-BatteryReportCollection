<#
.SYNOPSIS
    Invoke-BatteryReportCollection is a script to collect 'powercfg.exe /batteryreport' information and store it in a custom WMI class.
.DESCRIPTION
    The script creates a custom WMI namespace and class if they don't exist. 
    It then runs 'powercfg /batteryreport' command to generate a battery report in both XML and HTML format. 
    The XML report is parsed to extract battery information which is then stored in the custom WMI class. 
    The reports are saved in the 'C:\temp\batteryreport' directory, and are overwritten each time the script runs.
    Being a complex script that handles many operations, thoroughly test it in a safe environment before using it in a production scenario!
.PARAMETER None
    This script does not accept any parameters.
.EXAMPLE
    This example shows how to run the script.
        PS C:\> .\Invoke-BatteryReportCollection.ps1
.NOTES
    Version 1.5: Corrected issue #1 "Get-CimInstance: Invoke-BatteryReportCollection.ps1:186:21"
    Release Date: 2023-07-28
    Requires: PowerShell V3+. Requires elevated Admin privileges.
    https://github.com/bentman/PoSH-BatteryReportCollection
    Copyright (c) 2023 https://github.com/bentman
.LINK
    https://learn.microsoft.com/en-us/powershell/
#>

############################### VARIABLES ###############################
# Script Version
$scriptVer = "1.5"
# Define new Namespace name
$newClassName = "BatteryReport"
# Define Report folder root
$reportFolder = "C:\temp"

############################### GENERATED ###############################
# Define Battery report outputs
$reportFolderPath = Join-Path $reportFolder "$newClassName"
$reportPathXml = Join-Path $reportFolderPath "batteryreport.xml"
$reportPathHtml = Join-Path $reportFolderPath "batteryreport.html"
$transcriptPath = Join-Path $reportFolderPath "BatteryReportCollection.log"
# WMI namespace path
$namespacePath = "root\cimv2\$newClassName"

############################### FUNCTIONS ###############################
function New-WmiClass {
    [CmdletBinding()] param (
        [Parameter(Mandatory=$true)][string]$namespacePath,
        [Parameter(Mandatory=$true)][string]$newClassName
    )
    try {
        # Check if the Namespace exists, if not create it
        $namespace = Get-WmiObject -Namespace "root\cimv2" -Query "SELECT * FROM __Namespace WHERE Name='$newClassName'"
        if (!$namespace) {
            Write-Output "Creating namespace: $newClassName"
            $namespace = ([wmiclass]'root\cimv2:__Namespace').CreateInstance()
            $namespace.Name = "$newClassName"
            $namespace.Put() | Out-Null
        }
        # Check if the Class exists, if not create it
        $class = Get-WmiObject -Namespace $namespacePath -List | Where-Object {$_.Name -eq $newClassName}
        if (!$class) {
            Write-Output "Creating class: $newClassName"
            $class = New-Object System.Management.ManagementClass($namespacePath, $newClassName, $null)
            $class.Qualifiers.Add("Static", $true) | Out-Null
            $class.Properties.Add("ComputerName", [System.Management.CimType]::String, $false) | Out-Null
            $class.Properties["ComputerName"].Qualifiers.Add("key", $true) | Out-Null
            $class.Properties.Add("DesignCapacity", [System.Management.CimType]::UInt32, $false) | Out-Null
            $class.Properties.Add("FullChargeCapacity", [System.Management.CimType]::UInt32, $false) | Out-Null
            $class.Properties.Add("CycleCount", [System.Management.CimType]::UInt32, $false) | Out-Null
            $class.Properties.Add("ActiveRuntime", [System.Management.CimType]::String, $false) | Out-Null
            $class.Properties.Add("ActiveRuntimeAtDesignCapacity", [System.Management.CimType]::String, $false) | Out-Null
            $class.Properties.Add("ModernStandby", [System.Management.CimType]::String, $false) | Out-Null
            $class.Properties.Add("ModernStandbyAtDesignCapacity", [System.Management.CimType]::String, $false) | Out-Null
            $class.Put()
            Write-Output "Class $newClassName has been successfully created in the namespace $namespacePath."
        }
    }
    catch {
        Write-Error "Failed to create the class $newClassName in the namespace $namespacePath. Error: $_"
    }
}

function ConvertTo-StandardTimeFormat {
    [CmdletBinding()] param(
        # ISO 8601 duration passed as a string
        [Parameter(Mandatory=$true)]
        [string]$iso8601Duration
    )
    # Validate input: Must start with 'PT' (ISO 8601)
    if ($iso8601Duration -notmatch '^PT') {
        Write-Error "Invalid ISO 8601 duration format. The input should start with 'PT'."
        return
    }
    # Return a default value for empty or just 'PT' input
    if ([string]::IsNullOrEmpty($iso8601Duration) -or $iso8601Duration -eq 'PT') {
        return "00:00:00"
    }
    try {
        # Regular Expression pattern
        $match = [Regex]::Match($iso8601Duration, 'PT((?<hours>\d+)H)?((?<minutes>\d+)M)?((?<seconds>\d+)S)?')
        # Regular Expression matching
        $hours = [int]$match.Groups['hours'].Value
        $minutes = [int]$match.Groups['minutes'].Value
        $seconds = [int]$match.Groups['seconds'].Value
        # Validate if we could extract any time unit
        if ($hours -eq 0 -and $minutes -eq 0 -and $seconds -eq 0) {
            Write-Error "Invalid ISO 8601 duration format. No valid time unit (hour, minute, second) could be extracted."
            return
        }
        # Convert ISO 8601 timespan to standard "HH:MM:SS" format
        $timespan = New-TimeSpan -Hours $hours -Minutes $minutes -Seconds $seconds
        return $timespan.ToString("hh\:mm\:ss")
    }
    catch {
        Write-Error "Failed to convert ISO 8601 duration to standard time format. Error: $_"
    }
}

############################### EXECUTION ###############################

# Create output folder if it does not exist
if (-not (Test-Path -Path $reportFolderPath)) {
    New-Item -ItemType Directory -Path $reportFolderPath | Out-Null
}

# Start PowerShell Transcript
Start-Transcript -Path $transcriptPath -Force
Write-Host "" # Empty line for transcript readability
Write-Host "Script Version = $scriptVer"
(Get-PSCallStack).InvocationInfo.MyCommand.Name

<# Check if a battery is present in the system
$batteryPresent = Get-CimInstance -ClassName Win32_Battery
if (-not $batteryPresent) {
    # If a battery is not found, log an error and exit the script
    Write-Host "Error: No battery detected on this system - Skipping all operations." -ForegroundColor Red
    Stop-Transcript
    exit
}#>

# Execution wrapped in a Try/Catch
try { 
    try {
        # Create WMI class
        New-WmiClass -namespacePath $namespacePath -newClassName $newClassName -Verbose
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Stop-Transcript
        exit
    }
    
    # Generate HTML 'powercfg /batteryreport' for readability
    powercfg /batteryreport /output $reportPathHtml

    # Run XML 'powercfg /batteryreport'
    powercfg /batteryreport /output $reportPathXml /xml

    # Store XML 'powercfg /batteryreport' content
    [xml]$batteryReport = Get-Content $reportPathXml

    # Parse the battery and runtime estimates from the XML report
    $reportContent = $batteryReport.BatteryReport.Report
    $battery = $reportContent.Batteries.Battery
    $runtimeEstimates = $reportContent.RuntimeEstimates

    # Create a hash table with all the values
    $batteryData = @{
        DesignCapacity = $battery.DesignCapacity
        FullChargeCapacity = $battery.FullChargeCapacity
        CycleCount = $battery.CycleCount
        ActiveRuntime = $runtimeEstimates.FullCharge.ActiveRuntime
        ActiveRuntimeAtDesignCapacity = $runtimeEstimates.FullCharge.ActiveRuntimeAtDesignCapacity
        ModernStandby = $runtimeEstimates.FullCharge.ModernStandby
        ModernStandbyAtDesignCapacity = $runtimeEstimates.FullCharge.ModernStandbyAtDesignCapacity
    }

    # Convert ActiveRuntime from ISO 8601 duration format to HH:MM:SS
    foreach ($key in 'ActiveRuntime', 'ActiveRuntimeAtDesignCapacity', 'ModernStandby', 'ModernStandbyAtDesignCapacity') {
        if ($batteryData[$key]) {
            $batteryData[$key] = ConvertTo-StandardTimeFormat -iso8601Duration $batteryData[$key]
        }
    }

    # Check if an instance with the same ComputerName already exists
    $instance = Get-CimInstance -Namespace $namespacePath -ClassName $newClassName |
        Where-Object { $_.ComputerName -eq $env:COMPUTERNAME }

    # Instance does not exist, create it
    if ($null -eq $instance) {
        Set-WmiInstance -Namespace $namespacePath -Class $newClassName -Arguments @{
            ComputerName = $env:COMPUTERNAME
        } + $batteryData
    } else {
        # Instance exists, update it
        $instance | Set-CimInstance -Property $batteryData
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Get instances of the new BatteryReport CIM class for logging
    $logInstances = Get-CimInstance -Namespace $namespacePath -ClassName $newClassName

    Write-Host "`nNew Class: $newClassName"

# Define a list of properties to display
$propertiesToDisplay = @(
    'ComputerName', 
    'DesignCapacity', 
    'FullChargeCapacity', 
    'CycleCount', 
    'ActiveRuntime', 
    'ActiveRuntimeAtDesignCapacity', 
    'ModernStandby', 
    'ModernStandbyAtDesignCapacity'
)

# Loop through each instance and log the properties
foreach ($logInstance in $logInstances) {
    Write-Host "" # Empty line for transcript readability
    $logInstance.PSObject.Properties | 
        Where-Object { $_.Name -in $propertiesToDisplay } |  # Only include properties in the display list
        ForEach-Object {
            Write-Host "    $($_.Name): $($_.Value)"
        }
    Write-Host "" # Empty line for transcript readability
}
    Stop-Transcript
}
