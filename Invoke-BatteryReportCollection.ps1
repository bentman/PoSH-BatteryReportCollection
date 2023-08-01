<#
.SYNOPSIS
   Invoke-BatteryReportCollection is a script to collect 'powercfg.exe /batteryreport' and 'powercfg /energy' information, storing it in a custom WMI class.
.DESCRIPTION
   The script creates a custom WMI namespace and class if they don't exist. 
   It then runs 'powercfg /batteryreport' and 'powercfg /energy' commands to generate battery and system energy reports in HTML format. 
   Both reports are parsed to extract battery and energy consumption information which are then stored in the custom WMI class. 
   The reports are saved in the 'C:\temp\batteryreport' directory, and are overwritten each time the script runs.
.PARAMETER None
   This script does not accept any parameters.
.EXAMPLE
   PS C:\> .\Invoke-BatteryReportCollection.ps1
   This example shows how to run the script.
.NOTES
    NOTE: As this script is complex and handles many operations, thoroughly test it in a safe environment before using it in a production scenario.
    Version: 3.0
    Last Updated: 2023-07-31
    Copyright (c) 2023 https://github.com/bentman
    https://github.com/bentman/
    Requires: PowerShell V3+. Needs administrative privileges.
.CHANGE
   Version 1.0: Initial script
   Version 2.0: Corrected errors collected in feedback & added better logging.
   Version 3.0: Changed WMI Path to 'root\cimv2\BatteryReport:BatteryInfo'
                Expanded parsing capabilities to extract energy consumption information.
.LINK
    https://docs.microsoft.com/powershell
#>
############################## VARIABLES ###############################
# Script Version
$scriptVer = "3.0"
# Local hard drive location for processing
$reportFolder = "C:\temp"
# Name to assign WMI Namespace 
$newWmiNamespace = "BatteryReport"
# Name to assign WMI Namespace Class
$newWmiClass = "BatteryInfo"
# WMI Parent designation (recommended for SCCM) 
$namespaceParent = "root\cimv2"

############################### GENERATED ###############################
# Subfolder created on Local hard drive location for processing and transcript log
$reportFolderPath = Join-Path $reportFolder "$newWmiNamespace"
# XML Battery Report for parsing
$reportPathXml = Join-Path $reportFolderPath "$newWmiNamespace.xml"
# HTML Battery Report for human readability
$reportPathHtml = Join-Path $reportFolderPath "$newWmiNamespace.html"
# Script transcript for logging
$transcriptPath = Join-Path $reportFolderPath "$($newWmiNamespace)-Transcript.log"
# WMI Namespace path for operations (root\cimv2\BatteryReport)
$newWmiNamespacePath = "$namespaceParent\$newWmiNamespace"

############################### FUNCTION  ###############################
function Remove-WmiNamespaceClass { # Remove previous WMI Parent\Namespace:Class if exist 
    [CmdletBinding()] param(
        [Parameter(Mandatory=$true)][string] $oldWmiNamespace,
        [Parameter(Mandatory=$true)][string] $oldWmiClass
    )
    try {
        $existingBatteryReport = (Get-WmiObject -Namespace $oldWmiNamespace -Class $oldWmiClass -ErrorAction SilentlyContinue)
        if ($existingBatteryReport) {
            Write-Host "`n### Existing $($oldWmiNamespace):$($oldWmiClass) info was found."
            Write-Host "### $oldWmiNamespace will be removed before creating new."
            Remove-WmiObject -Namespace $oldWmiNamespace -Class $oldWmiClass -Verbose
        }
    } catch {
        Write-Host "ERROR: Removing WMI namespace" -ForegroundColor Green
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function ConvertTo-StandardTimeFormat {
    param( # ISO 8601 duration passed as a string
        [Parameter(Mandatory=$true)][string]$iso8601Duration
    )
    # Check for null, empty string, or whitespace
    if ([string]::IsNullOrWhiteSpace($iso8601Duration)) {
        Write-Error "`n### The ISO 8601 duration cannot be null, empty, or whitespace."
        return '00:00:00'
    }
    # Regular Expression pattern
    $match = [Regex]::Match($iso8601Duration, 'PT((?<hours>\d+)H)?((?<minutes>\d+)M)?((?<seconds>\d+)S)?')
    # Regular Expression matching
    $hours = [int]$match.Groups['hours'].Value
    $minutes = [int]$match.Groups['minutes'].Value
    $seconds = [int]$match.Groups['seconds'].Value
    # Check for match success
    if (!$match.Success) {
        Write-Error "`n### The provided string does not match the ISO 8601 duration format."
        return '00:00:00'
    }
    # Convert ISO 8601 timespan to standard "HH:MM:SS" format
    $timespan = New-TimeSpan -Hours $hours -Minutes $minutes -Seconds $seconds
    return $timespan.ToString("hh\:mm\:ss")
}

############################### EXECUTION ###############################
# Establish Logging
    # Create output folder if it does not exist
    if (!(Test-Path -Path $reportFolderPath)) {New-Item -ItemType Directory -Path $reportFolderPath}
    # Start transcript logging
    Start-Transcript -Path $transcriptPath
    # Log script name & version
    (Get-PSCallStack).InvocationInfo.MyCommand.Name
    Write-Host "Script Version = $scriptVer"

# Check if the system has a battery
    # Check if a battery is present in the system
    $batteryPresent = Get-WmiObject -Class Win32_Battery
    if (!$batteryPresent) {
        $noBattery = $true
        Write-Host "`n### Error: No battery detected on this system - Testing only WMI/CIM." -ForegroundColor Red
        # Stop-Transcript
        # Exit
    }

# Remove erroneous WMI 'root\cimv2\BatteryReport:BatteryReport' if exist
    Write-Host "`n### Removing $($newWmiNamespacePath):$($newWmiNamespace) if found."
    Remove-WmiNamespaceClass -oldWmiNamespace $newWmiNamespacePath -oldWmiClass $newWmiNamespace

# Remove previous WMI 'root\cimv2\BatteryReport:BatteryInfo' if exist
    Write-Host "`n### Removing $($newWmiNamespacePath):$($newWmiClass) if found."
    Remove-WmiNamespaceClass -oldWmiNamespace $newWmiNamespacePath -oldWmiClass $newWmiClass

# Generate PowerCfg Battery Reports
    # Generate HTML battery report for readability
    Write-Host "`n### Generating HTML Battery Report..."
    PowerCfg.exe /batteryreport /output $reportPathHtml
    Write-Host "`n### HTML Battery Report generated at $reportPathHtml"
    # Generate XML battery report for script consumption
    Write-Host "`n### Generating XML Battery Report..."
    PowerCfg.exe /batteryreport /output $reportPathXml /XML
    Write-Host "`n### XML Battery Report generated at $reportPathXml"

# Store XML 'powercfg /batteryreport' content
    Write-Host "`n### Getting contents of XML Battery Report..."
    [xml]$batteryReport = Get-Content $reportPathXml

# Parse PowerCfg Battery Report info from XML
    Write-Host "`n### Parsing XML Battery Report..."
    $computerName = $batteryReport.BatteryReport.SystemInformation.ComputerName
    $systemManufacturer = $batteryReport.BatteryReport.SystemInformation.SystemManufacturer
    $systemProductName = $batteryReport.BatteryReport.SystemInformation.SystemProductName
    $designCapacity = [uint64]$batteryReport.BatteryReport.Batteries.Battery.DesignCapacity
    $fullChargeCapacity = [uint64]$batteryReport.BatteryReport.Batteries.Battery.FullChargeCapacity
    $relativeCapacity = [uint32]$batteryReport.BatteryReport.Batteries.Battery.RelativeCapacity
    $cycleCount = [uint32]$batteryReport.BatteryReport.Batteries.Battery.CycleCount
    $designActiveRuntimeValue = $batteryReport.BatteryReport.RuntimeEstimates.DesignCapacity.ActiveRuntime
    $designStandbyRuntimeValue = $batteryReport.BatteryReport.RuntimeEstimates.DesignCapacity.ConnectedStandbyRuntime
    $fullChargeActiveRuntimeValue = $batteryReport.BatteryReport.RuntimeEstimates.FullChargeCapacity.ActiveRuntime
    $fullChargeStandbyRuntimeValue = $batteryReport.BatteryReport.RuntimeEstimates.FullChargeCapacity.ConnectedStandbyRuntime

# Convert 8601 durations to HH:MM:SS format
    Write-Host "`n### Converting ($)designActiveRuntime 8601 durations to HH:MM:SS format..."
    $designActiveRuntime = if ([string]::IsNullOrWhiteSpace($designActiveRuntimeValue)) {'00:00:00'} else {ConvertTo-StandardTimeFormat $designActiveRuntimeValue}
    Write-Host "`n### Converting ($)designStandbyRuntime 8601 durations to HH:MM:SS format..."
    $designStandbyRuntime = if ([string]::IsNullOrWhiteSpace($designStandbyRuntimeValue)) {'00:00:00'} else {ConvertTo-StandardTimeFormat $designStandbyRuntimeValue}
    Write-Host "`n### Converting ($)fullChargeActiveRuntime 8601 durations to HH:MM:SS format..."
    $fullChargeActiveRuntime = if ([string]::IsNullOrWhiteSpace($fullChargeActiveRuntimeValue)) {'00:00:00'} else {ConvertTo-StandardTimeFormat $fullChargeActiveRuntimeValue}
    Write-Host "`n### Converting ($)fullChargeStandbyRuntime 8601 durations to HH:MM:SS format..."
    $fullChargeStandbyRuntime = if ([string]::IsNullOrWhiteSpace($fullChargeStandbyRuntimeValue)) {'00:00:00'} else {ConvertTo-StandardTimeFormat $fullChargeStandbyRuntimeValue}

# Parse battery usage over last 21 days and convert to string
    Write-Host "`n### Converting Recent Usage history over last 21 days to a string..."
    $recentUsages = if ($noBattery) {'No recent battery usage reported'
    } else {
        $batteryReport.BatteryReport.Report.RecentUsage.Usage | ForEach-Object {
            "$($_.Timestamp.Substring(0, 10)), $($_.Duration), $($_.Ac), $($_.EntryType), $($_.ChargeCapacity), $($_.Discharge), $($_.FullChargeCapacity), $($_.IsNextOnBattery)"
        } | Select-Object -Last 21 | Out-String
    Write-Host "### The following information stored as a string... "
    Write-Host "### Duration | Ac [0,1] EntryType | ChargeCapacity | Discharge | FullChargeCapacity | IsNextOnBattery"
    Write-Host "### Retrieving this data and using it for reporting is discussed in documentation"
    }

# Create WMI Namespace, Class at 'root\cimv2\BatteryReport:BatteryInfo' 
Try { # Create WMI class if not exists
    Write-Host "`n### Creating WMI Namespace\Class\Properties at '$($newWmiNamespacePath):$($newWmiClass)'..."
    $class = Get-WmiObject -Namespace $newWmiNamespacePath -List | Where-Object {$_.Name -eq $newWmiClass}
    if (!$class) {
        $class = New-Object System.Management.ManagementClass("$newWmiNamespacePath", [string]::Empty, $null)
        $class["__CLASS"] = $newWmiClass
        $class.Qualifiers.Add("Static", $true) | Out-Null
        $class.Properties.Add("ComputerName", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties["ComputerName"].Qualifiers.Add("key", $true) | Out-Null
        $class.Properties.Add("SystemManufacturer", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties.Add("SystemProductName", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties.Add("DesignCapacity", [System.Management.CimType]::UInt32, $false) | Out-Null
        $class.Properties.Add("FullChargeCapacity", [System.Management.CimType]::UInt32, $false) | Out-Null
        $class.Properties.Add("CycleCount", [System.Management.CimType]::UInt32, $false) | Out-Null
        $class.Properties.Add("DesignActiveRuntime", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties.Add("DesignStandbyRuntime", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties.Add("FullChargeActiveRuntime", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties.Add("FullChargeStandbyRuntime", [System.Management.CimType]::String, $false) | Out-Null
        $class.Properties.Add("RecentUsages", [System.Management.CimType]::String, $false) | Out-Null
        $class.Put()
    }
}
Catch {
    Write-Host "`n### Failed to create 'root\cimv2\BatteryReport:BatteryInfo'" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    Exit
}

# Create an instance of the new BatteryInfo WMI class
    $newInstance = $class.CreateInstance()
    # Assign Battery Info values from Battery Report values
    $newInstance.ComputerName = $computerName
    $newInstance.SystemManufacturer = $systemManufacturer
    $newInstance.SystemProductName = $systemProductName
    $newInstance.DesignCapacity = $designCapacity
    $newInstance.FullChargeCapacity = $fullChargeCapacity
    $newInstance.CycleCount = $cycleCount
    $newInstance.DesignActiveRuntime = $designActiveRuntime
    $newInstance.DesignStandbyRuntime = $designStandbyRuntime
    $newInstance.FullChargeActiveRuntime = $fullChargeActiveRuntime
    $newInstance.FullChargeStandbyRuntime = $fullChargeStandbyRuntime
    $newInstance.RecentUsages = $recentUsages
    # Save the instance
    $newInstance.Put() 

# Logging
    Write-Host "Script Version = $scriptVer"
    (Get-PSCallStack).InvocationInfo.MyCommand.Name
    Write-Host "`nNew Class: $newWmiNamespacePath "
    # Get instances of the new BatteryReport CIM class for logging
    $logInstances = Get-CimInstance -Namespace $newWmiNamespacePath -ClassName $newWmiClass
    # Define a list of properties to display
    $propertiesToDisplay = @(
        'ComputerName',
        'SystemManufacturer',
        'SystemProductName',
        'DesignCapacity',
        'FullChargeCapacity',
        'RelativeCapacity',
        'CycleCount',
        'DesignActiveRuntime',
        'DesignStandbyRuntime',
        'FullChargeActiveRuntime',
        'FullChargeStandbyRuntime'
    )
    # Loop through each instance and log the properties
    foreach ($logInstance in $logInstances) {
        Write-Host "" # Empty line for transcript readability
        $logInstance.PSObject.Properties | 
            Where-Object { $_.Name -in $propertiesToDisplay } | 
            ForEach-Object {
                Write-Host "    $($_.Name): $($_.Value)"
            }
            Write-Host "" # Empty line for transcript readability
    }

# Terminate Logging
Stop-Transcript
