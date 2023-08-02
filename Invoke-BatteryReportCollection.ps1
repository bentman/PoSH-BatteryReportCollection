<#
.SYNOPSIS
   Invoke-BatteryReportCollection is a script to collect 'powercfg.exe /batteryreport' information, storing it in a custom WMI class.
.DESCRIPTION
   The script creates a custom WMI namespace and class if they don't exist. 
   It then runs 'powercfg /batteryreport' and 'powercfg /batteryreport /XML' commands to generate battery reports in HTML & XML formats. 
   XML report is parsed to extract battery information which are then stored in the custom WMI class. 
.PARAMETER None
   This script does not accept any parameters.
.EXAMPLE
   PS C:\> .\Invoke-BatteryReportCollection.ps1
   This example shows how to run the script.
.NOTES
    NOTE: As this script is complex and handles many operations, thoroughly test it in a safe environment before using it in a production scenario.
    Version: 4.0
    Last Updated: 2023-08-02
    Copyright (c) 2023 https://github.com/bentman
    https://github.com/bentman/PoSH-BatteryReportCollection
.CHANGE
   Version 1.0: Initial script
   Version 2.0: Corrected errors collected in feedback & added better logging.
   Version 3.0: Corrected 'function Remove-WmiNamespaceClass' + logging
   Version 4.0: Verified w10 + Error Handling
.LINK
    https://docs.microsoft.com/powershell
#>
############################## VARIABLES ###############################
# Script Version
$scriptVer = "4.0"
# Local hard drive location for processing
$reportFolder = "C:\temp"
# Name to assign WMI Namespace 
$newWmiClass = "BatteryReport"
# WMI Parent designation (recommended for SCCM) 
$namespaceParent = "root\cimv2"
# WMI Namespace path for operations (root\cimv2\BatteryReport)
$newWmiClassPath = "$namespaceParent\$newWmiClass"

############################### GENERATED ###############################
# Subfolder created on Local hard drive location for processing and transcript log
$reportFolderPath = Join-Path $reportFolder "$newWmiClass"
# XML Battery Report for parsing
$reportPathXml = Join-Path $reportFolderPath "$newWmiClass.xml"
# HTML Battery Report for human readability
$reportPathHtml = Join-Path $reportFolderPath "$newWmiClass.html"
# Script transcript for logging
$transcriptPath = Join-Path $reportFolderPath "$($newWmiClass)-Transcript.log"

############################### FUNCTION  ###############################
function Remove-WmiNamespaceClass { # Remove previous WMI Parent\Namespace:Class if exist 
    [CmdletBinding()] param(
        [Parameter(Mandatory=$true)][string] $oldWmiNamespace,
        [Parameter(Mandatory=$true)][string] $oldWmiClass
    )
    try {
        $existingBatteryReport = (Get-WmiObject -Namespace $oldWmiNamespace -List | Where-Object {$_.Name -eq $oldWmiClass})

        if (Get-WmiObject -Namespace $oldWmiNamespace -List | Where-Object {$_.Name -eq $oldWmiClass}) {
            Write-Warning "### Existing $existingBatteryReport will be removed before adding new."
            Remove-WmiObject -Namespace $oldWmiNamespace -Class $oldWmiClass -Verbose
        }
    } catch {
        Write-Error "`n### ERROR: Removing previous WMI $oldWmiClass"
        Write-Error "$($_.Exception.Message)"
    }
}

function ConvertTo-StandardTimeFormat {
    param( # ISO 8601 duration passed as a string
        [Parameter(Mandatory=$true)][string]$iso8601Duration
    )
    # Check for null, empty string, or whitespace
    if ([string]::IsNullOrWhiteSpace($iso8601Duration)) {
        Write-Warning "### The ISO 8601 duration cannot be null, empty, or whitespace."
        Write-Host "### Value '0:00:00:00' will be inserted for 'Days:Hours:Minutes:Seconds'. "
        return '0:00:00:00'
    }
    # Regular Expression pattern
    $match = [Regex]::Match($iso8601Duration, 'P((?<days>\d+)D)?((?<hours>\d+)H)?T((?<hours>\d+)H)?((?<minutes>\d+)M)?((?<seconds>\d+)S)?')
    # Regular Expression matching
    $days = [int]$match.Groups['days'].Value
    $hours = [int]$match.Groups['hours'].Value
    $minutes = [int]$match.Groups['minutes'].Value
    $seconds = [int]$match.Groups['seconds'].Value
    # Check for match success
    if (!$match.Success) {
        Write-Warning "### The provided string does not match the ISO 8601 duration format."
        Write-Host "### Value '0:00:00:00' will be inserted for 'Days:Hours:Minutes:Seconds'. "
        return '0:00:00:00'
    }
    # Convert ISO 8601 timespan to standard "HH:MM:SS" format
    $timespan = New-TimeSpan -Days $days -Hours $hours -Minutes $minutes -Seconds $seconds
    return $timespan.ToString("d\:hh\:mm\:ss")
}

############################### EXECUTION ###############################
# Establish Logging
    # Create output folder if it does not exist
    if (!(Test-Path -Path $reportFolderPath)) {New-Item -ItemType Directory -Path $reportFolderPath -Verbose}
    # Start transcript logging
    Start-Transcript -Path $transcriptPath
    # Log script name & version
    Write-Host "`nScript Version = $scriptVer"
    (Get-PSCallStack).InvocationInfo.MyCommand.Name

# Check if the system has a battery
    # Check if a battery is present in the system
    $batteryPresent = Get-WmiObject -Class Win32_Battery
    if (!$batteryPresent) {Write-Warning "### No battery detected - Testing WMI/CIM only."}

# Remove erroneous WMI 'root\cimv2\BatteryReport:BatteryReport' if exist
    Remove-WmiNamespaceClass -oldWmiNamespace $namespaceParent -oldWmiClass $newWmiClass

# Generate HTML 'PowerCfg /batteryreport' for human readability
    Write-Host "`n### Generating HTML Battery Report..."
    PowerCfg.exe /batteryreport /output $reportPathHtml
    if ($LASTEXITCODE -ne '0') {Write-Warning "### $reportPathHtml was not generated!"}
    else {Write-Host "### HTML Battery Report generated at $reportPathHtml"}

# Generate XML 'PowerCfg /batteryreport' for script consumption
    Write-Host "`n### Generating XML Battery Report..."
    PowerCfg.exe /batteryreport /output $reportPathXml /XML
    if ($LASTEXITCODE -ne '0') {Write-Warning "`### $reportPathXml was not generated!"}
    else {Write-Host "### XML Battery Report generated at $reportPathXml"}

# Store XML 'powercfg /batteryreport' content
    Write-Host "`n### Getting contents of XML Battery Report..."
    [xml]$batteryReport = Get-Content $reportPathXml
# Parse PowerCfg Battery Report info from XML
    Write-Host "### Parsing XML Battery Report..."
    $computerName = $batteryReport.BatteryReport.SystemInformation.ComputerName
    $systemManufacturer = $batteryReport.BatteryReport.SystemInformation.SystemManufacturer
    $systemProductName = $batteryReport.BatteryReport.SystemInformation.SystemProductName
    $designCapacity = [uint64]$batteryReport.BatteryReport.Batteries.Battery.DesignCapacity
    $fullChargeCapacity = [uint64]$batteryReport.BatteryReport.Batteries.Battery.FullChargeCapacity
    $cycleCount = [uint32]$batteryReport.BatteryReport.Batteries.Battery.CycleCount
    $designActiveRuntimeValue = $batteryReport.BatteryReport.RuntimeEstimates.DesignCapacity.ActiveRuntime
    $fullChargeActiveRuntimeValue = $batteryReport.BatteryReport.RuntimeEstimates.FullChargeCapacity.ActiveRuntime
# Convert 8601 durations to HH:MM:SS format
    Write-Host "### Converting ($)designActiveRuntime 8601 durations to D:HH:MM:SS format..."
    $designActiveRuntime = if ([string]::IsNullOrWhiteSpace($designActiveRuntimeValue)) {'0:00:00:00'} 
        else {ConvertTo-StandardTimeFormat $designActiveRuntimeValue}
    Write-Host "### Converting ($)fullChargeActiveRuntime 8601 durations to D:HH:MM:SS format..."
    $fullChargeActiveRuntime = if ([string]::IsNullOrWhiteSpace($fullChargeActiveRuntimeValue)) {'0:00:00:00'} 
        else {ConvertTo-StandardTimeFormat $fullChargeActiveRuntimeValue}

# Create WMI Namespace\Class at '$newWmiClassPath' 
    Try { # Create WMI class if not exists
        Write-Host "`n### Creating WMI Namespace\Class\Properties at '$newWmiClassPath'..."
        $class = (Get-WmiObject -Namespace $namespaceParent -List | Where-Object {$_.Name -eq $newWmiClass})
        if (!$class) {
            $class = New-Object System.Management.ManagementClass("$namespaceParent", [string]::Empty, $null)
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
            $class.Properties.Add("FullChargeActiveRuntime", [System.Management.CimType]::String, $false) | Out-Null
            $class.Put()
        }
    } Catch {
        Write-Error "`n### Failed to create '$newWmiClassPath'"
        Write-Error "$($_.Exception.Message)"
        Stop-Transcript
        Exit
    }

# Set values on '$newWmiClassPath' instances
    try {
        Write-Host "`n### Inserting WMI property values at '$newWmiClassPath'..."
        $newInstance = $class.CreateInstance()
        # Assign Battery Info values from Battery Report values
        $newInstance.ComputerName = $computerName
        $newInstance.SystemManufacturer = $systemManufacturer
        $newInstance.SystemProductName = $systemProductName
        $newInstance.DesignCapacity = $designCapacity
        $newInstance.FullChargeCapacity = $fullChargeCapacity
        $newInstance.CycleCount = $cycleCount
        $newInstance.DesignActiveRuntime = $designActiveRuntime
        $newInstance.FullChargeActiveRuntime = $fullChargeActiveRuntime
        # Save the instance
        $newInstance.Put()
    } catch {
        Write-Error "`n### Failed to insert WMI property values"
        Write-Error "$($_.Exception.Message)"
        Stop-Transcript
        Exit
    }

# Write Summary to Transcript
    Write-Host "`nScript Version = $scriptVer"
    (Get-PSCallStack).InvocationInfo.MyCommand.Name
    Write-Host "`nNew Class: $newWmiClassPath "
    # Get instances of the new BatteryReport CIM class for logging
    $logInstances = Get-CimInstance -Namespace $namespaceParent -ClassName $newWmiClass
    # Define a list of properties to display
    $propertiesToDisplay = @(
        'ComputerName',
        'SystemManufacturer',
        'SystemProductName',
        'DesignCapacity',
        'FullChargeCapacity',
        'CycleCount',
        'DesignActiveRuntime',
        'FullChargeActiveRuntime'
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