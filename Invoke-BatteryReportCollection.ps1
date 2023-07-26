<#
.SYNOPSIS
   Invoke-BatteryReportCollection is a script to collect 'powercfg.exe /batteryreport' information and store it in a custom WMI class.

.DESCRIPTION
   The script creates a custom WMI namespace and class if they don't exist. 
   It then runs 'powercfg /batteryreport' command to generate a battery report in both XML and HTML format. 
   The XML report is parsed to extract battery information which is then stored in the custom WMI class. 
   The reports are saved in the 'C:\temp\batteryreport' directory, and are overwritten each time the script runs.

.PARAMETER None
   This script does not accept any parameters.

.EXAMPLE
   PS C:\> .\Invoke-BatteryReportCollection.ps1
   This example shows how to run the script.

.NOTES
    Version: 1.1
    Creation Date: 2023-07-25
    Copyright (c) 2023 https://github.com/bentman
    https://github.com/bentman/
    Requires: PowerShell V3+. Needs administrative privileges.
   
    NOTE: As this script is complex and handles many operations, thoroughly test it in a safe environment before using it in a production scenario.

.CHANGE
   Version 1.0: Initial script
   Version 1.1: Added function for converting ISO 8601 timspan to HH:MM:SS format
                Added more error handling and logging
   
.LINK
    https://docs.microsoft.com/powershell/scripting/learn/deep-dives/everything-about-powershell-functions?view=powershell-7.1
#>

############################### VARIABLES ###############################

# Define output folder and create if not exists
$reportFolderPath = "C:\temp\batteryreport"

# Define Battery report outputs
$reportPathXml = Join-Path $reportFolderPath "batteryreport.xml"
$reportPathHtml = Join-Path $reportFolderPath "batteryreport.html"
$transcriptPath = Join-Path $reportFolderPath "BatteryReportCollection.log"

# Custom WMI namespace and class parameters
$namespacePath = "root\cimv2\BatteryReport"
$className = "BatteryReport"

############################### FUNCTIONS ###############################

function ConvertTo-StandardTimeFormat {
    param(
        # ISO 8601 duration passed as a string
        [Parameter(Mandatory=$true)][string]$iso8601Duration
    )
    # Return a default value for empty input
    if ([string]::IsNullOrEmpty($iso8601Duration)) {
        return "00:00:00"
    }
    # Regular Expression pattern
    $match = [Regex]::Match($iso8601Duration, 'PT((?<hours>\d+)H)?((?<minutes>\d+)M)?((?<seconds>\d+)S)?')
    # Regular Expression matching
    $hours = [int]$match.Groups['hours'].Value
    $minutes = [int]$match.Groups['minutes'].Value
    $seconds = [int]$match.Groups['seconds'].Value
    # Convert ISO 8601 timespan to standard "HH:MM:SS" format
    $timespan = New-TimeSpan -Hours $hours -Minutes $minutes -Seconds $seconds
    return $timespan.ToString("hh\:mm\:ss")
}

############################### EXECUTION ###############################

# Create output folder if not exist
if (!(Test-Path -Path $reportFolderPath)) {
    New-Item -ItemType Directory -Path $reportFolderPath | Out-Null
}

# Start PowerShell Transcript
Start-Transcript -Path $transcriptPath -Force

# Check if a battery is present in the system
$batteryPresent = Get-CimInstance -ClassName Win32_Battery
if (!$batteryPresent) {
    # If a battery is not found, log an error and exit the script
    Write-Host "Error: No battery detected on this system." -ForegroundColor Red
    Stop-Transcript
    Exit
}

# Script execution wrapped in a Try/Catch
Try { 
    Try { # Create WMI class if not exist
        # Using 'Get-WmiObject' for compatibility with "New-Object"
        $class = Get-WmiObject -Namespace $namespacePath -List | Where-Object {$_.Name -eq $className}
        if (!$class) {
            $class = New-Object System.Management.ManagementClass("$namespacePath", [string]::Empty, $null)
            $class["__CLASS"] = $className
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
        }
    } Catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Stop-Transcript
        Exit
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
    $batteryData.ActiveRuntime = ConvertTo-StandardTimeFormat -iso8601Duration $batteryData.ActiveRuntime
    $batteryData.ActiveRuntimeAtDesignCapacity = ConvertTo-StandardTimeFormat -iso8601Duration $batteryData.ActiveRuntimeAtDesignCapacity
    $batteryData.ModernStandby = ConvertTo-StandardTimeFormat -iso8601Duration $batteryData.ModernStandby
    $batteryData.ModernStandbyAtDesignCapacity = ConvertTo-StandardTimeFormat -iso8601Duration $batteryData.ModernStandbyAtDesignCapacity

    # Store the information into the WMI class
    Set-WmiInstance -Namespace root\cimv2\BatteryReport -Class BatteryReport -Arguments @{
        ComputerName = $env:COMPUTERNAME
        DesignCapacity = $batteryData.DesignCapacity
        FullChargeCapacity = $batteryData.FullChargeCapacity
        CycleCount = $batteryData.CycleCount
        ActiveRuntime = $batteryData.ActiveRuntime
        ActiveRuntimeAtDesignCapacity = $batteryData.ActiveRuntimeAtDesignCapacity
        ModernStandby = $batteryData.ModernStandby
        ModernStandbyAtDesignCapacity = $batteryData.ModernStandbyAtDesignCapacity
    }

    # Get instances of the new BatteryReport CIM class for logging
    $instances = Get-CimInstance -Namespace $namespacePath -ClassName $className

    # Loop through each instance and log the properties
    Write-Host "`nNew Class Name: $className"
    foreach ($instance in $instances) {
        Write-Host "" # empty line for readability
        Write-Host "    ComputerName: $($instance.ComputerName)"
        Write-Host "    DesignCapacity: $($instance.DesignCapacity)"
        Write-Host "    FullChargeCapacity: $($instance.FullChargeCapacity)"
        Write-Host "    CycleCount: $($instance.CycleCount)"
        Write-Host "    ActiveRuntime: $($instance.ActiveRuntime)"
        Write-Host "    ActiveRuntimeAtDesignCapacity: $($instance.ActiveRuntimeAtDesignCapacity)"
        Write-Host "    ModernStandby: $($instance.ModernStandby)"
        Write-Host "    ModernStandbyAtDesignCapacity: $($instance.ModernStandbyAtDesignCapacity)"
        Write-Host "" # empty line for readability
    }

} Catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} Finally {
    Stop-Transcript
}
