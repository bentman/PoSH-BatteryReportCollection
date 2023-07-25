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
    Version: 1.0
    Creation Date: 2023-07-25
    Copyright (c) 2023 https://github.com/bentman
    https://github.com/bentman/
    Requires: PowerShell V3+. Needs administrative privileges.
   
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

############################### EXECUTION ###############################

# Create output folder if not exist
if (!(Test-Path -Path $reportFolderPath)) {
    New-Item -ItemType Directory -Path $reportFolderPath | Out-Null
}

# Start PowerShell Transcript
Start-Transcript -Path $transcriptPath -Force

# Script execution wrapped in a Try/Catch
Try { 
    Try { # Create WMI class if not exists
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
    }
    Catch {
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

    # Convert ActiveRuntime from PT,H,M,S format to HH:MM:SS
    $batteryData.ActiveRuntime = ([timespan]::Parse($batteryData.ActiveRuntime)).ToString("hh\:mm\:ss")
    $batteryData.ActiveRuntimeAtDesignCapacity = ([timespan]::Parse($batteryData.ActiveRuntimeAtDesignCapacity)).ToString("hh\:mm\:ss")
    $batteryData.ModernStandby = ([timespan]::Parse($batteryData.ModernStandby)).ToString("hh\:mm\:ss")
    $batteryData.ModernStandbyAtDesignCapacity = ([timespan]::Parse($batteryData.ModernStandbyAtDesignCapacity)).ToString("hh\:mm\:ss")

    # Store the information into the WMI class
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
} Catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} Finally {
    Stop-Transcript
}
