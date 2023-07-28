############################## VARIABLES ###############################
$scriptVer = "2.0"
$newClassName = "BatteryReport"
$reportFolder = "C:\temp"

############################### GENERATED ###############################
$reportFolderPath = Join-Path $reportFolder "$newClassName"
$reportPathXml = Join-Path $reportFolderPath "BatteryReport.xml"
$reportPathHtml = Join-Path $reportFolderPath "BatteryReport.html"
$transcriptPath = Join-Path $reportFolderPath "BatteryReportCollection.log"
$namespacePath = "root\cimv2\$newClassName"

############################### FUNCTION  ###############################
# Function to convert ISO 8601 duration to standard time format
function ConvertTo-StandardTimeFormat {
    param(
        [Parameter(Mandatory=$true)]
        [string] $isoDuration
    )

    # Check for null or whitespace and return '00:00:00' if true
    if ([string]::IsNullOrWhiteSpace($isoDuration)) {
        return '00:00:00'
    }

    # Standard pattern matching for ISO 8601 duration format
    $pattern = 'P(?<years>\d+Y)?(?<months>\d+M)?(?<days>\d+D)?(T(?<hours>\d+H)?(?<minutes>\d+M)?(?<seconds>\d+(\.\d+)?S)?)?'

    if ($isoDuration -match $pattern) {
        $years = [int]$Matches.years.TrimEnd('Y')
        $months = [int]$Matches.months.TrimEnd('M')
        $days = [int]$Matches.days.TrimEnd('D')
        $hours = [int]$Matches.hours.TrimEnd('H')
        $minutes = [int]$Matches.minutes.TrimEnd('M')
        $seconds = [double]$Matches.seconds.TrimEnd('S')

        # Calculate total time in seconds
        $totalDays = $years * 365 + $months * 30 + $days
        $totalHours = $totalDays * 24 + $hours
        $totalMinutes = $totalHours * 60 + $minutes
        $totalSeconds = $totalMinutes * 60 + $seconds

        # Return in standard time format
        return "{0:D2}:{1:D2}:{2:D2}" -f $totalHours, $totalMinutes, [Math]::Round($totalSeconds)
    } else {
        return '00:00:00'
    }
}

############################### EXECUTION ###############################
try {
    # Create output folder if it does not exist
    if (-not (Test-Path -Path $reportFolderPath)) {
        New-Item -ItemType Directory -Path $reportFolderPath -ErrorAction Stop | Out-Null
    }

    # Start the transcript
    Start-Transcript -Path $transcriptPath -ErrorAction Stop
    Write-Host "Script Version = $scriptVer"
    (Get-PSCallStack).InvocationInfo.MyCommand.Name

    <# Check if a battery is present in the system
    $batteryPresent = Get-WmiObject -Class Win32_Battery -ErrorAction Stop
    if ($null -eq $batteryPresent) {
        Write-Host "Error: No battery detected on this system - Skipping all operations." -ForegroundColor Red
        Stop-Transcript
        exit
    }#>

    # Generate HTML battery report for readability
    powercfg /batteryreport /output $reportPathHtml
    Write-Host "Battery report generated at $reportPathHtml"

    # Generate XML battery report for script consumption
    powercfg /batteryreport /output $reportPathXml /XML
    Write-Host "Battery report generated at $reportPathXml"

    # Parse battery report
    [xml]$batteryReport = Get-Content $reportPathXml
    $designCapacity = $batteryReport.BatteryReport.Report.DesignCapacity.mWh
    $fullChargeCapacity = $batteryReport.BatteryReport.Report.FullChargeCapacity.mWh
    $cycleCount = $batteryReport.BatteryReport.Report.CycleCount.Count
    $activeTimeAcValue = $batteryReport.BatteryReport.Report.Active.PowerStateTimeAc.Value
    $activeRuntime = if ([string]::IsNullOrWhiteSpace($activeTimeAcValue)) {'00:00:00'} else {ConvertTo-StandardTimeFormat $activeTimeAcValue}
    $modernStandby = if ([string]::IsNullOrWhiteSpace($modernStandbyDuration)) {'00:00:00'} else {ConvertTo-StandardTimeFormat $modernStandbyDuration}
        
    # Get WMI namespace if it doesn't exist
    $namespace = Get-WmiObject -Namespace root\cimv2 -Class __Namespace -Filter "Name = '$newClassName'" -ErrorAction Stop
    if ($null -eq $namespace) {
        # This assumes that the current user has rights to create namespaces
        $namespace = ([wmiclass]"root\cimv2:__Namespace").CreateInstance()
        $namespace.Name = $newClassName
        $namespace.Put()
    }

    # WMI class properties
    $classProperties = @{
        FullChargeCapacity = $fullChargeCapacity
        DesignCapacity = $designCapacity
        CycleCount = $cycleCount
        ActiveTimeAcValue = $activeTimeAcValue
        ModernStandby = $modernStandby
        ActiveRuntime = $activeRuntime
    }

    # Create WMI class if it doesn't exist
    $class = Get-WmiObject -Namespace $namespacePath -List | Where-Object {$_.Name -eq $newClassName} -ErrorAction Stop
    if ($null -eq $class) {
        # Create new class
        $class = New-Object System.Management.ManagementClass($namespacePath, [String]::Empty, $null)
        $class["__CLASS"] = $newClassName
        foreach ($prop in $classProperties.Keys) {
            $class.Properties.Add($prop, [System.Management.CimType]::String, $false)
        }
        $class.Put()
    }

    # Create or update an instance of WMI class
    $instance = Get-CimInstance -Namespace $namespacePath -ClassName $newClassName -ErrorAction SilentlyContinue
    if ($null -eq $instance) {
        # Create new instance
        $newInstance = New-Object -TypeName PSObject
        foreach ($prop in $classProperties.Keys) {
            Add-Member -InputObject $newInstance -NotePropertyName $prop -NotePropertyValue $classProperties[$prop]
        }
        New-CimInstance -Namespace $namespacePath -ClassName $newClassName -Property $newInstance.PSObject.Properties
    } else {
        # Update existing instance
        foreach ($prop in $classProperties.Keys) {
            $instance | Set-CimInstance -Property @{ $prop = $classProperties[$prop] }
        }
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Get instances of the new BatteryReport CIM class for logging
    $logInstances = Get-CimInstance -Namespace $namespacePath -ClassName $newClassName
    Write-Host "`nNew Class: $newClassName"
    # Define a list of properties to display
    $propertiesToDisplay = @(
        'fullChargeCapacity',
        'designCapacity',
        'cycleCount',
        'activeTimeAcValue',
        'modernStandby'
        'activeRuntime',
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
