# Define the Namespace and Class Name
$namespacePath = "root\cimv2\BatteryReport"
$newClassName = "BatteryReport"
$logFolder = "C:\temp\BatteryReport\"
$logFile = Join-Path -Path $logFolder -ChildPath "Cleanup.log"

# Start the transcript
try {
    Start-Transcript -Path $logFile -ErrorAction Stop
    Write-Host "Transcript started, output file is $logFile"
}
catch {
    Write-Host "Failed to start transcript. Please make sure that the path $logFile is valid and accessible."
    return
}

# Check if the class exists
try {
    $class = Get-CimClass -Namespace $namespacePath -ClassName $newClassName -ErrorAction Stop
}
catch {
    Write-Host "The class $newClassName does not exist in the namespace $namespacePath."
    Stop-Transcript
    return
}

# If the class exists, then remove it
try {
    Remove-CimClass -CimClass $class -ErrorAction Stop
    Write-Host "Successfully removed the class $newClassName from the namespace $namespacePath."
}
catch {
    Write-Host "Failed to remove the class $newClassName from the namespace $namespacePath."
    Write-Host $_.Exception.Message
}

# Stop the transcript
Stop-Transcript
