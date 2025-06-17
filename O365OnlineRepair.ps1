# Setup logging
$ComputerName = $env:COMPUTERNAME
$LogFile = "C:\Windows\Temp\OfficeRepair_$ComputerName_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

Function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp [$Level] - $Message"
    Write-Output $entry
    Add-Content -Path $LogFile -Value $entry
}

Write-Log "Starting Office Online Repair check..."

# Check if Office Click-to-Run is installed
$OfficePath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
if (-Not (Test-Path $OfficePath)) {
    Write-Log "Office Click-to-Run not found. Skipping repair." "WARN"
    exit 0
}

# Check if repair is already running
$repairRunning = Get-Process -Name OfficeClickToRun -ErrorAction SilentlyContinue
if ($repairRunning) {
    Write-Log "Office repair already in progress. Exiting." "WARN"
    exit 0
}

# Start the silent online repair
Write-Log "Running silent Office Online Repair..."

Start-Process $OfficePath -ArgumentList "/repairmode=online /forceappshutdown /quiet" -Wait

# Confirm completion
Write-Log "Office repair command executed successfully."

exit 0  # success
