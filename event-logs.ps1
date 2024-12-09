# Capture the script start time
$ScriptStartTime = Get-Date

# Set output location for CSV files.
$LogOutputDirectory = 'C:\Users\Asus\Downloads\ps1'

# Ensure the output directory exists
if (-not (Test-Path -Path $LogOutputDirectory)) {
    Write-Output "Output directory does not exist. Creating it at $LogOutputDirectory."
    New-Item -ItemType Directory -Path $LogOutputDirectory | Out-Null
}

# Define Windows event log types to export.
$EventTypesToExport = @('Application', 'Security', 'Setup', 'System', 'ForwardedEvents')

# Process each log type
foreach ($EventType in $EventTypesToExport) {
    try {
        # Build the file path for the current log type.
        $LogOutputTopic = "Windows_Event_Log_$EventType"
        $CurrentTimeUTC = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        $LogOutputFileName = "$CurrentTimeUTC_$LogOutputTopic.csv"
        $LogOutputCSVFilePath = Join-Path -Path $LogOutputDirectory -ChildPath $LogOutputFileName

        # Inform the user about the process
        Write-Output "Creating CSV for Windows event log type: $EventType."
        Write-Output "Target CSV file path: $LogOutputCSVFilePath"

        # Export the event log to a CSV file
        Get-WinEvent -LogName $EventType -ErrorAction Stop | Export-Csv -Path $LogOutputCSVFilePath -NoTypeInformation -Force

        Write-Output "Finished creating CSV file for $EventType log."
    } catch {
        # Handle errors (e.g., log type not accessible)
        Write-Output "Failed to export $EventType log. Error: $_"
    }
}

# Calculate and display the script run time
$ScriptEndTime = Get-Date
$ScriptDuration = $ScriptEndTime - $ScriptStartTime
Write-Output "Log export process completed in: $ScriptDuration."