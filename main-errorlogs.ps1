# Capture the script start time
$ScriptStartTime = Get-Date

# Set output location for CSV files.
$LogOutputDirectory = 'C:\Users\Asus\Downloads\ps1'

# Ensure the output directory exists
if (-not (Test-Path -Path $LogOutputDirectory)) {
    Write-Output "Output directory does not exist. Creating it at $LogOutputDirectory."
    New-Item -ItemType Directory -Path $LogOutputDirectory | Out-Null
}

# Define Windows event log types to export
$EventTypesToExport = @('Application', 'Security', 'Setup', 'System', 'ForwardedEvents')

# Define specific Event IDs and sources related to potential system issues (for imaging purposes)
$CriticalEventIDs = @(41, 1001, 1000, 1026, 129, 4625, 7040, 20, 30)  # Event IDs related to crashes, kernel issues, and failures
$EventSourcesToMonitor = @('Microsoft-Windows-Kernel-Power', 'Microsoft-Windows-WindowsUpdateClient')

# Define a log file for errors during script execution
$ErrorLogFilePath = Join-Path -Path $LogOutputDirectory -ChildPath 'ErrorLog.txt'

# Function to log errors to a file
function Log-Error {
    param([string]$ErrorMessage)
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$Timestamp - $ErrorMessage" | Out-File -FilePath $ErrorLogFilePath -Append
}

# Function to check if the script is running with Administrator privileges
function Check-AdminRights {
    if (-not (Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System")) {
        Write-Output "Script is not running with Administrator privileges. Please run as Administrator."
        Log-Error "Script is not running with Administrator privileges."
        exit
    }
}

# Check if the script is running with Administrator privileges
Check-AdminRights

# Process each log type
foreach ($EventType in $EventTypesToExport) {
    try {
        # Build the file path for the current log type
        $LogOutputTopic = "Windows_Event_Log_$EventType"
        $CurrentTimeUTC = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        $LogOutputFileName = "$CurrentTimeUTC_$LogOutputTopic.csv"
        $LogOutputCSVFilePath = Join-Path -Path $LogOutputDirectory -ChildPath $LogOutputFileName

        # Inform the user about the process
        Write-Output "Creating CSV for Windows event log type: $EventType."
        Write-Output "Target CSV file path: $LogOutputCSVFilePath"

        # Attempt to export the event log
        Get-WinEvent -LogName $EventType -ErrorAction Stop | Export-Csv -Path $LogOutputCSVFilePath -NoTypeInformation -Force

        Write-Output "Finished creating CSV file for $EventType log."

        # Additional logic to monitor critical events (e.g., Kernel errors or update failures)
        if ($EventType -eq 'System') {
            Write-Output "Filtering critical events for imaging purposes..."

            # Fetch and filter critical events
            $CriticalEvents = Get-WinEvent -LogName 'System' | Where-Object {
                $_.Id -in $CriticalEventIDs -or $EventSourcesToMonitor -contains $_.ProviderName
            }

            if ($CriticalEvents) {
                $CriticalEventFilePath = Join-Path -Path $LogOutputDirectory -ChildPath "Critical_Events_$CurrentTimeUTC.csv"
                $CriticalEvents | Export-Csv -Path $CriticalEventFilePath -NoTypeInformation -Force
                Write-Output "Critical events saved to: $CriticalEventFilePath"
            } else {
                Write-Output "No critical events found."
            }
        }
    } catch {
        # Handle errors (e.g., log type not accessible)
        $ErrorMessage = "Failed to export $EventType log. Error: $_"
        Write-Output $ErrorMessage
        Log-Error $ErrorMessage
    }
}

# Calculate and display the script run time
$ScriptEndTime = Get-Date
$ScriptDuration = $ScriptEndTime - $ScriptStartTime
Write-Output "Log export process completed in: $ScriptDuration"
