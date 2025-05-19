################################################################################################################################################
#
#              Welcome to Blake's Master Dashboard script!  For every dashboard in Portland!
#                                                                                   ~ Made with love. #transrights
#
################################################################################################################################################ 

param (
    [hashtable]$DashboardUrls = @{  # A hashtable mapping computer names to their specific dashboard URLs (can contain multiple URLs)
        "Stord-BNA-BNAB7" = @(
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1317&D=2"
        )
        "Stord-BNA-BNAB8" = @(
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1163&D=3"
        )
        "Stord-BNA-BNAB9" = @(
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1288&D=4"
        )
        "Stord-BNA-0H4BT" = @(
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&D=1",
            "https://www.p3pl.com/Shipping/DashboardWarehouseMoves.aspx?P=3849ctyn2938hycfg3487ypqa1v2&W=19",
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1317&D=2",
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1163&D=3",
            "https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1288&D=4",
            "https://www.wrike.com/workspace.htm?acc=2577757#/dashboards/9701292"
        )
    },
    [hashtable]$ComputerSpecificSettings = @{  # A hashtable for per-computer variable overrides
        "Stord-BNA-0H4BT" = @{ 
            "RefreshIntervalMinutes" = 30
            "TabSwitchIntervalSeconds" = 15
        }
    },
    [int]$DefaultRefreshIntervalMinutes = 5,                                      # The interval in minutes between each refresh.
    [int]$DefaultChromeStartDelaySeconds = 10,                                    # Default delay in seconds before starting Chrome after stopping any running instances.
    [int]$DefaultTabSwitchIntervalSeconds = 60,                                   # Interval in minutes between tab switches. (only applicable if multiple URLs are open)
    [int]$DefaultLoopIntervalSeconds = 5                                          # Keeps the loop refreshing at a set interval.
)

# Global variables (Shouldn't be edited)
$refreshCounter = 0                                                               # Countdown to the next refresh.
$loopCounter = 0                                                                  # Counts how many times the script has looped.
$tabSwitchCounter = 0                                                             # Counter to keep track of tab switch interval.
$currentTabIndex = 0                                                              # Index of the current tab being viewed.
$loopIntervalSeconds = $DefaultLoopIntervalSeconds                                # Keeps the loop refreshing at a set interval.
$refreshed = 0                                                                    # Track the number of times the dashboards have been refreshed.
$logFilePath = "\\int3\shared\BlakeC\Script Logs\master_dashboard_script_log.txt" # Logging file - to monitor any errors.
$fallbackLogFilePath = "C:\scripts\master_dashboard_script_log.txt"               # Fall back Logging file

################################################################################################################################################

# Get the current computer name
$ComputerName = $env:COMPUTERNAME

# Apply per-computer settings if available
if ($ComputerSpecificSettings.ContainsKey($ComputerName)) {
    $settings = $ComputerSpecificSettings[$ComputerName]
    if ($settings.ContainsKey("RefreshIntervalMinutes")) {
        $RefreshIntervalMinutes = $settings["RefreshIntervalMinutes"]
    } else {
        $RefreshIntervalMinutes = $DefaultRefreshIntervalMinutes
    }
    if ($settings.ContainsKey("ChromeStartDelaySeconds")) {
        $ChromeStartDelaySeconds = $settings["ChromeStartDelaySeconds"]
    } else {
        $ChromeStartDelaySeconds = $DefaultChromeStartDelaySeconds
    }
    if ($settings.ContainsKey("TabSwitchIntervalSeconds")) {
        $TabSwitchIntervalSeconds = $settings["TabSwitchIntervalSeconds"]
    } else {
        $TabSwitchIntervalSeconds = $DefaultTabSwitchIntervalSeconds
    }
} else {
    $RefreshIntervalMinutes = $DefaultRefreshIntervalMinutes
    $ChromeStartDelaySeconds = $DefaultChromeStartDelaySeconds
    $TabSwitchIntervalSeconds = $DefaultTabSwitchIntervalSeconds
}

################################################################################################################################################
# All the real code is below!
################################################################################################################################################


# Load System.Windows.Forms to use Cursor methods
Add-Type -AssemblyName System.Windows.Forms

# Function to ensure Scroll Lock is in the desired state
function Set-ScrollLock {
    param ($desiredState)
    $wshell = New-Object -ComObject wscript.shell
    if ([console]::CapsLock -ne $desiredState) {
        $wshell.SendKeys("{SCROLLLOCK}")
    }
}

# Fallback logging directory
$logDir = Split-Path $fallbackLogFilePath
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Function to write log messages to the log file
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $env:COMPUTERNAME -> $message"
    
    # Check if the primary log file path is available
    if (Test-Path -Path (Split-Path $logFilePath)) {
        Add-Content -Path $logFilePath -Value $logMessage
    }
    else {
        Add-Content -Path $fallbackLogFilePath -Value $logMessage
    }
}



# Define the user32.dll function
Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class Taskbar {
        [DllImport("user32.dll")]
        public static extern int FindWindow(string className, string windowText);
        [DllImport("user32.dll")]
        public static extern int ShowWindow(int hwnd, int command);

        public static void HideTaskbar() {
            int taskbarHandle = FindWindow("Shell_TrayWnd", "");
            ShowWindow(taskbarHandle, 0); // 0 = Hide
        }

        public static void ShowTaskbar() {
            int taskbarHandle = FindWindow("Shell_TrayWnd", "");
            ShowWindow(taskbarHandle, 5); // 5 = Show
        }
    }
"@





Write-Log "Script started."

# Set Scroll Lock on and hide the mouse cursor
Set-ScrollLock -desiredState $true
[System.Windows.Forms.Cursor]::Hide()  # Hide the mouse!
# Hide the taskbar
[Taskbar]::HideTaskbar()

# Stop Chrome if running
$chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
if ($chromeProcess) {
    Stop-Process -Name "chrome" -Force
}

Start-Sleep -Seconds $ChromeStartDelaySeconds

# Replace exit_type if it's Crashed to Normal in Chrome preferences
$chromePrefsPath = "$env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default\Preferences"
if (Test-Path $chromePrefsPath) {
    (Get-Content -Path $chromePrefsPath) -replace '"exit_type":"Crashed"', '"exit_type":"Normal"' | Set-Content -Path $chromePrefsPath
} else {
    Write-Log "Chrome preferences not found. Cannot set exit_type to Normal."
}

# Open the appropriate dashboard URLs based on computer name
if ($DashboardUrls.ContainsKey($ComputerName)) {
    $urls = $DashboardUrls[$ComputerName]
    foreach ($url in $urls) {
        try {
            Start-Process -FilePath "chrome" -ArgumentList "--kiosk", $url -ErrorAction Stop
            Start-Sleep -Seconds 1 # Short delay between opening tabs
        } catch {
            Write-Log "Failed to start Chrome with URL $url. Error: $_"
        }
    }
} else {
    Write-Log "Computer name not found in Dashboard URLs."
    exit
}

# Activate Chrome and make it fullscreen
$wshell = New-Object -ComObject wscript.shell
$wshell.AppActivate('chrome')

# Main Script loop
while ($true) {
    # Check if Chrome is still running, if not, restart it
    if (-not (Get-Process -Name "chrome" -ErrorAction SilentlyContinue)) {
        Write-Log "Chrome not running. Restarting..."
        foreach ($url in $urls) {
            Start-Process -FilePath "chrome" -ArgumentList "--kiosk", $url
            Start-Sleep -Seconds 1
        }
    }

    $loopCounter += $loopIntervalSeconds
    $tabSwitchCounter += $loopIntervalSeconds
    $refreshCounter += $loopIntervalSeconds

    # Refresh the dashboard at the specified interval
    if ($refreshCounter -ge ($RefreshIntervalMinutes * 60)) {
        $wshell.SendKeys('^{f5}')
        $refreshCounter = 0
        $refreshed += 1
        $uptime = (($loopCounter) / 3600) / 24
        #Write-Log "Refreshing Web pages: Refreshed $refreshed times | Uptime: $([math]::Round($uptime, 2)) days"
    }

    # Switch tabs at the specified interval if there are multiple URLs
    if ($urls.Count -gt 1 -and $tabSwitchCounter -ge ($TabSwitchIntervalSeconds)) {
        $wshell.SendKeys('^{TAB}')
        $tabSwitchCounter = 0
        $currentTabIndex = ($currentTabIndex + 1) % $urls.Count
        #Write-Log "Switched to tab $($currentTabIndex + 1)"
    }

    # Reset the log file at the end of the month
    if ((Get-Date).Day -eq 1 -and (Get-Date).Hour -eq 0 -and (Get-Date).Minute -lt $loopIntervalSeconds) {
        Clear-Content -Path $logFilePath
        Write-Log "Log file has been reset. (beginning of the month)"
    }

    Start-Sleep -Seconds $loopIntervalSeconds
}





# Gracefully exit to show mouse cursor and turn off Scroll Lock
Register-EngineEvent PowerShell.Exiting -Action {
    [System.Windows.Forms.Cursor]::Show()  # Show the mouse when exiting
    [Taskbar]::ShowTaskbar() # Show the taskbar
    Set-ScrollLock -desiredState $false    # Turn off Scroll Lock
    Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue

    Write-Log "Dashboard script stopped gracefully."
}
