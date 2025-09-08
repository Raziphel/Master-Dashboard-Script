################################################################################################################################################
#
#              Welcome to Blake's Master Dashboard script!  For every dashboard in Portland!
#                                                                                   ~ Made with love. #transrights
#
################################################################################################################################################

param (
	[hashtable]$DashboardUrls = @{  # A hashtable mapping computer names to their specific dashboard URLs (can contain multiple URLs)
		"PS-BNA1-87E9A0052A74" = @(
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1317&D=2"
		)
		"PS-BNA1-A324BA4C0059" = @(
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1163&D=3"
		)
		"PS-BNA1-B7CC656D0fEB" = @(
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1288&D=4"
		)
		"PS-BNA1-D0H4BT2" = @(
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&D=1",
			"https://www.p3pl.com/Shipping/DashboardWarehouseMoves.aspx?P=3849ctyn2938hycfg3487ypqa1v2&W=19",
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1317&D=2",
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1163&D=3",
			"https://www.p3pl.com/Shipping/DashboardBasicClientV1.aspx?P=3849ctyn4365hycfg3487ypqa1cv1&W=19&ClientID=1288&D=4",
			"https://www.wrike.com/workspace.htm?acc=2577757#/dashboards/9701292"
		)
	},
	[hashtable]$ComputerSpecificSettings = @{  # A hashtable for per-computer variable overrides
		"PS-BNA1-D0H4BT2" = @{
			"RefreshIntervalMinutes" = 30
			"TabSwitchIntervalSeconds" = 15
		}
	},
	[int]$DefaultRefreshIntervalMinutes = 5,         # The interval in minutes between each refresh.
	[int]$DefaultChromeStartDelaySeconds = 10,       # Default delay in seconds before starting Chrome after stopping any running instances.
	[int]$DefaultTabSwitchIntervalSeconds = 60,      # Interval in SECONDS between tab switches. (only applicable if multiple URLs are open)
	[int]$DefaultLoopIntervalSeconds = 5             # Keeps the loop refreshing at a set interval (seconds).
)

# =========================================
# Global variables (Shouldn't be edited)
# =========================================
$refreshCounter      = 0
$loopCounter         = 0
$tabSwitchCounter    = 0
$currentTabIndex     = 0
$loopIntervalSeconds = $DefaultLoopIntervalSeconds
$refreshed           = 0
$logFilePath         = "\\int3\shared\BlakeC\Script Logs\master_dashboard_script_log.txt"
$fallbackLogFilePath = "C:\scripts\master_dashboard_script_log.txt"

################################################################################################################################################
# Helpers: Logging, Name Normalization, Scroll Lock, Taskbar
################################################################################################################################################

# Load System.Windows.Forms (for cursor + key state)
Add-Type -AssemblyName System.Windows.Forms

# Ensure fallback logging directory exists
$logDir = Split-Path $fallbackLogFilePath
if (-not (Test-Path $logDir)) {
	New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
	param([string]$message)
	$timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$logMessage = "[$timestamp] $env:COMPUTERNAME -> $message"

	# Try primary directory; fall back to local
	$primaryDir = Split-Path $logFilePath
	if (Test-Path -Path $primaryDir) {
		Add-Content -Path $logFilePath -Value $logMessage
	} else {
		Add-Content -Path $fallbackLogFilePath -Value $logMessage
	}
}

# Normalize a name to 15-char NetBIOS uppercase (what Windows actually uses for COMPUTERNAME)
function Get-NetBIOS15 {
	param([string]$Name)
	return $Name.Substring(0, [Math]::Min(15, $Name.Length)).ToUpperInvariant()
}

# Proper Scroll Lock state handling
function Set-ScrollLock {
	param([bool]$desiredState)
	# .NET can read ScrollLock via Control.IsKeyLocked
	$current = [System.Windows.Forms.Control]::IsKeyLocked([System.Windows.Forms.Keys]::Scroll)
	if ($current -ne $desiredState) {
		$wshell = New-Object -ComObject wscript.shell
		$wshell.SendKeys("{SCROLLLOCK}")
	}
}

# user32 wrappers for taskbar show/hide
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
			ShowWindow(taskbarHandle, 5); // 5 = Show (SW_SHOW)
		}
	}
"@

################################################################################################################################################
# Computer Name & Per-Computer Settings (normalized)
################################################################################################################################################

# Get the current computer name and 15-char normalized version
$ComputerName       = $env:COMPUTERNAME
$ComputerNameShort  = Get-NetBIOS15 $ComputerName

# Build normalized lookup tables (keys collapsed to 15-char uppercase)
$DashboardUrls15 = @{}
foreach ($k in $DashboardUrls.Keys) {
	$short = Get-NetBIOS15 $k
	$DashboardUrls15[$short] = $DashboardUrls[$k]
}

$ComputerSpecificSettings15 = @{}
foreach ($k in $ComputerSpecificSettings.Keys) {
	$short = Get-NetBIOS15 $k
	$ComputerSpecificSettings15[$short] = $ComputerSpecificSettings[$k]
}

# Apply per-computer settings if available; else defaults
if ($ComputerSpecificSettings15.ContainsKey($ComputerNameShort)) {
	$settings = $ComputerSpecificSettings15[$ComputerNameShort]

	if ($settings.ContainsKey("RefreshIntervalMinutes")) { $RefreshIntervalMinutes   = $settings["RefreshIntervalMinutes"] }   else { $RefreshIntervalMinutes   = $DefaultRefreshIntervalMinutes }
	if ($settings.ContainsKey("ChromeStartDelaySeconds")) { $ChromeStartDelaySeconds  = $settings["ChromeStartDelaySeconds"] }  else { $ChromeStartDelaySeconds  = $DefaultChromeStartDelaySeconds }
	if ($settings.ContainsKey("TabSwitchIntervalSeconds")) { $TabSwitchIntervalSeconds = $settings["TabSwitchIntervalSeconds"] } else { $TabSwitchIntervalSeconds = $DefaultTabSwitchIntervalSeconds }
} else {
	$RefreshIntervalMinutes   = $DefaultRefreshIntervalMinutes
	$ChromeStartDelaySeconds  = $DefaultChromeStartDelaySeconds
	$TabSwitchIntervalSeconds = $DefaultTabSwitchIntervalSeconds
}

################################################################################################################################################
# Startup Actions
################################################################################################################################################

Write-Log "Script started. COMPUTERNAME='$ComputerName' SHORT='$ComputerNameShort'"

# Toggle Scroll Lock on, hide cursor, hide taskbar
Set-ScrollLock -desiredState $true
[System.Windows.Forms.Cursor]::Hide()
[Taskbar]::HideTaskbar()

# Ensure graceful cleanup happens on exit (register BEFORE the loop)
Register-EngineEvent PowerShell.Exiting -Action {
	try {
		[System.Windows.Forms.Cursor]::Show()
		[Taskbar]::ShowTaskbar()
		Set-ScrollLock -desiredState $false
		Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
		Write-Log "Dashboard script stopped gracefully."
	} catch {
		# Swallow any shutdown errors
	}
} | Out-Null

# Stop Chrome if running
$chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
if ($chromeProcess) {
	Stop-Process -Name "chrome" -Force
}

Start-Sleep -Seconds $ChromeStartDelaySeconds

# Replace exit_type if it's Crashed to Normal in Chrome preferences
$chromePrefsPath = "$env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default\Preferences"
if (Test-Path $chromePrefsPath) {
	try {
		(Get-Content -Path $chromePrefsPath -Raw) -replace '"exit_type"\s*:\s*"Crashed"', '"exit_type":"Normal"' | Set-Content -Path $chromePrefsPath
	} catch {
		Write-Log "Failed to update Chrome preferences at '$chromePrefsPath'. Error: $_"
	}
} else {
	Write-Log "Chrome preferences not found. Cannot set exit_type to Normal. Path: $chromePrefsPath"
}

# Open the appropriate dashboard URLs based on computer name
if ($DashboardUrls15.ContainsKey($ComputerNameShort)) {
	$urls = $DashboardUrls15[$ComputerNameShort]
	foreach ($url in $urls) {
		try {
			Start-Process -FilePath "chrome" -ArgumentList "--kiosk", $url -ErrorAction Stop
			Start-Sleep -Seconds 1 # small delay so tabs reliably stack
		} catch {
			Write-Log "Failed to start Chrome with URL $url. Error: $_"
		}
	}
} else {
	Write-Log "Computer name '$ComputerName' (short '$ComputerNameShort') not found in Dashboard URLs. Exiting."
	exit
}

# Try to activate Chrome
$wshell = New-Object -ComObject wscript.shell
$null = $wshell.AppActivate('chrome')

################################################################################################################################################
# Main Loop
################################################################################################################################################

while ($true) {
	# If Chrome died, restart and re-open URLs
	if (-not (Get-Process -Name "chrome" -ErrorAction SilentlyContinue)) {
		Write-Log "Chrome not running. Restarting..."
		foreach ($url in $urls) {
			try {
				Start-Process -FilePath "chrome" -ArgumentList "--kiosk", $url
				Start-Sleep -Seconds 1
			} catch {
				Write-Log "Failed to restart Chrome with URL $url. Error: $_"
			}
		}
		$null = $wshell.AppActivate('chrome')
	}

	$loopCounter      += $loopIntervalSeconds
	$tabSwitchCounter += $loopIntervalSeconds
	$refreshCounter   += $loopIntervalSeconds

	# Refresh the dashboards at the specified interval
	if ($refreshCounter -ge ($RefreshIntervalMinutes * 60)) {
		$wshell.SendKeys('^{F5}')
		$refreshCounter = 0
		$refreshed += 1
		# $uptimeDays = [math]::Round(($loopCounter / 86400), 2)
		# Write-Log "Refreshed $refreshed times | Uptime: $uptimeDays days"
	}

	# Switch tabs at the specified interval if there are multiple URLs
	if ($urls.Count -gt 1 -and $tabSwitchCounter -ge $TabSwitchIntervalSeconds) {
		$wshell.SendKeys('^{TAB}')
		$tabSwitchCounter = 0
		$currentTabIndex = ($currentTabIndex + 1) % $urls.Count
		# Write-Log "Switched to tab $($currentTabIndex + 1)"
	}

	# Reset the log file at the beginning of the month (00:00 on the 1st)
	if ((Get-Date).Day -eq 1 -and (Get-Date).Hour -eq 0 -and (Get-Date).Minute -lt $loopIntervalSeconds) {
		if (Test-Path $logFilePath) {
			try { Clear-Content -Path $logFilePath } catch { Write-Log "Failed to clear primary log: $_" }
		}
		if (Test-Path $fallbackLogFilePath) {
			try { Clear-Content -Path $fallbackLogFilePath } catch { Write-Log "Failed to clear fallback log: $_" }
		}
		Write-Log "Log file(s) have been reset. (beginning of the month)"
	}

	Start-Sleep -Seconds $loopIntervalSeconds
}
