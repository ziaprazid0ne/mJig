# Helper function to get key name from key code
function Get-KeyName {
	param([int]$keyCode)
	
	# Common virtual key codes to names mapping
	$KeyMap = @{
		0x08 = "Backspace"; 0x09 = "Tab"; 0x0C = "Clear"; 0x0D = "Enter"
		0x10 = "Shift"; 0x11 = "Ctrl"; 0x12 = "Alt"; 0x13 = "Pause"
		0x14 = "CapsLock"; 0x1B = "Esc"; 0x20 = "Space"; 0x21 = "PageUp"
		0x22 = "PageDown"; 0x23 = "End"; 0x24 = "Home"; 0x25 = "Left"
		0x26 = "Up"; 0x27 = "Right"; 0x28 = "Down"; 0x2C = "PrintScreen"
		0x2D = "Insert"; 0x2E = "Delete"; 0x5B = "LWin"; 0x5C = "RWin"
		0x5D = "Apps"; 0x5F = "Sleep"; 0x70 = "F1"; 0x71 = "F2"
		0x72 = "F3"; 0x73 = "F4"; 0x74 = "F5"; 0x75 = "F6"
		0x76 = "F7"; 0x77 = "F8"; 0x78 = "F9"; 0x79 = "F10"
		0x7A = "F11"; 0x7B = "F12"; 0x7C = "F13"; 0x7D = "F14"
		0x7E = "F15"; 0x7F = "F16"; 0x80 = "F17"; 0x81 = "F18"
		0x82 = "F19"; 0x83 = "F20"; 0x84 = "F21"; 0x85 = "F22"
		0x86 = "F23"; 0x87 = "F24"; 0x90 = "NumLock"; 0x91 = "ScrollLock"
		0xA0 = "LShift"; 0xA1 = "RShift"; 0xA2 = "LCtrl"; 0xA3 = "RCtrl"
		0xA4 = "LAlt"; 0xA5 = "RAlt"; 0xA6 = "BrowserBack"; 0xA7 = "BrowserForward"
		0xA8 = "BrowserRefresh"; 0xA9 = "BrowserStop"; 0xAA = "BrowserSearch"
		0xAB = "BrowserFavorites"; 0xAC = "BrowserHome"; 0xAD = "VolumeMute"
		0xAE = "VolumeDown"; 0xAF = "VolumeUp"; 0xB0 = "MediaNext"
		0xB1 = "MediaPrev"; 0xB2 = "MediaStop"; 0xB3 = "MediaPlay"
		0xB4 = "LaunchMail"; 0xB5 = "LaunchMedia"; 0xB6 = "LaunchApp1"
		0xB7 = "LaunchApp2"; 0xBA = ";"; 0xBB = "="; 0xBC = ","
		0xBD = "-"; 0xBE = "."; 0xBF = "/"; 0xC0 = "Backtick"
		0xDB = "["; 0xDC = "Backslash"; 0xDD = "]"; 0xDE = "Quote"
	}
	
	# Check if we have a mapped name
	if ($KeyMap.ContainsKey($keyCode)) {
		return $KeyMap[$keyCode]
	}
	
	# For alphanumeric keys (0x30-0x39 for 0-9, 0x41-0x5A for A-Z)
	if ($keyCode -ge 0x30 -and $keyCode -le 0x39) {
		return [char]($keyCode)
	}
	if ($keyCode -ge 0x41 -and $keyCode -le 0x5A) {
		return [char]($keyCode)
	}
	
	# For numpad keys (0x60-0x69 for 0-9, 0x6A-0x6F for operators)
	if ($keyCode -ge 0x60 -and $keyCode -le 0x69) {
		return "Num" + [char]($keyCode - 0x30)
	}
	if ($keyCode -eq 0x6A) { return "Num*" }
	if ($keyCode -eq 0x6B) { return "Num+" }
	if ($keyCode -eq 0x6C) { return "NumEnter" }
	if ($keyCode -eq 0x6D) { return "Num-" }
	if ($keyCode -eq 0x6E) { return "Num." }
	if ($keyCode -eq 0x6F) { return "Num/" }
	
	# Additional common keys that might be missing
	# Mouse buttons are sometimes reported as keys (0x01-0x06)
	if ($keyCode -eq 0x01) { return "LButton" }
	if ($keyCode -eq 0x02) { return "RButton" }
	if ($keyCode -eq 0x04) { return "MButton" }
	if ($keyCode -eq 0x05) { return "XButton1" }
	if ($keyCode -eq 0x06) { return "XButton2" }
	
	# Try to get key name using Windows Forms Keys enum as fallback
	try {
		Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
		$keysEnum = [System.Windows.Forms.Keys]
		if ([Enum]::IsDefined($keysEnum, $keyCode)) {
			$keyName = [Enum]::GetName($keysEnum, $keyCode)
			if ($keyName) {
				return $keyName
			}
		}
	} catch {
		# If we can't use the enum, continue to return null
	}
	
	# Unknown key
	return $null
}

function Start-mJig {

	#############################################################
	## mJig - An overly complex powershell mouse jiggling tool ##
	#############################################################

	<#    _                       __
		/   \                  /      \
	   '      \              /          \
	  |       |Oo          o|            |
	  `    \  |OOOo......oOO|   /        |
	   `    \\OOOOOOOOOOOOOOO\//        /
		 \ _o\OOOOOOOOOOOOOOOO//. ___ /
	 ______OOOOOOOOOOOOOOOOOOOOOOOo.___
	  --- OO'* `OOOOOOOOOO'*  `OOOOO--
		  OO.   OOOOOOOOO'    .OOOOO o
		  `OOOooOOOOOOOOOooooOOOOOO'OOOo
		.OO "OOOOOOOOOOOOOOOOOOOO"OOOOOOOo
	__ OOOOOO`OOOOOOOOOOOOOOOO"OOOOOOOOOOOOo
	___OOOOOOOO_"OOOOOOOOOOO"_OOOOOOOOOOOOOOOO
	 OOOOO^OOOO0`(____)/"OOOOOOOOOOOOO^OOOOOO
	 OOOOO OO000/00||00\000000OOOOOOOO OOOOOO
	 OOOOO O0000000000000000 ppppoooooOOOOOO
	 `OOOOO 0000000000000000 QQQQ "OOOOOOO"
	  o"OOOO 000000000000000oooooOOoooooooO'
	  OOo"OOOO.00000000000000000OOOOOOOO'
	 OOOOOO QQQQ 0000000000000000000OOOOOOO
	OOOOOO00eeee00000000000000000000OOOOOOOO.
	OOOOOOOO000000000000000000000000OOOOOOOOOO
	OOOOOOOOO00000000000000000000000OOOOOOOOOO
	`OOOOOOOOO000000000000000000000OOOOOOOOOOO
	 "OOOOOOOO0000000000000000000OOOOOOOOOOO'
	   "OOOOOOO00000000000000000OOOOOOOOOO"
	.ooooOOOOOOOo"OOOOOOO000000000000OOOOOOOOOOO"
	.OOO"""""""""".oOOOOOOOOOOOOOOOOOOOOOOOOOOOOo
	OOO         QQQQO"'                     `"QQQQ
	OOO
	`OOo.
	`"OOOOOOOOOOOOoooooooo#>

    ###################
	## Ideas & Notes ##
	###################

	# Need More inconography in output.
	# Propper hidden option
	# stealth toggle for hiding ui (to be seperate from hidden option).
	# Add a routine to determine the direction the cursor moved and add a corosponding arrow emoji to the log. https://unicode.org/emoji/charts/full-emoji-list.html
	# Add indecator for current output mode in top bar

	param(
		[Parameter(Mandatory = $false)] 
		[ValidateSet("min", "full", "hidden", "dib")]
		[string]$Output = "min",
		[Parameter(Mandatory = $false)]
		[switch]$DebugMode,
		[Parameter(Mandatory = $false)] 
		[string]$EndTime = "0",  # 0 = no end time, otherwise 4-digit 24 hour format (e.g., 1807 = 6:07 PM)
		[Parameter(Mandatory = $false)]
		[int]$EndVariance = 0,  # Variance in minutes to randomly add/subtract from EndTime to avoid overly consistent end times. Only applies if EndTime is specified (not 0).
		[Parameter(Mandatory = $false)]
		[double]$IntervalSeconds = 10,  # sets the base interval time between refreshes
		[Parameter(Mandatory = $false)]
		[double]$IntervalVariance = 2,  # Sets the maximum random plus and minus variance in seconds each refresh
		[Parameter(Mandatory = $false)]
		[double]$MoveSpeed = 0.5,  # Base movement speed in seconds (time to complete movement)
		[Parameter(Mandatory = $false)]
		[double]$MoveVariance = 0.2,  # Maximum random variance in movement speed (in seconds)
		[Parameter(Mandatory = $false)]
		[double]$TravelDistance = 10,  # Base travel distance in pixels
		[Parameter(Mandatory = $false)]
		[double]$TravelVariance = 5,  # Maximum random variance in travel distance (in pixels)
		[Parameter(Mandatory = $false)]
		[double]$AutoResumeDelaySeconds = 0  # Timer in seconds that resets on user input detection. When > 0, coordinate updates and simulated key presses are skipped.
	)

	############
	## Preparing ##
	############ 

	# Initialize script-scoped variables from parameters (so they can be modified)
	# Parameters are read-only, so we use script-scoped variables that shadow them
	$script:IntervalSeconds = $IntervalSeconds
	$script:IntervalVariance = $IntervalVariance
	$script:MoveSpeed = $MoveSpeed
	$script:MoveVariance = $MoveVariance
	$script:TravelDistance = $TravelDistance
	$script:TravelVariance = $TravelVariance
	$script:AutoResumeDelaySeconds = $AutoResumeDelaySeconds
	$script:EndVariance = $EndVariance

	# Initialize Variables
	$LastPos = $null
	$OldBufferSize = $null
	$OldWindowSize = $null
	$Rows = 0
	$OldRows = 0
	$SkipUpdate = $false
	$PreviousView = $null  # Store the view before hiding to restore it later
	$PosUpdate = $false
	$Time = $false
	$LogArray = @()
	$HostWidth = 0
	$HostHeight = 0
	$OutputLine = 0
	$LastMovementTime = $null
	$LastMovementDurationMs = 0  # Track duration of last movement in milliseconds
	$LastSimulatedKeyPress = $null  # Track when we last sent a simulated key press
	$LastAutomatedMouseMovement = $null  # Track when we last performed automated mouse movement
	$LastUserInputTime = $null  # Track when user input was last detected (for auto-resume delay timer)
	$PressedKeys = @{}  # Track currently pressed keys for display
	$PreviousIntervalKeys = @()  # Track keys pressed in previous interval for display
	$PreviousMouseWheelDelta = 0  # Track mouse wheel position for scroll detection
	$LastResizeDetection = $null  # Track when we last detected a resize
	$PendingResize = $null  # Track pending resize to throttle redraws
	$ResizeThrottleMs = 200  # Wait 200ms after window stops resizing before processing resize
	$script:lastInputCheckTime = $null  # Track when we last logged input check (for debug mode)
	$script:DialogButtonClick = $null  # Track dialog button clicks detected from main loop ("Update" or "Cancel")
	$script:DialogButtonBounds = $null  # Store dialog button bounds when dialog is open {buttonRowY, updateStartX, updateEndX, cancelStartX, cancelEndX}
	$script:LastClickLogTime = $null  # Track when we last logged a click to prevent duplicate logs
	$script:WindowTitle = "mJig - mJigg"  # Fixed window title (same for all instances to enable duplicate detection)
	
	# Function to find window handle by title (for duplicate detection)
	function Find-WindowHandleByTitle {
		param(
			[string]$WindowTitle
		)
		
		$foundHandles = @()
		
		# Use EnumWindows to find all windows with matching title
		try {
			$hasEnumWindows = [mJiggAPI.Mouse].GetMethod("EnumWindows") -ne $null
			$hasGetWindowText = [mJiggAPI.Mouse].GetMethod("GetWindowText") -ne $null
			$hasGetWindowThreadProcessId = [mJiggAPI.Mouse].GetMethod("GetWindowThreadProcessId") -ne $null
			
			if ($hasEnumWindows -and $hasGetWindowText -and $hasGetWindowThreadProcessId) {
				# We need to use a C# callback, so we'll create a temporary class for this
				# For now, use FindWindow which finds the first match
				$hasFindWindow = [mJiggAPI.Mouse].GetMethod("FindWindow") -ne $null
				if ($hasFindWindow) {
					$handle = [mJiggAPI.Mouse]::FindWindow($null, $WindowTitle)
					if ($handle -ne [IntPtr]::Zero) {
						$foundHandles += $handle
					}
				}
			}
		} catch {
			# EnumWindows/FindWindow failed
		}
		
		return $foundHandles
	}
	
	# Function to check if another instance with the same title is running
	function Test-DuplicateWindow {
		param(
			[string]$WindowTitle
		)
		
		try {
			$handles = Find-WindowHandleByTitle -WindowTitle $WindowTitle
			if ($handles.Count -gt 0) {
				# Check if any of these windows belong to a different process
				$hasGetWindowThreadProcessId = [mJiggAPI.Mouse].GetMethod("GetWindowThreadProcessId") -ne $null
				if ($hasGetWindowThreadProcessId) {
					foreach ($handle in $handles) {
						$windowProcessId = 0
						[mJiggAPI.Mouse]::GetWindowThreadProcessId($handle, [ref]$windowProcessId) | Out-Null
						if ($windowProcessId -ne 0 -and $windowProcessId -ne $PID) {
							# Found a window with same title but different process
							try {
								$otherProcess = Get-Process -Id $windowProcessId -ErrorAction SilentlyContinue
								if ($null -ne $otherProcess) {
									return @{
										IsDuplicate = $true
										ProcessId = $windowProcessId
										ProcessName = $otherProcess.ProcessName
										WindowHandle = $handle
									}
								}
							} catch {
								# Process lookup failed, but we know it's a different PID
								return @{
									IsDuplicate = $true
									ProcessId = $windowProcessId
									ProcessName = "Unknown"
									WindowHandle = $handle
								}
							}
						}
					}
				}
			}
		} catch {
			# Check failed
		}
		
		return @{ IsDuplicate = $false }
	}
	
	# Function to find window handle using EnumWindows (like Get-ProcessWindow.ps1)
	function Find-WindowHandle {
		param(
			[int]$ProcessId = $PID
		)
		
		$foundHandle = [IntPtr]::Zero
		
		# First, try to find by window title (most reliable for our case)
		try {
			$hasFindWindow = [mJiggAPI.Mouse].GetMethod("FindWindow") -ne $null
			if ($hasFindWindow -and $null -ne $script:WindowTitle) {
				# Try exact title match
				$handle = [mJiggAPI.Mouse]::FindWindow($null, $script:WindowTitle)
				if ($handle -eq [IntPtr]::Zero -and $DebugMode) {
					# Try with DEBUGMODE suffix
					$handle = [mJiggAPI.Mouse]::FindWindow($null, "$script:WindowTitle - DEBUGMODE")
				}
				if ($handle -ne [IntPtr]::Zero) {
					# Verify it belongs to our process
					$hasGetWindowThreadProcessId = [mJiggAPI.Mouse].GetMethod("GetWindowThreadProcessId") -ne $null
					if ($hasGetWindowThreadProcessId) {
						$windowProcessId = 0
						[mJiggAPI.Mouse]::GetWindowThreadProcessId($handle, [ref]$windowProcessId) | Out-Null
						if ($windowProcessId -eq $ProcessId -or $windowProcessId -eq $PID) {
							return $handle
						}
					} else {
						# Can't verify, but return it anyway
						return $handle
					}
				}
			}
		} catch {
			# FindWindow by title failed
		}
		
		# Use the C# EnumWindows-based method (fallback)
		try {
			$hasFindWindowByProcessId = [mJiggAPI.Mouse].GetMethod("FindWindowByProcessId") -ne $null
			if ($hasFindWindowByProcessId) {
				$foundHandle = [mJiggAPI.Mouse]::FindWindowByProcessId($ProcessId)
				if ($foundHandle -ne [IntPtr]::Zero) {
					return $foundHandle
				}
			}
		} catch {
			# FindWindowByProcessId failed or not available
		}
		
		# Fallback: Try parent process if current process has no window
		try {
			$currentProcess = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
			if ($null -ne $currentProcess -and $null -ne $currentProcess.Parent) {
				$parentId = $currentProcess.Parent.Id
				$hasFindWindowByProcessId = [mJiggAPI.Mouse].GetMethod("FindWindowByProcessId") -ne $null
				if ($hasFindWindowByProcessId) {
					$foundHandle = [mJiggAPI.Mouse]::FindWindowByProcessId($parentId)
					if ($foundHandle -ne [IntPtr]::Zero) {
						return $foundHandle
					}
				}
			}
		} catch {
			# Parent process lookup failed
		}
		
		return [IntPtr]::Zero
	}

	# Prep the Host Console
	# Set window title FIRST so we can be found by duplicate detection
	try {
		$Host.UI.RawUI.WindowTitle = if ($DebugMode) { "$script:WindowTitle - DEBUGMODE" } else { $script:WindowTitle }
		if ($DebugMode) {
			Write-Host "[DEBUG] Set window title: $($Host.UI.RawUI.WindowTitle)" -ForegroundColor Cyan
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to set window title: $($_.Exception.Message)" -ForegroundColor Yellow
		}
	}
	
	# Check for duplicate windows - FAIL if another instance is running
	if ($DebugMode) {
		Write-Host "[DEBUG] Checking for duplicate mJig instances..." -ForegroundColor Cyan
	}
	
	$duplicateFound = $false
	$duplicateProcessId = 0
	$duplicateProcessName = "Unknown"
	
	# Search for any windows with our window title (excluding our own PID)
	try {
		$hasFindWindowByTitlePattern = [mJiggAPI.Mouse].GetMethod("FindWindowByTitlePattern") -ne $null
		if ($hasFindWindowByTitlePattern) {
			# Search for windows with our exact title, excluding our own PID
			$duplicateHandle = [mJiggAPI.Mouse]::FindWindowByTitlePattern($script:WindowTitle, $PID)
			if ($duplicateHandle -eq [IntPtr]::Zero -and $DebugMode) {
				# Try with DEBUGMODE suffix
				$duplicateHandle = [mJiggAPI.Mouse]::FindWindowByTitlePattern("$script:WindowTitle - DEBUGMODE", $PID)
			}
			
			if ($duplicateHandle -ne [IntPtr]::Zero) {
				# Found a window with matching title - verify the process is actually running
				$hasGetWindowThreadProcessId = [mJiggAPI.Mouse].GetMethod("GetWindowThreadProcessId") -ne $null
				if ($hasGetWindowThreadProcessId) {
					$windowProcessId = 0
					[mJiggAPI.Mouse]::GetWindowThreadProcessId($duplicateHandle, [ref]$windowProcessId) | Out-Null
					if ($windowProcessId -ne 0) {
						# Verify the process actually exists and is running
						try {
							$otherProcess = Get-Process -Id $windowProcessId -ErrorAction Stop
							if ($null -ne $otherProcess) {
								# Process exists - this is a real duplicate
								$duplicateFound = $true
								$duplicateProcessId = $windowProcessId
								$duplicateProcessName = $otherProcess.ProcessName
							}
						} catch {
							# Process doesn't exist - window handle is stale, ignore it
							if ($DebugMode) {
								Write-Host "  [INFO] Found window handle for PID $windowProcessId but process doesn't exist (stale handle)" -ForegroundColor Gray
							}
						}
					}
				}
			}
		} else {
			# Fallback: Try FindWindow with exact title match
			$hasFindWindow = [mJiggAPI.Mouse].GetMethod("FindWindow") -ne $null
			if ($hasFindWindow) {
				$testHandle = [mJiggAPI.Mouse]::FindWindow($null, $script:WindowTitle)
				if ($testHandle -eq [IntPtr]::Zero -and $DebugMode) {
					$testHandle = [mJiggAPI.Mouse]::FindWindow($null, "$script:WindowTitle - DEBUGMODE")
				}
				if ($testHandle -ne [IntPtr]::Zero) {
					$hasGetWindowThreadProcessId = [mJiggAPI.Mouse].GetMethod("GetWindowThreadProcessId") -ne $null
					if ($hasGetWindowThreadProcessId) {
						$windowProcessId = 0
						[mJiggAPI.Mouse]::GetWindowThreadProcessId($testHandle, [ref]$windowProcessId) | Out-Null
						if ($windowProcessId -ne 0 -and $windowProcessId -ne $PID) {
							# Verify the process actually exists and is running
							try {
								$otherProcess = Get-Process -Id $windowProcessId -ErrorAction Stop
								if ($null -ne $otherProcess) {
									# Process exists - this is a real duplicate
									$duplicateFound = $true
									$duplicateProcessId = $windowProcessId
									$duplicateProcessName = $otherProcess.ProcessName
								}
							} catch {
								# Process doesn't exist - window handle is stale, ignore it
								if ($DebugMode) {
									Write-Host "  [INFO] Found window handle for PID $windowProcessId but process doesn't exist (stale handle)" -ForegroundColor Gray
								}
							}
						}
					}
				}
			}
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Could not check for duplicates: $($_.Exception.Message)" -ForegroundColor Yellow
		}
	}
	
	# If duplicate found, exit with error
	if ($duplicateFound) {
		Write-Host ""
		Write-Host "ERROR: Another instance of mJig is already running!" -ForegroundColor Red
		Write-Host "  Process ID: $duplicateProcessId" -ForegroundColor Red
		Write-Host "  Process Name: $duplicateProcessName" -ForegroundColor Red
		Write-Host "  Window Title: $script:WindowTitle" -ForegroundColor Red
		Write-Host ""
		Write-Host "Please close the other instance before starting a new one." -ForegroundColor Yellow
		Write-Host ""
		exit 1
	} else {
		if ($DebugMode) {
			Write-Host "  [OK] No duplicate instances found" -ForegroundColor Green
		}
	}
	
	# Clear console first before any debug output
	if ($Output -ne "hidden") {
		try {
			Clear-Host
		} catch {
			# Silently ignore - console might not be clearable
		}
	}
	
	if ($DebugMode) {
		Write-Host "Initialization Debug" -ForegroundColor Magenta
		Write-Host ""
		Write-Host "[DEBUG] Initializing console..." -ForegroundColor Cyan
		# Window title already set above, just log it
		Write-Host "[DEBUG] Window title: $($Host.UI.RawUI.WindowTitle)" -ForegroundColor Cyan
		Write-Host "[DEBUG] DebugMode is ENABLED - click detection will be logged" -ForegroundColor Yellow
		Write-Host ""
	}
	try {
		$signature = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern IntPtr GetStdHandle(int nStdHandle);
'@
		$type = Add-Type -MemberDefinition $signature -Name Win32Utils -Namespace Console -PassThru -ErrorAction SilentlyContinue
		if ($type) {
			$STD_INPUT_HANDLE = -10
			$hConsole = $type::GetStdHandle($STD_INPUT_HANDLE)
			$mode = 0
			if ($type::GetConsoleMode($hConsole, [ref]$mode)) {
				$ENABLE_QUICK_EDIT_MODE = 0x0040
				$ENABLE_MOUSE_INPUT = 0x0010
				$ENABLE_EXTENDED_FLAGS = 0x0080
				# Disable Quick Edit Mode but enable Mouse Input
				$newMode = ($mode -band (-bnot $ENABLE_QUICK_EDIT_MODE)) -bor $ENABLE_MOUSE_INPUT
				if ($type::SetConsoleMode($hConsole, $newMode)) {
					if ($DebugMode) {
						Write-Host "  [OK] Quick Edit Mode disabled, Mouse Input enabled" -ForegroundColor Green
					}
				} else {
					if ($DebugMode) {
						Write-Host "  [WARN] Failed to set console mode (SetConsoleMode failed)" -ForegroundColor Yellow
					}
				}
			} else {
				if ($DebugMode) {
					Write-Host "  [WARN] Failed to disable Quick Edit Mode (GetConsoleMode failed)" -ForegroundColor Yellow
				}
			}
		} else {
			if ($DebugMode) {
				Write-Host "  [WARN] Failed to disable Quick Edit Mode (could not load Win32 API)" -ForegroundColor Yellow
			}
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to disable Quick Edit Mode: $($_.Exception.Message)" -ForegroundColor Yellow
		}
	}
	
	try {
		[Console]::CursorVisible = $false
		if ($DebugMode) {
			Write-Host "  [OK] Console cursor hidden" -ForegroundColor Green
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to hide cursor: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	if ($Output -ne "hidden") {
		# Console already cleared above, no need to clear again
	}
	
	# Capture Initial Buffer & Window Sizes (needed even for hidden mode)
	if ($DebugMode) {
		Write-Host "[DEBUG] Capturing console dimensions..." -ForegroundColor Cyan
	}
	try {
		$pshost = Get-Host
		$pswindow = $pshost.UI.RawUI
		$newWindowSize = $pswindow.WindowSize
		$newBufferSize = $pswindow.BufferSize
		if ($DebugMode) {
			Write-Host "  [OK] Got console dimensions" -ForegroundColor Green
			Write-Host "    Window Size: $($newWindowSize.Width)x$($newWindowSize.Height)" -ForegroundColor Gray
			Write-Host "    Buffer Size: $($newBufferSize.Width)x$($newBufferSize.Height)" -ForegroundColor Gray
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to get console dimensions: $($_.Exception.Message)" -ForegroundColor Red
		}
		throw  # Re-throw as this is critical
	}
	# Set vertical buffer to match window height, but let horizontal buffer be managed by PowerShell (for text zoom)
	try {
		$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($newBufferSize.Width, $newWindowSize.Height)
		$newBufferSize = $pswindow.BufferSize
		if ($DebugMode) {
			Write-Host "  [OK] Set buffer height to match window height" -ForegroundColor Green
		}
	} catch {
		# If setting buffer size fails, continue with current buffer size
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to set buffer size: $($_.Exception.Message)" -ForegroundColor Yellow
			Write-Host "    Continuing with current buffer size" -ForegroundColor Gray
		}
		$newBufferSize = $pswindow.BufferSize
	}
	$OldBufferSize = $newBufferSize
	$OldWindowSize = $newWindowSize
	$HostWidth = $newWindowSize.Width
	$HostHeight = $newWindowSize.Height
	if ($DebugMode) {
		Write-Host "    Final host dimensions: ${HostWidth}x${HostHeight}" -ForegroundColor Gray
	}

	# Initialize the Output Array
	if ($DebugMode) {
		Write-Host "[DEBUG] Initializing output array..." -ForegroundColor Cyan
	}
	try {
		if ($Output -ne "hidden") {
			$LogArray = @()
			if ($DebugMode) {
				Write-Host "  [OK] Output mode: $Output" -ForegroundColor Green
			}
		} else {
			if ($DebugMode) {
				Write-Host "  [OK] Output mode: hidden (no log array)" -ForegroundColor Green
			}
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to initialize output array: $($_.Exception.Message)" -ForegroundColor Red
		}
		throw  # Re-throw as this is critical
	}

	###############################
	## Calculating the End Times ##
	###############################
	
	if ($DebugMode) {
		Write-Host "[DEBUG] Calculating end times..." -ForegroundColor Cyan
	}
	
	# Convert EndTime to string and parse
	# Handle: "0" = none, "00" or "0000" = midnight (0000), 2-digit = hour on the hour, 4-digit = HHmm
	try {
		$endTimeTrimmed = $EndTime.Trim()
		
		# Check if it's "0" (single digit) - means no end time
		if ($endTimeTrimmed -eq "0") {
			$endTimeInt = -1
			$endTimeStr = ""
			if ($DebugMode) {
				Write-Host "  [OK] No end time specified - script will run indefinitely" -ForegroundColor Green
			}
		} elseif ($endTimeTrimmed.Length -eq 2) {
			# 2-digit input = hour on the hour (e.g., "12" = 1200, "00" = 0000)
			$hours = [int]$endTimeTrimmed
			if ($hours -ge 0 -and $hours -le 23) {
				$endTimeInt = $hours * 100  # Convert to HHmm format (e.g., 12 -> 1200)
				$endTimeStr = $endTimeInt.ToString().PadLeft(4, '0')
				if ($DebugMode) {
					Write-Host "  [OK] Parsed end time: $endTimeStr (hour on the hour)" -ForegroundColor Green
				}
			} else {
				Write-Host "Error: Invalid hour format. Hours must be 00-23. Got: $EndTime" -ForegroundColor Red
				throw "Invalid hour format: $EndTime"
			}
		} elseif ($endTimeTrimmed.Length -eq 4) {
			# 4-digit input = HHmm format
			$endTimeInt = [int]$endTimeTrimmed
			$hours = [int]$endTimeTrimmed.Substring(0, 2)
			$minutes = [int]$endTimeTrimmed.Substring(2, 2)
			
			# Validate HHmm format
			if ($hours -ge 0 -and $hours -le 23 -and $minutes -ge 0 -and $minutes -le 59) {
				$endTimeStr = $endTimeTrimmed
				if ($DebugMode) {
					Write-Host "  [OK] Parsed end time: $endTimeStr" -ForegroundColor Green
				}
			} else {
				if ($hours -gt 23) {
					Write-Host "Error: Invalid time format. Hours must be 00-23. Got: $EndTime" -ForegroundColor Red
				} elseif ($minutes -gt 59) {
					Write-Host "Error: Invalid time format. Minutes must be 00-59. Got: $EndTime" -ForegroundColor Red
				} else {
					Write-Host "Error: Invalid time format. Expected HHmm format (0000-2359). Got: $EndTime" -ForegroundColor Red
				}
				throw "Invalid time format: $EndTime"
			}
		} else {
			Write-Host "Error: Invalid time format. Expected '0' (none), 2-digit hour (00-23), or 4-digit HHmm (0000-2359). Got: $EndTime" -ForegroundColor Red
			throw "Invalid time format: $EndTime"
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to parse endTime: $($_.Exception.Message)" -ForegroundColor Red
		}
		if ($_.Exception.Message -notmatch "Invalid time format") {
			Write-Host "Error: Invalid EndTime format: $EndTime" -ForegroundColor Red
		}
		throw
	}
	
	# Time format has already been validated in the try-catch block above
	# Proceed with initialization
		if ($DebugMode) {
			Write-Host "[DEBUG] Loading System.Windows.Forms assembly..." -ForegroundColor Cyan
		}
		try {
			Add-Type -AssemblyName System.Windows.Forms
			if ($DebugMode) {
				Write-Host "  [OK] System.Windows.Forms loaded" -ForegroundColor Green
			}
		} catch {
			if ($DebugMode) {
				Write-Host "  [FAIL] Failed to load System.Windows.Forms: $($_.Exception.Message)" -ForegroundColor Red
			}
			throw  # Re-throw as this is critical
		}
		
		# Add Windows API for system-wide keyboard detection and key sending
		if ($DebugMode) {
			Write-Host "[DEBUG] Loading Windows API types..." -ForegroundColor Cyan
		}
		# Check if types already exist and have the required methods
		$typesNeedReload = $false
		try {
			# Use a safer method to check if types exist without throwing errors
			$existingKeyboard = $null
			$existingMouse = $null
			$existingMouseHook = $null
			
			# Try to get the types using Get-Type or by checking if they're loaded
			$allTypes = [System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.Namespace -eq 'mJiggAPI' }
			
			foreach ($type in $allTypes) {
				if ($type.Name -eq 'Keyboard') { $existingKeyboard = $type }
				if ($type.Name -eq 'Mouse') { $existingMouse = $type }
				if ($type.Name -eq 'MouseHook') { $existingMouseHook = $type }
			}
			
			if ($null -ne $existingMouse) {
				$hasGetCursorPos = $existingMouse.GetMethod("GetCursorPos") -ne $null
				$hasGetForegroundWindow = $existingMouse.GetMethod("GetForegroundWindow") -ne $null
				$hasFindWindow = $existingMouse.GetMethod("FindWindow") -ne $null
				$hasFindWindowByProcessId = $existingMouse.GetMethod("FindWindowByProcessId") -ne $null
				$hasFindWindowByTitlePattern = $existingMouse.GetMethod("FindWindowByTitlePattern") -ne $null
				if (-not $hasGetCursorPos -or -not $hasGetForegroundWindow -or -not $hasFindWindow -or -not $hasFindWindowByProcessId -or -not $hasFindWindowByTitlePattern) {
					# Type exists but missing required methods - need to reload
					# Note: PowerShell cannot remove types once loaded, so Add-Type will fail silently
					# User may need to restart PowerShell session to get updated types
					$typesNeedReload = $true
					if ($DebugMode) {
						Write-Host "  [WARN] Existing types found but missing required methods" -ForegroundColor Yellow
						Write-Host "  [WARN] Missing: GetCursorPos=$(-not $hasGetCursorPos), GetForegroundWindow=$(-not $hasGetForegroundWindow), FindWindow=$(-not $hasFindWindow), FindWindowByProcessId=$(-not $hasFindWindowByProcessId), FindWindowByTitlePattern=$(-not $hasFindWindowByTitlePattern)" -ForegroundColor Yellow
						Write-Host "  [WARN] Attempting reload (may fail if types already exist - restart PowerShell if needed)" -ForegroundColor Yellow
					}
				} else {
					# Types exist and have required methods - skip reload
					if ($DebugMode) {
						Write-Host "  [INFO] Types already loaded from previous run (with required methods)" -ForegroundColor Gray
					}
				}
			} else {
				# Types don't exist - need to load
				$typesNeedReload = $true
			}
		} catch {
			# Types don't exist or can't be accessed - need to load
			$typesNeedReload = $true
			if ($DebugMode) {
				Write-Host "  [INFO] Types not found, will load them: $($_.Exception.Message)" -ForegroundColor Gray
			}
		}
		
		# Only attempt to add types if they don't exist or are incomplete
		if ($typesNeedReload) {
			try {
				# Try to add the types - use ErrorAction Stop to catch failures
				$typeDefinition = @"
using System;
using System.Runtime.InteropServices;
namespace mJiggAPI {
	// Define POINT struct for P/Invoke (avoids dependency on System.Drawing.Primitives)
	[StructLayout(LayoutKind.Sequential)]
	public struct POINT {
		public int X;
		public int Y;
		
		public POINT(int x, int y) {
			X = x;
			Y = y;
		}
	}
	
	[StructLayout(LayoutKind.Sequential)]
	public struct RECT {
		public int Left;
		public int Top;
		public int Right;
		public int Bottom;
	}
	
	[StructLayout(LayoutKind.Sequential)]
	public struct CONSOLE_SCREEN_BUFFER_INFO {
		public COORD dwSize;
		public COORD dwCursorPosition;
		public short wAttributes;
		public SMALL_RECT srWindow;
		public COORD dwMaximumWindowSize;
	}
	
	[StructLayout(LayoutKind.Sequential)]
	public struct MOUSE_EVENT_RECORD {
		public COORD dwMousePosition;
		public uint dwButtonState;
		public uint dwControlKeyState;
		public uint dwEventFlags;
	}
	
	[StructLayout(LayoutKind.Explicit)]
	public struct INPUT_RECORD {
		[FieldOffset(0)]
		public ushort EventType;
		[FieldOffset(4)]
		public MOUSE_EVENT_RECORD MouseEvent;
	}
	
	[StructLayout(LayoutKind.Sequential)]
	public struct COORD {
		public short X;
		public short Y;
	}
	
	[StructLayout(LayoutKind.Sequential)]
	public struct SMALL_RECT {
		public short Left;
		public short Top;
		public short Right;
		public short Bottom;
	}
	
	public class Keyboard {
		[DllImport("user32.dll")]
		public static extern short GetAsyncKeyState(int vKey);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
		public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
		
		public const uint KEYEVENTF_KEYUP = 0x0002;
		public const int VK_RMENU = 0xA5;  // Right Alt key (modifier, won't type anything)
	}
	
	public class Mouse {
		[DllImport("user32.dll")]
		public static extern short GetAsyncKeyState(int vKey);
		
		[DllImport("user32.dll")]
		public static extern int GetSystemMetrics(int nIndex);
		
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetCursorPos(out POINT lpPoint);
		
		[DllImport("kernel32.dll")]
		public static extern IntPtr GetConsoleWindow();
		
		[DllImport("user32.dll")]
		public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);
		
		[DllImport("user32.dll")]
		public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
		
		[DllImport("kernel32.dll")]
		public static extern IntPtr GetStdHandle(int nStdHandle);
		
		[DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();
		
		[DllImport("kernel32.dll")]
		public static extern bool GetConsoleScreenBufferInfo(IntPtr hConsoleOutput, out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);
		
		[DllImport("kernel32.dll", SetLastError = true)]
		public static extern bool ReadConsoleInput(IntPtr hConsoleInput, [Out] INPUT_RECORD[] lpBuffer, uint nLength, out uint lpNumberOfEventsRead);
		
		[DllImport("kernel32.dll", SetLastError = true)]
		public static extern uint PeekConsoleInput(IntPtr hConsoleInput, [Out] INPUT_RECORD[] lpBuffer, uint nLength, out uint lpNumberOfEventsRead);
		
		// Window finding APIs
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
		
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
		
		public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
		
		[DllImport("user32.dll", SetLastError = true)]
		public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
		
		// Static storage for EnumWindows callback
		private static IntPtr foundWindowHandle = IntPtr.Zero;
		private static int targetProcessId = 0;
		
		// Callback for EnumWindows to find window by process ID
		private static bool EnumWindowsCallback(IntPtr hWnd, IntPtr lParam) {
			if (hWnd == IntPtr.Zero) return true;
			
			try {
				uint windowProcessId = 0;
				GetWindowThreadProcessId(hWnd, out windowProcessId);
				if (windowProcessId == targetProcessId) {
					foundWindowHandle = hWnd;
					return false; // Stop enumeration
				}
			} catch { }
			return true; // Continue enumeration
		}
		
		// Public method to find window handle by process ID
		public static IntPtr FindWindowByProcessId(int processId) {
			foundWindowHandle = IntPtr.Zero;
			targetProcessId = processId;
			try {
				EnumWindows(new EnumWindowsProc(EnumWindowsCallback), IntPtr.Zero);
			} catch { }
			return foundWindowHandle;
		}
		
		// Static storage for title-based search
		private static IntPtr foundWindowHandleByTitle = IntPtr.Zero;
		private static string targetTitlePattern = string.Empty;
		private static int excludeProcessId = 0;
		
		// Callback for EnumWindows to find window by title pattern
		private static bool EnumWindowsCallbackByTitle(IntPtr hWnd, IntPtr lParam) {
			if (hWnd == IntPtr.Zero) return true;
			
			try {
				uint windowProcessId = 0;
				GetWindowThreadProcessId(hWnd, out windowProcessId);
				
				// Skip if this is the process we want to exclude
				if (excludeProcessId != 0 && windowProcessId == excludeProcessId) {
					return true;
				}
				
				// Get window title
				System.Text.StringBuilder sb = new System.Text.StringBuilder(256);
				int length = GetWindowText(hWnd, sb, sb.Capacity);
				string windowTitle = sb.ToString();
				
				// Check if title matches pattern (starts with pattern)
				if (!string.IsNullOrEmpty(windowTitle) && windowTitle.StartsWith(targetTitlePattern, System.StringComparison.OrdinalIgnoreCase)) {
					foundWindowHandleByTitle = hWnd;
					return false; // Stop enumeration
				}
			} catch { }
			return true; // Continue enumeration
		}
		
		// Public method to find window handle by title pattern (excluding a specific process ID)
		public static IntPtr FindWindowByTitlePattern(string titlePattern, int excludePid) {
			foundWindowHandleByTitle = IntPtr.Zero;
			targetTitlePattern = titlePattern ?? string.Empty;
			excludeProcessId = excludePid;
			try {
				EnumWindows(new EnumWindowsProc(EnumWindowsCallbackByTitle), IntPtr.Zero);
			} catch { }
			return foundWindowHandleByTitle;
		}
		
		// Mouse button virtual key codes
		public const int VK_LBUTTON = 0x01;
		public const int VK_RBUTTON = 0x02;
		public const int VK_MBUTTON = 0x04;
		public const int VK_XBUTTON1 = 0x05;
		public const int VK_XBUTTON2 = 0x06;
		
		// Console input event constants
		public const ushort MOUSE_EVENT = 2;
		public const uint MOUSE_LEFT_BUTTON_DOWN = 0x0001;
		public const uint MOUSE_LEFT_BUTTON_UP = 0x0002;
		public const uint DOUBLE_CLICK = 0x0002;
		
		// For mouse wheel detection, we use GetAsyncKeyState with VK codes
		// Note: Mouse wheel doesn't have VK codes, so we need to use a hook
		// For now, we'll check for wheel button state changes
	}
	
	public class MouseHook {
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool UnhookWindowsHookEx(IntPtr hhk);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
		
		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr GetModuleHandle(string lpModuleName);
		
		[DllImport("user32.dll")]
		public static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);
		
		[StructLayout(LayoutKind.Sequential)]
		public struct MSG {
			public IntPtr hwnd;
			public uint message;
			public IntPtr wParam;
			public IntPtr lParam;
			public uint time;
			public POINT pt;
		}
		
		[StructLayout(LayoutKind.Sequential)]
		public struct MSLLHOOKSTRUCT {
			public POINT pt;
			public uint mouseData;
			public uint flags;
			public uint time;
			public IntPtr dwExtraInfo;
		}
		
		public const uint PM_REMOVE = 0x0001;
		public const uint PM_NOREMOVE = 0x0000;
		
		public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
		
		public const int WH_MOUSE_LL = 14;
		public const int WM_MOUSEWHEEL = 0x020A;
		public const int WM_MOUSEHWHEEL = 0x020E;
		
		public static IntPtr hHook = IntPtr.Zero;
		public static LowLevelMouseProc proc = HookCallback;
		public static int lastWheelDelta = 0;
		
		public static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
			if (nCode >= 0) {
				int msg = wParam.ToInt32();
				if (msg == WM_MOUSEWHEEL || msg == WM_MOUSEHWHEEL) {
					try {
						// Mouse wheel event detected
						// Use Marshal.PtrToStructure to properly read MSLLHOOKSTRUCT
						// This is more reliable than manual offset reading
						MSLLHOOKSTRUCT mouseHookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
						// Delta is in the high 16 bits of mouseData, sign-extended
						int delta = (short)((mouseHookStruct.mouseData >> 16) & 0xFFFF);
						lastWheelDelta += delta;  // Accumulate delta
					} catch {
						// If structure reading fails, try manual offset as fallback
						try {
							int mouseData = Marshal.ReadInt32(lParam, 8);
							int delta = (short)((mouseData >> 16) & 0xFFFF);
							lastWheelDelta += delta;
						} catch {
							// Both methods failed, skip this event
						}
					}
				}
			}
			return CallNextHookEx(hHook, nCode, wParam, lParam);
		}
		
		public static void InstallHook() {
			if (hHook == IntPtr.Zero) {
				// For low-level hooks (WH_MOUSE_LL), we can use GetModuleHandle(null) 
				// which gets the handle for the current process. This works fine for
				// low-level hooks and avoids dependencies on System.Diagnostics.Process
				IntPtr hMod = GetModuleHandle(null);
				hHook = SetWindowsHookEx(WH_MOUSE_LL, proc, hMod, 0);
			}
		}
		
		public static void ProcessMessages() {
			// Process Windows messages to ensure hook callback is invoked
			// For low-level hooks, this isn't strictly necessary, but it helps
			MSG msg;
			int count = 0;
			while (PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_REMOVE) && count < 10) {
				// Process a few messages - low-level hooks are called directly by Windows
				count++;
			}
		}
		
		public static void UninstallHook() {
			if (hHook != IntPtr.Zero) {
				UnhookWindowsHookEx(hHook);
				hHook = IntPtr.Zero;
			}
		}
	}
}
"@
				
				# Add-Type with explicit error handling and assembly references
				# Note: We use our own POINT struct, so we don't need System.Drawing.dll
				# We also avoid System.Diagnostics.Process by using GetModuleHandle(null)
				$addTypeResult = $null
				$addTypeError = $null
				try {
					if ($DebugMode) {
						Write-Host "  [DEBUG] Attempting to add types..." -ForegroundColor Cyan
					}
					$addTypeResult = Add-Type -TypeDefinition $typeDefinition -ReferencedAssemblies @("System.dll") -ErrorAction Stop
					if ($DebugMode) {
						Write-Host "  [OK] Add-Type completed successfully" -ForegroundColor Green
					}
				} catch {
					$addTypeError = $_
					# If Add-Type fails, it might be because types already exist
					# Check if the error is about duplicate types
					if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*duplicate*" -or $_.Exception.Message -like "*Cannot add type*") {
						if ($DebugMode) {
							Write-Host "  [INFO] Types may already exist: $($_.Exception.Message)" -ForegroundColor Yellow
						}
					} else {
						# Some other error occurred - log it
						if ($DebugMode) {
							Write-Host "  [WARN] Add-Type error: $($_.Exception.Message)" -ForegroundColor Yellow
							if ($_.Exception.InnerException) {
								Write-Host "  [WARN] Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
							}
						}
						# Don't throw yet - we'll check if types exist anyway
					}
				}
				
				# Always verify types were loaded, regardless of Add-Type result
				# Try both reflection and direct type access
				$loadedKeyboard = $null
				$loadedMouse = $null
				$loadedMouseHook = $null
				
				# First try direct type access (most reliable)
				try {
					$testType = [mJiggAPI.Keyboard]
					$loadedKeyboard = $testType
				} catch {
					# Type not accessible directly, try reflection
				}
				
				try {
					$testType = [mJiggAPI.Mouse]
					$loadedMouse = $testType
				} catch {
					# Type not accessible directly, try reflection
				}
				
				try {
					$testType = [mJiggAPI.MouseHook]
					$loadedMouseHook = $testType
				} catch {
					# Type not accessible directly, try reflection
				}
				
				# If direct access failed, try reflection
				if ($null -eq $loadedKeyboard -or $null -eq $loadedMouse -or $null -eq $loadedMouseHook) {
					try {
						$allTypes = [System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.Namespace -eq 'mJiggAPI' }
						foreach ($type in $allTypes) {
							if ($type.Name -eq 'Keyboard' -and $null -eq $loadedKeyboard) { $loadedKeyboard = $type }
							if ($type.Name -eq 'Mouse' -and $null -eq $loadedMouse) { $loadedMouse = $type }
							if ($type.Name -eq 'MouseHook' -and $null -eq $loadedMouseHook) { $loadedMouseHook = $type }
						}
					} catch {
						if ($DebugMode) {
							Write-Host "  [WARN] Error checking for loaded types: $($_.Exception.Message)" -ForegroundColor Yellow
						}
					}
				}
				
				# Check if we have all three types
				if ($null -ne $loadedKeyboard -and $null -ne $loadedMouse -and $null -ne $loadedMouseHook) {
					if ($DebugMode) {
						Write-Host "  [OK] All types verified: Keyboard, Mouse, MouseHook" -ForegroundColor Green
					}
				} else {
					# Types weren't loaded - check if they already exist from previous check
					if ($null -ne $existingKeyboard -and $null -ne $existingMouse -and $null -ne $existingMouseHook) {
						if ($DebugMode) {
							Write-Host "  [INFO] Types already exist from previous run" -ForegroundColor Gray
						}
					} else {
						# Types don't exist and failed to load - try to find them anywhere
						if ($DebugMode) {
							Write-Host "  [DEBUG] Searching all assemblies for mJiggAPI types..." -ForegroundColor Cyan
							try {
								$allAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()
								foreach ($assembly in $allAssemblies) {
									try {
										$types = $assembly.GetTypes() | Where-Object { $_.Name -in @('Keyboard', 'Mouse', 'MouseHook') }
										if ($types) {
											Write-Host "    Found types in $($assembly.FullName): $($types | ForEach-Object { $_.FullName } | Join-String -Separator ', ')" -ForegroundColor Gray
										}
									} catch {
										# Some assemblies can't be inspected
									}
								}
							} catch {
								Write-Host "    Error searching assemblies: $($_.Exception.Message)" -ForegroundColor Yellow
							}
						}
						
						# Types don't exist and failed to load
						$missingTypes = @()
						if ($null -eq $loadedKeyboard) { $missingTypes += "Keyboard" }
						if ($null -eq $loadedMouse) { $missingTypes += "Mouse" }
						if ($null -eq $loadedMouseHook) { $missingTypes += "MouseHook" }
						$errorMsg = "Failed to load required mJiggAPI types: $($missingTypes -join ', ')"
						if ($addTypeError) {
							$errorMsg += "`nAdd-Type error: $($addTypeError.Exception.Message)"
						}
						if ($DebugMode) {
							Write-Host "  [FAIL] $errorMsg" -ForegroundColor Red
						}
						throw $errorMsg
					}
				}
			} catch {
				# Final fallback - check if types exist anyway
				$finalKeyboard = $null
				$finalMouse = $null
				$finalMouseHook = $null
				try {
					$allTypes = [System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.Namespace -eq 'mJiggAPI' }
					foreach ($type in $allTypes) {
						if ($type.Name -eq 'Keyboard') { $finalKeyboard = $type }
						if ($type.Name -eq 'Mouse') { $finalMouse = $type }
						if ($type.Name -eq 'MouseHook') { $finalMouseHook = $type }
					}
				} catch {
					# Ignore errors when checking for existing types
				}
				
				if ($null -ne $finalKeyboard -and $null -ne $finalMouse -and $null -ne $finalMouseHook) {
					if ($DebugMode) {
						Write-Host "  [INFO] Types found after error recovery" -ForegroundColor Gray
					}
				} else {
					if ($DebugMode) {
						Write-Host "  [FAIL] Add-Type failed and types don't exist: $($_.Exception.Message)" -ForegroundColor Red
						Write-Host "  [INFO] This may require restarting PowerShell to reload types" -ForegroundColor Yellow
					}
					throw "Failed to load required mJiggAPI types: $($_.Exception.Message)"
				}
			}
		}
		
		# Verify types loaded correctly
		try {
			$testKey = [mJiggAPI.Keyboard]::GetAsyncKeyState(0x01)
			$testPoint = New-Object mJiggAPI.POINT
			$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
			if ($hasGetCursorPos) {
				$testMouse = [mJiggAPI.Mouse]::GetCursorPos([ref]$testPoint)
			}
			if ($DebugMode) {
				Write-Host "  [OK] Windows API types loaded successfully" -ForegroundColor Green
			}
		} catch {
			if ($DebugMode) {
				Write-Host "  [FAIL] Could not verify keyboard/mouse API: $($_.Exception.Message)" -ForegroundColor Red
			}
			Write-Host "Warning: Could not verify keyboard/mouse API. Some features may be disabled." -ForegroundColor Yellow
		}
		
		# Apply variance to end time if variance is set and end time is specified (not -1)
		if ($endTimeInt -ne -1 -and $script:EndVariance -gt 0) {
			try {
				$ras = Get-Random -Maximum 3 -Minimum 1
				if ($ras -eq 1) {
					$variance = -(Get-Random -Maximum $script:EndVariance)
					$endTimeInt = $endTimeInt + $variance
				} else {
					$variance = (Get-Random -Maximum $script:EndVariance)
					$endTimeInt = $endTimeInt + $variance
				}
				# Ensure time stays within valid range (0-2359)
				if ($endTimeInt -lt 0) {
					$endTimeInt = 0
				} elseif ($endTimeInt -gt 2359) {
					$endTimeInt = 2359
				}
				$endTimeStr = $endTimeInt.ToString().PadLeft(4, '0')
				if ($DebugMode) {
					Write-Host "  [OK] Applied variance: $variance minutes, final end time: $endTimeStr" -ForegroundColor Green
				}
			} catch {
				if ($DebugMode) {
					Write-Host "  [FAIL] Failed to apply variance: $($_.Exception.Message)" -ForegroundColor Red
				}
			}
		}
		
		# Calculate end date/time only if end time is set (not -1)
		if ($endTimeInt -ne -1) {
			try {
				$currentTime = Get-Date -Format "HHmm"
				if ($DebugMode) {
					Write-Host "  [OK] Current time: $currentTime" -ForegroundColor Green
				}
			} catch {
				if ($DebugMode) {
					Write-Host "  [FAIL] Failed to get current time: $($_.Exception.Message)" -ForegroundColor Red
				}
				throw
			}
			try {
				if ($endTimeInt -le [int]$currentTime) {
					$tommorow = (Get-Date).AddDays(1)
					$endDate = Get-Date $tommorow -Format "MMdd"
					if ($DebugMode) {
						Write-Host "  [OK] End time is today, using tomorrow's date: $endDate" -ForegroundColor Green
					}
				} else {
					$endDate = Get-Date -Format "MMdd"
					if ($DebugMode) {
						Write-Host "  [OK] End time is today, using today's date: $endDate" -ForegroundColor Green
					}
				}
				$end = "$endDate$endTimeStr"
				$time = $false
				if ($DebugMode) {
					Write-Host "  [OK] Final end datetime: $end" -ForegroundColor Green
				}
			} catch {
				if ($DebugMode) {
					Write-Host "  [FAIL] Failed to calculate end datetime: $($_.Exception.Message)" -ForegroundColor Red
				}
				throw
			}
		} else {
			# No end time - set end to empty and time to false
			$end = ""
			$time = $false
			if ($DebugMode) {
				Write-Host "  [OK] No end time - script will run indefinitely" -ForegroundColor Green
			}
		}

		# Initialize mouse wheel hook for scroll detection
		if ($DebugMode) {
			Write-Host "[DEBUG] Initializing mouse wheel hook..." -ForegroundColor Cyan
		}
		try {
			[mJiggAPI.MouseHook]::InstallHook()
			# Verify hook was installed
			if ([mJiggAPI.MouseHook]::hHook -eq [IntPtr]::Zero) {
				# Hook installation failed - low-level hooks may not work in PowerShell
				# This is expected as WH_MOUSE_LL requires the hook procedure to be in a DLL
				if ($DebugMode) {
					Write-Host "  [WARN] Mouse wheel hook: Not available (requires DLL)" -ForegroundColor Yellow
				}
			} else {
				if ($DebugMode) {
					Write-Host "  [OK] Mouse wheel hook: Installed successfully" -ForegroundColor Green
				}
			}
		} catch {
			# Hook initialization failed, wheel detection will be disabled
			# Silently continue - wheel detection is optional
			if ($DebugMode) {
				Write-Host "  [FAIL] Mouse wheel hook: Failed to install - $($_.Exception.Message)" -ForegroundColor Red
			}
		}
		
		# Initialize lastPos for mouse detection
		if ($DebugMode) {
			Write-Host "[DEBUG] Initializing mouse position tracking..." -ForegroundColor Cyan
		}
		try {
			if ($null -eq $LastPos) {
				# Use direct Windows API call for better performance (avoids .NET stutter)
				$point = New-Object mJiggAPI.POINT
				$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
				if ($hasGetCursorPos) {
					if ([mJiggAPI.Mouse]::GetCursorPos([ref]$point)) {
						# Convert POINT to System.Drawing.Point for compatibility with rest of code
						$LastPos = New-Object System.Drawing.Point($point.X, $point.Y)
					} else {
						throw "GetCursorPos API call failed"
					}
				} else {
					throw "GetCursorPos method not available"
				}
				if ($DebugMode) {
					Write-Host "  [OK] Initial mouse position: $($LastPos.X), $($LastPos.Y)" -ForegroundColor Green
				}
			} else {
				if ($DebugMode) {
					Write-Host "  [OK] Mouse position already set: $($LastPos.X), $($LastPos.Y)" -ForegroundColor Green
				}
			}
		} catch {
			if ($DebugMode) {
				Write-Host "  [FAIL] Failed to get mouse position: $($_.Exception.Message)" -ForegroundColor Red
			}
			# Don't throw - mouse position tracking is optional
		}

		# Track start time for runtime calculation
		$ScriptStartTime = Get-Date

		# Function to calculate smooth movement path with acceleration/deceleration
		# Returns an array of points and the total movement time in milliseconds
		function Get-SmoothMovementPath {
			param(
				[int]$startX,
				[int]$startY,
				[int]$endX,
				[int]$endY,
				[double]$baseSpeedSeconds,
				[double]$varianceSeconds
			)
			
			# Calculate distance
			$deltaX = $endX - $startX
			$deltaY = $endY - $startY
			$distance = [Math]::Sqrt($deltaX * $deltaX + $deltaY * $deltaY)
			
			# If distance is very small, return single point
			if ($distance -lt 1) {
				return @{
					Points = @([PSCustomObject]@{ X = $endX; Y = $endY })
					TotalTimeMs = 0
				}
			}
			
			# Calculate movement time with variance (in milliseconds)
			$baseSpeedMs = $baseSpeedSeconds * 1000
			$varianceMs = $varianceSeconds * 1000
			$varianceAmountMs = Get-Random -Minimum 0 -Maximum ($varianceMs + 1)
			$ras = Get-Random -Maximum 2 -Minimum 0
			if ($ras -eq 0) {
				$movementTimeMs = ($baseSpeedMs - $varianceAmountMs)
			} else {
				$movementTimeMs = ($baseSpeedMs + $varianceAmountMs)
			}
			
			# Ensure minimum movement time of 50ms
			if ($movementTimeMs -lt 50) {
				$movementTimeMs = 50
			}
			
			# Calculate number of points based on distance and time
			# More points for longer distances and longer times
			# Aim for roughly 1 point per 5-10 pixels, but adjust based on time
			$basePoints = [Math]::Max(2, [Math]::Floor($distance / 8))
			$timeBasedPoints = [Math]::Max(2, [Math]::Floor($movementTimeMs / 20))  # ~1 point per 20ms
			$numPoints = [Math]::Max(2, [Math]::Min($basePoints, $timeBasedPoints))
			
			# Randomly determine if we should apply a curve (chance of 0 = straight line)
			# Curve amount: 0 = straight line, > 0 = curved path
			# Maximum curve is 30% of the distance, with 30% chance of no curve
			$curveChance = Get-Random -Minimum 0 -Maximum 100
			$curveAmount = 0
			if ($curveChance -ge 30) {
				# Apply a curve - random amount between 5% and 30% of distance
				$curvePercent = Get-Random -Minimum 5 -Maximum 31  # 5-30%
				$curveAmount = ($distance * $curvePercent) / 100
			}
			
			# Calculate perpendicular direction for curve offset
			# Perpendicular vector: (-deltaY, deltaX) normalized
			$perpendicularX = 0
			$perpendicularY = 0
			if ($curveAmount -gt 0) {
				$normalizedLength = [Math]::Sqrt($deltaX * $deltaX + $deltaY * $deltaY)
				if ($normalizedLength -gt 0) {
					$perpendicularX = -$deltaY / $normalizedLength
					$perpendicularY = $deltaX / $normalizedLength
				}
				# Randomly choose curve direction (left or right)
				$curveDirection = Get-Random -Maximum 2  # 0 or 1
				if ($curveDirection -eq 0) {
					$perpendicularX = -$perpendicularX
					$perpendicularY = -$perpendicularY
				}
			}
			
			# Generate points with acceleration/deceleration curve and optional path curve
			# Use ease-in-out-cubic: accelerates in first half, decelerates in second half
			$points = @()
			for ($i = 0; $i -le $numPoints; $i++) {
				# Normalized progress (0 to 1)
				$t = $i / $numPoints
				
				# Ease-in-out-cubic function: accelerates then decelerates
				# f(t) = t < 0.5 ? 4t³ : 1 - pow(-2t + 2, 3)/2
				if ($t -lt 0.5) {
					$easedT = 4 * $t * $t * $t
				} else {
					$easedT = 1 - [Math]::Pow(-2 * $t + 2, 3) / 2
				}
				
				# Calculate base position along straight path
				$baseX = $startX + $deltaX * $easedT
				$baseY = $startY + $deltaY * $easedT
				
				# Apply curve offset if curveAmount > 0
				# Use quadratic curve: offset is maximum at midpoint (t=0.5) and 0 at start/end
				# This creates a smooth arc: offset = curveAmount * 4 * t * (1 - t)
				if ($curveAmount -gt 0) {
					$curveOffset = $curveAmount * 4 * $t * (1 - $t)  # Quadratic curve: peaks at t=0.5
					$baseX = $baseX + $perpendicularX * $curveOffset
					$baseY = $baseY + $perpendicularY * $curveOffset
				}
				
				# Round to integer pixel coordinates
				$x = [Math]::Round($baseX)
				$y = [Math]::Round($baseY)
				
				$points += [PSCustomObject]@{
					X = $x
					Y = $y
				}
			}
			
			return @{
				Points = $points
				TotalTimeMs = [Math]::Round($movementTimeMs)
			}
		}

		# Function to get direction arrow emoji based on movement delta
		# Options: "arrows" (emoji arrows), "text" (N/S/E/W/NE/etc), "simple" (←→↑↓↗↖↘↙)
		function Get-DirectionArrow {
			param(
				[int]$deltaX,
				[int]$deltaY,
				[string]$style = "simple"  # "arrows", "text", or "simple"
			)
			
			# Calculate angle and determine primary direction
			# Use a threshold to determine if movement is primarily horizontal, vertical, or diagonal
			$absX = [Math]::Abs($deltaX)
			$absY = [Math]::Abs($deltaY)
			
			# If movement is very small, return no arrow
			if ($absX -lt 5 -and $absY -lt 5) {
				return ""
			}
			
			# Determine if movement is primarily horizontal or vertical
			# If one axis is significantly larger, use cardinal direction
			# Otherwise use diagonal direction
			if ($absX -gt $absY * 2) {
				# Primarily horizontal
				if ($style -eq "text") {
					if ($deltaX -gt 0) { return "E" } else { return "W" }
				} elseif ($style -eq "arrows") {
					if ($deltaX -gt 0) { return "➡️" } else { return "⬅️" }
				} else {
					# simple style
					if ($deltaX -gt 0) { return "→" } else { return "←" }
				}
			} elseif ($absY -gt $absX * 2) {
				# Primarily vertical
				if ($style -eq "text") {
					if ($deltaY -gt 0) { return "S" } else { return "N" }
				} elseif ($style -eq "arrows") {
					if ($deltaY -gt 0) { return "⬇️" } else { return "⬆️" }
				} else {
					# simple style
					if ($deltaY -gt 0) { return "↓" } else { return "↑" }
				}
			} else {
				# Diagonal movement
				if ($style -eq "text") {
					if ($deltaX -gt 0 -and $deltaY -gt 0) {
						return "SE"
					} elseif ($deltaX -gt 0 -and $deltaY -lt 0) {
						return "NE"
					} elseif ($deltaX -lt 0 -and $deltaY -gt 0) {
						return "SW"
					} else {
						return "NW"
					}
				} elseif ($style -eq "arrows") {
					if ($deltaX -gt 0 -and $deltaY -gt 0) {
						return "↘️"
					} elseif ($deltaX -gt 0 -and $deltaY -lt 0) {
						return "↗️"
					} elseif ($deltaX -lt 0 -and $deltaY -gt 0) {
						return "↙️"
					} else {
						return "↖️"
					}
				} else {
					# simple style
					if ($deltaX -gt 0 -and $deltaY -gt 0) {
						return "↘"
					} elseif ($deltaX -gt 0 -and $deltaY -lt 0) {
						return "↗"
					} elseif ($deltaX -lt 0 -and $deltaY -gt 0) {
						return "↙"
					} else {
						return "↖"
					}
				}
			}
		}

		# Function to draw drop shadow for dialog boxes
		function Draw-DialogShadow {
			param(
				[int]$dialogX,
				[int]$dialogY,
				[int]$dialogWidth,
				[int]$dialogHeight
			)
			
			$shadowChar = [char]0x2591  # ░ light shade character
			
			# Draw shadow on right side (one column, starting from row 1 to avoid top corner)
			# The shadow should be one column to the right of the dialog
			for ($i = 1; $i -le $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX + $dialogWidth, $dialogY + $i)
				Write-Host $shadowChar -NoNewline -ForegroundColor DarkGray
			}
			
			# Draw shadow on bottom (one row below the dialog, starting from column 1 to avoid left corner)
			# The shadow should be one row below the dialog bottom border
			# Use dialogHeight + 1 to ensure it's definitely below the dialog
			for ($i = 1; $i -le $dialogWidth; $i++) {
				[Console]::SetCursorPosition($dialogX + $i, $dialogY + $dialogHeight + 1)
				Write-Host $shadowChar -NoNewline -ForegroundColor DarkGray
			}
			
			# Draw corner shadow (bottom-right corner, one row down and one column right)
			[Console]::SetCursorPosition($dialogX + $dialogWidth, $dialogY + $dialogHeight + 1)
			Write-Host $shadowChar -NoNewline -ForegroundColor DarkGray
		}
		
		# Function to clear drop shadow for dialog boxes
		function Clear-DialogShadow {
			param(
				[int]$dialogX,
				[int]$dialogY,
				[int]$dialogWidth,
				[int]$dialogHeight
			)
			
			# Clear shadow on right side (one column)
			for ($i = 1; $i -le $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX + $dialogWidth, $dialogY + $i)
				Write-Host " " -NoNewline
			}
			
			# Clear shadow on bottom (one row below the dialog)
			for ($i = 1; $i -le $dialogWidth; $i++) {
				[Console]::SetCursorPosition($dialogX + $i, $dialogY + $dialogHeight + 1)
				Write-Host " " -NoNewline
			}
			
			# Clear corner shadow
			[Console]::SetCursorPosition($dialogX + $dialogWidth, $dialogY + $dialogHeight + 1)
			Write-Host " " -NoNewline
		}

		# Function to show popup dialog for changing end time
		function Show-TimeChangeDialog {
			param(
				[int]$currentEndTime,
				[ref]$HostWidthRef,
				[ref]$HostHeightRef
			)
			
			# Get current host dimensions from references
			$currentHostWidth = $HostWidthRef.Value
			$currentHostHeight = $HostHeightRef.Value
			
			# Dialog dimensions (reduced width)
			$dialogWidth = 35
			$dialogHeight = 7
			$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
			$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
			
			# Save current cursor position and visibility
			$savedCursorVisible = [Console]::CursorVisible
			[Console]::CursorVisible = $true
			
			# Draw dialog box (exactly 35 characters per line)
			# Calculate spacing for bottom line (emojis display as 2 chars each but count as 1 in string length)
			$checkmark = [char]0x2705  # ✅ green checkmark
			$redX = [char]0x274C  # ❌ red X
			$bottomLineContent = "│  " + $checkmark + "|(u)pdate  " + $redX + "|(c)ancel"
			# Account for emojis: each emoji is 1 char in string but 2 display columns
			# String length: "│  " (3) + emoji (1) + "|" (1) + "(u)pdate  " (10) + emoji (1) + "|" (1) + "(c)ancel" (8) = 25
			# Display width: "│  " (3) + emoji (2) + "|" (1) + "(u)pdate  " (10) + emoji (2) + "|" (1) + "(c)ancel" (8) = 27
			# So we need: 35 - 27 - 1 = 7 spaces before closing │
			$bottomLineTextLength = $bottomLineContent.Length + 2  # +2 because emojis count as 2 display chars but 1 string char each
			$bottomLinePadding = Get-Padding -usedWidth ($bottomLineTextLength + 1) -totalWidth $dialogWidth
			
			# Build all lines to be exactly 35 characters using Get-Padding helper
			$line0 = "┌─────────────────────────────────┐"  # 35 chars
			$line1Text = "│  Change End Time"
			$line1Padding = Get-Padding -usedWidth ($line1Text.Length + 1) -totalWidth $dialogWidth
			$line1 = $line1Text + (" " * $line1Padding) + "│"
			
			$line2 = "│" + (" " * 33) + "│"  # 35 chars
			
			$line3Text = "│  Enter new time (HHmm format):"
			$line3Padding = Get-Padding -usedWidth ($line3Text.Length + 1) -totalWidth $dialogWidth
			$line3 = $line3Text + (" " * $line3Padding) + "│"
			
			# Line 4 will be drawn separately with highlighted field
			$line4Text = "│  "
			$line4Padding = Get-Padding -usedWidth ($line4Text.Length + 1 + 6) -totalWidth $dialogWidth  # +6 for "[    ]"
			$line4 = $line4Text + (" " * $line4Padding) + "│"
			
			$line5 = "│" + (" " * 33) + "│"  # 35 chars
			
			$line7 = "└─────────────────────────────────┘"  # 35 chars
			
			$dialogLines = @(
				$line0,
				$line1,
				$line2,
				$line3,
				$line4,
				$line5,
				$null,  # Bottom line will be written separately with colors
				$line7
			)
			
			# Draw dialog background (clear area)
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX, $dialogY + $i)
				Write-Host (" " * $dialogWidth) -NoNewline
			}
			
			# Draw dialog box
			for ($i = 0; $i -lt $dialogLines.Count; $i++) {
				if ($i -eq 1) {
					# Title line - write in magenta
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host "│  " -NoNewline
					Write-Host "Change End Time" -NoNewline -ForegroundColor Magenta
					$titleUsedWidth = 3 + "Change End Time".Length  # "│  " + title
					$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
					Write-Host (" " * $titlePadding) -NoNewline
					Write-Host "│" -NoNewline
				} elseif ($i -eq 4) {
					# Input field line - draw with highlighted field (same style as modify movement dialog)
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host "│  " -NoNewline
					# Get initial time input value for display
					$initialTimeDisplay = if ($currentEndTime -ne -1 -and $currentEndTime -ne 0) { 
						$currentEndTime.ToString().PadLeft(4, '0') 
					} else { 
						"" 
					}
					$fieldDisplay = $initialTimeDisplay.PadRight(4)
					# Draw highlighted field (same style as modify movement dialog)
					Write-Host "[" -NoNewline
					Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
					Write-Host "]" -NoNewline
					# Calculate padding to fill remaining width
					$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
					$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
					Write-Host (" " * $fieldPadding) -NoNewline
					Write-Host "│" -NoNewline
				} elseif ($i -eq 6) {
					# Bottom line - write with colored icons and hotkey letters
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host "│  " -NoNewline
					Write-Host $checkmark -NoNewline -ForegroundColor Green
					Write-Host "|" -NoNewline
					# Parse "(u)pdate" - parentheses white, letter yellow, text white
					Write-Host "(" -NoNewline
					Write-Host "u" -NoNewline -ForegroundColor Yellow
					Write-Host ")pdate  " -NoNewline
					Write-Host $redX -NoNewline -ForegroundColor Red
					Write-Host "|" -NoNewline
					# Parse "(c)ancel" - parentheses white, letter yellow, text white
					Write-Host "(" -NoNewline
					Write-Host "c" -NoNewline -ForegroundColor Yellow
					Write-Host ")ancel" -NoNewline
					Write-Host (" " * $bottomLinePadding) -NoNewline
					Write-Host "│" -NoNewline
				} else {
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host $dialogLines[$i] -NoNewline
				}
			}
			
			# Draw drop shadow
			Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
			
			# Calculate button bounds for click detection
			# Button row is at dialogY + 6 (line 6)
			$buttonRowY = $dialogY + 6
			# Update button: starts at dialogX + 3 ("│  " = 3 chars), emoji (2 display chars) + pipe (1) = 6 chars
			# "(u)pdate  " = 10 chars, so update button spans from X+3 to X+15 (inclusive)
			$updateButtonStartX = $dialogX + 3
			$updateButtonEndX = $dialogX + 15  # 3 + 2 (emoji) + 1 (pipe) + 10 (text) - 1 (inclusive)
			# Cancel button: starts after update button + spacing
			# Update button ends at X+15, then we have "  " (2 spaces) = X+17, then emoji (2) + pipe (1) = X+20
			# "(c)ancel" = 8 chars, so cancel button spans from X+20 to X+27 (inclusive)
			$cancelButtonStartX = $dialogX + 20
			$cancelButtonEndX = $dialogX + 27
			
			# Store button bounds in script scope for main loop click detection
			$script:DialogButtonBounds = @{
				buttonRowY = $buttonRowY
				updateStartX = $updateButtonStartX
				updateEndX = $updateButtonEndX
				cancelStartX = $cancelButtonStartX
				cancelEndX = $cancelButtonEndX
			}
			$script:DialogButtonClick = $null  # Clear any previous click  # 20 + 2 (emoji) + 1 (pipe) + 8 (text) - 1 (inclusive)
			
			# Position cursor in input field (inside the brackets, after "│  [")
			# Line 4 is "│  [" + 4 spaces + "]", so input starts at position 4
			$inputX = $dialogX + 4
			$inputY = $dialogY + 4
			# Don't show cursor initially - will be shown after first character is typed
			[Console]::CursorVisible = $false
			
			# Get input
			# Initialize with current end time if it exists (convert to 4-digit string)
			if ($currentEndTime -ne -1 -and $currentEndTime -ne 0) {
				$timeInput = $currentEndTime.ToString().PadLeft(4, '0')
			} else {
				$timeInput = ""
			}
			$result = $null
			$needsRedraw = $false
			$errorMessage = ""
			$isFirstChar = $true  # Track if this is the first character typed
			
			# Don't draw initial input value - cursor is hidden until first character is typed
			# Position cursor at the input field after initial draw (even if hidden, so it's ready when shown)
			[Console]::SetCursorPosition($inputX + $timeInput.Length, $inputY)
			
			# Debug: Log that dialog input loop has started
			if ($DebugMode) {
				if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
					$LogArray = @()
				}
				$LogArray += [PSCustomObject]@{
					logRow = $true
					components = @(
						@{
							priority = 1
							text = (Get-Date).ToString("HH:mm:ss")
							shortText = (Get-Date).ToString("HH:mm:ss")
						},
						@{
							priority = 2
							text = " - [DEBUG] Time dialog input loop started, button row Y: $buttonRowY"
							shortText = " - [DEBUG] Dialog started"
						}
					)
				}
			}
			
			:inputLoop do {
				# Check for window resize and update references
				$pshost = Get-Host
				$pswindow = $pshost.UI.RawUI
				$newWindowSize = $pswindow.WindowSize
				if ($newWindowSize.Width -ne $currentHostWidth -or $newWindowSize.Height -ne $currentHostHeight) {
					# Window was resized - update references and flag for main UI redraw
					# Don't force buffer size - let PowerShell manage it (allows text zoom to work)
					$HostWidthRef.Value = $newWindowSize.Width
					$HostHeightRef.Value = $newWindowSize.Height
					$currentHostWidth = $newWindowSize.Width
					$currentHostHeight = $newWindowSize.Height
					$needsRedraw = $true
					
					# Reposition dialog
					$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
					$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
					$inputX = $dialogX + 4
					$inputY = $dialogY + 4
					
					# Recalculate button bounds after repositioning
					$buttonRowY = $dialogY + 6
					$updateButtonStartX = $dialogX + 3
					$updateButtonEndX = $dialogX + 15
					$cancelButtonStartX = $dialogX + 20
					$cancelButtonEndX = $dialogX + 27
					
					# Update button bounds in script scope
					$script:DialogButtonBounds = @{
						buttonRowY = $buttonRowY
						updateStartX = $updateButtonStartX
						updateEndX = $updateButtonEndX
						cancelStartX = $cancelButtonStartX
						cancelEndX = $cancelButtonEndX
					}
					
					# Clear screen and redraw dialog
					Clear-Host
					for ($i = 0; $i -lt $dialogLines.Count; $i++) {
						if ($i -eq 1) {
							# Title line - write in magenta
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host "│  " -NoNewline
							Write-Host "Change End Time" -NoNewline -ForegroundColor Magenta
							$titleUsedWidth = 3 + "Change End Time".Length  # "│  " + title
							$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $titlePadding) -NoNewline
							Write-Host "│" -NoNewline
						} elseif ($i -eq 4) {
							# Input field line - draw with highlighted field
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host "│  " -NoNewline
							$fieldDisplay = $timeInput.PadRight(4)
							# Draw highlighted field (same style as modify movement dialog)
							Write-Host "[" -NoNewline
							Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
							Write-Host "]" -NoNewline
							# Calculate padding to fill remaining width
							$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
							$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $fieldPadding) -NoNewline
							Write-Host "│" -NoNewline
						} elseif ($i -eq 6) {
							# Bottom line - write with colored icons and hotkey letters
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host "│  " -NoNewline
							Write-Host $checkmark -NoNewline -ForegroundColor Green
							Write-Host "|" -NoNewline
							# Parse "(u)pdate" - parentheses white, letter yellow, text white
							Write-Host "(" -NoNewline
							Write-Host "u" -NoNewline -ForegroundColor Yellow
							Write-Host ")pdate  " -NoNewline
							Write-Host $redX -NoNewline -ForegroundColor Red
							Write-Host "|" -NoNewline
							# Parse "(c)ancel" - parentheses white, letter yellow, text white
							Write-Host "(" -NoNewline
							Write-Host "c" -NoNewline -ForegroundColor Yellow
							Write-Host ")ancel" -NoNewline
							Write-Host (" " * $bottomLinePadding) -NoNewline
							Write-Host "│" -NoNewline
						} else {
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host $dialogLines[$i] -NoNewline
						}
					}
					
					# Draw drop shadow
					Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
					
					# Reposition cursor and redraw input with highlight - redraw entire line 4
					[Console]::SetCursorPosition($dialogX, $inputY)
					Write-Host "│  " -NoNewline  # Redraw left border
					$fieldDisplay = $timeInput.PadRight(4)
					Write-Host "[" -NoNewline
					Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
					Write-Host "]" -NoNewline
					# Calculate padding to fill remaining width
					$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
					$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
					Write-Host (" " * $fieldPadding) -NoNewline
					Write-Host "│" -NoNewline  # Redraw right border
					# Redraw error line if there's an error
					[Console]::SetCursorPosition($dialogX, $dialogY + 5)
					if ($errorMessage -ne "") {
						Write-Host "│  " -NoNewline
						Write-Host $errorMessage -NoNewline -ForegroundColor Red
						$errorLineUsedWidth = 3 + $errorMessage.Length  # "│  " + error message
						$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
						Write-Host (" " * $errorLinePadding) -NoNewline
						Write-Host "│" -NoNewline
					} else {
						# Clear error line but keep box lines
						Write-Host "│" -NoNewline
						Write-Host (" " * ($dialogWidth - 2)) -NoNewline
						Write-Host "│" -NoNewline
					}
					# Position cursor at end of input (after opening bracket)
					[Console]::SetCursorPosition($inputX + $timeInput.Length, $inputY)
				}
				
				# Check for mouse button clicks on dialog buttons using GetAsyncKeyState (same as main menu)
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				
				# Debug: Log that we're checking for clicks in dialog
				if ($DebugMode) {
					if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
						$LogArray = @()
					}
					$LogArray += [PSCustomObject]@{
						logRow = $true
						components = @(
							@{
								priority = 1
								text = (Get-Date).ToString("HH:mm:ss")
								shortText = (Get-Date).ToString("HH:mm:ss")
							},
							@{
								priority = 2
								text = " - [DEBUG] Time dialog: Checking for mouse clicks..."
								shortText = " - [DEBUG] Checking clicks"
							}
						)
					}
				}
				
				# Debug: Log that we're checking for input (throttled to every 2 seconds)
				if ($DebugMode -and ($script:lastInputCheckTime -eq $null -or ((Get-Date) - $script:lastInputCheckTime).TotalSeconds -gt 2)) {
					$script:lastInputCheckTime = Get-Date
					if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
						$LogArray = @()
					}
					$LogArray += [PSCustomObject]@{
						logRow = $true
						components = @(
							@{
								priority = 1
								text = (Get-Date).ToString("HH:mm:ss")
								shortText = (Get-Date).ToString("HH:mm:ss")
							},
							@{
								priority = 2
								text = " - [DEBUG] Checking for mouse clicks (GetAsyncKeyState)..."
								shortText = " - [DEBUG] Checking clicks..."
							}
						)
					}
				}
				
				try {
					# Initialize previous key states if needed
					if ($null -eq $script:previousKeyStates) {
						$script:previousKeyStates = @{}
					}
					
					$leftMouseButtonCode = 0x01
					$currentKeyState = [mJiggAPI.Keyboard]::GetAsyncKeyState($leftMouseButtonCode)
					$isCurrentlyPressed = (($currentKeyState -band 0x8000) -ne 0)
					$wasJustPressed = (($currentKeyState -band 0x0001) -ne 0)
					$wasPreviouslyPressed = if ($script:previousKeyStates.ContainsKey($leftMouseButtonCode)) { $script:previousKeyStates[$leftMouseButtonCode] } else { $false }
					
					# Debug: Always log if button state is detected (not throttled)
					if ($DebugMode -and ($currentKeyState -ne 0)) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Mouse button detected! State: 0x$($currentKeyState.ToString('X4')), pressed=$isCurrentlyPressed, justPressed=$wasJustPressed, wasPrev=$wasPreviouslyPressed"
									shortText = " - [DEBUG] Mouse detected"
								}
							)
						}
					}
					
					# Debug: Log mouse button state check (throttled)
					if ($DebugMode -and ($isCurrentlyPressed -or $wasJustPressed)) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Mouse button state: pressed=$isCurrentlyPressed, justPressed=$wasJustPressed, wasPrev=$wasPreviouslyPressed"
									shortText = " - [DEBUG] Mouse state check"
								}
							)
						}
					}
					
					if ($wasJustPressed -or ($isCurrentlyPressed -and -not $wasPreviouslyPressed)) {
						# Debug: Log that we detected a click
						if ($DebugMode) {
							if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
								$LogArray = @()
							}
							$LogArray += [PSCustomObject]@{
								logRow = $true
								components = @(
									@{
										priority = 1
										text = (Get-Date).ToString("HH:mm:ss")
										shortText = (Get-Date).ToString("HH:mm:ss")
									},
									@{
										priority = 2
										text = " - [DEBUG] Mouse click detected! Starting coordinate conversion..."
										shortText = " - [DEBUG] Click detected"
									}
								)
							}
						}
						
						# Left mouse button clicked - check if it's on a dialog button
						$mousePoint = New-Object mJiggAPI.POINT
						$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
						if ($hasGetCursorPos -and [mJiggAPI.Mouse]::GetCursorPos([ref]$mousePoint)) {
							$consoleHandle = [mJiggAPI.Mouse]::GetConsoleWindow()
							if ($consoleHandle -ne [IntPtr]::Zero) {
								# Debug: Log that we got console handle
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Got console handle, screen pos: ($($mousePoint.X),$($mousePoint.Y))"
												shortText = " - [DEBUG] Got handle"
											}
										)
									}
								}
								
								$clientPoint = New-Object mJiggAPI.POINT
								$clientPoint.X = $mousePoint.X
								$clientPoint.Y = $mousePoint.Y
								if ([mJiggAPI.Mouse]::ScreenToClient($consoleHandle, [ref]$clientPoint)) {
									$windowRect = New-Object mJiggAPI.RECT
									if ([mJiggAPI.Mouse]::GetWindowRect($consoleHandle, [ref]$windowRect)) {
										$stdOutHandle = [mJiggAPI.Mouse]::GetStdHandle(-11)
										$bufferInfo = New-Object mJiggAPI.CONSOLE_SCREEN_BUFFER_INFO
										if ([mJiggAPI.Mouse]::GetConsoleScreenBufferInfo($stdOutHandle, [ref]$bufferInfo)) {
											$visibleLeft = $bufferInfo.srWindow.Left
											$visibleTop = $bufferInfo.srWindow.Top
											$visibleRight = $bufferInfo.srWindow.Right
											$visibleBottom = $bufferInfo.srWindow.Bottom
											$visibleWidth = $visibleRight - $visibleLeft + 1
											$visibleHeight = $visibleBottom - $visibleTop + 1
											$windowWidth = $windowRect.Right - $windowRect.Left
											$windowHeight = $windowRect.Bottom - $windowRect.Top
											$borderLeft = 8
											$borderTop = 30
											$borderRight = 8
											$borderBottom = 8
											$clientWidth = $windowWidth - $borderLeft - $borderRight
											$clientHeight = $windowHeight - $borderTop - $borderBottom
											$charWidth = $clientWidth / $visibleWidth
											$charHeight = $clientHeight / $visibleHeight
											$adjustedX = $clientPoint.X - $borderLeft
											$adjustedY = $clientPoint.Y - $borderTop
											$consoleX = [Math]::Floor($adjustedX / $charWidth) + $visibleLeft
											$consoleY = [Math]::Floor($adjustedY / $charHeight) + $visibleTop
											
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												$LogArray += [PSCustomObject]@{
													logRow = $true
													components = @(
														@{
															priority = 1
															text = (Get-Date).ToString("HH:mm:ss")
															shortText = (Get-Date).ToString("HH:mm:ss")
														},
														@{
															priority = 2
															text = " - [DEBUG] Mouse click detected at console ($consoleX,$consoleY)"
															shortText = " - [DEBUG] Click ($consoleX,$consoleY)"
														}
													)
												}
											}
											
											# Debug: Log button bounds for reference
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												$LogArray += [PSCustomObject]@{
													logRow = $true
													components = @(
														@{
															priority = 1
															text = (Get-Date).ToString("HH:mm:ss")
															shortText = (Get-Date).ToString("HH:mm:ss")
														},
														@{
															priority = 2
															text = " - [DEBUG] Button bounds - Row Y: $buttonRowY, Update: X$updateButtonStartX-$updateButtonEndX, Cancel: X$cancelButtonStartX-$cancelButtonEndX"
															shortText = " - [DEBUG] Button bounds"
														}
													)
												}
											}
											
											# Check if click is on update button
											$clickedButton = "none"
											$isOnButton = $false
											if ($consoleY -eq $buttonRowY -or $consoleY -eq ($buttonRowY - 1) -or $consoleY -eq ($buttonRowY + 1)) {
												if ($consoleX -ge $updateButtonStartX -and $consoleX -le $updateButtonEndX) {
													$clickedButton = "Update"
													$isOnButton = $true
													# Update button clicked - trigger update action
													$char = "u"
													$keyProcessed = $true
												} elseif ($consoleX -ge $cancelButtonStartX -and $consoleX -le $cancelButtonEndX) {
													$clickedButton = "Cancel"
													$isOnButton = $true
													# Cancel button clicked - trigger cancel action
													$char = "c"
													$keyProcessed = $true
												}
											}
											
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												if ($isOnButton) {
													$LogArray += [PSCustomObject]@{
														logRow = $true
														components = @(
															@{
																priority = 1
																text = (Get-Date).ToString("HH:mm:ss")
																shortText = (Get-Date).ToString("HH:mm:ss")
															},
															@{
																priority = 2
																text = " - [DEBUG] Button clicked: $clickedButton"
																shortText = " - [DEBUG] $clickedButton"
															}
														)
													}
												} else {
													$LogArray += [PSCustomObject]@{
														logRow = $true
														components = @(
															@{
																priority = 1
																text = (Get-Date).ToString("HH:mm:ss")
																shortText = (Get-Date).ToString("HH:mm:ss")
															},
															@{
																priority = 2
																text = " - [DEBUG] Click NOT on button (row Y: $buttonRowY, update X: $updateButtonStartX-$updateButtonEndX, cancel X: $cancelButtonStartX-$cancelButtonEndX)"
																shortText = " - [DEBUG] Not on button"
															}
														)
													}
												}
											}
										} else {
											# Debug: Log if GetConsoleScreenBufferInfo failed
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												$LogArray += [PSCustomObject]@{
													logRow = $true
													components = @(
														@{
															priority = 1
															text = (Get-Date).ToString("HH:mm:ss")
															shortText = (Get-Date).ToString("HH:mm:ss")
														},
														@{
															priority = 2
															text = " - [DEBUG] Failed to get console screen buffer info"
															shortText = " - [DEBUG] Buffer info failed"
														}
													)
												}
											}
										}
									} else {
										# Debug: Log if GetWindowRect failed
										if ($DebugMode) {
											if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
												$LogArray = @()
											}
											$LogArray += [PSCustomObject]@{
												logRow = $true
												components = @(
													@{
														priority = 1
														text = (Get-Date).ToString("HH:mm:ss")
														shortText = (Get-Date).ToString("HH:mm:ss")
													},
													@{
														priority = 2
														text = " - [DEBUG] Failed to get window rect"
														shortText = " - [DEBUG] Window rect failed"
													}
												)
											}
										}
									}
								} else {
									# Debug: Log if ScreenToClient failed
									if ($DebugMode) {
										if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
											$LogArray = @()
										}
										$LogArray += [PSCustomObject]@{
											logRow = $true
											components = @(
												@{
													priority = 1
													text = (Get-Date).ToString("HH:mm:ss")
													shortText = (Get-Date).ToString("HH:mm:ss")
												},
												@{
													priority = 2
													text = " - [DEBUG] Failed to convert screen to client coordinates"
													shortText = " - [DEBUG] ScreenToClient failed"
												}
											)
										}
									}
								}
							} else {
								# Debug: Log if console handle is zero
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Console handle is zero"
												shortText = " - [DEBUG] Handle zero"
											}
										)
									}
								}
							}
						} else {
							# Debug: Log if GetCursorPos failed
							if ($DebugMode) {
								if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
									$LogArray = @()
								}
								$LogArray += [PSCustomObject]@{
									logRow = $true
									components = @(
										@{
											priority = 1
											text = (Get-Date).ToString("HH:mm:ss")
											shortText = (Get-Date).ToString("HH:mm:ss")
										},
										@{
											priority = 2
											text = " - [DEBUG] Failed to get cursor position (hasGetCursorPos=$hasGetCursorPos)"
											shortText = " - [DEBUG] GetCursorPos failed"
										}
									)
								}
							}
						}
					}
					# Update previous key state
					$script:previousKeyStates[$leftMouseButtonCode] = $isCurrentlyPressed
				} catch {
					# Ignore errors in mouse click detection
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Mouse click detection error: $($_.Exception.Message)"
									shortText = " - [DEBUG] Error: $($_.Exception.Message)"
								}
							)
						}
					}
				}
				
				# Check for dialog button clicks detected by main loop
				if ($null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null  # Clear it after using
					
					if ($buttonClick -eq "Update") {
						$char = "u"
						$keyProcessed = $true
					} elseif ($buttonClick -eq "Cancel") {
						$char = "c"
						$keyProcessed = $true
					}
				}
				
				# Wait for key input (non-blocking check)
				# Read keys and only process key UP events (same as main menu)
				if (-not $keyProcessed) {
					while ($Host.UI.RawUI.KeyAvailable -and -not $keyProcessed) {
						$keyInfo = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyup,AllowCtrlC")
						$isKeyDown = $false
						if ($null -ne $keyInfo.KeyDown) {
							$isKeyDown = $keyInfo.KeyDown
						}
						
						# Only process key UP events (skip key down)
						if (-not $isKeyDown) {
							$key = $keyInfo.Key
							$char = $keyInfo.Character
							$keyProcessed = $true
						}
					}
				}
				
				if (-not $keyProcessed) {
					# No key available, sleep briefly and check for resize again
					Start-Sleep -Milliseconds 50
					continue
				}
				
				if ($char -eq "u" -or $char -eq "U" -or $key -eq "Enter" -or $char -eq [char]13 -or $char -eq [char]10) {
					# Update - allow blank input to clear end time (Enter key also works as hidden function)
					# Debug: Log time dialog update
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$updateValue = if ($timeInput.Length -eq 0) { "cleared" } else { $timeInput }
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Time dialog: Update clicked (value: $updateValue)"
									shortText = " - [DEBUG] Time: Update"
								}
							)
						}
					}
					if ($timeInput.Length -eq 0) {
						# Blank input - clear end time (use -1 as special value)
						$result = -1
						break
					} elseif ($timeInput.Length -eq 1 -and $timeInput -eq "0") {
						# Single "0" = no end time
						$result = -1
						break
					} elseif ($timeInput.Length -eq 2) {
						# 2 digits entered - treat as hours, auto-fill minutes as 00
						try {
							$hours = [int]$timeInput
							if ($hours -ge 0 -and $hours -le 23) {
								# Valid hours - pad with "00" for minutes
								$timeInput = $timeInput.PadRight(4, '0')
								$newTime = [int]$timeInput
								$result = $newTime
								break
							} else {
								# Invalid hours - show error
								$errorMessage = "Hours out of range (00-23)"
								# Redraw input field with highlight - redraw entire line 4
								[Console]::SetCursorPosition($dialogX, $inputY)
								Write-Host "│  " -NoNewline  # Redraw left border
								$fieldDisplay = $timeInput.PadRight(4)
								Write-Host "[" -NoNewline
								Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
								Write-Host "]" -NoNewline
								# Calculate padding to fill remaining width
								$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
								$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
								Write-Host (" " * $fieldPadding) -NoNewline
								Write-Host "│" -NoNewline  # Redraw right border
								# Show error
								[Console]::SetCursorPosition($dialogX, $dialogY + 5)
								Write-Host "│  " -NoNewline
								Write-Host $errorMessage -NoNewline -ForegroundColor Red
								$errorLineUsedWidth = 3 + $errorMessage.Length  # "│  " + error message
								$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
								Write-Host (" " * $errorLinePadding) -NoNewline
								Write-Host "│" -NoNewline
								[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
							}
						} catch {
							# Invalid input - show error
							$errorMessage = "Invalid hours"
							# Redraw input field with highlight - redraw entire line 4
							[Console]::SetCursorPosition($dialogX, $inputY)
							Write-Host "│  " -NoNewline  # Redraw left border
							$fieldDisplay = $timeInput.PadRight(4)
							Write-Host "[" -NoNewline
							Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
							Write-Host "]" -NoNewline
							# Calculate padding to fill remaining width
							$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
							$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $fieldPadding) -NoNewline
							Write-Host "│" -NoNewline  # Redraw right border
							# Show error
							[Console]::SetCursorPosition($dialogX, $dialogY + 5)
							Write-Host "│  " -NoNewline
							Write-Host $errorMessage -NoNewline -ForegroundColor Red
							$errorLineUsedWidth = 3 + $errorMessage.Length  # "│  " + error message
							$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $errorLinePadding) -NoNewline
							Write-Host "│" -NoNewline
							[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
						}
					} elseif ($timeInput.Length -eq 4) {
						try {
							$newTime = [int]$timeInput
							# Validate time format: HHmm where HH is 00-23 and mm is 00-59
							$hours = [int]$timeInput.Substring(0, 2)
							$minutes = [int]$timeInput.Substring(2, 2)
							
							# "0000" is midnight (12:00 AM), not "no end time"
							if ($newTime -ge 0 -and $newTime -le 2359 -and $hours -le 23 -and $minutes -le 59) {
								$result = $newTime
								break
							} else {
								# Invalid time - show error
								if ($hours -gt 23) {
									$errorMessage = "Hours out of range (00-23)"
								} elseif ($minutes -gt 59) {
									$errorMessage = "Minutes out of range (00-59)"
								} else {
									$errorMessage = "Time out of range (0000-2359)"
								}
								# Redraw input field with highlight - redraw entire line 4
								[Console]::SetCursorPosition($dialogX, $inputY)
								Write-Host "│  " -NoNewline  # Redraw left border
								$fieldDisplay = $timeInput.PadRight(4)
								Write-Host "[" -NoNewline
								Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
								Write-Host "]" -NoNewline
								# Calculate padding to fill remaining width
								$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
								$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
								Write-Host (" " * $fieldPadding) -NoNewline
								Write-Host "│" -NoNewline  # Redraw right border
								# Show error
								[Console]::SetCursorPosition($dialogX, $dialogY + 5)
								Write-Host "│  " -NoNewline
								Write-Host $errorMessage -NoNewline -ForegroundColor Red
								$errorLineUsedWidth = 3 + $errorMessage.Length  # "│  " + error message
								$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
								Write-Host (" " * $errorLinePadding) -NoNewline
								Write-Host "│" -NoNewline
								[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
							}
						} catch {
							# Invalid input - show error (shouldn't normally happen with numeric-only input)
							$errorMessage = "Number out of range"
							# Redraw input field with highlight - redraw entire line 4
							[Console]::SetCursorPosition($dialogX, $inputY)
							Write-Host "│  " -NoNewline  # Redraw left border
							$fieldDisplay = $timeInput.PadRight(4)
							Write-Host "[" -NoNewline
							Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
							Write-Host "]" -NoNewline
							# Calculate padding to fill remaining width
							$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
							$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $fieldPadding) -NoNewline
							Write-Host "│" -NoNewline  # Redraw right border
							# Show error
							[Console]::SetCursorPosition($dialogX, $dialogY + 5)
							Write-Host "│  " -NoNewline
							Write-Host $errorMessage -NoNewline -ForegroundColor Red
							$errorLineUsedWidth = 3 + $errorMessage.Length  # "│  " + error message
							$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $errorLinePadding) -NoNewline
							Write-Host "│" -NoNewline
							[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
						}
					} else {
						# Not 4 digits yet - show error
						$errorMessage = "Enter 4 digits (HHmm format)"
						# Redraw input field with highlight - redraw entire line 4
						[Console]::SetCursorPosition($dialogX, $inputY)
						Write-Host "│  " -NoNewline  # Redraw left border
						$fieldDisplay = $timeInput.PadRight(4)
						Write-Host "[" -NoNewline
						Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
						Write-Host "]" -NoNewline
						# Calculate padding to fill remaining width
						$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
						$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
						Write-Host (" " * $fieldPadding) -NoNewline
						Write-Host "│" -NoNewline  # Redraw right border
						# Show error
						[Console]::SetCursorPosition($dialogX, $dialogY + 5)
						Write-Host "│  " -NoNewline
						Write-Host $errorMessage -NoNewline -ForegroundColor Red
						$errorLineUsedWidth = 3 + $errorMessage.Length  # "│  " + error message
						$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
						Write-Host (" " * $errorLinePadding) -NoNewline
						Write-Host "│" -NoNewline
						[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
					}
				} elseif ($char -eq "c" -or $char -eq "C" -or $char -eq "t" -or $char -eq "T" -or $key -eq "Escape" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 27)) {
					# Cancel (Escape key and 't' key also work as hidden functions)
					# Debug: Log time dialog cancel
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Time dialog: Cancel clicked"
									shortText = " - [DEBUG] Time: Cancel"
								}
							)
						}
					}
					$result = $null
					$needsRedraw = $false  # No redraw needed on cancel
					break
				} elseif ($key -eq "Backspace" -or $char -eq [char]8 -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 8)) {
					# Backspace - handle multiple ways to ensure it works
					# Ensure value is a string (in case it was somehow set as a char)
					$timeInput = $timeInput.ToString()
					if ($timeInput.Length -gt 0) {
						$timeInput = $timeInput.Substring(0, $timeInput.Length - 1)
						# If field is now empty, reset the input tracking so next char will clear again
						if ($timeInput.Length -eq 0) {
							$isFirstChar = $true
							[Console]::CursorVisible = $false  # Hide cursor when field is empty
						}
						$errorMessage = ""  # Clear error when editing
						# Redraw input with highlight - redraw entire line 4 to ensure clean overwrite
						[Console]::SetCursorPosition($dialogX, $inputY)
						Write-Host "│  " -NoNewline  # Redraw left border
						$fieldDisplay = $timeInput.PadRight(4)
						Write-Host "[" -NoNewline
						Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
						Write-Host "]" -NoNewline
						# Calculate padding to fill remaining width
						$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
						$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
						Write-Host (" " * $fieldPadding) -NoNewline
						Write-Host "│" -NoNewline  # Redraw right border
						# Clear and redraw error line with box lines
						[Console]::SetCursorPosition($dialogX, $dialogY + 5)
						Write-Host "│" -NoNewline
						Write-Host (" " * 33) -NoNewline
						Write-Host "│" -NoNewline
						# Position cursor at end of input (only if field has content, after opening bracket)
						if ($timeInput.Length -gt 0) {
							[Console]::SetCursorPosition($inputX + $timeInput.Length, $inputY)
						}
					}
				} elseif ($char -match "[0-9]") {
					# Numeric input
					# If this is the first character typed, clear the field first and show cursor
					if ($isFirstChar) {
						$timeInput = $char.ToString()  # Convert char to string
						$isFirstChar = $false
						[Console]::CursorVisible = $true  # Show cursor after first character
					} elseif ($timeInput.Length -lt 4) {
						$timeInput += $char.ToString()  # Convert char to string
					}
					$errorMessage = ""  # Clear error when typing
					# Redraw input field with highlight - redraw entire line 4 to ensure clean overwrite
					[Console]::SetCursorPosition($dialogX, $inputY)
					Write-Host "│  " -NoNewline  # Redraw left border
					$fieldDisplay = $timeInput.PadRight(4)
					Write-Host "[" -NoNewline
					Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
					Write-Host "]" -NoNewline
					# Calculate padding to fill remaining width
					$fieldUsedWidth = 3 + 6  # "│  " + "[    ]"
					$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
					Write-Host (" " * $fieldPadding) -NoNewline
					Write-Host "│" -NoNewline  # Redraw right border
					# Clear error line if it was showing, redraw box lines
					[Console]::SetCursorPosition($dialogX, $dialogY + 5)
					Write-Host "│" -NoNewline
					Write-Host (" " * 33) -NoNewline
					Write-Host "│" -NoNewline
					# Position cursor at end of input (after opening bracket)
					[Console]::SetCursorPosition($inputX + $timeInput.Length, $inputY)
				}
				
				# Clear any remaining keys in buffer after processing
				try {
					while ($Host.UI.RawUI.KeyAvailable) {
						$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
					}
				} catch {
					# Silently ignore - buffer might not be clearable
				}
			} until ($false)
			
			# Clear shadow before clearing dialog area
			Clear-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
			
			# Clear dialog area
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX, $dialogY + $i)
				Write-Host (" " * $dialogWidth) -NoNewline
			}
			
			# Restore cursor visibility
			[Console]::CursorVisible = $savedCursorVisible
			
			# Clear dialog button bounds when dialog closes
			$script:DialogButtonBounds = $null
			$script:DialogButtonClick = $null
			
			# Return result object with result and redraw flag
			return @{
				Result = $result
				NeedsRedraw = $needsRedraw
			}
		}

		# Helper function: Calculate padding needed to fill remaining width
		function Get-Padding {
			param(
				[int]$usedWidth,
				[int]$totalWidth
			)
			return [Math]::Max(0, $totalWidth - $usedWidth)
		}
		
		# Helper function: Draw a horizontal line in a section
		function Write-SectionLine {
			param(
				[int]$x,
				[int]$y,
				[int]$width,
				[string]$leftChar = "│",
				[string]$rightChar = "│",
				[string]$fillChar = " ",
				[System.ConsoleColor]$borderColor = [System.ConsoleColor]::White,
				[System.ConsoleColor]$fillColor = [System.ConsoleColor]::White
			)
			
			[Console]::SetCursorPosition($x, $y)
			Write-Host $leftChar -NoNewline -ForegroundColor $borderColor
			$fillWidth = $width - 2  # Subtract left and right border
			Write-Host ($fillChar * $fillWidth) -NoNewline -ForegroundColor $fillColor
			Write-Host $rightChar -NoNewline -ForegroundColor $borderColor
		}
		
		# Helper function: Draw a simple dialog row (no description box)
		function Write-SimpleDialogRow {
			param(
				[int]$x,
				[int]$y,
				[int]$width,
				[string]$content = "",
				[System.ConsoleColor]$contentColor = [System.ConsoleColor]::White
			)
			
			[Console]::SetCursorPosition($x, $y)
			Write-Host "│" -NoNewline
			if ($content.Length -gt 0) {
				Write-Host " " -NoNewline
				Write-Host $content -NoNewline -ForegroundColor $contentColor
				$usedWidth = 1 + 1 + $content.Length  # │ + space + content
				$padding = Get-Padding -usedWidth ($usedWidth + 1) -totalWidth $width
				Write-Host (" " * $padding) -NoNewline
			} else {
				# Empty line
				Write-Host (" " * ($width - 2)) -NoNewline
			}
			Write-Host "│" -NoNewline
		}
		
		# Helper function: Draw a field row with input box (no description box)
		function Write-SimpleFieldRow {
			param(
				[int]$x,
				[int]$y,
				[int]$width,
				[string]$label,
				[int]$longestLabel,
				[string]$fieldValue,
				[int]$fieldWidth,
				[int]$fieldIndex,
				[int]$currentFieldIndex
			)
			
			[Console]::SetCursorPosition($x, $y)
			
			# Calculate label padding for alignment
			$labelPadding = [Math]::Max(0, $longestLabel - $label.Length)
			$labelText = "│  " + $label + (" " * $labelPadding)
			
			# Format field value
			$fieldDisplay = if ([string]::IsNullOrEmpty($fieldValue)) { "" } else { $fieldValue }
			$fieldDisplay = $fieldDisplay.PadRight($fieldWidth)
			$fieldContent = "[" + $fieldDisplay + "]"
			
			# Write label
			Write-Host $labelText -NoNewline
			
			# Write field (highlighted if current field)
			if ($fieldIndex -eq $currentFieldIndex) {
				Write-Host "[" -NoNewline
				Write-Host $fieldDisplay -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
				Write-Host "]" -NoNewline
			} else {
				Write-Host $fieldContent -NoNewline -ForegroundColor Cyan
			}
			
			# Calculate and write padding to fill remaining width
			$usedWidth = $labelText.Length + $fieldContent.Length
			$remainingPadding = Get-Padding -usedWidth ($usedWidth + 1) -totalWidth $width
			Write-Host (" " * $remainingPadding) -NoNewline
			Write-Host "│" -NoNewline
		}
		
		function Show-MovementModifyDialog {
			param(
				[double]$currentIntervalSeconds,
				[double]$currentIntervalVariance,
				[double]$currentMoveSpeed,
				[double]$currentMoveVariance,
				[double]$currentTravelDistance,
				[double]$currentTravelVariance,
				[double]$currentAutoResumeDelaySeconds,
				[ref]$HostWidthRef,
				[ref]$HostHeightRef
			)
			
			# Get current host dimensions from references
			$currentHostWidth = $HostWidthRef.Value
			$currentHostHeight = $HostHeightRef.Value
			
			# Dialog dimensions - simplified (no description box)
			$dialogWidth = 30  # Width for parameters section (reduced by 20)
			$dialogHeight = 17  # Increased by 1 for new auto-resume delay field
			$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
			$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
			
			# Save current cursor position and visibility
			$savedCursorVisible = [Console]::CursorVisible
			[Console]::CursorVisible = $false
			
			# Input field values
			$intervalSecondsInput = $currentIntervalSeconds.ToString()
			$intervalVarianceInput = $currentIntervalVariance.ToString()
			$moveSpeedInput = $currentMoveSpeed.ToString()
			$moveVarianceInput = $currentMoveVariance.ToString()
			$travelDistanceInput = $currentTravelDistance.ToString()
			$travelVarianceInput = $currentTravelVariance.ToString()
			$autoResumeDelaySecondsInput = $currentAutoResumeDelaySeconds.ToString()
			
			# Current field index (0-6)
			$currentField = 0
			$errorMessage = ""
			$lastFieldWithInput = -1  # Track which field last received input (to detect first character)
			
			# Field positions (Y coordinates relative to dialogY)
			$fieldYPositions = @(5, 6, 9, 10, 13, 14, 15)  # Y positions for each input field
			$fieldWidth = 6  # Width of input field (max 6 characters)
			# Calculate the longest label width for alignment
			$label1 = [Math]::Max("Interval (sec): ".Length, "Variance (sec): ".Length)
			$label2 = [Math]::Max("Distance (px): ".Length, "Variance (px): ".Length)
			$label3 = [Math]::Max($label1, $label2)
			$label4 = [Math]::Max($label3, "Speed (sec): ".Length)
			$longestLabel = [Math]::Max($label4, "Delay (sec): ".Length)
			$inputBoxStartX = 3 + $longestLabel  # "│  " + longest label = X position where all input boxes start
			
			# Draw dialog function - simplified (no description box)
			$drawDialog = {
				param($x, $y, $width, $height, $currentFieldIndex, $errorMsg, $inputBoxStartXPos, $fieldWidthValue, $intervalSec, $intervalVar, $moveSpeed, $moveVar, $travelDist, $travelDistVar, $autoResumeDelaySec)
				
				$fieldWidth = $fieldWidthValue
				
				# Calculate longest label for alignment
				$label1 = [Math]::Max("Interval (sec): ".Length, "Variance (sec): ".Length)
				$label2 = [Math]::Max("Distance (px): ".Length, "Variance (px): ".Length)
				$label3 = [Math]::Max($label1, $label2)
				$label4 = [Math]::Max($label3, "Speed (sec): ".Length)
				$longestLabel = [Math]::Max($label4, "Delay (sec): ".Length)
				
				# Define field data structure
				$fields = @(
					@{ Index = 0; Label = "Interval (sec): "; Value = $intervalSec },
					@{ Index = 1; Label = "Variance (sec): "; Value = $intervalVar },
					@{ Index = 2; Label = "Distance (px): "; Value = $travelDist },
					@{ Index = 3; Label = "Variance (px): "; Value = $travelDistVar },
					@{ Index = 4; Label = "Speed (sec): "; Value = $moveSpeed },
					@{ Index = 5; Label = "Variance (sec): "; Value = $moveVar },
					@{ Index = 6; Label = "Delay (sec): "; Value = $autoResumeDelaySec }
				)
				
				# Clear dialog area
				for ($i = 0; $i -lt $height; $i++) {
					[Console]::SetCursorPosition($x, $y + $i)
					Write-Host (" " * $width) -NoNewline
				}
				
				# Top border (spans full width)
				[Console]::SetCursorPosition($x, $y)
				Write-Host "┌" -NoNewline
				Write-Host ("─" * ($width - 2)) -NoNewline
				Write-Host "┐" -NoNewline
				
				# Title line
				Write-SimpleDialogRow -x $x -y ($y + 1) -width $width -content "Modify Movement Settings" -contentColor Magenta
				
				# Empty line (row 2)
				Write-SimpleDialogRow -x $x -y ($y + 2) -width $width
				
				# Interval section header (row 3)
				Write-SimpleDialogRow -x $x -y ($y + 3) -width $width -content "Interval:" -contentColor Yellow
				
				# Interval fields (rows 4-5)
				Write-SimpleFieldRow -x $x -y ($y + 4) -width $width `
					-label $fields[0].Label -longestLabel $longestLabel -fieldValue $fields[0].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[0].Index -currentFieldIndex $currentFieldIndex
				
				Write-SimpleFieldRow -x $x -y ($y + 5) -width $width `
					-label $fields[1].Label -longestLabel $longestLabel -fieldValue $fields[1].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[1].Index -currentFieldIndex $currentFieldIndex
				
				# Travel Distance section header (row 6)
				Write-SimpleDialogRow -x $x -y ($y + 6) -width $width -content "Travel Distance:" -contentColor Yellow
				
				# Travel Distance fields (rows 7-8)
				Write-SimpleFieldRow -x $x -y ($y + 7) -width $width `
					-label $fields[2].Label -longestLabel $longestLabel -fieldValue $fields[2].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[2].Index -currentFieldIndex $currentFieldIndex
				
				Write-SimpleFieldRow -x $x -y ($y + 8) -width $width `
					-label $fields[3].Label -longestLabel $longestLabel -fieldValue $fields[3].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[3].Index -currentFieldIndex $currentFieldIndex
				
				# Movement Speed section header (row 9)
				Write-SimpleDialogRow -x $x -y ($y + 9) -width $width -content "Movement Speed:" -contentColor Yellow
				
				# Movement Speed fields (rows 10-11)
				Write-SimpleFieldRow -x $x -y ($y + 10) -width $width `
					-label $fields[4].Label -longestLabel $longestLabel -fieldValue $fields[4].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[4].Index -currentFieldIndex $currentFieldIndex
				
				Write-SimpleFieldRow -x $x -y ($y + 11) -width $width `
					-label $fields[5].Label -longestLabel $longestLabel -fieldValue $fields[5].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[5].Index -currentFieldIndex $currentFieldIndex
				
				# Auto-Resume Delay section header (row 12)
				Write-SimpleDialogRow -x $x -y ($y + 12) -width $width -content "Auto-Resume Delay:" -contentColor Yellow
				
				# Auto-Resume Delay field (row 13)
				Write-SimpleFieldRow -x $x -y ($y + 13) -width $width `
					-label $fields[6].Label -longestLabel $longestLabel -fieldValue $fields[6].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[6].Index -currentFieldIndex $currentFieldIndex
				
				# Empty line (row 14)
				Write-SimpleDialogRow -x $x -y ($y + 14) -width $width
				
				# Error line (row 15)
				if ($errorMsg) {
					Write-SimpleDialogRow -x $x -y ($y + 15) -width $width -content $errorMsg -contentColor Red
				} else {
					Write-SimpleDialogRow -x $x -y ($y + 15) -width $width
				}
				
				# Bottom line with buttons (row 16)
				$checkmark = [char]0x2705
				$redX = [char]0x274C
				[Console]::SetCursorPosition($x, $y + 16)
				Write-Host "│" -NoNewline
				Write-Host " " -NoNewline
				Write-Host $checkmark -NoNewline -ForegroundColor Green
				Write-Host "|" -NoNewline
				Write-Host "(" -NoNewline
				Write-Host "u" -NoNewline -ForegroundColor Yellow
				Write-Host ")pdate  " -NoNewline
				Write-Host $redX -NoNewline -ForegroundColor Red
				Write-Host "|" -NoNewline
				Write-Host "(" -NoNewline
				Write-Host "c" -NoNewline -ForegroundColor Yellow
				Write-Host ")ancel" -NoNewline
				# Account for emojis: each emoji is 1 char in string but 2 display columns
				$buttonText = $checkmark + "|(u)pdate  " + $redX + "|(c)ancel"
				$buttonTextLength = $buttonText.Length + 2  # +2 because emojis count as 2 display chars but 1 string char each
				$buttonUsedWidth = 1 + 1 + $buttonTextLength  # │ + space + button text
				$buttonPadding = Get-Padding -usedWidth ($buttonUsedWidth + 1) -totalWidth $width
				Write-Host (" " * $buttonPadding) -NoNewline
				Write-Host "│" -NoNewline
				
				# Bottom border (spans full width)
				[Console]::SetCursorPosition($x, $y + 17)
				Write-Host "└" -NoNewline
				Write-Host ("─" * ($width - 2)) -NoNewline
				Write-Host "┘" -NoNewline
				
				# Draw drop shadow
				Draw-DialogShadow -dialogX $x -dialogY $y -dialogWidth $width -dialogHeight $height
			}
			
			# Initial draw
			& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
			
			# Calculate button bounds for click detection
			# Button row is at dialogY + 16 (row 16)
			$buttonRowY = $dialogY + 16
			# Update button: starts at dialogX + 2 ("│ " = 2 chars), emoji (2 display chars) + pipe (1) = 5 chars
			# "(u)pdate  " = 10 chars, so update button spans from X+2 to X+14 (inclusive)
			$updateButtonStartX = $dialogX + 2
			$updateButtonEndX = $dialogX + 14  # 2 + 2 (emoji) + 1 (pipe) + 10 (text) - 1 (inclusive)
			# Cancel button: starts after update button + spacing
			# Update button ends at X+14, then we have "  " (2 spaces) = X+16, then emoji (2) + pipe (1) = X+19
			# "(c)ancel" = 8 chars, so cancel button spans from X+19 to X+26 (inclusive)
			$cancelButtonStartX = $dialogX + 19
			$cancelButtonEndX = $dialogX + 26
			
			# Store button bounds in script scope for main loop click detection
			$script:DialogButtonBounds = @{
				buttonRowY = $buttonRowY
				updateStartX = $updateButtonStartX
				updateEndX = $updateButtonEndX
				cancelStartX = $cancelButtonStartX
				cancelEndX = $cancelButtonEndX
			}
			$script:DialogButtonClick = $null  # Clear any previous click  # 19 + 2 (emoji) + 1 (pipe) + 8 (text) - 1 (inclusive)
			
			# Position cursor at the first field after initial draw (ready for input)
			$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
			$fieldY = $dialogY + $fieldYOffsets[$currentField]
			# Get the current field's input value
			$currentInputRef = switch ($currentField) {
				0 { [ref]$intervalSecondsInput }
				1 { [ref]$intervalVarianceInput }
				2 { [ref]$travelDistanceInput }
				3 { [ref]$travelVarianceInput }
				4 { [ref]$moveSpeedInput }
				5 { [ref]$moveVarianceInput }
				6 { [ref]$autoResumeDelaySecondsInput }
			}
			# Cursor X: dialogX + inputBoxStartX + 1 (for opening bracket) + length of actual value
			$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
			[Console]::SetCursorPosition($cursorX, $fieldY)
			$result = $null
			$needsRedraw = $false
			
			# Debug: Log that dialog input loop has started
			if ($DebugMode) {
				if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
					$LogArray = @()
				}
				$LogArray += [PSCustomObject]@{
					logRow = $true
					components = @(
						@{
							priority = 1
							text = (Get-Date).ToString("HH:mm:ss")
							shortText = (Get-Date).ToString("HH:mm:ss")
						},
						@{
							priority = 2
							text = " - [DEBUG] Movement dialog input loop started, button row Y: $buttonRowY"
							shortText = " - [DEBUG] Dialog started"
						}
					)
				}
			}
			
			:inputLoop do {
				# Check for window resize
				$pshost = Get-Host
				$pswindow = $pshost.UI.RawUI
				$newWindowSize = $pswindow.WindowSize
				if ($newWindowSize.Width -ne $currentHostWidth -or $newWindowSize.Height -ne $currentHostHeight) {
					$HostWidthRef.Value = $newWindowSize.Width
					$HostHeightRef.Value = $newWindowSize.Height
					$currentHostWidth = $newWindowSize.Width
					$currentHostHeight = $newWindowSize.Height
					$needsRedraw = $true
					$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
					$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
					
					# Recalculate button bounds after repositioning
					$buttonRowY = $dialogY + 16
					$updateButtonStartX = $dialogX + 2
					$updateButtonEndX = $dialogX + 14
					$cancelButtonStartX = $dialogX + 19
					$cancelButtonEndX = $dialogX + 26
					
					# Update button bounds in script scope
					$script:DialogButtonBounds = @{
						buttonRowY = $buttonRowY
						updateStartX = $updateButtonStartX
						updateEndX = $updateButtonEndX
						cancelStartX = $cancelButtonStartX
						cancelEndX = $cancelButtonEndX
					}
					
					Clear-Host
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
					# Position cursor at the active field after resize
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldY = $dialogY + $fieldYOffsets[$currentField]
					$currentInputRef = switch ($currentField) {
						0 { [ref]$intervalSecondsInput }
						1 { [ref]$intervalVarianceInput }
						2 { [ref]$travelDistanceInput }
						3 { [ref]$travelVarianceInput }
						4 { [ref]$moveSpeedInput }
						5 { [ref]$moveVarianceInput }
						6 { [ref]$autoResumeDelaySecondsInput }
					}
					$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
					[Console]::SetCursorPosition($cursorX, $fieldY)
				}
				
				# Check for mouse button clicks on dialog buttons using GetAsyncKeyState (same as main menu)
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				
				# Debug: Log that we're checking for clicks in dialog
				if ($DebugMode) {
					if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
						$LogArray = @()
					}
					$LogArray += [PSCustomObject]@{
						logRow = $true
						components = @(
							@{
								priority = 1
								text = (Get-Date).ToString("HH:mm:ss")
								shortText = (Get-Date).ToString("HH:mm:ss")
							},
							@{
								priority = 2
								text = " - [DEBUG] Time dialog: Checking for mouse clicks..."
								shortText = " - [DEBUG] Checking clicks"
							}
						)
					}
				}
				
				# Debug: Log that we're checking for input (throttled to every 2 seconds)
				if ($DebugMode -and ($script:lastInputCheckTime -eq $null -or ((Get-Date) - $script:lastInputCheckTime).TotalSeconds -gt 2)) {
					$script:lastInputCheckTime = Get-Date
					if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
						$LogArray = @()
					}
					$LogArray += [PSCustomObject]@{
						logRow = $true
						components = @(
							@{
								priority = 1
								text = (Get-Date).ToString("HH:mm:ss")
								shortText = (Get-Date).ToString("HH:mm:ss")
							},
							@{
								priority = 2
								text = " - [DEBUG] Checking for mouse clicks (GetAsyncKeyState)..."
								shortText = " - [DEBUG] Checking clicks..."
							}
						)
					}
				}
				
				try {
					# Initialize previous key states if needed
					if ($null -eq $script:previousKeyStates) {
						$script:previousKeyStates = @{}
					}
					
					$leftMouseButtonCode = 0x01
					$currentKeyState = [mJiggAPI.Keyboard]::GetAsyncKeyState($leftMouseButtonCode)
					$isCurrentlyPressed = (($currentKeyState -band 0x8000) -ne 0)
					$wasJustPressed = (($currentKeyState -band 0x0001) -ne 0)
					$wasPreviouslyPressed = if ($script:previousKeyStates.ContainsKey($leftMouseButtonCode)) { $script:previousKeyStates[$leftMouseButtonCode] } else { $false }
					
					# Debug: Always log if button state is detected (not throttled)
					if ($DebugMode -and ($currentKeyState -ne 0)) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Mouse button detected! State: 0x$($currentKeyState.ToString('X4')), pressed=$isCurrentlyPressed, justPressed=$wasJustPressed, wasPrev=$wasPreviouslyPressed"
									shortText = " - [DEBUG] Mouse detected"
								}
							)
						}
					}
					
					# Debug: Log mouse button state check (throttled)
					if ($DebugMode -and ($isCurrentlyPressed -or $wasJustPressed)) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Mouse button state: pressed=$isCurrentlyPressed, justPressed=$wasJustPressed, wasPrev=$wasPreviouslyPressed"
									shortText = " - [DEBUG] Mouse state check"
								}
							)
						}
					}
					
					if ($wasJustPressed -or ($isCurrentlyPressed -and -not $wasPreviouslyPressed)) {
						# Debug: Log that we detected a click
						if ($DebugMode) {
							if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
								$LogArray = @()
							}
							$LogArray += [PSCustomObject]@{
								logRow = $true
								components = @(
									@{
										priority = 1
										text = (Get-Date).ToString("HH:mm:ss")
										shortText = (Get-Date).ToString("HH:mm:ss")
									},
									@{
										priority = 2
										text = " - [DEBUG] Mouse click detected! Starting coordinate conversion..."
										shortText = " - [DEBUG] Click detected"
									}
								)
							}
						}
						
						# Left mouse button clicked - check if it's on a dialog button
						$mousePoint = New-Object mJiggAPI.POINT
						$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
						if ($hasGetCursorPos -and [mJiggAPI.Mouse]::GetCursorPos([ref]$mousePoint)) {
							$consoleHandle = [mJiggAPI.Mouse]::GetConsoleWindow()
							if ($consoleHandle -ne [IntPtr]::Zero) {
								# Debug: Log that we got console handle
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Got console handle, screen pos: ($($mousePoint.X),$($mousePoint.Y))"
												shortText = " - [DEBUG] Got handle"
											}
										)
									}
								}
								
								$clientPoint = New-Object mJiggAPI.POINT
								$clientPoint.X = $mousePoint.X
								$clientPoint.Y = $mousePoint.Y
								if ([mJiggAPI.Mouse]::ScreenToClient($consoleHandle, [ref]$clientPoint)) {
									$windowRect = New-Object mJiggAPI.RECT
									if ([mJiggAPI.Mouse]::GetWindowRect($consoleHandle, [ref]$windowRect)) {
										$stdOutHandle = [mJiggAPI.Mouse]::GetStdHandle(-11)
										$bufferInfo = New-Object mJiggAPI.CONSOLE_SCREEN_BUFFER_INFO
										if ([mJiggAPI.Mouse]::GetConsoleScreenBufferInfo($stdOutHandle, [ref]$bufferInfo)) {
											$visibleLeft = $bufferInfo.srWindow.Left
											$visibleTop = $bufferInfo.srWindow.Top
											$visibleRight = $bufferInfo.srWindow.Right
											$visibleBottom = $bufferInfo.srWindow.Bottom
											$visibleWidth = $visibleRight - $visibleLeft + 1
											$visibleHeight = $visibleBottom - $visibleTop + 1
											$windowWidth = $windowRect.Right - $windowRect.Left
											$windowHeight = $windowRect.Bottom - $windowRect.Top
											$borderLeft = 8
											$borderTop = 30
											$borderRight = 8
											$borderBottom = 8
											$clientWidth = $windowWidth - $borderLeft - $borderRight
											$clientHeight = $windowHeight - $borderTop - $borderBottom
											$charWidth = $clientWidth / $visibleWidth
											$charHeight = $clientHeight / $visibleHeight
											$adjustedX = $clientPoint.X - $borderLeft
											$adjustedY = $clientPoint.Y - $borderTop
											$consoleX = [Math]::Floor($adjustedX / $charWidth) + $visibleLeft
											$consoleY = [Math]::Floor($adjustedY / $charHeight) + $visibleTop
											
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												$LogArray += [PSCustomObject]@{
													logRow = $true
													components = @(
														@{
															priority = 1
															text = (Get-Date).ToString("HH:mm:ss")
															shortText = (Get-Date).ToString("HH:mm:ss")
														},
														@{
															priority = 2
															text = " - [DEBUG] Mouse click detected at console ($consoleX,$consoleY)"
															shortText = " - [DEBUG] Click ($consoleX,$consoleY)"
														}
													)
												}
											}
											
											# Debug: Log button bounds for reference
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												$LogArray += [PSCustomObject]@{
													logRow = $true
													components = @(
														@{
															priority = 1
															text = (Get-Date).ToString("HH:mm:ss")
															shortText = (Get-Date).ToString("HH:mm:ss")
														},
														@{
															priority = 2
															text = " - [DEBUG] Button bounds - Row Y: $buttonRowY, Update: X$updateButtonStartX-$updateButtonEndX, Cancel: X$cancelButtonStartX-$cancelButtonEndX"
															shortText = " - [DEBUG] Button bounds"
														}
													)
												}
											}
											
											# Check if click is on update button
											$clickedButton = "none"
											$isOnButton = $false
											if ($consoleY -eq $buttonRowY -or $consoleY -eq ($buttonRowY - 1) -or $consoleY -eq ($buttonRowY + 1)) {
												if ($consoleX -ge $updateButtonStartX -and $consoleX -le $updateButtonEndX) {
													$clickedButton = "Update"
													$isOnButton = $true
													# Update button clicked - trigger update action
													$char = "u"
													$keyProcessed = $true
												} elseif ($consoleX -ge $cancelButtonStartX -and $consoleX -le $cancelButtonEndX) {
													$clickedButton = "Cancel"
													$isOnButton = $true
													# Cancel button clicked - trigger cancel action
													$char = "c"
													$keyProcessed = $true
												}
											}
											
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												if ($isOnButton) {
													$LogArray += [PSCustomObject]@{
														logRow = $true
														components = @(
															@{
																priority = 1
																text = (Get-Date).ToString("HH:mm:ss")
																shortText = (Get-Date).ToString("HH:mm:ss")
															},
															@{
																priority = 2
																text = " - [DEBUG] Button clicked: $clickedButton"
																shortText = " - [DEBUG] $clickedButton"
															}
														)
													}
												} else {
													$LogArray += [PSCustomObject]@{
														logRow = $true
														components = @(
															@{
																priority = 1
																text = (Get-Date).ToString("HH:mm:ss")
																shortText = (Get-Date).ToString("HH:mm:ss")
															},
															@{
																priority = 2
																text = " - [DEBUG] Click NOT on button (row Y: $buttonRowY, update X: $updateButtonStartX-$updateButtonEndX, cancel X: $cancelButtonStartX-$cancelButtonEndX)"
																shortText = " - [DEBUG] Not on button"
															}
														)
													}
												}
											}
										} else {
											# Debug: Log if GetConsoleScreenBufferInfo failed
											if ($DebugMode) {
												if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
													$LogArray = @()
												}
												$LogArray += [PSCustomObject]@{
													logRow = $true
													components = @(
														@{
															priority = 1
															text = (Get-Date).ToString("HH:mm:ss")
															shortText = (Get-Date).ToString("HH:mm:ss")
														},
														@{
															priority = 2
															text = " - [DEBUG] Failed to get console screen buffer info"
															shortText = " - [DEBUG] Buffer info failed"
														}
													)
												}
											}
										}
									} else {
										# Debug: Log if GetWindowRect failed
										if ($DebugMode) {
											if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
												$LogArray = @()
											}
											$LogArray += [PSCustomObject]@{
												logRow = $true
												components = @(
													@{
														priority = 1
														text = (Get-Date).ToString("HH:mm:ss")
														shortText = (Get-Date).ToString("HH:mm:ss")
													},
													@{
														priority = 2
														text = " - [DEBUG] Failed to get window rect"
														shortText = " - [DEBUG] Window rect failed"
													}
												)
											}
										}
									}
								} else {
									# Debug: Log if ScreenToClient failed
									if ($DebugMode) {
										if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
											$LogArray = @()
										}
										$LogArray += [PSCustomObject]@{
											logRow = $true
											components = @(
												@{
													priority = 1
													text = (Get-Date).ToString("HH:mm:ss")
													shortText = (Get-Date).ToString("HH:mm:ss")
												},
												@{
													priority = 2
													text = " - [DEBUG] Failed to convert screen to client coordinates"
													shortText = " - [DEBUG] ScreenToClient failed"
												}
											)
										}
									}
								}
							} else {
								# Debug: Log if console handle is zero
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Console handle is zero"
												shortText = " - [DEBUG] Handle zero"
											}
										)
									}
								}
							}
						} else {
							# Debug: Log if GetCursorPos failed
							if ($DebugMode) {
								if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
									$LogArray = @()
								}
								$LogArray += [PSCustomObject]@{
									logRow = $true
									components = @(
										@{
											priority = 1
											text = (Get-Date).ToString("HH:mm:ss")
											shortText = (Get-Date).ToString("HH:mm:ss")
										},
										@{
											priority = 2
											text = " - [DEBUG] Failed to get cursor position (hasGetCursorPos=$hasGetCursorPos)"
											shortText = " - [DEBUG] GetCursorPos failed"
										}
									)
								}
							}
						}
					}
					# Update previous key state
					$script:previousKeyStates[$leftMouseButtonCode] = $isCurrentlyPressed
				} catch {
					# Ignore errors in mouse click detection
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Mouse click detection error: $($_.Exception.Message)"
									shortText = " - [DEBUG] Error: $($_.Exception.Message)"
								}
							)
						}
					}
				}
				
				# Check for dialog button clicks detected by main loop
				if ($null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null  # Clear it after using
					
					if ($buttonClick -eq "Update") {
						$char = "u"
						$keyProcessed = $true
					} elseif ($buttonClick -eq "Cancel") {
						$char = "c"
						$keyProcessed = $true
					}
				}
				
				# Wait for key input (non-blocking check)
				if (-not $keyProcessed) {
					while ($Host.UI.RawUI.KeyAvailable -and -not $keyProcessed) {
						$keyInfo = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp,AllowCtrlC")
						$isKeyDown = $false
						if ($null -ne $keyInfo.KeyDown) {
							$isKeyDown = $keyInfo.KeyDown
						}
						# Process key-up events (when key is released)
						if (-not $isKeyDown) {
							$key = $keyInfo.Key
							$char = $keyInfo.Character
							$keyProcessed = $true
						}
					}
				}
				
				if (-not $keyProcessed) {
					Start-Sleep -Milliseconds 50
					continue
				}
				
				# Get current field input string reference
				$currentInputRef = switch ($currentField) {
					0 { [ref]$intervalSecondsInput }
					1 { [ref]$intervalVarianceInput }
					2 { [ref]$travelDistanceInput }      # Travel Distance (swapped)
					3 { [ref]$travelVarianceInput }  # Travel Variance (swapped)
					4 { [ref]$moveSpeedInput }           # Move Speed (swapped)
					5 { [ref]$moveVarianceInput }        # Move Variance (swapped)
					6 { [ref]$autoResumeDelaySecondsInput }  # Auto-Resume Delay
				}
				
				if ($char -eq "u" -or $char -eq "U" -or $key -eq "Enter" -or $char -eq [char]13 -or $char -eq [char]10) {
					# Update - validate and save all values
					# Debug: Log movement dialog update
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Movement dialog: Update clicked"
									shortText = " - [DEBUG] Movement: Update"
								}
							)
						}
					}
					$errorMessage = ""
					try {
						$newIntervalSeconds = [double]$intervalSecondsInput
						$newIntervalVariance = [double]$intervalVarianceInput
						$newMoveSpeed = [double]$moveSpeedInput
						$newMoveVariance = [double]$moveVarianceInput
						$newTravelDistance = [double]$travelDistanceInput
						$newTravelVariance = [double]$travelVarianceInput
						$newAutoResumeDelaySeconds = [double]$autoResumeDelaySecondsInput
						
						# Validate ranges
						if ($newIntervalSeconds -le 0) {
							$errorMessage = "Interval must be greater than 0"
						} elseif ($newIntervalVariance -lt 0) {
							$errorMessage = "Interval variance must be >= 0"
						} elseif ($newMoveSpeed -le 0) {
							$errorMessage = "Move speed must be greater than 0"
						} elseif ($newMoveVariance -lt 0) {
							$errorMessage = "Move variance must be >= 0"
						} elseif ($newTravelDistance -le 0) {
							$errorMessage = "Travel distance must be greater than 0"
						} elseif ($newTravelVariance -lt 0) {
							$errorMessage = "Travel variance must be >= 0"
						} elseif ($newAutoResumeDelaySeconds -lt 0) {
							$errorMessage = "Auto-resume delay must be >= 0"
						}
						
						if (-not $errorMessage) {
							$result = @{
								IntervalSeconds = $newIntervalSeconds
								IntervalVariance = $newIntervalVariance
								MoveSpeed = $newMoveSpeed
								MoveVariance = $newMoveVariance
								TravelDistance = $newTravelDistance
								TravelVariance = $newTravelVariance
								AutoResumeDelaySeconds = $newAutoResumeDelaySeconds
							}
							break
						}
					} catch {
						$errorMessage = "Invalid number format"
					}
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
				} elseif ($char -eq "c" -or $char -eq "C" -or $char -eq "t" -or $char -eq "T" -or $key -eq "Escape" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 27)) {
					# Cancel
					# Debug: Log movement dialog cancel
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Movement dialog: Cancel clicked"
									shortText = " - [DEBUG] Movement: Cancel"
								}
							)
						}
					}
					$result = $null
					$needsRedraw = $false
					break
				} elseif ($char -eq "m" -or $char -eq "M") {
					# Hidden option: Close dialog with 'm' key
					$result = $null
					$needsRedraw = $false
					break
				} elseif ($key -eq "Tab" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 9)) {
					# Tab - check if Shift is pressed for reverse tab
					# Use Windows API to check if Shift keys are currently pressed (more reliable than ControlKeyState)
					$VK_LSHIFT = 0xA0  # Left Shift virtual key code
					$VK_RSHIFT = 0xA1  # Right Shift virtual key code
					$isShiftPressed = $false
					try {
						$shiftState = [mJiggAPI.Keyboard]::GetAsyncKeyState($VK_LSHIFT) -bor [mJiggAPI.Keyboard]::GetAsyncKeyState($VK_RSHIFT)
						$isShiftPressed = (($shiftState -band 0x8000) -ne 0)
					} catch {
						# Fallback to ControlKeyState if API call fails
						if ($null -ne $keyInfo.ControlKeyState) {
							$isShiftPressed = (($keyInfo.ControlKeyState -band 4) -ne 0) -or 
											   (($keyInfo.ControlKeyState -band 1) -ne 0) -or 
											   (($keyInfo.ControlKeyState -band 2) -ne 0)
						}
					}
					if ($isShiftPressed) {
						# Shift+Tab - move to previous field
						$currentField = ($currentField - 1 + 7) % 7
					} else {
						# Tab - move to next field
						$currentField = ($currentField + 1) % 7
					}
					$errorMessage = ""
					$lastFieldWithInput = -1  # Reset input tracking when switching fields
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
					# Position cursor at the active field after tab navigation
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldY = $dialogY + $fieldYOffsets[$currentField]
					$currentInputRef = switch ($currentField) {
						0 { [ref]$intervalSecondsInput }
						1 { [ref]$intervalVarianceInput }
						2 { [ref]$travelDistanceInput }
						3 { [ref]$travelVarianceInput }
						4 { [ref]$moveSpeedInput }
						5 { [ref]$moveVarianceInput }
						6 { [ref]$autoResumeDelaySecondsInput }
					}
					$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
					[Console]::SetCursorPosition($cursorX, $fieldY)
				} elseif ($key -eq "Backspace" -or $char -eq [char]8 -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 8)) {
					# Backspace
					# Ensure value is a string (in case it was somehow set as a char)
					$currentInputRef.Value = $currentInputRef.Value.ToString()
					if ($currentInputRef.Value.Length -gt 0) {
						$currentInputRef.Value = $currentInputRef.Value.Substring(0, $currentInputRef.Value.Length - 1)
						# If field is now empty, reset the input tracking so next char will clear again
						if ($currentInputRef.Value.Length -eq 0) {
							$lastFieldWithInput = -1
							[Console]::CursorVisible = $false  # Hide cursor when field is cleared
						}
						$errorMessage = ""
						& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
						# Position cursor at end of input (must be after dialog draw)
						$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
						$fieldY = $dialogY + $fieldYOffsets[$currentField]
						$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
						[Console]::SetCursorPosition($cursorX, $fieldY)
					}
				} elseif ($char -match "[0-9]") {
					# Numeric input - limit to 6 characters
					# If this is the first character typed in this field, clear the field first
					$isFirstChar = ($lastFieldWithInput -ne $currentField)
					if ($isFirstChar) {
						$currentInputRef.Value = $char.ToString()  # Convert char to string
						$lastFieldWithInput = $currentField
						[Console]::CursorVisible = $true  # Show cursor when first character is typed
					} elseif ($currentInputRef.Value.Length -lt 6) {
						$currentInputRef.Value += $char.ToString()  # Convert char to string
					}
					$errorMessage = ""
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
					# Position cursor at end of input (must be after dialog draw)
					# Calculate Y position based on field index (matches drawing positions)
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldY = $dialogY + $fieldYOffsets[$currentField]
					# Cursor X: dialogX + inputBoxStartX + 1 (for opening bracket) + length of actual value
					# inputBoxStartX is relative to dialogX, so we need to add dialogX for absolute position
					$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
					# Position cursor - use Console method directly for reliability
					[Console]::SetCursorPosition($cursorX, $fieldY)
				} elseif ($char -eq ".") {
					# Decimal point for all fields - limit to 6 characters (including decimal point)
					# If this is the first character typed in this field, clear the field first
					$isFirstChar = ($lastFieldWithInput -ne $currentField)
					if ($isFirstChar) {
						$currentInputRef.Value = "."  # Already a string
						$lastFieldWithInput = $currentField
						[Console]::CursorVisible = $true  # Show cursor when first character is typed
					} elseif ($currentInputRef.Value -notmatch "\." -and $currentInputRef.Value.Length -lt 6) {
						$currentInputRef.Value += "."  # Already a string
					}
					$errorMessage = ""
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
					# Position cursor at end of input after clearing (must be after dialog draw)
					if ($isFirstChar) {
						# Calculate Y position based on field index (matches drawing positions)
						$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
						$fieldY = $dialogY + $fieldYOffsets[$currentField]
						# Cursor X: dialogX + inputBoxStartX + 1 (for opening bracket) + length of actual value
						# inputBoxStartX is relative to dialogX, so we need to add dialogX for absolute position
						$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
						# Position cursor - use Console method directly for reliability
						[Console]::SetCursorPosition($cursorX, $fieldY)
					}
				} elseif ($key -eq "UpArrow" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 38)) {
					# UpArrow - move to previous field (reverse tab)
					$currentField = ($currentField - 1 + 7) % 7
					$errorMessage = ""
					$lastFieldWithInput = -1  # Reset input tracking when switching fields
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
					# Position cursor at the active field after arrow navigation
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldY = $dialogY + $fieldYOffsets[$currentField]
					$currentInputRef = switch ($currentField) {
						0 { [ref]$intervalSecondsInput }
						1 { [ref]$intervalVarianceInput }
						2 { [ref]$travelDistanceInput }
						3 { [ref]$travelVarianceInput }
						4 { [ref]$moveSpeedInput }
						5 { [ref]$moveVarianceInput }
						6 { [ref]$autoResumeDelaySecondsInput }
					}
					$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
					[Console]::SetCursorPosition($cursorX, $fieldY)
				} elseif ($key -eq "DownArrow" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 40)) {
					# DownArrow - move to next field (forward tab)
					$currentField = ($currentField + 1) % 7
					$errorMessage = ""
					$lastFieldWithInput = -1  # Reset input tracking when switching fields
					& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
					# Position cursor at the active field after arrow navigation
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldY = $dialogY + $fieldYOffsets[$currentField]
					$currentInputRef = switch ($currentField) {
						0 { [ref]$intervalSecondsInput }
						1 { [ref]$intervalVarianceInput }
						2 { [ref]$travelDistanceInput }
						3 { [ref]$travelVarianceInput }
						4 { [ref]$moveSpeedInput }
						5 { [ref]$moveVarianceInput }
						6 { [ref]$autoResumeDelaySecondsInput }
					}
					$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
					[Console]::SetCursorPosition($cursorX, $fieldY)
				}
				
				# Clear any remaining keys in buffer
				try {
					while ($Host.UI.RawUI.KeyAvailable) {
						$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
					}
				} catch {
					# Silently ignore
				}
			} until ($false)
			
			# Clear shadow before clearing dialog area
			Clear-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
			
			# Clear dialog area
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX, $dialogY + $i)
				Write-Host (" " * $dialogWidth) -NoNewline
			}
			
			# Restore cursor visibility
			[Console]::CursorVisible = $savedCursorVisible
			
			# Return result object
			return @{
				Result = $result
				NeedsRedraw = $needsRedraw
			}
		}

		# Function to show quit confirmation dialog
		function Show-QuitConfirmationDialog {
			param(
				[ref]$HostWidthRef,
				[ref]$HostHeightRef
			)
			
			# Debug: Log that dialog was opened
			if ($DebugMode) {
				# Note: LogArray is in parent scope, accessible directly
				if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
					$LogArray = @()
				}
				$LogArray += [PSCustomObject]@{
					logRow = $true
					components = @(
						@{
							priority = 1
							text = (Get-Date).ToString("HH:mm:ss")
							shortText = (Get-Date).ToString("HH:mm:ss")
						},
						@{
							priority = 2
							text = " - [DEBUG] Quit confirmation dialog opened"
							shortText = " - [DEBUG] Quit dialog opened"
						}
					)
				}
			}
			
			# Get current host dimensions from references
			$currentHostWidth = $HostWidthRef.Value
			$currentHostHeight = $HostHeightRef.Value
			
			# Dialog dimensions (same as time change dialog)
			$dialogWidth = 35
			$dialogHeight = 8
			$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
			$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
			
			# Save current cursor position and visibility
			$savedCursorVisible = [Console]::CursorVisible
			[Console]::CursorVisible = $false
			
			# Draw dialog box (exactly 35 characters per line)
			# Calculate spacing for bottom line (emojis display as 2 chars each but count as 1 in string length)
			$checkmark = [char]0x2705  # ✅ green checkmark
			$redX = [char]0x274C  # ❌ red X
			$bottomLineContent = "│  " + $checkmark + "|(y)es  " + $redX + "|(n)o"
			# Account for emojis: each emoji is 1 char in string but 2 display columns
			# String length: "│  " (3) + emoji (1) + "|" (1) + "(y)es  " (7) + emoji (1) + "|" (1) + "(n)o" (4) = 18
			# Display width: "│  " (3) + emoji (2) + "|" (1) + "(y)es  " (7) + emoji (2) + "|" (1) + "(n)o" (4) = 20
			# So we need: 35 - 20 - 1 = 14 spaces before closing │
			$bottomLineTextLength = $bottomLineContent.Length + 2  # +2 because emojis count as 2 display chars but 1 string char each
			$bottomLinePadding = Get-Padding -usedWidth ($bottomLineTextLength + 1) -totalWidth $dialogWidth
			
			# Build all lines to be exactly 35 characters using Get-Padding helper
			$line0 = "┌─────────────────────────────────┐"  # 35 chars
			$line1Text = "│  Confirm Quit"
			$line1Padding = Get-Padding -usedWidth ($line1Text.Length + 1) -totalWidth $dialogWidth
			$line1 = $line1Text + (" " * $line1Padding) + "│"
			
			$line2 = "│" + (" " * 33) + "│"  # 35 chars
			
			$line3Text = "│  Are you sure you want to quit?"
			$line3Padding = Get-Padding -usedWidth ($line3Text.Length + 1) -totalWidth $dialogWidth
			$line3 = $line3Text + (" " * $line3Padding) + "│"
			
			$line4 = "│" + (" " * 33) + "│"  # 35 chars
			$line5 = "│" + (" " * 33) + "│"  # 35 chars
			
			$dialogLines = @(
				$line0,
				$line1,
				$line2,
				$line3,
				$line4,
				$line5,
				$null,  # Bottom line will be written separately with colors
				"└─────────────────────────────────┘"  # 35 chars
			)
			
			# Draw dialog background (clear area)
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX, $dialogY + $i)
				Write-Host (" " * $dialogWidth) -NoNewline
			}
			
			# Draw dialog box
			for ($i = 0; $i -lt $dialogLines.Count; $i++) {
				if ($i -eq 1) {
					# Title line - write in magenta
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host "│  " -NoNewline
					Write-Host "Confirm Quit" -NoNewline -ForegroundColor Magenta
					$titleUsedWidth = 3 + "Confirm Quit".Length  # "│  " + title
					$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
					Write-Host (" " * $titlePadding) -NoNewline
					Write-Host "│" -NoNewline
				} elseif ($i -eq 6) {
					# Bottom line - write with colored icons and hotkey letters
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host "│  " -NoNewline
					Write-Host $checkmark -NoNewline -ForegroundColor Green
					Write-Host "|" -NoNewline
					# Parse "(y)es" - parentheses white, letter yellow, text white
					Write-Host "(" -NoNewline
					Write-Host "y" -NoNewline -ForegroundColor Yellow
					Write-Host ")es  " -NoNewline
					Write-Host $redX -NoNewline -ForegroundColor Red
					Write-Host "|" -NoNewline
					# Parse "(n)o" - parentheses white, letter yellow, text white
					Write-Host "(" -NoNewline
					Write-Host "n" -NoNewline -ForegroundColor Yellow
					Write-Host ")o" -NoNewline
					Write-Host (" " * $bottomLinePadding) -NoNewline
					Write-Host "│" -NoNewline
				} else {
					[Console]::SetCursorPosition($dialogX, $dialogY + $i)
					Write-Host $dialogLines[$i] -NoNewline
				}
			}
			
			# Draw drop shadow
			Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
			
			# Get input
			$result = $null
			$needsRedraw = $false
			
			:inputLoop do {
				# Check for window resize and update references
				$pshost = Get-Host
				$pswindow = $pshost.UI.RawUI
				$newWindowSize = $pswindow.WindowSize
				if ($newWindowSize.Width -ne $currentHostWidth -or $newWindowSize.Height -ne $currentHostHeight) {
					# Window was resized - update references and flag for main UI redraw
					# Don't force buffer size - let PowerShell manage it (allows text zoom to work)
					$HostWidthRef.Value = $newWindowSize.Width
					$HostHeightRef.Value = $newWindowSize.Height
					$currentHostWidth = $newWindowSize.Width
					$currentHostHeight = $newWindowSize.Height
					$needsRedraw = $true
					
					# Reposition dialog
					$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
					$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
					
					# Clear screen and redraw dialog
					Clear-Host
					for ($i = 0; $i -lt $dialogLines.Count; $i++) {
						if ($i -eq 1) {
							# Title line - write in magenta
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host "│  " -NoNewline
							Write-Host "Confirm Quit" -NoNewline -ForegroundColor Magenta
							$titleUsedWidth = 3 + "Confirm Quit".Length  # "│  " + title
							$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
							Write-Host (" " * $titlePadding) -NoNewline
							Write-Host "│" -NoNewline
						} elseif ($i -eq 6) {
							# Bottom line - write with colored icons and hotkey letters
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host "│  " -NoNewline
							Write-Host $checkmark -NoNewline -ForegroundColor Green
							Write-Host "|" -NoNewline
							# Parse "(y)es" - parentheses white, letter yellow, text white
							Write-Host "(" -NoNewline
							Write-Host "y" -NoNewline -ForegroundColor Yellow
							Write-Host ")es  " -NoNewline
							Write-Host $redX -NoNewline -ForegroundColor Red
							Write-Host "|" -NoNewline
							# Parse "(n)o" - parentheses white, letter yellow, text white
							Write-Host "(" -NoNewline
							Write-Host "n" -NoNewline -ForegroundColor Yellow
							Write-Host ")o" -NoNewline
							Write-Host (" " * $bottomLinePadding) -NoNewline
							Write-Host "│" -NoNewline
						} else {
							[Console]::SetCursorPosition($dialogX, $dialogY + $i)
							Write-Host $dialogLines[$i] -NoNewline
						}
					}
					
					# Draw drop shadow
					Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
				}
				
				# Check for dialog button clicks detected by main loop
				if ($null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null  # Clear it after using
					
					if ($buttonClick -eq "Update") {
						$char = "u"
						$keyProcessed = $true
					} elseif ($buttonClick -eq "Cancel") {
						$char = "c"
						$keyProcessed = $true
					}
				}
				
				# Wait for key input (non-blocking check)
				# Read keys and only process key UP events (same as main menu)
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				while ($Host.UI.RawUI.KeyAvailable -and -not $keyProcessed) {
					$keyInfo = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyup,AllowCtrlC")
					$isKeyDown = $false
					if ($null -ne $keyInfo.KeyDown) {
						$isKeyDown = $keyInfo.KeyDown
					}
					
					# Only process key UP events (skip key down)
					if (-not $isKeyDown) {
						$key = $keyInfo.Key
						$char = $keyInfo.Character
						$keyProcessed = $true
					}
				}
				
				if (-not $keyProcessed) {
					# No key available, sleep briefly and check for resize again
					Start-Sleep -Milliseconds 50
					continue
				}
				
				if ($char -eq "y" -or $char -eq "Y" -or $key -eq "Enter" -or $char -eq [char]13 -or $char -eq [char]10) {
					# Yes - confirm quit (Enter key also works as hidden function)
					# Debug: Log quit confirmation
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Quit dialog: Confirmed"
									shortText = " - [DEBUG] Quit: Yes"
								}
							)
						}
					}
					$result = $true
					break
				} elseif ($char -eq "n" -or $char -eq "N" -or $char -eq "q" -or $char -eq "Q" -or $key -eq "Escape" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 27)) {
					# No - cancel quit (Escape key and 'q' key also work as hidden functions)
					# Debug: Log quit cancellation
					if ($DebugMode) {
						if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
							$LogArray = @()
						}
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(
								@{
									priority = 1
									text = (Get-Date).ToString("HH:mm:ss")
									shortText = (Get-Date).ToString("HH:mm:ss")
								},
								@{
									priority = 2
									text = " - [DEBUG] Quit dialog: Canceled"
									shortText = " - [DEBUG] Quit: No"
								}
							)
						}
					}
					$result = $false
					$needsRedraw = $false  # No redraw needed on cancel
					break
				}
				
				# Clear any remaining keys in buffer after processing
				try {
					while ($Host.UI.RawUI.KeyAvailable) {
						$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
					}
				} catch {
					# Silently ignore - buffer might not be clearable
				}
			} until ($false)
			
			# Clear shadow before clearing dialog area
			Clear-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
			
			# Clear dialog area
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				[Console]::SetCursorPosition($dialogX, $dialogY + $i)
				Write-Host (" " * $dialogWidth) -NoNewline
			}
			
			# Restore cursor visibility
			[Console]::CursorVisible = $savedCursorVisible
			
			# Clear dialog button bounds when dialog closes
			$script:DialogButtonBounds = $null
			$script:DialogButtonClick = $null
			
			# Return result object with result and redraw flag
			return @{
				Result = $result
				NeedsRedraw = $needsRedraw
			}
		}

		# Initialization complete - pause to read debug output if in debug mode
		if ($DebugMode) {
			Write-Host "`nPress any key to start mJig..." -ForegroundColor Yellow
			
			$keyPressed = $false
			
			# VK codes for Ctrl (to filter out Ctrl-only keys)
			$VK_LCONTROL = 0xA2
			$VK_RCONTROL = 0xA3
			
			while (-not $keyPressed) {
				# Process Windows messages to ensure smooth mouse movement
				try {
					if ([mJiggAPI.MouseHook]::hHook -ne [IntPtr]::Zero) {
						[mJiggAPI.MouseHook]::ProcessMessages()
					}
				} catch {
					# Ignore errors - hook might not be installed
				}
				
				# Check Ctrl state - if Ctrl is pressed alone, skip ALL key processing
				# This prevents Ctrl from triggering continuation
				$ctrlPressedAlone = $false
				try {
					$ctrlState = [mJiggAPI.Keyboard]::GetAsyncKeyState($VK_LCONTROL) -bor [mJiggAPI.Keyboard]::GetAsyncKeyState($VK_RCONTROL)
					$ctrlPressed = (($ctrlState -band 0x8000) -ne 0)
					if ($ctrlPressed) {
						$ctrlPressedAlone = $true
					}
				} catch {
					# Ignore errors
				}
				
				# If Ctrl is pressed alone, skip ALL key reading and continue loop
				if ($ctrlPressedAlone) {
					Start-Sleep -Milliseconds 50
					continue
				}
				
				# Use ReadKey but filter Ctrl-only keys
				# Consume Ctrl-only keys until we get a valid key
				try {
					# Check if any key is available
					if ($Host.UI.RawUI.KeyAvailable) {
						# Read keys until we get a non-Ctrl-only key
						$validKeyFound = $false
						while ($Host.UI.RawUI.KeyAvailable -and -not $validKeyFound) {
							$keyInfo = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
							# Only process key DOWN events
							$isKeyDown = if ($null -ne $keyInfo.KeyDown) { $keyInfo.KeyDown } else { $true }
							# Filter out Ctrl-only keys
							$isCtrlOnly = ($keyInfo.Key -eq "LeftCtrl" -or $keyInfo.Key -eq "RightCtrl")
							
							if ($isKeyDown -and -not $isCtrlOnly) {
								$keyPressed = $true
								$validKeyFound = $true
							}
							# If it was Ctrl-only or key-up, continue reading next key
						}
					} else {
						# No key available, sleep briefly
						Start-Sleep -Milliseconds 50
					}
				} catch {
					# Ignore read errors
					Start-Sleep -Milliseconds 50
				}
			}
			# Don't clear the debug output - let user read it
			# Clear-Host is called later when the main loop starts rendering
		}

		# Clear any keys that might be in the input buffer from before script started
		# This prevents the script from getting stuck on startup
		try {
			while ($Host.UI.RawUI.KeyAvailable) {
				$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
			}
		} catch {
			# Silently ignore - buffer might not be in a readable state
		}

		# Main Processing Loop
		:process do {
			
			# Reset state for this iteration
			$userInputDetected = $false
			$keyboardInputDetected = $false  # Track keyboard input separately
			$mouseInputDetected = $false     # Track mouse input separately
			$waitExecuted = $false
			$intervalKeys = @()  # Track keys pressed during this interval
			$intervalMouseInputs = @()  # Track mouse inputs (clicks and movement) separately
			$interval = 0
			$math = 0
			$date = Get-Date
			$currentTime = Get-Date -Format "HHmm"
			$forceRedraw = $false
			$automatedMovementPos = $null  # Track position after automated movement
			$directionArrow = ""  # Track direction arrow for log display
			$lastKeyPress = $null  # Reset key press tracking
			$lastKeyInfo = $null  # Reset key info tracking
			$pressedMenuKeys = @{}  # Track which menu keys are currently pressed (to detect key up)
			
			# Calculate interval and wait BEFORE doing movement (skip on first run or if forceRedraw)
			if ($null -ne $LastMovementTime -and -not $forceRedraw) {
				# Calculate random interval with variance
				# Variance can be any number, even larger than base interval
				# Convert to milliseconds for calculation, then back to seconds for display
				$intervalSecondsMs = $script:IntervalSeconds * 1000
				$intervalVarianceMs = $script:IntervalVariance * 1000
				
				# Get-Random -Maximum returns 0 to (max-1), so we need to add 1 to get 0 to max
				$varianceAmountMs = Get-Random -Minimum 0 -Maximum ($intervalVarianceMs + 1)
				$ras = Get-Random -Maximum 2 -Minimum 0
				if ($ras -eq 0) {
					# Subtract variance
					$intervalMs = ($intervalSecondsMs - $varianceAmountMs)
				} else {
					# Add variance
					$intervalMs = ($intervalSecondsMs + $varianceAmountMs) 
				}
				
				# Subtract the previous movement duration from the interval
				$intervalMs = $intervalMs - $LastMovementDurationMs
				
				# Ensure minimum interval of 1 second (variance can be larger than base interval)
				$minIntervalMs = 1000  # 1 second in milliseconds
				if ($intervalMs -lt $minIntervalMs) {
					$intervalMs = $minIntervalMs
				}
				
				# Convert back to seconds and round to 1 decimal place for display
				$interval = [math]::Round($intervalMs / 1000, 1)
				
				# Calculate number of 50ms iterations needed (1000ms / 50ms = 20 iterations per second)
				# Use the millisecond value for accurate calculation
				$math = [math]::Max(1, [math]::Floor($intervalMs / 50))
				
				$waitExecuted = $true
				# Use direct Windows API call for better performance (avoids .NET stutter)
				$point = New-Object mJiggAPI.POINT
				$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
				if ($hasGetCursorPos) {
					if ([mJiggAPI.Mouse]::GetCursorPos([ref]$point)) {
						# Convert POINT to System.Drawing.Point for compatibility
						$mousePosAtStart = New-Object System.Drawing.Point($point.X, $point.Y)
					} else {
						# API call failed - skip this check
						$mousePosAtStart = $null
					}
				} else {
					# Method not available - skip this check
					$mousePosAtStart = $null
				}
				
				# Wait Loop - Check window/buffer size changes, menu hotkeys, and keyboard input
				# Menu hotkeys checked every 200ms (every 4th iteration), keyboard input checked every 50ms for maximum reliability
				$x = 0
				:waitLoop do {
					$x++
					
					# Check for system-wide keyboard input every 50ms for maximum reliability
					# Skip checking if we recently sent a simulated key press (within last 300ms)
					$shouldCheckKeyboard = $true
					if ($null -ne $LastSimulatedKeyPress) {
						$timeSinceSimulatedKey = ((Get-Date) - $LastSimulatedKeyPress).TotalMilliseconds
						if ($timeSinceSimulatedKey -lt 300) {
							$shouldCheckKeyboard = $false
						} else {
							$LastSimulatedKeyPress = $null
						}
					}
					
					if ($shouldCheckKeyboard) {
						# Initialize previous key states lazily
						if ($null -eq $script:previousKeyStates) {
							$script:previousKeyStates = @{}
						}
						
						# Check mouse position every 50ms to detect movement for console skip
						# This prevents console updates from blocking during active mouse movement
						if ($null -eq $script:lastMousePosCheck) {
							$script:lastMousePosCheck = $null
						}
						try {
							$checkPoint = New-Object mJiggAPI.POINT
							$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
							if ($hasGetCursorPos -and [mJiggAPI.Mouse]::GetCursorPos([ref]$checkPoint)) {
								$currentCheckPos = New-Object System.Drawing.Point($checkPoint.X, $checkPoint.Y)
								if ($null -ne $script:lastMousePosCheck) {
									# Check if mouse moved more than 2 pixels
									$deltaX = [Math]::Abs($currentCheckPos.X - $script:lastMousePosCheck.X)
									$deltaY = [Math]::Abs($currentCheckPos.Y - $script:lastMousePosCheck.Y)
									if ($deltaX -gt 2 -or $deltaY -gt 2) {
										# Mouse is moving - set flag to skip console updates
										$script:LastMouseMovementTime = Get-Date
										$mouseInputDetected = $true
										if ($script:AutoResumeDelaySeconds -gt 0) {
											$LastUserInputTime = Get-Date
										}
									}
								}
								$script:lastMousePosCheck = $currentCheckPos
							}
						} catch {
							# Ignore errors in mouse position checking
						}
						
						# Track all pressed keys for display (only when Output is "full")
						if ($Output -eq "full") {
							for ($keyCode = 0; $keyCode -le 255; $keyCode++) {
								if ($keyCode -eq 0xA5) { continue }  # Skip Right Alt
								$currentKeyState = [mJiggAPI.Keyboard]::GetAsyncKeyState($keyCode)
								$isCurrentlyPressed = (($currentKeyState -band 0x8000) -ne 0)
								if ($isCurrentlyPressed) {
									$keyName = Get-KeyName -keyCode $keyCode
									if ($keyName) {
										$PressedKeys[$keyCode] = $keyName
									} else {
										$PressedKeys[$keyCode] = "Unknown(0x$($keyCode.ToString('X2')))"
									}
								} else {
									if ($PressedKeys.ContainsKey($keyCode)) {
										$PressedKeys.Remove($keyCode)
									}
								}
							}
						}
						
						# Check for key press transitions (for user input detection)
						for ($keyCode = 0; $keyCode -le 255; $keyCode++) {
							if ($keyCode -eq 0xA5) { continue }  # Skip Right Alt
							
							# Check mouse buttons (only actual mouse button codes: 0x01-0x06)
							# 0x01 = LButton, 0x02 = RButton, 0x04 = MButton, 0x05 = XButton1, 0x06 = XButton2
							# Note: 0x07-0x0F are NOT mouse buttons (they're keys like Backspace, Tab, etc.)
							if ($keyCode -ge 0x01 -and $keyCode -le 0x06) {
								$currentKeyState = [mJiggAPI.Keyboard]::GetAsyncKeyState($keyCode)
								$isCurrentlyPressed = (($currentKeyState -band 0x8000) -ne 0)
								$wasJustPressed = (($currentKeyState -band 0x0001) -ne 0)
								$wasPreviouslyPressed = if ($script:previousKeyStates.ContainsKey($keyCode)) { $script:previousKeyStates[$keyCode] } else { $false }
								
								if ($wasJustPressed -or ($isCurrentlyPressed -and -not $wasPreviouslyPressed)) {
									# Left mouse button clicked - check if it's on a menu item or dialog button
									# Debounce: Only log clicks if we haven't logged one in the last 300ms
									# Also check if mouse is currently moving - if so, skip logging entirely
									$shouldLogClick = $true
									if ($DebugMode) {
										# Check debounce timer
										if ($null -ne $script:LastClickLogTime) {
											$timeSinceLastClickLog = ((Get-Date) - $script:LastClickLogTime).TotalMilliseconds
											if ($timeSinceLastClickLog -lt 300) {
												$shouldLogClick = $false
											}
										}
										
										# Check if mouse is currently moving - if so, don't log at all
										if ($null -ne $script:LastMouseMovementTime) {
											$timeSinceMouseMovement = ((Get-Date) - $script:LastMouseMovementTime).TotalMilliseconds
											# Don't log if mouse moved within last 200ms
											if ($timeSinceMouseMovement -lt 200) {
												$shouldLogClick = $false
											}
										}
									}
									
									if ($shouldLogClick) {
										# Collect all debug information first, then log it all at once
										$debugInfo = @{
											focus = "unknown"
											screenX = "unknown"
											screenY = "unknown"
											consoleX = "unknown"
											consoleY = "unknown"
											button = "none"
										}
										
										if ($DebugMode) {
											try {
												# Get mouse position in screen coordinates
												$gotCursorPos = $false
												$mousePoint = New-Object mJiggAPI.POINT
												$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
												if ($hasGetCursorPos) {
													$gotCursorPos = [mJiggAPI.Mouse]::GetCursorPos([ref]$mousePoint)
													if ($gotCursorPos) {
														$debugInfo.screenX = $mousePoint.X
														$debugInfo.screenY = $mousePoint.Y
													}
												}
												
												# Get console window handle
												$consoleHandle = [IntPtr]::Zero
												$hasGetConsoleWindow = [mJiggAPI.Mouse].GetMethod("GetConsoleWindow") -ne $null
												if ($hasGetConsoleWindow) {
													try {
														$consoleHandle = [mJiggAPI.Mouse]::GetConsoleWindow()
													} catch {
														# Method exists but call failed
														$consoleHandle = [IntPtr]::Zero
													}
												}
												
												if ($consoleHandle -eq [IntPtr]::Zero) {
													# Console handle is zero - this can happen in Windows Terminal or other modern terminal emulators
													# Try alternative method to get console window
													try {
														# Try using Get-Process to find the console window
														$currentProcess = Get-Process -Id $PID
														$mainWindowHandle = $currentProcess.MainWindowHandle
														if ($mainWindowHandle -ne [IntPtr]::Zero) {
															$consoleHandle = $mainWindowHandle
														}
													} catch {
														# Alternative method also failed
													}
												}
												
												if ($consoleHandle -ne [IntPtr]::Zero) {
													# Check if console window is in focus
													$isConsoleFocused = $false
													try {
														$foregroundWindow = [mJiggAPI.Mouse]::GetForegroundWindow()
														
														# First check: Does foreground window match our console handle?
														$isConsoleFocused = ($foregroundWindow -eq $consoleHandle)
														
														# If not, check if foreground window belongs to our process or parent process
														if (-not $isConsoleFocused) {
															$hasGetWindowThreadProcessId = [mJiggAPI.Mouse].GetMethod("GetWindowThreadProcessId") -ne $null
															if ($hasGetWindowThreadProcessId) {
																$fgProcessId = 0
																[mJiggAPI.Mouse]::GetWindowThreadProcessId($foregroundWindow, [ref]$fgProcessId) | Out-Null
																
																# Check if foreground window belongs to our process
																if ($fgProcessId -eq $PID) {
																	# Foreground window belongs to our process - use it as the console handle
																	$consoleHandle = $foregroundWindow
																	$isConsoleFocused = $true
																} else {
																	# Check if foreground window belongs to parent process (Windows Terminal)
																	try {
																		$currentProcess = Get-Process -Id $PID -ErrorAction SilentlyContinue
																		if ($null -ne $currentProcess -and $null -ne $currentProcess.Parent -and $fgProcessId -eq $currentProcess.Parent.Id) {
																			# Foreground window belongs to parent - use it as the console handle
																			$consoleHandle = $foregroundWindow
																			$isConsoleFocused = $true
																		}
																	} catch {
																		# Parent check failed
																	}
																}
																
																if ($DebugMode) {
																	# Add debug info showing process IDs
																	$fgProcessName = "Unknown"
																	try {
																		$fgProcess = Get-Process -Id $fgProcessId -ErrorAction SilentlyContinue
																		if ($null -ne $fgProcess) {
																			$fgProcessName = $fgProcess.ProcessName
																		}
																	} catch { }
																	
																	$debugInfo.focus = if ($isConsoleFocused) { 
																		"yes (fg=$($foregroundWindow.ToInt64()), console=$($consoleHandle.ToInt64()), fgPID=$fgProcessId)" 
																	} else { 
																		"no (fg=$($foregroundWindow.ToInt64()), console=$($consoleHandle.ToInt64()), fgPID=$fgProcessId($fgProcessName))" 
																	}
																} else {
																	$debugInfo.focus = if ($isConsoleFocused) { "yes" } else { "no" }
																}
															} else {
																# Can't check process ID
																if ($DebugMode) {
																	$debugInfo.focus = "no (fg=$($foregroundWindow.ToInt64()), console=$($consoleHandle.ToInt64()), can't check PID)"
																} else {
																	$debugInfo.focus = "no"
																}
															}
														} else {
															# Handles match
															if ($DebugMode) {
																$debugInfo.focus = "yes (fg=$($foregroundWindow.ToInt64()), console=$($consoleHandle.ToInt64()))"
															} else {
																$debugInfo.focus = "yes"
															}
														}
													} catch {
														$isConsoleFocused = $true  # Fallback: assume focused
														$debugInfo.focus = "yes (fallback: $($_.Exception.Message))"
													}
													
													# Convert screen coordinates to client (console window) coordinates
													if ($gotCursorPos) {
														$clientPoint = New-Object mJiggAPI.POINT
														$clientPoint.X = $mousePoint.X
														$clientPoint.Y = $mousePoint.Y
														$hasScreenToClient = [mJiggAPI.Mouse].GetMethod("ScreenToClient") -ne $null
														if ($hasScreenToClient) {
															$screenToClientOk = [mJiggAPI.Mouse]::ScreenToClient($consoleHandle, [ref]$clientPoint)
															if ($screenToClientOk) {
																# Get console window rect to calculate character cell size
																$windowRect = New-Object mJiggAPI.RECT
																$hasGetWindowRect = [mJiggAPI.Mouse].GetMethod("GetWindowRect") -ne $null
																if ($hasGetWindowRect) {
																	$getWindowRectOk = [mJiggAPI.Mouse]::GetWindowRect($consoleHandle, [ref]$windowRect)
																	if ($getWindowRectOk) {
																		# Get console buffer info for accurate coordinate conversion
																		$stdOutHandle = [IntPtr]::Zero
																		$hasGetStdHandle = [mJiggAPI.Mouse].GetMethod("GetStdHandle") -ne $null
																		if ($hasGetStdHandle) {
																			$stdOutHandle = [mJiggAPI.Mouse]::GetStdHandle(-11)  # STD_OUTPUT_HANDLE
																		}
																		$bufferInfo = New-Object mJiggAPI.CONSOLE_SCREEN_BUFFER_INFO
																		$hasGetBufferInfo = [mJiggAPI.Mouse].GetMethod("GetConsoleScreenBufferInfo") -ne $null
																		if ($hasGetBufferInfo -and $stdOutHandle -ne [IntPtr]::Zero) {
																			$getBufferInfoOk = [mJiggAPI.Mouse]::GetConsoleScreenBufferInfo($stdOutHandle, [ref]$bufferInfo)
																			if ($getBufferInfoOk) {
																				# Use buffer info to get visible window area
																				$visibleLeft = $bufferInfo.srWindow.Left
																				$visibleTop = $bufferInfo.srWindow.Top
																				$visibleRight = $bufferInfo.srWindow.Right
																				$visibleBottom = $bufferInfo.srWindow.Bottom
																				
																				# Calculate character cell size from window dimensions and visible buffer area
																				$visibleWidth = $visibleRight - $visibleLeft + 1
																				$visibleHeight = $visibleBottom - $visibleTop + 1
																				$windowWidth = $windowRect.Right - $windowRect.Left
																				$windowHeight = $windowRect.Bottom - $windowRect.Top
																				
																				# Account for window borders (approximate: 8px left/right, 30px top for title bar, 8px bottom)
																				$borderLeft = 8
																				$borderTop = 30
																				$borderRight = 8
																				$borderBottom = 8
																				$clientWidth = $windowWidth - $borderLeft - $borderRight
																				$clientHeight = $windowHeight - $borderTop - $borderBottom
																				
																				# Calculate character cell size
																				$charWidth = $clientWidth / $visibleWidth
																				$charHeight = $clientHeight / $visibleHeight
																				
																				# Convert client coordinates to console buffer coordinates
																				# Account for borders
																				$adjustedX = $clientPoint.X - $borderLeft
																				$adjustedY = $clientPoint.Y - $borderTop
																				$consoleX = [Math]::Floor($adjustedX / $charWidth) + $visibleLeft
																				$consoleY = [Math]::Floor($adjustedY / $charHeight) + $visibleTop
																				
																				$debugInfo.consoleX = $consoleX
																				$debugInfo.consoleY = $consoleY
																				
																				# Check if a dialog is open and if click is on a dialog button
																				if ($null -ne $script:DialogButtonBounds -and $isConsoleFocused) {
																					$bounds = $script:DialogButtonBounds
																					
																					# Check if click Y matches button row (with tolerance)
																					if ($consoleY -eq $bounds.buttonRowY -or $consoleY -eq ($bounds.buttonRowY - 1) -or $consoleY -eq ($bounds.buttonRowY + 1)) {
																						# Check if click is on Update button
																						if ($consoleX -ge $bounds.updateStartX -and $consoleX -le $bounds.updateEndX) {
																							$script:DialogButtonClick = "Update"
																							$debugInfo.button = "Update"
																						}
																						# Check if click is on Cancel button
																						elseif ($consoleX -ge $bounds.cancelStartX -and $consoleX -le $bounds.cancelEndX) {
																							$script:DialogButtonClick = "Cancel"
																							$debugInfo.button = "Cancel"
																						}
																					}
																				}
																				
																				# Check if click is on any menu item (with tolerance for Y coordinate)
																				# Only check menu items if dialog is not open
																				if ($null -eq $script:DialogButtonBounds -and $Output -eq "full" -and $null -ne $script:MenuItemsBounds -and $script:MenuItemsBounds.Count -gt 0) {
																					# Check if click is on any menu item (with tolerance for Y coordinate)
																					# Also try simpler coordinate matching as fallback
																					$clickMatched = $false
																					foreach ($menuItem in $script:MenuItemsBounds) {
																						# Allow Y coordinate to match exactly or be within 1 row (for click detection tolerance)
																						$yMatches = ($consoleY -eq $menuItem.y -or $consoleY -eq ($menuItem.y - 1) -or $consoleY -eq ($menuItem.y + 1))
																						$xMatches = ($consoleX -ge $menuItem.startX -and $consoleX -le $menuItem.endX)
																						
																						if ($yMatches -and $xMatches) {
																							# Click matches this menu item - trigger the hotkey action
																							if ($null -ne $menuItem.hotkey) {
																								$script:MenuClickHotkey = $menuItem.hotkey
																								$clickMatched = $true
																								break
																							}
																						}
																					}
																					
																					# Fallback: if precise matching failed, try simpler approach
																					# Check if click Y is in menu row area, then find closest menu item by X
																					if (-not $clickMatched -and $script:MenuItemsBounds.Count -gt 0) {
																						# Get the menu row Y coordinate (should be same for all items)
																						$menuRowY = $script:MenuItemsBounds[0].y
																						# Check if click Y is close to menu row (within 3 rows for tolerance)
																						if ([Math]::Abs($consoleY - $menuRowY) -le 3) {
																							# Find the menu item whose X range the click is closest to
																							$closestItem = $null
																							$closestDistance = [int]::MaxValue
																							foreach ($menuItem in $script:MenuItemsBounds) {
																								# Calculate distance from click X to menu item center
																								$itemCenterX = ($menuItem.startX + $menuItem.endX) / 2
																								$distance = [Math]::Abs($consoleX - $itemCenterX)
																								if ($distance -lt $closestDistance) {
																									$closestDistance = $distance
																									$closestItem = $menuItem
																								}
																							}
																							# If click is reasonably close to a menu item (within 10 characters), trigger it
																							if ($null -ne $closestItem -and $closestDistance -le 10 -and $null -ne $closestItem.hotkey) {
																								$script:MenuClickHotkey = $closestItem.hotkey
																							}
																						}
																					}
																				}
																			}
																		}
																	}
																}
															}
														}
													}
												} else {
													# Console handle is zero - try alternative methods
													# This can happen in Windows Terminal or other modern terminal emulators
													try {
														$currentProcess = Get-Process -Id $PID
														$mainWindowHandle = $currentProcess.MainWindowHandle
														if ($mainWindowHandle -ne [IntPtr]::Zero) {
															$consoleHandle = $mainWindowHandle
															# Try to check focus with the alternative handle
																try {
																	$foregroundWindow = [mJiggAPI.Mouse]::GetForegroundWindow()
																	$isConsoleFocused = ($foregroundWindow -eq $consoleHandle)
																	if ($DebugMode) {
																		$debugInfo.focus = if ($isConsoleFocused) { 
																			"yes (alt, fg=$($foregroundWindow.ToInt64()), console=$($consoleHandle.ToInt64()))" 
																		} else { 
																			"no (alt, fg=$($foregroundWindow.ToInt64()), console=$($consoleHandle.ToInt64()))" 
																		}
																	} else {
																		$debugInfo.focus = if ($isConsoleFocused) { "yes (alt)" } else { "no (alt)" }
																	}
																} catch {
																	$debugInfo.focus = "yes (alt fallback: $($_.Exception.Message))"
																	$isConsoleFocused = $true
																}
														} else {
															# Try to find window handle using FindWindow
															try {
																$foundHandle = Find-WindowHandle -ProcessId $PID
																if ($foundHandle -ne [IntPtr]::Zero) {
																	$consoleHandle = $foundHandle
																	# Check focus with found window handle
																		try {
																			$foregroundWindow = [mJiggAPI.Mouse]::GetForegroundWindow()
																			$isConsoleFocused = ($foregroundWindow -eq $consoleHandle)
																			$debugInfo.focus = if ($isConsoleFocused) { "yes (found)" } else { "no (found)" }
																		} catch {
																			$debugInfo.focus = "yes (found fallback: $($_.Exception.Message))"
																			$isConsoleFocused = $true
																		}
																} else {
																	# Try to get parent process (Windows Terminal) window handle
																	try {
																		$parentProcess = Get-Process -Id $currentProcess.Parent.Id -ErrorAction SilentlyContinue
																		if ($null -ne $parentProcess -and $parentProcess.MainWindowHandle -ne [IntPtr]::Zero) {
																			$consoleHandle = $parentProcess.MainWindowHandle
																			# Check focus with parent window handle
																			try {
																				# Try to check focus - use try-catch instead of method existence check
																				try {
																					$foregroundWindow = [mJiggAPI.Mouse]::GetForegroundWindow()
																					$isConsoleFocused = ($foregroundWindow -eq $consoleHandle)
																					$debugInfo.focus = if ($isConsoleFocused) { "yes (parent)" } else { "no (parent)" }
																				} catch {
																					# GetForegroundWindow not available - likely types need reload
																					# Assume focused as fallback since user is clicking in the window
																					$isConsoleFocused = $true
																					if ($_.Exception.Message -match "does not contain a method") {
																						$debugInfo.focus = "assumed (restart PowerShell to load GetForegroundWindow)"
																					} else {
																						$debugInfo.focus = "assumed (GetForegroundWindow failed: $($_.Exception.Message))"
																					}
																				}
																			} catch {
																				$debugInfo.focus = "yes (parent fallback)"
																				$isConsoleFocused = $true
																			}
																		} else {
																			# Try finding Windows Terminal process by name
																			try {
																				$wtProcesses = Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }
																				if ($null -ne $wtProcesses -and $wtProcesses.Count -gt 0) {
																					# Use the first Windows Terminal window we find
																					$consoleHandle = $wtProcesses[0].MainWindowHandle
																						try {
																							$foregroundWindow = [mJiggAPI.Mouse]::GetForegroundWindow()
																							$isConsoleFocused = ($foregroundWindow -eq $consoleHandle)
																							$debugInfo.focus = if ($isConsoleFocused) { "yes (WT)" } else { "no (WT)" }
																						} catch {
																							$debugInfo.focus = "yes (WT fallback: $($_.Exception.Message))"
																							$isConsoleFocused = $true
																						}
																				} else {
																					# Last resort: assume focused and proceed without window handle
																					$debugInfo.focus = "assumed (no handle)"
																					$isConsoleFocused = $true
																				}
																			} catch {
																				$debugInfo.focus = "assumed (WT search failed)"
																				$isConsoleFocused = $true
																			}
																		}
																	} catch {
																		# Parent process method failed, assume focused
																		$debugInfo.focus = "assumed (parent failed)"
																		$isConsoleFocused = $true
																	}
																}
															} catch {
																# FindWindow method failed, assume focused
																$debugInfo.focus = "assumed (FindWindow failed)"
																$isConsoleFocused = $true
															}
														}
													} catch {
														# All methods failed, assume focused
														$debugInfo.focus = "assumed (Get-Process failed)"
														$isConsoleFocused = $true
													}
												}
											} catch {
												# Log error for debugging
												$debugInfo.focus = "error: $($_.Exception.Message)"
											}
										}
									
									# Create a single consolidated log entry with all debug information
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] LButton clicked - Focus: $($debugInfo.focus), Screen: ($($debugInfo.screenX),$($debugInfo.screenY)), Console: ($($debugInfo.consoleX),$($debugInfo.consoleY)), Button: $($debugInfo.button)"
												shortText = " - [DEBUG] LButton - Focus: $($debugInfo.focus), Coords: ($($debugInfo.consoleX),$($debugInfo.consoleY)), Button: $($debugInfo.button)"
											}
										)
									}
									
									# Record when we logged this click for debouncing
									$script:LastClickLogTime = Get-Date
									
									# Check if mouse movement was recent - if so, don't force immediate redraw
									# This respects the mouse movement delay to prevent console stutter
									$shouldForceRedraw = $true
									if ($null -ne $script:LastMouseMovementTime) {
										$timeSinceMouseMovement = ((Get-Date) - $script:LastMouseMovementTime).TotalMilliseconds
										# Don't force redraw if mouse moved within last 200ms
										if ($timeSinceMouseMovement -lt 200) {
											$shouldForceRedraw = $false
										}
									}
									
									# Only force immediate UI redraw if mouse hasn't moved recently
									if ($shouldForceRedraw) {
										$SkipUpdate = $false
										$forceRedraw = $true
										break waitLoop
									}
									# Otherwise, the log will appear on the next normal update cycle
								}
									
									$mouseButtonName = Get-KeyName -keyCode $keyCode
									if (-not $mouseButtonName) {
										if ($keyCode -eq 0x07) { $mouseButtonName = "XButton3" }
										elseif ($keyCode -eq 0x08) { $mouseButtonName = "XButton4" }
										elseif ($keyCode -eq 0x09) { $mouseButtonName = "XButton5" }
										elseif ($keyCode -eq 0x0A) { $mouseButtonName = "XButton6" }
										elseif ($keyCode -eq 0x0B) { $mouseButtonName = "XButton7" }
										elseif ($keyCode -eq 0x0C) { $mouseButtonName = "XButton8" }
										elseif ($keyCode -eq 0x0D) { $mouseButtonName = "XButton9" }
										elseif ($keyCode -eq 0x0E) { $mouseButtonName = "XButton10" }
										elseif ($keyCode -eq 0x0F) { $mouseButtonName = "XButton11" }
									}
									if ($mouseButtonName -and $intervalMouseInputs -notcontains $mouseButtonName) {
										$intervalMouseInputs += $mouseButtonName
										$userInputDetected = $true
										$mouseInputDetected = $true
										if ($script:AutoResumeDelaySeconds -gt 0) {
											$LastUserInputTime = Get-Date
										}
									}
								}
								$script:previousKeyStates[$keyCode] = $isCurrentlyPressed
								continue
							}
							
							# Check keyboard keys
							$currentKeyState = [mJiggAPI.Keyboard]::GetAsyncKeyState($keyCode)
							$isCurrentlyPressed = (($currentKeyState -band 0x8000) -ne 0)
							$wasJustPressed = (($currentKeyState -band 0x0001) -ne 0)
							$wasPreviouslyPressed = if ($script:previousKeyStates.ContainsKey($keyCode)) { $script:previousKeyStates[$keyCode] } else { $false }
							
							if ($wasJustPressed -or ($isCurrentlyPressed -and -not $wasPreviouslyPressed)) {
								$keyName = Get-KeyName -keyCode $keyCode
								$keyAdded = $false
								if ($keyName) {
									if ($intervalKeys -notcontains $keyName) {
										$intervalKeys += $keyName
										$keyAdded = $true
									}
								} else {
									$unknownKeyName = "Unknown(0x$($keyCode.ToString('X2')))"
									if ($intervalKeys -notcontains $unknownKeyName) {
										$intervalKeys += $unknownKeyName
										$keyAdded = $true
									}
								}
								if ($keyAdded) {
									$userInputDetected = $true
									$keyboardInputDetected = $true
									if ($script:AutoResumeDelaySeconds -gt 0) {
										$LastUserInputTime = Get-Date
									}
								}
							}
							$script:previousKeyStates[$keyCode] = $isCurrentlyPressed
						}
					}
					
					# Check for console keyboard input (menu hotkeys) - only every 200ms to avoid stutter
					# Also check for menu clicks immediately (they're set by mouse click handler)
					$menuHotkeyToProcess = $null
					if ($null -ne $script:MenuClickHotkey) {
						# Menu item was clicked - process it immediately
						$menuHotkeyToProcess = $script:MenuClickHotkey
						$script:MenuClickHotkey = $null  # Clear it after using
					} elseif ($x % 4 -eq 0) {
						# Read available keys for menu hotkeys (only every 200ms)
						$lastKeyPress = $null
						$lastKeyInfo = $null
						$keysRead = 0
						$maxKeysToRead = 10  # Limit to prevent infinite loops
						while ($Host.UI.RawUI.KeyAvailable -and $keysRead -lt $maxKeysToRead) {
							try {
								$keyInfo = $Host.UI.RawUI.ReadKey("IncludeKeyup,NoEcho")
								$keysRead++
								$keyPress = $keyInfo.Character
								$isEscape = ($keyInfo.Key -eq "Escape" -or $keyInfo.VirtualKeyCode -eq 27)
								$isKeyDown = if ($null -ne $keyInfo.KeyDown) { $keyInfo.KeyDown } else { $false }
								
								# Only process key up events
								if (-not $isKeyDown) {
									$keyId = if ($isEscape) { "Escape" } elseif ($keyPress) { $keyPress } else { $null }
									if ($keyId) {
										if ($isEscape) {
											$lastKeyPress = "Escape"
											$lastKeyInfo = $keyInfo
										} else {
											$lastKeyPress = $keyPress
											$lastKeyInfo = $keyInfo
										}
									}
								}
							} catch {
								break
							}
						}
					}
					
					# Process menu hotkeys (check both lastKeyPress and menuHotkeyToProcess)
					if ($null -ne $menuHotkeyToProcess) {
						# Process menu click hotkey immediately
						$lastKeyPress = $menuHotkeyToProcess
						$lastKeyInfo = $null
					}
					
					if ($null -ne $lastKeyPress -or $null -ne $lastKeyInfo) {
						$shouldProcessEscape = ($lastKeyPress -eq "Escape" -or ($null -ne $lastKeyInfo -and ($lastKeyInfo.Key -eq "Escape" -or $lastKeyInfo.VirtualKeyCode -eq 27)))
						if ($shouldProcessEscape) {
							$lastKeyPress = $null
							$lastKeyInfo = $null
							$HostWidthRef = [ref]$HostWidth
							$HostHeightRef = [ref]$HostHeight
							$quitResult = Show-QuitConfirmationDialog -hostWidthRef $HostWidthRef -hostHeightRef $HostHeightRef
							$HostWidth = $HostWidthRef.Value
							$HostHeight = $HostHeightRef.Value
							if ($quitResult.NeedsRedraw) {
								$SkipUpdate = $true
								$forceRedraw = $true
								clear-host
								break
							}
							if ($quitResult.Result -eq $true) {
								Clear-Host
								$runtime = (Get-Date) - $ScriptStartTime
								$hours = [math]::Floor($runtime.TotalHours)
								$minutes = $runtime.Minutes
								$seconds = $runtime.Seconds
								$runtimeStr = ""
								if ($hours -gt 0) {
									$runtimeStr = "$hours hour$(if ($hours -ne 1) { 's' }), $minutes minute$(if ($minutes -ne 1) { 's' })"
								} elseif ($minutes -gt 0) {
									$runtimeStr = "$minutes minute$(if ($minutes -ne 1) { 's' }), $seconds second$(if ($seconds -ne 1) { 's' })"
								} else {
									$runtimeStr = "$seconds second$(if ($seconds -ne 1) { 's' })"
								}
								Write-Host ""
								Write-Host "  mJig(`u{1F400}) " -NoNewline -ForegroundColor Magenta
								Write-Host "Stopped" -ForegroundColor Red
								Write-Host ""
								Write-Host "  Runtime: " -NoNewline -ForegroundColor Yellow
								Write-Host $runtimeStr -ForegroundColor Green
								Write-Host ""
								return
							} else {
								$SkipUpdate = $true
								$forceRedraw = $true
								clear-host
								break
							}
						} elseif ($lastKeyPress -eq "q") {
								$lastKeyPress = $null
								$lastKeyInfo = $null
								
								# Debug: Log quit dialog opened
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Quit dialog opened"
												shortText = " - [DEBUG] Quit opened"
											}
										)
									}
								}
								
								$HostWidthRef = [ref]$HostWidth
								$HostHeightRef = [ref]$HostHeight
								$quitResult = Show-QuitConfirmationDialog -hostWidthRef $HostWidthRef -hostHeightRef $HostHeightRef
								$HostWidth = $HostWidthRef.Value
								$HostHeight = $HostHeightRef.Value
								if ($quitResult.NeedsRedraw) {
									$SkipUpdate = $true
									$forceRedraw = $true
									clear-host
									break
								}
								if ($quitResult.Result -eq $true) {
									# Debug: Log quit confirmed
									if ($DebugMode) {
										if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
											$LogArray = @()
										}
										$LogArray += [PSCustomObject]@{
											logRow = $true
											components = @(
												@{
													priority = 1
													text = (Get-Date).ToString("HH:mm:ss")
													shortText = (Get-Date).ToString("HH:mm:ss")
												},
												@{
													priority = 2
													text = " - [DEBUG] Quit confirmed"
													shortText = " - [DEBUG] Quit confirmed"
												}
											)
										}
									}
									Clear-Host
									$runtime = (Get-Date) - $ScriptStartTime
									$hours = [math]::Floor($runtime.TotalHours)
									$minutes = $runtime.Minutes
									$seconds = $runtime.Seconds
									$runtimeStr = ""
									if ($hours -gt 0) {
										$runtimeStr = "$hours hour$(if ($hours -ne 1) { 's' }), $minutes minute$(if ($minutes -ne 1) { 's' })"
									} elseif ($minutes -gt 0) {
										$runtimeStr = "$minutes minute$(if ($minutes -ne 1) { 's' }), $seconds second$(if ($seconds -ne 1) { 's' })"
									} else {
										$runtimeStr = "$seconds second$(if ($seconds -ne 1) { 's' })"
									}
									Write-Host ""
									Write-Host "  mJig(" -NoNewline -ForegroundColor Magenta
									Write-Host "`u{1F400}" -NoNewline -ForegroundColor White
									Write-Host ") " -NoNewline -ForegroundColor Magenta
									Write-Host "Stopped" -ForegroundColor Red
									Write-Host ""
									Write-Host "  Runtime: " -NoNewline -ForegroundColor Yellow
									Write-Host $runtimeStr -ForegroundColor Green
									Write-Host ""
									return
								} else {
									# Debug: Log quit canceled
									if ($DebugMode) {
										if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
											$LogArray = @()
										}
										$LogArray += [PSCustomObject]@{
											logRow = $true
											components = @(
												@{
													priority = 1
													text = (Get-Date).ToString("HH:mm:ss")
													shortText = (Get-Date).ToString("HH:mm:ss")
												},
												@{
													priority = 2
													text = " - [DEBUG] Quit canceled"
													shortText = " - [DEBUG] Quit canceled"
												}
											)
										}
									}
									$SkipUpdate = $true
									$forceRedraw = $true
									clear-host
									break
								}
							} elseif ($lastKeyPress -eq "v") {
								$oldOutput = $Output
								if ($Output -eq "hidden") {
									if ($PreviousView -ne $null) {
										$Output = $PreviousView
									} else {
										$Output = "min"
									}
								} else {
									if ($Output -eq "full") {
										$Output = "min"
									} else {
										$Output = "full"
									}
								}
								# Debug: Log view toggle
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] View toggle: $oldOutput → $Output"
												shortText = " - [DEBUG] View: $oldOutput → $Output"
											}
										)
									}
								}
								$SkipUpdate = $true
								$forceRedraw = $true
								clear-host
								break
							} elseif ($lastKeyPress -eq "h") {
								$oldOutput = $Output
								if ($Output -eq "hidden") {
									if ($PreviousView -ne $null) {
										$Output = $PreviousView
									} else {
										$Output = "min"
									}
									$PreviousView = $null
								} else {
									$PreviousView = $Output
									$Output = "hidden"
								}
								# Debug: Log hide/show toggle
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Hide/Show toggle: $oldOutput → $Output"
												shortText = " - [DEBUG] Hide/Show: $oldOutput → $Output"
											}
										)
									}
								}
								$SkipUpdate = $true
								$forceRedraw = $true
								clear-host
								break
							} elseif ($lastKeyPress -eq "m" -and $Output -eq "full") {
								# Debug: Log movement dialog opened (before calling dialog)
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Movement dialog opened"
												shortText = " - [DEBUG] Movement opened"
											}
										)
									}
								}
								
								$HostWidthRef = [ref]$HostWidth
								$HostHeightRef = [ref]$HostHeight
								$dialogResult = Show-MovementModifyDialog -currentIntervalSeconds $script:IntervalSeconds -currentIntervalVariance $script:IntervalVariance -currentMoveSpeed $script:MoveSpeed -currentMoveVariance $script:MoveVariance -currentTravelDistance $script:TravelDistance -currentTravelVariance $script:TravelVariance -currentAutoResumeDelaySeconds $script:AutoResumeDelaySeconds -hostWidthRef $HostWidthRef -hostHeightRef $HostHeightRef
								$HostWidth = $HostWidthRef.Value
								$HostHeight = $HostHeightRef.Value
								
								# Debug: Verify logs are preserved after dialog closes
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Movement dialog closed"
												shortText = " - [DEBUG] Movement closed"
											}
										)
									}
								}
								
								if ($dialogResult.NeedsRedraw) {
									$SkipUpdate = $true
									$forceRedraw = $true
									clear-host
									break
								}
								if ($null -ne $dialogResult.Result) {
									$oldIntervalSeconds = $script:IntervalSeconds
									$oldIntervalVariance = $script:IntervalVariance
									$oldMoveSpeed = $script:MoveSpeed
									$oldMoveVariance = $script:MoveVariance
									$oldTravelDistance = $script:TravelDistance
									$oldTravelVariance = $script:TravelVariance
									$oldAutoResumeDelaySeconds = $script:AutoResumeDelaySeconds
									$script:IntervalSeconds = $dialogResult.Result.IntervalSeconds
									$script:IntervalVariance = $dialogResult.Result.IntervalVariance
									$script:MoveSpeed = $dialogResult.Result.MoveSpeed
									$script:MoveVariance = $dialogResult.Result.MoveVariance
									$script:TravelDistance = $dialogResult.Result.TravelDistance
									$script:TravelVariance = $dialogResult.Result.TravelVariance
									$script:AutoResumeDelaySeconds = $dialogResult.Result.AutoResumeDelaySeconds
									$changeDetails = @()
									if ($oldIntervalSeconds -ne $script:IntervalSeconds) { $changeDetails += "Interval: $oldIntervalSeconds → $($script:IntervalSeconds)" }
									if ($oldIntervalVariance -ne $script:IntervalVariance) { $changeDetails += "IntervalVar: $oldIntervalVariance → $($script:IntervalVariance)" }
									if ($oldMoveSpeed -ne $script:MoveSpeed) { $changeDetails += "Speed: $oldMoveSpeed → $($script:MoveSpeed)" }
									if ($oldMoveVariance -ne $script:MoveVariance) { $changeDetails += "SpeedVar: $oldMoveVariance → $($script:MoveVariance)" }
									if ($oldTravelDistance -ne $script:TravelDistance) { $changeDetails += "Distance: $oldTravelDistance → $($script:TravelDistance)" }
									if ($oldTravelVariance -ne $script:TravelVariance) { $changeDetails += "DistVar: $oldTravelVariance → $($script:TravelVariance)" }
									if ($oldAutoResumeDelaySeconds -ne $script:AutoResumeDelaySeconds) { $changeDetails += "Delay: $oldAutoResumeDelaySeconds → $($script:AutoResumeDelaySeconds)" }
									if ($changeDetails.Count -gt 0) {
										$changeDate = Get-Date
										$changeMessage = " - Settings updated: " + ($changeDetails -join ", ")
										$changeShortMessage = " - Updated: " + ($changeDetails -join ", ")
										$changeLogComponents = @(
											@{priority = 1; text = $changeDate.ToString(); shortText = $changeDate.ToString("HH:mm:ss")},
											@{priority = 2; text = $changeMessage; shortText = $changeShortMessage}
										)
										$LogArray += [PSCustomObject]@{logRow = $true; components = $changeLogComponents}
									}
								}
								$SkipUpdate = $true
								$forceRedraw = $true
								clear-host
								break
							} elseif ($lastKeyPress -eq "t") {
								# Debug: Log time dialog opened (before calling dialog)
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Time dialog opened"
												shortText = " - [DEBUG] Time opened"
											}
										)
									}
								}
								
								$HostWidthRef = [ref]$HostWidth
								$HostHeightRef = [ref]$HostHeight
								$dialogResult = Show-TimeChangeDialog -currentEndTime $endTimeInt -hostWidthRef $HostWidthRef -hostHeightRef $HostHeightRef
								$HostWidth = $HostWidthRef.Value
								$HostHeight = $HostHeightRef.Value
								
								# Debug: Verify logs are preserved after dialog closes
								if ($DebugMode) {
									if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
										$LogArray = @()
									}
									$LogArray += [PSCustomObject]@{
										logRow = $true
										components = @(
											@{
												priority = 1
												text = (Get-Date).ToString("HH:mm:ss")
												shortText = (Get-Date).ToString("HH:mm:ss")
											},
											@{
												priority = 2
												text = " - [DEBUG] Time dialog closed"
												shortText = " - [DEBUG] Time closed"
											}
										)
									}
								}
								
								if ($dialogResult.NeedsRedraw) {
									$SkipUpdate = $true
									$forceRedraw = $true
									clear-host
									break
								}
								if ($null -ne $dialogResult.Result) {
									$oldEndTimeInt = $endTimeInt
									$oldEndTimeStr = $endTimeStr
									if ($dialogResult.Result -eq -1) {
										$endTimeInt = -1
										$endTimeStr = ""
										$end = ""
										$changeDate = Get-Date
										$changeMessage = if ([string]::IsNullOrEmpty($oldEndTimeStr)) {" - End time cleared"} else {" - End time cleared (was: $oldEndTimeStr)"}
										$changeShortMessage = " - End time cleared"
										$changeLogComponents = @(
											@{priority = 1; text = $changeDate.ToString(); shortText = $changeDate.ToString("HH:mm:ss")},
											@{priority = 2; text = $changeMessage; shortText = $changeShortMessage}
										)
										$LogArray += [PSCustomObject]@{logRow = $true; components = $changeLogComponents}
									} else {
										$endTimeInt = $dialogResult.Result
										$endTimeStr = $endTimeInt.ToString().PadLeft(4, '0')
										$currentTime = Get-Date -Format "HHmm"
										if ($endTimeInt -le [int]$currentTime) {
											$tommorow = (Get-Date).AddDays(1)
											$endDate = Get-Date $tommorow -Format "MMdd"
										} else {
											$endDate = Get-Date -Format "MMdd"
										}
										$end = "$endDate$endTimeStr"
										$changeDate = Get-Date
										$changeMessage = if ($oldEndTimeInt -eq -1 -or [string]::IsNullOrEmpty($oldEndTimeStr)) {" - End time set: $endTimeStr"} else {" - End time changed: $oldEndTimeStr → $endTimeStr"}
										$changeShortMessage = " - End time: $endTimeStr"
										$changeLogComponents = @(
											@{priority = 1; text = $changeDate.ToString(); shortText = $changeDate.ToString("HH:mm:ss")},
											@{priority = 2; text = $changeMessage; shortText = $changeShortMessage}
										)
										$LogArray += [PSCustomObject]@{logRow = $true; components = $changeLogComponents}
									}
								}
								$SkipUpdate = $true
								$forceRedraw = $true
								clear-host
								break
							}
						}
					
					# Check for window size changes - process only when window stops resizing
					# Only check every 200ms (every 4th iteration) to avoid blocking Windows from processing mouse messages
					# Frequent window property access can interfere with cursor rendering during user mouse movement
					if ($Output -ne "hidden" -and ($x % 4 -eq 0)) {
						$pshost = Get-Host
						$pswindow = $pshost.UI.RawUI
						$newWindowSize = $pswindow.WindowSize
						$newBufferSize = $pswindow.BufferSize
						
						# Check if buffer size changed (e.g., from text zoom)
						# When text is zoomed, horizontal buffer size changes - this determines line length
						$bufferSizeChanged = ($null -eq $OldBufferSize -or 
							$newBufferSize.Width -ne $OldBufferSize.Width -or 
							$newBufferSize.Height -ne $OldBufferSize.Height)
						
						# Check if horizontal buffer changed but window width didn't (text zoom)
						# Also ensure vertical buffer matches window height
						$horizontalBufferChanged = ($null -ne $OldBufferSize -and $newBufferSize.Width -ne $OldBufferSize.Width)
						$windowWidthUnchanged = ($null -ne $oldWindowSize -and $newWindowSize.Width -eq $oldWindowSize.Width)
						
						# Set vertical buffer to match window height
						if ($newBufferSize.Height -ne $newWindowSize.Height) {
							try {
								$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($newBufferSize.Width, $newWindowSize.Height)
								$newBufferSize = $pswindow.BufferSize
							} catch {
								# If setting buffer size fails, continue with current buffer size
							}
						}
						
						# If horizontal buffer changed but window width didn't, it's text zoom - use buffer width for line length
						if ($horizontalBufferChanged -and $windowWidthUnchanged -and $null -ne $OldBufferSize) {
							# Text zoom detected - use buffer width for line length calculations
							$OldBufferSize = $newBufferSize
							# Use buffer width for HostWidth (determines line length), window height for HostHeight
							$HostWidth = $newBufferSize.Width
							$HostHeight = $newWindowSize.Height
							# Don't update oldWindowSize - keep the original so resize handler doesn't trigger
							$SkipUpdate = $true
							$forceRedraw = $true
							$waitExecuted = $false
							clear-host
							break
						}
						
						# Check if window size is different from what we last processed (oldWindowSize)
						# Skip this check if we just handled a text zoom
						$windowSizeChanged = ($null -eq $oldWindowSize -or 
							$newWindowSize.Width -ne $oldWindowSize.Width -or 
							$newWindowSize.Height -ne $oldWindowSize.Height)
						
						if ($windowSizeChanged) {
							# Check if this is a new size (different from pending resize) or if we don't have a pending resize
							$isNewSize = ($null -eq $PendingResize -or 
								$newWindowSize.Width -ne $PendingResize.Width -or 
								$newWindowSize.Height -ne $PendingResize.Height)
							
							if ($isNewSize) {
								# New size detected - store it and reset timer
								$PendingResize = $newWindowSize
								$lastResizeDetection = Get-Date
							}
						}
						
						# Check if window has been stable (matches pending resize) long enough to process
						if ($null -ne $PendingResize -and $null -ne $lastResizeDetection) {
							# Verify current size still matches pending resize (window stopped changing)
							$sizeMatchesPending = ($newWindowSize.Width -eq $PendingResize.Width -and 
								$newWindowSize.Height -eq $PendingResize.Height)
							
							if ($sizeMatchesPending) {
								$timeSinceResize = ((Get-Date) - $lastResizeDetection).TotalMilliseconds
								if ($timeSinceResize -ge $ResizeThrottleMs) {
									# Window has been stable long enough - process the resize
									$currentWindowSize = $pswindow.WindowSize
									$currentBufferSize = $pswindow.BufferSize
									# Set vertical buffer to match window height, preserve horizontal buffer width
									try {
										$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($currentBufferSize.Width, $currentWindowSize.Height)
										$currentBufferSize = $pswindow.BufferSize
									} catch {
										# If setting buffer size fails, continue with current buffer size
									}
									# Update tracking variables
									$OldBufferSize = $currentBufferSize
									$oldWindowSize = $currentWindowSize
									$HostWidth = $currentWindowSize.Width
									$HostHeight = $currentWindowSize.Height
									$SkipUpdate = $true
									$forceRedraw = $true
									$waitExecuted = $false  # Mark that wait was interrupted, don't log this
									$PendingResize = $null  # Clear pending resize
									$lastResizeDetection = $null  # Clear detection time
									clear-host
									# Break out of wait loop to immediately redraw
									break
								}
							}
						}
					}
					
					start-sleep -m 50
				} until ($x -ge $math)
			}
			
			# Keyboard and mouse input checking is now done every 200ms in the wait loop above
			# This provides more reliable detection compared to checking once per interval
			
			# Check for mouse wheel scrolling
			try {
				if ([mJiggAPI.MouseHook]::hHook -ne [IntPtr]::Zero) {
					[mJiggAPI.MouseHook]::ProcessMessages()
					$currentWheelDelta = [mJiggAPI.MouseHook]::lastWheelDelta
					if ($currentWheelDelta -ne $PreviousMouseWheelDelta) {
						$deltaChange = $currentWheelDelta - $PreviousMouseWheelDelta
						if ($deltaChange -gt 0) {
							$wheelText = "Scroll Up"
							if ($intervalMouseInputs -notcontains $wheelText) {
								$intervalMouseInputs += $wheelText
								$userInputDetected = $true
								$mouseInputDetected = $true
							}
						} elseif ($deltaChange -lt 0) {
							$wheelText = "Scroll Down"
							if ($intervalMouseInputs -notcontains $wheelText) {
								$intervalMouseInputs += $wheelText
								$userInputDetected = $true
								$mouseInputDetected = $true
							}
						}
						$PreviousMouseWheelDelta = $currentWheelDelta
					}
				}
			} catch {
				# Hook not available, skip
			}
			
			# Check for window size changes (also check outside wait loop)
			# Only check if we haven't already detected a resize in this iteration
			if ($Output -ne "hidden" -and -not $forceRedraw) {
				$pshost = Get-Host
				$pswindow = $pshost.UI.RawUI
				$newWindowSize = $pswindow.WindowSize
				$newBufferSize = $pswindow.BufferSize
				
				# Check if buffer size changed (e.g., from text zoom)
				# When text is zoomed, horizontal buffer size changes - this determines line length
				$bufferSizeChanged = ($null -eq $OldBufferSize -or 
					$newBufferSize.Width -ne $OldBufferSize.Width -or 
					$newBufferSize.Height -ne $OldBufferSize.Height)
				
				# Check if horizontal buffer changed but window width didn't (text zoom)
				# Also ensure vertical buffer matches window height
				$horizontalBufferChanged = ($null -ne $OldBufferSize -and $newBufferSize.Width -ne $OldBufferSize.Width)
				$windowWidthUnchanged = ($null -ne $oldWindowSize -and $newWindowSize.Width -eq $oldWindowSize.Width)
				
				# Set vertical buffer to match window height
				if ($newBufferSize.Height -ne $newWindowSize.Height) {
					try {
						$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($newBufferSize.Width, $newWindowSize.Height)
						$newBufferSize = $pswindow.BufferSize
					} catch {
						# If setting buffer size fails, continue with current buffer size
					}
				}
				
				# If horizontal buffer changed but window width didn't, it's text zoom - use buffer width for line length
				$isTextZoom = $false
				if ($horizontalBufferChanged -and $windowWidthUnchanged -and $null -ne $OldBufferSize) {
					# Text zoom detected - use buffer width for line length calculations
					$isTextZoom = $true
				}
				
				if ($isTextZoom) {
					# Text zoom detected - use buffer width for line length calculations
					$OldBufferSize = $newBufferSize
					# Use buffer width for hostWidth (determines line length), window height for hostHeight
					$HostWidth = $newBufferSize.Width
					$HostHeight = $newWindowSize.Height
					# Don't update oldWindowSize - keep the original so resize handler doesn't trigger
					$SkipUpdate = $true
					$forceRedraw = $true
					clear-host
					# Skip window resize check since this is just a zoom
					$windowSizeChanged = $false
				} else {
					# Check if window size is different from what we last processed (oldWindowSize)
					$windowSizeChanged = ($null -eq $oldWindowSize -or 
						$newWindowSize.Width -ne $oldWindowSize.Width -or 
						$newWindowSize.Height -ne $oldWindowSize.Height)
				}
				
				if ($windowSizeChanged) {
					# Check if this is a new size (different from pending resize) or if we don't have a pending resize
					$isNewSize = ($null -eq $PendingResize -or 
						$newWindowSize.Width -ne $PendingResize.Width -or 
						$newWindowSize.Height -ne $PendingResize.Height)
					
					if ($isNewSize) {
						# New size detected - store it and reset timer
						$PendingResize = $newWindowSize
						$lastResizeDetection = Get-Date
					}
				}
				
				# Check if window has been stable (matches pending resize) long enough to process
				if ($null -ne $PendingResize -and $null -ne $lastResizeDetection) {
					# Verify current size still matches pending resize (window stopped changing)
					$sizeMatchesPending = ($newWindowSize.Width -eq $PendingResize.Width -and 
						$newWindowSize.Height -eq $PendingResize.Height)
					
					if ($sizeMatchesPending) {
						$timeSinceResize = ((Get-Date) - $lastResizeDetection).TotalMilliseconds
						if ($timeSinceResize -ge $ResizeThrottleMs) {
							# Window has been stable long enough - process the resize
							$currentWindowSize = $pswindow.WindowSize
							$currentBufferSize = $pswindow.BufferSize
							# Set vertical buffer to match window height, preserve horizontal buffer width
							try {
								$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($currentBufferSize.Width, $currentWindowSize.Height)
								$currentBufferSize = $pswindow.BufferSize
							} catch {
								# If setting buffer size fails, continue with current buffer size
							}
							# Update tracking variables
							$OldBufferSize = $currentBufferSize
							$oldWindowSize = $currentWindowSize
							$HostWidth = $currentWindowSize.Width
							$HostHeight = $currentWindowSize.Height
							$SkipUpdate = $true
							$forceRedraw = $true
							$PendingResize = $null  # Clear pending resize
							$lastResizeDetection = $null  # Clear detection time
							clear-host
						}
					}
				}
			}
			
			# Check if this is the first run (before we modify lastMovementTime)
			$isFirstRun = ($null -eq $LastMovementTime)
			
			# Determine if we should skip the update based on user input or first run
			if ($userInputDetected) {
				$SkipUpdate = $true
			} elseif ($isFirstRun) {
				# Skip automated input on first run
				$SkipUpdate = $true
			} elseif (-not $forceRedraw) {
				# Only set skipUpdate to false if we're not forcing a redraw
				$SkipUpdate = $false
			}
			
			# Prepare UI dimensions
			$outputline = 0
			$oldRows = $Rows
			$Rows = $HostHeight - 6
			
			# Save current log array BEFORE building new one (this preserves the previous iteration's logs)
			# On first run, $LogArray might be null or empty, so handle that case
			if ($null -eq $LogArray -or $LogArray.Count -eq 0) {
				$tempOldLogArray = @()
			} else {
				$tempOldLogArray = $LogArray.Clone()
			}
			
			# Handle log array resizing when window height changes
			if ($oldRows -ne $Rows) {
				if ($oldRows -lt $Rows) {
					# Window got taller - add empty entries at the beginning
					$insertArray = @()
					$row = [PSCustomObject]@{
						logRow = $true
						components = @()
					}
					for ($i = 0; $i -lt ($Rows - $oldRows); $i++) {
						$insertArray += $row
					}
					$tempOldLogArray = $insertArray + $tempOldLogArray
				} else {
					# Window got shorter - trim old entries from the beginning
					$trimCount = $oldRows - $Rows
					if ($tempOldLogArray.Count -gt $trimCount) {
						$tempOldLogArray = $tempOldLogArray[$trimCount..($tempOldLogArray.Count - 1)]
					} else {
						$tempOldLogArray = @()
					}
				}
			}
			
			# Build new log array: take all entries from old array (they scroll up)
			# The old array already has the previous logs, we just need to keep them
			$LogArray = @()
			
			# Copy all old log entries (they will scroll up by one position)
			# We keep up to $Rows entries from the old array (we'll add a new one, then trim to $Rows)
			$maxOldEntries = $Rows
			$startIndex = [math]::Max(0, $tempOldLogArray.Count - $maxOldEntries)
			
			for ($i = $startIndex; $i -lt $tempOldLogArray.Count; $i++) {
				# Preserve components if they exist, otherwise create empty entry
				if ($tempOldLogArray[$i].components) {
					$LogArray += [PSCustomObject]@{
						logRow = $true
						components = $tempOldLogArray[$i].components
					}
				} else {
					# Legacy format - convert to components if needed
					if ($tempOldLogArray[$i].value) {
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @(@{
								priority = 1
								text = $tempOldLogArray[$i].value
								shortText = $tempOldLogArray[$i].value
							})
						}
					} else {
						$LogArray += [PSCustomObject]@{
							logRow = $true
							components = @()
						}
					}
				}
			}
			
			# Fill remaining slots with empty entries if we don't have enough old entries
			# We fill up to $Rows entries (before adding the new one)
			while ($LogArray.Count -lt $Rows) {
				$LogArray += [PSCustomObject]@{
					logRow = $true
					components = @()
				}
			}
			
			# Check current mouse position to detect user movement (simple approach - only check at end of interval)
			# Compare end position to start position to detect if user moved mouse during the interval
			# This is simpler and doesn't interfere with mouse movement like checking during the wait loop
			# Use direct Windows API call for better performance (avoids .NET stutter)
			$point = New-Object mJiggAPI.POINT
			$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
			if ($hasGetCursorPos) {
				if ([mJiggAPI.Mouse]::GetCursorPos([ref]$point)) {
					# Convert POINT to System.Drawing.Point for compatibility
					$currentPos = New-Object System.Drawing.Point($point.X, $point.Y)
				} else {
					# API call failed - skip this check
					$currentPos = $null
				}
			} else {
				# Method not available - skip this check
				$currentPos = $null
			}
			$PosUpdate = $false
			$x = 0
			$y = 0
			
			# Only check for mouse movement if we haven't already detected user input
			# Skip checking if we recently performed automated mouse movement (within last 300ms)
			# This prevents our own automated movement from being detected as user input
			$shouldCheckMouseAfterWait = $true
			if ($null -ne $LastAutomatedMouseMovement) {
				$timeSinceAutomatedMovement = ((Get-Date) - $LastAutomatedMouseMovement).TotalMilliseconds
				if ($timeSinceAutomatedMovement -lt 300) {
					# Too soon after our automated movement - skip mouse detection
					$shouldCheckMouseAfterWait = $false
				}
			}
			
			if ($shouldCheckMouseAfterWait -and -not $userInputDetected -and $null -ne $mousePosAtStart -and $null -ne $currentPos) {
				# Compare current position to position at start of interval (simple approach)
				$deltaX = [Math]::Abs($currentPos.X - $mousePosAtStart.X)
				$deltaY = [Math]::Abs($currentPos.Y - $mousePosAtStart.Y)
				$movementThreshold = 3  # Only detect movement if it's more than 3 pixels
				
				if ($deltaX -gt $movementThreshold -or $deltaY -gt $movementThreshold) {
					# Check if this movement is from our automated movement
					$isAutomatedPos = ($null -ne $automatedMovementPos -and 
									   $null -ne $currentPos -and
									   $currentPos.X -eq $automatedMovementPos.X -and 
									   $currentPos.Y -eq $automatedMovementPos.Y)
					if (-not $isAutomatedPos) {
						# User moved mouse during interval - skip automated movement
						$SkipUpdate = $true
						$PosUpdate = $false
						$mouseInputDetected = $true
						# Reset auto-resume delay timer on user input
						if ($script:AutoResumeDelaySeconds -gt 0) {
							$LastUserInputTime = Get-Date
						}
						# Add mouse movement to detected inputs with emoji only
						$mouseMoveText = "`u{1F400}"
						if ($intervalMouseInputs -notcontains $mouseMoveText) {
							$intervalMouseInputs += $mouseMoveText
						}
						$LastPos = $currentPos
						$automatedMovementPos = $null  # Clear automated position since user moved
					}
					# If it matches our automated position, ignore it (it's from our movement)
				}
			}
			
			# Check if auto-resume delay timer is active (check before skipUpdate logic)
			$cooldownActive = $false
			$secondsRemaining = 0
			if ($script:AutoResumeDelaySeconds -gt 0) {
				if ($null -eq $LastUserInputTime) {
					# Timer hasn't started yet (no user input detected yet) - allow movement
					$cooldownActive = $false
				} else {
					$timeSinceInput = ((Get-Date) - $LastUserInputTime).TotalSeconds
					if ($timeSinceInput -lt $script:AutoResumeDelaySeconds) {
						$cooldownActive = $true
						$secondsRemaining = [Math]::Ceiling($script:AutoResumeDelaySeconds - $timeSinceInput)
					} else {
						# Timer expired - clear it
						# Debug: Log resume (timer expired)
						if ($DebugMode -and $null -ne $LastUserInputTime) {
							if ($null -eq $LogArray -or -not ($LogArray -is [Array])) {
								$LogArray = @()
							}
							$LogArray += [PSCustomObject]@{
								logRow = $true
								components = @(
									@{
										priority = 1
										text = (Get-Date).ToString("HH:mm:ss")
										shortText = (Get-Date).ToString("HH:mm:ss")
									},
									@{
										priority = 2
										text = " - [DEBUG] Auto-resume delay expired, resuming"
										shortText = " - [DEBUG] Resumed"
									}
								)
							}
						}
						$LastUserInputTime = $null
						$cooldownActive = $false
					}
				}
			}
			
			if ($SkipUpdate -ne $true) {
				if ($cooldownActive) {
					# Timer is active - skip coordinate updates and simulated key presses
					$SkipUpdate = $true
					$PosUpdate = $false
					# Store cooldown state for log component building (don't log directly here)
				} else {
					# No user movement detected - perform automated movement
					# Get fresh position right before movement to avoid stutter
					# Try to use direct Windows API call for better performance
					$point = New-Object mJiggAPI.POINT
					$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
					if ($hasGetCursorPos) {
						if ([mJiggAPI.Mouse]::GetCursorPos([ref]$point)) {
							# Convert POINT to System.Drawing.Point for compatibility
							$pos = New-Object System.Drawing.Point($point.X, $point.Y)
						} else {
							# API call failed - use last known position
							$pos = $LastPos
						}
					} else {
						# Method not available - use last known position
						$pos = $LastPos
					}
					$PosUpdate = $true
				
				# Calculate travel distance with variance
				$baseDistance = $script:TravelDistance
				# Use double variance directly (Get-Random supports doubles, -Maximum is exclusive so add small epsilon)
				$varianceAmount = Get-Random -Minimum 0.0 -Maximum ($script:TravelVariance + 0.0001)
				$rasDist = Get-Random -Maximum 2
				if ($rasDist -eq 0) {
					$distance = $baseDistance - $varianceAmount
				} else {
					$distance = $baseDistance + $varianceAmount
				}
				# Ensure minimum distance of 1 pixel
				if ($distance -lt 1) {
					$distance = 1
				}
				
				# Calculate random direction (angle in radians)
				$angle = Get-Random -Minimum 0 -Maximum ([Math]::PI * 2)
				
				# Calculate target coordinates based on distance and angle
				$x = [Math]::Round($pos.X + ($distance * [Math]::Cos($angle)))
				$y = [Math]::Round($pos.Y + ($distance * [Math]::Sin($angle)))
				
				# Ensure coordinates stay within screen bounds
				$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
				$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
				$x = [Math]::Max(0, [Math]::Min($x, $screenWidth - 1))
				$y = [Math]::Max(0, [Math]::Min($y, $screenHeight - 1))
				
				# Calculate movement direction for arrow emoji
				try {
					$deltaX = $x - $pos.X
					$deltaY = $y - $pos.Y
					$directionArrow = Get-DirectionArrow -deltaX $deltaX -deltaY $deltaY -style "simple"
				} catch {
					# If arrow calculation fails, just use empty string
					$directionArrow = ""
				}
				
				# Calculate smooth movement path
				$movementPath = Get-SmoothMovementPath -startX $pos.X -startY $pos.Y -endX $x -endY $y -baseSpeedSeconds $script:MoveSpeed -varianceSeconds $script:MoveVariance
				$movementPoints = $movementPath.Points
				$LastMovementDurationMs = $movementPath.TotalTimeMs
				
				# Move through each point smoothly
				if ($movementPoints.Count -gt 1) {
					$timePerPoint = $LastMovementDurationMs / ($movementPoints.Count - 1)
					
					# Add a tiny initial delay to prevent stutter at movement start
					# This ensures smooth transition from wait loop to movement
					Start-Sleep -Milliseconds 1
					
					# Move to each intermediate point (skip first point as it's the start position)
					for ($i = 1; $i -lt $movementPoints.Count; $i++) {
						$point = $movementPoints[$i]
						[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point ($point.X, $point.Y)
						
						# Sleep for the calculated time per point (minimum 1ms)
						# Always sleep between points to ensure smooth movement
						if ($i -lt $movementPoints.Count - 1) {
							$sleepTime = [Math]::Max(1, [Math]::Round($timePerPoint))
							Start-Sleep -Milliseconds $sleepTime
						}
					}
				} else {
					# Single point or no movement needed - just move directly
					[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point ($x, $y)
				}
				
				# Use direct Windows API call for better performance (avoids .NET stutter)
				$point = New-Object mJiggAPI.POINT
				$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
				if ($hasGetCursorPos) {
					if ([mJiggAPI.Mouse]::GetCursorPos([ref]$point)) {
						# Convert POINT to System.Drawing.Point for compatibility
						$LastPos = New-Object System.Drawing.Point($point.X, $point.Y)
					} else {
						# API call failed - keep existing LastPos
					}
				} else {
					# Method not available - keep existing LastPos
				}
				$automatedMovementPos = $LastPos  # Track this as our automated movement position
				# Record when we performed this automated mouse movement to prevent it from being detected as user input
				$LastAutomatedMouseMovement = Get-Date
				
				# Send Right Alt key press (modifier key - won't type anything or interfere with apps)
				# This is needed for apps like Slack/Skype that check for keyboard activity
				# Using Right Alt specifically to avoid conflicts with Left Alt shortcuts
				try {
					$vkCode = [byte]0xA5  # VK_RMENU (Right Alt)
					[mJiggAPI.Keyboard]::keybd_event($vkCode, [byte]0, [uint32]0, [int]0)  # Key down
					Start-Sleep -Milliseconds 10
					[mJiggAPI.Keyboard]::keybd_event($vkCode, [byte]0, [uint32]0x0002, [int]0)  # Key up (KEYEVENTF_KEYUP = 0x0002)
					# Record when we sent this simulated key press to prevent it from being detected as user input
					$LastSimulatedKeyPress = Get-Date
					Start-Sleep -Milliseconds 50  # Small delay to let key state clear
				} catch {
					# If keybd_event fails, continue without keyboard input
					# Mouse movement alone should still work for most cases
				}
				
					$LastMovementTime = Get-Date
				}
			} else {
				# skipUpdate was set - just update tracking
				$PosUpdate = $false
				$LastPos = $currentPos
				if ($null -eq $LastMovementTime) {
					$LastMovementTime = Get-Date
				}
			}
			
			# Combine mouse inputs and keys for display
			# Mouse inputs (with emoji) should appear first, then keyboard keys
			$allInputs = @()
			# Add mouse movement first if present (it's just the emoji now)
			$mouseMovement = $intervalMouseInputs | Where-Object { $_ -eq "`u{1F400}" }
			if ($mouseMovement) {
				$allInputs += $mouseMovement
				# Add other mouse inputs (clicks) after movement
				$otherMouseInputs = $intervalMouseInputs | Where-Object { $_ -ne "`u{1F400}" }
				if ($otherMouseInputs) {
					$allInputs += $otherMouseInputs
				}
			} else {
				# No mouse movement, just add all mouse inputs
				if ($intervalMouseInputs.Count -gt 0) {
					$allInputs += $intervalMouseInputs
				}
			}
			# Add keyboard keys after mouse inputs
			if ($intervalKeys.Count -gt 0) {
				$allInputs += $intervalKeys
			}
			$PreviousIntervalKeys = $allInputs
			
			# Only create log entry when we complete a wait interval AND do something
			# Don't create log entries for window resize events
			$shouldCreateLogEntry = $false
			
			# If this is just a window resize (forceRedraw set), don't create log entry
			if ($forceRedraw -and -not $waitExecuted -and -not $PosUpdate) {
				# This is just a window resize redraw - skip log entry completely
				$shouldCreateLogEntry = $false
			} elseif ($PosUpdate) {
				# We did a movement - always log this
				$shouldCreateLogEntry = $true
			} elseif ($isFirstRun) {
				# First run - log this
				$shouldCreateLogEntry = $true
			} elseif ($waitExecuted -and -not $forceRedraw) {
				# We completed a wait interval (and it wasn't interrupted by resize) - log this
				$shouldCreateLogEntry = $true
			}
			
			if ($shouldCreateLogEntry) {
				# Build log entry components array (priority order: timestamp, message, coordinates, wait info, input detection)
				$logComponents = @()
				
				# Component 1: Timestamp (full format)
				$logComponents += @{
					priority = 1
					text = $date.ToString()
					shortText = $date.ToString("HH:mm:ss")
				}
				
				# Component 2: Main message
				if ($SkipUpdate -ne $true) {
					if ($PosUpdate) {
						# Get direction arrow if available
						$arrowText = if ($directionArrow) { " $directionArrow" } else { "" }
						$logComponents += @{
							priority = 2
							text = " - Coordinates updated$arrowText"
							shortText = " - Updated$arrowText"
						}
						# Component 3: Coordinates
						$logComponents += @{
							priority = 3
							text = " x$x/y$y"
							shortText = " x$x/y$y"
						}
					} else {
						$logComponents += @{
							priority = 2
							text = " - Input detected, skipping update"
							shortText = " - Input detected"
						}
					}
				} elseif ($isFirstRun) {
					# First run - show initialization message
					$logComponents += @{
						priority = 2
						text = " - Initialized"
						shortText = " - Initialized"
					}
				} elseif ($keyboardInputDetected -or $mouseInputDetected) {
					# User input was detected - show user input skip with KB/MS status
					$logComponents += @{
						priority = 2
						text = " - User input skip"
						shortText = " - Skipped"
					}
				} elseif ($cooldownActive) {
					# Auto-resume delay is active (no user input detected) - show custom message
					$logComponents += @{
						priority = 2
						text = " - Auto-Resume Delay"
						shortText = " - Auto-Resume Delay"
					}
					# Add resume timer component
					$logComponents += @{
						priority = 4
						text = " [Resume: ${secondsRemaining}s]"
						shortText = " [R: ${secondsRemaining}s]"
					}
				} else {
					$logComponents += @{
						priority = 2
						text = " - User input skip"
						shortText = " - Skipped"
					}
				}
				
				# Component 4: Wait interval info (only if not cooldown active or user input detected)
				if ($waitExecuted -and -not $cooldownActive) {
					$logComponents += @{
						priority = 4
						text = " [Interval:${interval}s]"
						shortText = " [Interval:${interval}s]"
					}
				} elseif (-not $isFirstRun -and -not $cooldownActive) {
					$logComponents += @{
						priority = 4
						text = " [First run]"
						shortText = " [First run]"
					}
				}
				
				# Component 5 & 6: Keyboard and Mouse detection (only when user input was detected, lowest priority - removed first)
				# These are the first to be removed when window gets narrow
				if ($SkipUpdate -eq $true -and -not $isFirstRun -and ($keyboardInputDetected -or $mouseInputDetected)) {
					# Keyboard detection status
					$kbStatus = if ($keyboardInputDetected) { "YES" } else { "NO" }
					$logComponents += @{
						priority = 5
						text = " [KB:$kbStatus]"
						shortText = " [K:" + $kbStatus.Substring(0,1) + "]"
					}
					
					# Mouse detection status
					$msStatus = if ($mouseInputDetected) { "YES" } else { "NO" }
					$logComponents += @{
						priority = 6
						text = " [MS:$msStatus]"
						shortText = " [M:" + $msStatus.Substring(0,1) + "]"
					}
				}
				
				# Add current log entry to array with components (append to end)
				$LogArray += [PSCustomObject]@{
					logRow = $true
					components = $logComponents
				}
				
				# Ensure we always have exactly $Rows entries
				# If we have more than $Rows, trim to keep the last $Rows entries (newest at bottom)
				if ($LogArray.Count -gt $Rows) {
					$LogArray = $LogArray[($LogArray.Count - $Rows)..($LogArray.Count - 1)]
				}
				# If we have fewer than $Rows, prepend empty entries at the beginning
				# This ensures the newest entry is always at the bottom (last index = $Rows - 1)
				while ($LogArray.Count -lt $Rows) {
					$LogArray = @([PSCustomObject]@{
						logRow = $true
						components = @()
					}) + $LogArray
				}
			} else {
				# No log entry created - ensure we have $Rows empty entries
				# Add empty entries at the beginning so existing entries stay at the bottom
				while ($LogArray.Count -lt $Rows) {
					$LogArray = @([PSCustomObject]@{
						logRow = $true
						components = @()
					}) + $LogArray
				}
			}
			
			# Final check: ensure we always have exactly $Rows entries before rendering
			# This handles edge cases where the array might not be properly sized
			# Initialize logArray if it's null or empty
			if ($null -eq $LogArray) {
				$LogArray = @()
			}
			# Ensure we have exactly $Rows entries
			while ($LogArray.Count -lt $Rows) {
				$LogArray = @([PSCustomObject]@{
					logRow = $true
					components = @()
				}) + $LogArray
			}
			# Trim if we somehow have more than $Rows (keep the last $Rows entries)
			if ($LogArray.Count -gt $Rows) {
				$LogArray = $LogArray[($LogArray.Count - $Rows)..($LogArray.Count - 1)]
			}
			# Final verification: logArray must have exactly $Rows entries at this point
			# If it doesn't, something went wrong - rebuild it with empty entries
			# This is a safety net to ensure we always have the correct structure
			if ($null -eq $LogArray -or $LogArray.Count -ne $Rows) {
				# Rebuild the array with exactly $Rows empty entries
				$LogArray = @()
				for ($fillIdx = 0; $fillIdx -lt $Rows; $fillIdx++) {
					$LogArray += [PSCustomObject]@{
						logRow = $true
						components = @()
					}
				}
				# If we had a log entry that was created, try to preserve it at the last index
				# (This shouldn't normally happen, but it's a safety check)
			}

			# Output Handling
			# Skip console updates if mouse movement was recently detected (every 50ms check) to prevent stutter
			# This prevents blocking console operations from interfering with Windows mouse message processing
			$skipConsoleUpdate = $false
			if ($null -ne $script:LastMouseMovementTime) {
				$timeSinceMouseMovement = ((Get-Date) - $script:LastMouseMovementTime).TotalMilliseconds
				# Skip console updates for 200ms after mouse movement to prevent stutter
				if ($timeSinceMouseMovement -lt 200) {
					$skipConsoleUpdate = $true
				}
			}
			# Force UI redraw when forceRedraw is true (e.g., after window resize) - override skipConsoleUpdate
			if ($forceRedraw) {
				$skipConsoleUpdate = $false
			}
			
			if ($Output -ne "hidden" -and -not $skipConsoleUpdate) {
				# Output blank line
				$t = $true
				try {
					[Console]::SetCursorPosition(0, $Outputline)
				} catch {
					clear-host
					$t = $false
				} finally {
					if ($t) {
						for ($i = $Host.UI.RawUI.CursorPosition.x; $i -lt $HostWidth; $i++) {
							write-host " " -NoNewline
						}
					}
				}
				$Outputline++

				# Output header
				$t = $true
				try {
					[Console]::SetCursorPosition(0, $Outputline)
				} catch {
					$t = $false
				} finally {
					if ($t) {
						# Calculate widths for centering times between mJig title and view tag
						# Left part: "  mJig(`u{1F400})" = 2 + 5 + 2 + 1 = 10 (with emoji)
						$headerLeftWidth = 2 + 5 + 2 + 1  # "  " + "mJig(" + emoji + ")"
						# Add DEBUGMODE text width if in debug mode
						if ($DebugMode) {
							$headerLeftWidth += 13  # " - DEBUGMODE" = 13 chars
						}
						
						# Time section: "Current`u{23F3}/" + time + " ➣  " + "End`u{23F3}/" + time (or "none")
						# Components: "Current" (7) + emoji (2) + "/" (1) + time + " ➣  " (4) + "End" (3) + emoji (2) + "/" (1) + time
						$timeSectionBaseWidth = 7 + 2 + 1 + 4 + 3 + 2 + 1  # Fixed text parts
						# Determine end time display text
						if ($endTimeInt -eq -1 -or [string]::IsNullOrEmpty($endTimeStr)) {
							$endTimeDisplay = "none"
						} else {
							$endTimeDisplay = $endTimeStr
						}
						$timeSectionTimeWidth = $currentTime.Length + $endTimeDisplay.Length
						$timeSectionWidth = $timeSectionBaseWidth + $timeSectionTimeWidth
						
						# Right part: view tag (with 2 spaces after for right margin)
						if ($Output -eq "full") {
							$viewTagText = " Full"
						} else {
							$viewTagText = " Minimum"
						}
						$viewTagWidth = $viewTagText.Length
						$rightMarginWidth = 2  # 2 spaces after view tag
						
						# Calculate spacing to center times between left and right parts
						# Account for 2 spaces after view tag
						$totalUsedWidth = $headerLeftWidth + $timeSectionWidth + $viewTagWidth + $rightMarginWidth
						$remainingSpace = $HostWidth - $totalUsedWidth
						$spacingBeforeTimes = [math]::Max(1, [math]::Floor($remainingSpace / 2))
						$spacingAfterTimes = [math]::Max(1, $remainingSpace - $spacingBeforeTimes)
						
						# Write left part (mJig title)
						Write-Host "  mJig(" -NoNewline -ForegroundColor Magenta
						Write-Host "`u{1F400}" -NoNewline -ForegroundColor White
						Write-Host ")" -NoNewline -ForegroundColor Magenta
						# Add DEBUGMODE indicator if in debug mode
						if ($DebugMode) {
							Write-Host " - DEBUGMODE" -NoNewline -ForegroundColor Red
						}
						
						# Add spacing before times
						write-host (" " * $spacingBeforeTimes) -NoNewLine
						
						# Write times (Current first, then End)
						Write-Host "Current" -NoNewline -ForegroundColor Yellow
						Write-Host "`u{23F3}/" -NoNewline -ForegroundColor White
						Write-Host "$currentTime" -NoNewline -ForegroundColor Green
						Write-Host " ➣  " -NoNewline
						Write-Host "End" -NoNewline -ForegroundColor Yellow
						Write-Host "`u{23F3}/" -NoNewline -ForegroundColor White
						Write-Host "$endTimeDisplay" -NoNewline -ForegroundColor Green
						
						# Add spacing after times and write view tag aligned to the right
						write-host (" " * $spacingAfterTimes) -NoNewLine
						if ($Output -eq "full") {
							write-host " Full" -ForeGroundColor Magenta -NoNewline
						} else {
							write-host " Minimum" -ForeGroundColor Magenta -NoNewline
						}
						
						# Add 2 spaces for right margin (view tag should end 2 spaces from right edge)
						write-host "  " -NoNewline
						
						# Clear any remaining characters on the line
						$currentX = $Host.UI.RawUI.CursorPosition.x
						if ($currentX -lt $HostWidth) {
							for ($i = $currentX; $i -lt $HostWidth; $i++) {
								write-host " " -NoNewline
							}
						}
					}
				}
				$Outputline++

				# Output Line Spacer
				$t = $true
				try {
					[Console]::SetCursorPosition(0, $Outputline)
				} catch {
					$t = $false
				} finally {
					if ($t) {
						for ($i = $Host.UI.RawUI.CursorPosition.x; $i -lt $HostWidth; $i++) {
							Write-Host " " -NoNewLine
							write-host ("─" * ($HostWidth - 2)) -ForegroundColor White -NoNewline
							for ($i = $Host.UI.RawUI.CursorPosition.x; $i -lt $HostWidth; $i++) {
								write-host " " -NoNewline
							}
						}
					}
				}
				$outputLine++

			# Only render console if not skipping updates (prevents stutter during mouse movement)
			if (-not $skipConsoleUpdate) {
				# Calculate view-dependent variables INSIDE the skip check to ensure they use current $Output value
				# This prevents stale view calculations when console updates resume after mouse movement
				$boxWidth = 50  # Width for stats box
				$boxPadding = 2  # Padding around box (1 space on each side)
				$verticalSeparatorWidth = 3  # " │ " = 3 characters
				$showStatsBox = ($Output -eq "full" -and $HostWidth -ge ($boxWidth + $boxPadding + $verticalSeparatorWidth + 50))  # Need at least 50 chars for logs
				$logWidth = if ($showStatsBox) { $HostWidth - $boxWidth - $boxPadding - $verticalSeparatorWidth } else { $HostWidth }  # Reserve space for box + padding + separator
				
				# Pre-calculate key text splitting for full view
				$keysFirstLine = ""
				$keysSecondLine = ""
				if ($showStatsBox) {
					if ($PreviousIntervalKeys.Count -gt 0) {
						# Filter out empty/null values to prevent leading commas
						$filteredKeys = $PreviousIntervalKeys | Where-Object { $_ -and $_.ToString().Trim() -ne "" }
						$keysText = if ($filteredKeys.Count -gt 0) { ($filteredKeys -join ", ") } else { "" }
						# Split into two lines if needed (only if we have text)
						if ($keysText -and $keysText.Length -gt ($boxWidth - 2)) {
							# Try to split at a comma
							$splitPos = $keysText.LastIndexOf(", ", ($boxWidth - 2))
							if ($splitPos -gt 0) {
								$keysFirstLine = $keysText.Substring(0, $splitPos)
								$keysSecondLine = $keysText.Substring($splitPos + 2)
								# Truncate second line if still too long
								if ($keysSecondLine.Length -gt ($boxWidth - 2)) {
									$keysSecondLine = $keysSecondLine.Substring(0, ($boxWidth - 5)) + "..."
								}
							} else {
								# No comma found, just truncate
								$keysFirstLine = $keysText.Substring(0, ($boxWidth - 5)) + "..."
								$keysSecondLine = ""
							}
						} elseif ($keysText) {
							$keysFirstLine = $keysText
							$keysSecondLine = ""
						}
					}
				}
				
				for ($i = 0; $i -lt $Rows; $i++) {
					$t = $true
					try {
						[Console]::SetCursorPosition(0, $Outputline)
					} catch {
						$t = $false
					} finally {
						if ($t) {
						# Always render all rows - check if we have a log entry with content for this row
						# We ensure logArray always has $Rows entries, so $i should always be < logArray.Count
						# Define availableWidth here so it's available in both if and else blocks
						$availableWidth = $logWidth
						
						# Safety check: ensure logArray has an entry for this index
						$hasLogEntry = ($i -lt $LogArray.Count -and $null -ne $LogArray[$i] -and $null -ne $LogArray[$i].components)
						$hasContent = ($hasLogEntry -and $LogArray[$i].components.Count -gt 0)
						
						if ($hasContent) {
							# Format log line based on available width with priority
							$formattedLine = ""
							$useShortTimestamp = $false
							
							# Calculate total length with full components (accounting for 2-space indent)
							$fullLength = 2  # Start with 2 for leading spaces
							foreach ($component in $LogArray[$i].components) {
								$fullLength += $component.text.Length
							}
							
							# If full length exceeds width, start using shortened timestamp
							if ($fullLength -gt $availableWidth) {
								$useShortTimestamp = $true
								# Recalculate with short timestamp
								$shortLength = 2  # Start with 2 for leading spaces
								foreach ($component in $LogArray[$i].components) {
									if ($component.priority -eq 1) {
										$shortLength += $component.shortText.Length
									} else {
										$shortLength += $component.text.Length
									}
								}
								$fullLength = $shortLength
							}
							
							# Build line with priority-based truncation (accounting for 2-space indent)
							$formattedLine = "  "  # Add 2 leading spaces
							$remainingWidth = $availableWidth - 2  # Subtract 2 for leading spaces
							foreach ($component in $LogArray[$i].components | Sort-Object priority) {
								$componentText = if ($component.priority -eq 1 -and $useShortTimestamp) {
									$component.shortText
								} else {
									$component.text
								}
								
								# Check if we have room for this component
								if ($componentText.Length -le $remainingWidth) {
									$formattedLine += $componentText
									$remainingWidth -= $componentText.Length
								} else {
									# Truncate this component if it's the last one and we have some room
									if ($remainingWidth -gt 3) {
										$formattedLine += $componentText.Substring(0, $remainingWidth - 3) + "..."
									}
									break
								}
							}
							
							# Clear the line first, then write the new content
							# Pad with spaces to clear any leftover characters and ensure exact width
							# Truncate if longer, pad if shorter to ensure exactly $availableWidth characters
							$truncatedLine = if ($formattedLine.Length -gt $availableWidth) {
								$formattedLine.Substring(0, $availableWidth)
							} else {
								$formattedLine
							}
							$paddedLine = $truncatedLine.PadRight($availableWidth)
							write-host $paddedLine -NoNewline
							
							# Draw vertical separator (always separate from box, for all rows)
							if ($showStatsBox) {
								write-host " │ " -NoNewline -ForegroundColor White
							}
							
							# Draw stats box in full view (with padding so it doesn't touch white lines)
							if ($showStatsBox) {
								# Add space before box (so it doesn't touch vertical separator)
								write-host " " -NoNewline
								
								# Draw box content
								if ($i -eq 0) {
									# Top border
									write-host "┌" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┐" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq 1) {
									# Header row
									$boxHeader = "Stats"
									write-host "│" -NoNewline -ForegroundColor Cyan
									write-host $boxHeader.PadRight($boxWidth - 2) -NoNewline -ForegroundColor Cyan
									write-host "│" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq 2) {
									# Separator row
									write-host "├" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┤" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq $Rows - 5) {
									# Fifth to last row - separator before keys (moved up one line)
									write-host "├" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┤" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq $Rows - 4) {
									# Fourth to last row - show detected keys label (moved up one line)
									write-host "│" -NoNewline -ForegroundColor Cyan
									write-host "Detected Inputs:".PadRight($boxWidth - 2) -NoNewline -ForegroundColor Cyan
									write-host "│" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq $Rows - 3) {
									# Third to last row - show first line of keys
									if ($PreviousIntervalKeys.Count -gt 0) {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host $keysFirstLine.PadRight($boxWidth - 2) -NoNewline -ForegroundColor Yellow
										write-host "│" -NoNewline -ForegroundColor Cyan
									} else {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host "(none)".PadRight($boxWidth - 2) -NoNewline -ForegroundColor DarkGray
										write-host "│" -NoNewline -ForegroundColor Cyan
									}
								} elseif ($i -eq $Rows - 2) {
									# Second to last row - show second line of keys if needed
									if ($PreviousIntervalKeys.Count -gt 0 -and $keysSecondLine -ne "") {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host $keysSecondLine.PadRight($boxWidth - 2) -NoNewline -ForegroundColor Yellow
										write-host "│" -NoNewline -ForegroundColor Cyan
									} else {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host (" " * ($boxWidth - 2)) -NoNewline
										write-host "│" -NoNewline -ForegroundColor Cyan
									}
								} elseif ($i -eq $Rows - 1) {
									# Last row - bottom border of box
									write-host "└" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┘" -NoNewline -ForegroundColor Cyan
								} else {
									# Empty rows in box
									write-host "│" -NoNewline -ForegroundColor Cyan
									write-host (" " * ($boxWidth - 2)) -NoNewline
									write-host "│" -NoNewline -ForegroundColor Cyan
								}
								
								# Add space after box (so it doesn't touch right edge)
								write-host " " -NoNewline
							}
							
							# Move to next line
							write-host ""
						} else {
							# Clear the line with spaces - ensure exactly $availableWidth characters
							# This must match the width used for log entries to maintain alignment
							# Use PadRight to ensure exact width, same as log entries
							$emptyLine = "".PadRight($availableWidth)
							write-host $emptyLine -NoNewline
							
							# Draw vertical separator (always separate from box, for all rows)
							if ($showStatsBox) {
								write-host " │ " -NoNewline -ForegroundColor White
							}
							
							# Draw stats box in full view (for empty log lines, with padding)
							if ($showStatsBox) {
								# Add space before box (so it doesn't touch vertical separator)
								write-host " " -NoNewline
								if ($i -eq 0) {
									write-host "┌" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┐" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq 1) {
									write-host "│" -NoNewline -ForegroundColor Cyan
									write-host "Stats".PadRight($boxWidth - 2) -NoNewline -ForegroundColor Cyan
									write-host "│" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq 2) {
									write-host "├" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┤" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq $Rows - 5) {
									# Fifth to last row - separator before keys (moved up one line)
									write-host "├" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┤" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq $Rows - 4) {
									# Fourth to last row - show detected keys label (moved up one line)
									write-host "│" -NoNewline -ForegroundColor Cyan
									write-host "Detected Inputs:".PadRight($boxWidth - 2) -NoNewline -ForegroundColor Cyan
									write-host "│" -NoNewline -ForegroundColor Cyan
								} elseif ($i -eq $Rows - 3) {
									# Third to last row - show first line of keys
									if ($PreviousIntervalKeys.Count -gt 0) {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host $keysFirstLine.PadRight($boxWidth - 2) -NoNewline -ForegroundColor Yellow
										write-host "│" -NoNewline -ForegroundColor Cyan
									} else {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host "(none)".PadRight($boxWidth - 2) -NoNewline -ForegroundColor DarkGray
										write-host "│" -NoNewline -ForegroundColor Cyan
									}
								} elseif ($i -eq $Rows - 2) {
									# Second to last row - show second line of keys if needed
									if ($PreviousIntervalKeys.Count -gt 0 -and $keysSecondLine -ne "") {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host $keysSecondLine.PadRight($boxWidth - 2) -NoNewline -ForegroundColor Yellow
										write-host "│" -NoNewline -ForegroundColor Cyan
									} else {
										write-host "│" -NoNewline -ForegroundColor Cyan
										write-host (" " * ($boxWidth - 2)) -NoNewline
										write-host "│" -NoNewline -ForegroundColor Cyan
									}
								} elseif ($i -eq $Rows - 1) {
									# Last row - bottom border of box
									write-host "└" -NoNewline -ForegroundColor Cyan
									write-host ("─" * ($boxWidth - 2)) -NoNewline -ForegroundColor Cyan
									write-host "┘" -NoNewline -ForegroundColor Cyan
								} else {
									write-host "│" -NoNewline -ForegroundColor Cyan
									write-host (" " * ($boxWidth - 2)) -NoNewline
									write-host "│" -NoNewline -ForegroundColor Cyan
								}
							}
							
							write-host ""
						}
					}
				}
				$outputLine++
			}
			}  # End of skipConsoleUpdate check

			# Output bottom separator (only if not skipping console updates)
			if ($Output -ne "hidden" -and -not $skipConsoleUpdate) {
				# Calculate if we should show stats box (full view, wide enough window)
				$boxWidth = 50  # Width for stats box
				$boxPadding = 2  # Padding around box (1 space on each side)
				$verticalSeparatorWidth = 3  # " │ " = 3 characters
				$showStatsBox = ($Output -eq "full" -and $HostWidth -ge ($boxWidth + $boxPadding + $verticalSeparatorWidth + 50))  # Need at least 50 chars for logs
				$logWidth = if ($showStatsBox) { $HostWidth - $boxWidth - $boxPadding - $verticalSeparatorWidth } else { $HostWidth }  # Reserve space for box + padding + separator
				
				$t = $true
				try {
					[Console]::SetCursorPosition(0, $Outputline)
				} catch {
					$t = $false
				} finally {
					if ($t) {
						# Draw continuous white separator line across full width
						# This separates the log section from the menu section
						# No vertical separator on this line - it's a continuous horizontal line
						# Match the pattern from the top separator exactly (one space on each side)
						for ($i = $Host.UI.RawUI.CursorPosition.x; $i -lt $HostWidth; $i++) {
							Write-Host " " -NoNewLine
							write-host ("─" * ($HostWidth - 2)) -ForegroundColor White -NoNewline
							for ($i = $Host.UI.RawUI.CursorPosition.x; $i -lt $HostWidth; $i++) {
								write-host " " -NoNewline
							}
						}
					}
				}
				$outputLine++

				## Menu Options ##
				$t = $true
				try {
					[Console]::SetCursorPosition(0, $Outputline)
				} catch {
					$t = $false
				} finally {
					if ($t) {
						# Dynamic menu truncation - similar to log truncation
						# Define menu items with different display formats
						# Format 0: Full with icons and pipes
						# Format 1: Without icons and pipes
						# Format 2: Shortened to just hotkey word
						
						# Determine menu text based on whether end time is set
						$timeMenuText = if ($endTimeInt -eq -1 -or [string]::IsNullOrEmpty($endTimeStr)) {
							"set_end_(t)ime"
						} else {
							"change_end_(t)ime"
						}
						
						# Build menu items - include modify movement only in full view
						$menuItemsList = @(
							@{
								full = "⏳|$timeMenuText"
								noIcons = $timeMenuText
								short = "(t)ime"
							},
							@{
								full = "🔄|toggle_(v)iew"
								noIcons = "toggle_(v)iew"
								short = "(v)iew"
							},
							@{
								full = "🔒|(h)ide_output"
								noIcons = "(h)ide_output"
								short = "(h)ide"
							}
						)
						
						# Add modify movement only in full view
						if ($Output -eq "full") {
							$menuItemsList += @{
								full = "⚙️|modify_(m)ovement"
								noIcons = "modify_(m)ovement"
								short = "(m)ove"
							}
						}
						
						# Always add quit at the end
						$menuItemsList += @{
							full = "❌|(q)uit"
							noIcons = "(q)uit"
							short = "(q)uit"
						}
						
						$menuItems = $menuItemsList
						
						# Calculate widths for each format (emojis count as 2 chars in console)
						# Format 0 (full): includes emoji (2) + pipe (1) + text
						# Format 1 (noIcons): just text
						# Format 2 (short): just hotkey word
						
						$format0Width = 2  # Leading spaces
						$format1Width = 2  # Leading spaces
						$format2Width = 2  # Leading spaces
						
						foreach ($item in $menuItems) {
							# Format 0: emoji (2 display chars) + pipe (1) + text length (after removing emoji and pipe)
							$textPart = $item.full -replace "^.+\|", ""  # Remove everything up to and including the pipe
							$format0Width += 2 + 1 + $textPart.Length + 2  # +2 for spacing between items
							# Format 1: just text length
							$format1Width += $item.noIcons.Length + 2  # +2 for spacing between items
							# Format 2: just short text length (single space between items)
							$format2Width += $item.short.Length + 1  # +1 for spacing between items
						}
						
						# Add 2 spaces for right margin
						$format0Width += 2
						$format1Width += 2
						$format2Width += 2
						
						# Determine which format to use based on available width
						$menuFormat = 0  # Default to full format
						if ($HostWidth -lt $format0Width) {
							if ($HostWidth -lt $format1Width) {
								$menuFormat = 2  # Use short format
							} else {
								$menuFormat = 1  # Use no-icons format
							}
						}
						
						# Calculate spacing for right-aligned quit (only for format 0 and 1)
						$spacing = 0
						if ($menuFormat -lt 2) {
							# Calculate total width of all items except quit
							$totalMenuWidth = 2  # Leading spaces
							$itemsBeforeQuit = $menuItems.Count - 1  # All items except quit
							for ($i = 0; $i -lt $itemsBeforeQuit; $i++) {
								$item = $menuItems[$i]
								if ($menuFormat -eq 0) {
									# Account for emoji (2 display chars) and pipe (1 char) + text
									$textPart = $item.full -replace "^.+\|", ""  # Remove everything up to and including the pipe
									$itemWidth = 2 + 1 + $textPart.Length
								} else {
									$itemText = $item.noIcons
									$itemWidth = $itemText.Length
								}
								$totalMenuWidth += $itemWidth + 2  # +2 for spacing between items
							}
							# Calculate quit item width (last item)
							$quitItem = $menuItems[$menuItems.Count - 1]
							if ($menuFormat -eq 0) {
								$textPart = $quitItem.full -replace "^.+\|", ""
								$quitWidth = 2 + 1 + $textPart.Length
							} else {
								$quitWidth = $quitItem.noIcons.Length
							}
							# Spacing = total width - items before quit - quit item - right margin
							$spacing = [math]::Max(1, $HostWidth - $totalMenuWidth - $quitWidth - 2)  # -2 for right margin
						}
						
						# Write menu items
						write-host "  " -NoNewLine
						
						# Initialize menu item bounds tracking
						$script:MenuItemsBounds = @()
						$menuY = $Outputline
						$currentMenuX = 2  # Start after "  " prefix
						
						# Write all items except quit
						$itemsBeforeQuit = $menuItems.Count - 1
						for ($i = 0; $i -lt $itemsBeforeQuit; $i++) {
							$item = $menuItems[$i]
							$itemStartX = $currentMenuX
							
							# Determine which text to use
							if ($menuFormat -eq 0) {
								$itemText = $item.full
							} elseif ($menuFormat -eq 1) {
								$itemText = $item.noIcons
							} else {
								$itemText = $item.short
							}
							
							# Calculate item width (accounting for emojis which display as 2 chars)
							$itemDisplayWidth = 0
							if ($menuFormat -eq 0) {
								$parts = $itemText -split "\|", 2
								if ($parts.Count -eq 2) {
									$emoji = $parts[0]
									$text = $parts[1]
									$itemDisplayWidth = 2 + 1 + $text.Length  # emoji (2) + pipe (1) + text
								}
							} else {
								$itemDisplayWidth = $itemText.Length
							}
							
							# Parse and write the menu item
							if ($menuFormat -eq 0) {
								# Full format: emoji|text
								$parts = $itemText -split "\|", 2
								if ($parts.Count -eq 2) {
									$emoji = $parts[0]
									$text = $parts[1]
									write-host $emoji -NoNewline
									write-host "|" -NoNewline -ForegroundColor White
									# Parse text for colors (look for parentheses with letter inside)
									$textParts = $text -split "([()])"
									for ($j = 0; $j -lt $textParts.Count; $j++) {
										$part = $textParts[$j]
										# Check if this is a hotkey pattern: (letter)
										if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
											# Hotkey letter - parentheses green, letter yellow
											write-host "(" -NoNewline -ForegroundColor Green
											write-host $textParts[$j + 1] -NoNewline -ForegroundColor Yellow
											write-host ")" -NoNewline -ForegroundColor Green
											$j += 2  # Skip the letter and closing paren
										} elseif ($part -ne "") {
											# Regular text
											write-host $part -NoNewline -ForegroundColor Green
										}
									}
								}
							} else {
								# No-icons or short format: just text
								# Parse text for colors (look for parentheses with letter inside)
								$textParts = $itemText -split "([()])"
								for ($j = 0; $j -lt $textParts.Count; $j++) {
									$part = $textParts[$j]
									# Check if this is a hotkey pattern: (letter)
									if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
										# Hotkey letter - parentheses green, letter yellow
										write-host "(" -NoNewline -ForegroundColor Green
										write-host $textParts[$j + 1] -NoNewline -ForegroundColor Yellow
										write-host ")" -NoNewline -ForegroundColor Green
										$j += 2  # Skip the letter and closing paren
									} elseif ($part -ne "") {
										# Regular text
										write-host $part -NoNewline -ForegroundColor Green
									}
								}
							}
							
							# Store menu item bounds
							$itemEndX = $itemStartX + $itemDisplayWidth - 1
							# Extract hotkey from item (look for pattern like "(t)" in the text)
							$hotkeyMatch = $itemText -match "\(([a-z])\)"
							$hotkey = if ($hotkeyMatch) { $matches[1] } else { $null }
							$script:MenuItemsBounds += @{
								startX = $itemStartX
								endX = $itemEndX
								y = $menuY
								hotkey = $hotkey
								index = $i
							}
							
							# Update current X position
							$currentMenuX = $itemEndX + 1
							
							# Add spacing between items
							if ($menuFormat -eq 2) {
								# Short format: use single space
								write-host " " -NoNewline
								$currentMenuX += 1
							} else {
								# Full or no-icons format: use two spaces
								write-host "  " -NoNewline
								$currentMenuX += 2
							}
						}
						
						# Add spacing before quit (only for format 0 and 1)
						if ($menuFormat -lt 2) {
							write-host (" " * $spacing) -NoNewLine
							$currentMenuX += $spacing
						}
						
						# Write quit item (last item)
						$quitItem = $menuItems[$menuItems.Count - 1]
						$quitStartX = $currentMenuX
						if ($menuFormat -eq 0) {
							$itemText = $quitItem.full
						} elseif ($menuFormat -eq 1) {
							$itemText = $quitItem.noIcons
						} else {
							$itemText = $quitItem.short
						}
						
						# Calculate quit item width
						$quitDisplayWidth = 0
						if ($menuFormat -eq 0) {
							$parts = $itemText -split "\|", 2
							if ($parts.Count -eq 2) {
								$emoji = $parts[0]
								$text = $parts[1]
								$quitDisplayWidth = 2 + 1 + $text.Length  # emoji (2) + pipe (1) + text
							}
						} else {
							$quitDisplayWidth = $itemText.Length
						}
						
						# Parse and write the quit menu item
						if ($menuFormat -eq 0) {
							# Full format: emoji|text
							$parts = $itemText -split "\|", 2
							if ($parts.Count -eq 2) {
								$emoji = $parts[0]
								$text = $parts[1]
								write-host $emoji -NoNewline
								write-host "|" -NoNewline -ForegroundColor White
								# Parse text for colors (look for parentheses with letter inside)
								$textParts = $text -split "([()])"
								for ($j = 0; $j -lt $textParts.Count; $j++) {
									$part = $textParts[$j]
									# Check if this is a hotkey pattern: (letter)
									if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
										# Hotkey letter - parentheses green, letter yellow
										write-host "(" -NoNewline -ForegroundColor Green
										write-host $textParts[$j + 1] -NoNewline -ForegroundColor Yellow
										write-host ")" -NoNewline -ForegroundColor Green
										$j += 2  # Skip the letter and closing paren
									} elseif ($part -ne "") {
										# Regular text
										write-host $part -NoNewline -ForegroundColor Green
									}
								}
							}
						} else {
							# No-icons or short format: just text
							# Parse text for colors (look for parentheses with letter inside)
							$textParts = $itemText -split "([()])"
							for ($j = 0; $j -lt $textParts.Count; $j++) {
								$part = $textParts[$j]
								# Check if this is a hotkey pattern: (letter)
								if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
									# Hotkey letter - parentheses green, letter yellow
									write-host "(" -NoNewline -ForegroundColor Green
									write-host $textParts[$j + 1] -NoNewline -ForegroundColor Yellow
									write-host ")" -NoNewline -ForegroundColor Green
									$j += 2  # Skip the letter and closing paren
								} elseif ($part -ne "") {
									# Regular text
									write-host $part -NoNewline -ForegroundColor Green
								}
							}
						}
						
						# Store quit item bounds
						$quitEndX = $quitStartX + $quitDisplayWidth - 1
						$quitHotkeyMatch = $itemText -match "\(([a-z])\)"
						$quitHotkey = if ($quitHotkeyMatch) { $matches[1] } else { $null }
						$script:MenuItemsBounds += @{
							startX = $quitStartX
							endX = $quitEndX
							y = $menuY
							hotkey = $quitHotkey
							index = $menuItems.Count - 1
						}
						
						# Add 2 spaces for right margin
						write-host "  " -NoNewline
						
						# Clear any remaining characters on the line to ensure proper display
						$currentX = $Host.UI.RawUI.CursorPosition.x
						if ($currentX -lt $HostWidth) {
							for ($i = $currentX; $i -lt $HostWidth; $i++) {
								write-host " " -NoNewline
							}
						}
					}
				}
				$Outputline++
				}
			} elseif ($Output -eq "hidden") {
				# Hidden view: show only "TIME | running..." on a single line that updates in place
				# Update the timestamp on every iteration to show the latest time
				# Only render if not skipping console updates (to prevent stutter during mouse movement)
				if (-not $skipConsoleUpdate) {
					$t = $true
					try {
						[Console]::SetCursorPosition(0, 0)
					} catch {
						$t = $false
					} finally {
						if ($t) {
							# Clear the line first
							for ($i = 0; $i -lt $HostWidth; $i++) {
								write-host " " -NoNewline
							}
							# Position at start of line and write the timestamp and running message
							[Console]::SetCursorPosition(0, 0)
							$timeStr = $date.ToString("HH:mm:ss")
							write-host "$timeStr | running..." -NoNewline
							# Clear any remaining characters on the line
							$currentX = $Host.UI.RawUI.CursorPosition.x
							if ($currentX -lt $HostWidth) {
								for ($i = $currentX; $i -lt $HostWidth; $i++) {
									write-host " " -NoNewline
								}
							}
						}
					}
				}
			}
			# If skipConsoleUpdate is true and Output is not "hidden", don't render anything (prevents stutter)
			
			# Check if end time reached (only if end time is set)
			if ($endTimeInt -ne -1 -and -not [string]::IsNullOrEmpty($end)) {
				$current = Get-Date -Format "MMddHHmm"
				if ($current -ge $end) {
					$time = $true
				}
			}
		} until ($time -eq $true)
		
		# End message
		if ($Output -ne "hidden") {
			[Console]::SetCursorPosition(0, $Outputline)
			Write-Host "       END TIME REACHED: " -NoNewline -ForegroundColor Red
			Write-Host "Stopping " -NoNewline
			Write-Host "mJig"
			write-host
		}
}

# Uncomment the line below to run the function when script is executed directly
# Start-mJig -Output full
