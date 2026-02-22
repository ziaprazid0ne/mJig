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

	param(
		[Parameter(Mandatory = $false)] 
		[ValidateSet("min", "full", "hidden", "dib")]
		[string]$Output = "min",
		[Parameter(Mandatory = $false)]
		[switch]$DebugMode,
		[Parameter(Mandatory = $false)]
		[switch]$Diag,
		[Parameter(Mandatory = $false)] 
		[string]$EndTime = "0",  # 0 = no end time, otherwise 4-digit 24 hour format (e.g., 1807 = 6:07 PM)
		[Parameter(Mandatory = $false)]
		[int]$EndVariance = 0,  # Variance in minutes to randomly add/subtract from EndTime to avoid overly consistent end times. Only applies if EndTime is specified (not 0).
		[Parameter(Mandatory = $false)]
		[double]$IntervalSeconds = 2,  # sets the base interval time between refreshes
		[Parameter(Mandatory = $false)]
		[double]$IntervalVariance = 2,  # Sets the maximum random plus and minus variance in seconds each refresh
		[Parameter(Mandatory = $false)]
		[double]$MoveSpeed = 0.5,  # Base movement speed in seconds (time to complete movement)
		[Parameter(Mandatory = $false)]
		[double]$MoveVariance = 0.2,  # Maximum random variance in movement speed (in seconds)
		[Parameter(Mandatory = $false)]
		[double]$TravelDistance = 100,  # Base travel distance in pixels
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
	$SkipUpdate = $false
	$PreviousView = $null  # Store the view before hiding to restore it later
	$PosUpdate = $false
	$LogArray = @()
	$HostWidth = 0
	$HostHeight = 0
	$OutputLine = 0
	$LastMovementTime = $null
	$LastMovementDurationMs = 0  # Track duration of last movement in milliseconds
	$LastSimulatedKeyPress = $null  # Track when we last sent a simulated key press
	$LastAutomatedMouseMovement = $null  # Track when we last performed automated mouse movement
	$LastUserInputTime = $null  # Track when user input was last detected (for auto-resume delay timer)

	$PreviousIntervalKeys = @()  # Track keys pressed in previous interval for display
	$LastResizeDetection = $null  # Track when we last detected a resize
	$PendingResize = $null  # Track pending resize to throttle redraws
	$ResizeThrottleMs = 1500  # Wait 2000ms after window stops resizing before processing resize
	$ResizeClearedScreen = $false  # Track if we've cleared the screen at the start of a resize
	$LastResizeLogoTime = $null  # Track when we last drew the resize logo
	$script:LoopIteration = 0  # Track loop iterations for diagnostics
	$script:lastInputCheckTime = $null  # Track when we last logged input check (for debug mode)
	$script:DialogButtonClick = $null  # Track dialog button clicks detected from main loop ("Update" or "Cancel")
	
	# Performance: Cache for reflection method lookups
	$script:MethodCache = @{}
	
	# Note: Screen bounds are cached later after System.Windows.Forms is loaded
	$script:ScreenWidth = $null
	$script:ScreenHeight = $null
	$script:DialogButtonBounds = $null  # Store dialog button bounds when dialog is open {buttonRowY, updateStartX, updateEndX, cancelStartX, cancelEndX}
	$script:LastClickLogTime = $null  # Track when we last logged a click to prevent duplicate logs
	$script:WindowTitle = "mJig - mJigg"  # Fixed window title (same for all instances to enable duplicate detection)
	$script:MenuClickHotkey = $null  # Menu item hotkey triggered by mouse click
	$script:RenderQueue = New-Object 'System.Collections.Generic.List[hashtable]'
	
	# Box-drawing characters (using Unicode code points to avoid encoding issues)
	$script:BoxTopLeft = [char]0x250C      # ┌
	$script:BoxTopRight = [char]0x2510     # ┐
	$script:BoxBottomLeft = [char]0x2514   # └
	$script:BoxBottomRight = [char]0x2518  # ┘
	$script:BoxHorizontal = [char]0x2500   # ─
	$script:BoxVertical = [char]0x2502     # │
	$script:BoxVerticalRight = [char]0x251C # ├
	$script:BoxVerticalLeft = [char]0x2524  # ┤
	
	# ============================================================================
	# Theme Colors - Centralized color definitions for easy customization
	# ============================================================================
	
	# Menu Bar (bottom of main screen)
	$script:MenuButtonBg = "DarkBlue"
	$script:MenuButtonText = "White"
	$script:MenuButtonHotkey = "Green"
	$script:MenuButtonPipe = "White"
	
	# Main Display - Header
	$script:HeaderAppName = "Magenta"
	$script:HeaderIcon = "White"
	$script:HeaderStatus = "Green"
	$script:HeaderPaused = "Yellow"
	$script:HeaderTimeLabel = "Yellow"
	$script:HeaderTimeValue = "Green"
	$script:HeaderViewTag = "Magenta"
	$script:HeaderSeparator = "White"
	
	# Main Display - Stats Box
	$script:StatsBoxBorder = "Cyan"
	$script:StatsBoxTitle = "Cyan"
	$script:StatsLabel = "White"
	$script:StatsValue = "Yellow"
	$script:StatsValueGood = "Green"
	$script:StatsValueBad = "Red"
	
	# Quit Dialog
	$script:QuitDialogBg = "DarkMagenta"
	$script:QuitDialogShadow = "DarkMagenta"
	$script:QuitDialogBorder = "White"
	$script:QuitDialogTitle = "Yellow"
	$script:QuitDialogText = "White"
	$script:QuitDialogButtonBg = "Magenta"
	$script:QuitDialogButtonText = "White"
	$script:QuitDialogButtonHotkey = "Yellow"
	
	# Set Time Dialog
	$script:TimeDialogBg = "DarkBlue"
	$script:TimeDialogShadow = "DarkBlue"
	$script:TimeDialogBorder = "White"
	$script:TimeDialogTitle = "Yellow"
	$script:TimeDialogText = "White"
	$script:TimeDialogButtonBg = "Blue"
	$script:TimeDialogButtonText = "White"
	$script:TimeDialogButtonHotkey = "Yellow"
	$script:TimeDialogFieldBg = "Blue"
	$script:TimeDialogFieldText = "White"
	
	# Modify Movement Dialog
	$script:MoveDialogBg = "DarkBlue"
	$script:MoveDialogShadow = "DarkBlue"
	$script:MoveDialogBorder = "White"
	$script:MoveDialogTitle = "Yellow"
	$script:MoveDialogSectionTitle = "Yellow"
	$script:MoveDialogText = "White"
	$script:MoveDialogButtonBg = "Blue"
	$script:MoveDialogButtonText = "White"
	$script:MoveDialogButtonHotkey = "Yellow"
	$script:MoveDialogFieldBg = "Blue"
	$script:MoveDialogFieldText = "White"
	
	# Resize Screen
	$script:ResizeBoxBorder = "White"
	$script:ResizeLogoName = "Magenta"
	$script:ResizeLogoIcon = "White"
	$script:ResizeQuoteText = "White"
	
	# General UI
	$script:TextDefault = "White"
	$script:TextMuted = "DarkGray"
	$script:TextHighlight = "Cyan"
	$script:TextSuccess = "Green"
	$script:TextWarning = "Yellow"
	$script:TextError = "Red"
	
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
			Write-Host "[DEBUG] Set window title: $($Host.UI.RawUI.WindowTitle)" -ForegroundColor $script:TextHighlight
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to set window title: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
		}
	}
	
	# Check for duplicate windows - FAIL if another instance is running
	if ($DebugMode) {
		Write-Host "[DEBUG] Checking for duplicate mJig instances..." -ForegroundColor $script:TextHighlight
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
			Write-Host "  [WARN] Could not check for duplicates: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
		}
	}
	
	# If duplicate found, exit with error
	if ($duplicateFound) {
		Write-Host ""
		Write-Host "ERROR: Another instance of mJig is already running!" -ForegroundColor $script:TextError
		Write-Host "  Process ID: $duplicateProcessId" -ForegroundColor $script:TextError
		Write-Host "  Process Name: $duplicateProcessName" -ForegroundColor $script:TextError
		Write-Host "  Window Title: $script:WindowTitle" -ForegroundColor $script:TextError
		Write-Host ""
		Write-Host "Please close the other instance before starting a new one." -ForegroundColor $script:TextWarning
		Write-Host ""
		exit 1
	} else {
		if ($DebugMode) {
			Write-Host "  [OK] No duplicate instances found" -ForegroundColor $script:TextSuccess
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
		Write-Host "Initialization Debug" -ForegroundColor $script:HeaderAppName
		Write-Host ""
		Write-Host "[DEBUG] Initializing console..." -ForegroundColor $script:TextHighlight
		# Window title already set above, just log it
		Write-Host "[DEBUG] Window title: $($Host.UI.RawUI.WindowTitle)" -ForegroundColor $script:TextHighlight
		Write-Host "[DEBUG] DebugMode is ENABLED - click detection will be logged" -ForegroundColor $script:TextWarning
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
						Write-Host "  [OK] Quick Edit Mode disabled, Mouse Input enabled" -ForegroundColor $script:TextSuccess
					}
				} else {
					if ($DebugMode) {
						Write-Host "  [WARN] Failed to set console mode (SetConsoleMode failed)" -ForegroundColor $script:TextWarning
					}
				}
			} else {
				if ($DebugMode) {
					Write-Host "  [WARN] Failed to disable Quick Edit Mode (GetConsoleMode failed)" -ForegroundColor $script:TextWarning
				}
			}
		} else {
			if ($DebugMode) {
				Write-Host "  [WARN] Failed to disable Quick Edit Mode (could not load Win32 API)" -ForegroundColor $script:TextWarning
			}
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to disable Quick Edit Mode: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
		}
	}
	
	# Enable VT100 processing on stdout for ANSI escape sequence rendering
	try {
		if ($type) {
			$STD_OUTPUT_HANDLE = -11
			$hStdOut = $type::GetStdHandle($STD_OUTPUT_HANDLE)
			$outMode = 0
			if ($type::GetConsoleMode($hStdOut, [ref]$outMode)) {
				$ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
				$newOutMode = $outMode -bor $ENABLE_VIRTUAL_TERMINAL_PROCESSING
				if ($type::SetConsoleMode($hStdOut, $newOutMode)) {
					if ($DebugMode) {
						Write-Host "  [OK] VT100 processing enabled on stdout" -ForegroundColor $script:TextSuccess
					}
				} else {
					if ($DebugMode) {
						Write-Host "  [WARN] Failed to enable VT100 processing" -ForegroundColor $script:TextWarning
					}
				}
			}
		}
		[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
		if ($DebugMode) {
			Write-Host "  [OK] Console output encoding set to UTF-8" -ForegroundColor $script:TextSuccess
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] VT100/UTF-8 setup: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
		}
	}
	
	try {
		[Console]::Write("$([char]27)[?25l")
		$script:CursorVisible = $false
		if ($DebugMode) {
			Write-Host "  [OK] Console cursor hidden" -ForegroundColor $script:TextSuccess
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to hide cursor: $($_.Exception.Message)" -ForegroundColor $script:TextError
		}
	}
	
	# Capture Initial Buffer & Window Sizes (needed even for hidden mode)
	if ($DebugMode) {
		Write-Host "[DEBUG] Capturing console dimensions..." -ForegroundColor $script:TextHighlight
	}
	try {
		$pshost = Get-Host
		$pswindow = $pshost.UI.RawUI
		$newWindowSize = $pswindow.WindowSize
		$newBufferSize = $pswindow.BufferSize
		if ($DebugMode) {
			Write-Host "  [OK] Got console dimensions" -ForegroundColor $script:TextSuccess
			Write-Host "    Window Size: $($newWindowSize.Width)x$($newWindowSize.Height)" -ForegroundColor Gray
			Write-Host "    Buffer Size: $($newBufferSize.Width)x$($newBufferSize.Height)" -ForegroundColor Gray
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to get console dimensions: $($_.Exception.Message)" -ForegroundColor $script:TextError
		}
		throw  # Re-throw as this is critical
	}
	# Set vertical buffer to match window height, but let horizontal buffer be managed by PowerShell (for text zoom)
	try {
		$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($newBufferSize.Width, $newWindowSize.Height)
		$newBufferSize = $pswindow.BufferSize
		if ($DebugMode) {
			Write-Host "  [OK] Set buffer height to match window height" -ForegroundColor $script:TextSuccess
		}
	} catch {
		# If setting buffer size fails, continue with current buffer size
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to set buffer size: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
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
		Write-Host "[DEBUG] Initializing output array..." -ForegroundColor $script:TextHighlight
	}
	try {
		if ($Output -ne "hidden") {
			$LogArray = @()
			if ($DebugMode) {
				Write-Host "  [OK] Output mode: $Output" -ForegroundColor $script:TextSuccess
			}
		} else {
			if ($DebugMode) {
				Write-Host "  [OK] Output mode: hidden (no log array)" -ForegroundColor $script:TextSuccess
			}
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to initialize output array: $($_.Exception.Message)" -ForegroundColor $script:TextError
		}
		throw  # Re-throw as this is critical
	}

	###############################
	## Calculating the End Times ##
	###############################
	
	if ($DebugMode) {
		Write-Host "[DEBUG] Calculating end times..." -ForegroundColor $script:TextHighlight
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
				Write-Host "  [OK] No end time specified - script will run indefinitely" -ForegroundColor $script:TextSuccess
			}
		} elseif ($endTimeTrimmed.Length -eq 2) {
			# 2-digit input = hour on the hour (e.g., "12" = 1200, "00" = 0000)
			$hours = [int]$endTimeTrimmed
			if ($hours -ge 0 -and $hours -le 23) {
				$endTimeInt = $hours * 100  # Convert to HHmm format (e.g., 12 -> 1200)
				$endTimeStr = $endTimeInt.ToString().PadLeft(4, '0')
				if ($DebugMode) {
					Write-Host "  [OK] Parsed end time: $endTimeStr (hour on the hour)" -ForegroundColor $script:TextSuccess
				}
			} else {
				Write-Host "Error: Invalid hour format. Hours must be 00-23. Got: $EndTime" -ForegroundColor $script:TextError
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
					Write-Host "  [OK] Parsed end time: $endTimeStr" -ForegroundColor $script:TextSuccess
				}
			} else {
				if ($hours -gt 23) {
					Write-Host "Error: Invalid time format. Hours must be 00-23. Got: $EndTime" -ForegroundColor $script:TextError
				} elseif ($minutes -gt 59) {
					Write-Host "Error: Invalid time format. Minutes must be 00-59. Got: $EndTime" -ForegroundColor $script:TextError
				} else {
					Write-Host "Error: Invalid time format. Expected HHmm format (0000-2359). Got: $EndTime" -ForegroundColor $script:TextError
				}
				throw "Invalid time format: $EndTime"
			}
		} else {
			Write-Host "Error: Invalid time format. Expected '0' (none), 2-digit hour (00-23), or 4-digit HHmm (0000-2359). Got: $EndTime" -ForegroundColor $script:TextError
			throw "Invalid time format: $EndTime"
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [FAIL] Failed to parse endTime: $($_.Exception.Message)" -ForegroundColor $script:TextError
		}
		if ($_.Exception.Message -notmatch "Invalid time format") {
			Write-Host "Error: Invalid EndTime format: $EndTime" -ForegroundColor $script:TextError
		}
		throw
	}
	
	# Time format has already been validated in the try-catch block above
	# Proceed with initialization
		# Diagnostics - initialize folder and file paths
		$script:DiagEnabled = $Diag
		if ($script:DiagEnabled) {
			$script:DiagFolder = Join-Path $PSScriptRoot "_diag"
			if (-not (Test-Path $script:DiagFolder)) {
				New-Item -ItemType Directory -Path $script:DiagFolder -Force | Out-Null
			}
			$script:StartupDiagFile = Join-Path $script:DiagFolder "startup.txt"
			$script:SettleDiagFile = Join-Path $script:DiagFolder "settle.txt"
			$script:InputDiagFile = Join-Path $script:DiagFolder "input.txt"
			
			$diagTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
			"=== mJig Startup Diag: $diagTimestamp ===" | Out-File $script:StartupDiagFile
			"$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 1: Starting initialization" | Out-File $script:StartupDiagFile -Append
			"  Diag enabled, folder: $script:DiagFolder" | Out-File $script:StartupDiagFile -Append
			"=== mJig Settle Diag: $diagTimestamp ===" | Out-File $script:SettleDiagFile
			"$(Get-Date -Format 'HH:mm:ss.fff') - Settle diagnostics started" | Out-File $script:SettleDiagFile -Append
			"=== mJig Input Diag: $diagTimestamp ===" | Out-File $script:InputDiagFile
			"$(Get-Date -Format 'HH:mm:ss.fff') - Input diagnostics started (PeekConsoleInput + GetLastInputInfo)" | Out-File $script:InputDiagFile -Append
		}
		
		if ($DebugMode) {
			Write-Host "[DEBUG] Loading System.Windows.Forms assembly..." -ForegroundColor $script:TextHighlight
		}
		try {
			Add-Type -AssemblyName System.Windows.Forms
			if ($DebugMode) {
				Write-Host "  [OK] System.Windows.Forms loaded" -ForegroundColor $script:TextSuccess
			}
			# Cache screen bounds now that the assembly is loaded
			$script:ScreenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
			$script:ScreenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
			if ($DebugMode) {
				Write-Host "  [OK] Screen bounds cached: $($script:ScreenWidth) x $($script:ScreenHeight)" -ForegroundColor $script:TextSuccess
			}
		} catch {
			if ($DebugMode) {
				Write-Host "  [FAIL] Failed to load System.Windows.Forms: $($_.Exception.Message)" -ForegroundColor $script:TextError
			}
			throw  # Re-throw as this is critical
		}
		
		# Add Windows API for system-wide keyboard detection and key sending
		if ($DebugMode) {
			Write-Host "[DEBUG] Loading Windows API types..." -ForegroundColor $script:TextHighlight
		}
		# Check if types already exist and have the required methods
		$typesNeedReload = $false
		try {
			# Use a safer method to check if types exist without throwing errors
			$existingKeyboard = $null
			$existingMouse = $null
			
			# Try to get the types using Get-Type or by checking if they're loaded
			$allTypes = [System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.Namespace -eq 'mJiggAPI' }
			
			foreach ($type in $allTypes) {
				if ($type.Name -eq 'Keyboard') { $existingKeyboard = $type }
				if ($type.Name -eq 'Mouse') { $existingMouse = $type }
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
						Write-Host "  [WARN] Existing types found but missing required methods" -ForegroundColor $script:TextWarning
						Write-Host "  [WARN] Missing: GetCursorPos=$(-not $hasGetCursorPos), GetForegroundWindow=$(-not $hasGetForegroundWindow), FindWindow=$(-not $hasFindWindow), FindWindowByProcessId=$(-not $hasFindWindowByProcessId), FindWindowByTitlePattern=$(-not $hasFindWindowByTitlePattern)" -ForegroundColor $script:TextWarning
						Write-Host "  [WARN] Attempting reload (may fail if types already exist - restart PowerShell if needed)" -ForegroundColor $script:TextWarning
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
	
	[StructLayout(LayoutKind.Sequential)]
	public struct KEY_EVENT_RECORD {
		public int bKeyDown;
		public ushort wRepeatCount;
		public ushort wVirtualKeyCode;
		public ushort wVirtualScanCode;
		public char UnicodeChar;
		public uint dwControlKeyState;
	}
	
	[StructLayout(LayoutKind.Explicit)]
	public struct INPUT_RECORD {
		[FieldOffset(0)]
		public ushort EventType;
		[FieldOffset(4)]
		public MOUSE_EVENT_RECORD MouseEvent;
		[FieldOffset(4)]
		public KEY_EVENT_RECORD KeyEvent;
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
	
	[StructLayout(LayoutKind.Sequential)]
	public struct LASTINPUTINFO {
		public uint cbSize;
		public uint dwTime;
	}
	
	public class Keyboard {
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
		
		[DllImport("kernel32.dll")]
		public static extern IntPtr GetStdHandle(int nStdHandle);
		
		[DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();
		
		[DllImport("user32.dll")]
		public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
		
		[DllImport("kernel32.dll")]
		public static extern ulong GetTickCount64();
		
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
		
	}
}
"@
				
				# Add-Type with explicit error handling and assembly references
				# Note: We use our own POINT struct, so we don't need System.Drawing.dll
				$addTypeResult = $null
				$addTypeError = $null
				try {
					if ($DebugMode) {
						Write-Host "  [DEBUG] Attempting to add types..." -ForegroundColor $script:TextHighlight
					}
					$addTypeResult = Add-Type -TypeDefinition $typeDefinition -ReferencedAssemblies @("System.dll") -ErrorAction Stop
					if ($DebugMode) {
						Write-Host "  [OK] Add-Type completed successfully" -ForegroundColor $script:TextSuccess
					}
				} catch {
					$addTypeError = $_
					# If Add-Type fails, it might be because types already exist
					# Check if the error is about duplicate types
					if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*duplicate*" -or $_.Exception.Message -like "*Cannot add type*") {
						if ($DebugMode) {
							Write-Host "  [INFO] Types may already exist: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
						}
					} else {
						# Some other error occurred - log it
						if ($DebugMode) {
							Write-Host "  [WARN] Add-Type error: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
							if ($_.Exception.InnerException) {
								Write-Host "  [WARN] Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor $script:TextWarning
							}
						}
						# Don't throw yet - we'll check if types exist anyway
					}
				}
				
				# Always verify types were loaded, regardless of Add-Type result
				# Try both reflection and direct type access
				$loadedKeyboard = $null
				$loadedMouse = $null
				
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
				
				# If direct access failed, try reflection
				if ($null -eq $loadedKeyboard -or $null -eq $loadedMouse) {
					try {
						$allTypes = [System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.Namespace -eq 'mJiggAPI' }
						foreach ($type in $allTypes) {
							if ($type.Name -eq 'Keyboard' -and $null -eq $loadedKeyboard) { $loadedKeyboard = $type }
							if ($type.Name -eq 'Mouse' -and $null -eq $loadedMouse) { $loadedMouse = $type }
						}
					} catch {
						if ($DebugMode) {
							Write-Host "  [WARN] Error checking for loaded types: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
						}
					}
				}
				
				# Check if we have both types
				if ($null -ne $loadedKeyboard -and $null -ne $loadedMouse) {
					if ($DebugMode) {
						Write-Host "  [OK] All types verified: Keyboard, Mouse" -ForegroundColor $script:TextSuccess
					}
				} else {
					# Types weren't loaded - check if they already exist from previous check
					if ($null -ne $existingKeyboard -and $null -ne $existingMouse) {
						if ($DebugMode) {
							Write-Host "  [INFO] Types already exist from previous run" -ForegroundColor Gray
						}
					} else {
						# Types don't exist and failed to load - try to find them anywhere
						if ($DebugMode) {
							Write-Host "  [DEBUG] Searching all assemblies for mJiggAPI types..." -ForegroundColor $script:TextHighlight
							try {
								$allAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()
								foreach ($assembly in $allAssemblies) {
									try {
										$types = $assembly.GetTypes() | Where-Object { $_.Name -in @('Keyboard', 'Mouse') }
										if ($types) {
											Write-Host "    Found types in $($assembly.FullName): $($types | ForEach-Object { $_.FullName } | Join-String -Separator ', ')" -ForegroundColor Gray
										}
									} catch {
										# Some assemblies can't be inspected
									}
								}
							} catch {
								Write-Host "    Error searching assemblies: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
							}
						}
						
						# Types don't exist and failed to load
						$missingTypes = @()
						if ($null -eq $loadedKeyboard) { $missingTypes += "Keyboard" }
						if ($null -eq $loadedMouse) { $missingTypes += "Mouse" }
						$errorMsg = "Failed to load required mJiggAPI types: $($missingTypes -join ', ')"
						if ($addTypeError) {
							$errorMsg += "`nAdd-Type error: $($addTypeError.Exception.Message)"
						}
						if ($DebugMode) {
							Write-Host "  [FAIL] $errorMsg" -ForegroundColor $script:TextError
						}
						throw $errorMsg
					}
				}
			} catch {
				# Final fallback - check if types exist anyway
				$finalKeyboard = $null
				$finalMouse = $null
				try {
					$allTypes = [System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.Namespace -eq 'mJiggAPI' }
					foreach ($type in $allTypes) {
						if ($type.Name -eq 'Keyboard') { $finalKeyboard = $type }
						if ($type.Name -eq 'Mouse') { $finalMouse = $type }
					}
				} catch {
					# Ignore errors when checking for existing types
				}
				
				if ($null -ne $finalKeyboard -and $null -ne $finalMouse) {
					if ($DebugMode) {
						Write-Host "  [INFO] Types found after error recovery" -ForegroundColor Gray
					}
				} else {
					if ($DebugMode) {
						Write-Host "  [FAIL] Add-Type failed and types don't exist: $($_.Exception.Message)" -ForegroundColor $script:TextError
						Write-Host "  [INFO] This may require restarting PowerShell to reload types" -ForegroundColor $script:TextWarning
					}
					throw "Failed to load required mJiggAPI types: $($_.Exception.Message)"
				}
			}
		}
		
		# Verify types loaded correctly
		if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 2: Types loaded, verifying" | Out-File $script:StartupDiagFile -Append }
		try {
			$testKey = [mJiggAPI.Mouse]::GetAsyncKeyState(0x01)
			$testPoint = New-Object mJiggAPI.POINT
			$hasGetCursorPos = [mJiggAPI.Mouse].GetMethod("GetCursorPos") -ne $null
			if ($hasGetCursorPos) {
				$testMouse = [mJiggAPI.Mouse]::GetCursorPos([ref]$testPoint)
			}
			if ($DebugMode) {
				Write-Host "  [OK] Windows API types loaded successfully" -ForegroundColor $script:TextSuccess
			}
		} catch {
			if ($DebugMode) {
				Write-Host "  [FAIL] Could not verify keyboard/mouse API: $($_.Exception.Message)" -ForegroundColor $script:TextError
			}
			Write-Host "Warning: Could not verify keyboard/mouse API. Some features may be disabled." -ForegroundColor $script:TextWarning
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
					Write-Host "  [OK] Applied variance: $variance minutes, final end time: $endTimeStr" -ForegroundColor $script:TextSuccess
				}
			} catch {
				if ($DebugMode) {
					Write-Host "  [FAIL] Failed to apply variance: $($_.Exception.Message)" -ForegroundColor $script:TextError
				}
			}
		}
		
		# Calculate end date/time only if end time is set (not -1)
		if ($endTimeInt -ne -1) {
			try {
				$currentTime = Get-Date -Format "HHmm"
				if ($DebugMode) {
					Write-Host "  [OK] Current time: $currentTime" -ForegroundColor $script:TextSuccess
				}
			} catch {
				if ($DebugMode) {
					Write-Host "  [FAIL] Failed to get current time: $($_.Exception.Message)" -ForegroundColor $script:TextError
				}
				throw
			}
			try {
				if ($endTimeInt -le [int]$currentTime) {
					$tommorow = (Get-Date).AddDays(1)
					$endDate = Get-Date $tommorow -Format "MMdd"
					if ($DebugMode) {
						Write-Host "  [OK] End time is today, using tomorrow's date: $endDate" -ForegroundColor $script:TextSuccess
					}
				} else {
					$endDate = Get-Date -Format "MMdd"
					if ($DebugMode) {
						Write-Host "  [OK] End time is today, using today's date: $endDate" -ForegroundColor $script:TextSuccess
					}
				}
				$end = "$endDate$endTimeStr"
				$time = $false
				if ($DebugMode) {
					Write-Host "  [OK] Final end datetime: $end" -ForegroundColor $script:TextSuccess
				}
			} catch {
				if ($DebugMode) {
					Write-Host "  [FAIL] Failed to calculate end datetime: $($_.Exception.Message)" -ForegroundColor $script:TextError
				}
				throw
			}
		} else {
			# No end time - set end to empty and time to false
			$end = ""
			$time = $false
			if ($DebugMode) {
				Write-Host "  [OK] No end time - script will run indefinitely" -ForegroundColor $script:TextSuccess
			}
		}

		# Initialize lastPos for mouse detection
		if ($DebugMode) {
			Write-Host "[DEBUG] Initializing mouse position tracking..." -ForegroundColor $script:TextHighlight
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
					Write-Host "  [OK] Initial mouse position: $($LastPos.X), $($LastPos.Y)" -ForegroundColor $script:TextSuccess
				}
			} else {
				if ($DebugMode) {
					Write-Host "  [OK] Mouse position already set: $($LastPos.X), $($LastPos.Y)" -ForegroundColor $script:TextSuccess
				}
			}
		} catch {
			if ($DebugMode) {
				Write-Host "  [FAIL] Failed to get mouse position: $($_.Exception.Message)" -ForegroundColor $script:TextError
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
			
			# Define emoji arrows using ConvertFromUtf32 for cross-version compatibility
			$arrowRight = [char]::ConvertFromUtf32(0x27A1)  # ➡
			$arrowLeft = [char]::ConvertFromUtf32(0x2B05)   # ⬅
			$arrowDown = [char]::ConvertFromUtf32(0x2B07)   # ⬇
			$arrowUp = [char]::ConvertFromUtf32(0x2B06)     # ⬆
			$arrowSE = [char]::ConvertFromUtf32(0x2198)     # ↘
			$arrowNE = [char]::ConvertFromUtf32(0x2197)     # ↗
			$arrowSW = [char]::ConvertFromUtf32(0x2199)     # ↙
			$arrowNW = [char]::ConvertFromUtf32(0x2196)     # ↖
			
			# Simple arrows (BMP characters, work with [char])
			$simpleRight = [char]0x2192  # →
			$simpleLeft = [char]0x2190   # ←
			$simpleDown = [char]0x2193   # ↓
			$simpleUp = [char]0x2191     # ↑
			$simpleSE = [char]0x2198     # ↘
			$simpleNE = [char]0x2197     # ↗
			$simpleSW = [char]0x2199     # ↙
			$simpleNW = [char]0x2196     # ↖
			
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
					if ($deltaX -gt 0) { return $arrowRight } else { return $arrowLeft }
				} else {
					# simple style
					if ($deltaX -gt 0) { return $simpleRight } else { return $simpleLeft }
				}
			} elseif ($absY -gt $absX * 2) {
				# Primarily vertical
				if ($style -eq "text") {
					if ($deltaY -gt 0) { return "S" } else { return "N" }
				} elseif ($style -eq "arrows") {
					if ($deltaY -gt 0) { return $arrowDown } else { return $arrowUp }
				} else {
					# simple style
					if ($deltaY -gt 0) { return $simpleDown } else { return $simpleUp }
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
						return $arrowSE
					} elseif ($deltaX -gt 0 -and $deltaY -lt 0) {
						return $arrowNE
					} elseif ($deltaX -lt 0 -and $deltaY -gt 0) {
						return $arrowSW
					} else {
						return $arrowNW
					}
				} else {
					# simple style
					if ($deltaX -gt 0 -and $deltaY -gt 0) {
						return $simpleSE
					} elseif ($deltaX -gt 0 -and $deltaY -lt 0) {
						return $simpleNE
					} elseif ($deltaX -lt 0 -and $deltaY -gt 0) {
						return $simpleSW
					} else {
						return $simpleNW
					}
				}
			}
		}

		if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 3: About to define helper functions" | Out-File $script:StartupDiagFile -Append }

		# ============================================
		# Buffered Rendering Functions
		# ============================================

		$script:ESC = [char]27
		$script:CursorVisible = $false
		$script:AnsiFG = @{
			[ConsoleColor]::Black = 30; [ConsoleColor]::DarkBlue = 34; [ConsoleColor]::DarkGreen = 32; [ConsoleColor]::DarkCyan = 36
			[ConsoleColor]::DarkRed = 31; [ConsoleColor]::DarkMagenta = 35; [ConsoleColor]::DarkYellow = 33; [ConsoleColor]::Gray = 37
			[ConsoleColor]::DarkGray = 90; [ConsoleColor]::Blue = 94; [ConsoleColor]::Green = 92; [ConsoleColor]::Cyan = 96
			[ConsoleColor]::Red = 91; [ConsoleColor]::Magenta = 95; [ConsoleColor]::Yellow = 93; [ConsoleColor]::White = 97
		}
		$script:AnsiBG = @{
			[ConsoleColor]::Black = 40; [ConsoleColor]::DarkBlue = 44; [ConsoleColor]::DarkGreen = 42; [ConsoleColor]::DarkCyan = 46
			[ConsoleColor]::DarkRed = 41; [ConsoleColor]::DarkMagenta = 45; [ConsoleColor]::DarkYellow = 43; [ConsoleColor]::Gray = 47
			[ConsoleColor]::DarkGray = 100; [ConsoleColor]::Blue = 104; [ConsoleColor]::Green = 102; [ConsoleColor]::Cyan = 106
			[ConsoleColor]::Red = 101; [ConsoleColor]::Magenta = 105; [ConsoleColor]::Yellow = 103; [ConsoleColor]::White = 107
		}

		function Write-Buffer {
			param(
				[int]$X = -1,
				[int]$Y = -1,
				[string]$Text,
				[object]$FG = $null,
				[object]$BG = $null,
				[switch]$Wide
			)
			if ($Wide -and $null -ne $BG) { $Text = $Text + " " }
			$script:RenderQueue.Add(@{ X = $X; Y = $Y; Text = $Text; FG = $FG; BG = $BG })
		}

		function Flush-Buffer {
			param([switch]$ClearFirst)
			if ($script:RenderQueue.Count -eq 0) { return }
			$csi = "$($script:ESC)["
			$sb = New-Object System.Text.StringBuilder (8192)
			[void]$sb.Append("${csi}?25l")
			if ($ClearFirst) { [void]$sb.Append("${csi}2J") }
			$lastFGCode = -1
			$lastBGCode = -1
			foreach ($seg in $script:RenderQueue) {
				$fgCode = if ($null -ne $seg.FG) { $script:AnsiFG[[ConsoleColor]$seg.FG] } else { 39 }
				$bgCode = if ($null -ne $seg.BG) { $script:AnsiBG[[ConsoleColor]$seg.BG] } else { 49 }
				if ($seg.X -ge 0 -and $seg.Y -ge 0) {
					[void]$sb.Append("${csi}$($seg.Y + 1);$($seg.X + 1)H")
				}
				if ($fgCode -ne $lastFGCode -or $bgCode -ne $lastBGCode) {
					[void]$sb.Append("${csi}${fgCode};${bgCode}m")
					$lastFGCode = $fgCode
					$lastBGCode = $bgCode
				}
				[void]$sb.Append($seg.Text)
			}
			[void]$sb.Append("${csi}0m")
			if ($script:CursorVisible) { [void]$sb.Append("${csi}?25h") }
			[Console]::Write($sb.ToString())
			$script:RenderQueue.Clear()
		}

		function Clear-Buffer {
			$script:RenderQueue.Clear()
		}

		# Function to draw drop shadow for dialog boxes
		function Draw-DialogShadow {
			param(
				[int]$dialogX,
				[int]$dialogY,
				[int]$dialogWidth,
				[int]$dialogHeight,
				[string]$shadowColor = "DarkGray"
			)
			
			$shadowChar = [char]0x2591  # ░ light shade character
			
			for ($i = 1; $i -le $dialogHeight; $i++) {
				Write-Buffer -X ($dialogX + $dialogWidth) -Y ($dialogY + $i) -Text "$shadowChar" -FG $shadowColor
			}
			for ($i = 1; $i -le $dialogWidth; $i++) {
				Write-Buffer -X ($dialogX + $i) -Y ($dialogY + $dialogHeight + 1) -Text "$shadowChar" -FG $shadowColor
			}
			Write-Buffer -X ($dialogX + $dialogWidth) -Y ($dialogY + $dialogHeight + 1) -Text "$shadowChar" -FG $shadowColor
		}
		
		# Function to clear drop shadow for dialog boxes
		function Clear-DialogShadow {
			param(
				[int]$dialogX,
				[int]$dialogY,
				[int]$dialogWidth,
				[int]$dialogHeight
			)
			
			for ($i = 1; $i -le $dialogHeight; $i++) {
				Write-Buffer -X ($dialogX + $dialogWidth) -Y ($dialogY + $i) -Text " "
			}
			for ($i = 1; $i -le $dialogWidth; $i++) {
				Write-Buffer -X ($dialogX + $i) -Y ($dialogY + $dialogHeight + 1) -Text " "
			}
			Write-Buffer -X ($dialogX + $dialogWidth) -Y ($dialogY + $dialogHeight + 1) -Text " "
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
			
			$savedCursorVisible = $script:CursorVisible
			$script:CursorVisible = $true
			[Console]::Write("$($script:ESC)[?25h")
			
			$checkmark = [char]::ConvertFromUtf32(0x2705)  # ✅ green checkmark
			$redX = [char]::ConvertFromUtf32(0x274C)  # ❌ red X
			# Button line display width calculation:
			# "$($script:BoxVertical) " (2) + checkmark (2) + "|" (1) + "(u)pdate" (8) + "  " (2) + redX (2) + "|" (1) + "(c)ancel" (8) = 26
			# So we need: 35 - 26 - 1 = 8 spaces before closing $($script:BoxVertical)
			$bottomLinePadding = 8
			
			# Build all lines to be exactly 35 characters using Get-Padding helper
			$line0 = "$($script:BoxTopLeft)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxTopRight)"  # 35 chars
			$line1Text = "$($script:BoxVertical)  Change End Time"
			$line1Padding = Get-Padding -usedWidth ($line1Text.Length + 1) -totalWidth $dialogWidth
			$line1 = $line1Text + (" " * $line1Padding) + "$($script:BoxVertical)"
			
			$line2 = "$($script:BoxVertical)" + (" " * 33) + "$($script:BoxVertical)"  # 35 chars
			
			$line3Text = "$($script:BoxVertical)  Enter new time (HHmm format):"
			$line3Padding = Get-Padding -usedWidth ($line3Text.Length + 1) -totalWidth $dialogWidth
			$line3 = $line3Text + (" " * $line3Padding) + "$($script:BoxVertical)"
			
			# Line 4 will be drawn separately with highlighted field
			$line4Text = "$($script:BoxVertical)  "
			$line4Padding = Get-Padding -usedWidth ($line4Text.Length + 1 + 6) -totalWidth $dialogWidth  # +6 for "[    ]"
			$line4 = $line4Text + (" " * $line4Padding) + "$($script:BoxVertical)"
			
			$line5 = "$($script:BoxVertical)" + (" " * 33) + "$($script:BoxVertical)"  # 35 chars
			
			$line7 = "$($script:BoxBottomLeft)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxBottomRight)"  # 35 chars
			
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
			
			# Draw dialog background (clear area) with themed background
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth) -BG $script:TimeDialogBg
			}
			
			# Draw dialog box with themed background
			for ($i = 0; $i -lt $dialogLines.Count; $i++) {
				if ($i -eq 1) {
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -Text "Change End Time" -FG $script:TimeDialogTitle -BG $script:TimeDialogBg
					$titleUsedWidth = 3 + "Change End Time".Length
					$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
					Write-Buffer -Text (" " * $titlePadding) -BG $script:TimeDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
				} elseif ($i -eq 4) {
					$initialTimeDisplay = if ($currentEndTime -ne -1 -and $currentEndTime -ne 0) { 
						$currentEndTime.ToString().PadLeft(4, '0') 
					} else { 
						"" 
					}
					$fieldDisplay = $initialTimeDisplay.PadRight(4)
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -Text "[" -FG $script:TimeDialogText -BG $script:TimeDialogBg
					Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
					Write-Buffer -Text "]" -FG $script:TimeDialogText -BG $script:TimeDialogBg
					$fieldUsedWidth = 3 + 6
					$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
					Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
				} elseif ($i -eq 6) {
					$checkmarkX = $dialogX + 2
					$redXX = $dialogX + 15
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical) " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -X $checkmarkX -Y ($dialogY + $i) -Text "$checkmark" -FG $script:TextSuccess -BG $script:TimeDialogButtonBg -Wide
				Write-Buffer -X ($checkmarkX + 2) -Y ($dialogY + $i) -Text "|(u)pdate" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
				Write-Buffer -X ($checkmarkX + 4) -Y ($dialogY + $i) -Text "u" -FG $script:TimeDialogButtonHotkey -BG $script:TimeDialogButtonBg
				Write-Buffer -X ($checkmarkX + 5) -Y ($dialogY + $i) -Text ")pdate" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
				Write-Buffer -Text "  " -BG $script:TimeDialogBg
				Write-Buffer -X $redXX -Y ($dialogY + $i) -Text "$redX" -FG $script:TextError -BG $script:TimeDialogButtonBg -Wide
				Write-Buffer -X ($redXX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
				Write-Buffer -Text "(" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
				Write-Buffer -Text "c" -FG $script:TimeDialogButtonHotkey -BG $script:TimeDialogButtonBg
				Write-Buffer -Text ")ancel" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
					Write-Buffer -Text (" " * $bottomLinePadding) -BG $script:TimeDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
				} else {
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text $dialogLines[$i] -FG $script:TimeDialogText -BG $script:TimeDialogBg
				}
			}
			
			# Draw drop shadow
			Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:TimeDialogShadow
			Flush-Buffer
			
			# Calculate button bounds for click detection (visible characters only)
			# Button row is at dialogY + 6 (line 6)
			$buttonRowY = $dialogY + 6
			# Rendered: +0:border +1:space +2-3:✅(2cells) +4:| +5-12:(u)pdate +13-14:spaces +15-16:❌(2cells) +17:| +18-25:(c)ancel +26-33:padding +34:border
			$updateButtonStartX = $dialogX + 2
			$updateButtonEndX = $dialogX + 12
			$cancelButtonStartX = $dialogX + 15
			$cancelButtonEndX = $dialogX + 25
			
			# Store button bounds in script scope for main loop click detection
			$script:DialogButtonBounds = @{
				buttonRowY = $buttonRowY
				updateStartX = $updateButtonStartX
				updateEndX = $updateButtonEndX
				cancelStartX = $cancelButtonStartX
				cancelEndX = $cancelButtonEndX
			}
			$script:DialogButtonClick = $null  # Clear any previous click  # 20 + 2 (emoji) + 1 (pipe) + 8 (text) - 1 (inclusive)
			
			# Position cursor in input field (inside the brackets, after "$($script:BoxVertical)  [")
			# Line 4 is "$($script:BoxVertical)  [" + 4 spaces + "]", so input starts at position 4
			$inputX = $dialogX + 4
			$inputY = $dialogY + 4
			$script:CursorVisible = $false
			[Console]::Write("$($script:ESC)[?25l")
			
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
					$updateButtonStartX = $dialogX + 2
					$updateButtonEndX = $dialogX + 12
					$cancelButtonStartX = $dialogX + 15
					$cancelButtonEndX = $dialogX + 25
					
					# Update button bounds in script scope
					$script:DialogButtonBounds = @{
						buttonRowY = $buttonRowY
						updateStartX = $updateButtonStartX
						updateEndX = $updateButtonEndX
						cancelStartX = $cancelButtonStartX
						cancelEndX = $cancelButtonEndX
					}
					
					for ($i = 0; $i -lt $dialogLines.Count; $i++) {
						if ($i -eq 1) {
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text "Change End Time" -FG $script:TimeDialogTitle -BG $script:TimeDialogBg
							$titleUsedWidth = 3 + "Change End Time".Length
							$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $titlePadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						} elseif ($i -eq 4) {
							$fieldDisplay = $timeInput.PadRight(4)
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text "[" -FG $script:TimeDialogText -BG $script:TimeDialogBg
							Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
							Write-Buffer -Text "]" -FG $script:TimeDialogText -BG $script:TimeDialogBg
							$fieldUsedWidth = 3 + 6
							$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						} elseif ($i -eq 6) {
							$checkmarkX = $dialogX + 2
							$redXX = $dialogX + 15
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical) " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -X $checkmarkX -Y ($dialogY + $i) -Text "$checkmark" -FG $script:TextSuccess -BG $script:TimeDialogButtonBg -Wide
						Write-Buffer -X ($checkmarkX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
						Write-Buffer -Text "(" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
						Write-Buffer -Text "u" -FG $script:TimeDialogButtonHotkey -BG $script:TimeDialogButtonBg
						Write-Buffer -Text ")pdate" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
						Write-Buffer -Text "  " -BG $script:TimeDialogBg
						Write-Buffer -X $redXX -Y ($dialogY + $i) -Text "$redX" -FG $script:TextError -BG $script:TimeDialogButtonBg -Wide
						Write-Buffer -X ($redXX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
						Write-Buffer -Text "(" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
						Write-Buffer -Text "c" -FG $script:TimeDialogButtonHotkey -BG $script:TimeDialogButtonBg
						Write-Buffer -Text ")ancel" -FG $script:TimeDialogButtonText -BG $script:TimeDialogButtonBg
							Write-Buffer -Text (" " * $bottomLinePadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						} else {
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text $dialogLines[$i] -FG $script:TimeDialogText -BG $script:TimeDialogBg
						}
					}
					
					Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:TimeDialogShadow
					
					$fieldDisplay = $timeInput.PadRight(4)
					Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -Text "[" -FG $script:TimeDialogText -BG $script:TimeDialogBg
					Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
					Write-Buffer -Text "]" -FG $script:TimeDialogText -BG $script:TimeDialogBg
					$fieldUsedWidth = 3 + 6
					$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
					Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					if ($errorMessage -ne "") {
						Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text $errorMessage -FG $script:TextError -BG $script:TimeDialogBg
						$errorLineUsedWidth = 3 + $errorMessage.Length
						$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
						Write-Buffer -Text (" " * $errorLinePadding) -BG $script:TimeDialogBg
						Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					} else {
						Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text (" " * ($dialogWidth - 2)) -BG $script:TimeDialogBg
						Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					}
					Flush-Buffer -ClearFirst
					[Console]::SetCursorPosition($inputX + $timeInput.Length, $inputY)
				}
				
				# Check for mouse button clicks on dialog buttons via console input buffer
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				
				try {
					$peekBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' 16
					$peekEvts = [uint32]0
					$hIn = [mJiggAPI.Mouse]::GetStdHandle(-10)
					if ([mJiggAPI.Mouse]::PeekConsoleInput($hIn, $peekBuf, 16, [ref]$peekEvts) -and $peekEvts -gt 0) {
						$lastClickIdx = -1
						$clickX = -1; $clickY = -1
						for ($e = 0; $e -lt $peekEvts; $e++) {
							if ($peekBuf[$e].EventType -eq 0x0002 -and $peekBuf[$e].MouseEvent.dwEventFlags -eq 0 -and ($peekBuf[$e].MouseEvent.dwButtonState -band 0x0001) -ne 0) {
								$clickX = $peekBuf[$e].MouseEvent.dwMousePosition.X
								$clickY = $peekBuf[$e].MouseEvent.dwMousePosition.Y
								$lastClickIdx = $e
							}
						}
						if ($lastClickIdx -ge 0) {
							$consumeCount = [uint32]($lastClickIdx + 1)
							$flushBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' $consumeCount
							$flushed = [uint32]0
							[mJiggAPI.Mouse]::ReadConsoleInput($hIn, $flushBuf, $consumeCount, [ref]$flushed) | Out-Null
							
							if ($clickY -eq $buttonRowY -and $clickX -ge $updateButtonStartX -and $clickX -le $updateButtonEndX) {
								$char = "u"; $keyProcessed = $true
							} elseif ($clickY -eq $buttonRowY -and $clickX -ge $cancelButtonStartX -and $clickX -le $cancelButtonEndX) {
								$char = "c"; $keyProcessed = $true
							}
							if ($DebugMode) {
								$clickTarget = if ($keyProcessed) { "button:$char" } else { "none" }
								if ($null -eq $LogArray -or -not ($LogArray -is [Array])) { $LogArray = @() }
								$LogArray += [PSCustomObject]@{
									logRow = $true
									components = @(
										@{ priority = 1; text = (Get-Date).ToString("HH:mm:ss"); shortText = (Get-Date).ToString("HH:mm:ss") },
										@{ priority = 2; text = " - [DEBUG] Time dialog click at ($clickX,$clickY), target: $clickTarget"; shortText = " - [DEBUG] Click ($clickX,$clickY) -> $clickTarget" }
									)
								}
							}
						}
					}
				} catch { }
				
				# Check for dialog button clicks detected by main loop
				if (-not $keyProcessed -and $null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null
					if ($buttonClick -eq "Update") { $char = "u"; $keyProcessed = $true }
					elseif ($buttonClick -eq "Cancel") { $char = "c"; $keyProcessed = $true }
				}
				
				# Wait for key input (non-blocking check)
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
								$fieldDisplay = $timeInput.PadRight(4)
								Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
								Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								$fieldUsedWidth = 3 + 6
								$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
								Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
								Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -Text $errorMessage -FG $script:TextError -BG $script:TimeDialogBg
								$errorLineUsedWidth = 3 + $errorMessage.Length
								$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
								Write-Buffer -Text (" " * $errorLinePadding) -BG $script:TimeDialogBg
								Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Flush-Buffer
								[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
							}
						} catch {
							# Invalid input - show error
							$errorMessage = "Invalid hours"
							# Redraw input field with highlight - redraw entire line 4
							$fieldDisplay = $timeInput.PadRight(4)
							Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
							Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							$fieldUsedWidth = 3 + 6
							$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text $errorMessage -FG $script:TextError -BG $script:TimeDialogBg
							$errorLineUsedWidth = 3 + $errorMessage.Length
							$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $errorLinePadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Flush-Buffer
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
								$fieldDisplay = $timeInput.PadRight(4)
								Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
								Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								$fieldUsedWidth = 3 + 6
								$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
								Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
								Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Write-Buffer -Text $errorMessage -FG $script:TextError -BG $script:TimeDialogBg
								$errorLineUsedWidth = 3 + $errorMessage.Length
								$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
								Write-Buffer -Text (" " * $errorLinePadding) -BG $script:TimeDialogBg
								Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
								Flush-Buffer
								[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
							}
						} catch {
							# Invalid input - show error (shouldn't normally happen with numeric-only input)
							$errorMessage = "Number out of range"
							# Redraw input field with highlight - redraw entire line 4
							$fieldDisplay = $timeInput.PadRight(4)
							Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
							Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							$fieldUsedWidth = 3 + 6
							$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Write-Buffer -Text $errorMessage -FG $script:TextError -BG $script:TimeDialogBg
							$errorLineUsedWidth = 3 + $errorMessage.Length
							$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $errorLinePadding) -BG $script:TimeDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
							Flush-Buffer
							[Console]::SetCursorPosition($inputX + 1 + $timeInput.Length, $inputY)
						}
					} else {
						# Not 4 digits yet - show error
						$errorMessage = "Enter 4 digits (HHmm format)"
						# Redraw input field with highlight - redraw entire line 4
						$fieldDisplay = $timeInput.PadRight(4)
						Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
						Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						$fieldUsedWidth = 3 + 6
						$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
						Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
						Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text $errorMessage -FG $script:TextError -BG $script:TimeDialogBg
						$errorLineUsedWidth = 3 + $errorMessage.Length
						$errorLinePadding = Get-Padding -usedWidth ($errorLineUsedWidth + 1) -totalWidth $dialogWidth
						Write-Buffer -Text (" " * $errorLinePadding) -BG $script:TimeDialogBg
						Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Flush-Buffer
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
							$script:CursorVisible = $false; [Console]::Write("$($script:ESC)[?25l")
						}
						$errorMessage = ""  # Clear error when editing
						# Redraw input with highlight - redraw entire line 4 to ensure clean overwrite
						$fieldDisplay = $timeInput.PadRight(4)
						Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
						Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						$fieldUsedWidth = 3 + 6
						$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
						Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
						Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Write-Buffer -Text (" " * 33) -BG $script:TimeDialogBg
						Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
						Flush-Buffer
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
						$script:CursorVisible = $true; [Console]::Write("$($script:ESC)[?25h")
					} elseif ($timeInput.Length -lt 4) {
						$timeInput += $char.ToString()  # Convert char to string
					}
					$errorMessage = ""  # Clear error when typing
					# Redraw input field with highlight - redraw entire line 4 to ensure clean overwrite
					$fieldDisplay = $timeInput.PadRight(4)
					Write-Buffer -X $dialogX -Y $inputY -Text "$($script:BoxVertical)  " -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -Text "[" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -Text $fieldDisplay -FG $script:TimeDialogFieldText -BG $script:TimeDialogFieldBg
					Write-Buffer -Text "]" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					$fieldUsedWidth = 3 + 6
					$fieldPadding = Get-Padding -usedWidth ($fieldUsedWidth + 1) -totalWidth $dialogWidth
					Write-Buffer -Text (" " * $fieldPadding) -BG $script:TimeDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Write-Buffer -Text (" " * 33) -BG $script:TimeDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:TimeDialogBorder -BG $script:TimeDialogBg
					Flush-Buffer
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
			
			Clear-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth)
			}
			Flush-Buffer
			
			$script:CursorVisible = $savedCursorVisible
			if ($script:CursorVisible) { [Console]::Write("$($script:ESC)[?25h") } else { [Console]::Write("$($script:ESC)[?25l") }
			
			$script:DialogButtonBounds = $null
			$script:DialogButtonClick = $null
			
			return @{
				Result = $result
				NeedsRedraw = $needsRedraw
			}
		}

		# ============================================
		# Performance Helper Functions
		# ============================================
		
		# Playful quotes for resize screen
		$script:ResizeQuotes = @(
			"Jiggling since the dawn of idle timeouts..."
			"A mouse in motion stays employed"
			"Wiggle wiggle wiggle"
			"Like jello, but for your cursor"
			"Making mice dance since 2024"
			"The early mouse gets the jiggle"
			"Shake it like a Polaroid picture"
			"Keep calm and jiggle on"
			"This mouse has moves"
			"Cursor cardio in progress"
			"Staying active so you don't have to"
			"Mice just wanna have fun"
			"Jiggle physics: enabled"
			"Not all who wander are lost, some are jiggling"
			"Professional mouse motivator"
			"Your mouse's personal trainer"
			"Wiggling through the workday"
		)
		$script:CurrentResizeQuote = $null
		
		# Helper function: Draw centered logo during window resize using buffered output
		function Draw-ResizeLogo {
			param([switch]$ClearFirst)
			try {
				$rawUI = $Host.UI.RawUI
				$winSize = $rawUI.WindowSize
				$winWidth = $winSize.Width
				$winHeight = $winSize.Height
				
				# Only draw if window is large enough
				if ($winWidth -lt 16 -or $winHeight -lt 14) {
					return
				}
				
				# Select a random quote if we don't have one yet
				if ($null -eq $script:CurrentResizeQuote) {
					$script:CurrentResizeQuote = $script:ResizeQuotes | Get-Random
				}
				
				# Box-drawing characters
				$boxTL = [char]0x250C      # ┌
				$boxTR = [char]0x2510      # ┐
				$boxBL = [char]0x2514      # └
				$boxBR = [char]0x2518      # ┘
				$boxH = [char]0x2500       # ─
				$boxV = [char]0x2502       # │
				
				# Logo display width: mJig( (5) + emoji (2) + ) (1) = 8
				$logoDisplayWidth = 8
				
				# Calculate center position for logo
				$centerX = [math]::Floor(($winWidth - $logoDisplayWidth) / 2)
				$centerY = [math]::Floor($winHeight / 2)
				
				# Box dimensions: scale with screen size while maintaining minimum padding
				$minPadding = 3
				# Calculate available space around logo
				$availableH = [math]::Min($centerX - 1, $winWidth - $centerX - $logoDisplayWidth - 1)
				$availableV = [math]::Min($centerY - 1, $winHeight - $centerY - 2)
				# Use 42% of available space as padding, with minimum
				$boxPaddingH = [math]::Max($minPadding * 2, [math]::Floor($availableH * 0.42))
				$boxPaddingV = [math]::Max($minPadding, [math]::Floor($availableV * 0.42))
				$boxLeft = $centerX - $boxPaddingH - 1
				$boxRight = $centerX + $logoDisplayWidth + $boxPaddingH
				$boxTop = $centerY - $boxPaddingV - 1
				$boxBottom = $centerY + $boxPaddingV + 1
				$boxInnerWidth = $boxRight - $boxLeft - 1
				
				# Build horizontal line string once
				$hLine = [string]::new($boxH, $boxInnerWidth)
				
				Write-Buffer -X $boxLeft -Y $boxTop -Text "$boxTL$hLine$boxTR"
				for ($y = $boxTop + 1; $y -lt $boxBottom; $y++) {
					Write-Buffer -X $boxLeft -Y $y -Text "$boxV"
					Write-Buffer -X $boxRight -Y $y -Text "$boxV"
				}
				Write-Buffer -X $boxLeft -Y $boxBottom -Text "$boxBL$hLine$boxBR"
				
				$emojiX = $centerX + 5
				Write-Buffer -X $centerX -Y $centerY -Text "mJig(" -FG $script:ResizeLogoName
				Write-Buffer -X $emojiX -Y $centerY -Text ([char]::ConvertFromUtf32(0x1F400)) -FG $script:ResizeLogoIcon
				Write-Buffer -X ($emojiX + 2) -Y $centerY -Text ")" -FG $script:ResizeLogoName
				
				$quoteY = $centerY + 2
				if ($quoteY -lt $boxBottom -and $null -ne $script:CurrentResizeQuote) {
					$quote = $script:CurrentResizeQuote
					$maxQuoteWidth = $boxInnerWidth - 2
					if ($quote.Length -gt $maxQuoteWidth) {
						$quote = $quote.Substring(0, $maxQuoteWidth - 3) + "..."
					}
					$quoteX = [math]::Floor(($winWidth - $quote.Length) / 2)
					Write-Buffer -X $quoteX -Y $quoteY -Text $quote -FG $script:ResizeQuoteText
				}
				if ($ClearFirst) { Flush-Buffer -ClearFirst } else { Flush-Buffer }
				
			} catch {
				try {
					$winSize = $Host.UI.RawUI.WindowSize
					$centerX = [math]::Max(0, [math]::Floor(($winSize.Width - 8) / 2))
					$centerY = [math]::Max(0, [math]::Floor($winSize.Height / 2))
					Write-Buffer -X $centerX -Y $centerY -Text "mJig(" -FG $script:ResizeLogoName
					Write-Buffer -Text ([char]::ConvertFromUtf32(0x1F400)) -FG $script:ResizeLogoIcon
					Write-Buffer -Text ")" -FG $script:ResizeLogoName
					if ($ClearFirst) { Flush-Buffer -ClearFirst } else { Flush-Buffer }
				} catch { }
			}
		}
		
		# Helper function: Get method safely (cached for performance)
		function Get-CachedMethod {
			param(
				$type,
				[string]$methodName
			)
			$cacheKey = "$($type.FullName).$methodName"
			if (-not $script:MethodCache.ContainsKey($cacheKey)) {
				$script:MethodCache[$cacheKey] = $type.GetMethod($methodName)
			}
			return $script:MethodCache[$cacheKey]
		}
		
		# Helper function: Get mouse position (uses cached method)
		function Get-MousePosition {
			$point = New-Object mJiggAPI.POINT
			$mouseType = [mJiggAPI.Mouse]
			$getCursorPosMethod = Get-CachedMethod -type $mouseType -methodName "GetCursorPos"
			if ($null -ne $getCursorPosMethod -and [mJiggAPI.Mouse]::GetCursorPos([ref]$point)) {
				return New-Object System.Drawing.Point($point.X, $point.Y)
			}
			return $null
		}
		
		# Helper function: Check mouse movement threshold
		function Test-MouseMoved {
			param(
				[System.Drawing.Point]$currentPos,
				[System.Drawing.Point]$lastPos,
				[int]$threshold = 2
			)
			if ($null -eq $lastPos) { return $false }
			$deltaX = [Math]::Abs($currentPos.X - $lastPos.X)
			$deltaY = [Math]::Abs($currentPos.Y - $lastPos.Y)
			return ($deltaX -gt $threshold -or $deltaY -gt $threshold)
		}
		
		# Helper function: Calculate time since (in milliseconds)
		# Returns MaxValue if startTime is null (allows safe comparison without null checks)
		function Get-TimeSinceMs {
			param($startTime)
			if ($null -eq $startTime) { return [double]::MaxValue }
			return ((Get-Date) - [DateTime]$startTime).TotalMilliseconds
		}
		
		# Helper function: Calculate value with random variance
		function Get-ValueWithVariance {
			param([double]$baseValue, [double]$variance)
			$varianceAmount = Get-Random -Minimum 0.0 -Maximum ($variance + 0.0001)
			if ((Get-Random -Maximum 2) -eq 0) {
				return $baseValue - $varianceAmount
			} else {
				return $baseValue + $varianceAmount
			}
		}
		
		# Helper function: Clamp coordinates to screen bounds
		function Set-CoordinateBounds {
			param([ref]$x, [ref]$y)
			$x.Value = [Math]::Max(0, [Math]::Min($x.Value, $script:ScreenWidth - 1))
			$y.Value = [Math]::Max(0, [Math]::Min($y.Value, $script:ScreenHeight - 1))
		}
		
		
		# ============================================
		# UI Helper Functions
		# ============================================
		
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
				[string]$leftChar = "$($script:BoxVertical)",
				[string]$rightChar = "$($script:BoxVertical)",
				[string]$fillChar = " ",
				[System.ConsoleColor]$borderColor = [System.ConsoleColor]::White,
				[System.ConsoleColor]$fillColor = [System.ConsoleColor]::White
			)
			
			$fillWidth = $width - 2
			Write-Buffer -X $x -Y $y -Text $leftChar -FG $borderColor
			Write-Buffer -Text ($fillChar * $fillWidth) -FG $fillColor
			Write-Buffer -Text $rightChar -FG $borderColor
		}
		
		# Helper function: Draw a simple dialog row (no description box)
		function Write-SimpleDialogRow {
			param(
				[int]$x,
				[int]$y,
				[int]$width,
				[string]$content = "",
				[System.ConsoleColor]$contentColor = [System.ConsoleColor]::White,
				[System.ConsoleColor]$backgroundColor = $null
			)
			
			$borderFG = if ($null -ne $backgroundColor) { $script:MoveDialogBorder } else { $null }
			Write-Buffer -X $x -Y $y -Text "$($script:BoxVertical)" -FG $borderFG -BG $backgroundColor
			if ($content.Length -gt 0) {
				Write-Buffer -Text " " -BG $backgroundColor
				Write-Buffer -Text $content -FG $contentColor -BG $backgroundColor
				$usedWidth = 1 + 1 + $content.Length
				$padding = Get-Padding -usedWidth ($usedWidth + 1) -totalWidth $width
				Write-Buffer -Text (" " * $padding) -BG $backgroundColor
			} else {
				Write-Buffer -Text (" " * ($width - 2)) -BG $backgroundColor
			}
			Write-Buffer -Text "$($script:BoxVertical)" -FG $borderFG -BG $backgroundColor
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
				[int]$currentFieldIndex,
				[System.ConsoleColor]$backgroundColor = $null
			)
			
			$labelPadding = [Math]::Max(0, $longestLabel - $label.Length)
			$labelText = "$($script:BoxVertical)  " + $label + (" " * $labelPadding)
			
			$fieldDisplay = if ([string]::IsNullOrEmpty($fieldValue)) { "" } else { $fieldValue }
			$fieldDisplay = $fieldDisplay.PadRight($fieldWidth)
			$fieldContent = "[" + $fieldDisplay + "]"
			
			$labelFG = if ($null -ne $backgroundColor) { $script:MoveDialogText } else { $null }
			$borderFG = if ($null -ne $backgroundColor) { $script:MoveDialogBorder } else { $null }
			$fieldFG = if ($fieldIndex -eq $currentFieldIndex) {
				if ($null -ne $backgroundColor) { $script:MoveDialogFieldText } else { $script:TimeDialogFieldText }
			} else {
				$script:TextHighlight
			}
			$fieldBG = if ($fieldIndex -eq $currentFieldIndex) {
				if ($null -ne $backgroundColor) { $script:MoveDialogFieldBg } else { $script:TimeDialogFieldBg }
			} else {
				$backgroundColor
			}
			
			Write-Buffer -X $x -Y $y -Text $labelText -FG $labelFG -BG $backgroundColor
			Write-Buffer -Text "[" -FG $labelFG -BG $backgroundColor
			Write-Buffer -Text $fieldDisplay -FG $fieldFG -BG $fieldBG
			Write-Buffer -Text "]" -FG $labelFG -BG $backgroundColor
			$usedWidth = $labelText.Length + $fieldContent.Length
			$remainingPadding = Get-Padding -usedWidth ($usedWidth + 1) -totalWidth $width
			Write-Buffer -Text (" " * $remainingPadding) -BG $backgroundColor
			Write-Buffer -Text "$($script:BoxVertical)" -FG $borderFG -BG $backgroundColor
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
			
			$savedCursorVisible = $script:CursorVisible
			$script:CursorVisible = $false
			[Console]::Write("$($script:ESC)[?25l")
			
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
			$inputBoxStartX = 3 + $longestLabel  # "$($script:BoxVertical)  " + longest label = X position where all input boxes start
			
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
				
			# Clear dialog area with themed background
			for ($i = 0; $i -lt $height; $i++) {
				Write-Buffer -X $x -Y ($y + $i) -Text (" " * $width) -BG $script:MoveDialogBg
			}
				
			# Top border (spans full width)
			Write-Buffer -X $x -Y $y -Text "$($script:BoxTopLeft)" -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
			Write-Buffer -Text ("$($script:BoxHorizontal)" * ($width - 2)) -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
			Write-Buffer -Text "$($script:BoxTopRight)" -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
				
				# Title line
				Write-SimpleDialogRow -x $x -y ($y + 1) -width $width -content "Modify Movement Settings" -contentColor $script:MoveDialogTitle -backgroundColor $script:MoveDialogBg
				
				# Empty line (row 2)
				Write-SimpleDialogRow -x $x -y ($y + 2) -width $width -backgroundColor $script:MoveDialogBg
				
				# Interval section header (row 3)
				Write-SimpleDialogRow -x $x -y ($y + 3) -width $width -content "Interval:" -contentColor $script:MoveDialogSectionTitle -backgroundColor $script:MoveDialogBg
				
				# Interval fields (rows 4-5)
				Write-SimpleFieldRow -x $x -y ($y + 4) -width $width `
					-label $fields[0].Label -longestLabel $longestLabel -fieldValue $fields[0].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[0].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				Write-SimpleFieldRow -x $x -y ($y + 5) -width $width `
					-label $fields[1].Label -longestLabel $longestLabel -fieldValue $fields[1].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[1].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				# Travel Distance section header (row 6)
				Write-SimpleDialogRow -x $x -y ($y + 6) -width $width -content "Travel Distance:" -contentColor $script:MoveDialogSectionTitle -backgroundColor $script:MoveDialogBg
				
				# Travel Distance fields (rows 7-8)
				Write-SimpleFieldRow -x $x -y ($y + 7) -width $width `
					-label $fields[2].Label -longestLabel $longestLabel -fieldValue $fields[2].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[2].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				Write-SimpleFieldRow -x $x -y ($y + 8) -width $width `
					-label $fields[3].Label -longestLabel $longestLabel -fieldValue $fields[3].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[3].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				# Movement Speed section header (row 9)
				Write-SimpleDialogRow -x $x -y ($y + 9) -width $width -content "Movement Speed:" -contentColor $script:MoveDialogSectionTitle -backgroundColor $script:MoveDialogBg
				
				# Movement Speed fields (rows 10-11)
				Write-SimpleFieldRow -x $x -y ($y + 10) -width $width `
					-label $fields[4].Label -longestLabel $longestLabel -fieldValue $fields[4].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[4].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				Write-SimpleFieldRow -x $x -y ($y + 11) -width $width `
					-label $fields[5].Label -longestLabel $longestLabel -fieldValue $fields[5].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[5].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				# Auto-Resume Delay section header (row 12)
				Write-SimpleDialogRow -x $x -y ($y + 12) -width $width -content "Auto-Resume Delay:" -contentColor $script:MoveDialogSectionTitle -backgroundColor $script:MoveDialogBg
				
				# Auto-Resume Delay field (row 13)
				Write-SimpleFieldRow -x $x -y ($y + 13) -width $width `
					-label $fields[6].Label -longestLabel $longestLabel -fieldValue $fields[6].Value `
					-fieldWidth $fieldWidth -fieldIndex $fields[6].Index -currentFieldIndex $currentFieldIndex -backgroundColor $script:MoveDialogBg
				
				# Empty line (row 14)
				Write-SimpleDialogRow -x $x -y ($y + 14) -width $width -backgroundColor $script:MoveDialogBg
				
				# Error line (row 15)
				if ($errorMsg) {
					Write-SimpleDialogRow -x $x -y ($y + 15) -width $width -content $errorMsg -contentColor $script:TextError -backgroundColor $script:MoveDialogBg
				} else {
					Write-SimpleDialogRow -x $x -y ($y + 15) -width $width -backgroundColor $script:MoveDialogBg
				}
				
			# Bottom line with buttons (row 16)
			$checkmark = [char]::ConvertFromUtf32(0x2705)
			$redX = [char]::ConvertFromUtf32(0x274C)
			$checkmarkX = $x + 2
			$redXX = $x + 15
			Write-Buffer -X $x -Y ($y + 16) -Text "$($script:BoxVertical)" -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
			Write-Buffer -Text " " -BG $script:MoveDialogBg
			Write-Buffer -X $checkmarkX -Y ($y + 16) -Text $checkmark -FG $script:TextSuccess -BG $script:MoveDialogButtonBg -Wide
		Write-Buffer -X ($checkmarkX + 2) -Y ($y + 16) -Text "|" -FG $script:MoveDialogButtonText -BG $script:MoveDialogButtonBg
		Write-Buffer -Text "(" -FG $script:MoveDialogButtonText -BG $script:MoveDialogButtonBg
		Write-Buffer -Text "u" -FG $script:MoveDialogButtonHotkey -BG $script:MoveDialogButtonBg
		Write-Buffer -Text ")pdate" -FG $script:MoveDialogButtonText -BG $script:MoveDialogButtonBg
		Write-Buffer -Text "  " -BG $script:MoveDialogBg
		Write-Buffer -X $redXX -Y ($y + 16) -Text $redX -FG $script:TextError -BG $script:MoveDialogButtonBg -Wide
		Write-Buffer -X ($redXX + 2) -Y ($y + 16) -Text "|" -FG $script:MoveDialogButtonText -BG $script:MoveDialogButtonBg
			Write-Buffer -Text "(" -FG $script:MoveDialogButtonText -BG $script:MoveDialogButtonBg
			Write-Buffer -Text "c" -FG $script:MoveDialogButtonHotkey -BG $script:MoveDialogButtonBg
			Write-Buffer -Text ")ancel" -FG $script:MoveDialogButtonText -BG $script:MoveDialogButtonBg
			# Button line display width calculation:
			# "$($script:BoxVertical) " (2) + checkmark (2) + "|" (1) + "(u)pdate" (8) + "  " (2) + redX (2) + "|" (1) + "(c)ancel" (8) = 26
			# So we need: 30 - 26 - 1 = 3 spaces before closing $($script:BoxVertical)
			$buttonPadding = 3
			Write-Buffer -Text (" " * $buttonPadding) -BG $script:MoveDialogBg
			Write-Buffer -Text "$($script:BoxVertical)" -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
				
			# Bottom border (spans full width)
			Write-Buffer -X $x -Y ($y + 17) -Text "$($script:BoxBottomLeft)" -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
			Write-Buffer -Text ("$($script:BoxHorizontal)" * ($width - 2)) -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
			Write-Buffer -Text "$($script:BoxBottomRight)" -FG $script:MoveDialogBorder -BG $script:MoveDialogBg
				
				# Draw drop shadow
				Draw-DialogShadow -dialogX $x -dialogY $y -dialogWidth $width -dialogHeight $height -shadowColor $script:MoveDialogShadow
			}
			
		# Initial draw
		& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
		Flush-Buffer
		
		# Calculate button bounds for click detection
			# Button row is at dialogY + 16 (row 16)
			$buttonRowY = $dialogY + 16
			# Rendered: +0:border +1:space +2-3:✅(2cells) +4:| +5-12:(u)pdate +13-14:spaces +15-16:❌(2cells) +17:| +18-25:(c)ancel +26-28:padding +29:border
			$updateButtonStartX = $dialogX + 2
			$updateButtonEndX = $dialogX + 12
			$cancelButtonStartX = $dialogX + 15
			$cancelButtonEndX = $dialogX + 25
			
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
					$updateButtonEndX = $dialogX + 12
					$cancelButtonStartX = $dialogX + 15
					$cancelButtonEndX = $dialogX + 25
					
					# Update button bounds in script scope
					$script:DialogButtonBounds = @{
						buttonRowY = $buttonRowY
						updateStartX = $updateButtonStartX
						updateEndX = $updateButtonEndX
						cancelStartX = $cancelButtonStartX
						cancelEndX = $cancelButtonEndX
					}
					
				& $drawDialog $dialogX $dialogY $dialogWidth $dialogHeight $currentField $errorMessage $inputBoxStartX $fieldWidth $intervalSecondsInput $intervalVarianceInput $moveSpeedInput $moveVarianceInput $travelDistanceInput $travelVarianceInput $autoResumeDelaySecondsInput
				Flush-Buffer -ClearFirst
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
				
				# Check for mouse button clicks on dialog buttons/fields via console input buffer
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				
				try {
					$peekBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' 16
					$peekEvts = [uint32]0
					$hIn = [mJiggAPI.Mouse]::GetStdHandle(-10)
					if ([mJiggAPI.Mouse]::PeekConsoleInput($hIn, $peekBuf, 16, [ref]$peekEvts) -and $peekEvts -gt 0) {
						$lastClickIdx = -1
						$clickX = -1; $clickY = -1
						for ($e = 0; $e -lt $peekEvts; $e++) {
							if ($peekBuf[$e].EventType -eq 0x0002 -and $peekBuf[$e].MouseEvent.dwEventFlags -eq 0 -and ($peekBuf[$e].MouseEvent.dwButtonState -band 0x0001) -ne 0) {
								$clickX = $peekBuf[$e].MouseEvent.dwMousePosition.X
								$clickY = $peekBuf[$e].MouseEvent.dwMousePosition.Y
								$lastClickIdx = $e
							}
						}
						if ($lastClickIdx -ge 0) {
							$consumeCount = [uint32]($lastClickIdx + 1)
							$flushBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' $consumeCount
							$flushed = [uint32]0
							[mJiggAPI.Mouse]::ReadConsoleInput($hIn, $flushBuf, $consumeCount, [ref]$flushed) | Out-Null
							
							$clickedField = -1
							if ($clickY -eq $buttonRowY -and $clickX -ge $updateButtonStartX -and $clickX -le $updateButtonEndX) {
								$char = "u"; $keyProcessed = $true
							} elseif ($clickY -eq $buttonRowY -and $clickX -ge $cancelButtonStartX -and $clickX -le $cancelButtonEndX) {
								$char = "c"; $keyProcessed = $true
							}
							if (-not $keyProcessed -and $clickX -ge $dialogX -and $clickX -lt ($dialogX + $dialogWidth)) {
								$clickFieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)
								for ($fi = 0; $fi -lt $clickFieldYOffsets.Count; $fi++) {
									$fy = $dialogY + $clickFieldYOffsets[$fi]
									if ($clickY -eq $fy) {
										$clickedField = $fi
										break
									}
								}
								if ($clickedField -ge 0 -and $clickedField -ne $currentField) {
									$previousField = $currentField
									$currentField = $clickedField
									$errorMessage = ""
									$lastFieldWithInput = -1
									
									$fieldLabels = @("Interval (sec): ", "Variance (sec): ", "Distance (px): ", "Variance (px): ", "Speed (sec): ", "Variance (sec): ", "Delay (sec): ")
									$fieldValues = @($intervalSecondsInput, $intervalVarianceInput, $travelDistanceInput, $travelVarianceInput, $moveSpeedInput, $moveVarianceInput, $autoResumeDelaySecondsInput)
									
									Write-SimpleFieldRow -x $dialogX -y ($dialogY + $clickFieldYOffsets[$previousField]) -width $dialogWidth `
										-label $fieldLabels[$previousField] -longestLabel $longestLabel -fieldValue $fieldValues[$previousField] `
										-fieldWidth $fieldWidth -fieldIndex $previousField -currentFieldIndex $currentField -backgroundColor DarkBlue
									
								Write-SimpleFieldRow -x $dialogX -y ($dialogY + $clickFieldYOffsets[$currentField]) -width $dialogWidth `
									-label $fieldLabels[$currentField] -longestLabel $longestLabel -fieldValue $fieldValues[$currentField] `
									-fieldWidth $fieldWidth -fieldIndex $currentField -currentFieldIndex $currentField -backgroundColor DarkBlue
								Flush-Buffer
								
								$fieldY = $dialogY + $clickFieldYOffsets[$currentField]
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
									$keyProcessed = $true
								}
							}
							if ($DebugMode) {
								$clickTarget = "none"
								if ($keyProcessed -and $char) { $clickTarget = "button:$char" }
								elseif ($clickedField -ge 0) { $clickTarget = "field:$clickedField" }
								if ($null -eq $LogArray -or -not ($LogArray -is [Array])) { $LogArray = @() }
								$LogArray += [PSCustomObject]@{
									logRow = $true
									components = @(
										@{ priority = 1; text = (Get-Date).ToString("HH:mm:ss"); shortText = (Get-Date).ToString("HH:mm:ss") },
										@{ priority = 2; text = " - [DEBUG] Movement dialog click at ($clickX,$clickY), target: $clickTarget"; shortText = " - [DEBUG] Click ($clickX,$clickY) -> $clickTarget" }
									)
								}
							}
						}
					}
				} catch { }
				
				# Check for dialog button clicks detected by main loop
				if (-not $keyProcessed -and $null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null
					if ($buttonClick -eq "Update") { $char = "u"; $keyProcessed = $true }
					elseif ($buttonClick -eq "Cancel") { $char = "c"; $keyProcessed = $true }
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
				Flush-Buffer
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
						$shiftState = [mJiggAPI.Mouse]::GetAsyncKeyState($VK_LSHIFT) -bor [mJiggAPI.Mouse]::GetAsyncKeyState($VK_RSHIFT)
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
				Flush-Buffer
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
							$script:CursorVisible = $false; [Console]::Write("$($script:ESC)[?25l")
						}
						$errorMessage = ""
						# Optimized: only redraw current field instead of entire dialog
						$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)
						$fieldLabels = @("Interval (sec): ", "Variance (sec): ", "Distance (px): ", "Variance (px): ", "Speed (sec): ", "Variance (sec): ", "Delay (sec): ")
					Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$currentField]) -width $dialogWidth `
						-label $fieldLabels[$currentField] -longestLabel $longestLabel -fieldValue $currentInputRef.Value `
						-fieldWidth $fieldWidth -fieldIndex $currentField -currentFieldIndex $currentField -backgroundColor DarkBlue
					Flush-Buffer
					# Position cursor at end of input
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
						$script:CursorVisible = $true; [Console]::Write("$($script:ESC)[?25h")
					} elseif ($currentInputRef.Value.Length -lt 6) {
						$currentInputRef.Value += $char.ToString()  # Convert char to string
					}
					$errorMessage = ""
					# Optimized: only redraw current field instead of entire dialog
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)
					$fieldLabels = @("Interval (sec): ", "Variance (sec): ", "Distance (px): ", "Variance (px): ", "Speed (sec): ", "Variance (sec): ", "Delay (sec): ")
				Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$currentField]) -width $dialogWidth `
					-label $fieldLabels[$currentField] -longestLabel $longestLabel -fieldValue $currentInputRef.Value `
					-fieldWidth $fieldWidth -fieldIndex $currentField -currentFieldIndex $currentField -backgroundColor DarkBlue
				Flush-Buffer
				# Position cursor at end of input
				$fieldY = $dialogY + $fieldYOffsets[$currentField]
				$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
				[Console]::SetCursorPosition($cursorX, $fieldY)
			} elseif ($char -eq ".") {
					# Decimal point for all fields - limit to 6 characters (including decimal point)
					# If this is the first character typed in this field, clear the field first
					$isFirstChar = ($lastFieldWithInput -ne $currentField)
					if ($isFirstChar) {
						$currentInputRef.Value = "."  # Already a string
						$lastFieldWithInput = $currentField
						$script:CursorVisible = $true; [Console]::Write("$($script:ESC)[?25h")
					} elseif ($currentInputRef.Value -notmatch "\." -and $currentInputRef.Value.Length -lt 6) {
						$currentInputRef.Value += "."  # Already a string
					}
					$errorMessage = ""
					# Optimized: only redraw current field instead of entire dialog
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)
					$fieldLabels = @("Interval (sec): ", "Variance (sec): ", "Distance (px): ", "Variance (px): ", "Speed (sec): ", "Variance (sec): ", "Delay (sec): ")
				Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$currentField]) -width $dialogWidth `
					-label $fieldLabels[$currentField] -longestLabel $longestLabel -fieldValue $currentInputRef.Value `
					-fieldWidth $fieldWidth -fieldIndex $currentField -currentFieldIndex $currentField -backgroundColor DarkBlue
				Flush-Buffer
				# Position cursor at end of input
				$fieldY = $dialogY + $fieldYOffsets[$currentField]
				$cursorX = $dialogX + $inputBoxStartX + 1 + $currentInputRef.Value.Length
				[Console]::SetCursorPosition($cursorX, $fieldY)
			} elseif ($key -eq "UpArrow" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 38)) {
					# UpArrow - move to previous field (reverse tab)
					$previousField = $currentField
					$currentField = ($currentField - 1 + 7) % 7
					$errorMessage = ""
					$lastFieldWithInput = -1  # Reset input tracking when switching fields
					
					# Optimized redraw: only update the two affected field rows instead of entire dialog
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldLabels = @("Interval (sec): ", "Variance (sec): ", "Distance (px): ", "Variance (px): ", "Speed (sec): ", "Variance (sec): ", "Delay (sec): ")
					$fieldValues = @($intervalSecondsInput, $intervalVarianceInput, $travelDistanceInput, $travelVarianceInput, $moveSpeedInput, $moveVarianceInput, $autoResumeDelaySecondsInput)
					
					# Redraw previous field (unhighlight)
					Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$previousField]) -width $dialogWidth `
						-label $fieldLabels[$previousField] -longestLabel $longestLabel -fieldValue $fieldValues[$previousField] `
						-fieldWidth $fieldWidth -fieldIndex $previousField -currentFieldIndex $currentField -backgroundColor DarkBlue
					
				# Redraw new field (highlight)
				Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$currentField]) -width $dialogWidth `
					-label $fieldLabels[$currentField] -longestLabel $longestLabel -fieldValue $fieldValues[$currentField] `
					-fieldWidth $fieldWidth -fieldIndex $currentField -currentFieldIndex $currentField -backgroundColor DarkBlue
				Flush-Buffer
				
				# Position cursor at the active field after arrow navigation
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
					$previousField = $currentField
					$currentField = ($currentField + 1) % 7
					$errorMessage = ""
					$lastFieldWithInput = -1  # Reset input tracking when switching fields
					
					# Optimized redraw: only update the two affected field rows instead of entire dialog
					$fieldYOffsets = @(4, 5, 7, 8, 10, 11, 13)  # Y offsets for each field relative to dialogY
					$fieldLabels = @("Interval (sec): ", "Variance (sec): ", "Distance (px): ", "Variance (px): ", "Speed (sec): ", "Variance (sec): ", "Delay (sec): ")
					$fieldValues = @($intervalSecondsInput, $intervalVarianceInput, $travelDistanceInput, $travelVarianceInput, $moveSpeedInput, $moveVarianceInput, $autoResumeDelaySecondsInput)
					
					# Redraw previous field (unhighlight)
					Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$previousField]) -width $dialogWidth `
						-label $fieldLabels[$previousField] -longestLabel $longestLabel -fieldValue $fieldValues[$previousField] `
						-fieldWidth $fieldWidth -fieldIndex $previousField -currentFieldIndex $currentField -backgroundColor DarkBlue
					
				# Redraw new field (highlight)
				Write-SimpleFieldRow -x $dialogX -y ($dialogY + $fieldYOffsets[$currentField]) -width $dialogWidth `
					-label $fieldLabels[$currentField] -longestLabel $longestLabel -fieldValue $fieldValues[$currentField] `
					-fieldWidth $fieldWidth -fieldIndex $currentField -currentFieldIndex $currentField -backgroundColor DarkBlue
				Flush-Buffer
				
				# Position cursor at the active field after arrow navigation
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
			Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth)
		}
		Flush-Buffer
			
			$script:CursorVisible = $savedCursorVisible
			if ($script:CursorVisible) { [Console]::Write("$($script:ESC)[?25h") } else { [Console]::Write("$($script:ESC)[?25l") }
			
			$script:DialogButtonBounds = $null
			$script:DialogButtonClick = $null
			
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
			$dialogHeight = 7
			$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth - $dialogWidth) / 2))
			$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
			
			$savedCursorVisible = $script:CursorVisible
			$script:CursorVisible = $false
			[Console]::Write("$($script:ESC)[?25l")
			
			# Draw dialog box (exactly 35 characters per line)
			$checkmark = [char]::ConvertFromUtf32(0x2705)  # ✅ green checkmark
			$redX = [char]::ConvertFromUtf32(0x274C)  # ❌ red X
			# Button line display width calculation:
			# "$($script:BoxVertical) " (2) + checkmark (2) + "|" (1) + "(y)es" (5) + "  " (2) + redX (2) + "|" (1) + "(n)o" (4) = 19
			# So we need: 35 - 19 - 1 = 15 spaces before closing $($script:BoxVertical)
			$bottomLinePadding = 15
			
			# Build all lines to be exactly 35 characters using Get-Padding helper
			$line0 = "$($script:BoxTopLeft)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxTopRight)"  # 35 chars
			$line1Text = "$($script:BoxVertical)  Confirm Quit"
			$line1Padding = Get-Padding -usedWidth ($line1Text.Length + 1) -totalWidth $dialogWidth
			$line1 = $line1Text + (" " * $line1Padding) + "$($script:BoxVertical)"
			
			$line2 = "$($script:BoxVertical)" + (" " * 33) + "$($script:BoxVertical)"  # 35 chars
			
			$line3Text = "$($script:BoxVertical)  Are you sure you want to quit?"
			$line3Padding = Get-Padding -usedWidth ($line3Text.Length + 1) -totalWidth $dialogWidth
			$line3 = $line3Text + (" " * $line3Padding) + "$($script:BoxVertical)"
			
			$line4 = "$($script:BoxVertical)" + (" " * 33) + "$($script:BoxVertical)"  # 35 chars
			$line5 = "$($script:BoxVertical)" + (" " * 33) + "$($script:BoxVertical)"  # 35 chars
			
			$dialogLines = @(
				$line0,
				$line1,
				$line2,
				$line3,
				$line4,
				$line5,
				$null,  # Bottom line will be written separately with colors
				"$($script:BoxBottomLeft)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxHorizontal)$($script:BoxBottomRight)"  # 35 chars
			)
			
			# Draw dialog background (clear area) with magenta background
			for ($i = 0; $i -lt $dialogHeight; $i++) {
				Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth) -BG DarkMagenta
			}
			
			# Draw dialog box with themed background
			for ($i = 0; $i -lt $dialogLines.Count; $i++) {
				if ($i -eq 1) {
					# Title line
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical)  " -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
					Write-Buffer -Text "Confirm Quit" -FG $script:QuitDialogTitle -BG $script:QuitDialogBg
					$titleUsedWidth = 3 + "Confirm Quit".Length  # "$($script:BoxVertical)  " + title
					$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
					Write-Buffer -Text (" " * $titlePadding) -BG $script:QuitDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
				} elseif ($i -eq 6) {
					# Bottom line - write with colored icons and hotkey letters
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical) " -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
					$checkmarkX = $dialogX + 2
					Write-Buffer -X $checkmarkX -Y ($dialogY + $i) -Text $checkmark -FG $script:TextSuccess -BG $script:QuitDialogButtonBg -Wide
					Write-Buffer -X ($checkmarkX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
					# Parse "(y)es" - parentheses white, letter yellow, text white
					Write-Buffer -Text "(" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
					Write-Buffer -Text "y" -FG $script:QuitDialogButtonHotkey -BG $script:QuitDialogButtonBg
					Write-Buffer -Text ")es" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
					Write-Buffer -Text "  " -BG $script:QuitDialogBg
					$redXX = $dialogX + 12
					Write-Buffer -X $redXX -Y ($dialogY + $i) -Text $redX -FG $script:TextError -BG $script:QuitDialogButtonBg -Wide
					Write-Buffer -X ($redXX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
					# Parse "(n)o" - parentheses white, letter yellow, text white
					Write-Buffer -Text "(" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
					Write-Buffer -Text "n" -FG $script:QuitDialogButtonHotkey -BG $script:QuitDialogButtonBg
					Write-Buffer -Text ")o" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
					Write-Buffer -Text (" " * $bottomLinePadding) -BG $script:QuitDialogBg
					Write-Buffer -Text "$($script:BoxVertical)" -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
				} else {
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text $dialogLines[$i] -FG $script:QuitDialogText -BG $script:QuitDialogBg
				}
			}
			
			# Draw drop shadow
			Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:QuitDialogShadow
			Flush-Buffer
			
			# Calculate button bounds for click detection (visible characters only)
			# Button row is at dialogY + 6
			# Rendered: +0:border +1:space +2-3:✅(2cells) +4:| +5-9:(y)es +10-11:spaces +12-13:❌(2cells) +14:| +15-18:(n)o +19-33:padding +34:border
			$buttonRowY = $dialogY + 6
			$yesButtonStartX = $dialogX + 2
			$yesButtonEndX = $dialogX + 9
			$noButtonStartX = $dialogX + 12
			$noButtonEndX = $dialogX + 18
			
			$script:DialogButtonBounds = @{
				buttonRowY = $buttonRowY
				updateStartX = $yesButtonStartX
				updateEndX = $yesButtonEndX
				cancelStartX = $noButtonStartX
				cancelEndX = $noButtonEndX
			}
			$script:DialogButtonClick = $null
			
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
					
					for ($i = 0; $i -lt $dialogLines.Count; $i++) {
						if ($i -eq 1) {
							# Title line
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical)  " -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
							Write-Buffer -Text "Confirm Quit" -FG $script:QuitDialogTitle -BG $script:QuitDialogBg
							$titleUsedWidth = 3 + "Confirm Quit".Length  # "$($script:BoxVertical)  " + title
							$titlePadding = Get-Padding -usedWidth ($titleUsedWidth + 1) -totalWidth $dialogWidth
							Write-Buffer -Text (" " * $titlePadding) -BG $script:QuitDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
						} elseif ($i -eq 6) {
							# Bottom line - write with colored icons and hotkey letters
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text "$($script:BoxVertical) " -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
							$checkmarkX = $dialogX + 2
							Write-Buffer -X $checkmarkX -Y ($dialogY + $i) -Text $checkmark -FG $script:TextSuccess -BG $script:QuitDialogButtonBg -Wide
							Write-Buffer -X ($checkmarkX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
							# Parse "(y)es" - parentheses white, letter yellow, text white
							Write-Buffer -Text "(" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
							Write-Buffer -Text "y" -FG $script:QuitDialogButtonHotkey -BG $script:QuitDialogButtonBg
							Write-Buffer -Text ")es" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
							Write-Buffer -Text "  " -BG $script:QuitDialogBg
							$redXX = $dialogX + 12
							Write-Buffer -X $redXX -Y ($dialogY + $i) -Text $redX -FG $script:TextError -BG $script:QuitDialogButtonBg -Wide
							Write-Buffer -X ($redXX + 2) -Y ($dialogY + $i) -Text "|" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
							# Parse "(n)o" - parentheses white, letter yellow, text white
							Write-Buffer -Text "(" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
							Write-Buffer -Text "n" -FG $script:QuitDialogButtonHotkey -BG $script:QuitDialogButtonBg
							Write-Buffer -Text ")o" -FG $script:QuitDialogButtonText -BG $script:QuitDialogButtonBg
							Write-Buffer -Text (" " * $bottomLinePadding) -BG $script:QuitDialogBg
							Write-Buffer -Text "$($script:BoxVertical)" -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
						} else {
							Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text $dialogLines[$i] -FG $script:QuitDialogText -BG $script:QuitDialogBg
						}
					}
					
					# Draw drop shadow
					Draw-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:QuitDialogShadow
					Flush-Buffer -ClearFirst
					
					$buttonRowY = $dialogY + 6
					$yesButtonStartX = $dialogX + 2
					$yesButtonEndX = $dialogX + 9
					$noButtonStartX = $dialogX + 12
					$noButtonEndX = $dialogX + 18
					
					$script:DialogButtonBounds = @{
						buttonRowY = $buttonRowY
						updateStartX = $yesButtonStartX
						updateEndX = $yesButtonEndX
						cancelStartX = $noButtonStartX
						cancelEndX = $noButtonEndX
					}
				}
				
				# Check for mouse button clicks on dialog buttons via console input buffer
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				
				try {
					$peekBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' 16
					$peekEvts = [uint32]0
					$hIn = [mJiggAPI.Mouse]::GetStdHandle(-10)
					if ([mJiggAPI.Mouse]::PeekConsoleInput($hIn, $peekBuf, 16, [ref]$peekEvts) -and $peekEvts -gt 0) {
						$lastClickIdx = -1
						$clickX = -1; $clickY = -1
						for ($e = 0; $e -lt $peekEvts; $e++) {
							if ($peekBuf[$e].EventType -eq 0x0002 -and $peekBuf[$e].MouseEvent.dwEventFlags -eq 0 -and ($peekBuf[$e].MouseEvent.dwButtonState -band 0x0001) -ne 0) {
								$clickX = $peekBuf[$e].MouseEvent.dwMousePosition.X
								$clickY = $peekBuf[$e].MouseEvent.dwMousePosition.Y
								$lastClickIdx = $e
							}
						}
						if ($lastClickIdx -ge 0) {
							$consumeCount = [uint32]($lastClickIdx + 1)
							$flushBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' $consumeCount
							$flushed = [uint32]0
							[mJiggAPI.Mouse]::ReadConsoleInput($hIn, $flushBuf, $consumeCount, [ref]$flushed) | Out-Null
							
							if ($clickY -eq $buttonRowY -and $clickX -ge $yesButtonStartX -and $clickX -le $yesButtonEndX) {
								$char = "y"; $keyProcessed = $true
							} elseif ($clickY -eq $buttonRowY -and $clickX -ge $noButtonStartX -and $clickX -le $noButtonEndX) {
								$char = "n"; $keyProcessed = $true
							}
							if ($DebugMode) {
								$clickTarget = if ($keyProcessed) { "button:$char" } else { "none" }
								if ($null -eq $LogArray -or -not ($LogArray -is [Array])) { $LogArray = @() }
								$LogArray += [PSCustomObject]@{
									logRow = $true
									components = @(
										@{ priority = 1; text = (Get-Date).ToString("HH:mm:ss"); shortText = (Get-Date).ToString("HH:mm:ss") },
										@{ priority = 2; text = " - [DEBUG] Quit dialog click at ($clickX,$clickY), target: $clickTarget"; shortText = " - [DEBUG] Click ($clickX,$clickY) -> $clickTarget" }
									)
								}
							}
						}
					}
				} catch { }
				
				# Check for dialog button clicks detected by main loop
				if (-not $keyProcessed -and $null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null
					if ($buttonClick -eq "Update") { $char = "y"; $keyProcessed = $true }
					elseif ($buttonClick -eq "Cancel") { $char = "n"; $keyProcessed = $true }
				}
				
				# Wait for key input (non-blocking check)
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
				Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth)
			}
			Flush-Buffer
			
			$script:CursorVisible = $savedCursorVisible
			if ($script:CursorVisible) { [Console]::Write("$($script:ESC)[?25h") } else { [Console]::Write("$($script:ESC)[?25l") }
			
			$script:DialogButtonBounds = $null
			$script:DialogButtonClick = $null
			
			# Return result object with result and redraw flag
			return @{
				Result = $result
				NeedsRedraw = $needsRedraw
			}
		}

		# Initialization complete - pause to read debug output if in debug mode
		if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 4: Before debug mode check (DebugMode=$DebugMode)" | Out-File $script:StartupDiagFile -Append }
		
		if ($DebugMode) {
			if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - ENTERED DEBUG MODE KEY WAIT LOOP" | Out-File $script:StartupDiagFile -Append }

			Write-Host "`nPress any key to start mJig..." -ForegroundColor $script:TextWarning
			
			$keyPressed = $false
			
			$VK_LCONTROL = 0xA2
			$VK_RCONTROL = 0xA3
			
			while (-not $keyPressed) {
				$ctrlPressedAlone = $false
				try {
					$ctrlState = [mJiggAPI.Mouse]::GetAsyncKeyState($VK_LCONTROL) -bor [mJiggAPI.Mouse]::GetAsyncKeyState($VK_RCONTROL)
					$ctrlPressed = (($ctrlState -band 0x8000) -ne 0)
					if ($ctrlPressed) {
						$ctrlPressedAlone = $true
					}
				} catch {}
				
				if ($ctrlPressedAlone) {
					Start-Sleep -Milliseconds 50
					continue
				}
				
				try {
					if ($Host.UI.RawUI.KeyAvailable) {
						$validKeyFound = $false
						while ($Host.UI.RawUI.KeyAvailable -and -not $validKeyFound) {
							$keyInfo = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
							$isKeyDown = if ($null -ne $keyInfo.KeyDown) { $keyInfo.KeyDown } else { $true }
							$isCtrlOnly = ($keyInfo.Key -eq "LeftCtrl" -or $keyInfo.Key -eq "RightCtrl")
							
							if ($isKeyDown -and -not $isCtrlOnly) {
								$keyPressed = $true
								$validKeyFound = $true
							}
						}
					} else {
						Start-Sleep -Milliseconds 50
					}
				} catch {
					Start-Sleep -Milliseconds 50
				}
			}

		}

		# Note: Previously had code here to clear key buffer, but $Host.UI.RawUI.KeyAvailable 
		# can return false positives and ReadKey blocks waiting for input, causing startup freeze.
		# The main loop handles stale input gracefully, so pre-clearing is not necessary.
		if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 5: Skipping key buffer clear (causes blocking)" | Out-File $script:StartupDiagFile -Append }
		if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 6: Entering main loop" | Out-File $script:StartupDiagFile -Append }

		# Main Processing Loop
		:process while ($true) {
			$script:LoopIteration++
			
			# Reset state for this iteration
			$time = $false
			$script:userInputDetected = $false
			$keyboardInputDetected = $false
			$mouseInputDetected = $false
			$scrollDetectedInInterval = $false
			$waitExecuted = $false
			$intervalMouseInputs = @()
			$interval = 0
			$math = 0
			$date = Get-Date
			$currentTime = $date.ToString("HHmm")
			$forceRedraw = $false
			$automatedMovementPos = $null  # Track position after automated movement
			$directionArrow = ""  # Track direction arrow for log display
			$lastKeyPress = $null  # Reset key press tracking
			$lastKeyInfo = $null  # Reset key info tracking
			$pressedMenuKeys = @{}  # Track which menu keys are currently pressed (to detect key up)
			
			# Calculate interval and wait BEFORE doing movement (skip on first run or if forceRedraw)
			if ($null -ne $LastMovementTime -and -not $forceRedraw) {
				# Calculate random interval with variance
				# Convert to milliseconds for calculation
				$intervalSecondsMs = $script:IntervalSeconds * 1000
				$intervalVarianceMs = $script:IntervalVariance * 1000
				$intervalMs = Get-ValueWithVariance -baseValue $intervalSecondsMs -variance $intervalVarianceMs
				
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
				$mousePosAtStart = Get-MousePosition
				
				# Wait Loop - Check window/buffer size changes, menu hotkeys, and keyboard input
				# Menu hotkeys checked every 200ms (every 4th iteration), keyboard input checked every 50ms for maximum reliability
				$x = 0
				:waitLoop while ($true) {
					$x++
					
					# Check for system-wide keyboard input every 50ms for maximum reliability
					# Skip checking if we recently sent a simulated key press (within last 300ms)
					$shouldCheckKeyboard = (Get-TimeSinceMs -startTime $LastSimulatedKeyPress) -ge 300
					if ($shouldCheckKeyboard) {
						$LastSimulatedKeyPress = $null
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
							$currentCheckPos = Get-MousePosition
							if ($script:DiagEnabled -and $null -ne $currentCheckPos) {
								$lastX = if ($null -ne $script:lastMousePosCheck) { $script:lastMousePosCheck.X } else { "null" }
								$lastY = if ($null -ne $script:lastMousePosCheck) { $script:lastMousePosCheck.Y } else { "null" }
								$moved = Test-MouseMoved -currentPos $currentCheckPos -lastPos $script:lastMousePosCheck -threshold 2
								"$(Get-Date -Format 'HH:mm:ss.fff') - MOUSEPOS cur=($($currentCheckPos.X),$($currentCheckPos.Y)) last=($lastX,$lastY) moved=$moved" | Out-File $script:InputDiagFile -Append
							}
							if ($null -ne $currentCheckPos) {
								if (Test-MouseMoved -currentPos $currentCheckPos -lastPos $script:lastMousePosCheck -threshold 2) {
									$script:LastMouseMovementTime = Get-Date
									$mouseInputDetected = $true
									$mouseMoveText = "Mouse"
									if ($intervalMouseInputs -notcontains $mouseMoveText) {
										$intervalMouseInputs += $mouseMoveText
									}
									if ($script:AutoResumeDelaySeconds -gt 0) {
										$LastUserInputTime = Get-Date
									}
								}
								$script:lastMousePosCheck = $currentCheckPos
							} elseif ($script:DiagEnabled) {
								"$(Get-Date -Format 'HH:mm:ss.fff') - MOUSEPOS: Get-MousePosition returned NULL" | Out-File $script:InputDiagFile -Append
							}
						} catch {
							if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - MOUSEPOS ERROR: $($_.Exception.Message)" | Out-File $script:InputDiagFile -Append }
						}
						
						# Detect scroll, keyboard, and mouse clicks via PeekConsoleInput (works when console is focused)
						# Keyboard events are only peeked (not consumed) so the menu hotkey handler can still read them
						$scrollDetected = $false
						$script:ConsoleClickCoords = $null
						try {
							$peekBuffer = New-Object 'mJiggAPI.INPUT_RECORD[]' 32
							$peekEvents = [uint32]0
							$hStdIn = [mJiggAPI.Mouse]::GetStdHandle(-10)  # STD_INPUT_HANDLE
							if ([mJiggAPI.Mouse]::PeekConsoleInput($hStdIn, $peekBuffer, 32, [ref]$peekEvents) -and $peekEvents -gt 0) {
								$hasScrollEvent = $false
								$hasKeyboardEvent = $false
								$lastScrollIdx = -1
								$lastClickIdx = -1
								for ($e = 0; $e -lt $peekEvents; $e++) {
									if ($peekBuffer[$e].EventType -eq 0x0002) {
										$mouseFlags = $peekBuffer[$e].MouseEvent.dwEventFlags
										$mouseButtons = $peekBuffer[$e].MouseEvent.dwButtonState
										if ($mouseFlags -eq 0x0004) {
											$hasScrollEvent = $true
											$lastScrollIdx = $e
										} elseif ($mouseFlags -eq 0 -and ($mouseButtons -band 0x0001) -ne 0) {
											# Left button press (dwEventFlags=0 means button press/release, bit 0 = left button)
											$script:ConsoleClickCoords = @{
												X = $peekBuffer[$e].MouseEvent.dwMousePosition.X
												Y = $peekBuffer[$e].MouseEvent.dwMousePosition.Y
											}
											$lastClickIdx = $e
										}
									}
									if ($peekBuffer[$e].EventType -eq 0x0001 -and $peekBuffer[$e].KeyEvent.wVirtualKeyCode -ne 0xA5) {
										$hasKeyboardEvent = $true
									}
								}
								# Consume scroll and click events to prevent buffer buildup
								$maxConsumeIdx = [Math]::Max($lastScrollIdx, $lastClickIdx)
								if ($maxConsumeIdx -ge 0) {
									$consumeCount = [uint32]($maxConsumeIdx + 1)
									$flushBuffer = New-Object 'mJiggAPI.INPUT_RECORD[]' $consumeCount
									$flushed = [uint32]0
									[mJiggAPI.Mouse]::ReadConsoleInput($hStdIn, $flushBuffer, $consumeCount, [ref]$flushed) | Out-Null
								}
								if ($hasScrollEvent) {
									$scrollDetected = $true
									$scrollDetectedInInterval = $true
									$otherText = "Scroll/Other"
									if ($intervalMouseInputs -notcontains $otherText) {
										$intervalMouseInputs += $otherText
									}
									$mouseInputDetected = $true
									$script:userInputDetected = $true
									if ($script:AutoResumeDelaySeconds -gt 0) {
										$LastUserInputTime = Get-Date
									}
									if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - PeekConsoleInput: scroll detected (events=$peekEvents)" | Out-File $script:InputDiagFile -Append }
								}
								if ($hasKeyboardEvent) {
									$keyboardInputDetected = $true
									$script:userInputDetected = $true
									if ($script:AutoResumeDelaySeconds -gt 0) {
										$LastUserInputTime = Get-Date
									}
									if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - PeekConsoleInput: keyboard detected (events=$peekEvents)" | Out-File $script:InputDiagFile -Append }
								}
							}
						} catch {
							if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - PeekConsoleInput ERROR: $($_.Exception.Message)" | Out-File $script:InputDiagFile -Append }
						}
						
						# Detect user input via GetLastInputInfo (system-wide, passive)
						# Keyboard and scroll are evidence-based (PeekConsoleInput).
						# If GetLastInputInfo sees activity that wasn't classified as keyboard or scroll,
						# it's almost certainly mouse movement.
						try {
							$lii = New-Object mJiggAPI.LASTINPUTINFO
							$lii.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf([type][mJiggAPI.LASTINPUTINFO])
							$liiResult = [mJiggAPI.Mouse]::GetLastInputInfo([ref]$lii)
							if ($liiResult) {
								$tickNow = [uint64][mJiggAPI.Mouse]::GetTickCount64()
								$lastInputTick = [uint64]$lii.dwTime
								$systemIdleMs = $tickNow - $lastInputTick
								$recentSimulated = ($null -ne $LastSimulatedKeyPress) -and ((Get-TimeSinceMs -startTime $LastSimulatedKeyPress) -lt 500)
								$recentAutoMove = ($null -ne $LastAutomatedMouseMovement) -and ((Get-TimeSinceMs -startTime $LastAutomatedMouseMovement) -lt 500)
								if ($script:DiagEnabled) {
									$ts = Get-Date -Format 'HH:mm:ss.fff'
									"$ts - LII idleMs=$systemIdleMs simFilter=$recentSimulated autoFilter=$recentAutoMove kbDet=$keyboardInputDetected msDet=$mouseInputDetected scrollInt=$scrollDetectedInInterval" | Out-File $script:InputDiagFile -Append
								}
								if ($systemIdleMs -lt 300 -and -not $recentSimulated -and -not $recentAutoMove) {
									$script:userInputDetected = $true
									if ($script:AutoResumeDelaySeconds -gt 0) {
										$LastUserInputTime = Get-Date
									}
									if (-not $keyboardInputDetected -and -not $scrollDetectedInInterval -and -not $mouseInputDetected) {
										$mouseInputDetected = $true
										$script:LastMouseMovementTime = Get-Date
										$mouseMoveText = "Mouse"
										if ($intervalMouseInputs -notcontains $mouseMoveText) {
											$intervalMouseInputs += $mouseMoveText
										}
										if ($script:DiagEnabled) { "  >> userInput=TRUE idleMs=$systemIdleMs -> mouse (no kb/scroll/click evidence)" | Out-File $script:InputDiagFile -Append }
									} else {
										if ($script:DiagEnabled) { "  >> userInput=TRUE idleMs=$systemIdleMs (already classified: kb=$keyboardInputDetected ms=$mouseInputDetected scroll=$scrollDetectedInInterval)" | Out-File $script:InputDiagFile -Append }
									}
								}
							}
						} catch {
							if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - GetLastInputInfo ERROR: $($_.Exception.Message)" | Out-File $script:InputDiagFile -Append }
						}
						
						# Check for left-click via console input buffer (exact cell coordinates from the console)
						if ($null -ne $script:ConsoleClickCoords) {
							$consoleX = $script:ConsoleClickCoords.X
							$consoleY = $script:ConsoleClickCoords.Y
							
							# Check dialog buttons first (if a dialog is open)
							if ($null -ne $script:DialogButtonBounds) {
								$bounds = $script:DialogButtonBounds
								if ($consoleY -eq $bounds.buttonRowY -and $consoleX -ge $bounds.updateStartX -and $consoleX -le $bounds.updateEndX) {
									$script:DialogButtonClick = "Update"
								} elseif ($consoleY -eq $bounds.buttonRowY -and $consoleX -ge $bounds.cancelStartX -and $consoleX -le $bounds.cancelEndX) {
									$script:DialogButtonClick = "Cancel"
								}
							}
							
							# Check menu items (only when no dialog is open)
							if ($null -eq $script:DialogButtonBounds -and $null -ne $script:MenuItemsBounds -and $script:MenuItemsBounds.Count -gt 0) {
								foreach ($menuItem in $script:MenuItemsBounds) {
									if ($null -ne $menuItem.hotkey -and $consoleY -eq $menuItem.y -and $consoleX -ge $menuItem.startX -and $consoleX -le $menuItem.endX) {
										$script:MenuClickHotkey = $menuItem.hotkey
										break
									}
								}
							}
							
							if ($DebugMode) {
								$clickTarget = "none"
								if ($null -ne $script:DialogButtonClick) { $clickTarget = "Dialog:$($script:DialogButtonClick)" }
								elseif ($null -ne $script:MenuClickHotkey) { $clickTarget = "Menu:$($script:MenuClickHotkey)" }
								if ($null -eq $LogArray -or -not ($LogArray -is [Array])) { $LogArray = @() }
								$LogArray += [PSCustomObject]@{
									logRow = $true
									components = @(
										@{ priority = 1; text = (Get-Date).ToString("HH:mm:ss"); shortText = (Get-Date).ToString("HH:mm:ss") },
										@{ priority = 2; text = " - [DEBUG] LButton click at console ($consoleX,$consoleY), target: $clickTarget"; shortText = " - [DEBUG] Click ($consoleX,$consoleY) -> $clickTarget" }
									)
								}
							}
						}
						
						# Check mouse buttons (0x01-0x06) for input detection (pause jiggler)
						for ($keyCode = 0x01; $keyCode -le 0x06; $keyCode++) {
							if ($keyCode -eq 0x03) { continue }  # 0x03 is VK_CANCEL, not a mouse button
							$currentKeyState = [mJiggAPI.Mouse]::GetAsyncKeyState($keyCode)
							$isCurrentlyPressed = (($currentKeyState -band 0x8000) -ne 0)
							$wasJustPressed = (($currentKeyState -band 0x0001) -ne 0)
							$wasPreviouslyPressed = if ($script:previousKeyStates.ContainsKey($keyCode)) { $script:previousKeyStates[$keyCode] } else { $false }
							
							if ($wasJustPressed -or ($isCurrentlyPressed -and -not $wasPreviouslyPressed)) {
								
								$mouseButtonName = switch ($keyCode) {
									0x01 { "LButton" }
									0x02 { "RButton" }
									0x04 { "MButton" }
									0x05 { "XButton1" }
									0x06 { "XButton2" }
								}
								if ($mouseButtonName -and $intervalMouseInputs -notcontains $mouseButtonName) {
									$intervalMouseInputs += $mouseButtonName
									$script:userInputDetected = $true
									$mouseInputDetected = $true
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
								$mouseEmoji = [char]::ConvertFromUtf32(0x1F400)
								Write-Host "  mJig($mouseEmoji) " -NoNewline -ForegroundColor $script:HeaderAppName
								Write-Host "Stopped" -ForegroundColor $script:TextError
								Write-Host ""
								Write-Host "  Runtime: " -NoNewline -ForegroundColor $script:StatsLabel
								Write-Host $runtimeStr -ForegroundColor $script:StatsValue
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
									$mouseEmoji = [char]::ConvertFromUtf32(0x1F400)
									Write-Host "  mJig(" -NoNewline -ForegroundColor $script:HeaderAppName
									$mouseEmojiX = $Host.UI.RawUI.CursorPosition.X
									$mouseEmojiY = $Host.UI.RawUI.CursorPosition.Y
									Write-Host $mouseEmoji -NoNewline -ForegroundColor $script:HeaderIcon
									[Console]::SetCursorPosition($mouseEmojiX + 2, $mouseEmojiY)
									Write-Host ") " -NoNewline -ForegroundColor $script:HeaderAppName
									Write-Host "Stopped" -ForegroundColor $script:TextError
									Write-Host ""
									Write-Host "  Runtime: " -NoNewline -ForegroundColor $script:StatsLabel
									Write-Host $runtimeStr -ForegroundColor $script:StatsValue
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
												priority = [int]2
												text = " - [DEBUG] View toggle: $oldOutput $([char]0x2192) $Output"
												shortText = " - [DEBUG] View: $oldOutput $([char]0x2192) $Output"
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
								$script:MenuItemsBounds = @()
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
												priority = [int]2
												text = " - [DEBUG] Hide/Show toggle: $oldOutput $([char]0x2192) $Output"
												shortText = " - [DEBUG] Hide/Show: $oldOutput $([char]0x2192) $Output"
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
									$arrowChar = [char]0x2192
									if ($oldIntervalSeconds -ne $script:IntervalSeconds) { $changeDetails += "Interval: $oldIntervalSeconds $arrowChar $($script:IntervalSeconds)" }
									if ($oldIntervalVariance -ne $script:IntervalVariance) { $changeDetails += "IntervalVar: $oldIntervalVariance $arrowChar $($script:IntervalVariance)" }
									if ($oldMoveSpeed -ne $script:MoveSpeed) { $changeDetails += "Speed: $oldMoveSpeed $arrowChar $($script:MoveSpeed)" }
									if ($oldMoveVariance -ne $script:MoveVariance) { $changeDetails += "SpeedVar: $oldMoveVariance $arrowChar $($script:MoveVariance)" }
									if ($oldTravelDistance -ne $script:TravelDistance) { $changeDetails += "Distance: $oldTravelDistance $arrowChar $($script:TravelDistance)" }
									if ($oldTravelVariance -ne $script:TravelVariance) { $changeDetails += "DistVar: $oldTravelVariance $arrowChar $($script:TravelVariance)" }
									if ($oldAutoResumeDelaySeconds -ne $script:AutoResumeDelaySeconds) { $changeDetails += "Delay: $oldAutoResumeDelaySeconds $arrowChar $($script:AutoResumeDelaySeconds)" }
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
										$isTomorrow = $endTimeInt -le [int]$currentTime
										if ($isTomorrow) {
											$tommorow = (Get-Date).AddDays(1)
											$endDate = Get-Date $tommorow -Format "MMdd"
										} else {
											$endDate = Get-Date -Format "MMdd"
										}
										$end = "$endDate$endTimeStr"
										$changeDate = Get-Date
										$arrowChar = [char]0x2192
										$dayLabel = if ($isTomorrow) { " (tomorrow)" } else { " (today)" }
										$endDateDisplay = $endDate.Substring(0,2) + "/" + $endDate.Substring(2,2)
										$endTimeDisplay = $endTimeStr.Substring(0,2) + ":" + $endTimeStr.Substring(2,2)
										$changeMessage = if ($oldEndTimeInt -eq -1 -or [string]::IsNullOrEmpty($oldEndTimeStr)) {" - End time set: $endDateDisplay $endTimeDisplay$dayLabel"} else {" - End time changed: $oldEndTimeStr $arrowChar $endDateDisplay $endTimeDisplay$dayLabel"}
										$changeShortMessage = " - End time: $endDateDisplay $endTimeDisplay"
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
								# Clear screen once when resize starts
								if (-not $ResizeClearedScreen) {
									$ResizeClearedScreen = $true
									$script:CurrentResizeQuote = $null
									Draw-ResizeLogo -ClearFirst
									$LastResizeLogoTime = Get-Date
								}
								# New size detected - store it and reset timer
								$PendingResize = $newWindowSize
								$lastResizeDetection = Get-Date
							}
						}
						
						# Tight redraw loop while resizing - bypasses the normal 50ms sleep
						while ($ResizeClearedScreen) {
							# Check for size changes
							$newWindowSize = $pswindow.WindowSize
							$isNewSize = ($null -eq $PendingResize -or 
								$newWindowSize.Width -ne $PendingResize.Width -or 
								$newWindowSize.Height -ne $PendingResize.Height)
							
							if ($isNewSize) {
								$PendingResize = $newWindowSize
								$lastResizeDetection = Get-Date
								Draw-ResizeLogo -ClearFirst
								$LastResizeLogoTime = Get-Date
							}
							
							# Check if resize is complete (stable for threshold)
							$sizeMatchesPending = ($newWindowSize.Width -eq $PendingResize.Width -and 
								$newWindowSize.Height -eq $PendingResize.Height)
							if ($sizeMatchesPending) {
								$timeSinceResize = Get-TimeSinceMs -startTime $lastResizeDetection
								if ($timeSinceResize -ge $ResizeThrottleMs) {
									# Resize complete - exit tight loop
									$currentWindowSize = $pswindow.WindowSize
									$currentBufferSize = $pswindow.BufferSize
									try {
										$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($currentBufferSize.Width, $currentWindowSize.Height)
										$currentBufferSize = $pswindow.BufferSize
									} catch { }
									$OldBufferSize = $currentBufferSize
									$oldWindowSize = $currentWindowSize
									$HostWidth = $currentWindowSize.Width
									$HostHeight = $currentWindowSize.Height
									$SkipUpdate = $true
									$forceRedraw = $true
									$waitExecuted = $false
									$PendingResize = $null
									$lastResizeDetection = $null
									$ResizeClearedScreen = $false
									$LastResizeLogoTime = $null
									break
								}
							}
							
							# Small sleep to prevent CPU spinning (1ms)
							Start-Sleep -Milliseconds 1
						}
						
						# If resize completed, break out of wait loop
						if ($forceRedraw) {
							break
						}
						
						# Check if window has been stable (matches pending resize) long enough to process
						if ($null -ne $PendingResize -and $null -ne $lastResizeDetection) {
							# Verify current size still matches pending resize (window stopped changing)
							$sizeMatchesPending = ($newWindowSize.Width -eq $PendingResize.Width -and 
								$newWindowSize.Height -eq $PendingResize.Height)
							
							if ($sizeMatchesPending) {
								$timeSinceResize = Get-TimeSinceMs -startTime $lastResizeDetection
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
									$ResizeClearedScreen = $false  # Reset for next resize
									$LastResizeLogoTime = $null  # Clear logo timer
									# Break out of wait loop to immediately redraw
									break
								}
							}
						}
					}
					
					start-sleep -m 50
					
					# Check if we've waited long enough
					if ($x -ge $math) {
						break
					}
				} # end :waitLoop
			}
			
			# Keyboard and mouse input checking is now done every 200ms in the wait loop above
			# This provides more reliable detection compared to checking once per interval
			
			# Safety net: detect user input via GetLastInputInfo after wait loop.
			# Same inference as wait-loop: unclassified activity → mouse movement.
			try {
				$lii = New-Object mJiggAPI.LASTINPUTINFO
				$lii.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf([type][mJiggAPI.LASTINPUTINFO])
				if ([mJiggAPI.Mouse]::GetLastInputInfo([ref]$lii)) {
					$tickNow = [uint64][mJiggAPI.Mouse]::GetTickCount64()
					$lastInputTick = [uint64]$lii.dwTime
					$systemIdleMs = $tickNow - $lastInputTick
					$recentSimulated = ($null -ne $LastSimulatedKeyPress) -and ((Get-TimeSinceMs -startTime $LastSimulatedKeyPress) -lt 500)
					$recentAutoMove = ($null -ne $LastAutomatedMouseMovement) -and ((Get-TimeSinceMs -startTime $LastAutomatedMouseMovement) -lt 500)

					if ($systemIdleMs -lt 300 -and -not $recentSimulated -and -not $recentAutoMove) {
						$script:userInputDetected = $true
						if ($script:AutoResumeDelaySeconds -gt 0) {
							$LastUserInputTime = Get-Date
						}
						if (-not $keyboardInputDetected -and -not $scrollDetectedInInterval -and -not $mouseInputDetected) {
							$mouseInputDetected = $true
							$script:LastMouseMovementTime = Get-Date
							$mouseMoveText = "Mouse"
							if ($intervalMouseInputs -notcontains $mouseMoveText) {
								$intervalMouseInputs += $mouseMoveText
							}
						}
					}
				}
			} catch {
				# GetLastInputInfo not available, skip
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
						if (-not $ResizeClearedScreen) {
							$ResizeClearedScreen = $true
							$script:CurrentResizeQuote = $null
							Draw-ResizeLogo -ClearFirst
							$LastResizeLogoTime = Get-Date
						}
						$PendingResize = $newWindowSize
						$lastResizeDetection = Get-Date
					}
				}
				
				# Redraw logo every 100ms while resizing
				if ($ResizeClearedScreen -and $null -ne $LastResizeLogoTime) {
					$timeSinceLogoDraw = Get-TimeSinceMs -startTime $LastResizeLogoTime
					if ($timeSinceLogoDraw -ge 3) {
						Draw-ResizeLogo -ClearFirst
						$LastResizeLogoTime = Get-Date
					}
				}
				
				# Check if window has been stable (matches pending resize) long enough to process
				if ($null -ne $PendingResize -and $null -ne $lastResizeDetection) {
					# Verify current size still matches pending resize (window stopped changing)
					$sizeMatchesPending = ($newWindowSize.Width -eq $PendingResize.Width -and 
						$newWindowSize.Height -eq $PendingResize.Height)
					
					if ($sizeMatchesPending) {
						$timeSinceResize = Get-TimeSinceMs -startTime $lastResizeDetection
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
							$ResizeClearedScreen = $false  # Reset for next resize
							$LastResizeLogoTime = $null  # Clear logo timer
						}
					}
				}
			}
			
			# Check if this is the first run (before we modify lastMovementTime)
			$isFirstRun = ($null -eq $LastMovementTime)
			
			# Wait for mouse to stop moving before proceeding
			# This prevents stutter by ensuring the mouse is settled before we do expensive operations
			# Only do this if we actually waited (not on first run or force redraw)
			if (-not $isFirstRun -and -not $forceRedraw) {
				$mouseSettleMs = 150  # Must be still for this long
				$lastSettleCheckPos = Get-MousePosition
				$mouseSettledTime = $null
				$settleLoopCount = 0
				$maxMoveDelta = 0
				
				if ($script:DiagEnabled) {
					"$(Get-Date -Format 'HH:mm:ss.fff') - Loop $($script:LoopIteration): Starting settle wait, pos: $($lastSettleCheckPos.X),$($lastSettleCheckPos.Y)" | Out-File $script:SettleDiagFile -Append
				}
				
				while ($true) {
					$settleLoopCount++
					Start-Sleep -Milliseconds 25
					$currentSettlePos = Get-MousePosition
					
					$mouseMoved = $false
					if ($null -ne $currentSettlePos -and $null -ne $lastSettleCheckPos) {
						$deltaX = [Math]::Abs($currentSettlePos.X - $lastSettleCheckPos.X)
						$deltaY = [Math]::Abs($currentSettlePos.Y - $lastSettleCheckPos.Y)
						$moveDelta = [Math]::Max($deltaX, $deltaY)
						if ($moveDelta -gt $maxMoveDelta) { $maxMoveDelta = $moveDelta }
						if ($deltaX -gt 2 -or $deltaY -gt 2) {
							$mouseMoved = $true
						}
					}
					$lastSettleCheckPos = $currentSettlePos
					
					if ($mouseMoved) {
						$mouseSettledTime = $null
					} else {
						if ($null -eq $mouseSettledTime) {
							$mouseSettledTime = Get-Date
						} elseif (((Get-Date) - $mouseSettledTime).TotalMilliseconds -ge $mouseSettleMs) {
							if ($script:DiagEnabled) {
								"$(Get-Date -Format 'HH:mm:ss.fff') - Loop $($script:LoopIteration): Settled after $settleLoopCount checks, max delta: $maxMoveDelta" | Out-File $script:SettleDiagFile -Append
							}
							break
						}
					}
				}
			}
			
			# Determine if we should skip the update based on user input or first run
			if ($script:userInputDetected) {
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
			$currentPos = Get-MousePosition
			$PosUpdate = $false
			$x = 0
			$y = 0
			
			# Only check for mouse movement if we haven't already detected user input
			# Skip checking if we recently performed automated mouse movement (within last 300ms)
			# This prevents our own automated movement from being detected as user input
			$shouldCheckMouseAfterWait = $true
			if ($null -ne $LastAutomatedMouseMovement) {
				$timeSinceAutomatedMovement = Get-TimeSinceMs -startTime $LastAutomatedMouseMovement
				if ($timeSinceAutomatedMovement -lt 300) {
					# Too soon after our automated movement - skip mouse detection
					$shouldCheckMouseAfterWait = $false
				}
			}
			
			if ($shouldCheckMouseAfterWait -and -not $script:userInputDetected -and $null -ne $mousePosAtStart -and $null -ne $currentPos) {
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
						$mouseMoveText = "Mouse"
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
					$pos = Get-MousePosition
					if ($null -eq $pos) {
						# API call failed - use last known position
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
				
				# Update last position using cached method for better performance
				$newPos = Get-MousePosition
				if ($null -ne $newPos) {
					$LastPos = $newPos
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
					# Flush any simulated key events from the console input buffer
					# so they aren't mistaken for user keyboard input on the next check
					try {
						$hStdIn = [mJiggAPI.Mouse]::GetStdHandle(-10)
						$flushBuf = New-Object 'mJiggAPI.INPUT_RECORD[]' 32
						$flushCount = [uint32]0
						if ([mJiggAPI.Mouse]::PeekConsoleInput($hStdIn, $flushBuf, 32, [ref]$flushCount) -and $flushCount -gt 0) {
							[mJiggAPI.Mouse]::ReadConsoleInput($hStdIn, $flushBuf, $flushCount, [ref]$flushCount) | Out-Null
						}
					} catch { }
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
			# Mouse movement first, then clicks/scroll, then keyboard
			$allInputs = @()
			$mouseMovement = $intervalMouseInputs | Where-Object { $_ -eq "Mouse" }
			if ($mouseMovement) {
				$allInputs += $mouseMovement
				$otherMouseInputs = $intervalMouseInputs | Where-Object { $_ -ne "Mouse" }
				if ($otherMouseInputs) {
					$allInputs += $otherMouseInputs
				}
			} else {
				if ($intervalMouseInputs.Count -gt 0) {
					$allInputs += $intervalMouseInputs
				}
			}
			if ($keyboardInputDetected) {
				$allInputs += "Keyboard"
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
					priority = [int]1
					text = $date.ToString()
					shortText = $date.ToString("HH:mm:ss")
				}
				
				# Component 2: Main message
				if ($SkipUpdate -ne $true) {
					if ($PosUpdate) {
						# Get direction arrow if available
						$arrowText = if ($directionArrow) { " $directionArrow" } else { "" }
						$logComponents += @{
							priority = [int]2
							text = " - Coordinates updated$arrowText"
							shortText = " - Updated$arrowText"
						}
						# Component 3: Coordinates
						$logComponents += @{
							priority = [int]3
							text = " x$x/y$y"
							shortText = " x$x/y$y"
						}
					} else {
						$logComponents += @{
							priority = [int]2
							text = " - Input detected, skipping update"
							shortText = " - Input detected"
						}
					}
				} elseif ($isFirstRun) {
					# First run - show initialization message
					$logComponents += @{
						priority = [int]2
						text = " - Initialized"
						shortText = " - Initialized"
					}
				} elseif ($keyboardInputDetected -or $mouseInputDetected) {
					# User input was detected - show user input skip with KB/MS status
					$logComponents += @{
						priority = [int]2
						text = " - User input skip"
						shortText = " - Skipped"
					}
				} elseif ($cooldownActive) {
					# Auto-resume delay is active (no user input detected) - show custom message
					$logComponents += @{
						priority = [int]2
						text = " - Auto-Resume Delay"
						shortText = " - Auto-Resume Delay"
					}
					# Add resume timer component
					$logComponents += @{
						priority = [int]4
						text = " [Resume: ${secondsRemaining}s]"
						shortText = " [R: ${secondsRemaining}s]"
					}
				} else {
					$logComponents += @{
						priority = [int]2
						text = " - User input skip"
						shortText = " - Skipped"
					}
				}
				
				# Component 4: Wait interval info (only if not cooldown active or user input detected)
				if ($waitExecuted -and -not $cooldownActive) {
					$logComponents += @{
						priority = [int]4
						text = " [Interval:${interval}s]"
						shortText = " [Interval:${interval}s]"
					}
				} elseif (-not $isFirstRun -and -not $cooldownActive) {
					$logComponents += @{
						priority = [int]4
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
						priority = [int]5
						text = " [KB:$kbStatus]"
						shortText = " [K:" + $kbStatus.Substring(0,1) + "]"
					}
					
					# Mouse detection status
					$msStatus = if ($mouseInputDetected) { "YES" } else { "NO" }
					$logComponents += @{
						priority = [int]6
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
			$skipConsoleUpdate = (Get-TimeSinceMs -startTime $script:LastMouseMovementTime) -lt 200
			# Force UI redraw when forceRedraw is true (e.g., after window resize) - override skipConsoleUpdate
			if ($forceRedraw) {
				$skipConsoleUpdate = $false
			}
			
			if ($Output -ne "hidden" -and -not $skipConsoleUpdate) {
				# Output blank line
				Write-Buffer -X 0 -Y $Outputline -Text (" " * $HostWidth)
				$Outputline++

				# Output header
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
				
				# Write left part (mJig title) via buffer with static emoji positioning
				$mouseEmoji = [char]::ConvertFromUtf32(0x1F400)  # 🐀
				$hourglassEmoji = [char]::ConvertFromUtf32(0x23F3)  # ⏳
				Write-Buffer -X 0 -Y $Outputline -Text "  mJig(" -FG $script:HeaderAppName
				$curX = 7  # "  mJig(" = 7 chars
				Write-Buffer -Text $mouseEmoji -FG $script:HeaderIcon
				Write-Buffer -X ($curX + 2) -Y $Outputline -Text ")" -FG $script:HeaderAppName
				$curX = $curX + 2 + 1  # emoji (2) + ")" (1)
				# Add DEBUGMODE indicator if in debug mode
				if ($DebugMode) {
					Write-Buffer -Text " - DEBUGMODE" -FG $script:TextError
					$curX += 13
				}
				
				# Add spacing before times
				Write-Buffer -Text (" " * $spacingBeforeTimes)
				$curX += $spacingBeforeTimes
				
				# Write times (Current first, then End)
				Write-Buffer -Text "Current" -FG $script:HeaderTimeLabel
				$hourglassX1 = $curX + 7  # "Current" = 7 chars
				Write-Buffer -Text $hourglassEmoji -FG $script:HeaderIcon
				Write-Buffer -X ($hourglassX1 + 2) -Y $Outputline -Text "/" -FG $script:TextDefault
				$curX = $hourglassX1 + 2 + 1  # emoji (2) + "/" (1)
				Write-Buffer -Text "$currentTime" -FG $script:HeaderTimeValue
				$curX += $currentTime.Length
				$arrowTriangle = [char]0x27A3  # ➣
				Write-Buffer -Text " $arrowTriangle  "
				$curX += 4  # " ➣  " = 4 display chars
				Write-Buffer -Text "End" -FG $script:HeaderTimeLabel
				$hourglassX2 = $curX + 3  # "End" = 3 chars
				Write-Buffer -Text $hourglassEmoji -FG $script:HeaderIcon
				Write-Buffer -X ($hourglassX2 + 2) -Y $Outputline -Text "/" -FG $script:TextDefault
				$curX = $hourglassX2 + 2 + 1  # emoji (2) + "/" (1)
				Write-Buffer -Text "$endTimeDisplay" -FG $script:HeaderTimeValue
				$curX += $endTimeDisplay.Length
				
				# Add spacing after times and write view tag aligned to the right
				Write-Buffer -Text (" " * $spacingAfterTimes)
				$curX += $spacingAfterTimes
				if ($Output -eq "full") {
					Write-Buffer -Text " Full" -FG $script:HeaderViewTag
					$curX += 5
				} else {
					Write-Buffer -Text " Minimum" -FG $script:HeaderViewTag
					$curX += 8
				}
				
				# Add 2 spaces for right margin
				Write-Buffer -Text "  "
				$curX += 2
				
				# Clear any remaining characters on the line
				if ($curX -lt $HostWidth) {
					Write-Buffer -Text (" " * ($HostWidth - $curX))
				}
				$Outputline++

				# Output Line Spacer
				Write-Buffer -X 0 -Y $Outputline -Text " "
				Write-Buffer -Text ("$($script:BoxHorizontal)" * ($HostWidth - 2)) -FG $script:HeaderSeparator
				Write-Buffer -Text " "
				$outputLine++

			# Only render console if not skipping updates (prevents stutter during mouse movement)
			if (-not $skipConsoleUpdate) {
				# Calculate view-dependent variables INSIDE the skip check to ensure they use current $Output value
				# This prevents stale view calculations when console updates resume after mouse movement
				$boxWidth = 50  # Width for stats box
				$boxPadding = 2  # Padding around box (1 space on each side)
				$verticalSeparatorWidth = 3  # " $($script:BoxVertical) " = 3 characters
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
					$rowY = $Outputline + $i
					$availableWidth = $logWidth
					
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
							# Sort components by priority (ascending) - lower priority numbers appear first
							$sortedComponents = $LogArray[$i].components | Sort-Object { [int]$_.priority }
							foreach ($component in $sortedComponents) {
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
							Write-Buffer -X 0 -Y $rowY -Text $paddedLine
							
							if ($showStatsBox) {
								Write-Buffer -X $logWidth -Y $rowY -Text " $($script:BoxVertical) " -FG $script:StatsBoxBorder
							}
							
							# Draw stats box in full view (with padding so it doesn't touch white lines)
							if ($showStatsBox) {
								Write-Buffer -Text " "
								
								# Draw box content
								if ($i -eq 0) {
									# Top border
									Write-Buffer -Text "$($script:BoxTopLeft)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxTopRight)" -FG $script:StatsBoxBorder
								} elseif ($i -eq 1) {
									# Header row
									$boxHeader = "Stats"
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									Write-Buffer -Text $boxHeader.PadRight($boxWidth - 2) -FG $script:StatsBoxTitle
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
								} elseif ($i -eq 2) {
									# Separator row
									Write-Buffer -Text "$($script:BoxVerticalRight)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxVerticalLeft)" -FG $script:StatsBoxBorder
								} elseif ($i -eq $Rows - 5) {
									Write-Buffer -Text "$($script:BoxVerticalRight)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxVerticalLeft)" -FG $script:StatsBoxBorder
								} elseif ($i -eq $Rows - 4) {
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									Write-Buffer -Text "Detected Inputs:".PadRight($boxWidth - 2) -FG $script:StatsBoxTitle
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
								} elseif ($i -eq $Rows - 3) {
									if ($PreviousIntervalKeys.Count -gt 0) {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text $keysFirstLine.PadRight($boxWidth - 2) -FG $script:StatsValue
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									} else {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text "(none)".PadRight($boxWidth - 2) -FG $script:TextMuted
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									}
								} elseif ($i -eq $Rows - 2) {
									if ($PreviousIntervalKeys.Count -gt 0 -and $keysSecondLine -ne "") {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text $keysSecondLine.PadRight($boxWidth - 2) -FG $script:StatsValue
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									} else {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text (" " * ($boxWidth - 2))
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									}
								} elseif ($i -eq $Rows - 1) {
									Write-Buffer -Text "$($script:BoxBottomLeft)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxBottomRight)" -FG $script:StatsBoxBorder
								} else {
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									Write-Buffer -Text (" " * ($boxWidth - 2))
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
								}
								
								Write-Buffer -Text " "
							}
						} else {
							$emptyLine = "".PadRight($availableWidth)
							Write-Buffer -X 0 -Y $rowY -Text $emptyLine
							
							if ($showStatsBox) {
								Write-Buffer -X $logWidth -Y $rowY -Text " $($script:BoxVertical) " -FG $script:StatsBoxBorder
							}
							
							if ($showStatsBox) {
								Write-Buffer -Text " "
								if ($i -eq 0) {
									Write-Buffer -Text "$($script:BoxTopLeft)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxTopRight)" -FG $script:StatsBoxBorder
								} elseif ($i -eq 1) {
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									Write-Buffer -Text "Stats".PadRight($boxWidth - 2) -FG $script:StatsBoxTitle
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
								} elseif ($i -eq 2) {
									Write-Buffer -Text "$($script:BoxVerticalRight)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxVerticalLeft)" -FG $script:StatsBoxBorder
								} elseif ($i -eq $Rows - 5) {
									Write-Buffer -Text "$($script:BoxVerticalRight)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxVerticalLeft)" -FG $script:StatsBoxBorder
								} elseif ($i -eq $Rows - 4) {
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									Write-Buffer -Text "Detected Inputs:".PadRight($boxWidth - 2) -FG $script:StatsBoxTitle
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
								} elseif ($i -eq $Rows - 3) {
									if ($PreviousIntervalKeys.Count -gt 0) {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text $keysFirstLine.PadRight($boxWidth - 2) -FG $script:StatsValue
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									} else {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text "(none)".PadRight($boxWidth - 2) -FG $script:TextMuted
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									}
								} elseif ($i -eq $Rows - 2) {
									if ($PreviousIntervalKeys.Count -gt 0 -and $keysSecondLine -ne "") {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text $keysSecondLine.PadRight($boxWidth - 2) -FG $script:StatsValue
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									} else {
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
										Write-Buffer -Text (" " * ($boxWidth - 2))
										Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									}
								} elseif ($i -eq $Rows - 1) {
									Write-Buffer -Text "$($script:BoxBottomLeft)" -FG $script:StatsBoxBorder
									Write-Buffer -Text ("$($script:BoxHorizontal)" * ($boxWidth - 2)) -FG $script:StatsBoxBorder
									Write-Buffer -Text "$($script:BoxBottomRight)" -FG $script:StatsBoxBorder
								} else {
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
									Write-Buffer -Text (" " * ($boxWidth - 2))
									Write-Buffer -Text "$($script:BoxVertical)" -FG $script:StatsBoxBorder
								}
								Write-Buffer -Text " "
							}
						}
				}
				$outputLine += $Rows
			}
			}  # End of skipConsoleUpdate check

			# Output bottom separator (only if not skipping console updates)
			if ($Output -ne "hidden" -and -not $skipConsoleUpdate) {
				# Calculate if we should show stats box (full view, wide enough window)
				$boxWidth = 50  # Width for stats box
				$boxPadding = 2  # Padding around box (1 space on each side)
				$verticalSeparatorWidth = 3  # " $($script:BoxVertical) " = 3 characters
				$showStatsBox = ($Output -eq "full" -and $HostWidth -ge ($boxWidth + $boxPadding + $verticalSeparatorWidth + 50))  # Need at least 50 chars for logs
				$logWidth = if ($showStatsBox) { $HostWidth - $boxWidth - $boxPadding - $verticalSeparatorWidth } else { $HostWidth }  # Reserve space for box + padding + separator
				
				# Bottom separator line
				Write-Buffer -X 0 -Y $Outputline -Text " "
				Write-Buffer -Text ("$($script:BoxHorizontal)" * ($HostWidth - 2)) -FG $script:HeaderSeparator
				Write-Buffer -Text " "
				$outputLine++

				## Menu Options ##
				$timeMenuText = if ($endTimeInt -eq -1 -or [string]::IsNullOrEmpty($endTimeStr)) {
					"set_end_(t)ime"
				} else {
					"change_end_(t)ime"
				}
				
				$emojiHourglass = [char]::ConvertFromUtf32(0x23F3)  # ⏳
				$emojiRefresh = [char]::ConvertFromUtf32(0x1F441)  # 👁️
				$emojiLock = [char]::ConvertFromUtf32(0x1F512)  # 🔒
				$emojiGear = [char]::ConvertFromUtf32(0x1F6E0)  # 🛠
				$emojiRedX = [char]::ConvertFromUtf32(0x274C)  # ❌
				
				$menuItemsList = @(
					@{
						full = "$emojiHourglass|$timeMenuText"
						noIcons = $timeMenuText
						short = "(t)ime"
					},
					@{
						full = "$emojiRefresh|toggle_(v)iew"
						noIcons = "toggle_(v)iew"
						short = "(v)iew"
					},
					@{
						full = "$emojiLock|(h)ide_output"
						noIcons = "(h)ide_output"
						short = "(h)ide"
					}
				)
				
				if ($Output -eq "full") {
					$menuItemsList += @{
						full = "$emojiGear|modify_(m)ovement"
						noIcons = "modify_(m)ovement"
						short = "(m)ove"
					}
				}
				
				$menuItemsList += @{
					full = "$emojiRedX|(q)uit"
					noIcons = "(q)uit"
					short = "(q)uit"
				}
				
				$menuItems = $menuItemsList
				
				# Calculate widths for each format (emojis = 2 display chars)
				$format0Width = 2  # Leading spaces
				$format1Width = 2
				$format2Width = 2
				
				foreach ($item in $menuItems) {
					$textPart = $item.full -replace "^.+\|", ""
					$format0Width += 2 + 1 + $textPart.Length + 2
					$format1Width += $item.noIcons.Length + 2
					$format2Width += $item.short.Length + 1
				}
				
				$format0Width += 2
				$format1Width += 2
				$format2Width += 2
				
				$menuFormat = 0
				if ($HostWidth -lt $format0Width) {
					if ($HostWidth -lt $format1Width) {
						$menuFormat = 2
					} else {
						$menuFormat = 1
					}
				}
				
				$quitItem = $menuItems[$menuItems.Count - 1]
				if ($menuFormat -eq 0) {
					$textPart = $quitItem.full -replace "^.+\|", ""
					$quitWidth = 2 + 1 + $textPart.Length
				} elseif ($menuFormat -eq 1) {
					$quitWidth = $quitItem.noIcons.Length
				} else {
					$quitWidth = $quitItem.short.Length
				}
				
				# Write menu items via buffer with static position tracking
				$menuY = $Outputline
				$currentMenuX = 2  # Start after "  " prefix
				Write-Buffer -X 0 -Y $menuY -Text "  "
				
				$script:MenuItemsBounds = @()
				$itemsBeforeQuit = $menuItems.Count - 1
				for ($mi = 0; $mi -lt $itemsBeforeQuit; $mi++) {
					$item = $menuItems[$mi]
					$itemStartX = $currentMenuX
					
					if ($menuFormat -eq 0) {
						$itemText = $item.full
					} elseif ($menuFormat -eq 1) {
						$itemText = $item.noIcons
					} else {
						$itemText = $item.short
					}
					
					# Calculate item display width statically (emoji = 2 display cells)
					$itemDisplayWidth = 0
					if ($menuFormat -eq 0) {
						$parts = $itemText -split "\|", 2
						if ($parts.Count -eq 2) {
							$itemDisplayWidth = 2 + 1 + $parts[1].Length
						}
					} else {
						$itemDisplayWidth = $itemText.Length
					}
					
					# Render the menu item
					if ($menuFormat -eq 0) {
						$parts = $itemText -split "\|", 2
						if ($parts.Count -eq 2) {
							$emoji = $parts[0]
							$text = $parts[1]
							Write-Buffer -X $itemStartX -Y $menuY -Text $emoji -BG $script:MenuButtonBg -Wide
						$pipeX = $itemStartX + 2
						Write-Buffer -X $pipeX -Y $menuY -Text "|" -FG $script:MenuButtonPipe -BG $script:MenuButtonBg
						$textParts = $text -split "([()])"
							for ($j = 0; $j -lt $textParts.Count; $j++) {
								$part = $textParts[$j]
								if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
									Write-Buffer -Text "(" -FG $script:MenuButtonText -BG $script:MenuButtonBg
									Write-Buffer -Text $textParts[$j + 1] -FG $script:MenuButtonHotkey -BG $script:MenuButtonBg
									Write-Buffer -Text ")" -FG $script:MenuButtonText -BG $script:MenuButtonBg
									$j += 2
								} elseif ($part -ne "") {
									Write-Buffer -Text $part -FG $script:MenuButtonText -BG $script:MenuButtonBg
								}
							}
						}
					} else {
						Write-Buffer -X $itemStartX -Y $menuY -Text "" -BG $script:MenuButtonBg
						$textParts = $itemText -split "([()])"
						for ($j = 0; $j -lt $textParts.Count; $j++) {
							$part = $textParts[$j]
							if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
								Write-Buffer -Text "(" -FG $script:MenuButtonText -BG $script:MenuButtonBg
								Write-Buffer -Text $textParts[$j + 1] -FG $script:MenuButtonHotkey -BG $script:MenuButtonBg
								Write-Buffer -Text ")" -FG $script:MenuButtonText -BG $script:MenuButtonBg
								$j += 2
							} elseif ($part -ne "") {
								Write-Buffer -Text $part -FG $script:MenuButtonText -BG $script:MenuButtonBg
							}
						}
					}
					
					# Store menu item bounds (computed statically)
					$itemEndX = $itemStartX + $itemDisplayWidth - 1
					$hotkeyMatch = $itemText -match "\(([a-z])\)"
					$hotkey = if ($hotkeyMatch) { $matches[1] } else { $null }
					$script:MenuItemsBounds += @{
						startX = $itemStartX
						endX = $itemEndX
						y = $menuY
						hotkey = $hotkey
						index = $mi
					}
					
					# Advance position statically
					$currentMenuX = $itemStartX + $itemDisplayWidth
					
					if ($menuFormat -eq 2) {
						Write-Buffer -Text " "
						$currentMenuX += 1
					} else {
						Write-Buffer -Text "  "
						$currentMenuX += 2
					}
				}
				
				# Add spacing before quit (right-align in full/noIcons formats)
				if ($menuFormat -lt 2) {
					$desiredQuitX = $HostWidth - $quitWidth - 2
					$spacing = [math]::Max(1, $desiredQuitX - $currentMenuX)
					Write-Buffer -Text (" " * $spacing)
					$currentMenuX += $spacing
				}
				
				# Write quit item
				$quitStartX = $currentMenuX
				if ($menuFormat -eq 0) {
					$itemText = $quitItem.full
				} elseif ($menuFormat -eq 1) {
					$itemText = $quitItem.noIcons
				} else {
					$itemText = $quitItem.short
				}
				
				if ($menuFormat -eq 0) {
					$parts = $itemText -split "\|", 2
					if ($parts.Count -eq 2) {
						$emoji = $parts[0]
						$text = $parts[1]
						Write-Buffer -X $quitStartX -Y $menuY -Text $emoji -BG $script:MenuButtonBg -Wide
					$pipeX = $quitStartX + 2
						Write-Buffer -X $pipeX -Y $menuY -Text "|" -FG $script:MenuButtonPipe -BG $script:MenuButtonBg
						$textParts = $text -split "([()])"
						for ($j = 0; $j -lt $textParts.Count; $j++) {
							$part = $textParts[$j]
							if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
								Write-Buffer -Text "(" -FG $script:MenuButtonText -BG $script:MenuButtonBg
								Write-Buffer -Text $textParts[$j + 1] -FG $script:MenuButtonHotkey -BG $script:MenuButtonBg
								Write-Buffer -Text ")" -FG $script:MenuButtonText -BG $script:MenuButtonBg
								$j += 2
							} elseif ($part -ne "") {
								Write-Buffer -Text $part -FG $script:MenuButtonText -BG $script:MenuButtonBg
							}
						}
					}
				} else {
					Write-Buffer -X $quitStartX -Y $menuY -Text "" -BG $script:MenuButtonBg
					$textParts = $itemText -split "([()])"
					for ($j = 0; $j -lt $textParts.Count; $j++) {
						$part = $textParts[$j]
						if ($part -eq "(" -and $j + 2 -lt $textParts.Count -and $textParts[$j + 1] -match "^[a-z]$" -and $textParts[$j + 2] -eq ")") {
							Write-Buffer -Text "(" -FG $script:MenuButtonText -BG $script:MenuButtonBg
							Write-Buffer -Text $textParts[$j + 1] -FG $script:MenuButtonHotkey -BG $script:MenuButtonBg
							Write-Buffer -Text ")" -FG $script:MenuButtonText -BG $script:MenuButtonBg
							$j += 2
						} elseif ($part -ne "") {
							Write-Buffer -Text $part -FG $script:MenuButtonText -BG $script:MenuButtonBg
						}
					}
				}
				
				# Store quit item bounds (computed statically)
				$quitEndX = $quitStartX + $quitWidth - 1
				$quitHotkeyMatch = $itemText -match "\(([a-z])\)"
				$quitHotkey = if ($quitHotkeyMatch) { $matches[1] } else { $null }
				$script:MenuItemsBounds += @{
					startX = $quitStartX
					endX = $quitEndX
					y = $menuY
					hotkey = $quitHotkey
					index = $menuItems.Count - 1
				}
				
				# Right margin and clear remaining
				$menuEndX = $quitStartX + $quitWidth
				Write-Buffer -Text "  "
				$menuEndX += 2
				if ($menuEndX -lt $HostWidth) {
					Write-Buffer -Text (" " * ($HostWidth - $menuEndX))
				}
				$Outputline++
				
				# Flush entire UI to console in one operation
				Flush-Buffer
			} elseif ($Output -eq "hidden") {
				if (-not $skipConsoleUpdate) {
					# Detect resize while in hidden mode
					$pswindow = (Get-Host).UI.RawUI
					$newW = $pswindow.WindowSize.Width
					$newH = $pswindow.WindowSize.Height
					if ($newW -ne $HostWidth -or $newH -ne $HostHeight) {
						$HostWidth = $newW
						$HostHeight = $newH
						$forceRedraw = $true
					}
					
					$timeStr = $date.ToString("HH:mm:ss")
					$statusLine = "$timeStr | running..."
					
					Write-Buffer -X 0 -Y 0 -Text $statusLine.PadRight($HostWidth)
					
					$hBtnY = [math]::Max(1, $HostHeight - 2)
					$hBtnX = [math]::Max(0, $HostWidth - 4)
					Write-Buffer -X $hBtnX -Y $hBtnY -Text "(" -FG $script:MenuButtonText -BG $script:MenuButtonBg
					Write-Buffer -Text "h" -FG $script:MenuButtonHotkey -BG $script:MenuButtonBg
					Write-Buffer -Text ")" -FG $script:MenuButtonText -BG $script:MenuButtonBg
					
					if ($forceRedraw) { Flush-Buffer -ClearFirst } else { Flush-Buffer }
					
					$script:MenuItemsBounds = @(@{
						startX = $hBtnX
						endX = $hBtnX + 2
						y = $hBtnY
						hotkey = "h"
						index = 0
					})
				}
			}
			# If skipConsoleUpdate is true and Output is not "hidden", don't render anything (prevents stutter)
			
			# Reset resize cleared screen flag after we've completed a redraw
			# This ensures the screen will be cleared again if user starts a new resize
			if ($forceRedraw -and -not $skipConsoleUpdate) {
				$ResizeClearedScreen = $false
			}
			
			# Check if end time reached (only if end time is set)
			# Compare full MMddHHmm values to handle overnight runs correctly
			if ($endTimeInt -ne -1 -and -not [string]::IsNullOrEmpty($end)) {
				try {
					$currentDateTimeInt = [int]($date.ToString("MMddHHmm"))
					$endDateTimeInt = [int]$end
					if ($currentDateTimeInt -ge $endDateTimeInt) {
						$time = $true
					}
				} catch {
					# If comparison fails, don't stop the script
				}
			}
			
			# Only break if time is explicitly set to true
			if ($time -eq $true) {
				# End message
				if ($Output -ne "hidden") {
					[Console]::SetCursorPosition(0, $Outputline)
					Write-Host "       END TIME REACHED: " -NoNewline -ForegroundColor $script:TextError
					Write-Host "Stopping " -NoNewline
					Write-Host "mJig"
					write-host
				}
				break
			}
		} # end :process
}

# Uncomment the line below to run the function when script is executed directly
# Start-mJig -Output full
