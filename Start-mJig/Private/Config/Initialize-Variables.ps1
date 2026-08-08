	# Copy parameters to script-scoped variables so they can be modified at runtime
	$script:IntervalSeconds = $IntervalSeconds
	$script:IntervalVariance = $IntervalVariance
	$script:MoveSpeed = $MoveSpeed
	$script:MoveVariance = $MoveVariance
	$script:TravelDistance = $TravelDistance
	$script:TravelVariance = $TravelVariance
	$script:AutoResumeDelaySeconds = $AutoResumeDelaySeconds
	$script:EndVariance = $EndVariance
	$script:Output = $Output
	$script:DebugMode = [bool]$DebugMode

	# Derive stealth session identifier for pipe/mutex/encryption names
	# (Get-SessionIdentifier is dot-sourced in Start-mJig.psm1 before this file is loaded)
	$script:SessionId = Get-SessionIdentifier
	$script:PipeEncryptionKey = $script:SessionId.AesKey
	$script:PipeAuthToken = $script:SessionId.AuthToken
	if ($_PipeName -eq 'mJig_IPC') {
		$script:PipeName = $script:SessionId.PipeName
	} else {
		$script:PipeName = $_PipeName
	}
	if ($_WorkerMode -and $script:_wsDiagFile) {
		"$(Get-Date -Format 'HH:mm:ss.fff') [2] Session derived  PipeName=$($script:PipeName)  KeyLen=$($script:PipeEncryptionKey.Length)  TokenLen=$($script:PipeAuthToken.Length)" | Out-File $script:_wsDiagFile -Append
	}

	# General script-scoped state
	$script:PendingForceRedraw = $false
	$script:CurrentScreenState = "startup"

	# Stats tracking — cumulative counters updated every move/skip cycle
	$script:StatsMoveCount             = 0
	$script:StatsSkipCount             = 0
	$script:StatsCurrentStreak         = 0
	$script:StatsLongestStreak         = 0
	$script:StatsTotalDistancePx       = 0.0
	$script:StatsLastMoveDist          = 0.0
	$script:StatsMinMoveDist           = [double]::MaxValue
	$script:StatsMaxMoveDist           = 0.0
	$script:StatsKbInterruptCount      = 0
	$script:StatsMsInterruptCount      = 0
	$script:StatsLongestCleanStreak    = 0
	$script:StatsCleanStreak           = 0
	$script:StatsAvgActualIntervalSecs = 0.0
	$script:StatsLastMoveTick          = $null
	$script:StatsAvgDurationMs         = 0.0
	$script:StatsMinDurationMs         = [int]::MaxValue
	$script:StatsMaxDurationMs         = 0
	$script:StatsDirectionCounts       = @{ N = 0.0; NE = 0.0; E = 0.0; SE = 0.0; S = 0.0; SW = 0.0; W = 0.0; NW = 0.0 }
	$script:StatsLastCurveParams       = $null

	# Curve animation — pending flag cleared by animation loop when it consumes new params
	$script:StatsCurveAnimPending  = $false
	$script:_LastCurveParamKey     = ""

	# Input and hotkey state
	$script:PreviousIntervalKeys = @()
	$script:ResizeThrottleMs = 100
	$script:LoopIteration = 0
	$script:LastInputCheckTime = $null
	$script:DialogButtonClick = $null
	$script:ManualPause = $false
	$script:_HotkeyDebounce = $false

	# Display sleep state
	$script:DisplaySleepMode            = $false
	$script:DisplaySleepAudioEnabled    = $true
	$script:DisplaySleepAutoEnabled     = $false
	$script:DisplaySleepAutoTimeoutSecs = 60
	$script:DisplaySleepActivatedAt     = $null
	# Single unified activity clock: drives auto-sleep idle and auto-resume cooldown.
	# Initialized to module load time so auto-sleep does not fire immediately at startup.
	$script:LastUserActivityTime        = Get-Date
	# Armed flag: $false until first real user input; prevents spurious cooldown on startup.
	$script:_CooldownArmed              = $false

	# Performance and UI state
	$script:MethodCache = @{}
	$script:ScreenWidth = $null
	$script:ScreenHeight = $null
	$script:DialogButtonBounds = $null
	$script:LastClickLogTime = $null

	# Title presets
	$script:TitlePresets = @(
		@{ Name = "mJig";                  Emoji = 0x1F400 }
		@{ Name = "Windows Update";        Emoji = 0x1F504 }
		@{ Name = "System Health Check";   Emoji = 0x1FA7A }
		@{ Name = "Background Services";   Emoji = 0x2699  }
		@{ Name = "Windows Defender Scan"; Emoji = 0x1F6E1 }
		@{ Name = "Performance Monitor";   Emoji = 0x1F4CA }
	)
	$script:TitlePresetIndex = 0
	$script:NotificationsEnabled = $true
	if ($Title.Length -gt 0) {
		$script:WindowTitle = $Title
		$script:TitleEmoji  = 0x1F400
		if ($script:TitlePresets[0].Name -ne $Title) {
			$script:TitlePresets = @(@{ Name = $Title; Emoji = 0x1F400 }) + $script:TitlePresets
		}
	} else {
		$script:WindowTitle = "mJig"
		$script:TitleEmoji  = 0x1F400
	}

	# Layout cache for Write-MainFrame geometry (invalidated by stamp change)
	$script:_FrameLayout = $null

	# Click and button tracking
	$script:MenuClickHotkey        = $null
	$script:ModeButtonBounds       = $null
	$script:ModeLabelBounds        = $null
	$script:HeaderEndTimeBounds    = $null
	$script:HeaderCurrentTimeBounds = $null
	$script:HeaderLogoBounds       = $null
	$script:Version                = "1.0.0"
	$script:VersionCheckCache      = $null
	$script:PressedMenuButton      = $null
	$script:ButtonClickedAt        = $null
	$script:PendingDialogCheck     = $false
	$script:LButtonWasDown         = $false
	$script:RenderQueue    = New-Object 'System.Collections.Generic.List[object[]]'
	$script:FrameBuilder   = New-Object System.Text.StringBuilder (8192)
	$script:MenuItemsBounds = New-Object 'System.Collections.Generic.List[hashtable]'

	# Box-drawing characters (using Unicode code points to avoid encoding issues)
	$script:BoxTopLeft       = [char]0x250C   # ┌
	$script:BoxTopRight      = [char]0x2510   # ┐
	$script:BoxBottomLeft    = [char]0x2514   # └
	$script:BoxBottomRight   = [char]0x2518   # ┘
	$script:BoxHorizontal    = [char]0x2500   # ─
	$script:BoxVertical      = [char]0x2502   # │
	$script:BoxVerticalRight = [char]0x251C   # ├
	$script:BoxVerticalLeft  = [char]0x2524   # ┤

	# VT100 / ANSI escape sequence helpers
	$script:ESC = [char]27
	$script:CSI = "$([char]27)["
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

	# Pre-allocated point for cursor movement animation (avoids per-step allocation)
	$script:_MovementPoint = New-Object System.Drawing.Point

	# Cached emoji constants (avoid recomputation every frame)
	$script:HourglassEmoji = [char]::ConvertFromUtf32(0x23F3)   # U+23F3 hourglass
	$script:LockEmoji      = [char]::ConvertFromUtf32(0x1F512)  # U+1F512 lock
	$script:GearEmoji      = [char]::ConvertFromUtf32(0x1F6E0)  # U+1F6E0 gear
	$script:RedXEmoji      = [char]::ConvertFromUtf32(0x274C)   # U+274C red X
	$script:CheckmarkEmoji = [char]::ConvertFromUtf32(0x2705)   # U+2705 checkmark
	$script:PauseEmoji     = "$([char]0x275A)$([char]0x275A)"   # U+275A x2 pause bars
	$script:PlayEmoji      = [char]::ConvertFromUtf32(0x25B6)   # U+25B6 play triangle

	# Cached virtual screen bounds (display config rarely changes during a run)
	$script:_VirtualScreen = [System.Windows.Forms.SystemInformation]::VirtualScreen

	# Playful quotes displayed on the resize screen
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
