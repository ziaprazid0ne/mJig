	# Prep the Host Console
	try {
		$Host.UI.RawUI.WindowTitle = if ($DebugMode) { "$script:WindowTitle - Debug Mode" } else { $script:WindowTitle }
		if ($DebugMode) {
			Write-Host "[DEBUG] Set window title: $($Host.UI.RawUI.WindowTitle)" -ForegroundColor $script:TextHighlight
		}
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Failed to set window title: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
		}
	}

	# Duplicate instance detection via named mutex
	if ($DebugMode) {
		Write-Host "[DEBUG] Checking for duplicate mJig instances (mutex)..." -ForegroundColor $script:TextHighlight
	}
	$script:InstanceMutex = $null
	$mutexAcquired = $false
	try {
		$script:InstanceMutex = New-Object System.Threading.Mutex($false, "Global\$($script:SessionId.PipeName)")
		$mutexAcquired = $script:InstanceMutex.WaitOne(0)
	} catch [System.Threading.AbandonedMutexException] {
		$mutexAcquired = $true
	} catch {
		if ($DebugMode) {
			Write-Host "  [WARN] Mutex check failed: $($_.Exception.Message)" -ForegroundColor $script:TextWarning
		}
	}
	$_viewerReconnect = $false
	if (-not $mutexAcquired -and -not $_WorkerMode) {
		# Mutex already held — connect as viewer after initialization
		if ($null -ne $script:InstanceMutex) { $script:InstanceMutex.Dispose() }
		$_viewerReconnect = $true
	}
	if ($DebugMode -and -not $_viewerReconnect) {
		Write-Host "  [OK] Mutex acquired — no other instance running" -ForegroundColor $script:TextSuccess
	}
	if ($_WorkerMode -and $script:_wsDiagFile) {
		"$(Get-Date -Format 'HH:mm:ss.fff') [3] Mutex check  acquired=$mutexAcquired  viewerReconnect=$_viewerReconnect" | Out-File $script:_wsDiagFile -Append
	}
	
	# Worker mode skips console rendering setup (headless)
	if (-not $_WorkerMode) {
	if (-not $_viewerReconnect -and $Output -ne "hidden") {
		if (-not $DebugMode) {
			Show-StartupScreen
		} else {
			try { Clear-Host } catch {}
		}
	}
	
	if ($DebugMode) {
		Write-Host "Initialization Debug" -ForegroundColor $script:HeaderAppName
		Write-Host ""
		$_childRsId = if ([System.Management.Automation.Runspaces.Runspace]::DefaultRunspace) {
			[System.Management.Automation.Runspaces.Runspace]::DefaultRunspace.Id } else { '(unknown)' }
		Write-Host "[RUNSPACE] Confirmed: running inside isolated runspace" -ForegroundColor Cyan
		Write-Host "  Runspace ID : $_childRsId"  -ForegroundColor Gray
		Write-Host "  Thread ID   : $([System.Threading.Thread]::CurrentThread.ManagedThreadId)" -ForegroundColor Gray
		Write-Host "  Modules     : $((Get-Module | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor Gray
		Write-Host ""
		Write-Host "[DEBUG] Initializing console..." -ForegroundColor $script:TextHighlight
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
	# Match buffer height to window height (horizontal buffer managed by PowerShell)
	try {
		$pswindow.BufferSize = New-Object System.Management.Automation.Host.Size($newBufferSize.Width, $newWindowSize.Height)
		$newBufferSize = $pswindow.BufferSize
		if ($DebugMode) {
			Write-Host "  [OK] Set buffer height to match window height" -ForegroundColor $script:TextSuccess
		}
	} catch {
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
	# Suppress PSScriptAnalyzer "assigned but never used" — these locals are consumed by the
	# dot-source caller (Start-mJig.psm1). The reads below prevent false-positive linter warnings.
	$null = $OldBufferSize; $null = $OldWindowSize
	if ($DebugMode) {
		Write-Host "    Final host dimensions: ${HostWidth}x${HostHeight}" -ForegroundColor Gray
	}

	# Initialize the Output Array
	if ($DebugMode) {
		Write-Host "[DEBUG] Initializing output array..." -ForegroundColor $script:TextHighlight
	}
	if ($DebugMode) {
		Write-Host "  [OK] Output mode: $Output" -ForegroundColor $script:TextSuccess
	}

	} # end if (-not $_WorkerMode) — console setup guard
