	# Verify types loaded correctly
	if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 2: Types loaded, verifying" | Out-File $script:StartupDiagFile -Append }
	try {
		$null = $script:MouseAPI::GetAsyncKeyState(0x01)
		$_testPoint = New-Object $script:PointType
		$_hasGetCursorPos = $null -ne $script:MouseAPI.GetMethod("GetCursorPos")
		if ($_hasGetCursorPos) {
			$null = $script:MouseAPI::GetCursorPos([ref]$_testPoint)
		}
		if ($DebugMode) {
			Write-Host "  [OK] Windows API types loaded successfully" -ForegroundColor $script:TextSuccess
		}
		if ($_WorkerMode -and $script:_wsDiagFile) {
			"$(Get-Date -Format 'HH:mm:ss.fff') [4c] Type verification OK" | Out-File $script:_wsDiagFile -Append
		}
	} catch {
		if ($_WorkerMode -and $script:_wsDiagFile) {
			"$(Get-Date -Format 'HH:mm:ss.fff') [4c] Type verification FAILED: $($_.Exception.Message)" | Out-File $script:_wsDiagFile -Append
		}
		if ($DebugMode) {
			Write-Host "  [FAIL] Could not verify keyboard/mouse API: $($_.Exception.Message)" -ForegroundColor $script:TextError
		}
		Write-Host "Warning: Could not verify keyboard/mouse API. Some features may be disabled." -ForegroundColor $script:TextWarning
	}

	# Auto-detect headless mode when console window is hidden (e.g. scheduled task)
	if (-not $Headless -and -not $_WorkerMode) {
		try {
			$_consoleHwnd = $script:MouseAPI::GetConsoleWindow()
			if ($_consoleHwnd -eq [IntPtr]::Zero -or -not $script:MouseAPI::IsWindowVisible($_consoleHwnd)) {
				$Headless = $true
			}
		} catch {}
	}
