	# Initialize lastPos for mouse detection
	if ($DebugMode) {
		Write-Host "[DEBUG] Initializing mouse position tracking..." -ForegroundColor $script:TextHighlight
	}
	try {
		if ($null -eq $LastPos) {
			$_initPoint      = New-Object $script:PointType
			$_hasCursorPos   = $null -ne $script:MouseAPI.GetMethod("GetCursorPos")
			if ($_hasCursorPos) {
				if ($script:MouseAPI::GetCursorPos([ref]$_initPoint)) {
					$LastPos = New-Object System.Drawing.Point($_initPoint.X, $_initPoint.Y)
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
		# Mouse position tracking is optional — continue without it
	}
	# Suppress PSScriptAnalyzer "assigned but never used" — $LastPos is consumed by caller scope
	$null = $LastPos
