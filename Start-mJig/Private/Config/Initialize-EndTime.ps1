	# Convert EndTime to string and parse
	# Handle: "0" = none, "00" or "0000" = midnight (0000), 2-digit = hour on the hour, 4-digit = HHmm
	try {
		$endTimeTrimmed = $EndTime.Trim()
		
		# Check if it is "0" (single digit) - means no end time
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
			$tomorrow = (Get-Date).AddDays(1)
			$endDate = Get-Date $tomorrow -Format "MMdd"
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
		$end = ""
		if ($DebugMode) {
			Write-Host "  [OK] No end time - script will run indefinitely" -ForegroundColor $script:TextSuccess
		}
	}
