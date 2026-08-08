	function Test-UserInputActivity {
		param(
			[object]$LastSimulatedKeyPressTime      = $null,
			[object]$LastAutomatedMouseMovementTime = $null
		)

		# Calls GetLastInputInfo and applies the simulated/automated filters.
		# Returns $true if genuine user activity was detected (system idle < 300 ms after filtering).
		# Caller is responsible for updating flags and the activity clock via Register-UserInput.
		try {
			$_lii        = New-Object $script:LastInputType
			$_lii.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf($_lii)
			if ($script:MouseAPI::GetLastInputInfo([ref]$_lii)) {
				$_tickNow        = [uint64]$script:MouseAPI::GetTickCount64()
				$_systemIdleMs   = $_tickNow - [uint64]$_lii.dwTime
				$_recentSim      = ($null -ne $LastSimulatedKeyPressTime) -and
				                   ((Get-TimeSinceMs -StartTime $LastSimulatedKeyPressTime) -lt 500)
				$_recentAutoMove = ($null -ne $LastAutomatedMouseMovementTime) -and
				                   ((Get-TimeSinceMs -StartTime $LastAutomatedMouseMovementTime) -lt 500)
				return ($_systemIdleMs -lt 300 -and -not $_recentSim -and -not $_recentAutoMove)
			}
		} catch {}
		return $false
	}
