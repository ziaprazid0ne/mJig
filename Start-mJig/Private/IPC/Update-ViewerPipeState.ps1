# Shared IPC message dispatcher for viewer mode.
# Stored as a scriptblock and dot-invoked (. $_handleIpcMsg) at both the
# top-of-loop and wait-loop read sites so it runs in the caller's scope,
# allowing it to read $msg and mutate caller locals directly.
$_handleIpcMsg = {
	switch ($msg.type) {
		'state' {
			$messageEpoch = if ($null -ne $msg.epoch) { [int]$msg.epoch } else { 0 }
			if ($messageEpoch -lt $_settingsEpoch) { break }
			$script:IntervalSeconds = [double]$msg.intervalSeconds
			$script:IntervalVariance = [double]$msg.intervalVariance
			$script:MoveSpeed = [double]$msg.moveSpeed
			$script:MoveVariance = [double]$msg.moveVariance
			$script:TravelDistance = [double]$msg.travelDistance
			$script:TravelVariance = [double]$msg.travelVariance
			$script:AutoResumeDelaySeconds = [double]$msg.autoResumeDelaySeconds
			$script:LoopIteration = [int]$msg.loopIteration
			$_incomingIter = [int]$msg.loopIteration
			if ($_incomingIter -ne $_viewerWorkerIter) { $_viewerWorkerIterChanged = $true; $_viewerWorkerIter = $_incomingIter }
			$endTimeStr = [string]$msg.endTimeStr
			$endTimeInt = [int]$msg.endTimeInt
			$end = [string]$msg.end
			$cooldownActive = [bool]$msg.cooldownActive
			$secondsRemaining = if ($null -ne $msg.cooldownRemaining) { [int]$msg.cooldownRemaining } else { 0 }
			if ([bool]$msg.mouseInputDetected) {
				$mouseInputDetected = $true
				$null = $intervalMouseInputs.Add("Mouse")
			}
			if ([bool]$msg.keyboardInputDetected) { $keyboardInputDetected = $true }
			if ([bool]$msg.keyboardInferred) { $_keyboardInferred = $true }
			$SkipUpdate = [bool]$msg.userInputDetected -or $cooldownActive
			# Real user activity from worker also re-arms the recurring auto-sleep idle clock
			if ([bool]$msg.userInputDetected) {
				$script:_DisplaySleepLastInputTime = Get-Date
			}
			# Keep viewer display-sleep UI in sync with worker-owned toggles (global hotkey)
			if ($null -ne $msg.displaySleepMode) {
				$_incomingSleep = [bool]$msg.displaySleepMode
				if ($_incomingSleep -ne $script:DisplaySleepMode) {
					$script:DisplaySleepMode = $_incomingSleep
					if ($_incomingSleep) {
						$script:DisplaySleepActivatedAt = Get-Date
						$script:DisplaySleepPrePos      = Get-MousePosition
					} else {
						$script:DisplaySleepActivatedAt = $null
						$script:_DisplaySleepLastInputTime = Get-Date
					}
				}
			}
			if ($null -ne $msg.statsWorkerStartTime) { try { $ScriptStartTime = [DateTime]::Parse([string]$msg.statsWorkerStartTime) } catch {} }
			if ($null -ne $msg.statsMoveCount)             { $script:StatsMoveCount             = [int]$msg.statsMoveCount }
			if ($null -ne $msg.statsSkipCount)             { $script:StatsSkipCount             = [int]$msg.statsSkipCount }
			if ($null -ne $msg.statsCurrentStreak)         { $script:StatsCurrentStreak         = [int]$msg.statsCurrentStreak }
			if ($null -ne $msg.statsLongestStreak)         { $script:StatsLongestStreak         = [int]$msg.statsLongestStreak }
			if ($null -ne $msg.statsTotalDistancePx)       { $script:StatsTotalDistancePx       = [double]$msg.statsTotalDistancePx }
			if ($null -ne $msg.statsLastMoveDist)          { $script:StatsLastMoveDist          = [double]$msg.statsLastMoveDist }
			if ($null -ne $msg.statsMinMoveDist)           { $script:StatsMinMoveDist           = [double]$msg.statsMinMoveDist }
			if ($null -ne $msg.statsMaxMoveDist)           { $script:StatsMaxMoveDist           = [double]$msg.statsMaxMoveDist }
			if ($null -ne $msg.statsLastMoveDurationMs)    { $LastMovementDurationMs            = [int]$msg.statsLastMoveDurationMs }
			if ($null -ne $msg.statsLastMoveSecondsAgo)    { $LastMovementTime                  = (Get-Date).AddSeconds(-[int]$msg.statsLastMoveSecondsAgo) }
			if ($null -ne $msg.statsKbInterruptCount)      { $script:StatsKbInterruptCount      = [int]$msg.statsKbInterruptCount }
			if ($null -ne $msg.statsMsInterruptCount)      { $script:StatsMsInterruptCount      = [int]$msg.statsMsInterruptCount }
			if ($null -ne $msg.statsLongestCleanStreak)    { $script:StatsLongestCleanStreak    = [int]$msg.statsLongestCleanStreak }
			if ($null -ne $msg.statsAvgActualIntervalSecs) { $script:StatsAvgActualIntervalSecs = [double]$msg.statsAvgActualIntervalSecs }
			if ($null -ne $msg.statsAvgDurationMs)         { $script:StatsAvgDurationMs         = [double]$msg.statsAvgDurationMs }
			if ($null -ne $msg.statsMinDurationMs)         { $script:StatsMinDurationMs         = [int]$msg.statsMinDurationMs }
			if ($null -ne $msg.statsMaxDurationMs)         { $script:StatsMaxDurationMs         = [int]$msg.statsMaxDurationMs }
			if ($null -ne $msg.statsDirectionCounts) {
				foreach ($_sdir in @('N','NE','E','SE','S','SW','W','NW')) {
					$script:StatsDirectionCounts[$_sdir] = [double]$msg.statsDirectionCounts.$_sdir
				}
			}
			if ($null -ne $msg.statsLastCurveParams -and $null -ne $msg.statsLastCurveParams.Distance) {
				$script:StatsLastCurveParams = @{
					Distance      = [double]$msg.statsLastCurveParams.Distance
					StartArcAmt   = [double]$msg.statsLastCurveParams.StartArcAmt
					StartArcSign  = [int]$msg.statsLastCurveParams.StartArcSign
					BodyCurveAmt  = [double]$msg.statsLastCurveParams.BodyCurveAmt
					BodyCurveSign = [int]$msg.statsLastCurveParams.BodyCurveSign
					BodyCurveType = [int]$msg.statsLastCurveParams.BodyCurveType
				}
				$_ck = "$($script:StatsLastCurveParams.Distance)|$($script:StatsLastCurveParams.StartArcAmt)|$($script:StatsLastCurveParams.BodyCurveAmt)|$($script:StatsLastCurveParams.BodyCurveType)"
				if ($_ck -ne $script:_LastCurveParamKey) { $script:_LastCurveParamKey = $_ck; $script:StatsCurveAnimPending = $true }
			}
			# Suppress PSScriptAnalyzer "assigned but never used" — all variables above are
			# exported to the caller's scope via dot-invoke; these reads prevent false-positive warnings.
			$null = $endTimeStr; $null = $endTimeInt; $null = $end; $null = $cooldownActive
			$null = $secondsRemaining; $null = $mouseInputDetected; $null = $keyboardInputDetected
			$null = $_keyboardInferred; $null = $SkipUpdate; $null = $ScriptStartTime
			$null = $LastMovementDurationMs; $null = $LastMovementTime
			$null = $_viewerWorkerIterChanged; $null = $_viewerWorkerIter
		}
		'log' {
			if ($null -ne $LogArray -and -not $script:ManualPause) {
				$components = @()
				foreach ($c in $msg.components) {
					$components += @{
						priority = [int]$c.priority
						text = [string]$c.text
						shortText = [string]$c.shortText
					}
				}
				if ($LogArray.Count -gt 0 -and $LogArray.Count -ge $Rows) {
					$LogArray.RemoveAt(0)
				}
				$null = $LogArray.Add([PSCustomObject]@{ logRow = $true; components = $components })
			}
		}
		'togglePause' {
			$script:ManualPause = [bool]$msg.paused
			if ($null -ne $msg.logMsg -and $null -ne $LogArray) {
				$components = @()
				foreach ($c in $msg.logMsg.components) {
					$components += @{
						priority = [int]$c.priority
						text = [string]$c.text
						shortText = [string]$c.shortText
					}
				}
				if ($LogArray.Count -gt 0 -and $LogArray.Count -ge $Rows) {
					$LogArray.RemoveAt(0)
				}
				$null = $LogArray.Add([PSCustomObject]@{ logRow = $true; components = $components })
			}
		}
		'displaySleep' {
			$_incomingSleep = [bool]$msg.active
			if ($_incomingSleep -ne $script:DisplaySleepMode) {
				$script:DisplaySleepMode = $_incomingSleep
				if ($_incomingSleep) {
					$script:DisplaySleepActivatedAt = Get-Date
					$script:DisplaySleepPrePos      = Get-MousePosition
				} else {
					$script:DisplaySleepActivatedAt = $null
					$script:_DisplaySleepLastInputTime = Get-Date
				}
			}
		}
		'stopped' {
			$_viewerStopped = $true
			$_viewerStopReason = if ($null -ne $msg.reason) { $msg.reason } else { 'unknown' }
			# Suppress PSScriptAnalyzer false-positive — exported to caller via dot-invoke
			$null = $_viewerStopped; $null = $_viewerStopReason
		}
		'focus' { Invoke-ViewerFocusWindow }
	}
}
# Suppress PSScriptAnalyzer "assigned but never used" — $_handleIpcMsg is dot-invoked
# via ". $_handleIpcMsg" in Start-mJig.psm1. This read prevents the false-positive warning.
$null = $_handleIpcMsg