	function Invoke-GlobalHotkeyAction {
		param(
			[ValidateSet('togglePause', 'toggleDisplaySleep', 'quit')]
			[string]$Action,

			# Source label used in log entries (hotkey or tray)
			[ValidateSet('hotkey', 'tray')]
			[string]$Source = 'hotkey',

			# Mutable state refs — write back to caller variable
			[ref]$ManualPauseRef,
			[ref]$DisplaySleepModeRef,

			[datetime]$Date,

			# Worker-specific mode
			[switch]$IsWorker,

			# Pipe sync scriptblock; receives message hashtable.
			# Worker: handles viewer-connected guard + send; use GetNewClosure at definition site.
			# Quit messages must be sent synchronously; non-blocking for all others.
			[scriptblock]$PipeSyncAction = $null,

			# Inline-specific: log array for Add-LogEntry; omit in worker mode
			[System.Collections.Generic.List[object]]$LogArray = $null,
			[int]$Rows = 0
		)

		switch ($Action) {

			'togglePause' {
				$ManualPauseRef.Value = -not $ManualPauseRef.Value
				$paused = $ManualPauseRef.Value
				$body   = if ($paused) { 'Paused' } else { 'Resumed' }
				$notify = if ($paused) { 'paused' } else { 'resumed' }
				Show-Notification -Body $body -Action $notify

				if ($IsWorker) {
					try { Update-TrayPauseLabel -Paused $paused } catch {}
					$logMsg = @{
						type = 'log'
						components = @(
							@{ priority = 1; text = $Date.ToString(); shortText = $Date.ToString('HH:mm:ss') }
							@{ priority = 2; text = " - $body via $Source"; shortText = " - $body" }
						)
					}
					if ($script:LogReplayBuffer.Count -ge 30) { $null = $script:LogReplayBuffer.Dequeue() }
					$null = $script:LogReplayBuffer.Enqueue($logMsg)
					if ($null -ne $PipeSyncAction) {
						& $PipeSyncAction @{ type = 'togglePause'; paused = $paused; logMsg = $logMsg }
					}
				} else {
					$logText = " - $($script:WindowTitle) $($body.ToLower())"
					if ($null -ne $LogArray) {
						Add-LogEntry -LogArray $LogArray -Rows $Rows -Date $Date -Text "$logText via $Source" -ShortText $logText
					}
				}
			}

			'toggleDisplaySleep' {
				if (-not $DisplaySleepModeRef.Value) {
					Start-Sleep -Milliseconds 500
					$null = Invoke-DisplaySleep -Action Sleep
					$DisplaySleepModeRef.Value = $true
					if ($IsWorker) {
						if ($null -ne $PipeSyncAction) {
							& $PipeSyncAction @{ type = 'displaySleep'; active = $true }
						}
					} else {
						$script:DisplaySleepActivatedAt = $Date
						if ($null -ne $LogArray) {
							Add-LogEntry -LogArray $LogArray -Rows $Rows -Date $Date -Text " - Display sleep activated via $Source" -ShortText ' - Display sleep on'
						}
					}
				} else {
					if (Invoke-DisplaySleep -Action Wake) {
						$DisplaySleepModeRef.Value = $false
						$script:LastUserActivityTime = $Date
						if ($IsWorker) {
							if ($null -ne $PipeSyncAction) {
								& $PipeSyncAction @{ type = 'displaySleep'; active = $false }
							}
						} else {
							$script:DisplaySleepActivatedAt = $null
							if ($null -ne $LogArray) {
								Add-LogEntry -LogArray $LogArray -Rows $Rows -Date $Date -Text " - Display wake confirmed via $Source" -ShortText ' - Display wake ok'
							}
						}
					}
				}
			}

			'quit' {
				Show-Notification -Body 'Stopped' -Action quit
				if ($null -ne $PipeSyncAction) {
					& $PipeSyncAction @{ type = 'stopped'; reason = 'quit' }
				}
				return $true
			}
		}
		return $false
	}
