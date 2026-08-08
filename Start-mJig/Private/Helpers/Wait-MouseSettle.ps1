function Wait-MouseSettle {
	$mouseSettleMs      = 150
	$lastSettleCheckPos = Get-MousePosition
	$mouseSettledTime   = $null
	$settleLoopCount    = 0
	$maxMoveDelta       = 0

	if ($script:DiagEnabled) {
		"$(Get-Date -Format 'HH:mm:ss.fff') - Loop $($script:LoopIteration): Starting settle wait, pos: $($lastSettleCheckPos.X),$($lastSettleCheckPos.Y)" | Out-File $script:SettleDiagFile -Append
	}

	while ($true) {
		$settleLoopCount++
		Start-Sleep -Milliseconds 25
		$currentSettlePos = Get-MousePosition

		$mouseMoved = $false
		if ($null -ne $currentSettlePos -and $null -ne $lastSettleCheckPos) {
			$deltaX    = [Math]::Abs($currentSettlePos.X - $lastSettleCheckPos.X)
			$deltaY    = [Math]::Abs($currentSettlePos.Y - $lastSettleCheckPos.Y)
			$moveDelta = [Math]::Max($deltaX, $deltaY)
			if ($moveDelta -gt $maxMoveDelta) { $maxMoveDelta = $moveDelta }
			if ($deltaX -gt 2 -or $deltaY -gt 2) { $mouseMoved = $true }
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
