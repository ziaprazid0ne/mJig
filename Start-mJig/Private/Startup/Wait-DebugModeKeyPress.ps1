function Wait-DebugModeKeyPress {
	if ($script:DiagEnabled) { "$(Get-Date -Format 'HH:mm:ss.fff') - ENTERED DEBUG MODE KEY WAIT LOOP" | Out-File $script:StartupDiagFile -Append }

	Write-Host "`nPress any key to start mJig..." -ForegroundColor $script:TextWarning

	$dbgModifierVKs = @(0x10, 0x11, 0x12, 0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0x5B, 0x5C)
	$hIn      = $script:MouseAPI::GetStdHandle(-10)
	$peekBuf  = New-Object "$($script:_ApiNamespace).INPUT_RECORD[]" 32
	$peekEvts = [uint32]0
	# Drain events buffered before the prompt appeared (e.g. Enter key-up from launch)
	try { $Host.UI.RawUI.FlushInputBuffer() } catch {}
	$detected = $false
	while (-not $detected) {
		Start-Sleep -Milliseconds 5
		try {
			if ($script:MouseAPI::PeekConsoleInput($hIn, $peekBuf, 32, [ref]$peekEvts) -and $peekEvts -gt 0) {
				for ($e = 0; $e -lt [int]$peekEvts; $e++) {
					if ($peekBuf[$e].EventType -eq 0x0001 -and $peekBuf[$e].KeyEvent.bKeyDown -eq 0 -and
					    $peekBuf[$e].KeyEvent.wVirtualKeyCode -notin $dbgModifierVKs) {
						$detected = $true; break
					}
				}
				if ($detected) {
					$flushBuf = New-Object "$($script:_ApiNamespace).INPUT_RECORD[]" $peekEvts
					$flushed  = [uint32]0
					$script:MouseAPI::ReadConsoleInput($hIn, $flushBuf, $peekEvts, [ref]$flushed) | Out-Null
				}
			}
		} catch {
			if ($Host.UI.RawUI.KeyAvailable) {
				try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp,AllowCtrlC") } catch {}
				$detected = $true
			} else {
				Start-Sleep -Milliseconds 45
			}
		}
	}
}
