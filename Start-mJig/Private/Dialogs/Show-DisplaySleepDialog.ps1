	function Show-DisplaySleepDialog {
		param(
			[ref]$HostWidthRef,
			[ref]$HostHeightRef
		)

		$script:CurrentScreenState = "dialog-display-sleep"
		$currentHostWidth  = $HostWidthRef.Value
		$currentHostHeight = $HostHeightRef.Value

		$dialogWidth  = 40
		# Rows: 0=border,1=title,2=divider,3=blank,4-6=info,7=blank,
		#       8=audio,9=blank,10=auto-sleep,11=blank,12=timeout,
		#       13=blank,14=buttons,15=blank,16=border
		$dialogHeight = 16

		$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth  - $dialogWidth)  / 2))
		$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))

		$savedCursorVisible = $script:CursorVisible
		$script:CursorVisible = $true
		[Console]::Write("$($script:ESC)[?25h")

		# Box strings — [string] casts prevent [char]*int integer arithmetic
		$_bh    = [string]$script:BoxHorizontal
		$_inner = $dialogWidth - 2
		$line0  = [string]$script:BoxTopLeft    + ($_bh * $_inner) + [string]$script:BoxTopRight
		$line2  = [string]$script:BoxVerticalRight + ($_bh * $_inner) + [string]$script:BoxVerticalLeft
		$lineBl = [string]$script:BoxVertical   + (" " * $_inner) + [string]$script:BoxVertical
		$lineB  = [string]$script:BoxBottomLeft + ($_bh * $_inner) + [string]$script:BoxBottomRight

		$emojiSpeaker = [char]::ConvertFromUtf32(0x1F50A)  # 🔊
		$emojiClock   = [char]::ConvertFromUtf32(0x23F0)   # ⏰
		$emojiTimer   = [char]::ConvertFromUtf32(0x231B)   # ⌛
		$emojiZzz     = [char]::ConvertFromUtf32(0x1F4A4)  # 💤
		$emojiApply   = [char]::ConvertFromUtf32(0x2705)   # ✅
		$emojiClose   = [char]::ConvertFromUtf32(0x274C)   # ❌

		$buttonLayout       = Get-DialogButtonLayout
		$dialogIconWidth    = $buttonLayout.IconWidth
		$dialogBracketWidth = $buttonLayout.BracketWidth
		$dialogParenOffset  = $buttonLayout.ParenAdjustment

		# (s)leep = 7 chars, (a)pply = 7 chars, (c)ancel = 8 chars
		$sleepBtnW  = $dialogBracketWidth + $dialogIconWidth + 7 + $dialogParenOffset
		$applyBtnW  = $dialogBracketWidth + $dialogIconWidth + 7 + $dialogParenOffset
		$cancelBtnW = $dialogBracketWidth + $dialogIconWidth + 8 + $dialogParenOffset
		# trailing pad in button row before closing │
		$btnTrail   = [math]::Max(0, $dialogWidth - 2 - $sleepBtnW - 2 - $applyBtnW - 2 - $cancelBtnW - 1)

		# Capture initial values for Cancel revert
		$initialAudio       = $script:DisplaySleepAudioEnabled
		$initialAutoEnabled = $script:DisplaySleepAutoEnabled
		$initialAutoTimeout = $script:DisplaySleepAutoTimeoutSecs

		# Local timeout input string (left-aligned in 4-char field)
		$timeoutStr  = [string]$script:DisplaySleepAutoTimeoutSecs
		if ($timeoutStr -eq "0" -or $timeoutStr -eq "") { $timeoutStr = "60" }

		$needsRedraw = $false
		$action      = 'cancel'

		# ── draw helpers ────────────────────────────────────────────────────

		$drawToggleRow = {
			param($DX, $RowY, $emoji, $hotkeyChar, $labelSuffix)
			$localBg  = $script:SettingsDialogBg
			$rowPad   = [math]::Max(0, $dialogWidth - (2 + $dialogBracketWidth + $dialogIconWidth + 3 + $labelSuffix.Length + $dialogParenOffset + 1))
			$cp       = if ($script:DialogButtonShowHotkeyParens) { ")" } else { "" }
			$btnX     = $DX + 2
			Write-Buffer -X $DX -Y $RowY -Text "$([string]$script:BoxVertical) " -FG $script:SettingsDialogBorder -BG $localBg
			if ($script:DialogButtonShowBrackets) { Write-Buffer -X $btnX -Y $RowY -Text "[" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			$cx = $btnX + [int]$script:DialogButtonShowBrackets
			if ($script:DialogButtonShowIcon) {
				Write-Buffer -X $cx -Y $RowY -Text $emoji -BG $script:SettingsDialogButtonBg -Wide
				Write-Buffer -X ($cx + 2) -Y $RowY -Text $script:DialogButtonSeparator -FG $script:SettingsDialogButtonText -BG $script:SettingsDialogButtonBg
			} else {
				Write-Buffer -X $cx -Y $RowY -Text "" -BG $script:SettingsDialogButtonBg
			}
			if ($script:DialogButtonShowHotkeyParens) { Write-Buffer -Text "(" -FG $script:SettingsDialogButtonText -BG $script:SettingsDialogButtonBg }
			Write-Buffer -Text $hotkeyChar -FG $script:SettingsDialogButtonHotkey -BG $script:SettingsDialogButtonBg
			Write-Buffer -Text "$cp$labelSuffix" -FG $script:SettingsDialogButtonText -BG $script:SettingsDialogButtonBg
			if ($script:DialogButtonShowBrackets) { Write-Buffer -Text "]" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			Write-Buffer -Text (" " * $rowPad) -BG $localBg
			Write-Buffer -Text "$([string]$script:BoxVertical)" -FG $script:SettingsDialogBorder -BG $localBg
		}

		$drawTimeoutRow = {
			param($DX, $RowY)
			$localBg  = $script:SettingsDialogBg
			$fieldBg  = if ($script:DisplaySleepAutoEnabled) { $script:SettingsDialogButtonBg } else { $localBg }
			$fieldFg  = if ($script:DisplaySleepAutoEnabled) { $script:SettingsDialogButtonText } else { $script:TextMuted }
			$labelFg  = if ($script:DisplaySleepAutoEnabled) { $script:SettingsDialogText } else { $script:TextMuted }
			$fdisp    = ($timeoutStr + "    ").Substring(0, 4)
			# Positions: │(DX) + "  "(+2) + ⏳(+2wide) + " Timeout: "(+10) + [(+1) + field(+4) + ](+1) + " sec"(+4) + trail + │
			# Trailing = dialogWidth - (1+2+2+10+1+4+1+4+1) = dialogWidth - 26
			$trail = [math]::Max(0, $dialogWidth - 26)
			Write-Buffer -X $DX -Y $RowY -Text "$([string]$script:BoxVertical)  " -FG $script:SettingsDialogBorder -BG $localBg
			Write-Buffer -X ($DX + 3) -Y $RowY -Text $emojiTimer -FG $labelFg -BG $localBg -Wide
			Write-Buffer -X ($DX + 5) -Y $RowY -Text " Timeout: " -FG $labelFg -BG $localBg
			Write-Buffer -X ($DX + 15) -Y $RowY -Text "[" -FG $script:SettingsDialogBorder -BG $localBg
			Write-Buffer -X ($DX + 16) -Y $RowY -Text $fdisp -FG $fieldFg -BG $fieldBg
			Write-Buffer -X ($DX + 20) -Y $RowY -Text "] sec" -FG $labelFg -BG $localBg
			Write-Buffer -X ($DX + 25) -Y $RowY -Text (" " * $trail) -BG $localBg
			Write-Buffer -Text "$([string]$script:BoxVertical)" -FG $script:SettingsDialogBorder -BG $localBg
		}

		$drawButtonRow = {
			param($DX, $RowY)
			$localBg    = $script:SettingsDialogBg
			$localBtnBg = $script:SettingsDialogButtonBg
			$localBtnFg = $script:SettingsDialogButtonText
			$localBtnHk = $script:SettingsDialogButtonHotkey
			$cp         = if ($script:DialogButtonShowHotkeyParens) { ")" } else { "" }
			Write-Buffer -X $DX -Y $RowY -Text "$([string]$script:BoxVertical) " -FG $script:SettingsDialogBorder -BG $localBg
			# Sleep button
			$slpX = $DX + 2
			if ($script:DialogButtonShowBrackets) { Write-Buffer -X $slpX -Y $RowY -Text "[" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			$scx = $slpX + [int]$script:DialogButtonShowBrackets
			if ($script:DialogButtonShowIcon) {
				Write-Buffer -X $scx -Y $RowY -Text $emojiZzz -BG $localBtnBg -Wide
				Write-Buffer -X ($scx + 2) -Y $RowY -Text $script:DialogButtonSeparator -FG $localBtnFg -BG $localBtnBg
			} else { Write-Buffer -X $scx -Y $RowY -Text "" -BG $localBtnBg }
			if ($script:DialogButtonShowHotkeyParens) { Write-Buffer -Text "(" -FG $localBtnFg -BG $localBtnBg }
			Write-Buffer -Text "s" -FG $localBtnHk -BG $localBtnBg
			Write-Buffer -Text "${cp}leep" -FG $localBtnFg -BG $localBtnBg
			if ($script:DialogButtonShowBrackets) { Write-Buffer -Text "]" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			Write-Buffer -Text "  " -BG $localBg
			# Apply button
			$aplX = $slpX + $sleepBtnW + 2
			if ($script:DialogButtonShowBrackets) { Write-Buffer -X $aplX -Y $RowY -Text "[" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			$acx = $aplX + [int]$script:DialogButtonShowBrackets
			if ($script:DialogButtonShowIcon) {
				Write-Buffer -X $acx -Y $RowY -Text $emojiApply -BG $localBtnBg -Wide
				Write-Buffer -X ($acx + 2) -Y $RowY -Text $script:DialogButtonSeparator -FG $localBtnFg -BG $localBtnBg
			} else { Write-Buffer -X $acx -Y $RowY -Text "" -BG $localBtnBg }
			if ($script:DialogButtonShowHotkeyParens) { Write-Buffer -Text "(" -FG $localBtnFg -BG $localBtnBg }
			Write-Buffer -Text "a" -FG $localBtnHk -BG $localBtnBg
			Write-Buffer -Text "${cp}pply" -FG $localBtnFg -BG $localBtnBg
			if ($script:DialogButtonShowBrackets) { Write-Buffer -Text "]" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			Write-Buffer -Text "  " -BG $localBg
			# Cancel button
			$cncX = $aplX + $applyBtnW + 2
			if ($script:DialogButtonShowBrackets) { Write-Buffer -X $cncX -Y $RowY -Text "[" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			$ccx = $cncX + [int]$script:DialogButtonShowBrackets
			if ($script:DialogButtonShowIcon) {
				Write-Buffer -X $ccx -Y $RowY -Text $emojiClose -FG $script:TextError -BG $localBtnBg -Wide
				Write-Buffer -X ($ccx + 2) -Y $RowY -Text $script:DialogButtonSeparator -FG $localBtnFg -BG $localBtnBg
			} else { Write-Buffer -X $ccx -Y $RowY -Text "" -BG $localBtnBg }
			if ($script:DialogButtonShowHotkeyParens) { Write-Buffer -Text "(" -FG $localBtnFg -BG $localBtnBg }
			Write-Buffer -Text "c" -FG $localBtnHk -BG $localBtnBg
			Write-Buffer -Text "${cp}ancel" -FG $localBtnFg -BG $localBtnBg
			if ($script:DialogButtonShowBrackets) { Write-Buffer -Text "]" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }
			Write-Buffer -Text (" " * $btnTrail) -BG $localBg
			Write-Buffer -Text "$([string]$script:BoxVertical)" -FG $script:SettingsDialogBorder -BG $localBg
		}

		$drawDialog = {
			param($DX, $DY)
			$localBg  = $script:SettingsDialogBg
			$localBd  = $script:SettingsDialogBorder
			$localTx  = $script:SettingsDialogText

			for ($i = 0; $i -le $dialogHeight; $i++) {
				$ry = $DY + $i
				Write-Buffer -X $DX -Y $ry -Text (" " * $dialogWidth) -BG $localBg
				if ($i -eq 0) {
					Write-Buffer -X $DX -Y $ry -Text $line0 -FG $localBd -BG $localBg
				} elseif ($i -eq 1) {
					$tpad = Get-Padding -UsedWidth (4 + "Sleep Display".Length + 1) -TotalWidth $dialogWidth
					Write-Buffer -X $DX -Y $ry -Text "$([string]$script:BoxVertical)  " -FG $localBd -BG $localBg
					Write-Buffer -Text "Sleep Display" -FG $script:SettingsDialogTitle -BG $localBg
					Write-Buffer -Text (" " * $tpad) -BG $localBg
					Write-Buffer -Text "$([string]$script:BoxVertical)" -FG $localBd -BG $localBg
				} elseif ($i -eq 2) {
					Write-Buffer -X $DX -Y $ry -Text $line2 -FG $localBd -BG $localBg
				} elseif ($i -eq 4) {
					$t = "  Sends a DDC/CI power command"
					$p = Get-Padding -UsedWidth ($t.Length + 2) -TotalWidth $dialogWidth
					Write-Buffer -X $DX -Y $ry -Text "$([string]$script:BoxVertical)$t$(' ' * $p)$([string]$script:BoxVertical)" -FG $localTx -BG $localBg
				} elseif ($i -eq 5) {
					$t = "  directly to the monitor firmware."
					$p = Get-Padding -UsedWidth ($t.Length + 2) -TotalWidth $dialogWidth
					Write-Buffer -X $DX -Y $ry -Text "$([string]$script:BoxVertical)$t$(' ' * $p)$([string]$script:BoxVertical)" -FG $localTx -BG $localBg
				} elseif ($i -eq 6) {
					$t = "  Apps remain active."
					$p = Get-Padding -UsedWidth ($t.Length + 2) -TotalWidth $dialogWidth
					Write-Buffer -X $DX -Y $ry -Text "$([string]$script:BoxVertical)$t$(' ' * $p)$([string]$script:BoxVertical)" -FG $localTx -BG $localBg
				} elseif ($i -eq 8) {
					$auSuffix = if ($script:DisplaySleepAudioEnabled) { "dio cues: On " } else { "dio cues: Off" }
					& $drawToggleRow $DX $ry $emojiSpeaker "u" $auSuffix
				} elseif ($i -eq 10) {
					$autoSuffix = if ($script:DisplaySleepAutoEnabled) { "imed sleep: On " } else { "imed sleep: Off" }
					& $drawToggleRow $DX $ry $emojiClock "t" $autoSuffix
				} elseif ($i -eq 12) {
					& $drawTimeoutRow $DX $ry
				} elseif ($i -eq 14) {
					& $drawButtonRow $DX $ry
				} elseif ($i -eq $dialogHeight) {
					Write-Buffer -X $DX -Y $ry -Text $lineB -FG $localBd -BG $localBg
				} elseif ($i -eq 3 -or $i -eq 7 -or $i -eq 9 -or $i -eq 11 -or $i -eq 13 -or $i -eq 15) {
					Write-Buffer -X $DX -Y $ry -Text $lineBl -FG $localTx -BG $localBg
				}
			}
		}

		& $drawDialog $dialogX $dialogY
		Write-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:SettingsDialogShadow
		Flush-Buffer
		[Console]::SetCursorPosition($dialogX + 16 + $timeoutStr.Length, $dialogY + 12)

		# Click-detection button bounds for the main-loop menu handler
		$buttonRowY = $dialogY + 14
		$slpStartX  = $dialogX + 2
		$slpEndX    = $slpStartX + $sleepBtnW - 1
		$aplStartX  = $slpStartX + $sleepBtnW + 2
		$aplEndX    = $aplStartX + $applyBtnW - 1
		$cncStartX  = $aplStartX + $applyBtnW + 2
		$cncEndX    = $cncStartX + $cancelBtnW - 1

		$script:DialogButtonBounds = @{
			buttonRowY   = $buttonRowY
			updateStartX = $slpStartX
			updateEndX   = $slpEndX
			cancelStartX = $cncStartX
			cancelEndX   = $cncEndX
		}
		$script:DialogButtonClick = $null

		:dlgLoop do {
			# Resize check
			$newWindowSize = (Get-Host).UI.RawUI.WindowSize
			if ($newWindowSize.Width -ne $currentHostWidth -or $newWindowSize.Height -ne $currentHostHeight) {
				$stableSize          = Invoke-ResizeHandler -PreviousScreenState "dialog-display-sleep"
				$HostWidthRef.Value  = $stableSize.Width
				$HostHeightRef.Value = $stableSize.Height
				$currentHostWidth    = $stableSize.Width
				$currentHostHeight   = $stableSize.Height
				$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth  - $dialogWidth)  / 2))
				$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
				$needsRedraw = $true
				Write-MainFrame -Force -NoFlush
				& $drawDialog $dialogX $dialogY
				Write-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:SettingsDialogShadow
				Flush-Buffer -ClearFirst
				[Console]::SetCursorPosition($dialogX + 16 + $timeoutStr.Length, $dialogY + 12)
				$buttonRowY = $dialogY + 14
				$slpStartX  = $dialogX + 2
				$slpEndX    = $slpStartX + $sleepBtnW - 1
				$aplStartX  = $slpStartX + $sleepBtnW + 2
				$aplEndX    = $aplStartX + $applyBtnW - 1
				$cncStartX  = $aplStartX + $applyBtnW + 2
				$cncEndX    = $cncStartX + $cancelBtnW - 1
				$script:DialogButtonBounds = @{
					buttonRowY   = $buttonRowY
					updateStartX = $slpStartX
					updateEndX   = $slpEndX
					cancelStartX = $cncStartX
					cancelEndX   = $cncEndX
				}
			}

			$keyProcessed = $false
			$char = $null; $key = $null; $keyInfo = $null

			# Mouse clicks
			$_click = Get-DialogMouseClick -PeekBuffer $script:_DialogPeekBuffer
			if ($null -ne $_click) {
				$cx = $_click.X; $cy = $_click.Y
				if ($cx -lt $dialogX -or $cx -ge ($dialogX + $dialogWidth) -or $cy -lt $dialogY -or $cy -gt ($dialogY + $dialogHeight)) {
					$char = "c"; $keyProcessed = $true
				} else {
					$row = $cy - $dialogY
					if ($row -eq 8)  { $char = "u"; $keyProcessed = $true }
					elseif ($row -eq 10) { $char = "t"; $keyProcessed = $true }
					elseif ($row -eq 14) {
						if ($cx -ge $slpStartX -and $cx -le $slpEndX)   { $char = "s"; $keyProcessed = $true }
						elseif ($cx -ge $aplStartX -and $cx -le $aplEndX) { $char = "a"; $keyProcessed = $true }
						elseif ($cx -ge $cncStartX -and $cx -le $cncEndX) { $char = "c"; $keyProcessed = $true }
					}
				}
			}

			if (-not $keyProcessed -and $null -ne $script:DialogButtonClick) {
				$bc = $script:DialogButtonClick; $script:DialogButtonClick = $null
				if ($bc -eq "Update") { $char = "s"; $keyProcessed = $true }
				elseif ($bc -eq "Cancel") { $char = "c"; $keyProcessed = $true }
			}

			if (-not $keyProcessed) {
				$keyInfo = Read-DialogKeyInput
				if ($null -ne $keyInfo) { $key = $keyInfo.Key; $char = $keyInfo.Character; $keyProcessed = $true }
			}

			if (-not $keyProcessed) { Start-Sleep -Milliseconds 50; continue }

			# Dispatch
			if ($char -eq "u" -or $char -eq "U") {
				$script:DisplaySleepAudioEnabled = -not $script:DisplaySleepAudioEnabled
				$auSuffix = if ($script:DisplaySleepAudioEnabled) { "dio cues: On " } else { "dio cues: Off" }
				& $drawToggleRow $dialogX ($dialogY + 8) $emojiSpeaker "u" $auSuffix
				Flush-Buffer
				[Console]::SetCursorPosition($dialogX + 16 + $timeoutStr.Length, $dialogY + 12)
			} elseif ($char -eq "t" -or $char -eq "T") {
				$script:DisplaySleepAutoEnabled = -not $script:DisplaySleepAutoEnabled
				$needsRedraw = $true
				$autoSuffix = if ($script:DisplaySleepAutoEnabled) { "imed sleep: On " } else { "imed sleep: Off" }
				& $drawToggleRow $dialogX ($dialogY + 10) $emojiClock "t" $autoSuffix
				& $drawTimeoutRow $dialogX ($dialogY + 12)
				Flush-Buffer
				[Console]::SetCursorPosition($dialogX + 16 + $timeoutStr.Length, $dialogY + 12)
			} elseif ($null -ne $char -and [char]::IsDigit($char)) {
				# Digit key — append to timeout field (max 4 digits)
				if ($timeoutStr.Length -lt 4) {
					$timeoutStr += [string]$char
					& $drawTimeoutRow $dialogX ($dialogY + 12)
					Flush-Buffer
					[Console]::SetCursorPosition($dialogX + 16 + $timeoutStr.Length, $dialogY + 12)
				}
			} elseif ($key -eq "Backspace" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 8)) {
				if ($timeoutStr.Length -gt 0) {
					$timeoutStr = $timeoutStr.Substring(0, $timeoutStr.Length - 1)
					& $drawTimeoutRow $dialogX ($dialogY + 12)
					Flush-Buffer
					[Console]::SetCursorPosition($dialogX + 16 + $timeoutStr.Length, $dialogY + 12)
				}
			} elseif ($char -eq "s" -or $char -eq "S" -or $key -eq "Enter" -or $char -eq [char]13) {
				$action = 'sleep'; break :dlgLoop
			} elseif ($char -eq "a" -or $char -eq "A") {
				$action = 'apply'; break :dlgLoop
			} elseif ($char -eq "c" -or $char -eq "C" -or $key -eq "Escape" -or
					($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 27)) {
				$action = 'cancel'
				$script:DisplaySleepAudioEnabled = $initialAudio
				$script:DisplaySleepAutoEnabled  = $initialAutoEnabled
				break :dlgLoop
			}

			try {
				while ($Host.UI.RawUI.KeyAvailable) { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC") }
			} catch {}
		} until ($false)

		# Commit timeout on apply or sleep; revert on cancel
		if ($action -ne 'cancel') {
			$parsedTimeout = 0
			if ([int]::TryParse($timeoutStr, [ref]$parsedTimeout) -and $parsedTimeout -gt 0) {
				$script:DisplaySleepAutoTimeoutSecs = $parsedTimeout
			} else {
				$script:DisplaySleepAutoTimeoutSecs = $initialAutoTimeout
			}
		} else {
			$script:DisplaySleepAutoTimeoutSecs = $initialAutoTimeout
		}

		Invoke-DialogCleanup -DialogX $dialogX -DialogY $dialogY -DialogWidth $dialogWidth -DialogHeight $dialogHeight -SavedCursorVisible $savedCursorVisible -ClearShadow

		return @{ Action = $action; NeedsRedraw = $needsRedraw }
	}
