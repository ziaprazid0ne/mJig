		function Show-QuitConfirmationDialog {
			param(
				[ref]$HostWidthRef,
				[ref]$HostHeightRef
			)
			
			$script:CurrentScreenState = "dialog-quit"

	if ($DebugMode) {
		Add-DebugLogEntry -LogArray $LogArray -Message "Quit confirmation dialog opened" -ShortMessage "Quit dialog opened"
	}

		$currentHostWidth = $HostWidthRef.Value
		$currentHostHeight = $HostHeightRef.Value

		$dialogWidth = 35
		$dialogHeight = 7
		# Right edge aligns with the right border padding column
		$_bpH    = [math]::Max(1, $script:BorderPadH)
			$dialogX = [math]::Max(0, $currentHostWidth - $dialogWidth - ($_bpH - 1))
			$menuBarY = if ($null -ne $script:MenuBarY) { $script:MenuBarY } else { $currentHostHeight - 2 }
			$dialogY = [math]::Max(0, $menuBarY - 2 - $dialogHeight)
			
			$savedCursorVisible = $script:CursorVisible
			$script:CursorVisible = $false
			[Console]::Write("$($script:ESC)[?25l")
			
$checkmark = [char]::ConvertFromUtf32(0x2705)
$redX = [char]::ConvertFromUtf32(0x274C)
$buttonLayout = Get-DialogButtonLayout
$dialogIconWidth = $buttonLayout.IconWidth; $dialogBracketWidth = $buttonLayout.BracketWidth; $dialogParenOffset = $buttonLayout.ParenAdjustment
$bottomLinePadding = $dialogWidth - (13 + 2 * $dialogParenOffset + 2 * $dialogIconWidth + 2 * $dialogBracketWidth) - 1
		$_bv      = [string]$script:BoxVertical
		$_hLine   = [string]$script:BoxHorizontal
		$_inner   = $dialogWidth - 2
		$line0      = $script:BoxTopLeft    + ($_hLine * $_inner) + $script:BoxTopRight
		$lineBottom = $script:BoxBottomLeft + ($_hLine * $_inner) + $script:BoxBottomRight
			
	# Slide-up animation: reveal rows progressively from behind the menu bar
	$clipY        = $menuBarY - 1   # separator row — nothing drawn at or below this Y
		$animSteps    = $dialogHeight + 1  # steps to fully reveal the box
		$frameDelayMs = 15  # 15ms per frame — at the Windows timer floor for consistency
		for ($step = 2; $step -le ($animSteps + 1); $step += 2) {
			$s     = [math]::Min($step, $animSteps)
			$animY = $menuBarY - 1 - $s  # top of box this step (rises 2 rows each step)
			for ($r = 0; $r -lt $s -and $r -le $dialogHeight; $r++) {
					$rowY = $animY + $r
					if ($rowY -ge $clipY) { continue }  # safety: never draw over separator
					# Side padding (terminal default background) — left only
					if ($dialogX -gt 0) {
						Write-Buffer -X ($dialogX - 1) -Y $rowY -Text " "
					}
					# Background fill
					Write-Buffer -X $dialogX -Y $rowY -Text (" " * $dialogWidth) -BG $script:QuitDialogBg
					# Row content
					if ($r -eq 0) {
						Write-Buffer -X $dialogX -Y $rowY -Text $line0 -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
					} elseif ($r -eq $dialogHeight) {
						Write-Buffer -X $dialogX -Y $rowY -Text $lineBottom -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
					} else {
						Write-Buffer -X $dialogX                      -Y $rowY -Text $script:BoxVertical -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
						Write-Buffer -X ($dialogX + $dialogWidth - 1) -Y $rowY -Text $script:BoxVertical -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
					}
				}
			# On the last step the box is fully revealed; draw the blank top-padding row
			if ($s -eq $animSteps -and $animY -gt 0) {
				$aPadLeft  = [math]::Max(0, $dialogX - 1)
				$aPadWidth = $dialogWidth + ($dialogX - $aPadLeft)
				Write-Buffer -X $aPadLeft -Y ($animY - 1) -Text (" " * $aPadWidth)
			}
			Flush-Buffer
			if ($frameDelayMs -gt 0) { Start-Sleep -Milliseconds $frameDelayMs }
		}

		# Draw blank padding (terminal default background) — top and left; no bottom or right
			if ($dialogY -gt 0) {
				$padLeft  = [math]::Max(0, $dialogX - 1)
				$padWidth = $dialogWidth + ($dialogX - $padLeft)
				Write-Buffer -X $padLeft -Y ($dialogY - 1) -Text (" " * $padWidth)
			}
			for ($i = 0; $i -le $dialogHeight; $i++) {
				if ($dialogX -gt 0) {
					Write-Buffer -X ($dialogX - 1) -Y ($dialogY + $i) -Text " "
				}
			}

		# Draw dialog background
		for ($i = 0; $i -le $dialogHeight; $i++) {
			Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth) -BG $script:QuitDialogBg
		}

		# Frame: top border + title + divider + bottom border
		Write-DialogFrame -X $dialogX -Y $dialogY -Width $dialogWidth -Height $dialogHeight `
			-Title "Confirm Quit" -BorderFG $script:QuitDialogBorder -TitleFG $script:QuitDialogTitle -BG $script:QuitDialogBg

		# Row 3: prompt text
		$line3Text    = "$_bv  Are you sure you want to quit?"
		$line3Padding = Get-Padding -UsedWidth ($line3Text.Length + 1) -TotalWidth $dialogWidth
		Write-Buffer -X $dialogX -Y ($dialogY + 3) -Text ($line3Text + (" " * $line3Padding) + $_bv) -FG $script:QuitDialogText -BG $script:QuitDialogBg

		# Rows 4–5: blank
		Write-Buffer -X $dialogX -Y ($dialogY + 4) -Text ($_bv + (" " * $_inner) + $_bv) -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
		Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text ($_bv + (" " * $_inner) + $_bv) -FG $script:QuitDialogBorder -BG $script:QuitDialogBg

		# Row 6: buttons
		$applyButtonX = $dialogX + 2
		Write-Buffer -X $dialogX -Y ($dialogY + 6) -Text "$_bv " -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
		$_yesW = Write-DialogButton -X $applyButtonX -Y ($dialogY + 6) -Hotkey "y" -Suffix "es" `
			-Emoji $checkmark -EmojiColor $script:TextSuccess -TextColor $script:QuitDialogButtonText -BgColor $script:QuitDialogButtonBg -HotkeyColor $script:QuitDialogButtonHotkey
		$cancelButtonX = $applyButtonX + $_yesW + 2
		Write-Buffer -Text "  " -BG $script:QuitDialogBg
		$null = Write-DialogButton -X $cancelButtonX -Y ($dialogY + 6) -Hotkey "n" -Suffix "o" `
			-Emoji $redX -EmojiColor $script:TextError -TextColor $script:QuitDialogButtonText -BgColor $script:QuitDialogButtonBg -HotkeyColor $script:QuitDialogButtonHotkey
		Write-Buffer -Text (" " * $bottomLinePadding) -BG $script:QuitDialogBg
		Write-Buffer -Text $_bv -FG $script:QuitDialogBorder -BG $script:QuitDialogBg

Flush-Buffer

# Calculate button bounds for click detection (visible characters only)
# Button row is at dialogY + 6
$buttonRowY = $dialogY + 6
$yesButtonStartX = $dialogX + 2
$yesButtonEndX   = $dialogX + 2 + $dialogBracketWidth + $dialogIconWidth + 5 + $dialogParenOffset - 1   # bracket + icon + "(y)es"(5) - 1 inclusive
$noButtonStartX  = $dialogX + 2 + $dialogBracketWidth + $dialogIconWidth + 5 + $dialogParenOffset + 2   # after btn1 + gap(2)
$noButtonEndX    = $noButtonStartX + $dialogBracketWidth + $dialogIconWidth + 4 + $dialogParenOffset - 1 # bracket + icon + "(n)o"(4) - 1 inclusive
			
			$script:DialogButtonBounds = @{
				buttonRowY = $buttonRowY
				updateStartX = $yesButtonStartX
				updateEndX = $yesButtonEndX
				cancelStartX = $noButtonStartX
				cancelEndX = $noButtonEndX
			}
			$script:DialogButtonClick = $null
			
		# Get input
		$result    = $false
		$needsRedraw = $false
			
			:inputLoop do {
				# Check for window resize and update references
				$pshost = Get-Host
				$pswindow = $pshost.UI.RawUI
				$newWindowSize = $pswindow.WindowSize
				if ($newWindowSize.Width -ne $currentHostWidth -or $newWindowSize.Height -ne $currentHostHeight) {
					$stableSize = Invoke-DialogResize -HostWidthRef $HostWidthRef -HostHeightRef $HostHeightRef `
						-ScreenState "dialog-quit"
					$currentHostWidth  = $stableSize.Width
					$currentHostHeight = $stableSize.Height
					$needsRedraw = $true
					
				# Reposition dialog: right edge at first column of right-side border padding
				$_bpH    = [math]::Max(1, $script:BorderPadH)
				$dialogX = [math]::Max(0, $currentHostWidth - $dialogWidth - ($_bpH - 1))
				$menuBarY = if ($null -ne $script:MenuBarY) { $script:MenuBarY } else { $currentHostHeight - 2 }
				$dialogY = [math]::Max(0, $menuBarY - 2 - $dialogHeight)

			# Draw dialog background
				for ($i = 0; $i -le $dialogHeight; $i++) {
					Write-Buffer -X $dialogX -Y ($dialogY + $i) -Text (" " * $dialogWidth) -BG $script:QuitDialogBg
				}

				# Frame: top border + title + divider + bottom border
				Write-DialogFrame -X $dialogX -Y $dialogY -Width $dialogWidth -Height $dialogHeight `
					-Title "Confirm Quit" -BorderFG $script:QuitDialogBorder -TitleFG $script:QuitDialogTitle -BG $script:QuitDialogBg

				# Row 3: prompt text
				$_line3Text    = "$_bv  Are you sure you want to quit?"
				$_line3Padding = Get-Padding -UsedWidth ($_line3Text.Length + 1) -TotalWidth $dialogWidth
				Write-Buffer -X $dialogX -Y ($dialogY + 3) -Text ($_line3Text + (" " * $_line3Padding) + $_bv) -FG $script:QuitDialogText -BG $script:QuitDialogBg

				# Rows 4–5: blank
				Write-Buffer -X $dialogX -Y ($dialogY + 4) -Text ($_bv + (" " * $_inner) + $_bv) -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
				Write-Buffer -X $dialogX -Y ($dialogY + 5) -Text ($_bv + (" " * $_inner) + $_bv) -FG $script:QuitDialogBorder -BG $script:QuitDialogBg

				# Row 6: buttons
				$applyButtonX = $dialogX + 2
				Write-Buffer -X $dialogX -Y ($dialogY + 6) -Text "$_bv " -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
				$_yesW = Write-DialogButton -X $applyButtonX -Y ($dialogY + 6) -Hotkey "y" -Suffix "es" `
					-Emoji $checkmark -EmojiColor $script:TextSuccess -TextColor $script:QuitDialogButtonText -BgColor $script:QuitDialogButtonBg -HotkeyColor $script:QuitDialogButtonHotkey
				$cancelButtonX = $applyButtonX + $_yesW + 2
				Write-Buffer -Text "  " -BG $script:QuitDialogBg
				$null = Write-DialogButton -X $cancelButtonX -Y ($dialogY + 6) -Hotkey "n" -Suffix "o" `
					-Emoji $redX -EmojiColor $script:TextError -TextColor $script:QuitDialogButtonText -BgColor $script:QuitDialogButtonBg -HotkeyColor $script:QuitDialogButtonHotkey
				Write-Buffer -Text (" " * $bottomLinePadding) -BG $script:QuitDialogBg
				Write-Buffer -Text $_bv -FG $script:QuitDialogBorder -BG $script:QuitDialogBg
		
		Flush-Buffer -ClearFirst
		
		$buttonRowY = $dialogY + 6
		$yesButtonStartX = $dialogX + 2
		$yesButtonEndX   = $dialogX + 2 + $dialogBracketWidth + $dialogIconWidth + 5 + $dialogParenOffset - 1
		$noButtonStartX  = $dialogX + 2 + $dialogBracketWidth + $dialogIconWidth + 5 + $dialogParenOffset + 2
		$noButtonEndX    = $noButtonStartX + $dialogBracketWidth + $dialogIconWidth + 4 + $dialogParenOffset - 1
					
					$script:DialogButtonBounds = @{
						buttonRowY = $buttonRowY
						updateStartX = $yesButtonStartX
						updateEndX = $yesButtonEndX
						cancelStartX = $noButtonStartX
						cancelEndX = $noButtonEndX
					}
				}
				
				# Check for mouse button clicks on dialog buttons via console input buffer
				$keyProcessed = $false
				$keyInfo = $null
				$key = $null
				$char = $null
				
			$_click = Get-DialogMouseClick -PeekBuffer $script:_DialogPeekBuffer
			if ($null -ne $_click) {
				$clickX = $_click.X; $clickY = $_click.Y
				if ($clickX -lt $dialogX -or $clickX -ge ($dialogX + $dialogWidth) -or $clickY -lt $dialogY -or $clickY -gt ($dialogY + $dialogHeight)) {
					$char = "n"; $keyProcessed = $true
				} elseif ($clickY -eq $buttonRowY -and $clickX -ge $yesButtonStartX -and $clickX -le $yesButtonEndX) {
					$char = "y"; $keyProcessed = $true
				} elseif ($clickY -eq $buttonRowY -and $clickX -ge $noButtonStartX -and $clickX -le $noButtonEndX) {
					$char = "n"; $keyProcessed = $true
				}
				if ($DebugMode) {
					$clickTarget = if ($keyProcessed) { "button:$char" } else { "none" }
					Add-DebugLogEntry -LogArray $LogArray -Message "Quit dialog click at ($clickX,$clickY), target: $clickTarget" -ShortMessage "Click ($clickX,$clickY) -> $clickTarget"
				}
			}
				
				# Check for dialog button clicks detected by main loop
				if (-not $keyProcessed -and $null -ne $script:DialogButtonClick) {
					$buttonClick = $script:DialogButtonClick
					$script:DialogButtonClick = $null
					if ($buttonClick -eq "Update") { $char = "y"; $keyProcessed = $true }
					elseif ($buttonClick -eq "Cancel") { $char = "n"; $keyProcessed = $true }
				}
				
			if (-not $keyProcessed) {
				$keyInfo = Read-DialogKeyInput
				if ($null -ne $keyInfo) {
					$key = $keyInfo.Key; $char = $keyInfo.Character; $keyProcessed = $true
				}
			}
				
				if (-not $keyProcessed) {
					# No key available, sleep briefly and check for resize again
					Start-Sleep -Milliseconds 50
					continue
				}
				
				if ($char -eq "y" -or $char -eq "Y" -or $key -eq "Enter" -or $char -eq [char]13 -or $char -eq [char]10) {
					# Yes - confirm quit (Enter key also works as hidden function)
					# Debug: Log quit confirmation
				if ($DebugMode) {
					Add-DebugLogEntry -LogArray $LogArray -Message "Quit dialog: Confirmed" -ShortMessage "Quit: Yes"
				}
					$result = $true
					break
				} elseif ($char -eq "n" -or $char -eq "N" -or $char -eq "q" -or $char -eq "Q" -or $key -eq "Escape" -or ($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 27)) {
					# No - cancel quit (Escape key and 'q' key also work as hidden functions)
					# Debug: Log quit cancellation
				if ($DebugMode) {
					Add-DebugLogEntry -LogArray $LogArray -Message "Quit dialog: Canceled" -ShortMessage "Quit: No"
				}
					$result = $false
					$needsRedraw = $false  # No redraw needed on cancel
					break
				}
				
				# Clear any remaining keys in buffer after processing
				try {
					while ($Host.UI.RawUI.KeyAvailable) {
						$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
					}
				} catch {
					# Silently ignore - buffer might not be clearable
				}
			} until ($false)
			
		Invoke-DialogCleanup -DialogX $dialogX -DialogY $dialogY -DialogWidth $dialogWidth -DialogHeight $dialogHeight -SavedCursorVisible $savedCursorVisible -ClearShadow
			# Return result object with result and redraw flag
			return @{
				Result = $result
				NeedsRedraw = $needsRedraw
			}
		}
