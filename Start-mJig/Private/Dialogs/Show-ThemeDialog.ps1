	function Show-ThemeDialog {
		param(
			[ref]$HostWidthRef,
			[ref]$HostHeightRef,
			[System.IO.StreamWriter]$ViewerPipeWriter = $null
		)

		$script:CurrentScreenState = "dialog-theme"

		$currentHostWidth  = $HostWidthRef.Value
		$currentHostHeight = $HostHeightRef.Value

		$dialogWidth  = 32
		# Rows: 0=border, 1=title, 2=divider, 3=blank, 4=theme-name, 5=blank, 6=next, 7=blank, 8=apply+cancel, 9=border
		$dialogHeight = 9
		$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth  - $dialogWidth)  / 2))
		$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))

		$savedCursorVisible = $script:CursorVisible
		$script:CursorVisible = $false
		[Console]::Write("$($script:ESC)[?25l")

		$buttonLayout       = Get-DialogButtonLayout
		$dialogIconWidth    = $buttonLayout.IconWidth
		$dialogBracketWidth = $buttonLayout.BracketWidth
		$_parenWidth        = if ($script:DialogButtonShowHotkeyParens) { 2 } else { 0 }
		$_nextBtnW          = $dialogBracketWidth + $dialogIconWidth + $_parenWidth + 10  # "n" + "ext theme"
		$_applyBtnW         = $dialogBracketWidth + $dialogIconWidth + $_parenWidth + 5   # "a" + "pply"
		$_cancelBtnW        = $dialogBracketWidth + $dialogIconWidth + $_parenWidth + 6   # "c" + "ancel"

		$paletteEmoji  = [char]::ConvertFromUtf32(0x1F3A8) # artist palette
		$emojiApply    = [char]::ConvertFromUtf32(0x2705)  # checkmark
		$emojiClose    = [char]::ConvertFromUtf32(0x274C)  # red X

		$drawThemeDialog = {
			param($dx, $dy)
			$_bv    = [string]$script:BoxVertical
			$inner  = $dialogWidth - 2

			# Clear background
			for ($i = 0; $i -le $dialogHeight; $i++) {
				Write-Buffer -X $dx -Y ($dy + $i) -Text (" " * $dialogWidth) -BG $script:ThemeDialogBg
			}

			# — Lines 0/1/2/$dialogHeight: frame (top border + title + divider + bottom) --
			Write-DialogFrame -X $dx -Y $dy -Width $dialogWidth -Height $dialogHeight `
				-Title "Theme" -BorderFG $script:ThemeDialogBorder -TitleFG $script:ThemeDialogTitle -BG $script:ThemeDialogBg

			# — Line 3: blank ----------------------------------------------------
			Write-Buffer -X $dx -Y ($dy + 3) -Text ($_bv + (" " * $inner) + $_bv) -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg

			# — Line 4: current theme name (centered) ----------------------------
			$themeName   = if ($script:CurrentThemeName -ne "") { $script:CurrentThemeName } else { "default" }
			$nameLen     = $themeName.Length
			$namePad     = [math]::Max(0, [math]::Floor(($inner - $nameLen) / 2))
			$nameContent = (" " * $namePad) + $themeName + (" " * ($inner - $namePad - $nameLen))
			Write-Buffer -X $dx -Y ($dy + 4) -Text $script:BoxVertical -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg
			Write-Buffer -Text $nameContent -FG $script:ThemeDialogNameFg -BG $script:ThemeDialogBg
			Write-Buffer -Text $script:BoxVertical -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg

			# — Line 5: blank ----------------------------------------------------
			Write-Buffer -X $dx -Y ($dy + 5) -Text ($_bv + (" " * $inner) + $_bv) -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg

			# — Line 6: next-theme button ----------------------------------------
			Write-Buffer -X $dx -Y ($dy + 6) -Text "$($script:BoxVertical) " -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg
			$null = Write-DialogButton -X ($dx + 2) -Y ($dy + 6) -Hotkey "n" -Suffix "ext theme" `
				-Emoji $paletteEmoji -EmojiColor $null -TextColor $script:ThemeDialogButtonText -BgColor $script:ThemeDialogButtonBg -HotkeyColor $script:ThemeDialogButtonHotkey
			Write-Buffer -Text (" " * [math]::Max(0, $dialogWidth - 3 - $_nextBtnW)) -BG $script:ThemeDialogBg
			Write-Buffer -Text $script:BoxVertical -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg

			# — Line 7: blank ----------------------------------------------------
			Write-Buffer -X $dx -Y ($dy + 7) -Text ($_bv + (" " * $inner) + $_bv) -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg

			# — Line 8: apply + cancel buttons -----------------------------------
			$_applyX  = $dx + 2
			$_cancelX = $_applyX + $_applyBtnW + 2
			Write-Buffer -X $dx -Y ($dy + 8) -Text "$($script:BoxVertical) " -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg
			$null = Write-DialogButton -X $_applyX -Y ($dy + 8) -Hotkey "a" -Suffix "pply" `
				-Emoji $emojiApply -EmojiColor $null -TextColor $script:ThemeDialogButtonText -BgColor $script:ThemeDialogButtonBg -HotkeyColor $script:ThemeDialogButtonHotkey
			Write-Buffer -Text "  " -BG $script:ThemeDialogBg
			$null = Write-DialogButton -X $_cancelX -Y ($dy + 8) -Hotkey "c" -Suffix "ancel" `
				-Emoji $emojiClose -EmojiColor $script:TextError -TextColor $script:ThemeDialogButtonText -BgColor $script:ThemeDialogButtonBg -HotkeyColor $script:ThemeDialogButtonHotkey
			Write-Buffer -Text (" " * [math]::Max(0, $dialogWidth - 3 - $_applyBtnW - 2 - $_cancelBtnW)) -BG $script:ThemeDialogBg
			Write-Buffer -Text $script:BoxVertical -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg

		}

		# Send viewerState before opening
		if ($null -ne $ViewerPipeWriter) {
			Send-PipeMessage -Writer $ViewerPipeWriter -Message @{ type='viewerState'; activeDialog='theme' }
		}

		# Save initial theme name so Cancel can revert via Set-ThemeProfile
		$initialThemeName = $script:CurrentThemeName

		# Initial draw
		& $drawThemeDialog $dialogX $dialogY
		Write-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:ThemeDialogShadow
		Flush-Buffer

		$nextButtonRowY     = $dialogY + 6
		$nextButtonStartX   = $dialogX + 2
		$nextButtonEndX     = $nextButtonStartX + $_nextBtnW - 1
		$applyButtonRowY    = $dialogY + 8
		$applyButtonStartX  = $dialogX + 2
		$applyButtonEndX    = $applyButtonStartX + $_applyBtnW - 1
		$cancelButtonStartX = $applyButtonStartX + $_applyBtnW + 2
		$cancelButtonEndX   = $cancelButtonStartX + $_cancelBtnW - 1

		$script:DialogButtonBounds = @{
			buttonRowY   = $applyButtonRowY
			updateStartX = $applyButtonStartX
			updateEndX   = $applyButtonEndX
			cancelStartX = $cancelButtonStartX
			cancelEndX   = $cancelButtonEndX
		}
		$script:DialogButtonClick = $null

		$needsRedraw = $false

		:inputLoop do {
			# Resize check
			$pswindow = (Get-Host).UI.RawUI
			$newWindowSize = $pswindow.WindowSize
			if ($newWindowSize.Width -ne $currentHostWidth -or $newWindowSize.Height -ne $currentHostHeight) {
				$stableSize = Invoke-DialogResize -HostWidthRef $HostWidthRef -HostHeightRef $HostHeightRef `
					-ScreenState "dialog-theme"
				$currentHostWidth  = $stableSize.Width
				$currentHostHeight = $stableSize.Height
				$needsRedraw = $true
				$dialogX = [math]::Max(0, [math]::Floor(($currentHostWidth  - $dialogWidth)  / 2))
				$dialogY = [math]::Max(0, [math]::Floor(($currentHostHeight - $dialogHeight) / 2))
				& $drawThemeDialog $dialogX $dialogY
				Write-DialogShadow -dialogX $dialogX -dialogY $dialogY -dialogWidth $dialogWidth -dialogHeight $dialogHeight -shadowColor $script:ThemeDialogShadow
				Flush-Buffer -ClearFirst
			$nextButtonRowY     = $dialogY + 6
			$nextButtonStartX   = $dialogX + 2
			$nextButtonEndX     = $nextButtonStartX + $_nextBtnW - 1
			$applyButtonRowY    = $dialogY + 8
			$applyButtonStartX  = $dialogX + 2
			$applyButtonEndX    = $applyButtonStartX + $_applyBtnW - 1
			$cancelButtonStartX = $applyButtonStartX + $_applyBtnW + 2
			$cancelButtonEndX   = $cancelButtonStartX + $_cancelBtnW - 1
				$script:DialogButtonBounds = @{
					buttonRowY   = $applyButtonRowY
					updateStartX = $applyButtonStartX
					updateEndX   = $applyButtonEndX
					cancelStartX = $cancelButtonStartX
					cancelEndX   = $cancelButtonEndX
				}
			}

			# Mouse click detection
			$keyProcessed = $false
			$keyInfo = $null
			$key     = $null
			$char    = $null

			$_click = Get-DialogMouseClick -PeekBuffer $script:_DialogPeekBuffer
			if ($null -ne $_click) {
				$clickX = $_click.X; $clickY = $_click.Y
				if ($clickX -lt $dialogX -or $clickX -ge ($dialogX + $dialogWidth) -or
					$clickY -lt $dialogY -or $clickY -gt ($dialogY + $dialogHeight)) {
					$char = "c"; $keyProcessed = $true
				} elseif ($clickY -eq $nextButtonRowY -and $clickX -ge $nextButtonStartX -and $clickX -le $nextButtonEndX) {
					$char = "n"; $keyProcessed = $true
				} elseif ($clickY -eq $applyButtonRowY -and $clickX -ge $applyButtonStartX -and $clickX -le $applyButtonEndX) {
					$char = "a"; $keyProcessed = $true
				} elseif ($clickY -eq $applyButtonRowY -and $clickX -ge $cancelButtonStartX -and $clickX -le $cancelButtonEndX) {
					$char = "c"; $keyProcessed = $true
				}
			}

			if (-not $keyProcessed -and $null -ne $script:DialogButtonClick) {
				$script:DialogButtonClick = $null
				$char = "a"; $keyProcessed = $true
			}

			if (-not $keyProcessed) {
				$keyInfo = Read-DialogKeyInput
				if ($null -ne $keyInfo) {
					$key = $keyInfo.Key; $char = $keyInfo.Character; $keyProcessed = $true
				}
			}

			if (-not $keyProcessed) { Start-Sleep -Milliseconds 50; continue }

			if ($char -eq "n" -or $char -eq "N") {
				# Cycle to the next theme
				$nextIndex = ($script:CurrentThemeIndex + 1) % $script:ThemeProfiles.Count
				$nextName  = $script:ThemeProfiles[$nextIndex].Name
				Set-ThemeProfile -Name $nextName

				# Redraw the name row only
				$themeName   = $script:CurrentThemeName
				$nameLen     = $themeName.Length
				$inner       = $dialogWidth - 2
				$namePad     = [math]::Max(0, [math]::Floor(($inner - $nameLen) / 2))
				$nameContent = (" " * $namePad) + $themeName + (" " * ($inner - $namePad - $nameLen))
				Write-Buffer -X $dialogX -Y ($dialogY + 4) -Text $script:BoxVertical -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg
				Write-Buffer -Text $nameContent -FG $script:ThemeDialogNameFg -BG $script:ThemeDialogBg
				Write-Buffer -Text $script:BoxVertical -FG $script:ThemeDialogBorder -BG $script:ThemeDialogBg
				Flush-Buffer

				$needsRedraw = $true
				continue
			}

			if ($char -eq "a" -or $char -eq "A" -or $key -eq "Enter" -or
				$char -eq [char]13 -or $char -eq [char]10) {
				# Apply — keep current theme and close
				break inputLoop
			}

			if ($key -eq "Escape" -or $char -eq "c" -or $char -eq "C" -or
				($null -ne $keyInfo -and $keyInfo.VirtualKeyCode -eq 27)) {
				# Cancel — revert to theme that was active when dialog opened
				if ($script:CurrentThemeName -ne $initialThemeName) {
					Set-ThemeProfile -Name $initialThemeName
					$needsRedraw = $true
				}
				break inputLoop
			}

		} while ($true)

		try { while ($Host.UI.RawUI.KeyAvailable) { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC") } } catch {}

		# Send viewerState after close
		if ($null -ne $ViewerPipeWriter) {
			Send-PipeMessage -Writer $ViewerPipeWriter -Message @{ type='viewerState'; activeDialog=$null }
		}

		Invoke-DialogCleanup -DialogX $dialogX -DialogY $dialogY -DialogWidth $dialogWidth -DialogHeight $dialogHeight -SavedCursorVisible $savedCursorVisible -ClearShadow

		return @{ NeedsRedraw = $needsRedraw }
	}
