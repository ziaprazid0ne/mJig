	function Write-DialogButton {
		param(
			[int]$X,
			[int]$Y,
			[string]$Hotkey,               # Single hotkey character, e.g. "a"
			[string]$Suffix,               # Text after hotkey letter, e.g. "pply"
		$Emoji = $null,                # 2-cell wide emoji string; untyped so $null stays $null
		$EmojiColor = $null,           # FG for emoji; $null = terminal default; untyped so $null stays $null
			[string]$TextColor,            # Non-hotkey text color
			[string]$BgColor,              # Button background color
			[string]$HotkeyColor           # Hotkey letter color
		)

		$contentX = $X
		if ($script:DialogButtonShowBrackets) {
			Write-Buffer -X $X -Y $Y -Text "[" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg
			$contentX = $X + 1
		}
		if ($script:DialogButtonShowIcon -and $null -ne $Emoji) {
			Write-Buffer -X $contentX -Y $Y -Text $Emoji -FG $EmojiColor -BG $BgColor -Wide
			Write-Buffer -X ($contentX + 2) -Y $Y -Text $script:DialogButtonSeparator -FG $TextColor -BG $BgColor
		} else {
			Write-Buffer -X $contentX -Y $Y -Text "" -BG $BgColor
		}
		$closingParen = if ($script:DialogButtonShowHotkeyParens) { ")" } else { "" }
		if ($script:DialogButtonShowHotkeyParens) { Write-Buffer -Text "(" -FG $TextColor -BG $BgColor }
		Write-Buffer -Text $Hotkey -FG $HotkeyColor -BG $BgColor
		Write-Buffer -Text "$($closingParen)$Suffix" -FG $TextColor -BG $BgColor
		if ($script:DialogButtonShowBrackets) { Write-Buffer -Text "]" -FG $script:DialogButtonBracketFg -BG $script:DialogButtonBracketBg }

		# Return rendered width so the caller can derive click bounds and inter-button spacing.
		$iconWidth    = if ($script:DialogButtonShowIcon -and $null -ne $Emoji) { 2 + $script:DialogButtonSeparator.Length } else { 0 }
		$bracketWidth = if ($script:DialogButtonShowBrackets) { 2 } else { 0 }
		$parenWidth   = if ($script:DialogButtonShowHotkeyParens) { 2 } else { 0 }
		return $bracketWidth + $iconWidth + $parenWidth + $Hotkey.Length + $Suffix.Length
	}
