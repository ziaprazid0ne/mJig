	function Write-DialogFrame {
		param(
			[int]$X,               # dialog top-left X
			[int]$Y,               # dialog top-left Y
			[int]$Width,           # full dialog width (includes border chars)
			[int]$Height,          # last row index; bottom border drawn at Y+Height
			[string]$Title = "",   # title text; "" skips the title row
			[string]$BorderFG,     # foreground for all border/chrome characters
			[string]$TitleFG,      # foreground for title text
			[string]$BG,           # background for all frame rows
			[switch]$NoDivider     # when set: row Y+2 is blank instead of BoxVerticalRight/Left
		)

		$inner = $Width - 2
		$hLine = [string]$script:BoxHorizontal
		$bv    = [string]$script:BoxVertical

		# Top border
		Write-Buffer -X $X -Y $Y `
			-Text ($script:BoxTopLeft + ($hLine * $inner) + $script:BoxTopRight) `
			-FG $BorderFG -BG $BG

		# Title row (Y+1) — only when a title is supplied
		if ($Title -ne "") {
			Write-Buffer -X $X -Y ($Y + 1) -Text "$bv  " -FG $BorderFG -BG $BG
			Write-Buffer -Text $Title -FG $TitleFG -BG $BG
			$titlePadding = Get-Padding -UsedWidth (3 + $Title.Length + 1) -TotalWidth $Width
			Write-Buffer -Text (" " * $titlePadding) -BG $BG
			Write-Buffer -Text $bv -FG $BorderFG -BG $BG
		}

		# Row Y+2: ├─┤ divider or blank
		if ($NoDivider) {
			Write-Buffer -X $X -Y ($Y + 2) -Text ($bv + (" " * $inner) + $bv) -FG $BorderFG -BG $BG
		} else {
			Write-Buffer -X $X -Y ($Y + 2) `
				-Text ($script:BoxVerticalRight + ($hLine * $inner) + $script:BoxVerticalLeft) `
				-FG $BorderFG -BG $BG
		}

		# Bottom border
		Write-Buffer -X $X -Y ($Y + $Height) `
			-Text ($script:BoxBottomLeft + ($hLine * $inner) + $script:BoxBottomRight) `
			-FG $BorderFG -BG $BG
	}
