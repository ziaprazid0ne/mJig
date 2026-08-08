		function Test-GlobalHotkey {
			$shift = ($script:MouseAPI::GetAsyncKeyState(0x10) -band 0x8000) -ne 0
			if (-not $shift) { $script:_HotkeyDebounce = $false; return $null }

			$mKey = ($script:MouseAPI::GetAsyncKeyState(0x4D) -band 0x8000) -ne 0
			if (-not $mKey) { $script:_HotkeyDebounce = $false; return $null }

			if ($script:_HotkeyDebounce) { return $null }

			$pKey = ($script:MouseAPI::GetAsyncKeyState(0x50) -band 0x8000) -ne 0
			$qKey = ($script:MouseAPI::GetAsyncKeyState(0x51) -band 0x8000) -ne 0
			$dKey = ($script:MouseAPI::GetAsyncKeyState(0x44) -band 0x8000) -ne 0
			$sKey = ($script:MouseAPI::GetAsyncKeyState(0x53) -band 0x8000) -ne 0

			if ($pKey) { $script:_HotkeyDebounce = $true; return 'togglePause' }
			if ($qKey) { $script:_HotkeyDebounce = $true; return 'quit' }
			if ($dKey) { $script:_HotkeyDebounce = $true; return 'toggleDisplaySleep' }
			if ($sKey) { $script:_HotkeyDebounce = $true; return 'toggleDisplaySleep' }

			return $null
		}
