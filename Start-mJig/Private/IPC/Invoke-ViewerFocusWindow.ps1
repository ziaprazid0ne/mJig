function Invoke-ViewerFocusWindow {
	# Resolves the viewer's actual terminal window handle (cached in $script:_ViewerTerminalHwnd)
	# and brings it to the foreground. Called when the worker sends a 'focus' IPC message.
	if ($null -eq $script:_ViewerTerminalHwnd) {
		$_fHwnd = $script:MouseAPI::GetConsoleWindow()
		if ($_fHwnd -ne [IntPtr]::Zero) {
			$_fRoot = $script:MouseAPI::GetAncestor($_fHwnd, 3)
			if ($_fRoot -ne [IntPtr]::Zero) { $_fHwnd = $_fRoot }
		}
		if ($_fHwnd -eq [IntPtr]::Zero -or -not $script:MouseAPI::IsWindowVisible($_fHwnd)) {
			# Windows Terminal: GetConsoleWindow returned a hidden ConPTY pseudo-window.
			# Walk the parent process chain to find the actual terminal window.
			$_fWalkPid = $PID
			$_fSkip    = @('pwsh','powershell','powershell_ise','cmd','conhost','openconsole',
			               'csrss','wininit','services','svchost','lsass','system','idle')
			$_fAllow   = @('windowsterminal','alacritty','wezterm-gui','wezterm','mintty',
			               'conemu64','conemuc64','cmder','hyper','terminus','tabby','fluent-terminal')
			$_fVisited = @{}
			while ($_fWalkPid -gt 0 -and -not $_fVisited.ContainsKey($_fWalkPid)) {
				$_fVisited[$_fWalkPid] = $true
				try {
					$_fWmi = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $_fWalkPid" -EA SilentlyContinue
					if (-not $_fWmi) { break }
					$_fParentPid = [int]$_fWmi.ParentProcessId
					if ($_fParentPid -le 0 -or $_fParentPid -eq $_fWalkPid) { break }
					$_fParent = Get-Process -Id $_fParentPid -EA SilentlyContinue
					if (-not $_fParent) { break }
					$_fExe = [System.IO.Path]::GetFileNameWithoutExtension($_fParent.ProcessName).ToLower()
					if ($_fExe -in $_fAllow) {
						$_fHwnd = $script:MouseAPI::FindMainWindowByProcessId($_fParentPid)
						break
					} elseif ($_fExe -in $_fSkip) {
						$_fWalkPid = $_fParentPid
					} else { break }
				} catch { break }
			}
		}
		$script:_ViewerTerminalHwnd = $_fHwnd
	}
	$_fwHwnd = $script:_ViewerTerminalHwnd
	if ($_fwHwnd -ne [IntPtr]::Zero) {
		if ($script:MouseAPI::IsIconic($_fwHwnd)) {
			$null = $script:MouseAPI::ShowWindow($_fwHwnd, 9)
			$_restoreMs = 0
			while ($script:MouseAPI::IsIconic($_fwHwnd) -and $_restoreMs -lt 500) {
				Start-Sleep -Milliseconds 20
				$_restoreMs += 20
			}
			# Fallback for Windows Terminal: PostMessage WM_SYSCOMMAND/SC_RESTORE.
			if ($script:MouseAPI::IsIconic($_fwHwnd)) {
				$null = $script:MouseAPI::PostMessage($_fwHwnd, [uint32]0x0112, [IntPtr]0xF120, [IntPtr]::Zero)
				$_restoreMs = 0
				while ($script:MouseAPI::IsIconic($_fwHwnd) -and $_restoreMs -lt 500) {
					Start-Sleep -Milliseconds 20
					$_restoreMs += 20
				}
			}
		}
		$_foregroundWindow = $script:MouseAPI::GetForegroundWindow()
		$_foregroundThread = $script:MouseAPI::GetWindowThreadProcessId($_foregroundWindow, [ref]0)
		$_currentThread    = $script:MouseAPI::GetCurrentThreadId()
		if ($_foregroundThread -ne 0 -and $_foregroundThread -ne $_currentThread) {
			$null = $script:MouseAPI::AttachThreadInput($_foregroundThread, $_currentThread, $true)
		}
		$null = $script:MouseAPI::BringWindowToTop($_fwHwnd)
		$null = $script:MouseAPI::SetForegroundWindow($_fwHwnd)
		if ($_foregroundThread -ne 0 -and $_foregroundThread -ne $_currentThread) {
			$null = $script:MouseAPI::AttachThreadInput($_foregroundThread, $_currentThread, $false)
		}
	}
}
