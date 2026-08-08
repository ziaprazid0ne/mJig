function Invoke-ModuleRunspaceProvisioner {
	param(
		[string]$ModuleRoot,
		[System.Collections.Generic.Dictionary[string, object]]$BoundParameters,
		[switch]$DebugMode
	)

	$_modPath = Join-Path $ModuleRoot 'Start-mJig.psm1'

	# Save process-global console state (restored on exit)
	$_savedTitle = try { $Host.UI.RawUI.WindowTitle } catch { $null }
	$_k32Loaded = $false
	try {
		try { [void][_mJigProv._K32] } catch {
			Add-Type -Name '_K32' -Namespace '_mJigProv' -ErrorAction Stop -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int n);
[DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out uint m);
[DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, uint m);
'@
		}
		$_hIn  = [_mJigProv._K32]::GetStdHandle(-10)
		$_hOut = [_mJigProv._K32]::GetStdHandle(-11)
		$_savedInMode = [uint32]0; $_savedOutMode = [uint32]0
		$null = [_mJigProv._K32]::GetConsoleMode($_hIn,  [ref]$_savedInMode)
		$null = [_mJigProv._K32]::GetConsoleMode($_hOut, [ref]$_savedOutMode)
		$_k32Loaded = $true
	} catch { }

	# Minimal InitialSessionState — no profiles, no snap-ins, no Format-Table overhead
	$_iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
	$_iss.Formats.Clear()
	$_iss.ThrowOnRunspaceOpenError = $true

	$_rs  = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $_iss)
	$_rs.ApartmentState = [System.Threading.ApartmentState]::STA
	$_rs.Open()
	$_ps = [System.Management.Automation.PowerShell]::Create()
	$_ps.Runspace = $_rs

	if ($DebugMode) {
		$_parentRsId = if ([System.Management.Automation.Runspaces.Runspace]::DefaultRunspace) {
			[System.Management.Automation.Runspaces.Runspace]::DefaultRunspace.Id } else { '(none)' }
		Write-Host "[RUNSPACE] Provisioner: launching isolated child" -ForegroundColor Cyan
		Write-Host "  Parent Runspace  : ID $_parentRsId"  -ForegroundColor Gray
		Write-Host "  Child  Runspace  : ID $($_rs.Id)"    -ForegroundColor Gray
		Write-Host "  Thread           : $([System.Threading.Thread]::CurrentThread.ManagedThreadId)" -ForegroundColor Gray
		Write-Host "  ApartmentState   : $($_rs.ApartmentState)" -ForegroundColor Gray
		Write-Host "  ISS base         : CreateDefault2 (Core + Utility + Management)" -ForegroundColor Gray
		Write-Host "  ISS Formats      : $($_iss.Formats.Count) (cleared)" -ForegroundColor Gray
		Write-Host "  Console IN mode  : 0x$($_savedInMode.ToString('X4'))  $(if ($_k32Loaded) {'(saved)'} else {'(save failed)'})" -ForegroundColor Gray
		Write-Host "  Console OUT mode : 0x$($_savedOutMode.ToString('X4'))  $(if ($_k32Loaded) {'(saved)'} else {'(save failed)'})" -ForegroundColor Gray
		Write-Host "  Window title     : $_savedTitle (saved)" -ForegroundColor Gray
		Write-Host ""
	}

	# Instant terminal-close handler — fires on CTRL_CLOSE_EVENT before PS grace period
	try { [void][_mJigCloseHandlerX] } catch {
		Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class _mJigCloseHandlerX {
    public delegate bool HandlerDelegate(uint ctrlType);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetConsoleCtrlHandler(HandlerDelegate h, bool add);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool TerminateProcess(IntPtr hProcess, uint exitCode);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetCurrentProcess();
    private static HandlerDelegate _delegate;
    public static void Register() {
        _delegate = new HandlerDelegate(OnClose);
        SetConsoleCtrlHandler(_delegate, true);
    }
    public static void Unregister() {
        if (_delegate != null) { SetConsoleCtrlHandler(_delegate, false); _delegate = null; }
    }
    private static bool OnClose(uint t) {
        if (t == 2 || t == 5 || t == 6) { TerminateProcess(GetCurrentProcess(), 0); }
        return false;
    }
}
'@ -ErrorAction SilentlyContinue
	}
	$_closeHandlerRegistered = $false
	try { [_mJigCloseHandlerX]::Register(); $_closeHandlerRegistered = $true } catch {}

	try {
		$null = $_ps.AddScript("Import-Module '$_modPath'")
		$null = $_ps.AddStatement().AddCommand('Start-mJig')
		$null = $_ps.AddParameter('_InModuleRunspace', $true)
		foreach ($_kvp in $BoundParameters.GetEnumerator()) {
			$null = $_ps.AddParameter($_kvp.Key, $_kvp.Value)
		}
		# BeginInvoke + poll — keeps outer thread interruptible for instant close
		$_asyncResult = $_ps.BeginInvoke()
		while (-not $_asyncResult.IsCompleted) {
			Start-Sleep -Milliseconds 50
		}
		try { $null = $_ps.EndInvoke($_asyncResult) } catch {}
		if ($_ps.HadErrors -and $DebugMode) {
			$_errStream = $_ps.Streams.Error
			if ($_errStream.Count -gt 0) {
				$_uniqueErrors = @{}
				foreach ($_err in $_errStream) {
					$_key = "$($_err.Exception.Message)|$($_err.InvocationInfo.ScriptLineNumber)"
					if (-not $_uniqueErrors.ContainsKey($_key)) {
						$_uniqueErrors[$_key] = @{ Error = $_err; Count = 1 }
					} else {
						$_uniqueErrors[$_key].Count++
					}
				}
				Write-Host ""
				Write-Host "[RUNSPACE] $($_errStream.Count) non-terminating error(s) in child ($($_uniqueErrors.Count) unique):" -ForegroundColor DarkYellow
				foreach ($_entry in $_uniqueErrors.Values) {
					$_e = $_entry.Error
					$_c = $_entry.Count
					$_line = if ($_e.InvocationInfo) { $_e.InvocationInfo.ScriptLineNumber } else { '?' }
					Write-Host "  Line $_line (x$_c): $($_e.Exception.Message)" -ForegroundColor DarkYellow
				}
			}
		}
	} finally {
		if ($_closeHandlerRegistered) { try { [_mJigCloseHandlerX]::Unregister() } catch {} }
		# Stop + EndInvoke: waits for BatchInvocationWorkItem thread to fully exit
		if ($null -ne $_asyncResult -and -not $_asyncResult.IsCompleted) {
			try { $null = $_ps.Stop() } catch {}
		}
		if ($null -ne $_asyncResult) {
			try { $null = $_ps.EndInvoke($_asyncResult) } catch {}
		}
		$_ps.Dispose()
		$_rs.Close()
		$_rs.Dispose()

		# Restore process-global console state (VT100, mode flags, title)
		try { [Console]::Write("$([char]27)[?25h$([char]27)[?7h$([char]27)[0m") } catch {}
		if ($_k32Loaded) {
			try { $null = [_mJigProv._K32]::SetConsoleMode($_hIn,  $_savedInMode)  } catch {}
			try { $null = [_mJigProv._K32]::SetConsoleMode($_hOut, $_savedOutMode) } catch {}
		}
		if ($null -ne $_savedTitle) {
			try { $Host.UI.RawUI.WindowTitle = $_savedTitle } catch {}
		}
		# In debug mode, pause before clearing so any errors or init output can be read.
		if ($DebugMode) {
			Write-Host ""
			Write-Host "Press any key to exit..." -ForegroundColor Yellow
			try {
				$Host.UI.RawUI.FlushInputBuffer()
				while ($true) {
					Start-Sleep -Milliseconds 50
					if ($Host.UI.RawUI.KeyAvailable) {
						try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp,AllowCtrlC") } catch {}
						break
					}
				}
			} catch {}
		}
		try { [Console]::Clear() } catch {}
		if ($DebugMode) {
			Write-Host ""
			Write-Host "[RUNSPACE] Child exited - console state restored" -ForegroundColor Cyan
			if ($_k32Loaded) {
				$_finalIn = [uint32]0; $_finalOut = [uint32]0
				$null = [_mJigProv._K32]::GetConsoleMode($_hIn,  [ref]$_finalIn)
				$null = [_mJigProv._K32]::GetConsoleMode($_hOut, [ref]$_finalOut)
				Write-Host "  Console IN mode  : 0x$($_finalIn.ToString('X4'))"  -ForegroundColor Gray
				Write-Host "  Console OUT mode : 0x$($_finalOut.ToString('X4'))" -ForegroundColor Gray
			}
			Write-Host "  Window title     : $($Host.UI.RawUI.WindowTitle)" -ForegroundColor Gray
		}
	}
}
