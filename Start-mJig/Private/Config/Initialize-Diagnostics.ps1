	$script:DiagEnabled = $Diag
	if ($script:DiagEnabled) {
		# $_diagRoot is set by Start-mJig.psm1 before dot-sourcing this file,
		# ensuring the module root is used (not this file's Private\Config\ directory).
		$script:DiagFolder = Join-Path $_diagRoot "_diag"
		if (-not (Test-Path $script:DiagFolder)) {
			New-Item -ItemType Directory -Path $script:DiagFolder -Force | Out-Null
		}
		$script:StartupDiagFile = Join-Path $script:DiagFolder "startup.txt"
		$script:SettleDiagFile  = Join-Path $script:DiagFolder "settle.txt"
		$script:InputDiagFile   = Join-Path $script:DiagFolder "input.txt"
		$script:IpcDiagFile     = Join-Path $script:DiagFolder "ipc.txt"
		$script:NotifyDiagFile  = Join-Path $script:DiagFolder "notify.txt"

		$diagTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
		$diagRsId = if ([System.Management.Automation.Runspaces.Runspace]::DefaultRunspace) {
			[System.Management.Automation.Runspaces.Runspace]::DefaultRunspace.Id } else { '(unknown)' }
		"=== mJig Startup Diag: $diagTimestamp ===" | Out-File $script:StartupDiagFile
		"$(Get-Date -Format 'HH:mm:ss.fff') - CHECKPOINT 1: Starting initialization" | Out-File $script:StartupDiagFile -Append
		"  Diag enabled, folder: $script:DiagFolder" | Out-File $script:StartupDiagFile -Append
		"  Runspace ID: $diagRsId  Thread: $([System.Threading.Thread]::CurrentThread.ManagedThreadId)  Modules: $((Get-Module | ForEach-Object { $_.Name }) -join ', ')" | Out-File $script:StartupDiagFile -Append
		"=== mJig Notify Diag: $diagTimestamp ===" | Out-File $script:NotifyDiagFile
		"$(Get-Date -Format 'HH:mm:ss.fff') [INIT] Notification diagnostics started  Mode=$(if ($_WorkerMode) { 'worker' } else { 'viewer' })" | Out-File $script:NotifyDiagFile -Append
		"=== mJig Settle Diag: $diagTimestamp ===" | Out-File $script:SettleDiagFile
		"$(Get-Date -Format 'HH:mm:ss.fff') - Settle diagnostics started" | Out-File $script:SettleDiagFile -Append
		"=== mJig Input Diag: $diagTimestamp ===" | Out-File $script:InputDiagFile
		"$(Get-Date -Format 'HH:mm:ss.fff') - Input diagnostics started (PeekConsoleInput + GetLastInputInfo)" | Out-File $script:InputDiagFile -Append
	}
