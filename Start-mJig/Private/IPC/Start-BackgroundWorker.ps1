function Start-BackgroundWorker {
	param(
		[System.Collections.IDictionary]$BoundParameters,
		[ref]$InlineRef,
		[switch]$Headless,
		[switch]$Diag,
		[string]$ModuleRoot
	)

	$modulePath = Join-Path $ModuleRoot 'Start-mJig.psm1'
	$workerCmd = "Import-Module '$modulePath'; Start-mJig -_WorkerMode -_InModuleRunspace -_PipeName '$($script:PipeName)'"

	# Forward movement/timing parameters to the worker
	if ($BoundParameters.ContainsKey('IntervalSeconds'))        { $workerCmd += " -IntervalSeconds $($BoundParameters['IntervalSeconds'])" }
	if ($BoundParameters.ContainsKey('IntervalVariance'))       { $workerCmd += " -IntervalVariance $($BoundParameters['IntervalVariance'])" }
	if ($BoundParameters.ContainsKey('MoveSpeed'))              { $workerCmd += " -MoveSpeed $($BoundParameters['MoveSpeed'])" }
	if ($BoundParameters.ContainsKey('MoveVariance'))           { $workerCmd += " -MoveVariance $($BoundParameters['MoveVariance'])" }
	if ($BoundParameters.ContainsKey('TravelDistance'))         { $workerCmd += " -TravelDistance $($BoundParameters['TravelDistance'])" }
	if ($BoundParameters.ContainsKey('TravelVariance'))         { $workerCmd += " -TravelVariance $($BoundParameters['TravelVariance'])" }
	if ($BoundParameters.ContainsKey('AutoResumeDelaySeconds')) { $workerCmd += " -AutoResumeDelaySeconds $($BoundParameters['AutoResumeDelaySeconds'])" }
	if ($BoundParameters.ContainsKey('EndTime'))                { $workerCmd += " -EndTime '$($BoundParameters['EndTime'])'" }
	if ($BoundParameters.ContainsKey('EndVariance'))            { $workerCmd += " -EndVariance $($BoundParameters['EndVariance'])" }
	if ($BoundParameters.ContainsKey('Output'))                 { $workerCmd += " -Output '$($BoundParameters['Output'])'" }
	if ($BoundParameters.ContainsKey('Title') -and ($BoundParameters['Title']).Length -gt 0) { $workerCmd += " -Title '$($BoundParameters['Title'])'" }
	if ($Diag) { $workerCmd += " -Diag" }

	$powershellExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName

	# Release mutex before spawning — worker acquires its own
	if ($null -ne $script:InstanceMutex) {
		try { $script:InstanceMutex.ReleaseMutex() } catch {}
		$script:InstanceMutex.Dispose()
		$script:InstanceMutex = $null
	}

	# WMI spawn: detaches from terminal job object so worker survives viewer close
	try {
		$encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($workerCmd))
		$cmdLine = "`"$powershellExe`" -NoProfile -NoLogo -WindowStyle Hidden -EncodedCommand $encodedCmd"
		$cimResult = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $cmdLine }
		if ($cimResult.ReturnValue -ne 0) { throw "WMI return code $($cimResult.ReturnValue)" }
		$null = Get-Process -Id $cimResult.ProcessId -ErrorAction Stop
	} catch {
		try {
			$encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($workerCmd))
			$workerArgs = @('-NoProfile', '-NoLogo', '-WindowStyle', 'Hidden', '-EncodedCommand', $encodedCmd)
			$null = Start-Process -FilePath $powershellExe -ArgumentList $workerArgs -WindowStyle Hidden -PassThru
		} catch {
			Write-Host "WARNING: Could not start background process. Running in direct mode." -ForegroundColor $script:TextWarning
			Write-Host "  $($_.Exception.Message)" -ForegroundColor Gray
			try {
				$script:InstanceMutex = New-Object System.Threading.Mutex($false, "Global\$($script:SessionId.PipeName)")
				$mutexAcq = $script:InstanceMutex.WaitOne(0)
				$null = $mutexAcq
			} catch {}
			$InlineRef.Value = $true
			return @{ ShouldReturn = $false; IsViewerMode = $false }
		}
	}

	# Headless: worker spawned, exit immediately (no TUI)
	if ($Headless) {
		return @{ ShouldReturn = $true; IsViewerMode = $false }
	}

	if (-not $InlineRef.Value) {
		$pipeRes = Connect-WorkerPipe -PipeName $script:PipeName -ConnectTimeoutMs 15000
		if ($null -eq $pipeRes) {
			if ($script:DiagEnabled) { Show-DiagnosticFiles }
			return @{ ShouldReturn = $true; IsViewerMode = $false }
		}
		return @{
			ShouldReturn  = $false
			IsViewerMode  = $true
			PipeResult    = $pipeRes
			PipeClient    = $pipeRes.Client
			PipeReader    = $pipeRes.Reader
			PipeWriter    = $pipeRes.Writer
		}
	}

	return @{ ShouldReturn = $false; IsViewerMode = $false }
}
