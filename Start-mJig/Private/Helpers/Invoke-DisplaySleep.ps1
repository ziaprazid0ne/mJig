	function Invoke-DisplaySleep {
	param(
		[Parameter(Mandatory = $true)]
		[ValidateSet('Sleep', 'Wake')]
		[string]$Action
	)

		# Enumerates fresh physical monitor handles from all current HMONITORs.
		# Returns a List[IntPtr] of physical monitor handles. Caller must call
		# DestroyPhysicalMonitor on each handle when done.
		# Re-enumerating on every call is mandatory: DDC/CI handles are tied to the
		# I2C session, which is dropped when the monitor enters deep sleep.
		$getPhysHandles = {
			$handles = [System.Collections.Generic.List[IntPtr]]::new()
			$delegateType = ($script:_ApiNamespace + '.DisplayControl+EnumMonitorsDelegate') -as [type]
			$hMons = [System.Collections.Generic.List[IntPtr]]::new()
			$enumCb = { param($hm, $h, $r, $d); $null = $hMons.Add($hm); return $true } -as $delegateType
			$null = $script:DisplayAPI::EnumDisplayMonitors(
				[IntPtr]::Zero, [IntPtr]::Zero, $enumCb, [IntPtr]::Zero)
			foreach ($hMon in $hMons) {
				$physCount = [uint32]0
				if (-not $script:DisplayAPI::GetNumberOfPhysicalMonitorsFromHMONITOR(
						$hMon, [ref]$physCount) -or $physCount -eq 0) { continue }
				$physArr = [System.Array]::CreateInstance(
					($script:PhysicalMonitorType -as [type]), [int]$physCount)
				if ($script:DisplayAPI::GetPhysicalMonitorsFromHMONITOR($hMon, $physCount, $physArr)) {
					foreach ($pm in $physArr) { $null = $handles.Add($pm.hPhysicalMonitor) }
				}
			}
			return $handles
		}

		$sendVcpKnock = {
			$handles = & $getPhysHandles
			foreach ($h in $handles) {
				$null = $script:DisplayAPI::SetVCPFeature($h, [byte]0xD6, [uint32]1)
				$null = $script:DisplayAPI::DestroyPhysicalMonitor($h)
			}
		}

		$readVcpOn = {
			$anyOn = $false
			$handles = & $getPhysHandles
			foreach ($h in $handles) {
				$vcpType    = [uint32]0
				$currentVal = [uint32]0
				$maxVal     = [uint32]0
				if ($script:DisplayAPI::GetVCPFeatureAndVCPFeatureReply(
						$h, [byte]0xD6, [ref]$vcpType, [ref]$currentVal, [ref]$maxVal)) {
					if ($currentVal -eq 1) { $anyOn = $true }
				}
				$null = $script:DisplayAPI::DestroyPhysicalMonitor($h)
			}
			return $anyOn
		}

		if ($Action -eq 'Sleep') {
			# VCP 0xD6 = 2 (Standby), not 4 (Off/DPM). Value 4 often powers down the
			# monitor's DDC/CI I2C channel after a short time, so SetVCPFeature(On=1)
			# can never reach firmware and a physical power-cycle is required.
			# Standby blanks the panel while keeping DDC reachable for wake.
			$handles = & $getPhysHandles
			foreach ($h in $handles) {
				$null = $script:DisplayAPI::SetVCPFeature($h, [byte]0xD6, [uint32]2)
				$null = $script:DisplayAPI::DestroyPhysicalMonitor($h)
			}
			if ($script:DisplaySleepAudioEnabled) {
				[Console]::Beep(660, 150)
				Start-Sleep -Milliseconds 80
				[Console]::Beep(440, 200)
			}
			return $true
		}

	# Wake: multi-attempt VCP knock + verify.
		# After minutes in DDC soft-off (VCP 0xD6=4), I2C is often dormant and a single
		# SetVCPFeature may return true at the API level without reaching firmware.
		# Do NOT use SC_MONITORPOWER=2 here: that puts Windows into "monitors off", after
		# which mJig's own simulated keypress is treated as user input and spontaneously
		# wakes the display. Only broadcast SC_MONITORPOWER=-1 (on) as a nudge.
		$waitMsList = @(700, 1200, 2000, 3000)
		$anyOk = $false
		for ($attempt = 0; $attempt -lt $waitMsList.Count; $attempt++) {
			$null = $script:MouseAPI::PostMessage(
				[IntPtr]::new(0xFFFF), [uint32]0x0112,
				[IntPtr]::new(0xF170), [IntPtr]::new(-1))

			& $sendVcpKnock
			Start-Sleep -Milliseconds $waitMsList[$attempt]

			if ([bool](& $readVcpOn)) {
				$anyOk = $true
				break
			}
		}

		if ($script:DisplaySleepAudioEnabled) {
			if ($anyOk) {
				[Console]::Beep(440, 200)
				Start-Sleep -Milliseconds 80
				[Console]::Beep(660, 150)
			} else {
				[Console]::Beep(440, 100)
			}
		}

		return $anyOk
	}
