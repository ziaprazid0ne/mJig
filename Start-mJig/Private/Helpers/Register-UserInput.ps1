	function Register-UserInput {
		param(
			[Parameter(Mandatory = $true)]
			[ValidateSet('Mouse', 'Keyboard', 'Scroll', 'Click')]
			[string]$Source,
			[Parameter(Mandatory = $true)]
			[datetime]$Date,
		$UserInputDetectedRef = $null,
		$MouseDetectedRef     = $null,
		$KeyboardDetectedRef  = $null
		)

		# Unified activity clock: single write point for auto-sleep and cooldown
		$script:LastUserActivityTime = $Date
		$script:_CooldownArmed       = $true

		if ($null -ne $UserInputDetectedRef) { $UserInputDetectedRef.Value = $true }

		switch ($Source) {
			'Mouse' {
				if ($null -ne $MouseDetectedRef) { $MouseDetectedRef.Value = $true }
				$script:LastMouseMovementTime = $Date
			}
			'Keyboard' {
				if ($null -ne $KeyboardDetectedRef) { $KeyboardDetectedRef.Value = $true }
			}
			'Scroll' {
				if ($null -ne $MouseDetectedRef) { $MouseDetectedRef.Value = $true }
			}
			'Click' {
				if ($null -ne $MouseDetectedRef) { $MouseDetectedRef.Value = $true }
			}
		}
	}
