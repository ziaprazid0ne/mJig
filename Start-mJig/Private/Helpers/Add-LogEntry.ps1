	function Add-LogEntry {
		param(
			[System.Collections.Generic.List[object]]$LogArray,
			[int]$Rows,
			[datetime]$Date,
			[Parameter(Mandatory = $true)]
			[string]$Text,
			[string]$ShortText = ""
		)
		if ([string]::IsNullOrEmpty($ShortText)) { $ShortText = $Text }
		$_ts = $Date.ToString("HH:mm:ss")
		if ($LogArray.Count -gt 0 -and $LogArray.Count -ge $Rows) { $LogArray.RemoveAt(0) }
		$null = $LogArray.Add([PSCustomObject]@{
			logRow = $true
			components = @(
				@{ priority = 1; text = $_ts; shortText = $_ts },
				@{ priority = 2; text = $Text; shortText = $ShortText }
			)
		})
	}
