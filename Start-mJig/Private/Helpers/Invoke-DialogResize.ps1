	function Invoke-DialogResize {
		param(
			[ref]$HostWidthRef,
			[ref]$HostHeightRef,
			[string]$ScreenState,
			[scriptblock]$ParentRedrawCallback = $null
		)
		$stableSize = Invoke-ResizeHandler -PreviousScreenState $ScreenState
		$HostWidthRef.Value  = $stableSize.Width
		$HostHeightRef.Value = $stableSize.Height
		Write-MainFrame -Force -NoFlush
		if ($null -ne $ParentRedrawCallback) {
			& $ParentRedrawCallback $stableSize.Width $stableSize.Height
		}
		return $stableSize
	}
