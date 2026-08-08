	function Send-ViewerMessage {
		param([hashtable]$Message)
		# Guard + error-swallow wrapper for all viewer pipe sends in the main loop.
		# $_isViewerMode and $_viewerPipeWriter are resolved via scope chain (defined inside Start-mJig).
		if ($_isViewerMode) { try { Send-PipeMessage -Writer $_viewerPipeWriter -Message $Message } catch {} }
	}
