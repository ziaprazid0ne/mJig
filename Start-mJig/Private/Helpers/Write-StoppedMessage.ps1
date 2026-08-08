function Get-RuntimeString {
	param([datetime]$StartTime)
	$runtime = (Get-Date) - $StartTime
	$hours   = [math]::Floor($runtime.TotalHours)
	$minutes = $runtime.Minutes
	$seconds = $runtime.Seconds
	if ($hours -gt 0) {
		return "$hours hour$(if ($hours -ne 1) { 's' }), $minutes minute$(if ($minutes -ne 1) { 's' })"
	} elseif ($minutes -gt 0) {
		return "$minutes minute$(if ($minutes -ne 1) { 's' }), $seconds second$(if ($seconds -ne 1) { 's' })"
	} else {
		return "$seconds second$(if ($seconds -ne 1) { 's' })"
	}
}

function Write-StoppedMessage {
	param([datetime]$ScriptStartTime)
	Clear-Host
	$runtimeStr = Get-RuntimeString -StartTime $ScriptStartTime
	Write-Host ""
	$mouseEmoji = [char]::ConvertFromUtf32(0x1F400)
	Write-Host "  mJig(" -NoNewline -ForegroundColor $script:HeaderAppName
	$mouseEmojiX = $Host.UI.RawUI.CursorPosition.X
	$mouseEmojiY = $Host.UI.RawUI.CursorPosition.Y
	Write-Host $mouseEmoji -NoNewline -ForegroundColor $script:HeaderIcon
	[Console]::SetCursorPosition($mouseEmojiX + 2, $mouseEmojiY)
	Write-Host ") " -NoNewline -ForegroundColor $script:HeaderAppName
	Write-Host "Stopped" -ForegroundColor $script:TextError
	Write-Host ""
	Write-Host "  Runtime: " -NoNewline -ForegroundColor $script:TextMuted
	Write-Host $runtimeStr -ForegroundColor $script:TextDefault
	Write-Host ""
}
