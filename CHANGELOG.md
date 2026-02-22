# mJig Changelog

All notable changes to `start-mjig.ps1` are documented in this file.

---

## [Latest] - Unreleased

Changes since last commit (3f27144 - "still working towords initial release"):

### Added
- `mJiggAPI.LASTINPUTINFO` struct, `GetLastInputInfo`, and `GetTickCount64` P/Invoke for passive system-wide input detection
- `mJiggAPI.KEY_EVENT_RECORD` struct for reading keyboard events from the console input buffer
- `KEY_EVENT_RECORD` overlay added to `INPUT_RECORD` union (at FieldOffset 4, alongside MouseEvent)
- `PeekConsoleInput`-based keyboard event detection -- reads KEY_EVENT records (EventType 0x0001) from the console input buffer to provide evidence-based keyboard detection without scanning key codes
- `PeekConsoleInput`-based scroll wheel detection -- reads MOUSE_EVENT records with scroll flag (EventType 0x0002, dwEventFlags 0x0004)
- Mouse movement inference via `GetLastInputInfo` -- when system activity is detected but no keyboard, scroll, or click evidence exists, it is classified as mouse movement
- Console input buffer flush after simulated keypress to prevent stale Right Alt events from being detected as user keyboard input
- `_diag/input.txt` diagnostic log for PeekConsoleInput + GetLastInputInfo detection
- `.gitignore` to exclude `_diag/` folder and backup files from git

### Changed
- **Input Detection Architecture**: Complete overhaul of input classification. All input types are now evidence-based:
  - **Keyboard**: Detected via `PeekConsoleInput` KEY_EVENT records (filtered for simulated VK 0xA5). Only peeked, not consumed, so menu hotkey handler can still read them.
  - **Scroll**: Detected via `PeekConsoleInput` MOUSE_EVENT with scroll flag. Consumed to prevent buffer buildup.
  - **Mouse clicks**: Detected via `GetAsyncKeyState` VK 0x01-0x06 (unchanged).
  - **Mouse movement**: Detected via `Test-MouseMoved` position polling or inferred by `GetLastInputInfo` when no other input type explains the activity.
  - **`GetLastInputInfo`**: No longer infers "keyboard" -- only sets `$script:userInputDetected` and infers mouse movement by elimination.
- **Mouse movement display label**: Changed from emoji (🐀) to text "Mouse" in the detected inputs display
- **Scroll/Input Detection**: Replaced system-wide low-level mouse hook (`WH_MOUSE_LL` / `mJiggAPI.MouseHook`) with `PeekConsoleInput` + `GetLastInputInfo`-based detection. Uses `GetTickCount64` for 64-bit tick math to avoid overflow on systems with >24.9 days uptime.
- **Diagnostics folder**: Moved from `$env:TEMP\mjig_diag\` to `_diag/` relative to the script location, so agents and users can access logs directly in the project directory
- Resize quote color changed from `DarkGray` to `White`

### Removed
- `mJiggAPI.MouseHook` class and all associated P/Invoke definitions (`SetWindowsHookEx`, `UnhookWindowsHookEx`, `CallNextHookEx`, `GetModuleHandle`, `PeekMessage`, `MSG`, `MSLLHOOKSTRUCT`, `LowLevelMouseProc`)
- `$PreviousMouseWheelDelta` tracking variable (no longer needed)
- Mouse hook install/uninstall/ProcessMessages calls from initialization, debug pause, and main loop
- Keyboard inference from `GetLastInputInfo` -- the flawed "if not mouse, must be keyboard" logic has been removed entirely

### Fixed
- Mouse cursor lag during debug "press any key" pause and potentially during busy main loop sections, caused by `WH_MOUSE_LL` hook starving the system message pump
- False "Keyboard" labels appearing when only mouse movement or scroll wheel was used -- caused by `GetLastInputInfo` incorrectly defaulting to keyboard when `Test-MouseMoved` had brief polling gaps
- Menu hotkeys not responding -- `PeekConsoleInput` was consuming KEY_EVENT records from the buffer before the menu hotkey handler (`$Host.UI.RawUI.ReadKey`) could read them. Fixed by only peeking (not consuming) keyboard events.

### Security
- **Removed `Get-KeyName` function** -- eliminated VK-code-to-name mapping table that security scanners flag as keylogger pattern
- **Removed full 256-code `GetAsyncKeyState` keyboard scan** -- the `for ($keyCode = 0; $keyCode -le 255; ...)` loop that polled every virtual key code every ~50ms has been replaced with a focused mouse-button-only loop (VK 0x01-0x06)
- **Removed `GetAsyncKeyState` from `Keyboard` class** -- the API is now only exposed in the `Mouse` class, reducing the P/Invoke surface
- **Removed `PressedKeys` real-time display scan** -- the secondary 256-code scan that populated real-time key state for the stats box has been removed entirely
- **Keyboard detection now evidence-based** -- `$keyboardInputDetected` is set only when actual KEY_EVENT records are found in the console input buffer via `PeekConsoleInput`. No key identity is captured beyond filtering the simulated VK 0xA5.
- **Stats box shows categories, not key names** -- "Detected Inputs" now displays `Mouse`, `Keyboard`, `LButton`, `Scroll/Other` etc. instead of specific key names like `A, LShift, Space`
- **Removed `$PressedKeys` and `$intervalKeys`** -- variables that accumulated specific key names are removed; only boolean flags and category labels remain

---

## [3f27144] - 2026-02-21

### Commit Message
"still working towords initial release, there is now a changelog.md for better tracking of changes across commits. and a context.md for quicker training of agents."

Changes since commit 8014293 ("a bit broken but with a bunch of updates"):

### Added

#### New Parameters
- `-DebugMode` switch - Enables verbose logging during initialization and runtime
- `-Diag` switch - Enables file-based diagnostics to `$env:TEMP\mjig_diag\`
- `-EndVariance` (int) - Random variance in minutes for end time
- `-IntervalSeconds` (double) - Base interval between movement cycles (was hardcoded)
- `-IntervalVariance` (double) - Random variance for intervals (was hardcoded)
- `-MoveSpeed` (double) - Movement animation duration in seconds
- `-MoveVariance` (double) - Random variance for movement speed
- `-TravelDistance` (double) - Cursor travel distance in pixels
- `-TravelVariance` (double) - Random variance for travel distance
- `-AutoResumeDelaySeconds` (double) - Cooldown timer after user input

#### New Functions
- `Get-KeyName` - Standalone helper function for mapping VK codes to readable names *(removed in latest -- see Security section)*
- `Find-WindowHandle` - Window handle lookup using EnumWindows
- `Get-Padding` - Calculate padding for dialog layouts
- `Get-TimeSinceMs` - Calculate milliseconds elapsed since a timestamp
- `Get-ValueWithVariance` - Generate random values with variance
- `Get-MousePosition` - Wrapper for GetCursorPos API
- `Test-MouseMoved` - Check if mouse has moved beyond threshold
- `Draw-DialogShadow` / `Clear-DialogShadow` - Dialog drop shadow rendering
- `Write-SimpleDialogRow` / `Write-SimpleFieldRow` - Dialog row rendering helpers
- `Show-MovementModifyDialog` - Runtime movement settings modification
- `Show-QuitConfirmationDialog` - Quit confirmation with runtime stats
- `Show-TimeChangeDialog` - Runtime end time modification
- `Draw-ResizeLogo` - Centered logo during window resize

#### New Features
- **Theme System**: Centralized color variables (`$script:MenuButton*`, `$script:Header*`, `$script:StatsBox*`, `$script:QuitDialog*`, `$script:TimeDialog*`, `$script:MoveDialog*`, `$script:Resize*`, `$script:Text*`)
- **Box-Drawing Characters**: All box chars now use `[char]` casts to avoid encoding issues
- **Duplicate Instance Detection**: Prevents running multiple mJig instances
- **Mouse Click Support**: Menu items and dialog buttons are clickable
- **Menu Item Bounds Tracking**: `$script:MenuItemsBounds` for click detection
- **Stats Box**: Real-time display of detected keyboard/mouse inputs (full view)
- **Window Resize Handling**: Clears screen, shows centered logo with decorative box
- **Resize Quotes**: Random playful quotes displayed during resize (`$script:ResizeQuotes`)
- **Mouse Stutter Prevention**: Waits for mouse to "settle" before starting next movement cycle
- **Movement Animation**: Smooth cursor movement with configurable speed
- **Auto-Resume Delay**: Configurable cooldown after user input before resuming automation
- **Diagnostics System**: File-based logging to `$env:TEMP\mjig_diag\` with `-Diag` flag

#### New P/Invoke Types
- `mJiggAPI.POINT` - Coordinate struct
- `mJiggAPI.RECT` - Rectangle struct
- `mJiggAPI.COORD` - Console coordinate struct
- `mJiggAPI.CONSOLE_SCREEN_BUFFER_INFO` - Console buffer info
- `mJiggAPI.MOUSE_EVENT_RECORD` - Mouse event data
- `mJiggAPI.INPUT_RECORD` - Input record union
- `mJiggAPI.SMALL_RECT` - Small rectangle struct
- `mJiggAPI.Keyboard` - Keyboard APIs (GetAsyncKeyState, keybd_event) *(GetAsyncKeyState later moved to Mouse class -- see latest)*
- `mJiggAPI.Mouse` - Mouse/Window APIs (GetCursorPos, SetCursorPos, FindWindow, EnumWindows, etc.)
- `mJiggAPI.MouseHook` - Mouse wheel hook (WH_MOUSE_LL)

#### New Documentation
- `CHANGELOG.md` - Structured change tracking across commits
- `CONTEXT.md` - Deep codebase context for AI agents
- `README.md` - Rewritten with full parameter table, usage examples, and architecture notes

### Changed

#### Parameters
- `-endTime` renamed to `-EndTime`, default changed from `"2400"` to `"0"` (no end time)
- `-Output` default changed from `"full"` to `"min"`
- Added `ValidateSet` for `-Output`: `"min"`, `"full"`, `"hidden"`, `"dib"`

#### Code Structure
- Moved from simple inline code to modular helper functions
- Parameters now copied to `$script:` variables for runtime modification
- Configuration section removed (settings now exposed as parameters)
- Variable naming standardized to PascalCase (`$lastPos` → `$LastPos`)
- P/Invoke types moved to `mJiggAPI` namespace with full struct definitions

#### UI
- Header now shows app icon (🐀) and view mode indicator
- Menu bar with emoji icons and colored hotkeys
- Interactive dialogs with drop shadows and themed colors
- Log entries use component-based structure for dynamic truncation

### Removed
- "Ideas & Notes" comment block
- Hardcoded configuration variables (`$defualtEndTime`, `$defualtEndMaxVariance`, `$intervalSeconds`, `$intervalVariance`)
- Simple `Keyboard` class (replaced by `mJiggAPI.Keyboard`)
- Direct cursor position access (replaced by helper functions)

### Fixed
- Encoding issues with box-drawing characters (now use `[char]` casts)
- Mouse stutter when movement cycle starts while user is moving mouse
- Console buffer size issues during window resize

---

## [8014293] - 2026-02-10

### Commit Message
"a bit broken but with a bunch of updates"

*This commit represents the baseline. Changes documented in [3f27144] above are relative to this commit.*

---

## [06f12d6] - 2026-01-22

### Commit Message
"major feature changes and bug fixes"

*Details not available - this was the state before 8014293.*

---

## Format Guidelines

When adding to this changelog:

1. **Latest Section**: Always add new changes under `[Latest] - Unreleased`
2. **On Commit**: Move `[Latest]` content to a new versioned section with commit hash and date
3. **Categories**: Use Added, Changed, Removed, Fixed, Deprecated, Security
4. **Be Specific**: Include function names, parameter names, line numbers where helpful
5. **Group Related**: Group related changes under descriptive subheadings
