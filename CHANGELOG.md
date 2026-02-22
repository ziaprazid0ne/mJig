# mJig Changelog

All notable changes to `start-mjig.ps1` are documented in this file.

---

## [Latest] - Unreleased

Changes since last commit (4ddbfc2 - "Evidence-based input detection, security hardening, scroll/mouse hook removal"):

### Added
- `$script:MenuClickHotkey` - Script-scoped variable for menu click hotkey (was previously used but never initialized)
- `$script:ConsoleClickCoords` - Script-scoped variable storing character cell X/Y from the last left-click detected via `PeekConsoleInput` MOUSE_EVENT records
- **Mouse click detection via PeekConsoleInput** - Main loop, Time dialog, Movement dialog, and Quit dialog all detect left-button clicks by reading native `MOUSE_EVENT` records from the console input buffer. The console provides exact character cell coordinates, eliminating all pixel-to-character math.
- **Movement dialog field click selection** - Clicking on a field row in the Modify Movement dialog now switches the active field to the clicked row, redraws only the affected rows (no flicker), and positions the cursor at the end of that field's input
- **Quit dialog Yes/No button click support** - The quit confirmation dialog now responds to mouse clicks on the Yes and No buttons
- **Buffered frame rendering** - `Write-Buffer`, `Flush-Buffer`, `Clear-Buffer` functions and `$script:RenderQueue` (`System.Collections.Generic.List[hashtable]`). All UI rendering now writes to an in-memory queue first, then flushes to the console in a single burst per frame.
- **VT100/ANSI single-write rendering** - `Flush-Buffer` builds a single string with embedded VT100 escape sequences (cursor positioning, color changes, cursor visibility) via `StringBuilder`, then outputs the entire frame with one `[Console]::Write()` call. Eliminates 55-80 separate `Write-Host` calls per frame.
- **VT100 processing enabled on stdout** - `ENABLE_VIRTUAL_TERMINAL_PROCESSING` (0x0004) enabled on the output handle via `SetConsoleMode` during console setup
- **UTF-8 console encoding** - `[Console]::OutputEncoding` set to `[System.Text.Encoding]::UTF8` for correct emoji rendering via `[Console]::Write()`
- **ConsoleColor-to-ANSI mapping tables** - `$script:AnsiFG` and `$script:AnsiBG` hashtables mapping all 16 `[ConsoleColor]` values to ANSI SGR codes (30-37/90-97 for FG, 40-47/100-107 for BG)
- **`$script:CursorVisible` tracking variable** - Tracks cursor visibility state for conditional cursor show/hide in VT100 sequences
- **`$script:ESC` variable** - Stores `[char]27` for VT100 escape sequence construction
- **`-ClearFirst` switch on `Flush-Buffer`** - Embeds `ESC[2J` (clear screen) in the VT100 string for atomic clear+redraw with no visible flash
- **`-ClearFirst` switch on `Draw-ResizeLogo`** - Passes through to `Flush-Buffer -ClearFirst`
- **Hidden view `(h)` button** - Clickable button in bottom-right corner of hidden mode, behaves like the hide_output toggle

### Changed
- **All rendering paths use buffered output** - Header, separators, log entries, stats box, menu bar, all three dialogs (Time, Movement, Quit), dialog shadows, field redraws, resize logo, and dialog cleanup all write through `Write-Buffer` instead of directly to the console. `Flush-Buffer` is called at well-defined frame boundaries.
- **Emoji positions computed statically** - Menu bar, header, and dialog button rows no longer read `$Host.UI.RawUI.CursorPosition.X` after writing emojis. Instead, positions are calculated mathematically (emoji = 2 display cells), and menu item bounds (`$script:MenuItemsBounds`) are tracked by arithmetic rather than cursor position reads.
- **Click detection architecture** - Replaced the previous `GetAsyncKeyState` + `Get-MousePosition` + screen-pixel-to-character-cell conversion approach with native console input buffer events. `PeekConsoleInput` reads `MOUSE_EVENT` records which provide exact character cell coordinates directly from the console.
- **Left-button guard added** - Click-to-button mapping in the main loop only processes left button presses (dwButtonState & 0x0001), preventing accidental menu triggers from other mouse buttons.
- **Debug mode gate removed** - Click detection logic now runs in all modes, not just when `$DebugMode` is enabled. Debug logging remains conditional.
- **Simplified hit-testing** - All click detection now uses simple character cell comparisons. No pixel math, tolerance percentages, or expanded bounding boxes.
- **Cursor visibility via VT100** - All `[Console]::CursorVisible` assignments replaced with VT100 `ESC[?25l`/`ESC[?25h` sequences and `$script:CursorVisible` tracking. `Flush-Buffer` conditionally shows cursor at end of frame based on tracked state.
- **Default colors use ANSI 39/49** - Segments without explicit FG/BG colors emit ANSI "default foreground" (39) and "default background" (49) codes instead of mapping console's current color, preserving the terminal's true default/transparent background.
- **Atomic screen clear+redraw** - All `Clear-Host` + rendering pairs replaced with `Flush-Buffer -ClearFirst`, embedding the screen clear in the VT100 string so clear and redraw happen in a single `[Console]::Write()` call with no visible blank flash. Affects resize logo, dialog resize, and hidden view.
- **Hidden view clears old menu bounds** - `$script:MenuItemsBounds` cleared when entering hidden mode so old buttons are not clickable.

### Removed
- `Convert-ScreenToConsole` function - Screen pixel to console cell conversion (replaced by PeekConsoleInput native coords)
- `Test-ClickInBounds` function - Pixel-based hit-test with tolerance expansion (replaced by exact cell matching)
- `Get-ConsoleWindowHandle` function - Cached window handle lookup (no longer needed without pixel conversion)
- `$script:CachedConsoleHandle`, `$script:LastClickDebug`, `$script:ClickTolerancePct` variables
- `RECT` struct, `ScreenToClient`, `GetWindowRect`, `GetClientRect` P/Invoke declarations
- `CONSOLE_FONT_INFOEX` struct, `GetCurrentConsoleFontEx` P/Invoke declaration
- `ReadConsoleOutputAttribute` and `WriteConsoleOutputAttribute` P/Invoke declarations (unused remnants from previous emoji background fix attempt)
- All `[Console]::CursorVisible` references (replaced by VT100 sequences)
- Segment merging optimization in `Flush-Buffer` (no longer needed -- VT100 color changes are just bytes in the string, not separate API calls)
- Separate `Clear-Host` calls before rendering (replaced by `-ClearFirst` flag)

### Fixed
- **Clickable menu buttons not working** - The entire click-to-button coordinate conversion pipeline was gated behind `if ($DebugMode)`, making click detection silently fail in normal operation. Replaced with always-on PeekConsoleInput approach.
- **Undefined `$gotCursorPos` variable** - The main loop checked `if ($gotCursorPos)` which was never defined, causing coordinate conversion to always be skipped.
- **`$mousePoint` vs `$mousePos` variable mismatch** - Code retrieved mouse position into `$mousePos` but attempted to read from `$mousePoint`, which was never populated.
- **Quit dialog click detection** - `$keyProcessed` and `$char` were reset immediately after being set by `$script:DialogButtonClick`, wiping out the click result.
- **Click detection broke after Modify Movement dialog** - Added `$script:DialogButtonBounds = $null` cleanup on dialog close.
- **Window resize triggered quit dialog** - Resolved by switching to PeekConsoleInput, which only fires MOUSE_EVENT records when the console is focused.
- **Button bounds tightened to visible characters** - Click areas now correspond exactly to rendered emoji+pipe+text characters.
- **UI strobing/flicker** - Replaced 55-80 `Write-Host` calls per frame with single VT100 `[Console]::Write()` call, reducing frame render time from 55-160ms to sub-millisecond.
- **Grey background on default areas** - VT100 renderer now uses ANSI code 49 (default background) instead of mapping `[Console]::BackgroundColor` to an explicit color.
- **Emoji background on 2-column emoji** - `Write-Buffer -Wide` appends a trailing space with the background color for wide emojis; the explicit pipe positioning overwrites the space.
- **Hidden view resize crash** - `SetCursorPosition` out-of-bounds during resize replaced by VT100 positioning (no exception possible).
- **Hidden view strobing** - Full-frame clear+redraw now atomic via `Flush-Buffer -ClearFirst`.

---

## [4ddbfc2] - 2026-02-21

### Commit Message
"Evidence-based input detection, security hardening, scroll/mouse hook removal"

Changes since commit 3f27144 ("still working towords initial release"):

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
