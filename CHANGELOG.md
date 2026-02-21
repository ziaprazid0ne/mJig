# mJig Changelog

All notable changes to `start-mjig.ps1` are documented in this file.

---

## [Latest] - Unreleased

Changes since last commit (8014293 - "a bit broken but with a bunch of updates"):

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
- `Get-KeyName` - Standalone helper function (lines 1-80) for mapping VK codes to readable names
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
- `mJiggAPI.Keyboard` - Keyboard APIs (GetAsyncKeyState, keybd_event)
- `mJiggAPI.Mouse` - Mouse/Window APIs (GetCursorPos, SetCursorPos, FindWindow, EnumWindows, etc.)
- `mJiggAPI.MouseHook` - Mouse wheel hook (WH_MOUSE_LL)

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

## [8014293] - 2026-02-20

### Commit Message
"a bit broken but with a bunch of updates"

*This commit represents the baseline. Changes documented above in [Latest] are relative to this commit.*

---

## [06f12d6] - Previous

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
