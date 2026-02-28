# mJig Code Context for AI Agents

This document provides deep context for AI agents working on the `start-mjig.ps1` codebase.

> **IMPORTANT FOR AI AGENTS**: When modifying `start-mjig.ps1`, you must also update this `AGENTS.md` file and `README.md` to reflect any changes. This includes:
> - New or modified parameters
> - New or renamed functions
> - Changes to line number ranges in the code structure
> - New features or behaviors
> - New theme colors or UI components
> - Changes to hotkeys or interactive controls
> - New gotchas or patterns discovered during development
>
> Keeping documentation in sync prevents knowledge drift and ensures future AI agents have accurate context.

> **COMMIT WORKFLOW**: When the user says they are ready to commit (e.g. "I have everything staged", "please commit"), follow this process:
> 1. Run `git status` and `git log --oneline -5` to see staged files and recent commit style.
> 2. Write a commit message: first line is a concise summary of the biggest feature changes. Body is a bulleted list of key changes. End with "See CHANGELOG.md for full details."
> 3. Write the message to a temp file (`_commit_msg.txt`) and use `git commit -F _commit_msg.txt` (PowerShell does not support heredoc in git commands). Delete the temp file after.
> 4. Verify with `git status` that the working tree is clean.
> 5. **After the commit is confirmed**, prep `CHANGELOG.md` for the next commit:
>    - Move the current `[Latest] - Unreleased` content into a new versioned section with the commit hash and today's date.
>    - Replace `[Latest] - Unreleased` with a fresh empty section referencing the new commit.
>    - Do NOT stage or commit this changelog prep -- it becomes the starting point for the next round of changes.
>
> The user expects this full workflow every time. Do not skip the changelog prep step.

---

## Architecture Overview

The script is a single-file PowerShell application (~6,000 lines) implementing a console-based TUI mouse jiggler. It uses Win32 API calls via P/Invoke for low-level mouse/keyboard interaction.

### High-Level Flow

```
1. Load assemblies (System.Windows.Forms)
2. Define P/Invoke types (mJiggAPI namespace)
3. Initialize variables and theme colors
4. Define helper functions (including Invoke-ResizeHandler)
5. Parse and validate parameters
6. Show-StartupScreen  (skipped in -DebugMode / -Output hidden)
7. Show-StartupComplete (skipped in -DebugMode / -Output hidden)
   └── Calls Invoke-ResizeHandler if window is resized during the screen
8. Initialize $oldWindowSize / $OldBufferSize to current state
9. Enter main processing loop
   ├── Wait for interval (with input monitoring)
   ├── Check for user input / hotkeys
   ├── Detect window resize → Invoke-ResizeHandler (any context)
   ├── Wait for mouse to settle (stutter prevention)
   ├── Perform automated mouse movement
   ├── Send simulated keypress
   ├── Render UI (header, logs, stats, menu)
   └── Handle window resize
10. Cleanup on exit
```

### Code Structure Map

```
start-mjig.ps1
└── Start-mJig function (lines 1-end)
    │
    ├── ASCII Art Banner (lines 88-120)
    │   └── Decorative mouse ASCII art in comment block
    │
    ├── Parameters (lines 122-148)
    │   ├── $Output - View mode (min/full/hidden/dib)
    │   ├── $DebugMode - Verbose logging switch
    │   ├── $Diag - File-based diagnostics switch
    │   ├── $EndTime - Stop time in HHmm format
    │   ├── $EndVariance - Random variance for end time
    │   ├── $IntervalSeconds - Base interval between cycles
    │   ├── $IntervalVariance - Random variance for intervals
    │   ├── $MoveSpeed - Movement animation duration
    │   ├── $MoveVariance - Random variance for speed
    │   ├── $TravelDistance - Cursor travel distance in pixels
    │   ├── $TravelVariance - Random variance for distance
    │   └── $AutoResumeDelaySeconds - Cooldown after user input
    │
    ├── Initialization Variables (lines 150-212)
    │   ├── Script-scoped copies of parameters
    │   ├── State tracking variables
    │   ├── Resize handling variables
    │   └── Box-drawing character definitions
    │
    ├── Theme Colors Section (lines 214-289)
    │   ├── Menu bar colors (incl. OnClick pressed-state colors)
    │   ├── Header colors
    │   ├── Stats box colors
    │   ├── Dialog colors (Quit, Time, Movement)
    │   ├── Resize screen colors
    │   └── General UI colors
    │
    ├── Startup Screen Functions (lines ~220-367)
    │   ├── Show-StartupScreen  — initial "Initializing…" screen (Write-Host, pre-VT100)
    │   └── Show-StartupComplete — "Initialization Complete" box; keypress-wait or 7-s countdown;
    │       │                       nested helpers: getSize, drainWakeKeys, handleResize
    │       │                       getSize calls PeekConsoleInput before reading WindowSize (ConPTY flush)
    │       │                       handleResize is self-contained; does NOT call Invoke-ResizeHandler
    │
    ├── Invoke-ResizeHandler (lines ~383-425)
    │   └── Unified blocking resize handler for main loop and hidden-mode contexts
    │
    ├── Helper Functions (lines ~384-3700)
    │   ├── Find-WindowHandle (~384-470)
    │   ├── Buffered Rendering (Write-Buffer, Flush-Buffer, Clear-Buffer, Write-ButtonImmediate) (~1663-1780)
    │   ├── Draw-DialogShadow / Clear-DialogShadow (~1785-1830)
    │   ├── Show-TimeChangeDialog (~1835-2410)
    │   ├── Draw-ResizeLogo (~2495-2610)
    │   ├── Get-MousePosition (~2615-2630)
    │   ├── Test-MouseMoved (~2632-2650)
    │   ├── Get-TimeSinceMs (~2652-2660)
    │   ├── Get-ValueWithVariance (~2662-2685)
    │   ├── Get-Padding (~2686-2715)
    │   ├── Write-SimpleDialogRow (~2716-2745)
    │   ├── Write-SimpleFieldRow (~2746-2790)
    │   ├── Show-MovementModifyDialog (~2792-3520)
    │   └── Show-QuitConfirmationDialog (~3525-3720)
    │
    ├── P/Invoke Type Definitions (lines ~700-900)
    │   ├── POINT struct
    │   ├── CONSOLE_SCREEN_BUFFER_INFO struct
    │   ├── MOUSE_EVENT_RECORD struct
    │   ├── KEY_EVENT_RECORD struct
    │   ├── INPUT_RECORD struct (union: MouseEvent + KeyEvent)
    │   ├── COORD struct
    │   ├── SMALL_RECT struct
    │   ├── Keyboard class (keybd_event only)
    │   └── Mouse class (GetCursorPos, SetCursorPos, GetAsyncKeyState, FindWindow, GetLastInputInfo, PeekConsoleInput, ReadConsoleInput, etc.)
    │
    ├── Assembly Loading & Verification (lines ~700-1260)
    │   ├── Load System.Windows.Forms
    │   ├── Check for existing mJiggAPI types
    │   ├── Define types via Add-Type
    │   └── Verify API functionality
    │
    ├── End Time Calculation (lines ~1280-1350)
    │   ├── Apply variance to end time
    │   └── Determine if end time is today/tomorrow
    │
    ├── Main Loop (lines ~3654-6011)
    │   │
    │   ├── Loop Initialization (~3654-3700)
    │   │   └── Reset per-iteration state variables
    │   │
    │   ├── Interval Calculation (~3700-3730)
    │   │   └── Calculate random wait time with variance
    │   │
    │   ├── Wait Loop (~3706-4200)
    │   │   ├── Mouse position monitoring (Test-MouseMoved)
    │   │   ├── PeekConsoleInput (scroll + keyboard + mouse click detection)
    │   │   ├── GetLastInputInfo (system-wide activity + mouse inference)
    │   │   ├── Menu hotkey detection (console ReadKey)
    │   │   ├── Window resize detection
    │   │   └── Dialog invocation
    │   │
    │   ├── Mouse Settle Detection (~4200-4400)
    │   │   └── Wait for mouse to stop moving
    │   │
    │   ├── Resize Handling Loop (~4400-4800)
    │   │   ├── Draw-ResizeLogo -ClearFirst (atomic clear+draw)
    │   │   └── Wait for resize completion
    │   │
    │   ├── Movement Execution (~4800-5000)
    │   │   ├── Calculate random direction
    │   │   ├── Animate cursor movement
    │   │   └── Send simulated keypress
    │   │
    │   ├── UI Rendering (~5000-5930)
    │   │   ├── Header line
    │   │   ├── Horizontal separator
    │   │   ├── Log entries (full view)
    │   │   ├── Stats box (full view, wide window)
    │   │   ├── Bottom separator
    │   │   ├── Menu bar
    │   │   └── Hidden view (status line + (h) button)
    │   │
    │   └── End Time Check (~5990)
    │       └── Exit if scheduled time reached
    │
    └── Cleanup (~6000-6011)
        └── Display runtime statistics
```

---

## Key Concepts

### 1. Script-Scoped Variables

Parameters are copied to `$script:` variables because PowerShell parameters are read-only:

```powershell
$script:IntervalSeconds = $IntervalSeconds
$script:MoveSpeed = $MoveSpeed
$script:TravelDistance = $TravelDistance
```

These can be modified at runtime via the Modify Movement dialog. When accessing these in nested functions, always use the `$script:` prefix.

**Common script-scoped variables:**
- `$script:IntervalSeconds`, `$script:IntervalVariance` - Timing
- `$script:MoveSpeed`, `$script:MoveVariance` - Animation speed
- `$script:TravelDistance`, `$script:TravelVariance` - Movement distance
- `$script:AutoResumeDelaySeconds` - User input cooldown
- `$script:DiagEnabled` - Diagnostics flag
- `$script:LoopIteration` - Main loop counter
- `$script:MenuItemsBounds` - Click detection bounds array; each entry now also carries `displayText`, `format`, `fg`, `bg`, `hotkeyFg`, `onClickFg`, `onClickBg`, `onClickHotkeyFg`
- `$script:MenuClickHotkey` - Menu item hotkey triggered by mouse click
- `$script:ConsoleClickCoords` - Character cell X/Y from last PeekConsoleInput left-click event
- `$script:PressedMenuButton` - Hotkey of the menu button currently held down (LMB pressed); cleared when pressed state is restored
- `$script:ButtonClickedAt` - `[DateTime]` timestamp of a confirmed click (UP over button); used alongside `PendingDialogCheck`
- `$script:PendingDialogCheck` - `$true` after a confirmed click; render loop clears it on the first execution after the action, immediately restoring the button color unless a dialog is open
- `$script:LButtonWasDown` - Tracks previous LMB state from console `PeekConsoleInput` events for UP-transition detection
- `$script:RenderQueue` - `System.Collections.Generic.List[hashtable]` used by buffered rendering (`Write-Buffer`/`Flush-Buffer`)
- `$script:ESC` - `[char]27` for VT100 escape sequences
- `$script:CursorVisible` - Boolean tracking cursor visibility state for VT100 sequences
- `$script:AnsiFG` / `$script:AnsiBG` - ConsoleColor-to-ANSI SGR code lookup hashtables
- `$script:LastMouseMovementTime` - Stutter prevention timing
- `$script:ResizeQuotes` - Playful quotes array
- `$script:CurrentResizeQuote` - Currently displayed quote

### 2. P/Invoke (Platform Invoke)

The script defines Win32 API types in a C# code block via `Add-Type`. All types are in the `mJiggAPI` namespace:

```powershell
# Mouse position
$point = New-Object mJiggAPI.POINT
[mJiggAPI.Mouse]::GetCursorPos([ref]$point)
[mJiggAPI.Mouse]::SetCursorPos($x, $y)

# Mouse button state (only used for 0x01-0x06 mouse buttons)
$state = [mJiggAPI.Mouse]::GetAsyncKeyState($keyCode)

# Simulate keypress
[mJiggAPI.Keyboard]::keybd_event($VK_RMENU, 0, 0, 0)  # Key down
[mJiggAPI.Keyboard]::keybd_event($VK_RMENU, 0, $KEYEVENTF_KEYUP, 0)  # Key up

# System-wide input detection (keyboard, mouse, scroll -- passive, no scanning)
$lii = New-Object mJiggAPI.LASTINPUTINFO
[mJiggAPI.Mouse]::GetLastInputInfo([ref]$lii)

# Window detection
$handle = [mJiggAPI.Mouse]::GetForegroundWindow()
$consoleHandle = [mJiggAPI.Mouse]::GetConsoleWindow()
```

**Key structs:**
- `mJiggAPI.POINT` - X/Y coordinates
- `mJiggAPI.COORD` - Console coordinates (short X, short Y)
- `mJiggAPI.LASTINPUTINFO` - System idle time tracking (cbSize, dwTime)
- `mJiggAPI.KEY_EVENT_RECORD` - Console keyboard event (bKeyDown, wVirtualKeyCode, etc.)
- `mJiggAPI.MOUSE_EVENT_RECORD` - Console mouse event (dwMousePosition, dwEventFlags, etc.)
- `mJiggAPI.INPUT_RECORD` - Console input union (EventType + MouseEvent/KeyEvent overlay at offset 4)

**Key APIs:**
- `GetCursorPos` / `SetCursorPos` - Mouse position read/write
- `GetAsyncKeyState` - Mouse button state only (VK 0x01-0x06); also used for Shift/Ctrl modifier checks
- `keybd_event` - Simulate key presses
- `GetLastInputInfo` - Passive system-wide last input timestamp (detects all input: keyboard, mouse, scroll)
- `PeekConsoleInput` / `ReadConsoleInput` - Console input buffer access for scroll and keyboard event detection
- `FindWindow` / `EnumWindows` - Window handle lookup
- `GetForegroundWindow` - Currently active window
- `GetConsoleWindow` - This script's console window

### 3. Console TUI Rendering (VT100 Buffered)

All rendering goes through a buffered rendering system backed by VT100/ANSI escape sequences. Code calls `Write-Buffer` to queue positioned, colored text segments, then `Flush-Buffer` builds a single string with embedded VT100 escape codes and outputs the entire frame with one `[Console]::Write()` call.

```powershell
# Queue segments at specific positions with colors
Write-Buffer -X $col -Y $row -Text $text -FG $fgColor -BG $bgColor

# Queue sequential segments (continue from last position)
Write-Buffer -Text "more text" -FG $color

# Flush all queued segments to console (single atomic write)
Flush-Buffer

# Atomic clear screen + redraw (no visible blank flash)
Flush-Buffer -ClearFirst
```

**VT100 setup** (console mode block, ~line 458):
- `ENABLE_VIRTUAL_TERMINAL_PROCESSING` (0x0004) enabled on stdout handle via `SetConsoleMode`
- `[Console]::OutputEncoding` set to `[System.Text.Encoding]::UTF8` for correct emoji rendering

**ANSI color tables** (`$script:AnsiFG`, `$script:AnsiBG`, ~line 1464):
- Map all 16 `[ConsoleColor]` enum values to ANSI SGR codes (FG: 30-37/90-97, BG: 40-47/100-107)
- Segments with `$null` FG/BG use ANSI codes 39/49 (terminal default colors) instead of explicit color codes

**Buffer infrastructure** (`$script:RenderQueue`, `Write-Buffer`, `Flush-Buffer`, `Clear-Buffer`):
- `$script:RenderQueue` - `System.Collections.Generic.List[hashtable]` holding `@{ X; Y; Text; FG; BG }` segments
- `Write-Buffer` - Adds a segment. `X`/`Y` of `-1` = continue from last position. `FG`/`BG` of `$null` = use terminal default color. `-Wide` switch appends a trailing space for 2-column emoji background fill.
- `Flush-Buffer` - Builds a `StringBuilder` with VT100 sequences: `ESC[?25l` (hide cursor), `ESC[row;colH` (positioning), `ESC[fg;bgm` (colors, only emitted on change), segment text, `ESC[0m` (reset), optional `ESC[?25h` (show cursor if `$script:CursorVisible` is true). Single `[Console]::Write()` outputs the entire frame atomically. `-ClearFirst` switch prepends `ESC[2J` for atomic screen clear+redraw.
- `Clear-Buffer` - Discards all queued segments without writing.

**Cursor visibility** is tracked via `$script:CursorVisible` (boolean) and controlled with VT100 sequences (`ESC[?25l` / `ESC[?25h`) instead of `[Console]::CursorVisible`. The cursor is hidden during flush and conditionally shown at the end based on the tracked state (e.g., shown during dialog text input, hidden otherwise).

**Frame boundaries** (where `Flush-Buffer` is called):
1. After initial dialog draw (shadow + borders + fields + buttons)
2. After dialog field redraw (2 affected rows on navigation/click)
3. After dialog input value change (character typed, backspace, validation error)
4. After dialog resize handler redraw (`Flush-Buffer -ClearFirst`)
5. After dialog cleanup (clear shadow + clear area)
6. After full main UI render (header + separator + logs + stats + separator + menu)
7. After resize logo draw (`Draw-ResizeLogo -ClearFirst` on first draw)

**What stays as direct writes:**
- Debug/diagnostic logging to files (`Out-File`)
- One-off `Write-Host` calls during initialization/startup (before main loop)
- `[Console]::SetCursorPosition` for positioning the text input cursor in dialogs (after Flush-Buffer)
- Direct `[Console]::Write()` of VT100 cursor show/hide sequences during dialog text input state changes

**Box-drawing characters** are stored as variables to avoid encoding issues:

```powershell
$script:BoxTopLeft = [char]0x250C      # ┌
$script:BoxTopRight = [char]0x2510     # ┐
$script:BoxBottomLeft = [char]0x2514   # └
$script:BoxBottomRight = [char]0x2518  # ┘
$script:BoxHorizontal = [char]0x2500   # ─
$script:BoxVertical = [char]0x2502     # │
$script:BoxVerticalRight = [char]0x251C # ├
$script:BoxVerticalLeft = [char]0x2524  # ┤
```

### 4. Theme System

All colors are centralized as `$script:` variables (lines 214-289):

```powershell
# Menu Bar (normal state)
$script:MenuButtonBg = "DarkBlue"
$script:MenuButtonText = "White"
$script:MenuButtonHotkey = "Green"
$script:MenuButtonPipe = "White"
# Menu Bar (pressed / onclick state)
$script:MenuButtonOnClickBg     = "DarkCyan"
$script:MenuButtonOnClickFg     = "Black"
$script:MenuButtonOnClickHotkey = "Black"

# Dialogs
$script:QuitDialogBg = "DarkMagenta"
$script:QuitDialogShadow = "DarkMagenta"
$script:QuitDialogBorder = "White"
$script:QuitDialogTitle = "Yellow"
# ... etc
```

**Color categories:**
| Prefix | Component |
|--------|-----------|
| `MenuButton*` | Bottom menu bar (normal + `OnClick*` pressed state) |
| `Header*` | Top header line |
| `StatsBox*` | Right-side stats panel |
| `QuitDialog*` | Quit confirmation dialog |
| `TimeDialog*` | Set end time dialog |
| `MoveDialog*` | Modify movement dialog |
| `Resize*` | Window resize splash screen |
| `Text*` | General purpose colors |

### 4a. Button Click System

Menu buttons use a multi-phase click model:

**Phase 1 — Mouse DOWN** (`PeekConsoleInput` handler):
- Detect which `$script:MenuItemsBounds` entry is under the cursor
- Set `$script:PressedMenuButton = $btn.hotkey`
- Immediately call `Write-ButtonImmediate` with `onClickFg`/`onClickBg`/`onClickHotkeyFg` + `Flush-Buffer` — gives instant visual feedback without waiting for the next frame

**Phase 2 — Mouse UP over same button** (confirmed click):
- Set `$script:ConsoleClickCoords` to trigger the action
- Set `$script:ButtonClickedAt = Get-Date` and `$script:PendingDialogCheck = $true`
- Leave `$script:PressedMenuButton` set — render loop handles restoration

**Phase 2 — Mouse UP outside button** (cancelled click):
- `Start-Sleep 100ms` brief delay, then `Write-ButtonImmediate` with normal colors
- Clear `$script:PressedMenuButton` immediately

**Phase 3 — Render loop restoration** (top of menu bar render, checks `$script:PendingDialogCheck`):
- `$script:DialogButtonBounds -eq $null` (no dialog open) → clears `$script:PressedMenuButton` immediately — handles toggles (v, h) and instant actions
- `$script:DialogButtonBounds -ne $null` (dialog open) → skips; button stays pressed while dialog is open

**Popup persistence**: Dialog-opening actions (q, t, m) call `Show-*Dialog` synchronously, blocking the main loop. The button stays visually highlighted (from Phase 1) throughout because no main render runs during the dialog. When the dialog closes, `DialogButtonBounds` is cleared and the next render's `PendingDialogCheck` fires the restore.

**`Write-ButtonImmediate` function** (near `Flush-Buffer` definition):
- Params: `$btn` (bounds entry), `$fg`, `$bg`, `$hotkeyFg`
- Reads `$btn.displayText` and `$btn.format` to render full button text with emoji/pipe splitting
- Calls `Flush-Buffer` at the end for immediate console output

**`$script:MenuItemsBounds` entry schema:**
```
startX, endX, y        — click hit area
hotkey                 — single character hotkey
index                  — position in menuItems array
displayText            — current format text string (for Write-ButtonImmediate)
format                 — menuFormat int (0=emoji|pipe, 1=noIcons, 2=short)
fg, bg, hotkeyFg       — normal render colors
onClickFg, onClickBg, onClickHotkeyFg  — pressed-state colors
```

### 5. Mouse Stutter Prevention

The "settle" logic prevents the next movement cycle from starting while the user is moving the mouse. This is critical for user experience:

```powershell
# Settle detection loop (simplified)
$stableChecks = 0
$requiredStableChecks = 3  # ~75ms of stability
while ($true) {
    Start-Sleep -Milliseconds 25
    $currentPos = Get-MousePosition
    
    if ($currentPos.X -eq $lastPos.X -and $currentPos.Y -eq $lastPos.Y) {
        $stableChecks++
        if ($stableChecks -ge $requiredStableChecks) {
            break  # Mouse has settled
        }
    } else {
        $stableChecks = 0  # Reset - mouse still moving
    }
    $lastPos = $currentPos
}
```

Input detection during the wait loop uses `PeekConsoleInput` for keyboard, scroll, and mouse click events (with exact character cell coordinates for click-to-button mapping), `GetLastInputInfo` for system-wide activity, `Test-MouseMoved` for cursor position changes, and a focused `GetAsyncKeyState` loop over mouse buttons only (VK 0x01-0x06) for general input detection. No keyboard scanning (`GetAsyncKeyState` over key codes) is performed.

### 6. Window Resize Handling

There are two resize handler paths: a **self-contained `handleResize`** for the welcome screen, and the **`Invoke-ResizeHandler`** function for everything after initialization.

#### Size detection: use `GetConsoleScreenBufferInfo` directly

`$Host.UI.RawUI.WindowSize` and `[Console]::WindowWidth` both go through managed wrapper code that can return stale values. The welcome screen's `getSize` and `handleResize` use `[mJiggAPI.Mouse]::GetConsoleScreenBufferInfo` (the raw Win32 P/Invoke on the **stdout** handle, `GetStdHandle(-11)`) and read `srWindow.Right - srWindow.Left + 1` / `srWindow.Bottom - srWindow.Top + 1` directly. This is the lowest-level path possible and always returns the current terminal dimensions.

```powershell
$csbi = New-Object mJiggAPI.CONSOLE_SCREEN_BUFFER_INFO
$hOut = [mJiggAPI.Mouse]::GetStdHandle(-11)   # STD_OUTPUT_HANDLE
if ([mJiggAPI.Mouse]::GetConsoleScreenBufferInfo($hOut, [ref]$csbi)) {
    $w = [int]($csbi.srWindow.Right  - $csbi.srWindow.Left + 1)
    $h = [int]($csbi.srWindow.Bottom - $csbi.srWindow.Top  + 1)
}
```

#### Welcome screen: `handleResize` (nested inside `Show-StartupComplete`)

Self-contained; does **not** call `Invoke-ResizeHandler` or `Send-ResizeExitWakeKey`.

Before the outer polling loop starts, `Restore-ConsoleInputMode` and `Send-ResizeExitWakeKey` are called once to prime Windows Terminal's input routing (same mechanism used after main-loop resizes). `drainWakeKeys` is then called to consume the injected events before real keypress detection begins.

1. Read initial size via `getSize` (direct CSBI call)
2. Draw logo (or clear on error)
3. 10ms poll loop: `GetConsoleScreenBufferInfo` directly → detect change → redraw logo
4. Stability: 1500ms with no size change AND LMB released → break
5. `[Console]::Clear()`, `Restore-ConsoleInputMode`, `drainWakeKeys`, redraw welcome box

#### `drainWakeKeys` (nested inside `Show-StartupComplete`)

Drains the **entire** console input buffer and returns `$true` if any genuine keypress was found.

Critical rules (each learned from a diagnosed bug):
- **`IncludeKeyDown,IncludeKeyUp`** — `IncludeKeyDown` alone causes `ReadKey` to block indefinitely on KeyUp events, freezing the polling loop.
- **Drain the whole buffer, never return early** — if a stale KeyUp (e.g. the Enter used to run the script) causes an early return, the synthetic wake key events are left behind in the buffer and counted as real keypresses on the next tick.
- **Filter `VK_MENU` (18) as well as `VK_RMENU` (165)** — `Send-ResizeExitWakeKey` injects `VK_RMENU` (0xA5), but the Windows console input layer reports it as `VK_MENU` (18) in `INPUT_RECORD` keyboard events. See gotcha below.
- **Filter all modifier VKs** — Shift (16), Ctrl (17), Alt (18), and their L/R variants (160–165) are never "press any key".
- **Only count `KeyDown=true` as a real keypress** — stale KeyUp events from any previous key are discarded.

```powershell
$_wakeVKs = @(16, 17, 18, 160, 161, 162, 163, 164, 165)
function drainWakeKeys {
    $_real = $false
    try {
        while ($Host.UI.RawUI.KeyAvailable) {
            $k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp,AllowCtrlC")
            if ($k.KeyDown -and $k.VirtualKeyCode -notin $_wakeVKs) { $_real = $true }
        }
    } catch {}
    return $_real
}
```

#### Main loop: `Invoke-ResizeHandler`

Called from: the main wait loop resize check, and the per-iteration outside-wait-loop resize check. Not called from the welcome screen.

1. **Enter**: Reset `$script:CurrentResizeQuote` and `$script:ResizeLogoLockedHeight`
2. **Initial draw**: `Draw-ResizeLogo -ClearFirst` (normal mode) or `[Console]::Clear()` (hidden mode)
3. **1ms poll loop**:
   - Read `$psw.WindowSize` on every iteration (main loop already calls PeekConsoleInput externally)
   - If size changed: update pending size, reset stability timer, redraw logo (or nothing in hidden mode)
   - Every 50 redraws: `[Console]::Clear()` + `Restore-ConsoleInputMode` to prevent artifact buildup
   - Check stability: if `$elapsed -ge $ResizeThrottleMs` (1500ms) AND LMB not held → exit
4. **Exit**: `[Console]::Clear()`, `Restore-ConsoleInputMode`, `Send-ResizeExitWakeKey`, `return $pendingSize`

```powershell
# Callers use the returned stable size to update their tracking state:
$stableSize = Invoke-ResizeHandler
$oldWindowSize = $stableSize
$HostWidth     = $stableSize.Width
$HostHeight    = $stableSize.Height
```

**LMB gate**: After the stability timer expires, the exit is deferred if `GetAsyncKeyState(0x01) -band 0x8000` is set (mouse button still held). The timer is **not** reset by mouse state — only new size changes reset it.

**`$oldWindowSize` initialization**: Both `$oldWindowSize` and `$OldBufferSize` are set to the current live values immediately before the `:process while ($true)` main loop starts. This prevents the first-iteration `$null` comparison from triggering a spurious resize screen on every startup.

**`Draw-ResizeLogo` function:**
- Accepts `-ClearFirst` switch (passed through to `Flush-Buffer`)
- Calculates center position for logo
- Draws box with dynamic padding (42% of available space)
- Locks height during resize to absorb ±1 transient row fluctuations (Windows Terminal reflow)
- Queues all segments via `Write-Buffer`, then `Flush-Buffer -ClearFirst` at the end
- Shows random quote 2 lines below logo

### 7. Dialog System

Dialogs are modal functions that take control of input and rendering:

**Structure:**
1. Save cursor visibility state
2. Calculate dialog position (centered)
3. Draw drop shadow
4. Draw dialog box with borders
5. Enter input loop
6. Handle keypresses (Enter, Escape, Tab, arrows, etc.)
7. Handle window resize (redraw dialog)
8. Return result hashtable
9. Clear shadow and dialog area
10. Restore cursor visibility

**Dialog helper functions** (all write through `Write-Buffer`, do NOT call `Flush-Buffer` themselves -- the caller decides when to flush):

```powershell
# Draw a row with borders and background
Write-SimpleDialogRow -x $x -y $y -width $w -content "Hello" -contentColor White -backgroundColor $bg

# Draw an input field row
Write-SimpleFieldRow -x $x -y $y -width $w -label "Value:" -longestLabel $ll -fieldValue $val -fieldWidth 4 -fieldIndex 0 -currentFieldIndex $cur -backgroundColor $bg

# Draw offset shadow effect
Draw-DialogShadow -dialogX $x -dialogY $y -dialogWidth $w -dialogHeight $h -shadowColor DarkGray
```

**Result format:**
```powershell
return @{
    Result = $userInput      # The data (or $null if cancelled)
    NeedsRedraw = $true      # Whether main UI needs full refresh
}
```

### 8. Menu Item Bounds Tracking & Click Detection

For mouse click detection, menu items track their console coordinates:

```powershell
$script:MenuItemsBounds = @()
$script:MenuItemsBounds += @{
    startX = $itemStartX      # Left edge X coordinate
    endX = $itemEndX          # Right edge X coordinate  
    y = $menuY                # Row number
    hotkey = "t"              # Associated hotkey
    index = $i                # Item index
}
```

**Click detection uses `PeekConsoleInput` MOUSE_EVENT records.** The console input buffer provides native `MOUSE_EVENT` records with exact character cell coordinates (`dwMousePosition.X/Y`), eliminating all pixel-to-character conversion math. This is the same buffer used for keyboard and scroll detection.

Click detection in the main loop's PeekConsoleInput block:

```powershell
$script:ConsoleClickCoords = $null
# Inside the existing PeekConsoleInput loop:
if ($peekBuffer[$e].EventType -eq 0x0002) {  # MOUSE_EVENT
    $mouseFlags = $peekBuffer[$e].MouseEvent.dwEventFlags
    $mouseButtons = $peekBuffer[$e].MouseEvent.dwButtonState
    if ($mouseFlags -eq 0 -and ($mouseButtons -band 0x0001) -ne 0) {  # Left button press
        $script:ConsoleClickCoords = @{
            X = $peekBuffer[$e].MouseEvent.dwMousePosition.X
            Y = $peekBuffer[$e].MouseEvent.dwMousePosition.Y
        }
    }
}
# After the peek loop, consume click events via ReadConsoleInput
# Then match against dialog buttons and menu items:
if ($null -ne $script:ConsoleClickCoords) {
    $clickX = $script:ConsoleClickCoords.X
    $clickY = $script:ConsoleClickCoords.Y
    # Exact character cell comparisons:
    if ($clickY -eq $bounds.buttonRowY -and $clickX -ge $bounds.startX -and $clickX -le $bounds.endX) { ... }
}
```

**Hit-testing uses exact character cell matching** — no tolerance, no pixel math, no expanded bounding boxes. Button and menu item bounds map to the exact visible characters (emoji+pipe+text). The `PeekConsoleInput` approach inherently handles focus (events only appear when the console is focused) and provides coordinates that match the console's own rendering.

Dialog button click detection (in Show-TimeChangeDialog, Show-MovementModifyDialog, and Show-QuitConfirmationDialog) uses the same `PeekConsoleInput` pattern within each dialog's own input loop. Each dialog peeks for MOUSE_EVENT left button press records, consumes them via `ReadConsoleInput`, then matches against `$buttonRowY`, `$updateButtonStartX/$updateButtonEndX`, and `$cancelButtonStartX/$cancelButtonEndX` (all mapped to visible characters only).

Show-MovementModifyDialog also supports **field click selection**: after button checks, clicks within the dialog area are matched against field Y offsets (`@(4, 5, 7, 8, 10, 11, 13)` relative to `$dialogY`). A matched field click switches `$currentField`, redraws only the two affected rows (previous and new selection), and repositions the cursor.

**Important**: All three dialogs must clear `$script:DialogButtonBounds = $null` and `$script:DialogButtonClick = $null` in their cleanup code. The main loop's menu click detection is guarded by `$null -eq $script:DialogButtonBounds` — stale bounds will block all menu item clicks.

### 9. Emoji Handling

Emojis display as 2 columns in the console but have string length of 1. With buffered rendering, emoji positions are computed statically (assuming 2 display cells) rather than reading `$Host.UI.RawUI.CursorPosition.X` after writing:

```powershell
$emojiX = $itemStartX
$pipeX = $emojiX + 2  # Emoji takes 2 display columns
Write-Buffer -X $emojiX -Y $menuY -Text $emoji -FG $iconColor -BG $bg
Write-Buffer -X $pipeX -Y $menuY -Text "|" -FG $pipeColor -BG $bg
```

**Common emojis used:**
```powershell
$emojiHourglass = [char]::ConvertFromUtf32(0x23F3)   # ⏳
$emojiEye = [char]::ConvertFromUtf32(0x1F441)        # 👁️
$emojiLock = [char]::ConvertFromUtf32(0x1F512)       # 🔒
$emojiGear = [char]::ConvertFromUtf32(0x1F6E0)       # 🛠
$emojiRedX = [char]::ConvertFromUtf32(0x274C)        # ❌
$emojiMouse = [char]::ConvertFromUtf32(0x1F400)      # 🐀
```

### 10. Log Array Structure

Log entries use a component-based structure for dynamic truncation:

```powershell
$LogArray += [PSCustomObject]@{
    logRow = $true
    components = @(
        @{ 
            priority = 1              # Lower = more important
            text = "full text"        # Displayed when space allows
            shortText = "short"       # Displayed when truncated
        },
        @{ 
            priority = 2
            text = " - detailed message"
            shortText = " - msg"
        }
    )
}
```

Priority determines display order when truncating. Components with lower priority numbers are shown first.

### 11. Input Detection and State Tracking

Input detection uses four complementary mechanisms, each providing evidence for a specific input type:

1. **`PeekConsoleInput`** (keyboard + scroll + mouse clicks) - Peeks at the console input buffer for event records. Detects `KEY_EVENT` (EventType 0x0001) for keyboard, `MOUSE_EVENT` with scroll flag (EventType 0x0002, dwEventFlags 0x0004) for scroll wheel, and `MOUSE_EVENT` with left button press (dwEventFlags 0, dwButtonState & 0x0001) for click detection. Keyboard events are only **peeked** (not consumed) so the menu hotkey handler can still read them. Scroll and click events are consumed to prevent buffer buildup. The simulated Right Alt key (VK 0xA5) is filtered out. Only works when console is focused.

2. **`GetAsyncKeyState`** (mouse buttons only) - Focused loop over VK codes 0x01-0x06 for general input detection (pausing the jiggler). Not used for click-to-button mapping — that is handled entirely by PeekConsoleInput MOUSE_EVENT records.

3. **`GetLastInputInfo`** (system-wide catch-all) - Passive API returning the timestamp of the last user input of any type. Used to set `$script:userInputDetected = $true` (pauses the jiggler). Also infers **mouse movement** when activity is detected but no keyboard, scroll, or click evidence was found by the other mechanisms.

4. **`Test-MouseMoved`** (position polling) - Compares cursor position against previous check with a pixel threshold. Provides direct evidence of mouse movement.

**Classification logic (evidence-based, inference by elimination):**
- **Mouse clicks**: `PeekConsoleInput` MOUSE_EVENT left button press → direct evidence (with cell coords for button mapping); `GetAsyncKeyState` VK 0x01-0x06 → general detection for jiggler pause
- **Scroll**: `PeekConsoleInput` MOUSE_EVENT with scroll flag → direct evidence
- **Keyboard**: `PeekConsoleInput` KEY_EVENT records (excluding VK 0xA5) → direct evidence
- **Mouse movement**: `Test-MouseMoved` position change → direct evidence; OR `GetLastInputInfo` activity with no keyboard/scroll/click evidence → inference by elimination

```powershell
# Mouse button state (0x01-0x06 only)
$state = [mJiggAPI.Mouse]::GetAsyncKeyState($keyCode)
$isCurrentlyPressed = ($state -band 0x8000) -ne 0
$wasJustPressed = ($state -band 0x0001) -ne 0
```

**State tracking variables:**
- `$script:previousKeyStates` - Hashtable of previous mouse button states (for edge detection, VK 0x01-0x06 only)
- `$script:LastSimulatedKeyPress` - Timestamp of last simulated press (for filtering)
- `$keyboardInputDetected` - Boolean, set by `PeekConsoleInput` KEY_EVENT records
- `$mouseInputDetected` - Boolean, set by mouse movement (Test-MouseMoved or GetLastInputInfo inference) or button clicks
- `$scrollDetectedInInterval` - Boolean, set by `PeekConsoleInput` scroll events, persists across wait loop iterations
- `$script:userInputDetected` - Boolean, set by any detection mechanism, triggers jiggler pause

### 12. Movement Animation

Mouse movement is animated over time for a natural appearance. The path is generated by `Get-SmoothMovementPath` which produces points with ease-in-out-cubic easing and optional curved paths:

```powershell
# Calculate path and animate
$movementPath = Get-SmoothMovementPath -startX $pos.X -startY $pos.Y -endX $x -endY $y ...
$movementPoints = $movementPath.Points
$timePerPoint = $movementPath.TotalTimeMs / ($movementPoints.Count - 1)

for ($i = 1; $i -lt $movementPoints.Count; $i++) {
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point ($point.X, $point.Y)
    Start-Sleep -Milliseconds $sleepTime

    # Abort if user moved mouse during animation (drift > 3px from placed position)
    $actualPos = Get-MousePosition
    if ([Math]::Abs($actualPos.X - $point.X) -gt 3 -or [Math]::Abs($actualPos.Y - $point.Y) -gt 3) {
        $movementAborted = $true
        break
    }
}
```

**User input during animation**: After each `SetCursorPos` + sleep, the loop reads the actual cursor position and compares it to where the cursor was just placed. If the position has drifted by more than 3 pixels in either axis, the user is moving the mouse and the animation aborts immediately. On abort: `$script:userInputDetected` and `$mouseInputDetected` are set, the simulated keypress is skipped, and the auto-resume delay timer is started (if configured).

---

## Common Modification Patterns

### Adding a New Theme Color

1. Add variable in Theme Colors section (~line 214):
```powershell
$script:NewComponentColor = "Cyan"
$script:NewComponentBg = "DarkBlue"
```

2. Use in code:
```powershell
Write-Buffer -Text "text" -FG $script:NewComponentColor -BG $script:NewComponentBg
```

3. Update `resources/AGENTS.md` color categories table.

### Adding a New Parameter

1. Add to param block (~line 122):
```powershell
[Parameter(Mandatory = $false)]
[int]$NewParam = 10
```

2. Copy to script scope (~line 156):
```powershell
$script:NewParam = $NewParam
```

3. Update README.md parameters table.

### Adding a New Dialog

1. Create function following pattern of `Show-TimeChangeDialog`
2. Key elements:
   - Save `$savedCursorVisible = $script:CursorVisible`
   - Calculate centered position
   - Queue all rendering via `Write-Buffer` (borders, content, fields, buttons)
   - Call `Draw-DialogShadow` (also uses `Write-Buffer`)
   - Call `Flush-Buffer` after the complete dialog is queued
   - Input loop with resize detection (use `Flush-Buffer -ClearFirst` on resize)
   - On field/input redraws: queue affected rows via `Write-Buffer`, then `Flush-Buffer`
   - Cursor visibility: `$script:CursorVisible = $true; [Console]::Write("$($script:ESC)[?25h")` to show, `$script:CursorVisible = $false; [Console]::Write("$($script:ESC)[?25l")` to hide
   - Call `Clear-DialogShadow` + queue clear area via `Write-Buffer`, then `Flush-Buffer`
   - Restore `$script:CursorVisible = $savedCursorVisible` and write appropriate VT100 sequence
   - Return `@{ Result = $data; NeedsRedraw = $bool }`

3. Add hotkey handler in wait loop (~line 5400)
4. Update README.md interactive controls

### Adding a Menu Item

1. Add to `$menuItemsList` array (~line 7092):
```powershell
@{
    full = "$emojiNew|new_(x)feature"
    noIcons = "new_(x)feature"
    short = "(x)new"
}
```

2. Add hotkey handler in wait loop input processing
3. Update README.md interactive controls

### Adding a New Box-Drawing Character

1. Define at top of initialization (~line 210):
```powershell
$script:BoxNewChar = [char]0xXXXX  # Character name
```

2. Never use literal box characters in code - always use variables

### Modifying Movement Behavior

Key locations:
- `$script:IntervalSeconds` - Wait time between cycles
- `$script:TravelDistance` - Pixels to move
- `$script:MoveSpeed` - Animation duration
- Movement calculation: ~line 5900
- Animation loop: ~line 5950

---

## Important Gotchas

### Encoding Issues (CRITICAL)

Box-drawing characters can corrupt if file encoding changes. **Always use `[char]` casts:**

```powershell
# SAFE - generates character at runtime
$char = [char]0x250C  # ┌

# RISKY - can corrupt to â"Œ if encoding changes
$char = "┌"
```

If you see `â"Œ` or similar garbage, the file encoding has been corrupted. Fix by:
1. Re-saving with UTF-8 BOM encoding
2. Better: Convert all literal box chars to `[char]` casts

### Resize Detection: Use `GetConsoleScreenBufferInfo` Directly (CRITICAL)

`$Host.UI.RawUI.WindowSize` and `[Console]::WindowWidth` both go through managed wrappers that can return stale values. For reliable resize detection in a polling loop, call `GetConsoleScreenBufferInfo` on the **stdout** handle directly and read `srWindow`:

```powershell
# CORRECT - direct Win32, always current
$csbi = New-Object mJiggAPI.CONSOLE_SCREEN_BUFFER_INFO
$hOut = [mJiggAPI.Mouse]::GetStdHandle(-11)   # STD_OUTPUT_HANDLE (-11)
if ([mJiggAPI.Mouse]::GetConsoleScreenBufferInfo($hOut, [ref]$csbi)) {
    $w = [int]($csbi.srWindow.Right  - $csbi.srWindow.Left + 1)
    $h = [int]($csbi.srWindow.Bottom - $csbi.srWindow.Top  + 1)
}
```

Note: use `GetStdHandle(-11)` (stdout) for `GetConsoleScreenBufferInfo`, not `-10` (stdin).

### `VK_RMENU` (165) Appears as `VK_MENU` (18) in Console Input Records (CRITICAL)

`Send-ResizeExitWakeKey` injects `VK_RMENU` (0xA5 = 165) via `keybd_event`. However, the Windows console input layer reports this in `INPUT_RECORD` keyboard events with `wVirtualKeyCode = 18` (`VK_MENU`), **not** 165. Any code that filters wake keys by checking `VirtualKeyCode -eq 165` will miss them entirely. Always filter both 18 and 165 (and all other modifier VKs 16, 160–165) when reading from the console input buffer.

### `ReadKey("IncludeKeyDown")` Blocks Indefinitely on KeyUp Events

`$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")` waits for a KeyDown event and **skips** KeyUp events — meaning if a KeyUp event is at the front of the buffer, `ReadKey` hangs waiting for the next KeyDown. `KeyAvailable` returns `$true` for both KeyDown and KeyUp, so calling `ReadKey("IncludeKeyDown")` after a truthy `KeyAvailable` check can freeze the entire polling loop indefinitely.

**Always use `"NoEcho,IncludeKeyDown,IncludeKeyUp,AllowCtrlC"`** when draining the input buffer. Filter KeyDown/KeyUp in code after the read.

### Console Buffer vs Window Size

```powershell
$Host.UI.RawUI.WindowSize   # Visible area (use this for UI layout)
$Host.UI.RawUI.BufferSize   # Total scrollable area (larger)
```

Always use `WindowSize` for calculating UI positions and widths.

### Script Scope vs Local Scope

Variables modified in nested functions need `$script:` prefix to persist:

```powershell
# WRONG - creates local variable, doesn't modify script state
function Update-Setting {
    $IntervalSeconds = 5  # Local only!
}

# CORRECT - modifies script-scoped variable
function Update-Setting {
    $script:IntervalSeconds = 5  # Persists!
}
```

### GetAsyncKeyState Return Values

```powershell
$state = [mJiggAPI.Mouse]::GetAsyncKeyState($keyCode)

# Bit 15 (0x8000) - Key is currently down
$isPressed = ($state -band 0x8000) -ne 0

# Bit 0 (0x0001) - Key was pressed since last GetAsyncKeyState call
$wasPressed = ($state -band 0x0001) -ne 0
```

Note: The "was pressed" bit is consumed on read, so only check it once per call. Only used for mouse buttons (0x01-0x06) and specific modifier keys (Shift, Ctrl).

### Type Reloading Limitations

PowerShell cannot unload types once loaded via `Add-Type`. If you modify the C# type definitions, users must restart their PowerShell session. The script checks for existing types and skips reload if they exist.

### Simulated Key Press Filtering

When checking `GetLastInputInfo`, filter out the script's own simulated key presses and automated mouse movements:

```powershell
$recentSimulated = ($null -ne $LastSimulatedKeyPress) -and ((Get-TimeSinceMs -startTime $LastSimulatedKeyPress) -lt 500)
$recentAutoMove = ($null -ne $LastAutomatedMouseMovement) -and ((Get-TimeSinceMs -startTime $LastAutomatedMouseMovement) -lt 500)
```

The script simulates Right Alt (VK_RMENU = 0xA5) via `keybd_event`. After the simulated keypress, the console input buffer is flushed to prevent stale simulated events from being detected as user keyboard input by `PeekConsoleInput`. The `PeekConsoleInput` keyboard scan also explicitly filters out VK 0xA5.

### Emoji Display Width Variations

Some emojis render as 1 column, others as 2. With buffered rendering, emoji positions are computed statically (assuming 2 display cells) and explicit X positions are used after each emoji:

```powershell
$emojiX = $startX
Write-Buffer -X $emojiX -Y $row -Text $emoji -FG $color -BG $bg
Write-Buffer -X ($emojiX + 2) -Y $row -Text "|rest" -FG $color -BG $bg
```

### Windows Terminal Color Override

Windows Terminal has a setting "Automatically adjust lightness of indistinguishable text" that can override foreground colors. This cannot be controlled from PowerShell - users must disable it in Windows Terminal settings if they encounter color issues.

---

## State Machine Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                         MAIN LOOP                               │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌──────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐   │
│  │  WAIT    │───►│  SETTLE  │───►│  MOVE    │───►│  RENDER  │   │
│  │  LOOP    │    │  CHECK   │    │  CURSOR  │    │  UI      │   │
│  └────┬─────┘    └──────────┘    └──────────┘    └────┬─────┘   │
│       │                                               │         │
│       │  ┌──────────────────────────────────────┐     │         │
│       └──┤  Hotkey / Click / Resize Detection   ├─────┘         │
│          └──────────────────────────────────────┘               │
│                          │                                      │
│          ┌───────────────┼───────────────┐                      │
│          ▼               ▼               ▼                      │
│    ┌──────────┐    ┌──────────┐    ┌──────────┐                 │
│    │  QUIT    │    │  TIME    │    │  MOVE    │                 │
│    │  DIALOG  │    │  DIALOG  │    │  DIALOG  │                 │
│    └──────────┘    └──────────┘    └──────────┘                 │
│                                                                 │
│    ┌──────────────────────────────────────────────────────┐     │
│    │               RESIZE HANDLING LOOP                   │     │
│    │  (Clear screen → Draw logo → Wait for completion)    │     │
│    └──────────────────────────────────────────────────────┘     │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## Testing Tips

1. **Debug Mode**: Run with `-DebugMode` for verbose console logging during initialization
2. **Diagnostics**: Run with `-Diag` for file-based logs in `_diag/` (relative to script location)
3. **Settle Detection**: Test by moving mouse during interval countdown - movement should be deferred
4. **Resize Handling**: Drag window edges to test logo centering and quote display
5. **Dialog Rendering**: Test dialogs at various window sizes (they should stay centered)
6. **Click Detection**: Test clicking menu items vs clicking elsewhere
7. **Encoding**: After any file modification, verify box characters render correctly

---

## File Locations

| File | Purpose |
|------|---------|
| `start-mjig.ps1` | Main script (single file) |
| `README.md` | User documentation |
| `resources/AGENTS.md` | AI agent context (this file) |
| `resources/test-logs.ps1` | Temporary test script for log rendering (git-ignored) |
| `CHANGELOG.md` | Change tracking across commits |
| `.gitignore` | Excludes `_diag/`, backup files, and `resources/*.ps1` from git |
| `_diag/startup.txt` | Initialization diagnostics (created with `-Diag`) |
| `_diag/settle.txt` | Mouse settle detection logs (created with `-Diag`) |
| `_diag/input.txt` | PeekConsoleInput + GetLastInputInfo input detection logs (created with `-Diag`) |
| `_diag/welcome.txt` | Welcome screen resize detection diagnostics (**always written**, no `-Diag` flag needed) |

The `_diag/` folder is at the **project root** (`c:\Projects\mJigg\_diag\`), one level above the script (`Start-mJig\`). The script uses `Split-Path $PSScriptRoot -Parent` to build the path so it lands at the project root regardless of where the script file lives within the repo. `welcome.txt` is always written regardless of `-Diag`; all other diag files require the `-Diag` flag. All diag files are git-ignored.

> **TEMPORARY TEST SCRIPTS**: When an agent creates a throwaway `.ps1` script to test or experiment with something (e.g. testing rendering logic, validating a calculation), place it in `resources/`. All `resources/*.ps1` files are git-ignored. Do NOT place temp scripts in the project root or elsewhere. Note: `_diag/` (at the project root, not inside `Start-mJig/`) is separate — it is for runtime diagnostic output produced by the script itself (via `-Diag` or always-on), not for agent-authored test scripts.

**When reviewing diagnostic output with the user**, always provide a ready-to-run command to print the relevant diag file. The user expects this every time. Use:

```powershell
# Run these from the project root (c:\Projects\mJigg)
Get-Content ".\_diag\input.txt"
Get-Content ".\_diag\startup.txt"
Get-Content ".\_diag\settle.txt"
Get-Content ".\_diag\welcome.txt"   # always present, no -Diag flag needed
# Or full paths:
Get-Content "c:\Projects\mJigg\_diag\welcome.txt"
```

No external dependencies - the script is fully self-contained.

---

## Quick Reference: Key Line Numbers

| Component | Approximate Lines |
|-----------|------------------|
| Parameters | 41-120 |
| Box Characters | 124-132 |
| Theme Colors | 134-210 |
| Show-StartupScreen | ~220-252 |
| Show-StartupComplete | ~253-355 |
| Invoke-ResizeHandler | ~341-382 |
| Find-WindowHandle | ~384-470 |
| P/Invoke Types | ~700-980 |
| Buffered Rendering Functions (Write-Buffer, Flush-Buffer, Write-ButtonImmediate) | ~1663-1780 |
| Draw-DialogShadow / Clear-DialogShadow | ~1785-1830 |
| Show-TimeChangeDialog | ~1835-2410 |
| Draw-ResizeLogo | ~2495-2610 |
| Get-MousePosition / Test-MouseMoved | ~2615-2650 |
| Write-SimpleDialogRow / Write-SimpleFieldRow | ~2716-2790 |
| Show-MovementModifyDialog | ~2792-3520 |
| Show-QuitConfirmationDialog | ~3525-3720 |
| $oldWindowSize / $OldBufferSize init (pre-main-loop) | ~3944-3950 |
| Main Loop Start | ~3955 |
| Wait Loop | ~3970-4860 |
| Resize Detection (wait loop, calls Invoke-ResizeHandler) | ~4760-4855 |
| Resize Detection (outside wait loop, calls Invoke-ResizeHandler) | ~4900-4960 |
| UI Rendering | ~5100-6050 |
| Menu Rendering | ~5700-6050 |

*Note: Line numbers are approximate and shift as code is modified.*
