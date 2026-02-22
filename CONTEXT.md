# mJig Code Context for AI Agents

This document provides deep context for AI agents working on the `start-mjig.ps1` codebase.

> **IMPORTANT FOR AI AGENTS**: When modifying `start-mjig.ps1`, you must also update this `CONTEXT.md` file and `README.md` to reflect any changes. This includes:
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
4. Define helper functions
5. Parse and validate parameters
6. Enter main processing loop
   ├── Wait for interval (with input monitoring)
   ├── Check for user input / hotkeys
   ├── Wait for mouse to settle (stutter prevention)
   ├── Perform automated mouse movement
   ├── Send simulated keypress
   ├── Render UI (header, logs, stats, menu)
   └── Handle window resize
7. Cleanup on exit
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
    │   ├── Menu bar colors
    │   ├── Header colors
    │   ├── Stats box colors
    │   ├── Dialog colors (Quit, Time, Movement)
    │   ├── Resize screen colors
    │   └── General UI colors
    │
    ├── Helper Functions (lines 212-3600)
    │   ├── Find-WindowHandle (~212-400)
    │   ├── Buffered Rendering (Write-Buffer, Flush-Buffer, Clear-Buffer) (~1454-1515)
    │   ├── Draw-DialogShadow / Clear-DialogShadow (~1518-1555)
    │   ├── Show-TimeChangeDialog (~1557-2220)
    │   ├── Draw-ResizeLogo (~2225-2320)
    │   ├── Get-MousePosition (~2326-2340)
    │   ├── Test-MouseMoved (~2337-2355)
    │   ├── Get-TimeSinceMs (~2351-2358)
    │   ├── Get-ValueWithVariance (~2358-2380)
    │   ├── Get-Padding (~2381-2408)
    │   ├── Write-SimpleDialogRow (~2409-2434)
    │   ├── Write-SimpleFieldRow (~2434-2475)
    │   ├── Show-MovementModifyDialog (~2478-3215)
    │   └── Show-QuitConfirmationDialog (~3219-3600)
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
- `$script:MenuItemsBounds` - Click detection bounds
- `$script:MenuClickHotkey` - Menu item hotkey triggered by mouse click
- `$script:ConsoleClickCoords` - Character cell X/Y from last PeekConsoleInput left-click event
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
# Menu Bar
$script:MenuButtonBg = "DarkBlue"
$script:MenuButtonText = "White"
$script:MenuButtonHotkey = "Green"
$script:MenuButtonPipe = "White"

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
| `MenuButton*` | Bottom menu bar |
| `Header*` | Top header line |
| `StatsBox*` | Right-side stats panel |
| `QuitDialog*` | Quit confirmation dialog |
| `TimeDialog*` | Set end time dialog |
| `MoveDialog*` | Modify movement dialog |
| `Resize*` | Window resize splash screen |
| `Text*` | General purpose colors |

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

When the console window is resized:

1. **Detection**: Compare `$Host.UI.RawUI.WindowSize` against stored size
2. **Clear+Draw**: Atomically clear screen and draw logo via `Draw-ResizeLogo -ClearFirst`
3. **Logo**: Draw centered "mJig(🐀)" with decorative box
4. **Quote**: Display random playful quote from `$script:ResizeQuotes`
5. **Wait**: Stay in tight loop, redrawing only on size change
6. **Debounce**: Wait 2 seconds after resize stops before full UI redraw

```powershell
# Resize detection
$currentSize = $Host.UI.RawUI.WindowSize
$isNewSize = ($currentSize.Width -ne $PendingResizeWidth) -or 
             ($currentSize.Height -ne $PendingResizeHeight)

if ($isNewSize -and -not $ResizeClearedScreen) {
    $script:CurrentResizeQuote = $null
    Draw-ResizeLogo -ClearFirst
    $ResizeClearedScreen = $true
}
```

The `Draw-ResizeLogo` function:
- Accepts `-ClearFirst` switch (passed through to `Flush-Buffer`)
- Calculates center position for logo
- Draws box with dynamic padding (42% of available space)
- Queues all segments via `Write-Buffer`, then `Flush-Buffer` (or `Flush-Buffer -ClearFirst`) at the end
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

Mouse movement is animated over time for a natural appearance:

```powershell
# Calculate random direction and distance
$angle = Get-Random -Minimum 0 -Maximum 360
$distance = Get-ValueWithVariance -baseValue $script:TravelDistance -variance $script:TravelVariance
$targetX = $currentX + [math]::Cos($angle * [math]::PI / 180) * $distance
$targetY = $currentY + [math]::Sin($angle * [math]::PI / 180) * $distance

# Animate movement
$moveDuration = Get-ValueWithVariance -baseValue $script:MoveSpeed -variance $script:MoveVariance
$steps = [math]::Max(1, [math]::Floor($moveDuration * 1000 / 16))  # ~60fps
for ($i = 1; $i -le $steps; $i++) {
    $progress = $i / $steps
    $newX = $currentX + ($targetX - $currentX) * $progress
    $newY = $currentY + ($targetY - $currentY) * $progress
    [mJiggAPI.Mouse]::SetCursorPos([int]$newX, [int]$newY)
    Start-Sleep -Milliseconds 16
}
```

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

3. Update CONTEXT.md color categories table.

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
│  ┌──────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐  │
│  │  WAIT    │───►│  SETTLE  │───►│  MOVE    │───►│  RENDER  │  │
│  │  LOOP    │    │  CHECK   │    │  CURSOR  │    │  UI      │  │
│  └────┬─────┘    └──────────┘    └──────────┘    └────┬─────┘  │
│       │                                               │         │
│       │  ┌──────────────────────────────────────┐    │         │
│       └──┤  Hotkey / Click / Resize Detection   ├────┘         │
│          └──────────────────────────────────────┘              │
│                          │                                      │
│          ┌───────────────┼───────────────┐                     │
│          ▼               ▼               ▼                     │
│    ┌──────────┐    ┌──────────┐    ┌──────────┐               │
│    │  QUIT    │    │  TIME    │    │  MOVE    │               │
│    │  DIALOG  │    │  DIALOG  │    │  DIALOG  │               │
│    └──────────┘    └──────────┘    └──────────┘               │
│                                                                 │
│    ┌──────────────────────────────────────────────────────┐    │
│    │               RESIZE HANDLING LOOP                    │    │
│    │  (Clear screen → Draw logo → Wait for completion)     │    │
│    └──────────────────────────────────────────────────────┘    │
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
| `CONTEXT.md` | AI agent context (this file) |
| `CHANGELOG.md` | Change tracking across commits |
| `.gitignore` | Excludes `_diag/` and backup files from git |
| `_diag/startup.txt` | Initialization diagnostics (created with `-Diag`) |
| `_diag/settle.txt` | Mouse settle detection logs (created with `-Diag`) |
| `_diag/input.txt` | PeekConsoleInput + GetLastInputInfo input detection logs (created with `-Diag`) |

The `_diag/` folder is created in the same directory as `start-mjig.ps1` when run with `-Diag`. It is git-ignored. AI agents can read these files directly from the project directory to diagnose runtime issues.

**When reviewing diagnostic output with the user**, always provide a ready-to-run command to print the relevant diag file. The user expects this every time. Use:

```powershell
Get-Content ".\_diag\input.txt"
Get-Content ".\_diag\startup.txt"
Get-Content ".\_diag\settle.txt"
```

No external dependencies - the script is fully self-contained.

---

## Quick Reference: Key Line Numbers

| Component | Approximate Lines |
|-----------|------------------|
| Parameters | 41-120 |
| Box Characters | 124-132 |
| Theme Colors | 134-210 |
| Buffered Rendering Functions | 1454-1515 |
| P/Invoke Types | 700-900 |
| Draw-DialogShadow / Clear-DialogShadow | 1518-1555 |
| Show-TimeChangeDialog | 1557-2220 |
| Draw-ResizeLogo | 2225-2320 |
| Get-MousePosition / Test-MouseMoved | 2326-2355 |
| Write-SimpleDialogRow / Write-SimpleFieldRow | 2409-2475 |
| Show-MovementModifyDialog | 2478-3215 |
| Show-QuitConfirmationDialog | 3219-3600 |
| Main Loop Start | 3654 |
| Wait Loop | 3706-4400 |
| Resize Handling | 4400-4800 |
| UI Rendering | 5000-5930 |
| Menu Rendering | 5600-5930 |

*Note: Line numbers are approximate and may shift as code is modified.*
