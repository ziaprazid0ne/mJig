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

---

## Architecture Overview

The script is a single-file PowerShell application (~7,500 lines) implementing a console-based TUI mouse jiggler. It uses Win32 API calls via P/Invoke for low-level mouse/keyboard interaction.

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
├── Get-KeyName function (lines 1-80)
│   └── Standalone helper for mapping VK codes to readable names
│
└── Start-mJig function (lines 82-end)
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
    ├── Helper Functions (lines 291-2800)
    │   ├── Find-WindowHandle (~291-400)
    │   ├── Get-Padding (~400-420)
    │   ├── Get-TimeSinceMs (~420-440)
    │   ├── Get-ValueWithVariance (~440-460)
    │   ├── Get-MousePosition (~460-500)
    │   ├── Test-MouseMoved (~500-520)
    │   ├── Draw-DialogShadow (~520-600)
    │   ├── Clear-DialogShadow (~600-650)
    │   ├── Write-SimpleDialogRow (~650-750)
    │   ├── Write-SimpleFieldRow (~750-850)
    │   ├── Show-MovementModifyDialog (~1600-2400)
    │   ├── Show-QuitConfirmationDialog (~2400-2600)
    │   ├── Show-TimeChangeDialog (~2600-2800)
    │   └── Draw-ResizeLogo (~2800-2950)
    │
    ├── P/Invoke Type Definitions (lines ~700-1200)
    │   ├── POINT struct
    │   ├── RECT struct
    │   ├── CONSOLE_SCREEN_BUFFER_INFO struct
    │   ├── MOUSE_EVENT_RECORD struct
    │   ├── INPUT_RECORD struct
    │   ├── COORD struct
    │   ├── SMALL_RECT struct
    │   ├── Keyboard class (GetAsyncKeyState, keybd_event)
    │   ├── Mouse class (GetCursorPos, SetCursorPos, FindWindow, etc.)
    │   └── MouseHook class (wheel detection)
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
    ├── Main Loop (lines ~4600-7400)
    │   │
    │   ├── Loop Initialization (~4620-4640)
    │   │   └── Reset per-iteration state variables
    │   │
    │   ├── Interval Calculation (~4640-4670)
    │   │   └── Calculate random wait time with variance
    │   │
    │   ├── Wait Loop (~4670-5400)
    │   │   ├── Mouse position monitoring
    │   │   ├── Keyboard state scanning
    │   │   ├── Menu hotkey detection
    │   │   ├── Mouse click handling
    │   │   ├── Window resize detection
    │   │   └── Dialog invocation
    │   │
    │   ├── Mouse Settle Detection (~5400-5600)
    │   │   └── Wait for mouse to stop moving
    │   │
    │   ├── Resize Handling Loop (~5600-5800)
    │   │   ├── Clear screen on resize start
    │   │   ├── Draw centered logo/box
    │   │   └── Wait for resize completion
    │   │
    │   ├── Movement Execution (~5800-6200)
    │   │   ├── Calculate random direction
    │   │   ├── Animate cursor movement
    │   │   └── Send simulated keypress
    │   │
    │   ├── UI Rendering (~6200-7200)
    │   │   ├── Header line
    │   │   ├── Horizontal separator
    │   │   ├── Log entries (full view)
    │   │   ├── Stats box (full view, wide window)
    │   │   ├── Bottom separator
    │   │   └── Menu bar
    │   │
    │   └── End Time Check (~7400)
    │       └── Exit if scheduled time reached
    │
    └── Cleanup (~7450-7476)
        ├── Uninstall mouse hook
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

# Keyboard state
$state = [mJiggAPI.Keyboard]::GetAsyncKeyState($keyCode)

# Simulate keypress
[mJiggAPI.Keyboard]::keybd_event($VK_RMENU, 0, 0, 0)  # Key down
[mJiggAPI.Keyboard]::keybd_event($VK_RMENU, 0, $KEYEVENTF_KEYUP, 0)  # Key up

# Window detection
$handle = [mJiggAPI.Mouse]::GetForegroundWindow()
$consoleHandle = [mJiggAPI.Mouse]::GetConsoleWindow()
```

**Key structs:**
- `mJiggAPI.POINT` - X/Y coordinates
- `mJiggAPI.RECT` - Window rectangle (Left, Top, Right, Bottom)
- `mJiggAPI.COORD` - Console coordinates (short X, short Y)

**Key APIs:**
- `GetCursorPos` / `SetCursorPos` - Mouse position read/write
- `GetAsyncKeyState` - System-wide keyboard/mouse button state
- `keybd_event` - Simulate key presses
- `FindWindow` / `EnumWindows` - Window handle lookup
- `GetForegroundWindow` - Currently active window
- `GetConsoleWindow` - This script's console window
- `ScreenToClient` - Convert screen coords to window coords

### 3. Console TUI Rendering

The UI uses `[Console]::SetCursorPosition()` and `Write-Host` for precise character placement:

```powershell
[Console]::SetCursorPosition($x, $y)
Write-Host "text" -NoNewline -ForegroundColor $color -BackgroundColor $bg
```

**Key patterns:**

```powershell
# Draw at specific position
[Console]::SetCursorPosition($col, $row)
Write-Host $text -NoNewline

# Fill remaining line width
$remaining = $HostWidth - [Console]::CursorPosition.X
Write-Host (" " * $remaining) -NoNewline

# Draw box border
Write-Host "$($script:BoxTopLeft)$($script:BoxHorizontal * $width)$($script:BoxTopRight)"
```

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

Additionally, the wait loop skips expensive `GetAsyncKeyState` scanning when mouse has moved recently:

```powershell
$mouseRecentlyMoving = ($null -ne $script:LastMouseMovementTime) -and 
                       ((Get-TimeSinceMs -startTime $script:LastMouseMovementTime) -lt 200)

if ($mouseRecentlyMoving) {
    Start-Sleep -Milliseconds 50
    continue  # Skip keyboard scanning
}
```

### 6. Window Resize Handling

When the console window is resized:

1. **Detection**: Compare `$Host.UI.RawUI.WindowSize` against stored size
2. **Clear**: Immediately `Clear-Host` when resize starts
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
    Clear-Host
    $script:CurrentResizeQuote = $null  # Get new quote
    Draw-ResizeLogo
    $ResizeClearedScreen = $true
}
```

The `Draw-ResizeLogo` function:
- Calculates center position for logo
- Draws box with dynamic padding (42% of available space)
- Uses `[Console]::Write()` for performance
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

**Dialog helper functions:**

```powershell
# Draw a row with borders and background
Write-SimpleDialogRow -text "Hello" -dialogX $x -dialogWidth $w -bgColor $bg -borderColor $border

# Draw an input field row
Write-SimpleFieldRow -label "Value:" -value $val -fieldWidth 4 -dialogX $x -dialogWidth $w

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

### 8. Menu Item Bounds Tracking

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

Click detection in the wait loop:

```powershell
$clickPos = Get-MousePosition
# Convert to console coordinates
$consoleX = # ... (involves ScreenToClient and font size calculation)
$consoleY = # ...

foreach ($item in $script:MenuItemsBounds) {
    if ($consoleY -eq $item.y -and $consoleX -ge $item.startX -and $consoleX -le $item.endX) {
        $lastKeyPress = $item.hotkey  # Simulate hotkey press
        break
    }
}
```

### 9. Emoji Handling

Emojis display as 2 columns in the console but have string length of 1. The code accounts for this:

```powershell
$pipeX = $itemStartX + 2  # Emoji takes 2 display columns
Write-Host $emoji -NoNewline

# Check actual cursor position after emoji
$cursorAfterEmoji = [Console]::CursorPosition.X

if ($cursorAfterEmoji -lt $pipeX) {
    # Single-column emoji (like 👁️) - fill the gap
    Write-Host " " -NoNewline -BackgroundColor $bg
}
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

The script uses `GetAsyncKeyState` for system-wide input detection:

```powershell
$state = [mJiggAPI.Keyboard]::GetAsyncKeyState($keyCode)
$isCurrentlyPressed = ($state -band 0x8000) -ne 0    # Key is down now
$wasJustPressed = ($state -band 0x0001) -ne 0        # Key was pressed since last check
```

**State tracking variables:**
- `$script:previousKeyStates` - Hashtable of previous key states (for edge detection)
- `$PressedKeys` - Currently pressed keys (for stats display)
- `$intervalKeys` - Keys pressed during current interval
- `$script:LastSimulatedKeyPress` - Timestamp of last simulated press (for filtering)

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
Write-Host "text" -ForegroundColor $script:NewComponentColor -BackgroundColor $script:NewComponentBg
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
   - Save `$savedCursorVisible = [Console]::CursorVisible`
   - Calculate centered position
   - Call `Draw-DialogShadow`
   - Draw dialog box with theme colors
   - Input loop with resize detection
   - Call `Clear-DialogShadow` before cleanup
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
$state = [mJiggAPI.Keyboard]::GetAsyncKeyState($keyCode)

# Bit 15 (0x8000) - Key is currently down
$isPressed = ($state -band 0x8000) -ne 0

# Bit 0 (0x0001) - Key was pressed since last GetAsyncKeyState call
$wasPressed = ($state -band 0x0001) -ne 0
```

Note: The "was pressed" bit is consumed on read, so only check it once per call.

### Type Reloading Limitations

PowerShell cannot unload types once loaded via `Add-Type`. If you modify the C# type definitions, users must restart their PowerShell session. The script checks for existing types and skips reload if they exist.

### Simulated Key Press Filtering

When checking for keyboard input, filter out the script's own simulated key presses:

```powershell
$shouldCheckKeyboard = (Get-TimeSinceMs -startTime $LastSimulatedKeyPress) -ge 300
```

The script uses Right Alt (VK_RMENU = 0xA5) which is also skipped in keyboard scanning.

### Emoji Display Width Variations

Some emojis render as 1 column, others as 2. After writing an emoji, check the actual cursor position and fill gaps if needed:

```powershell
Write-Host $emoji -NoNewline
$actualX = [Console]::CursorPosition.X
if ($actualX -lt $expectedX) {
    Write-Host (" " * ($expectedX - $actualX)) -NoNewline -BackgroundColor $bg
}
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
2. **Diagnostics**: Run with `-Diag` for file-based logs in `$env:TEMP\mjig_diag\`
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
| `$env:TEMP\mjig_diag\startup.txt` | Initialization diagnostics |
| `$env:TEMP\mjig_diag\settle.txt` | Mouse settle detection logs |

No external dependencies - the script is fully self-contained.

---

## Quick Reference: Key Line Numbers

| Component | Approximate Lines |
|-----------|------------------|
| Parameters | 122-148 |
| Theme Colors | 214-289 |
| Box Characters | 203-212 |
| P/Invoke Types | 780-1050 |
| Get-MousePosition | 460-500 |
| Draw-ResizeLogo | 2824-2950 |
| Show-MovementModifyDialog | 1600-2400 |
| Show-QuitConfirmationDialog | 2400-2600 |
| Show-TimeChangeDialog | 2600-2800 |
| Main Loop Start | 4619 |
| Wait Loop | 4671-5400 |
| Resize Handling | 5600-5800 |
| UI Rendering | 6200-7200 |
| Menu Rendering | 7063-7350 |

*Note: Line numbers are approximate and may shift as code is modified.*
