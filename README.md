# mJig 🐀

A feature-rich PowerShell mouse jiggler with a console-based TUI, designed to keep your system active with natural-looking mouse movements and intelligent user input detection.

## Features

- **Smart Mouse Movement**: Randomized cursor movements with configurable distance, speed, and variance
- **User Input Detection**: Automatically pauses when you're actively using mouse/keyboard
- **Auto-Resume Delay**: Configurable cooldown timer after user input before resuming automation
- **Scheduled Stop Time**: Set a specific end time with optional variance for natural patterns
- **Multiple View Modes**: Full, minimal, or hidden interface
- **Interactive Dialogs**: Modify settings on-the-fly without restarting
- **Mouse Stutter Prevention**: Waits for mouse to settle before starting next movement cycle
- **Window Resize Handling**: Beautiful centered logo with playful quotes during resize
- **Themeable UI**: Centralized color variables for easy customization
- **Stats Box**: Real-time display of detected input categories (Mouse, Keyboard, mouse buttons, Scroll/Other) and movement statistics
- **Click Support**: Mouse-clickable menu buttons and dialog interactions
- **Flicker-Free Rendering**: VT100/ANSI escape sequence rendering with atomic single-write frame output

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Windows OS (uses Win32 API for mouse/keyboard interaction)

## Installation

1. Download `start-mjig.ps1`
2. Ensure PowerShell execution policy allows scripts:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

### Basic Usage

```powershell
# Run with defaults (no end time, minimal view)
.\start-mjig.ps1
Start-mJig

# Run with full interface
Start-mJig -Output full

# Run until 5:30 PM
Start-mJig -EndTime 1730

# Run hidden in background
Start-mJig -Output hidden

# Debugging one-liner: launch in a new PowerShell session with full output and debug mode
$mJig = "C:\Path\To\start-mjig.ps1"
powershell -Command ". `"$mJig`"; Start-mJig -Output full -DebugMode"
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Output` | string | `"min"` | View mode: `full`, `min`, or `hidden` |
| `-EndTime` | string | `"0"` | Stop time in 24hr format (e.g., `1730` for 5:30 PM). `0` = no end time |
| `-EndVariance` | int | `0` | Random variance in minutes for end time |
| `-IntervalSeconds` | double | `2` | Base interval between movement cycles |
| `-IntervalVariance` | double | `2` | Random variance for interval timing |
| `-MoveSpeed` | double | `0.5` | Movement animation duration in seconds |
| `-MoveVariance` | double | `0.2` | Random variance for movement speed |
| `-TravelDistance` | double | `100` | Base cursor travel distance in pixels |
| `-TravelVariance` | double | `5` | Random variance for travel distance |
| `-AutoResumeDelaySeconds` | double | `0` | Cooldown after user input before resuming |
| `-DebugMode` | switch | `$false` | Enable debug logging |
| `-Diag` | switch | `$false` | Enable diagnostic file output |

### Interactive Controls

While running, use these keyboard shortcuts:

| Key | Action |
|-----|--------|
| `q` | Open quit confirmation dialog |
| `t` | Set/change end time |
| `v` | Toggle between full/min view |
| `h` | Toggle hidden mode |
| `m` | Open movement settings dialog (full view only) |

You can also click menu buttons with your mouse.

### Dialogs

**Modify Movement Settings** (`m` key in full view):
- Interval timing and variance
- Travel distance and variance  
- Movement speed and variance
- Auto-resume delay timer

**Set End Time** (`t` key):
- Enter time in HHmm format (e.g., 1730 for 5:30 PM)

**Quit Confirmation** (`q` key):
- Displays runtime statistics before exiting

## View Modes

### Full Mode
- Header with logo, current/end times, view indicator
- Activity log with timestamped entries
- Stats box showing detected inputs
- Interactive menu bar with icons

### Minimal Mode
- Header with essential info
- Menu bar only (no log or stats)

### Hidden Mode
- Minimal display: status line and clickable `(h)` button in bottom-right corner
- Hotkeys still functional
- Click `(h)` or press `h` to return to previous view
- Perfect for background operation

## Configuration

Movement and timing can be adjusted via:
1. Command-line parameters at startup
2. The "Modify Movement" dialog during runtime (`m` key)

### Theme Customization

Colors are defined as `$script:` variables in the Theme Colors section (around line 214). Groups include:
- Menu bar colors
- Header colors
- Stats box colors
- Dialog colors (Quit, Time, Movement)
- Resize screen colors
- General UI colors

## How It Works

1. **Movement Cycle**: At each interval, the script:
   - Checks if user is actively moving the mouse
   - Waits for mouse to "settle" (stop moving) if needed
   - Moves cursor a random distance in a random direction
   - Sends a simulated Right Alt keypress (non-intrusive)

2. **Input Detection**: Monitors user activity using multiple mechanisms:
   - `PeekConsoleInput` for keyboard and scroll wheel events (console-focused)
   - `GetLastInputInfo` for system-wide activity detection (passive)
   - `GetAsyncKeyState` for mouse button clicks (VK 0x01-0x06)
   - Position polling for mouse movement

3. **Stutter Prevention**: Before each movement cycle, verifies the mouse has been stationary for a brief period to avoid interfering with user actions

## Diagnostics

Enable with `-Diag` flag. Creates log files in `_diag/` (same directory as the script):
- `startup.txt` - Initialization diagnostics
- `settle.txt` - Mouse settle detection logs
- `input.txt` - Input detection logs (PeekConsoleInput + GetLastInputInfo)

## License

Shield: [![CC BY-ND 4.0][cc-by-nd-shield]][cc-by-nd]

This work is licensed under a
[Creative Commons Attribution-NoDerivs 4.0 International License][cc-by-nd].

[![CC BY-ND 4.0][cc-by-nd-image]][cc-by-nd]

[cc-by-nd]: https://creativecommons.org/licenses/by-nd/4.0/
[cc-by-nd-image]: https://licensebuttons.net/l/by-nd/4.0/88x31.png
[cc-by-nd-shield]: https://img.shields.io/badge/License-CC%20BY--ND%204.0-lightgrey.svg
