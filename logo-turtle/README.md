# logo-turtle

A self-contained browser-based LOGO turtle graphics interpreter built with HTML5 Canvas and vanilla JavaScript. No installation, no dependencies, no server — open either file in any modern browser and start drawing.

Two editions are available:

| File | Edition | Description |
|------|---------|-------------|
| `logo-turtle.html` | Standard | Clean light/dark interface with Nexus design system |
| `logo-turtle-3d.html` | Royal 3D | Velvet + gold theme, 3D panel depth effects, starfield background |

---

## Purpose

Replicates the classic LOGO turtle experience (as used in MS LOGO and similar educational tools) entirely in the browser. The turtle cursor draws lines on a canvas as it moves, controlled by typed commands in a command scroll (text editor).

---

## Usage

1. Open `logo-turtle.html` or `logo-turtle-3d.html` in any modern browser (Chrome, Edge, Firefox, Safari).
2. Type LOGO-style commands into the command scroll on the right.
3. Click **Run** to execute the program.
4. Click **Clear** to wipe the command scroll and start fresh.
5. Click **Reset** to clear the canvas and return the turtle to the centre.
6. Use the quick-access **example chips** (Square, Triangle, Circle, Flower, Octagon, Star) to load and run pre-written programs.

Alternatively, open `shell.html` to always load the latest version of `logo-turtle-3d.html` directly from GitHub on every page load.

---

## Button Reference

| Button | Action |
|--------|--------|
| **Run** | Executes all commands in the command scroll |
| **Clear** | Clears the command scroll (text editor) only — canvas is unchanged |
| **Reset** | Clears the canvas and returns the turtle to centre, heading 0° |

---

## Supported Commands

| Command | Aliases | Description |
|---------|---------|-------------|
| `fd N` | `forward N` | Move forward N steps, drawing a line if pen is down |
| `bk N` | `back N` | Move backward N steps |
| `rt N` | `right N` | Turn right N degrees |
| `lt N` | `left N` | Turn left N degrees |
| `pu` | `penup` | Lift the pen — turtle moves without drawing |
| `pd` | `pendown` | Lower the pen — turtle draws as it moves |
| `home` | — | Return turtle to centre, heading 0° |
| `cs` | `clearscreen` | Clear the canvas and reset turtle to centre |
| `repeat N [ ... ]` | — | Repeat the block of commands N times |

### Shapes quick reference

```logo
; Square
repeat 4 [fd 100 rt 90]

; Triangle
repeat 3 [fd 120 rt 120]

; Circle (approximated by 360 small steps)
repeat 360 [fd 1 rt 1]

; Flower
repeat 36 [fd 120 rt 170]

; Octagon
repeat 8 [fd 80 rt 45]

; Star
repeat 5 [fd 100 rt 144]
```

---

## HUD (heads-up display)

Below the canvas, a live status row shows:
- **X / Y** — turtle position relative to canvas centre
- **Heading** — current angle in degrees
- **Pen** — whether the pen is up or down

---

## Features

- Two editions: Standard and Royal 3D (velvet + gold, starfield background, 3D panel depth)
- Light and dark mode with a theme toggle button (top-right)
- Six quick-spell chips: Square, Triangle, Circle, Flower, Octagon, Star
- Accessible: keyboard-navigable, focus rings, `aria-live` status region, skip link
- Responsive layout — stacks to a single column on screens under 900 px
- Inline error messages displayed in the status bar for unknown commands or bad syntax

---

## Limitations

- No `to … end` procedure definitions in this version
- No colour or pen-width commands
- No `arc` command (circle is approximated via `repeat 360 [fd 1 rt 1]`)
- No file save / PNG export

---

## Files

| File | Description |
|------|-------------|
| `logo-turtle.html` | Standard edition |
| `logo-turtle-3d.html` | Royal 3D edition |
| `shell.html` | Thin loader — fetches latest `logo-turtle-3d.html` from GitHub on every page load |
| `README.md` | This file |
| `TESTING.md` | Manual test cases |

---

## Version history

| Version | Summary |
|---------|---------|
| v1 | Initial release — fd, bk, rt, lt, pu, pd, home, cs, repeat; light/dark mode; four example chips |
| v2 | Added Royal 3D Edition (`logo-turtle-3d.html`) — velvet/gold theme, starfield background, 3D panel depth effects, six quick-spell chips (added Octagon and Star) |
| v3 | Fixed Clear button in both editions — Clear now clears the command scroll (text editor) only; Reset handles canvas clear and turtle state reset |
| v4 | Added `shell.html` — thin loader that fetches latest `logo-turtle-3d.html` from GitHub raw on every page load; cache-busted with `?t=Date.now()` |
