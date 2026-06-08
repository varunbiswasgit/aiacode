# logo-turtle

A self-contained browser-based LOGO turtle graphics interpreter built with HTML5 Canvas and vanilla JavaScript. No installation, no dependencies, no server — open `logo-turtle.html` in any modern browser and start drawing.

---

## Purpose

Replicates the classic LOGO turtle experience (as used in MS LOGO and similar educational tools) entirely in the browser. The turtle cursor draws lines on a canvas as it moves, controlled by typed commands in a command box.

---

## Usage

1. Open `logo-turtle.html` in any modern browser (Chrome, Edge, Firefox, Safari).
2. Type LOGO-style commands into the command box on the right.
3. Click **Run** to execute.
4. Use the quick-access **example chips** (Square, Triangle, Circle, Flower) to load and run pre-written programs.

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
```

---

## HUD (heads-up display)

Below the canvas, a live status row shows:
- **X / Y** — turtle position relative to canvas centre
- **Heading** — current angle in degrees
- **Pen** — whether the pen is up or down

---

## Features

- Light and dark mode with a theme toggle button (top-right)
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

## Version history

| Version | Summary |
|---------|---------|
| v1 | Initial release — fd, bk, rt, lt, pu, pd, home, cs, repeat; light/dark mode; four example chips |
