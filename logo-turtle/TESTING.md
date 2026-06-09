# logo-turtle — Test Plan

## Test environment

- **Browser:** Any modern browser (Chrome 120+, Edge 120+, Firefox 121+, Safari 17+)
- **Files:** `logo-turtle.html` (Standard) and `logo-turtle-3d.html` (Royal 3D) — open directly from disk, no server required
- **Screen sizes tested:** Desktop (1280 px wide), mobile (375 px wide)
- **Applies to:** Both editions unless a test case specifies one file

---

## Manual test cases

### TC-01 — Default load

| Field | Detail |
|-------|--------|
| **Setup** | Open the file in a browser |
| **Action** | Observe the page on load |
| **Expected** | Canvas shows the default program output. Turtle arrow visible. Status bar reads `Program completed.` (Standard) or `✦ Program completed with honour. ✦` (Royal 3D) |
| **Pass criteria** | Default drawing rendered; no error message shown |

---

### TC-02 — `fd` and `rt` (straight line and turn)

| Field | Detail |
|-------|--------|
| **Setup** | Clear the command scroll |
| **Action** | Type `fd 100 rt 90 fd 100`, click **Run** |
| **Expected** | Two perpendicular lines drawn from centre. Turtle arrow is at the end of the second line, heading down |
| **Pass criteria** | Two lines visible; HUD shows X ≈ 100, Y ≈ −100 |

---

### TC-03 — `repeat` draws a square

| Field | Detail |
|-------|--------|
| **Setup** | Clear the command scroll |
| **Action** | Type `repeat 4 [fd 100 rt 90]`, click **Run** |
| **Expected** | A closed square drawn from the centre |
| **Pass criteria** | Turtle returns to starting position; HUD heading = 0° |

---

### TC-04 — Circle approximation

| Field | Detail |
|-------|--------|
| **Setup** | Clear the command scroll |
| **Action** | Click the **Circle** chip |
| **Expected** | A circle drawn from the centre; status bar reads completed message |
| **Pass criteria** | Visually circular shape; no error in status bar |

---

### TC-05 — Pen up / pen down

| Field | Detail |
|-------|--------|
| **Setup** | Clear the command scroll |
| **Action** | Type `fd 50 pu fd 50 pd fd 50`, click **Run** |
| **Expected** | Two separate line segments with a gap in the middle |
| **Pass criteria** | Gap is visible; HUD shows Pen: down after the program |

---

### TC-06 — `home` resets position

| Field | Detail |
|-------|--------|
| **Setup** | Clear the command scroll |
| **Action** | Type `fd 200 rt 45 home`, click **Run** |
| **Expected** | Line drawn, then turtle snaps back to centre heading 0° |
| **Pass criteria** | HUD shows X: 0, Y: 0, Heading: 0° |

---

### TC-07 — `cs` clears the canvas mid-program

| Field | Detail |
|-------|--------|
| **Setup** | Run any program so lines are on screen |
| **Action** | Add `cs` as the first line and re-run |
| **Expected** | Canvas is blank before the remaining commands execute |
| **Pass criteria** | No residual lines from the previous run visible |

---

### TC-08 — Clear button clears the command scroll only

| Field | Detail |
|-------|--------|
| **Setup** | Run any program so lines are on canvas and commands are in the scroll |
| **Action** | Click **Clear** |
| **Expected** | Command scroll (text editor) is emptied. Canvas drawing remains unchanged. Status bar updates. |
| **Pass criteria** | Text editor is blank; canvas lines still visible; turtle arrow unchanged |

---

### TC-09 — Reset button clears canvas and resets turtle

| Field | Detail |
|-------|--------|
| **Setup** | Run a program that moves the turtle away from centre |
| **Action** | Click **Reset** |
| **Expected** | Canvas clears; turtle returns to centre; HUD shows X: 0, Y: 0, Heading: 0°; command scroll is untouched |
| **Pass criteria** | All HUD values reset; turtle arrow at canvas centre; text editor content unchanged |

---

### TC-10 — Clear then Run on empty scroll

| Field | Detail |
|-------|--------|
| **Setup** | Run a program, then click **Clear** |
| **Action** | Click **Run** with the empty scroll |
| **Expected** | No crash. Canvas resets to blank with turtle at centre. Status bar shows completed message. |
| **Pass criteria** | No JavaScript error; page remains functional |

---

### TC-11 — Unknown command error handling

| Field | Detail |
|-------|--------|
| **Setup** | Clear the command scroll |
| **Action** | Type `fly 100`, click **Run** |
| **Expected** | Status bar shows an error message: `Unknown command: "fly"` |
| **Pass criteria** | No crash; canvas cleared; error message visible |

---

### TC-12 — Dark mode toggle

| Field | Detail |
|-------|--------|
| **Setup** | Page loaded in light mode |
| **Action** | Click the theme toggle button (top-right) |
| **Expected** | Page switches to dark mode; canvas background and text colours update; turtle redraws |
| **Pass criteria** | Dark surface colours visible; no layout breakage; program re-runs correctly |

---

### TC-13 — Mobile layout (375 px)

| Field | Detail |
|-------|--------|
| **Setup** | Open DevTools, set viewport to 375 px wide |
| **Action** | Load the page and run a program |
| **Expected** | Layout stacks to single column; canvas fills full width; all buttons remain tappable |
| **Pass criteria** | No horizontal overflow; canvas and editor both visible without zooming |

---

### TC-14 — Keyboard navigation

| Field | Detail |
|-------|--------|
| **Setup** | Open the page |
| **Action** | Tab through all interactive elements |
| **Expected** | Focus rings visible on all buttons, textarea, and chips |
| **Pass criteria** | Every control is reachable and activatable via keyboard alone |

---

### TC-15 — Royal 3D Edition: Quick-spell chips (Octagon and Star)

| Field | Detail |
|-------|--------|
| **File** | `logo-turtle-3d.html` only |
| **Setup** | Open the Royal 3D edition |
| **Action** | Click the **Octagon** chip, then the **Star** chip |
| **Expected** | Each chip loads its program into the command scroll and runs it immediately |
| **Pass criteria** | Correct geometric shape drawn for each chip; no errors in status bar |
