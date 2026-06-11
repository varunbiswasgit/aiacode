# tetris-game — Test Plan

All tests are manual. Open `tetris-royal-3d.html` in a modern browser (Chrome, Edge, or Firefox) and follow each case.

---

## Environment

- Browser: Chrome 120+ / Edge 120+ / Firefox 120+
- No server required — open directly via `file://` or any web server
- For mobile tests: Chrome on Android / Safari on iOS, or DevTools device emulation

---

## Test Cases

### TC-01 — Start screen

| Field | Detail |
|-------|--------|
| Setup | Open `tetris-royal-3d.html` fresh |
| Action | Observe the overlay |
| Expected | Royal modal shows title "TETRIS", subtitle "Royal 3D Edition", and a ◆ Start Game button |
| Pass | Modal visible; board empty; no pieces moving |

---

### TC-02 — Game starts on button click

| Field | Detail |
|-------|--------|
| Setup | Start screen visible |
| Action | Click ◆ Start Game |
| Expected | Modal hides; a piece begins falling; Score/Level/Lines show 0/1/0; Pause button enables |
| Pass | Piece animates; HUD updates; Pause button no longer greyed out |

---

### TC-03 — Left / right movement

| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Press ← and → |
| Expected | Active piece moves one column per keypress; stops at left/right wall |
| Pass | No movement beyond column 0 or column 9 |

---

### TC-04 — Rotation with wall kick

| Field | Detail |
|-------|--------|
| Setup | Move an I-piece to the far right wall |
| Action | Press ↑ |
| Expected | Piece rotates; if blocked at wall, it kicks left to fit |
| Pass | Rotation completes without clipping through wall |

---

### TC-05 — Soft drop scoring

| Field | Detail |
|-------|--------|
| Setup | Game running; note current score |
| Action | Hold ↓ for several cells |
| Expected | Score increases by 1 per cell dropped |
| Pass | Score delta equals number of cells soft-dropped |

---

### TC-06 — Hard drop scoring

| Field | Detail |
|-------|--------|
| Setup | Game running; note current score |
| Action | Press Space |
| Expected | Piece drops instantly; score increases by 2 × distance dropped |
| Pass | Score delta = 2 × cells dropped; piece locks immediately |

---

### TC-07 — Line clear and score

| Field | Detail |
|-------|--------|
| Setup | Manually fill one complete row (use hard drops) |
| Action | Complete the row |
| Expected | Row disappears; rows above shift down; score increases by 100 × level |
| Pass | Board collapses correctly; score updates |

---

### TC-08 — Tetris (4-line clear)

| Field | Detail |
|-------|--------|
| Setup | Stack pieces leaving a single-column gap on the right |
| Action | Drop an I-piece vertically into the gap |
| Expected | 4 rows clear simultaneously; score increases by 800 × level |
| Pass | All 4 rows removed; score delta = 800 × level |

---

### TC-09 — Level progression

| Field | Detail |
|-------|--------|
| Setup | Game running at level 1 |
| Action | Clear 10 lines |
| Expected | Level counter increments to 2; piece drop speed noticeably increases |
| Pass | Level = 2 shown in HUD; visual speed increase confirmed |

---

### TC-10 — Pause and resume

| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Press P; wait 3 seconds; press P again |
| Expected | Piece freezes on P; "PAUSED" text overlaid on canvas in gold; resumes on second P; no position drift |
| Pass | Game state identical before and after pause; PAUSED overlay visible |

---

### TC-11 — Ghost piece

| Field | Detail |
|-------|--------|
| Setup | Game running; board partially filled |
| Action | Move piece left and right |
| Expected | Transparent ghost shows projected landing position and updates with each move |
| Pass | Ghost aligns with where piece locks when dropped |

---

### TC-12 — Game over

| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Stack pieces until they reach the top |
| Expected | Royal Game Over modal appears showing final score; ▶ Play Again button resets the game |
| Pass | Modal shown; clicking Play Again resets score/level/lines and clears board |

---

### TC-13 — Responsive layout (desktop resize)

| Field | Detail |
|-------|--------|
| Setup | Game open in desktop browser |
| Action | Resize browser window from wide to narrow and back |
| Expected | Canvas resizes fluidly; block size recalculates; no overflow or scrollbar; side panel repositions below board on narrow widths |
| Pass | Board always fits within viewport; no horizontal scrollbar |

---

### TC-14 — Responsive layout (mobile / DevTools emulation)

| Field | Detail |
|-------|--------|
| Setup | Open in DevTools device emulation (e.g. iPhone 14, 390×844) |
| Action | Load page; observe layout |
| Expected | Canvas fills available height; side panel appears below board; Controls box hidden; all buttons visible |
| Pass | No clipping; board uses maximum available space |

---

### TC-15 — Touch rotate (tap)

| Field | Detail |
|-------|--------|
| Setup | Game running on mobile or DevTools touch emulation |
| Action | Tap the canvas (short press, no swipe) |
| Expected | Active piece rotates; wall-kick applies if needed |
| Pass | Piece rotates on each tap |

---

### TC-16 — Touch move (swipe left / right)

| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Swipe left and right across the canvas |
| Expected | Piece moves in the swipe direction; number of columns moved proportional to swipe distance |
| Pass | Piece moves without rotation; stops at walls |

---

### TC-17 — Touch hard drop (swipe up)

| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Swipe upward on the canvas |
| Expected | Piece hard-drops instantly to the lowest valid position and locks |
| Pass | Score increases by 2 × rows dropped; piece locks immediately |

---

### TC-18 — Touch soft drop (swipe down)

| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Swipe downward on the canvas |
| Expected | Piece soft-drops to the bottom and locks; score increases by 1 per cell |
| Pass | Piece reaches bottom; score increments correctly |

---

### TC-19 — Rules modal

| Field | Detail |
|-------|--------|
| Setup | Any state (game running or not) |
| Action | Click ◆ Rules button |
| Expected | Royal glassmorphism rules overlay opens covering the page; shows Controls, Touch, Scoring, Levelling sections |
| Pass | Overlay opens; ✓ Got It — Close button dismisses it; Esc key also closes it |

---

### TC-20 — Shell loader

| Field | Detail |
|-------|--------|
| Setup | Internet connection available |
| Action | Open `shell.html` in browser |
| Expected | Spinner shows briefly then game loads inside full-viewport iframe from GitHub raw URL |
| Pass | Game fully playable inside iframe; no error message shown |
