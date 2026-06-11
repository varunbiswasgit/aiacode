# tetris-game — Test Plan

All tests are manual. Open `index.html` in a modern browser (Chrome, Edge, or Firefox) and follow each case.

---

## Environment

- Browser: Chrome 120+ / Edge 120+ / Firefox 120+
- No server required
- No dependencies

---

## Test Cases

### TC-01 — Start screen

| Field | Detail |
|-------|--------|
| Setup | Open `index.html` fresh |
| Action | Observe the overlay |
| Expected | Overlay shows title "TETRIS", subtitle, and a Start button |
| Pass | Overlay visible; board empty; no pieces moving |

---

### TC-02 — Game starts on button click

| Field | Detail |
|-------|--------|
| Setup | Start screen visible |
| Action | Click Start |
| Expected | Overlay hides; a piece begins falling; Score/Level/Lines show 0/1/0 |
| Pass | Piece animates; HUD updates |

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
| Expected | Piece freezes on P; resumes movement on second P; no position drift |
| Pass | Game state identical before and after pause |

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
| Expected | Game over overlay appears showing final score; Start button resets the game |
| Pass | Overlay shown; clicking Start resets score/level/lines and clears board |
