# TESTING — triangles-game

All tests are manual. Open `final_triangles_game.html` in a browser to run each case.

---

## Test Environment

| Item | Requirement |
|------|-------------|
| Browser | Chrome 120+, Firefox 120+, Edge 120+, or Safari 17+ |
| Server | None required — open file directly |
| Dependencies | None |

---

## Test Cases

### TC-01 — Initial Load

| Field | Value |
|-------|-------|
| **Setup** | Open `final_triangles_game.html` fresh |
| **Action** | Observe the page |
| **Expected** | Score shows `Player 1: 0 \| Player 2: 0`; turn indicator shows Player 1; **Roll Dice** button enabled |
| **Pass criteria** | All three conditions met with no console errors |

---

### TC-02 — Dice Roll

| Field | Value |
|-------|-------|
| **Setup** | Game loaded, Player 1's turn |
| **Action** | Click **Roll Dice** |
| **Expected** | Moves left counter updates to a value between 1 and 6; **Roll Dice** button disables |
| **Pass criteria** | Counter changes; no duplicate rolls possible until turn ends |

---

### TC-03 — Turn Alternation

| Field | Value |
|-------|-------|
| **Setup** | Player 1 rolls dice; moves left = 0 |
| **Action** | Exhaust all moves |
| **Expected** | Turn indicator switches to Player 2 |
| **Pass criteria** | Indicator reads "Player 2's turn" |

---

### TC-04 — Score Increment

| Field | Value |
|-------|-------|
| **Setup** | Active game, valid triangle completable |
| **Action** | Connect the three dots that form a triangle |
| **Expected** | Active player's score increments by 1 |
| **Pass criteria** | Score display updates immediately; opposing player's score unchanged |

---

### TC-05 — Finish Game (Early)

| Field | Value |
|-------|-------|
| **Setup** | Game in progress |
| **Action** | Click **Finish Game** |
| **Expected** | Game ends; winner announced based on current scores |
| **Pass criteria** | Winner message displayed; game state locked |

---

### TC-06 — New Game / Play Again

| Field | Value |
|-------|-------|
| **Setup** | Game ended |
| **Action** | Click **New Game** or **Play Again** |
| **Expected** | Scores reset to 0; board clears; Player 1's turn |
| **Pass criteria** | All state reset; no residual dot connections or scores |

---

### TC-07 — Two-Player Win Declaration

| Field | Value |
|-------|-------|
| **Setup** | Play to completion — Player 1 more triangles |
| **Action** | Allow game to end naturally |
| **Expected** | "Player 1 wins!" displayed |
| **Pass criteria** | Winner matches player with higher score |

---

## Known Limitations

- Single-device, pass-and-play only (no network multiplayer).
- No persistent score tracking across sessions.
