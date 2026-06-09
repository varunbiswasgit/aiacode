# TESTING — dots-game

All tests are manual. Open `dots-royal-edition.html` in a modern browser before running each case.

---

## Environment

| Item | Value |
|------|-------|
| Browser | Chrome 124+, Edge 124+, Firefox 125+, Safari 17+ |
| Device | Desktop and mobile (touch) |
| Dependencies | None (offline after first font load) |

---

## Test Cases

### TC-01 — New game initialises correctly

**Setup:** Open the file in a browser.

**Action:** Observe the initial state.

**Expected:**
- 5×5 grid of 25 dots rendered.
- All segment placeholders shown as faint white lines.
- Scoreboard shows Player 1: 0, Player 2: 0, Boxes Left: 16.
- Turn line reads "Player 1's turn".

**Pass criteria:** All scores zero, 16 boxes remaining, no coloured lines visible.

---

### TC-02 — First dot tap selects and highlights

**Setup:** TC-01 complete.

**Action:** Tap any dot.

**Expected:**
- Tapped dot glows gold.
- No line is drawn.
- No score change.

**Pass criteria:** Exactly one dot highlighted gold; board otherwise unchanged.

---

### TC-03 — Second tap on adjacent dot draws one line

**Setup:** TC-02 complete (one dot selected).

**Action:** Tap a horizontally or vertically adjacent dot.

**Expected:**
- One line drawn in Player 1's blue colour between the two dots.
- Gold glow removed from first dot.
- Turn switches to Player 2.

**Pass criteria:** Exactly one blue line visible; turn indicator shows Player 2.

---

### TC-04 — Diagonal or non-adjacent tap moves selection

**Setup:** One dot selected (gold glow).

**Action:** Tap a dot that is not a direct neighbour.

**Expected:**
- Selection moves to the newly tapped dot.
- No line drawn.
- Turn does not change.

**Pass criteria:** Gold glow moves; no new line; same player's turn.

---

### TC-05 — Tapping selected dot deselects

**Setup:** One dot selected (gold glow).

**Action:** Tap the same dot again.

**Expected:**
- Gold glow removed.
- No line drawn.
- Turn does not change.

**Pass criteria:** No dot glowing; board unchanged.

---

### TC-06 — Completing a box claims it and grants extra turn

**Setup:** Three sides of one box already drawn by Player 1.

**Action:** Player 1 draws the fourth side.

**Expected:**
- Box fills with blue tint.
- Player 1 score increments by 1.
- Boxes Left decrements by 1.
- Turn remains with Player 1.

**Pass criteria:** Score +1 for Player 1; turn indicator still shows Player 1.

---

### TC-07 — Finish button ends game early

**Setup:** Game in progress with at least one box claimed.

**Action:** Click Finish.

**Expected:**
- Game-over modal appears.
- Winner declared based on current box counts.
- Play Again button visible.

**Pass criteria:** Modal shown with correct winner or draw message.

---

### TC-08 — New Game resets everything

**Setup:** Mid-game state with lines and boxes claimed.

**Action:** Click New Game.

**Expected:**
- All lines cleared.
- All box fills removed.
- Scores reset to 0.
- Boxes Left reset to 16.
- Turn reset to Player 1.
- No modal visible.

**Pass criteria:** Board visually identical to TC-01 initial state.

---

### TC-09 — All lines drawn triggers auto game-over

**Setup:** New game.

**Action:** Play until every segment between dots is claimed.

**Expected:**
- Game-over modal appears automatically.
- Winner or draw declared correctly based on box counts.

**Pass criteria:** Modal appears without clicking Finish; scores sum to 16.

---

### TC-10 — Touch interaction works on mobile

**Setup:** Open on a touch-screen device or browser DevTools mobile emulation.

**Action:** Play several turns using finger taps.

**Expected:**
- Dot selection and line drawing behave identically to mouse clicks.
- No ghost lines or double-draw on corner taps.

**Pass criteria:** No visual glitches; correct turn and score progression.

---

### TC-11 — Rules modal opens and closes

**Setup:** New game.

**Action:** Click Rules, then click Close.

**Expected:**
- Rules overlay appears on click.
- Overlay dismisses on Close.
- Game state unchanged.

**Pass criteria:** Modal toggles correctly; no score or turn change.
