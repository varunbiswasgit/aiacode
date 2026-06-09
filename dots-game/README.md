# dots-game — Dots & Boxes: Royal Edition

A self-contained two-player browser game built in the same Royal Edition visual style as the `triangles-game`. No install, no server — open `dots-royal-edition.html` in any modern browser.

---

## Purpose

Classic Dots and Boxes on a 5×5 grid (16 claimable boxes). Players take turns selecting pairs of adjacent dots to draw lines. Completing the fourth side of a box claims it. The player with the most boxes at the end wins.

---

## How to Run

1. Open `dots-royal-edition.html` in any modern browser (Chrome, Edge, Firefox, Safari).
2. No dependencies, no internet connection required after page load (fonts load from Google Fonts on first open).

---

## Interaction Model

| Step | Action |
|------|--------|
| 1 | Tap a dot — it glows gold to confirm selection |
| 2 | Tap an adjacent dot (horizontal or vertical neighbour only) |
| Result | One line is drawn in the current player's colour |

- Tapping the same dot again deselects it.
- Tapping a non-adjacent dot moves the selection.
- An already-drawn line cannot be redrawn.

---

## Controls

| Button | Action |
|--------|--------|
| Rules | Opens the in-game rules overlay |
| Finish | Ends the current game and shows the winner based on boxes claimed so far |
| New Game | Resets the board, scores, and turn to Player 1 |
| Play Again | Available in the end-game modal — resets and starts fresh |

---

## Scoring

- Claiming a box earns 1 point and grants an extra turn.
- The player with more claimed boxes when all lines are drawn wins.
- If both players claim 8 boxes, the result is a draw.

---

## Visual Design

Shares the Royal Edition aesthetic from `triangles-game`:
- Cinzel + Raleway fonts
- Layered cosmic gradient background
- Opal-animated rainbow board border
- Player 1: blue (`#60b8ff`), Player 2: red (`#ff7070`)
- Gold-glow dot selection indicator
- Jewelled button styles: gold (Rules), purple (Finish), cyan (New Game), green (Play Again)

---

## Files

| File | Description |
|------|-------------|
| `dots-royal-edition.html` | Complete self-contained game |
| `README.md` | This file |
| `TESTING.md` | Manual test cases |

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release — 5×5 grid, dot-to-dot selection, box claiming, extra-turn logic, Rules / Finish / New Game controls, Royal Edition visual style |
