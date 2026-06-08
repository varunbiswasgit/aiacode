# triangles-game

A two-player browser-based dice game built with vanilla HTML, CSS, and JavaScript. Players take turns rolling dice and connecting dots to form triangles on a 24-dot grid. The player who completes the most triangles wins.

---

## Files

| File | Description |
|------|-------------|
| `final_triangles_game.html` | Complete self-contained game (no external dependencies) |

---

## How to Play

1. Open `final_triangles_game.html` in any modern browser — no server or install required.
2. **Player 1** clicks **Roll Dice** to start their turn.
3. Use the moves granted by the dice roll to connect dots and form triangles.
4. The game ends automatically when no legal line can be drawn anywhere on the board.
5. Click **Finish Game** to end early at any point.
6. The player with the highest score wins. Click **New Game** or **Play Again** to reset.

---

## Configuration

| Setting | Default | Notes |
|---------|---------|-------|
| Players | 2 | Fixed at runtime |
| Dot grid | 24 dots | Fixed layout |
| Dice | Standard 1–6 | Determines moves per turn |

---

## Dependencies

None. Pure HTML/CSS/JavaScript — runs entirely in the browser.

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release — two-player triangles game with dice mechanic and 24-dot grid |
| v2 | Auto-detect game end via `anyLegalMoveExists()` (O(n²) dot-pair scan); comment unused `linesToD2` variable; fix duplicate-triangle check to be vertex-order-independent |
| v3 | Fix Player 2 turn loss — player switch now executes before board-exhaustion check |
