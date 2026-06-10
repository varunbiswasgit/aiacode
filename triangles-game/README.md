# triangles-game

A two-player browser-based dice game built with vanilla HTML, CSS, and JavaScript. Players take turns rolling dice and connecting dots to form triangles on a 24-dot grid. The player who completes the most triangles wins.

---

## Files

| File | Description |
|------|-------------|
| `final_triangles_game.html` | Complete self-contained game (no external dependencies) |
| `shell.html` | Thin loader — fetches latest `final_triangles_game.html` from GitHub on every page load |

---

## How to Play

1. Open `final_triangles_game.html` in any modern browser — no server or install required.
2. **Player 1** clicks **Roll Dice** to start their turn.
3. Use the moves granted by the dice roll to connect dots and form triangles.
4. The game ends automatically when no legal line can be drawn anywhere on the board.
5. Click **Finish Game** to end early at any point.
6. The player with the highest score wins. Click **New Game** or **Play Again** to reset.

Alternatively, open `shell.html` to always load the latest version of `final_triangles_game.html` directly from GitHub on every page load.

---

## Device Support

| Device / Orientation | Canvas size |
|----------------------|-------------|
| iPhone portrait | ~355px — full width |
| iPhone landscape | ~height-constrained, ~230px |
| iPad portrait | 500px (capped) |
| iPad landscape | 500px (capped) |
| Desktop | 500px (capped) |

Canvas resizes automatically on `resize` and `orientationchange` events. Dot positions are stored as normalised ratios (0–1) and projected to pixel coordinates at draw time, so all geometry scales correctly.

---

## Configuration

| Setting | Default | Notes |
|---------|---------|-------|
| Players | 2 | Fixed at runtime |
| Dot grid | 24 dots | Fixed layout, scales with canvas |
| Dice | Standard 1–6 | Determines moves per turn |
| Canvas cap | 500px | Prevents oversizing on large screens |

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
| v4 | Responsive canvas — dots stored as normalised ratios; canvas resizes on resize/orientationchange; tap radius scales proportionally; supports portrait, landscape, iPad, iPhone |
| v5 | Added `shell.html` — thin loader that fetches latest `final_triangles_game.html` from GitHub raw on every page load; cache-busted with `?t=Date.now()` |
