# tetris-game

A self-contained single-file browser Tetris game built with HTML5 Canvas and vanilla JavaScript. No install or server required — open `index.html` in any modern browser.

---

## Features

- All 7 standard Tetrominoes (I, O, T, S, Z, J, L) with wall-kick rotation
- Ghost piece showing projected landing position
- Hard drop, soft drop, and pause
- Progressive leveling — speed increases every 10 lines cleared
- Next-piece preview panel
- Score, Level, and Lines HUD
- Dark navy / red premium UI consistent with other games in this repo

---

## Controls

| Key | Action |
|-----|--------|
| ← → | Move left / right |
| ↑ | Rotate (with wall kick) |
| ↓ | Soft drop (+1 pt/cell) |
| Space | Hard drop (+2 pt/cell) |
| P | Pause / resume |

---

## Scoring

| Lines cleared | Points (× level) |
|---------------|------------------|
| 1 | 100 |
| 2 | 300 |
| 3 | 500 |
| 4 (Tetris) | 800 |
| Soft drop | +1 per cell |
| Hard drop | +2 per cell |

Level increases every 10 lines. Drop speed starts at 1000 ms and decreases by 90 ms per level, capped at 100 ms.

---

## File Structure

```
tetris-game/
  index.html    — Full game (HTML + CSS + JS, single file)
  README.md     — This file
  TESTING.md    — Manual test cases
```

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release — all 7 pieces, ghost piece, hard drop, leveling, HUD |
