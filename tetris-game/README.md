# tetris-game

A self-contained single-file browser Tetris game built with HTML5 Canvas and vanilla JavaScript — Royal 3D Edition. No install or server required — open `tetris-royal-3d.html` in any modern browser.

---

## Features

- All 7 standard Tetrominoes (I, O, T, S, Z, J, L) with wall-kick rotation
- Ghost piece showing projected landing position
- Hard drop, soft drop, and pause
- Progressive levelling — speed increases every 10 lines cleared
- Next-piece preview panel
- Score, Level, and Lines HUD
- **Royal 3D visual theme** — deep navy/purple background, gold gradient title, opal-shift animated rainbow border, Cinzel + Raleway fonts
- **3D bevelled blocks** — each block has a colour glow layer, light top-left bevel, dark bottom-right shadow, and a shine spot
- **Fluid responsive canvas** — `syncSize()` calculates block size dynamically from viewport on load, resize, and orientation change; no hardcoded pixel dimensions
- **Touch controls** — tap to rotate, swipe left/right to move, swipe down to soft-drop, swipe up to hard-drop
- Side panel reflows below the board on narrow screens (≤520 px)
- In-game Rules modal (Royal glassmorphism style)
- **Dice Mode** — toggle between Classic and Dice mode before a game starts; toggle is disabled and greyed-out once a game is in progress
- **Member words** — a pool of 100 words loaded from `member-words.json` at runtime; used for in-game word overlays
- Shell loader (`shell.html`) fetches `tetris-royal-3d.html` from GitHub raw API at load time

---

## Controls

### Keyboard

| Key | Action |
|-----|--------|
| ← → | Move left / right |
| ↑ | Rotate (with wall kick) |
| ↓ | Soft drop (+1 pt/cell) |
| Space | Hard drop (+2 pt/cell) |
| P | Pause / resume |

### Touch / Mobile

| Gesture | Action |
|---------|--------|
| Tap | Rotate |
| Swipe left / right | Move piece |
| Swipe down | Soft drop |
| Swipe up | Hard drop |

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

Level increases every 10 lines. Drop speed starts at 1000 ms and decreases by 92 ms per level, floored at 80 ms.

---

## File Structure

```
tetris-game/
  tetris-royal-3d.html      — Full game (HTML + CSS + JS, single file, Royal 3D Edition)
  shell.html                — Thin iframe loader — fetches tetris-royal-3d.html from GitHub raw API
  member-words.json         — Pool of 100 words used by the game at runtime
  member-words-admin.html   — Admin UI to add/remove/import words and download the updated JSON
  scores.json               — Persisted high-score data
  README.md                 — This file
  TESTING.md                — Manual test cases
```

---

## Managing Member Words

Open `member-words-admin.html` in any browser (no server needed):

1. **Add** a word — type and press Enter or click Add Word
2. **Remove** a word — click the ✕ on any chip
3. **Import** — paste an existing `member-words.json` to load a new baseline
4. **Download** the updated `member-words.json`
5. Commit the downloaded file to `tetris-game/member-words.json` on `main`

The game fetches `member-words.json` live on every load — no rebuild needed.

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release — all 7 pieces, ghost piece, hard drop, levelling, HUD |
| v2 | Royal 3D Edition — fluid responsive canvas, 3D bevelled blocks, opal-shift border, Royal theme (Cinzel/Raleway), touch controls, rules modal, side panel reflow on mobile; renamed index.html → tetris-royal-3d.html |
| v3 | Shell loader; member-words.json pool; Dice Mode toggle (disabled mid-game); blob-URL fix for words fetch; member-words-admin.html admin page |
