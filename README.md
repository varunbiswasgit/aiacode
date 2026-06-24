# aiacode — AI-Assisted Automation Scripts

A curated collection of VBA macros, PowerShell, Python scripts, and browser-based tools developed with AI assistance. All scripts are tested and documented; each tool folder contains its own README and TESTING file.

---

## ⚠️ AI Contributor Policy — README Updates

**This applies to any AI agent, Copilot, or automated tool pushing changes to this repository.**

Whenever you modify a script or any file in a tool folder, you **must** also update all of the following files in the same commit or PR:

| File | Location | What it covers |
|------|----------|----------------|
| Global README | [`README.md`](README.md) | Script table, descriptions, repository overview |
| Per-tool README | `<tool-folder>/README.md` | Purpose, fields, logic, configuration, version history |
| Per-tool TESTING | `<tool-folder>/TESTING.md` | Manual test cases, pass criteria, test environment |

### What triggers a README update

Any of the following changes require all files above to be reviewed and updated:

- Adding, removing, or renaming a field in a config structure or input parameter
- Adding, changing, or removing a function, launch strategy, or fallback step
- Changing default configuration values
- Changing what is logged, how errors are handled, or what output is produced
- Any change that affects documented test cases
- Adding a new script to the repository

### Minimum update requirement per file

- **Global README** — update the script description in the Scripts table if behaviour changed.
- **Per-tool README** — update the relevant section (fields, logic flow, configuration). Always add a version history entry.
- **Per-tool TESTING** — add, update, or remove test cases to match the new behaviour.

### Version history

Every code change must add a new row to the version history table in the per-tool README. Do not edit existing rows. Format:

```
| vN | One-line summary of what changed |
```

---

## How This Repository Is Organised

Each tool lives in its own self-contained folder. A folder holds the source script(s), a `README.md` explaining the tool, and a `TESTING.md` with test cases. There are no shared `scripts/`, `tests/`, or `README/` folders — everything a contributor needs for a given tool is in that tool's folder.

---

## File Types

| Extension | What it is |
|-----------|------------|
| `.bas` | VBA module exported from an Office application (Word, Excel, or Outlook). Import into the VBA editor (`Alt+F11`) to use. |
| `.ps1` | PowerShell script. Run from a PowerShell terminal or triggered by a VBA launcher. |
| `.py` | Python script. Requires Python 3 and dependencies listed in the tool's README. |
| `.html` | Self-contained browser tool or game. Open directly in any modern browser — no server required. |
| `README.md` | Human-readable documentation for the tool — purpose, usage, configuration, and version history. |
| `TESTING.md` | Test plan for the tool — manual test cases, automated test instructions, and pass criteria. |

---

## HTML Projects — Structure & Contributors Rule

**All HTML projects (games and browser tools) follow the same two-file pattern. This rule applies to all contributors, including AI agents.**

### Required files in every HTML project folder

| File | Purpose |
|------|---------|
| `index.html` | **Core game/tool** — complete HTML + CSS + JavaScript in a single self-contained file. Fully functional when opened directly in a browser. |
| `shell.html` | **Shell loader** — thin iframe loader that fetches `index.html` from the GitHub raw API on page open and injects it into a full-viewport iframe. No auto-refresh. |
| `README.md` | Purpose, features, controls, file structure, and version history. |
| `TESTING.md` | Step-by-step manual test cases with Setup, Action, Expected, and Pass columns. |

### Creating a new HTML project subfolder

1. **Create the folder** with a descriptive kebab-case name (e.g. `my-new-game/`).
2. **Write `index.html`** — single-file, no external CDN dependencies, works via `file://` or any web server.
3. **Write `shell.html`** — copy the loader pattern from [`triangles-game/shell.html`](triangles-game/shell.html) and update:
   - `<title>` — your game/tool name
   - `border-top-color` on `.spinner` — your accent colour (e.g. `#e94560`)
   - `iframe title` attribute — your game/tool name
   - `SOURCE` constant — raw GitHub URL pointing to your folder's `index.html`:
     ```
     https://raw.githubusercontent.com/varunbiswasgit/aiacode/main/YOUR_FOLDER_NAME/index.html
     ```
4. **Write `README.md`** — include features, controls (if a game), scoring (if relevant), file structure table, and a version history table starting at v1.
5. **Write `TESTING.md`** — minimum one test case per distinct behaviour; use the four-column format (Setup / Action / Expected / Pass).
6. **Update this file** — add a row to the Scripts table below with folder name, type `HTML / JavaScript`, and a one-line description.

### Updating an existing HTML project

| What changed | Files to update |
|---|---|
| `index.html` content or logic | `README.md` (version history row), `TESTING.md` (affected test cases), global `README.md` (description if behaviour changed) |
| `shell.html` loader | `README.md` (version history row) |
| New feature added | All of the above |

### Running HTML projects

| File | How to run |
|------|------------|
| `index.html` | Open directly in any browser — works via `file://` or any web server |
| `shell.html` | Open in any browser with an internet connection — fetches the latest `index.html` from GitHub raw API at load time |

---

## HTML Projects — Royal Edition Design System

All HTML projects in this repository that carry the **Royal Edition** visual theme share a single design system defined below. When making any CSS change to a Royal Edition file — or creating a new one — every token and pattern listed here must be applied consistently across **all** in-scope files.

> **Scope:** `dots-game/dots-royal-edition.html`, `triangles-game/final_triangles_game.html`, `tetris-game/tetris-royal-3d.html`, `worldcup2026-scores/worldcup2026_scores.html`
>
> **Exempt:** `logo-turtle/` — uses a gold/velvet sub-theme intentionally distinct from the cosmic-purple palette. Do not sync logo-turtle to the tokens below.

---

### Fonts

Every Royal Edition file must load exactly these two Google Fonts families and include both `preconnect` hints:

```html
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700&family=Raleway:wght@400;500;600&display=swap" rel="stylesheet">
```

| Role | Family | Weight |
|------|--------|--------|
| Headings, labels, buttons, HUD | `'Cinzel', serif` | 600 or 700 |
| Body copy, status lines, descriptions | `'Raleway', sans-serif` | 400, 500, or 600 |

---

### Background

The `body` background is always the fixed cosmic-purple radial stack:

```css
body {
  background:
    radial-gradient(ellipse at 20% 10%, #1a0a3a 0%, transparent 55%),
    radial-gradient(ellipse at 80% 90%, #0a1a3a 0%, transparent 55%),
    radial-gradient(ellipse at 60% 40%, #1a1040 0%, transparent 60%),
    linear-gradient(160deg, #0d0820 0%, #12103a 40%, #0a1828 100%);
  background-attachment: fixed;
}
```

The `body::before` pseudo-element overlays a fixed noise texture at z-index 0:

```css
body::before {
  content: '';
  position: fixed; inset: 0;
  background:
    repeating-linear-gradient(105deg, transparent 0px, transparent 80px,
      rgba(180,140,255,0.03) 80px, rgba(180,140,255,0.03) 81px),
    repeating-linear-gradient(15deg, transparent 0px, transparent 120px,
      rgba(100,220,255,0.02) 120px, rgba(100,220,255,0.02) 121px);
  pointer-events: none; z-index: 0;
}
```

---

### Layout Tokens

These values are the canonical current tag and must not drift between files:

| Token | Value | CSS property |
|-------|-------|--------------|
| Body padding | `10px 8px 14px` | `body { padding }` |
| UI section bottom gap | `8px` | `#ui { margin-bottom }` |
| H1 font size | `clamp(1.2rem, 4vw, 2rem)` | `h1 { font-size }` |
| H1 bottom gap | `7px` | `h1 { margin-bottom }` |
| Scoreboard / HUD padding | `6px 18px` | `#scoreboard, #hud { padding }` |
| Scoreboard / HUD gap | `10px` | `#scoreboard, #hud { gap }` |
| Scoreboard / HUD font-size | `clamp(0.75rem, 2vw, 0.95rem)` | `#scoreboard, #hud { font-size }` |
| Button padding | `8px 16px` | `button { padding }` |
| Button border-radius | `10px` | `button { border-radius }` |
| Canvas-wrap border-radius | `18px` | `#canvas-wrap { border-radius }` |
| Canvas inner border-radius | `15px` | `canvas { border-radius }` |

---

### Colour Palette

| Role | Value | Usage |
|------|-------|-------|
| Gold gradient (h1, scores) | `linear-gradient(135deg, #f5e083 0%, #e8c84a 30%, #fff8d0 55%, #c8960a 80%, #f5e083 100%)` | `h1`, HUD values, score display |
| Player 1 blue | `#60b8ff` | Player 1 labels, lines, fills |
| Player 2 red | `#ff7070` | Player 2 labels, lines, fills |
| Purple accent | `#c084fc` | Stage badges, section headings, panel titles |
| Green accent | `#4ade80` | Live status, roll prompts, bonus notes |
| Gold accent | `#f5e083` | HUD values, timestamps, level indicator |
| Body text | `#e8dfc8` | All readable body copy |
| Muted text | `rgba(200,190,230,0.7)` | Labels, captions, subtitles |
| Divider | `rgba(200,170,255,0.4)` | `·` or `✦` separators in HUD/scoreboard |
| Error red | `#ff7070` / `rgba(255,112,112,0.08)` bg | Error messages |

---

### H1 — Gold Gradient Heading

```css
h1 {
  font-family: 'Cinzel', serif;
  font-size: clamp(1.2rem, 4vw, 2rem);
  font-weight: 700;
  letter-spacing: 0.12em;
  background: linear-gradient(135deg, #f5e083 0%, #e8c84a 30%, #fff8d0 55%, #c8960a 80%, #f5e083 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  margin-bottom: 7px;
  filter: drop-shadow(0 2px 8px rgba(220,180,50,0.35));
}
```

---

### Scoreboard / HUD Pill

Used for score displays, match stats, and any persistent data row at the top of the UI:

```css
#scoreboard, #hud {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  justify-content: center;
  background: linear-gradient(135deg,
    rgba(255,255,255,0.06) 0%,
    rgba(180,140,255,0.08) 50%,
    rgba(100,200,255,0.06) 100%);
  border: 1px solid rgba(200,170,255,0.25);
  border-radius: 50px;
  padding: 6px 18px;
  font-family: 'Cinzel', serif;
  font-size: clamp(0.75rem, 2vw, 0.95rem);
  letter-spacing: 0.05em;
  margin: 5px auto;
  backdrop-filter: blur(10px);
  box-shadow: 0 0 20px rgba(160,120,255,0.12), inset 0 1px 0 rgba(255,255,255,0.12);
}
```

---

### Buttons

All buttons share this base rule:

```css
button {
  font-family: 'Cinzel', serif;
  font-size: clamp(0.66rem, 1.7vw, 0.85rem);
  font-weight: 600;
  letter-spacing: 0.08em;
  padding: 8px 16px;
  border: none;
  border-radius: 10px;
  cursor: pointer;
  transition: transform 80ms ease, box-shadow 80ms ease, opacity 200ms ease;
  -webkit-tap-highlight-color: transparent;
  user-select: none;
}
```

Each button variant uses a different colour gradient but shares the same 3D press pattern:
- **Shadow offset**: `0 5px 0 <dark-colour>` (creates the depth)
- **Active state**: `transform: translateY(4px); box-shadow: 0 1px 0 <dark-colour>;`
- **Disabled state**: `opacity: 0.4; cursor: not-allowed; pointer-events: none;`

#### Standard button colour variants

| Variant | Top colour | Mid colour | Dark base | Shadow base | Use |
|---------|-----------|-----------|-----------|-------------|-----|
| Gold (primary CTA) | `#ffe57a` | `#d4a017` | `#8a6000` | `#5a3d00` | Rules, main action |
| Purple | `#c084fc` | `#7c3aed` | `#3b0764` | `#2e0650` | Finish, secondary |
| Cyan-blue | `#67e8f9` | `#0284c7` | `#0a1a4a` | `#062040` | New Game, reset |
| Green | `#6ee7b7` | `#059669` | `#022c22` | `#014d38` | Play Again, confirm |
| Red-pink | `#fda4af` | `#e11d48` | `#4c0519` | `#3b0a1e` | Rules (alt), danger |
| Amber (dice/roll) | `#fde68a` | `#f59e0b` | `#78350f` | `#92400e` | Roll, in-game action |
| Grey (off/disabled toggle) | `#94a3b8` | `#475569` | `#1e293b` | `#0f172a` | Toggled-off states |

---

### Canvas Wrap — Opal Animated Border

All canvas-based projects wrap the canvas in an animated opal gradient border:

```css
#canvas-wrap {
  display: inline-block;
  border-radius: 18px;
  padding: 3px;
  background: linear-gradient(135deg,
    #c084fc 0%, #60b8ff 20%, #f5e083 40%,
    #6ee7b7 60%, #f472b6 80%, #c084fc 100%);
  box-shadow:
    0 0 40px rgba(150,100,255,0.3),
    0 0 80px rgba(80,160,255,0.15),
    0 20px 60px rgba(0,0,0,0.6);
  animation: opalShift 8s linear infinite;
}

@keyframes opalShift {
  0%   { filter: hue-rotate(0deg)  brightness(1);    }
  50%  { filter: hue-rotate(30deg) brightness(1.08); }
  100% { filter: hue-rotate(0deg)  brightness(1);    }
}

canvas {
  display: block;
  border-radius: 15px;
  background: linear-gradient(145deg, #0f0c29 0%, #1a1560 50%, #0d1b3e 100%);
}
```

For non-canvas projects (e.g. `worldcup2026-scores`), the opal wrap is applied to the main content container (`#content-wrap`) using the same CSS above with an inner `border-radius: 15px` div.

---

### Modals & Overlays

```css
.modal {
  position: fixed; top: 50%; left: 50%;
  transform: translate(-50%, -50%);
  z-index: 100;
  padding: 30px 36px;
  border-radius: 20px;
  background: linear-gradient(145deg, rgba(20,10,50,0.97) 0%, rgba(30,15,70,0.97) 100%);
  border: 1px solid rgba(200,160,255,0.3);
  box-shadow:
    0 0 60px rgba(150,80,255,0.3),
    0 30px 80px rgba(0,0,0,0.8),
    inset 0 1px 0 rgba(255,255,255,0.1);
  backdrop-filter: blur(20px);
  width: min(92vw, 520px);
}
```

Rules overlays (full-screen) use a semi-transparent backdrop:

```css
#rules-overlay {
  position: fixed; inset: 0; z-index: 200;
  background: rgba(8,4,28,0.82);
  backdrop-filter: blur(6px);
  display: flex; align-items: center; justify-content: center;
  padding: 16px;
}
```

---

### Stage / Section Badges

Used to label rounds, groups, or sections (e.g. World Cup group headings, game-mode badges):

```css
.round, .stage-badge {
  font-family: 'Cinzel', serif;
  font-size: clamp(0.72rem, 2vw, 0.9rem);
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: #c084fc;
  background: linear-gradient(135deg, rgba(192,132,252,0.12) 0%, rgba(96,184,255,0.08) 100%);
  border: 1px solid rgba(192,132,252,0.25);
  border-radius: 50px;
  padding: 5px 20px;
  display: inline-block;
  backdrop-filter: blur(6px);
}
```

---

### CSS Tag Sync — Rules for Developers and AI Agents

The term **"current tag"** refers to the canonical token values defined in this section. The current tag is versioned informally — when a value changes here, it changes in **all in-scope files simultaneously** in the same commit.

#### When to sync

| Trigger | Action required |
|---------|-----------------|
| Editing any CSS in a Royal Edition file | Read this section first; verify all tokens in that file match the current tag |
| Creating a new Royal Edition HTML project | Apply every token and pattern from this section before writing any project-specific CSS |
| A token value changes in this README | Update every in-scope file to the new value in the same commit |

#### How to do a CSS tag sync

1. **Read this README section** in full before touching any file.
2. **Read each in-scope HTML file** and compare every token against the tables above.
3. **Apply only diffs** — do not rewrite working game/tool logic, JS, or layout structure.
4. **Push all changed files in a single commit** with message format:
   ```
   style: harmonise inline CSS tag across <list of files>
   ```
5. **Do not touch exempt files** (`logo-turtle/`) unless the task explicitly targets them.

#### What is project-specific (do not sync)

Each project has CSS that is intentionally unique and must not be overwritten during a sync:

| Project | Project-specific CSS |
|---------|---------------------|
| `dots-game` | `.legend`, `.dice-panel`, `#dice-face` animation, `#btn-roll.used` |
| `triangles-game` | `#mode-badge`, `.rules-section h3`, dot/line/triangle canvas draw styles |
| `tetris-game` | `#side-panel`, `.panel-box`, `canvas#next`, `#controls-box`, Tetromino block 3D shading |
| `worldcup2026-scores` | `.match-card` grid layout, `.team`, `.team-flag`, `.match-score`, `.venue` |

---

## Scripts

| Folder | Type | Description |
|--------|------|-------------|
| [`bold-list-prefixes-word-outlook`](bold-list-prefixes-word-outlook/) | VBA (Word / Outlook) | Bolds the prefix of every bulleted or numbered list item up to the first colon or dash. Works in Word and the Outlook message editor. |
| [`clean-data-whatsapp`](clean-data-whatsapp/) | PowerShell | Cleans exported WhatsApp chat text files — removes formatting artifacts, validates paths, and extracts structured data. |
| [`dots-game`](dots-game/) | HTML / JavaScript | Two-player browser-based Dots and Boxes game on a 5×5 grid. Players take turns selecting pairs of adjacent dots to draw lines; completing the fourth side of a box claims it. Royal Edition visual style matching the Triangles game. No install required — open in any modern browser. |
| [`email-campaign-tracker`](email-campaign-tracker/) | VBA (Outlook) | Scans a selected Outlook folder for manager replies, matches them against a manager list Excel file, and overwrites a fixed tracker report `Manager_Response_Tracke.xlsx` with response status, attachment detection, and clarification flags. |
| [`excel-formatting`](excel-formatting/) | VBA (Excel) | Unified Excel formatting and cleanup macro with four interactive modes: simple formatting, advanced formatting, optional column splitting, and generic table extraction. |
| [`logo-turtle`](logo-turtle/) | HTML / JavaScript | Browser-based LOGO turtle graphics interpreter with two editions: Standard and Royal 3D. Draw with classic commands — `fd`, `bk`, `rt`, `lt`, `repeat`, and more. Clear button wipes the command scroll; Reset button clears the canvas and returns the turtle to centre. Exempt from Royal Edition CSS tag sync — uses a distinct gold/velvet sub-theme. No install required — open in any modern browser. |
| [`outlook-keyword-search`](outlook-keyword-search/) | VBA + PowerShell | Single and batch keyword search across Outlook folders. Available in a standalone VBA version and a PS-assisted version where Outlook stays fully responsive. |
| [`split-excel-by-manager`](split-excel-by-manager/) | VBA + Python | Splits an Excel workbook into separate files, one per unique manager name. Available as a VBA macro and a cross-platform Python script producing identical output. |
| [`tetris-game`](tetris-game/) | HTML / JavaScript | Single-file browser Tetris game with all 7 Tetrominoes, ghost piece, hard drop, wall-kick rotation, and progressive leveling. Royal Edition visual style. No install required — open in any modern browser. |
| [`triangles-game`](triangles-game/) | HTML / JavaScript | Two-player browser-based dice game on a 24-dot triangle grid. Royal Edition visual style. No install required — open in any modern browser. |
| [`win11-startup`](win11-startup/) | PowerShell | Self-healing Windows 11 startup launcher with Win32 shortcut-based repair, dynamic Appx AUMID resolution, and runtime presence-mode detection (Window vs Tray) that removes the need for per-app flags. |
| [`word-normalize-table`](word-normalize-table/) | VBA (Word) | Normalises all tables in the active document — sets width to 100%, clears row/cell constraints, centres cell content vertically. |
| [`word-resize-border-images`](word-resize-border-images/) | VBA (Word) | Resizes inline images to a user-specified width range, applies a configurable border, and cleans blank paragraphs including Unicode invisible characters. |
| [`worldcup2026-scores`](worldcup2026-scores/) | HTML / JavaScript | Single-file browser app showing all 104 FIFA World Cup 2026 matches — group stage through final — with live scores, kickoff times, venues, and country flags. Fetches real-time data from the Zafronix World Cup API (free tier, 250 req/day). Royal Edition visual style. No install required — open in any modern browser. |

---

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/varunbiswasgit/aiacode.git
cd aiacode
```

### 2. Find the script you need

Browse the Scripts table above. Each folder's `README.md` covers purpose, configuration, and usage.

---

## Running Tests

Each tool folder contains a `TESTING.md` with step-by-step manual test cases and, where applicable, an automated test script.

### Automated tests — Python

```bash
pip install pandas openpyxl pytest
pytest <tool-folder>/ -v
```

### Automated tests — PowerShell (Pester)

```powershell
cd <tool-folder>
Invoke-Pester . -Output Detailed
```

### Automated tests — VBA

1. Open the target Office application (Word, Excel, or Outlook).
2. Open the VBA editor (`Alt+F11`).
3. Import both the main script and the test script from the tool folder.
4. Run the `RunAllTests` macro (`Alt+F8 → RunAllTests → Run`).
5. Results appear in the Immediate window (`Ctrl+G`) and a summary dialog.

### Manual tests

For scripts without automated tests, follow the step-by-step test cases in `TESTING.md`. Each case specifies:

- **Setup** — the document or data state required before running
- **Action** — what to run or trigger
- **Expected** — what the output should look like
- **Pass criteria** — unambiguous pass/fail condition

---

## Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting a pull request.

---

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).
