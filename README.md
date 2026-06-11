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

## Scripts

| Folder | Type | Description |
|--------|------|-------------|
| [`bold-list-prefixes-word-outlook`](bold-list-prefixes-word-outlook/) | VBA (Word / Outlook) | Bolds the prefix of every bulleted or numbered list item up to the first colon or dash. Works in Word and the Outlook message editor. |
| [`clean-data-whatsapp`](clean-data-whatsapp/) | PowerShell | Cleans exported WhatsApp chat text files — removes formatting artifacts, validates paths, and extracts structured data. |
| [`dots-game`](dots-game/) | HTML / JavaScript | Two-player browser-based Dots and Boxes game on a 5×5 grid. Players take turns selecting pairs of adjacent dots to draw lines; completing the fourth side of a box claims it. Royal Edition visual style matching the Triangles game. No install required — open in any modern browser. |
| [`excel-formatting`](excel-formatting/) | VBA (Excel) | Unified Excel formatting and cleanup macro with four interactive modes: simple formatting, advanced formatting, optional column splitting, and generic table extraction. |
| [`logo-turtle`](logo-turtle/) | HTML / JavaScript | Browser-based LOGO turtle graphics interpreter with two editions: Standard and Royal 3D. Draw with classic commands — `fd`, `bk`, `rt`, `lt`, `repeat`, and more. Clear button wipes the command scroll; Reset button clears the canvas and returns the turtle to centre. No install required — open in any modern browser. |
| [`outlook-keyword-search`](outlook-keyword-search/) | VBA + PowerShell | Single and batch keyword search across Outlook folders. Available in a standalone VBA version and a PS-assisted version where Outlook stays fully responsive. |
| [`split-excel-by-manager`](split-excel-by-manager/) | VBA + Python | Splits an Excel workbook into separate files, one per unique manager name. Available as a VBA macro and a cross-platform Python script producing identical output. |
| [`tetris-game`](tetris-game/) | HTML / JavaScript | Single-file browser Tetris game with all 7 Tetrominoes, ghost piece, hard drop, wall-kick rotation, and progressive leveling. No install required — open in any modern browser. |
| [`triangles-game`](triangles-game/) | HTML / JavaScript | Two-player browser-based dice game on a 24-dot triangle grid. No install required — open in any modern browser. |
| [`win11-startup`](win11-startup/) | PowerShell | Self-healing Windows 11 startup launcher with Win32 shortcut-based repair, dynamic Appx AUMID resolution, and runtime presence-mode detection (Window vs Tray) that removes the need for per-app flags. |
| [`word-normalize-table`](word-normalize-table/) | VBA (Word) | Normalises all tables in the active document — sets width to 100%, clears row/cell constraints, centres cell content vertically. |
| [`word-resize-border-images`](word-resize-border-images/) | VBA (Word) | Resizes inline images to a user-specified width range, applies a configurable border, and cleans blank paragraphs including Unicode invisible characters. |
| [`worldcup2026-scores`](worldcup2026-scores/) | HTML / JavaScript | Single-file browser app showing all 104 FIFA World Cup 2026 matches — group stage through final — with live scores, kickoff times, venues, and country flags. Fetches real-time data from the Zafronix World Cup API (free tier, 250 req/day). No install required — open in any modern browser. |

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
