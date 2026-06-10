# worldcup2026-scores

A self-contained browser app that displays all 104 FIFA World Cup 2026 matches — group stage through final — with live scores, kickoff times, venues, and country flags.

---

## Usage

Open `worldcup2026_scores.html` directly in any modern browser. No server or install required.

---

## How It Works

- Fetches live match data from the [Zafronix World Cup API](https://api.zafronix.com) on load.
- Groups matches by stage (Group A–L, Round of 16, Quarter-finals, Semi-finals, Final).
- Displays home team vs away team with country flag emoji, kickoff time, stadium, and live or final score.
- **Refresh Scores** button re-fetches the latest data without reloading the page.

---

## API

| Field | Value |
|-------|-------|
| Provider | [Zafronix](https://api.zafronix.com) |
| Endpoint | `https://api.zafronix.com/fifa/worldcup/v1/matches?year=2026` |
| Auth | `X-API-Key` header |
| Free tier | 250 requests/day, no card required |
| Key in file | `zwc_free_63de5553ae6aa1f348523270` |

> **Note:** Replace the embedded key with your own from [api.zafronix.com/signup](https://api.zafronix.com/signup) if you fork or redeploy.

---

## File Structure

```
worldcup2026-scores/
├── worldcup2026_scores.html   # Single-file app — HTML + CSS + JS
├── README.md                  # This file
└── TESTING.md                 # Manual test cases
```

---

## Configuration

All configuration is inline in the HTML `<script>` block:

| Constant | Purpose |
|----------|---------|
| `API_KEY` | Zafronix API key |
| `API_URL` | Matches endpoint with `?year=2026` |

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release — 104 matches, live scores, country flags, group/stage grouping, responsive layout |
