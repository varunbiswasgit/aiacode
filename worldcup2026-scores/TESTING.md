# TESTING — worldcup2026-scores

---

## Test Environment

- Browser: Any modern browser (Chrome, Edge, Firefox, Safari)
- Network: Active internet connection required for live data
- File: Open `worldcup2026_scores.html` directly (no server needed)

---

## Test Cases

### TC-01 — Page loads and auto-fetches matches

| Field | Detail |
|-------|--------|
| **Setup** | Open `worldcup2026_scores.html` in a browser with internet access |
| **Action** | No action — page auto-fetches on load |
| **Expected** | All 104 matches displayed, grouped by stage (Group A through Final) |
| **Pass criteria** | At least one match card visible with team names, date, and kickoff time |

---

### TC-02 — Country flags display correctly

| Field | Detail |
|-------|--------|
| **Setup** | Page loaded with matches visible |
| **Action** | Visually inspect flags next to team names |
| **Expected** | Flag emoji renders for all 48 qualified nations (e.g., 🇸🇪 Sweden, 🇺🇸 USA, 🇧🇷 Brazil) |
| **Pass criteria** | No 🏳️ fallback flag for any named qualified team |

---

### TC-03 — Refresh button reloads data

| Field | Detail |
|-------|--------|
| **Setup** | Page loaded with matches visible |
| **Action** | Click **Refresh Scores** |
| **Expected** | Button shows “Loading…”, matches reload, button re-enables |
| **Pass criteria** | Match list re-renders without page reload; no JS errors in console |

---

### TC-04 — Scores show when available

| Field | Detail |
|-------|--------|
| **Setup** | At least one match has been played |
| **Action** | View match cards for completed matches |
| **Expected** | Score displayed as `X - Y` in cyan; unplayed matches show `-` in grey |
| **Pass criteria** | No null or undefined values rendered in score column |

---

### TC-05 — Mobile responsive layout

| Field | Detail |
|-------|--------|
| **Setup** | Open page in browser DevTools at 375px width |
| **Action** | Scroll through all match cards |
| **Expected** | Cards stack vertically; no horizontal overflow; text readable |
| **Pass criteria** | No clipped content or horizontal scrollbar |

---

### TC-06 — API failure shows error message

| Field | Detail |
|-------|--------|
| **Setup** | Disconnect internet or set an invalid API key |
| **Action** | Load or refresh the page |
| **Expected** | Red error box shown with HTTP status or message |
| **Pass criteria** | App does not crash; error message is visible and descriptive |
