# tetris-game — Test Plan

All tests are manual. Open `tetris-royal-3d.html` in a modern browser (Chrome, Edge, or Firefox) and follow each case.

---

## Environment

- Browser: Chrome 120+ / Edge 120+ / Firefox 120+
- No server required — open directly via `file://` or any web server
- For mobile tests: Chrome on Android / Safari on iOS, or DevTools device emulation

---

## Test Cases

### TC-01 — Start screen
| Field | Detail |
|-------|--------|
| Setup | Open `tetris-royal-3d.html` fresh |
| Action | Observe the overlay |
| Expected | Royal modal shows title, subtitle, and ◆ Start Game button (disabled until member word entered) |
| Pass | Modal visible; board empty; no pieces moving |

### TC-02 — Member word unlock
| Field | Detail |
|-------|--------|
| Setup | Start screen visible |
| Action | Type a valid member word in the Member Word field |
| Expected | Green hint appears; Start button enables; Member badge shows |
| Pass | Button unlocked; invalid word keeps button disabled |

### TC-03 — Game starts
| Field | Detail |
|-------|--------|
| Setup | Member word accepted |
| Action | Click ◆ Start Game |
| Expected | Modal hides; piece begins falling; Score/Level/Lines show 0/1/0 |
| Pass | Piece animates; HUD updates |

### TC-04 — Left / right movement
| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Press ← and → |
| Expected | Active piece moves one column per keypress; stops at walls |
| Pass | No movement beyond column 0 or column 9 |

### TC-05 — Rotation with wall kick
| Field | Detail |
|-------|--------|
| Setup | Move an I-piece to the far right wall |
| Action | Press ↑ |
| Expected | Piece rotates; kicks left if blocked |
| Pass | Rotation completes without clipping |

### TC-06 — Soft drop scoring
| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Hold ↓ for several cells |
| Expected | Score increases by 1 per cell |
| Pass | Score delta equals cells dropped |

### TC-07 — Hard drop scoring
| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Press Space |
| Expected | Piece drops instantly; score +2 × distance |
| Pass | Piece locks; score delta correct |

### TC-08 — Line clear and score
| Field | Detail |
|-------|--------|
| Setup | Fill one complete row |
| Action | Complete the row |
| Expected | Row disappears; score +100 × level |
| Pass | Board collapses; score updates |

### TC-09 — Tetris (4-line clear)
| Field | Detail |
|-------|--------|
| Setup | Stack pieces leaving a single-column gap |
| Action | Drop an I-piece into the gap |
| Expected | 4 rows clear; score +800 × level |
| Pass | All 4 rows removed; score correct |

### TC-10 — Level progression
| Field | Detail |
|-------|--------|
| Setup | Game running at level 1 |
| Action | Clear 10 lines |
| Expected | Level increments to 2; speed increases |
| Pass | Level = 2 in HUD; faster drop confirmed |

### TC-11 — Pause and resume
| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Press P; wait; press P again |
| Expected | PAUSED overlay in gold; resumes correctly |
| Pass | State identical before/after pause |

### TC-12 — Ghost piece
| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Move piece left and right |
| Expected | Ghost updates to projected landing position |
| Pass | Ghost aligns with lock position |

### TC-13 — Game over
| Field | Detail |
|-------|--------|
| Setup | Game running |
| Action | Stack to top |
| Expected | Game Over modal with score; Play Again resets |
| Pass | Modal shown; board resets on Play Again |

### TC-14 — Responsive layout (desktop)
| Field | Detail |
|-------|--------|
| Setup | Desktop browser |
| Action | Resize window wide to narrow |
| Expected | Canvas resizes; no overflow |
| Pass | Board fits viewport; no scrollbar |

### TC-15 — Responsive layout (mobile)
| Field | Detail |
|-------|--------|
| Setup | DevTools emulation (e.g. iPhone 14) |
| Action | Load page |
| Expected | Canvas fills height; side panel below board |
| Pass | No clipping |

### TC-16 — Touch rotate
| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Tap canvas |
| Expected | Piece rotates |
| Pass | Rotates on each tap |

### TC-17 — Touch move
| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Swipe left / right |
| Expected | Piece moves; stops at walls |
| Pass | No rotation triggered |

### TC-18 — Touch hard drop
| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Swipe up |
| Expected | Piece hard-drops; score +2 × rows |
| Pass | Piece locks immediately |

### TC-19 — Touch soft drop
| Field | Detail |
|-------|--------|
| Setup | Game running on touch device |
| Action | Swipe down |
| Expected | Soft drop; score +1 per cell |
| Pass | Score correct |

### TC-20 — Rules modal
| Field | Detail |
|-------|--------|
| Setup | Any state |
| Action | Click ◆ Rules |
| Expected | Rules overlay opens; Close button and Esc dismiss it |
| Pass | Overlay opens and closes correctly |

### TC-21 — Shell loader
| Field | Detail |
|-------|--------|
| Setup | Internet connection available |
| Action | Open `shell.html` |
| Expected | Spinner then game loads in iframe |
| Pass | Game playable inside iframe |

### TC-22 — Member words load (shell context)
| Field | Detail |
|-------|--------|
| Setup | Open `shell.html` |
| Action | Enter valid member word; start |
| Expected | Word accepted; game starts; no fetch errors |
| Pass | DevTools shows no errors |

### TC-23 — Member words admin — add and download
| Field | Detail |
|-------|--------|
| Setup | Open `member-words-admin.html` |
| Action | Add word e.g. "galaxy"; click Download |
| Expected | Chip appears; JSON contains "galaxy" |
| Pass | JSON valid; no duplicates |

### TC-24 — Member words admin — remove word
| Field | Detail |
|-------|--------|
| Setup | `member-words-admin.html` open |
| Action | Click ✕ on a word chip |
| Expected | Chip removed; JSON updated |
| Pass | Word absent from JSON |

### TC-25 — Member words admin — import JSON
| Field | Detail |
|-------|--------|
| Setup | `member-words-admin.html` open |
| Action | Paste valid JSON; click Load JSON |
| Expected | Word list replaced; count updates |
| Pass | New list rendered correctly |
