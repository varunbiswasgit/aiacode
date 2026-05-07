# Contributing to aiacode

`aiacode` is a personal collection of AI-assisted automation scripts covering VBA (Excel and Word), PowerShell, and Python. Contributions that improve reliability, extend functionality, or enhance documentation are welcome.

---

## What This Repository Contains

- **VBA macros** for Microsoft Excel and Word automation
- **PowerShell scripts** for file and text processing
- **Python scripts** as cross-platform equivalents of VBA tools
- **Per-script READMEs** in the `README/` folder
- **Test documentation and test scripts** in the `tests/` folder

---

## How to Contribute

### 1. Fork and clone

```bash
git clone https://github.com/<your-username>/aiacode.git
cd aiacode
```

### 2. Create a feature branch

```bash
git checkout -b feature/your-feature-name
```

### 3. Make your changes

- Place all scripts in the `scripts/` folder.
- Add or update the script-specific README in the `README/` folder, named `<script filename> README.md`.
- Add or update the Testing Readme in the `tests/` folder, named `<script filename> Testing Readme.md`.
- If adding a Python script, include or update automated test cases in `tests/`.

### 4. Test your changes

- For Python scripts, run the pytest suite:
  ```bash
  pytest tests/ -v
  ```
- For VBA and PowerShell scripts, follow the manual test cases in the relevant Testing Readme.

### 5. Commit with a clear message

Use the following prefix conventions:

| Prefix | Use for |
|--------|---------|
| `feat:` | New script or significant new capability |
| `fix:` | Bug fix |
| `docs:` | README, CONTRIBUTING, or Testing Readme changes only |
| `chore:` | File moves, renames, or housekeeping with no logic change |
| `test:` | Adding or updating test scripts or test documentation |

Example:
```bash
git commit -m "feat(ExcelFormatting): add keyword-based table anchor for Option 3"
```

### 6. Submit a pull request

Open a PR against the `main` branch with a clear description of what changed and why.

---

## Code Standards

### VBA
- Use `Option Explicit` in all modules.
- Use meaningful variable names; avoid single-letter names except loop counters.
- Comment the *why*, not the *what*.
- All user-facing messages use `MsgBox` with appropriate `vbInformation`, `vbExclamation`, or `vbCritical` icons.

### PowerShell
- Follow PowerShell best practices: approved verbs, full parameter names in scripts (not aliases).
- Validate all inputs before processing.
- Use `Write-Host` only for user-facing output; use `Write-Verbose` for diagnostic messages.

### Python
- Target Python 3.9+.
- Follow PEP 8 formatting.
- All functions must have docstrings.
- New scripts must include pytest test cases in `tests/`.

---

## Reporting Issues

Use the [GitHub issue tracker](https://github.com/varunbiswasgit/aiacode/issues) to report bugs or suggest features. Provide:
- The script name and version (if visible in the file header).
- A clear description of the problem.
- Sample input and expected vs. actual output where applicable.

---

## License

By contributing, you agree that your contributions will be licensed under the [GNU General Public License v3.0](LICENSE).
