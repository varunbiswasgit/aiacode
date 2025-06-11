# Running Tests

This project uses [Pester](https://github.com/pester/Pester) for unit tests. Ensure you have PowerShell (v7 or newer) and the Pester module installed.

From the repository root, run the tests with:

```powershell
Invoke-Pester
```

Or from a shell:

```bash
pwsh -NoLogo -NoProfile -Command "Invoke-Pester -Output Detailed"
```

The command will discover tests in the `tests/` directory and report the results.
