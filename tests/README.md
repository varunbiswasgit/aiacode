# tests/

This folder contains all test scripts and their corresponding Testing Readme files.

## Structure

Each script in `scripts/` has a dedicated Testing Readme in this folder. The Testing Readme is the primary reference for that script's test cases — it lists every manual test case, the expected outcome, and the pass criteria. Where an automated test file also exists (Python or PowerShell), it lives alongside the Testing Readme in this folder and is named after the script it tests.

The Testing Readme for a script should always be read together with two other files: the script file itself in `scripts/`, and the script's main README in the `README/` folder. The main README explains what the script does and how to install it; the Testing Readme explains how to verify it works correctly. The script file is the source of truth for behaviour — if a test case and the code appear to conflict, the code takes precedence.

Naming is consistent across all three files for each script. A script named `ExcelFormatting.bas` will have a main README at `README/ExcelFormatting.bas README.md` and a Testing Readme at `tests/ExcelFormatting.bas Testing Readme.md`. This one-to-one-to-one relationship makes it straightforward to locate all documentation for any given script.

## Naming Convention

Every Testing Readme follows the pattern:
```
<script filename> Testing Readme.md
```

## Running Automated Tests

### Python (pytest)

#### One-time setup

**1. Install Python**

Download and install Python 3.9 or later from [python.org/downloads](https://www.python.org/downloads/). During installation on Windows, tick **Add Python to PATH** before clicking Install Now. Verify the installation by opening a terminal and running:

```bash
python --version
```

**2. Confirm pip is available**

`pip` is Python’s package installer and is bundled with Python 3.4+. Confirm it is present:

```bash
pip --version
```

If pip is missing, reinstall Python or run `python -m ensurepip --upgrade`.

**3. Create a virtual environment (recommended)**

A virtual environment isolates this project’s dependencies from other Python projects on the same machine.

```bash
# From the repository root
python -m venv .venv
```

Activate it before installing packages or running tests:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Your terminal prompt will show `(.venv)` when the environment is active. Run `deactivate` to exit it.

**4. Install dependencies**

Install pytest and all packages the scripts under test require:

```bash
pip install pytest openpyxl pandas
```

If a `requirements.txt` file exists in the repository root, install from that instead:

```bash
pip install -r requirements.txt
```

#### Running the tests

Run all Python tests from the repository root with verbose output:

```bash
pytest tests/test_split_excel_by_manager.py -v
```

To run all test files discovered in the `tests/` folder at once:

```bash
pytest tests/ -v
```

> Always activate the virtual environment before running pytest. If `pytest` is not found, use `python -m pytest` instead.

---

### PowerShell (Pester)

#### One-time setup

**1. Confirm PowerShell version**

Pester 5 requires PowerShell 5.1 or later. Check the version:

```powershell
$PSVersionTable.PSVersion
```

PowerShell 5.1 ships with Windows 10 and later. PowerShell 7+ can be downloaded from [github.com/PowerShell/PowerShell/releases](https://github.com/PowerShell/PowerShell/releases) and runs side-by-side with 5.1.

**2. Set execution policy**

PowerShell blocks unsigned scripts by default. Allow locally written scripts to run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

`RemoteSigned` permits local scripts and only requires a signature for scripts downloaded from the internet. This change applies to the current user only and does not affect system policy.

**3. Install Pester**

Pester is not included with Windows by default (an outdated version 3 stub may be present). Install the current version from the PowerShell Gallery:

```powershell
Install-Module -Name Pester -Force -SkipPublisherCheck
```

`-Force` overwrites any pre-installed stub. `-SkipPublisherCheck` is needed because the publisher certificate may differ between the bundled and gallery versions. Verify the installed version:

```powershell
Get-Module -Name Pester -ListAvailable | Select-Object Name, Version
```

The version should be 5.x or later.

**4. Import Pester before first use (optional)**

Pester is auto-imported when you call `Invoke-Pester`, but you can import it explicitly to confirm the correct version is loaded:

```powershell
Import-Module Pester -MinimumVersion 5.0
```

#### Running the tests

Run the PowerShell test file with detailed output from the repository root:

```powershell
Invoke-Pester .\tests\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

To run all Pester test files discovered in the `tests/` folder:

```powershell
Invoke-Pester .\tests\ -Output Detailed
```

> If you see a version conflict warning, run `Remove-Module Pester` first and then re-import with the minimum version flag above.

---

### VBA / Word Macros

Manual only. Follow the environment setup steps in the respective Testing Readme.  
Complete the one-time IDE setup below before running any `.bas` script for the first time.

---

#### Step 1 — Enable the Developer tab

1. Open Excel or Word.
2. Go to **File → Options → Customize Ribbon**.
3. In the right-hand column, tick **Developer**, then click **OK**.
4. The **Developer** tab now appears in the ribbon.

> This step is required only once per Office installation.

---

#### Step 2 — Open the VBA IDE (Alt + F11)

1. With a workbook or document open, press **Alt + F11**.
2. The **Visual Basic Editor (VBE)** opens in a new window.
3. The **Project Explorer** panel (left side) lists all open workbooks/documents and their components.

> If Project Explorer is not visible, press **Ctrl + R** to show it.

---

#### Step 3 — Insert a new module

1. In Project Explorer, right-click the target workbook/document node (e.g., `VBAProject (Book1)`).
2. Choose **Insert → Module**.
3. A new `Module1` appears under the **Modules** folder.
4. The blank module opens in the code editor on the right.

> Alternatively, use the menu bar: **Insert → Module** while any project node is selected.

**To import an existing `.bas` file instead of inserting a blank module:**

1. Right-click the project node in Project Explorer.
2. Choose **Import File…**
3. Browse to the `.bas` file (e.g., `scripts/ExcelFormatting.bas`) and click **Open**.
4. The module appears in Project Explorer under **Modules**.

---

#### Step 4 — Add required library references

Some macros in this repository use late-binding (`CreateObject`) and require no extra references.  
Others use early-binding declarations (e.g., `Scripting.Dictionary`, `Word.Application`) and need the corresponding library checked in the References dialog.

1. In the VBE, go to **Tools → References…**
2. The **Available References** list appears. Tick the libraries below as needed, then click **OK**.

| Library | Required by | Notes |
|---------|-------------|-------|
| **Microsoft Scripting Runtime** | `ExcelFormatting.bas`, `SplitExcelByManager.bas` | Enables early-bound `Scripting.Dictionary` and `FileSystemObject`. Located at `scrrun.dll`. |
| **Microsoft Word 16.0 Object Library** | Word macro scripts | Required when automating Word from Excel VBA. Version number matches your Office installation (16.0 = Office 2016/2019/365). |
| **Microsoft Excel 16.0 Object Library** | Word macro scripts that write back to Excel | Required when an Excel object model reference is needed from outside Excel. |
| **Microsoft Office 16.0 Object Library** | Any script using `Office.CommandBars` or shared Office objects | Usually already checked by default. |

> **Tip — late-binding vs. early-binding:**  
> If a `.bas` file uses `Dim dict As Object` and `Set dict = CreateObject("Scripting.Dictionary")`, it is late-bound and the library reference is optional (but adds IntelliSense when checked).  
> If it uses `Dim dict As Scripting.Dictionary`, it is early-bound and the reference is mandatory.

---

#### Step 5 — Enable macro execution (Trust Center)

1. Go to **File → Options → Trust Center → Trust Center Settings…**
2. Click **Macro Settings**.
3. Select **Enable all macros** for development, or **Disable all macros with notification** for day-to-day use.
4. Tick **Trust access to the VBA project object model** if any script programmatically reads or modifies VBA project components.
5. Click **OK** twice.

> Re-open the workbook/document after changing macro settings for the change to take effect.

---

#### Step 6 — Run a macro

- **From the ribbon:** Developer → Macros → select the macro name → **Run**.
- **From the VBE:** place the cursor inside the `Sub` body and press **F5**, or press **F8** to step through line by line.
- **Keyboard shortcut:** assign a shortcut in Developer → Macros → select macro → **Options…** → set a `Ctrl+` key combination.

---

#### Quick-reference keyboard shortcuts

| Shortcut | Action |
|----------|--------|
| **Alt + F11** | Open / close the VBA IDE |
| **Ctrl + R** | Show Project Explorer |
| **F5** | Run the current Sub / selected macro |
| **F8** | Step into — execute one line at a time |
| **Ctrl + G** | Show the Immediate window (for `Debug.Print` output) |
| **Ctrl + Break** | Interrupt a running macro |
| **Ctrl + Space** | Trigger IntelliSense autocomplete |
| **Ctrl + F** | Find in current module |
| **Ctrl + H** | Find and replace in current module |
