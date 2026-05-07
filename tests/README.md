# tests/

This folder contains all test scripts and their corresponding Testing Readme files.

## Structure

| Script under test | Test file | Testing Readme |
|---|---|---|
| `scripts/ExcelFormatting.bas` | *(manual only — see readme)* | `ExcelFormatting.bas Testing Readme.md` |
| `scripts/SplitExcelByManager.bas` | *(manual only — see readme)* | `SplitExcelByManager.bas Testing Readme.md` |
| `scripts/split_excel_by_manager.py` | `test_split_excel_by_manager.py` | `test_split_excel_by_manager.py Testing Readme.md` |
| `scripts/Clean_data_whatsapp.ps1` | `Clean_data_whatsapp.Tests.ps1` | `Clean_data_whatsapp.ps1 Testing Readme.md` |

## Naming Convention

Every Testing Readme follows the pattern:
```
<script filename> Testing Readme.md
```

## Running Automated Tests

### Python (pytest)
```bash
pytest tests/test_split_excel_by_manager.py -v
```

### PowerShell (Pester)
```powershell
Invoke-Pester .\tests\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

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
