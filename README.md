---

# AI Assisted Microsoft Macro, Bash, and PowerShell Scripts

## Overview
This repository contains a collection of PowerShell, MS Macros, and Bash scripts developed using AI tools. The scripts have been tested and worked to the best of my knowledge at the time of posting. The repository also includes tools and scripts for automating workflows, such as bulk issue creation in GitHub.

## Scripts Index
1. **Clean_data_Whatsapp.ps1**: A script for cleaning and processing text files, such as removing extra quotes, validating paths, and extracting data matching specific patterns.
2. **WordResizeBorderImagesCleanlines.bas**: This program, written in VBA (Visual Basic for Applications), performs Resize and border Images:  and Cleans the Document.
2. _[Placeholder for additional scripts]_: Add descriptions for other scripts as you add them.

---

## Features
- Modular scripts with clear functionality.
- Easy-to-use prompts and customizable logic.
- Automated GitHub issue creation via CLI.
- Well-documented examples and use cases (WIP).

---

## Getting Started
### 1. Clone the repository:
   ```
   git clone https://github.com/varunbiswasgit/aiacode.git
   ```
### 2. Navigate to the specific script directory:
   ```
   cd aiacode
   ```

### 3. Adding Task Lists in Script-Specific `README.md` Files.
Use task lists in the script's `README.md` file for each script in this repository to track specific tasks, planned improvements, or testing workflows. Task lists are a Markdown feature that allows easy tracking and progress management.

#### **Task List Syntax**
To add a task list, use the following syntax:
```
- [ ] Task 1 description
- [ ] Task 2 description
- [x] Completed Task description
```
- `[ ]` indicates an **incomplete task**.
- `[x]` indicates a **completed task**.

#### **Example**
For the `Clean_data_Whatsapp.ps1` script:
```
## Task List for Clean_data_Whatsapp.ps1
- [x] Validate and sanitize input file paths.
- [x] Handle multi-line WhatsApp messages.
- [ ] Add advanced Unicode character handling.
- [ ] Optimize regex performance for large files.
- [ ] Add support for additional export formats.
```

#### Task lists can also link directly to GitHub Issues or PRs:
```
- [ ] [#4 Optimize regular expressions](https://github.com/varunbiswasgit/aiacode/issues/4)
```

#### Viewing Task Progress
When the `README.md` is viewed in GitHub, task lists display progress with checkboxes:
```
- Example:
  - [ ] Incomplete Task
  - [x] Completed Task3. Follow instructions in each script's individual `README.md`.
```

### 4. Individual README files are named `<filename.extension README.md>`, e.g., `Clean_data_whatsapp.ps1 README.md`.

---

## Automated GitHub Issue Creation
This repository includes the `bulk_create_issues.tmp` script, which automates the creation of multiple GitHub issues. This is useful for managing tasks and tracking features or bugs.

### Prerequisites
To use the script, ensure the following tools are installed on your system:

#### 1. **GitHub CLI (`gh`)**
   - Install the GitHub CLI:
     - **macOS** or **Linux** (via Homebrew):
       ```
       brew install gh
       ```
     - **Windows** (via Chocolatey):
       ```
       choco install gh
       ```
   - Authenticate with your GitHub account:
     ```
     gh auth login
     ```

#### 2. **Chocolatey (`choco`)**
   - If you don’t have Chocolatey installed (Windows users), follow these steps:
     1. Open a Command Prompt or PowerShell with admin rights.
     2. Install Chocolatey by running:
        ```
        Set-ExecutionPolicy Bypass -Scope Process -Force; `
        [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; `
        iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        ```
     3. Verify installation:
        ```
        choco --version
        ```

#### 3. **jq**
   - jq is required to handle JSON data in the bulk issue creation script:
     - **macOS** or **Linux**:
       ```
       brew install jq
       ```
     - **Windows**:
       ```
       choco install jq
       ```

---

### Bulk Issue Creation Workflow
1. Prepare your `issues.tmp` file in the following format:
   ```
   Issue Title 1::Issue description for Title 1
   Issue Title 2::Issue description for Title 2
   Add Error Handling::Improve error handling and provide detailed error messages.
   Parameterize Input Paths::Allow users to pass file paths as command-line parameters.
   ```

2. Run the `bulk_create_issues.tmp` script:
   ```
   bash bulk_create_issues.tmp
   ```

3. The script will:
   - Read titles and descriptions from `issues.tmp`.
   - Check for duplicate issues in your repository to avoid duplicates.
   - Create new issues using the GitHub CLI.

---

## Contributing
Contributions are welcome! Please read our [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

---

## License
This project is licensed under the [GNU General Public License v3.0](LICENSE).

---

## Notes
1. If you encounter warnings or errors during script execution, check your `issues.tmp` file for formatting issues or verify tool installations.
2. Would you like to automate additional processes like creating milestones or assigning labels? Let me know! 😊

---
