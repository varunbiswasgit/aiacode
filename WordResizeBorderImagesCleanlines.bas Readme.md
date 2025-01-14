# Resize Images and Clean Documents - VBA Program

This program, written in VBA (Visual Basic for Applications), performs three key tasks in a Microsoft Word document:

1. **Resize Images**
2. **The resized images are also styled with a border for a consistent appearance.**
3. **Cleans the Document**

---

## How to Use

### Prerequisites

- Microsoft Word (with support for running VBA macros).

### Steps to Run

1. Open the Word document you want to modify.
2. Press `Alt + F11` to open the VBA editor.
3. Create a new module:
   - In the VBA editor, click `Insert > Module`.
4. Paste the provided code into the module.
5. Close the VBA editor.
6. Run the macro:
   - Press `Alt + F8` in Word.
   - Select `ResizeImagesAndCleanDocument` and click `Run`.

---

## Code Overview

### Variables

- **`minWidth`** and **`maxWidth`**:
  - Define images' minimum and maximum width (in points, converted from inches).
- **`docRange`**:
  - Represents the content range of the document for processing empty lines.

### Functions

- **Image Resizing**:
  - Check each inline image and resize it if it exceeds the minimum width.
  - Adds a single-line border with a weight of points to resize images.
- **Document Cleanup**:
  - Using wildcards, find and replace three or more consecutive empty lines with two empty lines.

---

## Customization

### Adjust Image Width

To modify the resizing thresholds:

- Update the values of `minWidth` and `maxWidth`:
  ```vb
  minWidth = InchesToPoints(<your_min_width_in_inches>)
  maxWidth = InchesToPoints(<your_max_width_in_inches>)
  ```

### Change Border Style

To customize the image border:

- Update the properties of `.Line`:
  ```vb
  .Line.Weight = <your_line_weight>
  .Line.Style = <your_line_style>
  .Line.ForeColor.RGB = RGB(<red>, <green>, <blue>)
  ```

---

## Notes

- Ensure macros are enabled in Word to execute this program.
- Always save a backup of your document before running macros.
- This macro processes inline shapes only. Floating shapes (e.g., text-wrapped images) are not modified.

---

## License

This program is licensed under the [GNU General Public License v3.0](LICENSE).

