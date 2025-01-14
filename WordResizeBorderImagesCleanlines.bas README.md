# Resize Images and Clean Documents - VBA Program

This program, written in VBA (Visual Basic for Applications), performs two key tasks in a Microsoft Word document:

1. **Resize Images**: This function adjusts the width of inline images if they exceed a specified minimum width. The resized images are also styled with a border for a consistent appearance.
2. **Cleans the Document**: This function removes unnecessary empty lines (more than two consecutive line breaks) to enhance document readability.

---

## Features

1. **Image Resizing**
   - Ensures all images wider than the minimum width are resized to a maximum width while maintaining their aspect ratio.
   - Adds a customizable border to resized images for visual consistency.
   - Option to apply the border to all images or only to those that meet the resizing criteria (minimum size). This can be adjusted by modifying the border application code inside the size-checking loop.

2. **Empty Line Cleanup**
   - Detects and removes three or more consecutive empty lines, replacing them with two empty lines.

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

7. Enter the required parameters when prompted:
   - Minimum width (in inches).
   - Maximum width (in inches).
   - Border width (in points).
   - Red, Green, and Blue components of the border color (each ranging from 0 to 255).

---

## Code Overview

### User Inputs

- **Minimum Width**: The minimum width (in inches) for resizing images.
- **Maximum Width**: The maximum width (in inches) for resized images.
- **Border Width**: The thickness of the border (in poborder's thickness**: The RGB color values for the border.

### Variables

- **`minWidth`** and **`maxWidth`**:
  - Define the user-specified minimum and maximum width for images (converted toimage widthhes).
- **`borderWidth`**, **`borderColorR`**, **`borderColorG`**, **`borderColorB`**:
  - Define the border thickness and color based on user inputs.
- **`docRange`**:
  - Represents the content range of the document for processing empty lines.

### Functions

- **Image Resizing**:
  - Checks each inline image and resizes it if it exceeds the minimum width.
  - Applies a customizable border with the specified width and color.
  - Border application can be adjusted to all images or only those that fit the resizing criteria.
- **Document Cleanup**:
  - Use wildcards to find and replace three or more consecutive empty lines with two empty lines.

---

## Customization

The program allows full customization through user inputs. You will be prompted to provide values for:

1. Minimum width (in inches).
2. Maximum width (in inches).
3. Border width (in points).
4. Border color (RGB values).

To adjust the border application behavior:
- Move the border-related code inside or outside the resizing condition loop, depending on whether borders should apply to all images or only resized ones.

---

## Notes

- Ensure macros are enabled in Word to execute this program.
- Always save a backup of your document before running macros.
- This macro processes inline shapes only. Floating shapes (e.g., text-wrapped images) are not modified.

---

## License

This program is licensed under the [GNU General Public License v3.0](LICENSE).

