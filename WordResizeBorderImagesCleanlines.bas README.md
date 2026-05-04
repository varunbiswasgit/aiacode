# Resize Images and Clean Document — VBA Macro

This VBA macro for Microsoft Word performs two tasks on the active document:

1. **Resize Images**: Adjusts inline images wider than a user-specified minimum width to a user-specified maximum width, maintaining aspect ratio.
2. **Clean Empty Lines**: Removes three or more consecutive blank lines, replacing them with two.

---

## Features

### Image Resizing
- Prompts for minimum and maximum width (in inches) with cancel-safe, validated input.
- Resizes only inline images that exceed the minimum width.
- Locks aspect ratio during resizing to prevent distortion.
- Processes `wdInlineShapePicture` types only; floating/anchored shapes are not affected.

### Border Application
- Prompts for border width (in points) and RGB color values.
- Applies a solid single-line border to each resized image using `.Line.Weight`, `.Line.Style`, and `.Line.ForeColor.RGB`.
- Border is applied **only to images that meet the minimum width threshold** and are resized — images below the minimum width are not bordered.
- Default color (RGB 68, 114, 198) renders as a medium blue, matching a standard corporate blue tone.

### Empty Line Cleanup
- Uses a wildcard find-and-replace pattern (`^13{3,}`) with `MatchWildcards = True` to locate three or more consecutive paragraph marks and replace them with two.

---

## How to Use

### Prerequisites
- Microsoft Word with VBA macros enabled.

### Steps

1. Open the Word document you want to modify.
2. Press `Alt + F11` to open the VBA Editor.
3. Click `Insert > Module` to create a new module.
4. Paste the macro code into the module.
5. Close the VBA Editor.
6. Press `Alt + F8`, select `ResizeImagesAndCleanDocument`, and click **Run**.

### Input Prompts

You will be prompted for the following values. Press **Cancel** on any prompt to exit the macro safely.

| Prompt | Unit | Default |
|---|---|---|
| Minimum image width | Inches | 3 |
| Maximum image width | Inches | 6.3 |
| Border width | Points | 1.2 |
| Border color — Red | 0–255 | 68 |
| Border color — Green | 0–255 | 114 |
| Border color — Blue | 0–255 | 198 |

All inputs are validated. Non-numeric or out-of-range values will prompt a retry message.

---

## Known Limitations

- **Inline shapes only**: The macro does not process floating images or shapes wrapped with text.
- **Cancel behavior**: Pressing Cancel on any input box exits the macro immediately without making changes.
- **Border scope**: Only images wider than the minimum width threshold receive a border. Images at or below the minimum are not resized or bordered.

---

## Customization

- To apply borders to **all** images regardless of size, move the `.Line` border code outside the `If .Width > minWidth Then` block.
- To skip border prompts entirely, remove the border input variables and the `.Line` assignment lines from the code.
- To change the default border color, update the default values in the `InputBox` calls for Red, Green, and Blue.

---

## Notes

- Always save a backup of your document before running any macro.
- Ensure **Enable All Macros** or a trusted publisher setting is active in Word's Trust Center.

---

## License

Licensed under the [GNU General Public License v3.0](LICENSE).
