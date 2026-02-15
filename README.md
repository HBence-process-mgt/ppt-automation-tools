# ppt-automation-tools
PPT automation through pywin32

## Installation

This project depends on the Windows-only `pywin32` package for COM automation.

Install from the included `requirements.txt`:

```bash
python -m pip install -r requirements.txt
```

If you prefer to install the dependency directly (Windows only):

```bash
python -m pip install 
pip install pywin32
```

## Title Font Size Updater
### Description

This Python script updates the font size of all slide titles in the currently active PowerPoint presentation.
It uses pywin32 to interface with PowerPoint via COM automation.

**Important: The script also resets the title formatting to the template defaults before applying the new font size. This ensures that the slide titles follow the master template’s style (font, paragraph format, and indentation) while only changing the size.**

### Usage
1. Open your PowerPoint presentation.
2. Run the script.
3. Enter the desired font size (maximum 409).
4. The script will update all slide titles and print the number of titles updated.
5. If needed, you can undo the changes in PowerPoint with Ctrl+Z.

### Notes
- The script preserves the title text content while resetting all other template formatting.
- Any manually overridden formatting (font type, paragraph spacing, indent) will be reset to match the template.
- The script only affects title placeholders (HasTitle = True). Other text boxes are not modified.

## Recurring Shape Hider

### Description
This script searches through all slides in the currently active PowerPoint presentation and hides shapes whose name matches a user-provided search string.

The script is useful for hiding recurring elements such as icons, decorative shapes, watermark-like objects, or template artifacts that appear across many slides.

**Important: Shape name searching is NOT done by perfect/exact match. The script checks whether your input text is contained anywhere in the shape name (partial match).**

### Usage
1. Open the PowerPoint presentation you want to edit.
2. Run the script.
3. When prompted, enter the shape name (or part of the shape name) you want to search for.
4. The script will hide all matching shapes and print how many were hidden.

### How It Works
- The script loops through all slides in the active presentation.
- It checks every shape name on each slide.
- If the user input is found inside the shape name, the shape is hidden by setting:

shape.Visible = False

### Output Example
Hide recurring shapes with name:
logo
Hidden 14 shape(s)

### Notes
- The search is not case-insensitive.
- Hidden shapes are not deleted, only made invisible.
- Script effects can be reversed with the Undo button (Ctrl+Z) in PowerPoint.