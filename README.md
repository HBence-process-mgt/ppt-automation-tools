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