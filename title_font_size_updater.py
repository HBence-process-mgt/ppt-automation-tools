import win32com.client
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.ActivePresentation

print("Update title font size to:")
size = min(int(input()), 409)
placeholder = ""
updated_count = 0

for slide in presentation.Slides:
    if slide.Shapes.HasTitle:
        title = slide.Shapes.Title
        if title is not None:
            placeholder = title.TextFrame.TextRange.Text
            title.TextFrame.TextRange.Text = ""
            title.Delete()
            slide.Shapes.AddTitle()
            title = slide.Shapes.Title
            title.TextFrame.AutoSize = 0
            title.TextFrame.TextRange.Text = placeholder
            title.TextFrame.TextRange.Font.Size = size
            updated_count += 1
        else:
            continue
    else:
        continue

print(f"Updated {updated_count} title(s)")
print("Script effects can be reversed with the Undo button or Ctrl+Z in PowerPoint.")
print("Press Enter to exit...")