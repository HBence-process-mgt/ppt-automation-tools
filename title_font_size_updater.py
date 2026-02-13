import win32com.client
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.ActivePresentation

print("Update title font size to:")
size = min(int(input()), 409)
placeholder = ""
updated_count = 0

for slide in presentation.Slides:
    for shape in slide.Shapes:
        if shape.HasTitle:
            placeholder = shape.Title.TextFrame.TextRange.Text
            shape.Title.TextFrame.TextRange.text = ""
            shape.TextFrame.TextRange.Font.Reset()
            shape.TextFrame.TextRange.ParagraphFormat.Reset()
            shape.TextFrame.TextRange.IndentLevel.Reset()
            shape.TextFrame.TextRange.Text = placeholder
            shape.TextFrame.TextRange.Font.Size = size
            updated_count += 1
        else:
            continue

print(f"Updated {updated_count} title(s)")
print("Script effects can be reversed with the Undo button or Ctrl+Z in PowerPoint.")
print("Press Enter to exit...")