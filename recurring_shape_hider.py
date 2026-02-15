import win32com.client
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.ActivePresentation

print("Hide recurring shapes with name:")
name = input().strip().lower()
hidden_count = 0
if name == "":
    print("No name entered. Exiting...")
    print("Press Enter to exit...")
    exit()

for slide in presentation.Slides:
    if slide.Shapes.Count > 0:
        for shape in slide.Shapes:
            if name in shape.Name.lower():
                shape.Visible = False
                hidden_count += 1
            else:
                continue
    else:
        continue

print(f"Hidden {hidden_count} shape(s)")
print("Script effects can be reversed with the Undo button or Ctrl+Z in PowerPoint.")
print("Press Enter to exit...")