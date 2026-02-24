import win32com.client
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.ActivePresentation

print("Hide recurring shapes with name:")
name = input().strip().lower()
hidden_count = 0

if name == "":
    print("Please enter a valid shape name:")
    name = input().strip().lower()

for slide in presentation.Slides:
    if slide.Shapes.Count > 0:
        for shape in slide.Shapes:
            try:
                if name in shape.Name.lower():
                    shape.Visible = False
                    hidden_count += 1
            except Exception as e:
                print(f"On slide {slide.SlideNumber} hiding shape '{name}' failed: {e}")

if hidden_count == 0:
    print(f"No shapes with name '{name}' found.")
else:
    print(f"Hidden {hidden_count} shape(s)")

print("Script effects can be reversed with the Undo button or Ctrl+Z in PowerPoint.")
print("Press Enter to exit...")