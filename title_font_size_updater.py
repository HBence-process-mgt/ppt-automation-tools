import win32com.client
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.ActivePresentation

print("Update title font size to (number):")
size = min(float(input()), 409)
if size <= 0:
    print("Invalid size. Please enter a positive number:")
    size = min(float(input()), 409)

if presentation.Slides.Count == 0:
    print("No slides found in the presentation.")
    exit()

updated_count = 0

for slide in presentation.Slides:
    if slide.Shapes.HasTitle:  
        try:
            title = slide.Shapes.Title
            placeholder = title.TextFrame.TextRange.Text
            title.TextFrame.TextRange.Text = ""
            title.Delete()
            slide.Shapes.AddTitle()
            title = slide.Shapes.Title
            title.TextFrame.AutoSize = 0
            title.TextFrame.TextRange.Text = placeholder
            title.TextFrame.TextRange.Font.Size = size
            updated_count += 1
        except Exception as e:
            print(f"Slide {slide.SlideNumber} failed: {e}")
            continue
    

print(f"Updated {updated_count} title(s)")
print("Script effects can be reversed with the Undo button or Ctrl+Z in PowerPoint.")
print("Press Enter to exit...")