from pathlib import Path
import win32com.client
import ast
import time

print("TTTTT  EEEEE  N   N   GGG   K   K   U   U")
print("  T    E      NN  N  G      K  K    U   U")
print("  T    EEE    N N N  G  GG  KKK     U   U")
print("  T    E      N  NN  G   G  K  K    U   U")
print("  T    EEEEE  N   N   GGG   K   K   UUUU")
print("\n")
print("https://tengkulist.web.app/")
print("\n")

# Path settings
current_dir = Path.cwd()
input_file = current_dir / "Data.pptx"
output_file = current_dir / "New_Data.pptx"
changelist_file = current_dir / "Changelist.txt"

# Read the keywords and replacements from the text file
with open(changelist_file, "r") as f:
    replacements = ast.literal_eval(f.read())

# Open PowerPoint
ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
ppt_app.Visible = True

# Open the input presentation and replace keywords
ppt = ppt_app.Presentations.Open(str(input_file))

# Loop through the list of replacements
for keyword, replacement in replacements:
    # Show loading message
    print(f"Replacing '{keyword}' with '{replacement}'...")

    # Loop through all the slides
    for slide in ppt.Slides:
        # Loop through all the shapes in the slide
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    # Replace the keywords in the text
                    shape.TextFrame.TextRange.Replace(keyword, replacement)

    # Add a delay to show the loading message
    time.sleep(0.5)

# Show completion message
print("All replacements completed.")

# Save the new file
ppt.SaveAs(str(output_file))
ppt.Close()

# Quit PowerPoint
ppt_app.Quit()
