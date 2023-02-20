from pathlib import Path  # core python module
import win32com.client  # pip install pywin32
import ast  # core python module
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
input_file = current_dir / "Data.docx"
output_file = current_dir / "New_Data.docx"
changelist_file = current_dir / "Changelist.txt"

# Read the keywords and replacements from the text file
with open(changelist_file, "r") as f:
    replacements = ast.literal_eval(f.read())

# Replace keywords and their replacements
wd_replace = 2  # 2=replace all occurrences, 1=replace one occurrence, 0=replace no occurrences
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached

# Open Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

# Open the input document and replace keywords
word_app.Documents.Open(str(input_file))

# Loop through the list of replacements
for keyword, replacement in replacements:
    # Show loading message
    print(f"Replacing '{keyword}' with '{replacement}'...")

    # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
    word_app.Selection.Find.Execute(
        FindText=keyword,
        ReplaceWith=replacement,
        Replace=wd_replace,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )

    # Replace the keywords in shapes
    # VBA SO reference: https://stackoverflow.com/a/26266598
    # Loop through all the shapes
    for i, shape in enumerate(word_app.ActiveDocument.Shapes):
        if shape.TextFrame.HasText:
            words = shape.TextFrame.TextRange.Words
            # Loop through each word. This method preserves formatting.
            for j, word in enumerate(words):
                # If a word exists, replace the text of it, but keep the formatting.
                if word.Text == keyword:
                    word.Text = replacement

    # Add a delay to show the loading message
    time.sleep(0.5)

# Show completion message
print("All replacements completed.")

# Save the new file
word_app.ActiveDocument.SaveAs(str(output_file))
word_app.ActiveDocument.Close(SaveChanges=False)

# Quit Word
word_app.Application.Quit()
