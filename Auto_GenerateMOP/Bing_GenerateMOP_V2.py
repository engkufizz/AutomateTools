from pathlib import Path  # core python module
import win32com.client  # pip install pywin32

# Path settings
current_dir = Path.cwd()
input_file = current_dir / "Data.docx"
output_file = current_dir / "New_Data.docx"

# Ask the user how many replacements they want to do
while True:
    try:
        num_replacements = int(input("How many replacements do you want to do? "))
        if num_replacements > 0:
            break
        else:
            print("Please enter a positive number.")
    except ValueError:
        print("Please enter a valid number.")

# Create a dictionary of keywords and replacements
replacements = {}
for i in range(num_replacements):
    keyword = input(f"What is the keyword for replacement {i + 1}? ")
    replacement = input(f"What is the replacement for keyword {keyword}? ")
    replacements[keyword] = replacement

# Replace keywords and their replacements
wd_replace = 2  # 2=replace all occurrences, 1=replace one occurrence, 0=replace no occurrences
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached

# Open Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

# Open the input document and replace keywords
word_app.Documents.Open(str(input_file))
# Loop through the dictionary of replacements
for keyword, replacement in replacements.items():
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

# Save the new file
word_app.ActiveDocument.SaveAs(str(output_file))
word_app.ActiveDocument.Close(SaveChanges=False)

# Quit Word
word_app.Application.Quit()