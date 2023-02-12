from pathlib import Path  # core python module
import win32com.client  # pip install pywin32

# Path settings
current_dir = Path.cwd()
input_file = current_dir / "Data.docx"
output_file = current_dir / "New_Data.docx"

# Replace keywords and their replacements
replacements = [("2022", "2023"), ("keyword2", "replacement2"), ("keyword3", "replacement3")]
wd_replace = 2  # 2=replace all occurrences, 1=replace one occurrence, 0=replace no occurrences
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached

# Open Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

# Replace all keywords
for doc_file in Path(input_dir).rglob("*.doc*"):
    # Open each document and replace string
    word_app.Documents.Open(str(doc_file))
    # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
    word_app.Selection.Find.Execute(
        FindText=find_str,
        ReplaceWith=replace_with,
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
    for i in range(word_app.ActiveDocument.Shapes.Count):
        if word_app.ActiveDocument.Shapes(i + 1).TextFrame.HasText:
            words = word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words
            # Loop through each word. This method preserves formatting.
            for j in range(words.Count):
                # If a word exists, replace the text of it, but keep the formatting.
                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with

# Save the new file
output_path = Path.cwd() / "New_Data.docx"
word_app.ActiveDocument.SaveAs(str(output_path))
word_app.ActiveDocument.Close(SaveChanges=False)
word_app.Application.Quit()
