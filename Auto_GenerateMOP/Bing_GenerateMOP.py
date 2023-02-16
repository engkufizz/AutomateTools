# Import the docx module
import docx

# Define a function that takes a document and a dictionary of words to replace
def replace_words(doc, word_dict):
  # Loop through all the paragraphs in the document
  for paragraph in doc.paragraphs:
    # Loop through all the runs in the paragraph
    for run in paragraph.runs:
      # Loop through all the words to replace
      for old_word, new_word in word_dict.items():
        # Replace the old word with the new word in the run text
        run.text = run.text.replace(old_word, new_word)
  # Return the modified document
  return doc

# Open the document Data.docx
doc = docx.Document("Data.docx")

# Define a dictionary of words to replace
# For example, replace "apple" with "banana" and "red" with "blue"
word_dict = {"Hasrul": "Badrul", "XX Company": "Maxis Company"}

# Call the function to replace the words
new_doc = replace_words(doc, word_dict)

# Save the new document as New_Data.docx
new_doc.save("New_Data.docx")