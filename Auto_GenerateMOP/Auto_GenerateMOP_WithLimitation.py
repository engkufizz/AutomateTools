import re
import docx

def replace_words(doc_name, new_doc_name, replace_dict):
    doc = docx.Document(doc_name)
    
    for para in doc.paragraphs:
        for key, value in replace_dict.items():
            para.text = re.sub(r"\b" + re.escape(key) + r"\b", value, para.text)
                    
    doc.save(new_doc_name)

replace_dict = {
    'old_word_1': 'new_word_1',
    'old_word_2': 'new_word_2',
    'old_word_3': 'new_word_3',
}

replace_words('Data.docx', 'New_Data.docx', replace_dict)
