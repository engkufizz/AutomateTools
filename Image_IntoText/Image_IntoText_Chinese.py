import easyocr
from PIL import Image
import numpy as np

# Install easyocr using pip
# pip install easyocr

# Define the OCR reader and specify the language(s) to use
reader = easyocr.Reader(['ch_sim'])  # Use 'ch_sim' for simplified Chinese

# Open the image file
image_path = r'image.jpg'
image = Image.open(image_path)

# Convert the image to a NumPy array
image_np = np.array(image)

# Perform OCR on the image
result = reader.readtext(image_np)

# Extract the text from the result
text = ' '.join([res[1] for res in result])

# Print the extracted text
print(text)
