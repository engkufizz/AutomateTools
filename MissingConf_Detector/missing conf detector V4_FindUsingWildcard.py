import os
import shutil
import chardet
import re

# Path to the folder containing the configuration files
folder_path = 'Input'

# Path to the folder where we will save the files with missing configuration
output_folder_path = 'Output'

# Regular expression pattern to look for in the files
config_pattern = r'acl\s+.*\sinbound'

# Compile the regular expression pattern
pattern = re.compile(config_pattern)

# Loop through all files and subfolders in the folder
for root, dirs, files in os.walk(folder_path):
    # Loop through all files in the current directory
    for filename in files:
        # Check if the file has the '.txt' extension
        if filename.endswith('.txt'):
            # Construct the full path to the file
            file_path = os.path.join(root, filename)
            # Open the file
            with open(file_path, 'rb') as f:
                # Read the contents of the file and detect the encoding
                file_contents = f.read()
                result = chardet.detect(file_contents)
                # Decode the file contents using the detected encoding
                file_contents = file_contents.decode(result['encoding'], errors='ignore')
                # Check if the regular expression pattern is in the file
                if not re.search(pattern, file_contents):
                    # If not, copy the file to the output folder
                    shutil.copy(file_path, output_folder_path)
