import os
import shutil
import chardet

# Path to the folder containing the configuration files
folder_path = 'Input'

# Path to the folder where we will save the files with the configuration string
output_folder_path = 'Output'

# Configuration strings to look for in the files
config_strings = [
'sysname WWLSWBESI015',
'sysname BWLSWSHTI001',
'sysname BWLSWSHTI015',
'sysname BWUPEMEA001',
'sysname WWUPEASTRO02',
'sysname WWUPEASTRO01',
'sysname BWLSWSHTI002',
'sysname BWLSWSHTI003',
'sysname BWLSWSHTI007',
'sysname BWLSWSHTI008',
'sysname BWLSWSHTI013',
'sysname BWLSWSHTI014',
'sysname BWLSWSHTI016',
'sysname WWLSWBESI007',
'sysname WWLSWBESI008',
'sysname WWLSWBESI016',
'sysname WWLSWBESI009',
'sysname WWLSWBESI002',
'sysname WWLSWBESI003',
'sysname BWLSWSHTI018', 
'sysname BWLSWSHTI017']

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
                # Check if any of the configuration strings is in the file
                if any(config_string in file_contents for config_string in config_strings):
                    # If so, copy the file to the output folder
                    shutil.copy(file_path, output_folder_path)
