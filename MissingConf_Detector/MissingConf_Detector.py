import os
import shutil

# Path to the folder containing the configuration files
folder_path = 'C:/Users/engkufizz/Desktop/MissingConf_Detector/Input'

# Path to the folder where we will save the files with missing configuration
output_folder_path = 'C:/Users/engkufizz/Desktop/MissingConf_Detector/Output'

# Configuration string to look for in the files
config_string = 'ip ip-prefix XXX_IP_Filter index XXX permit 172.X.X.X 27 greater-equal 27 less-equal 32'

# Loop through all files and subfolders in the folder
for root, dirs, files in os.walk(folder_path):
  # Loop through all files in the current directory
  for filename in files:
    # Check if the file has the '.txt' extension
    if filename.endswith('.txt'):
      # Construct the full path to the file
      file_path = os.path.join(root, filename)
      # Open the file
      with open(file_path, encoding='utf-8') as f:
        # Read the contents of the file
        file_contents = f.read()
        # Check if the configuration string is in the file
        if config_string not in file_contents:
          # If not, copy the file to the output folder
          shutil.copy(file_path, output_folder_path)
