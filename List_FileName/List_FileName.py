import os

input_folder = "Input" # Replace with the path to your input folder
output_file = "namelist.txt"

with open(output_file, "w") as f:
    for file_name in os.listdir(input_folder):
        f.write(file_name + "\n")
