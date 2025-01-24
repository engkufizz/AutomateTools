import os

def find_string_in_log_files(directory, search_string, output_file, match_whole_word=False):
    # List to store the names of files containing the search string
    files_with_string = []

    # Iterate over all files in the specified directory
    for filename in os.listdir(directory):
        # Check if the file is a .log file
        if filename.endswith('.log'):
            file_path = os.path.join(directory, filename)
            try:
                # Open the .log file
                with open(file_path, 'r', encoding='utf-8') as file:
                    found = False

                    # Iterate over each line in the file
                    for line in file:
                        line_text = line.lower()
                        search_text = search_string.lower()

                        if match_whole_word:
                            # Split the line into words and check for exact matches
                            words = line_text.split()
                            if search_text in words:
                                found = True
                                break
                        else:
                            if search_text in line_text:
                                found = True
                                break

                if found:
                    files_with_string.append(filename)

            except Exception as e:
                print(f"Error reading {filename}: {e}")

    # Write the results to the output file
    with open(output_file, 'w') as f:
        for file in files_with_string:
            f.write(file + '\n')

# Define the directory containing the .log files
directory = r'<ENTER_DIRECTORY_PATH_HERE>'
# Define the string to search for
search_string = '<ENTER_SEARCH_STRING_HERE>'
# Define the output file
output_file = 'result.txt'

# Call the function
find_string_in_log_files(directory, search_string, output_file, match_whole_word=False)

print(f"Search complete. Results saved to {output_file}.")
