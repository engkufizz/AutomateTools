import os
import tkinter as tk
from tkinter import filedialog

def find_string_in_top_directory(directory, search_string, output_file, file_extensions, match_whole_word=False):
    """Searches for a string ONLY in the top-level of a given directory."""
    files_with_string = []
    print(f"Searching for '{search_string}' in '{directory}' only (no subfolders)...")

    for filename in os.listdir(directory):
        if filename.endswith(file_extensions):
            file_path = os.path.join(directory, filename)
            # Ensure we are checking a file, not a directory
            if os.path.isfile(file_path):
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                        for line in file:
                            if (search_string.lower() in line.lower()):
                                files_with_string.append(file_path)
                                break # Move to the next file once found
                except Exception as e:
                    print(f"Error processing {file_path}: {e}")
    
    write_results(output_file, search_string, directory, files_with_string, recursive=False)


def find_string_in_files_recursively(directory, search_string, output_file, file_extensions, match_whole_word=False):
    """Searches for a string in a directory and ALL its subdirectories."""
    files_with_string = []
    print(f"Searching for '{search_string}' in '{directory}' and all subfolders...")

    for dirpath, _, filenames in os.walk(directory):
        for filename in filenames:
            if filename.endswith(file_extensions):
                file_path = os.path.join(dirpath, filename)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                        for line in file:
                            if (search_string.lower() in line.lower()):
                                files_with_string.append(file_path)
                                break # Move to the next file once found
                except Exception as e:
                    print(f"Error processing {file_path}: {e}")

    write_results(output_file, search_string, directory, files_with_string, recursive=True)


def write_results(output_file, search_string, directory, file_list, recursive):
    """Writes the list of found files to the output file."""
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            if file_list:
                f.write(f"Files containing the string '{search_string}':\n\n")
                for file_path in file_list:
                    f.write(file_path + '\n')
            else:
                search_area = f"'{directory}' and its subdirectories" if recursive else f"'{directory}'"
                f.write(f"No files found in {search_area} containing the string '{search_string}'.\n")
        print(f"\nSearch complete. Results saved to {output_file}.")
    except Exception as e:
        print(f"Error writing to output file {output_file}: {e}")


# --- Main Script Execution ---

# 1. Set up and hide the root Tkinter window
root = tk.Tk()
root.withdraw()

# 2. Prompt user to select a directory
print("A folder selection window will now open.")
directory = filedialog.askdirectory(title="Select the folder to search")

# 3. Proceed only if a directory was selected
if directory:
    print(f"Directory selected: {directory}\n")

    # 4. Ask user whether to include subfolders
    include_subfolders_choice = input("Include subfolders in the search? (yes/no): ").lower().strip()
    
    # 5. Get the search string from the user
    search_string = input("Please enter the string to search for: ")

    if search_string:
        allowed_extensions = ('.log', '.txt', '.cfg')
        output_file = 'result.txt'
        
        # 6. Call the appropriate function based on user's choice
        if include_subfolders_choice.startswith('y'):
            find_string_in_files_recursively(directory, search_string, output_file, allowed_extensions)
        else:
            find_string_in_top_directory(directory, search_string, output_file, allowed_extensions)
    else:
        print("No search string entered. Exiting programme.")
else:
    print("No directory was selected. Exiting programme.")
