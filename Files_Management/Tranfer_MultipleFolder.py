import os
import shutil

# List of folders to be moved
folders = ['folder1', 'folder2', 'folder3', 'folder4', 'folder5', 'folder6', 'folder7', 'folder8']

# Destination folder
destination = 'output'

# Create the destination folder if it doesn't exist
if not os.path.exists(destination):
    os.makedirs(destination)

# Iterate over the folders
for folder in folders:
    # Search for folders that contain the provided input
    matching_folders = [f for f in os.listdir('.') if folder.lower() in f.lower()]

    if matching_folders:
        # Move each matching folder to the destination
        for matching_folder in matching_folders:
            shutil.move(matching_folder, destination)
    else:
        print(f'No folder found containing "{folder}"')
