# Import the pyfiglet module for ASCII art text
import pyfiglet

# Define the name and URL
name = "TENGKU"
url = "https://tengkulist.web.app/"

# Use the pyfiglet module to create ASCII art text for the name
name_ascii = pyfiglet.figlet_format(name)

# Combine the ASCII art text for the name and URL into a single string
combined_ascii = name_ascii + "\n" + url + "\n"

# Print the combined ASCII art text
print(combined_ascii)
