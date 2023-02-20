# Import the time module to simulate a delay
import time

# Import the pyfiglet module for ASCII art text
import pyfiglet

# Define the name and URL
name = "TENGKU"
url = "https://tengkulist.web.app/"

# Use the pyfiglet module to create ASCII art text for the name
name_ascii = pyfiglet.figlet_format(name)

# Combine the ASCII art text for the name and URL into a single string
combined_ascii = name_ascii + "\n" + url + "\n"

# Define the length of the loading bar and the delay between updates
loading_bar_length = 20
loading_bar_delay = 0.1

# Define the characters to use for the loading bar
loading_chars = ["|", "/", "-", "\\"]

# Print the initial ASCII art text
print(combined_ascii)

# Simulate a delay while the script is executing
for i in range(loading_bar_length):
    # Define the loading bar string
    loading_bar = f"[{loading_chars[i % len(loading_chars)]}{' ' * (loading_bar_length - i - 1)}]"
    
    # Print the loading bar and move the cursor back to the beginning of the line
    print(f"\r{loading_bar} buffering...", end="")
    
    # Wait for a short period of time
    time.sleep(loading_bar_delay)

# Print a message indicating that the script has completed executing
print("\r" + " " * (loading_bar_length + 13) + "\r" + "Script completed!")
