# -*- coding: utf-8 -*-
"""
Created on Sat Nov  9 17:33:09 2024

@author: TENGKU
"""

import paramiko
import pandas as pd
import getpass
import time
import tkinter as tk
from tkinter import filedialog
import os

def choose_file(dialog_title):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=dialog_title,
        filetypes=[("Excel files", "*.xlsx")]
    )
    return file_path

def choose_output_directory(dialog_title):
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title=dialog_title)
    return directory

def send_command_and_wait(shell, command, wait_time=2):
    shell.send(command + '\n')
    time.sleep(wait_time)  # Initial wait
    
    # Continue receiving data until no more data is available
    output = ''
    while True:
        time.sleep(1)  # Additional small delay between checks
        if shell.recv_ready():
            chunk = shell.recv(65535).decode()
            output += chunk
        else:
            # Check one more time after a short delay
            time.sleep(2)
            if not shell.recv_ready():
                break
    
    return output

def process_batch(batch_routers, username, password, output_directory, results):
    failure_count = 0
    
    for _, row in batch_routers.iterrows():
        ne_name = row['NE Name']
        ip_address = row['IP Address']
        
        try:
            # Initialize SSH client
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            # Connect to the router
            ssh.connect(ip_address, username=username, password=password, timeout=30)
            print(f"Connected to {ne_name} ({ip_address})")
            
            # Create SSH shell
            shell = ssh.invoke_shell()
            time.sleep(2)  # Wait for shell to initialize
            
            # Execute 's 0 t' command first and capture its output
            print(f"Setting screen length to 0 for {ne_name}...")
            screen_output = send_command_and_wait(shell, 's 0 t', 3)  # Increased wait time
            
            # Collect display isis lsdb verbose
            print(f"Collecting ISIS LSDB from {ne_name}...")
            print(f"This may take several minutes...")
            isis_output = send_command_and_wait(shell, 'display isis lsdb verbose', 30)  # Increased wait time significantly

            # Save outputs to file using NE Name
            log_file_path = os.path.join(output_directory, f"{ne_name}.log")
            with open(log_file_path, 'w', encoding='utf-8') as f:
                # Write the device name and command headers
                f.write(f"<{ne_name}>\n")
                f.write("=== DISPLAY ISIS LSDB VERBOSE ===\n")
                f.write("display isis lsdb verbose\n")
                
                # Process the ISIS output to remove extra blank lines and normalize spacing
                isis_lines = isis_output.split('\n')
                processed_lines = []
                prev_line_empty = False
                
                for line in isis_lines:
                    # Skip the command echo line
                    if line.strip() == 'display isis lsdb verbose':
                        continue
                    
                    # Remove extra spaces at the end of lines
                    current_line = line.rstrip()
                    
                    # Handle empty lines
                    if not current_line:
                        if not prev_line_empty:  # Only keep one empty line
                            processed_lines.append('')
                            prev_line_empty = True
                    else:
                        processed_lines.append(current_line)
                        prev_line_empty = False
                
                # Write the processed lines
                f.write('\n'.join(processed_lines))

            # Log success
            results.append({
                'NE Name': ne_name,
                'IP Address': ip_address,
                'Result': "Success",
                'Output File': log_file_path
            })
            
            # Close the connection
            ssh.close()
            print(f"Successfully collected data from {ne_name}")
            
            # Add delay between devices
            time.sleep(5)  # Wait between devices

        except paramiko.AuthenticationException as e:
            failure_count += 1
            print(f"Authentication failed for {ne_name} ({ip_address})")
            results.append({
                'NE Name': ne_name,
                'IP Address': ip_address,
                'Result': "Authentication Failed",
                'Output File': 'N/A'
            })
            if failure_count >= 7:
                return False
                
        except Exception as e:
            failure_count += 1
            print(f"Failed to connect or execute on {ne_name} ({ip_address}): {str(e)}")
            results.append({
                'NE Name': ne_name,
                'IP Address': ip_address,
                'Result': "Failed",
                'Output File': 'N/A'
            })
            if failure_count >= 7:
                return False
    
    return True

def main():
    # Prompt the user to select the NE List Excel file
    input_file_path = choose_file("Select the NE List Excel file")

    # Prompt for output directory for log files
    output_directory = choose_output_directory("Select directory to save log files")

    # Prompt for the status report Excel file location
    status_file_path = filedialog.asksaveasfilename(
        title="Specify the location and name to save the status report Excel file",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    # Ensure user selected valid files and directory
    if not input_file_path or not output_directory or not status_file_path:
        print("You need to select input file, output directory, and status file location.")
        return

    # Load the router details from Excel
    router_details = pd.read_excel(input_file_path, sheet_name=0)

    # Prompt for username and password
    username = input("Enter your SSH username: ")
    password = getpass.getpass("Enter your SSH password: ")

    # Prepare to store results
    results = []

    # Process routers in batches of 7
    batch_size = 7
    total_routers = len(router_details)

    for i in range(0, total_routers, batch_size):
        batch = router_details.iloc[i:min(i+batch_size, total_routers)]
        print(f"\nProcessing batch {(i//batch_size)+1} of {(total_routers+batch_size-1)//batch_size}")
        
        # Process the batch
        batch_success = process_batch(batch, username, password, output_directory, results)
        
        if not batch_success:
            print("\nWARNING: 7 or more failures detected in this batch!")
            continue_prompt = input("Do you want to continue with the next batch? (yes/no): ").lower()
            
            if continue_prompt != 'yes':
                print("Operation cancelled by user.")
                break
            else:
                print("Continuing with next batch...")

    # Save the results to Excel file
    results_df = pd.DataFrame(results)
    results_df.to_excel(status_file_path, sheet_name='Connection Status', index=False)

    print(f"\nStatus report saved to {status_file_path}")
    print(f"Log files saved in {output_directory}")

if __name__ == "__main__":
    main()
