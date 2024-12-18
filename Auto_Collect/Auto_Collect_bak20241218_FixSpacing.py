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
import re

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

def get_sysname_from_output(output):
    # Search for sysname in the configuration
    match = re.search(r'sysname\s+(\S+)', output)
    if match:
        return match.group(1)
    return None

def send_command_and_wait(shell, command, wait_time=2):
    shell.send(command + '\n')
    time.sleep(wait_time)
    output = ''
    while shell.recv_ready():
        output += shell.recv(65535).decode()
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
            ssh.connect(ip_address, username=username, password=password)
            print(f"Connected to {ne_name} ({ip_address})")
            
            # Create SSH shell
            shell = ssh.invoke_shell()
            
            # Execute 's 0 t' command first and capture its output
            print(f"Setting screen length to 0 for {ne_name}...")
            screen_output = send_command_and_wait(shell, 's 0 t', 1)
            
            # Collect display current-configuration
            print(f"Collecting configuration from {ne_name}...")
            config_output = send_command_and_wait(shell, 'display current-configuration', 5)

            # Collect display interface brief
            print(f"Collecting interface brief from {ne_name}...")
            interface_output = send_command_and_wait(shell, 'display interface brief', 2)

            # Get sysname from configuration
            sysname = get_sysname_from_output(config_output)
            if not sysname:
                sysname = ne_name  # Use NE name from Excel if sysname not found

            # Save outputs to file
            log_file_path = os.path.join(output_directory, f"{sysname}.log")
            with open(log_file_path, 'w', encoding='utf-8') as f:
                # Write device name
                f.write(f"<{sysname}>\n")
                
                # Process and write screen length output
                f.write("=== SCREEN LENGTH SETTING ===\n")
                screen_lines = screen_output.split('\n')
                processed_screen = '\n'.join(line.rstrip() for line in screen_lines if line.strip())
                f.write(f"{processed_screen}\n\n")
                
                # Process and write configuration output
                f.write("=== DISPLAY CURRENT-CONFIGURATION ===\n")
                config_lines = config_output.split('\n')
                processed_config = []
                prev_empty = False
                
                for line in config_lines:
                    line = line.rstrip()
                    if line.strip():
                        processed_config.append(line)
                        prev_empty = False
                    elif not prev_empty:
                        processed_config.append('')
                        prev_empty = True
                f.write('\n'.join(processed_config))
                f.write('\n\n')
                
                # Process and write interface brief output
                f.write("=== DISPLAY INTERFACE BRIEF ===\n")
                interface_lines = interface_output.split('\n')
                processed_interface = []
                prev_empty = False
                
                for line in interface_lines:
                    line = line.rstrip()
                    if line.strip():
                        processed_interface.append(line)
                        prev_empty = False
                    elif not prev_empty:
                        processed_interface.append('')
                        prev_empty = True
                f.write('\n'.join(processed_interface))

            # Log success
            results.append({
                'NE Name': ne_name,
                'IP Address': ip_address,
                'Sysname': sysname,
                'Result': "Success",
                'Output File': log_file_path
            })
            
            # Close the connection
            ssh.close()
            print(f"Successfully collected data from {ne_name}")

        except paramiko.AuthenticationException as e:
            failure_count += 1
            print(f"Authentication failed for {ne_name} ({ip_address})")
            results.append({
                'NE Name': ne_name,
                'IP Address': ip_address,
                'Sysname': 'N/A',
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
                'Sysname': 'N/A',
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
