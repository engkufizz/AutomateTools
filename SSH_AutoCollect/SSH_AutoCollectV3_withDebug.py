import paramiko
import logging

logging.basicConfig(filename='automation.log', level=logging.INFO,
                    format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

# Read the list of network elements from the text file
with open('NElist.txt', 'r') as f:
    lines = f.readlines()

# Connect to each network element
for line in lines:
    hostname, port, username, password = line.strip().split(',')
    port = int(port)

    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname, port=port, username=username, password=password, timeout=10)

        # Execute the command
        stdin, stdout, stderr = ssh.exec_command('dis health', timeout=10)

        # Save the logs to a separate file based on the hostname
        with open(f'{hostname}_logs.txt', 'w') as f:
            f.write(stdout.read().decode())

    except Exception as e:
        logging.error(f'Error connecting to {hostname}: {e}')

    finally:
        # Close the connection
        ssh.close()
