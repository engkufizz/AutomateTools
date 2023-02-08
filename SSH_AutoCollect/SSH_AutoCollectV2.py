import paramiko

# Read the list of network elements from the text file
with open('NElist.txt', 'r') as f:
    lines = f.readlines()

# Connect to each network element
for line in lines:
    hostname, port, username, password = line.strip().split(',')
    port = int(port)

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname, port=port, username=username, password=password)

    # Execute the command
    stdin, stdout, stderr = ssh.exec_command('uname -r')

    # Save the logs to a separate file based on the hostname
    with open(f'{hostname}_logs.txt', 'w') as f:
        f.write(stdout.read().decode())

    # Close the connection
    ssh.close()
