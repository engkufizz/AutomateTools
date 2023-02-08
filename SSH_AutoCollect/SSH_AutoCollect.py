import paramiko

# Connect to the network element
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect('hostname', port=22, username='username', password='password')

# Execute the command
stdin, stdout, stderr = ssh.exec_command('uname -r')

# Save the logs
with open('logs.txt', 'w') as f:
    f.write(stdout.read().decode())

# Close the connection
ssh.close()
