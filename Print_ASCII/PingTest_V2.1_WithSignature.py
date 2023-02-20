import subprocess
import re
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

hosts = ['google.com', 'facebook.com', 'twitter.com']
latencies = {}

# Print a message indicating that the program is currently loading
print("Loading...\n")

# Ping each host and calculate the average latency
for host in hosts:
    if subprocess.call('ping -n 1 -w 100 ' + host, stdout=subprocess.PIPE, shell=True) == 0:
        ping = subprocess.Popen(
            ['ping', '-n', '4', '-w', '100', host],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        out, error = ping.communicate()
        out = out.decode('utf-8')
        match = re.search(r'TTL=\d+', out)
        if match:
            times = re.findall(r'time=(\d+)ms', out)
            avg_time = sum([int(t) for t in times]) / len(times)
            latencies[host] = avg_time
        else:
            print(f"No match found for {host}: {out}")
    else:
        print(f"Could not reach {host}")

sorted_latencies = sorted(latencies.items(), key=lambda x: x[1])

# Print the latency results as an ASCII graph
print("\nLatencies:")
for host, latency in sorted_latencies:
    print(f"{host.ljust(15)} : {'#' * int(latency/5)} {latency:.3f} ms")

# Run a speedtest
try:
    # Print a message indicating that the speedtest is currently running
    print("\nRunning speedtest...")
    
    import speedtest
    st = speedtest.Speedtest()
    st.get_best_server()
    download_speed = st.download() / 1_000_000
    upload_speed = st.upload() / 1_000_000
    server_name = st.results.server['sponsor']
    print("\nSpeed Test Results:")
    print(f"Download speed: {download_speed:.2f} Mbit/s")
    print(f"Upload speed: {upload_speed:.2f} Mbit/s")
    print(f"Server: {server_name}")
except ImportError:
    print("\nSpeedtest module not installed.")
