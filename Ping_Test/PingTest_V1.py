import subprocess
import re
import speedtest

hosts = ['google.com', 'facebook.com', 'twitter.com']
latencies = {}
st = speedtest.Speedtest()

for host in hosts:
    ping = subprocess.Popen(
        ['ping', '-c', '1', '-w', '100', host],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    out, error = ping.communicate()
    out = out.decode('utf-8')
    match = re.search(r'time=([\d.]+) ms|Minimum = (\d+)ms, Maximum = (\d+)ms, Average = (\d+)ms', out)
    if match:
        time = match.group(1) or match.group(4)
        latencies[host] = float(time)
    else:
        print(f"No match found for {host}: {out}")

print("\nLatencies:")
sorted_latencies = sorted(latencies.items(), key=lambda x: x[1])
for host, latency in sorted_latencies:
    print(f"{host.ljust(15)} : {latency:>8.3f} ms")

print("\nSpeed Test Results:")
st.get_best_server()
download_speed = st.download() / 1_000_000
upload_speed = st.upload() / 1_000_000
server_info = st.results.server
server_name = server_info['sponsor']
print(f"Download speed: {download_speed:.2f} Mbit/s")
print(f"Upload speed: {upload_speed:.2f} Mbit/s")
print(f"Server: {server_name}")

