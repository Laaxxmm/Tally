import socket
import sys

def check_remote(host, port):
    print(f"Checking connection to REMOTE Tally at {host}:{port}...")
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(3)
    try:
        s.connect((host, port))
        print(f"✅ Successfully connected to {host}:{port}")
        s.close()
        return True
    except socket.error as e:
        print(f"❌ Failed to connect to {host}:{port}")
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    host = "172.16.1.121"
    port = 8501
    if check_remote(host, port):
        print(f"\nRemote Tally is reachable on port {port}!")
    else:
        print("\nRemote Tally is NOT reachable.")
        print("Possible causes:")
        print("1. Firewall on the Windows laptop is blocking port 9000.")
        print("2. Tally is not configured to listen on the network interface.")
        print("   (Check Tally Config > Connectivity > Client/Server > Port)")
