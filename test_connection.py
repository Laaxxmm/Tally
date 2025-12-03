import socket
import sys

def check_port(host, port):
    print(f"Checking connection to {host}:{port}...")
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(2)
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
    host = "127.0.0.1"
    port = 9000
    if check_port(host, port):
        print("\nTally appears to be running and reachable.")
    else:
        print("\nTally is NOT reachable. Please ensure:")
        print("1. Tally Prime is running.")
        print("2. 'Enable ODBC/HTTP' is set to 'Yes' in Tally Configuration.")
        print(f"3. Port is set to {port} in Tally Configuration.")
