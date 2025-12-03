import socket

def scan_tally():
    hosts = ["127.0.0.1", "localhost"]
    ports = range(9000, 9010)
    
    print("🔍 Scanning for Tally instance...")
    
    found = False
    for host in hosts:
        for port in ports:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.5)
            try:
                result = s.connect_ex((host, port))
                if result == 0:
                    print(f"\n✅ FOUND Tally at {host}:{port}")
                    found = True
                    # Try to send a simple handshake to confirm it's HTTP
                    try:
                        s.send(b"GET / HTTP/1.1\r\nHost: localhost\r\n\r\n")
                        resp = s.recv(100)
                        print(f"   Response: {resp.decode('utf-8', errors='ignore').strip()}")
                    except:
                        pass
                s.close()
            except:
                pass
                
    if not found:
        print("\n❌ Could not find Tally on ports 9000-9010.")
        print("Please check Tally Configuration > Advanced Configuration > Connectivity.")

if __name__ == "__main__":
    scan_tally()
