import socket

def identify_service(host, port):
    print(f"🕵️‍♀️ Identifying service at {host}:{port}...")
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(3)
        s.connect((host, port))
        
        # Send a generic HTTP GET
        s.send(b"GET / HTTP/1.1\r\nHost: test\r\n\r\n")
        response = s.recv(1024).decode('utf-8', errors='ignore')
        s.close()
        
        print("\n--- Response Headers ---")
        print(response[:500])
        print("------------------------\n")
        
        if "Streamlit" in response or "streamlit" in response:
            print("🚨 DIAGNOSIS: This is a Streamlit App!")
            print("You cannot connect the Dashboard to ITSELF.")
            print("You must connect to Tally Prime (usually Port 9000).")
        elif "Tally" in response:
            print("✅ DIAGNOSIS: This looks like Tally.")
        else:
            print("❓ DIAGNOSIS: Unknown service. Check headers above.")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    identify_service("172.16.1.121", 8501)
