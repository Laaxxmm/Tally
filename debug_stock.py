
import sys
import os
from datetime import date

# Add src to path
sys.path.append(os.path.join(os.getcwd(), "src"))

from tally_client import fetch_group_balance, _post_xml, _clean_tally_xml
import xml.etree.ElementTree as ET

HOST = "172.16.1.121" # User's remote host
PORT = 9000

def list_groups(host, port):
    print(f"Fetching groups from {host}:{port}...")
    xml = """<ENVELOPE>
      <HEADER>
        <VERSION>1</VERSION>
        <TALLYREQUEST>Export</TALLYREQUEST>
        <TYPE>Collection</TYPE>
        <ID>ListofGroups</ID>
      </HEADER>
      <BODY>
        <DESC>
          <STATICVARIABLES>
            <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
          </STATICVARIABLES>
          <TDL>
            <TDLMESSAGE>
              <COLLECTION NAME="ListofGroups">
                <TYPE>Group</TYPE>
                <FETCH>Name, Parent</FETCH>
              </COLLECTION>
            </TDLMESSAGE>
          </TDL>
        </DESC>
      </BODY>
    </ENVELOPE>"""
    
    try:
        raw = _clean_tally_xml(_post_xml(xml, host, port))
        root = ET.fromstring(raw)
        groups = [g.get("NAME") for g in root.findall(".//GROUP")]
        print(f"Found {len(groups)} groups.")
        if "Stock-in-Hand" in groups:
            print("SUCCESS: 'Stock-in-Hand' group found.")
        else:
            print("WARNING: 'Stock-in-Hand' group NOT found.")
            print("Available groups:", groups)
            
        return groups
    except Exception as e:
        print(f"Error listing groups: {e}")
        return []

def test_stock_fetch(host, port):
    print("\nTesting Stock Fetch...")
    # Test for a likely date
    test_date = date(2024, 3, 31)
    print(f"Fetching 'Stock-in-Hand' for {test_date}...")
    val = fetch_group_balance("Any", "Stock-in-Hand", test_date, host, port)
    print(f"Value returned: {val}")
    
    # Try fetching "Current Assets" just to see if *any* balance comes back
    print(f"Fetching 'Current Assets' for {test_date}...")
    val_ca = fetch_group_balance("Any", "Current Assets", test_date, host, port)
    print(f"Value returned for Current Assets: {val_ca}")

if __name__ == "__main__":
    list_groups(HOST, PORT)
    test_stock_fetch(HOST, PORT)
