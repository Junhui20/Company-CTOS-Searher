
import sqlite3
import json
import os
from app.history import TaskHistoryManager

DB_FILE = "scrape_history.db"

def verify_persistence():
    print("Initializing History Manager...")
    mgr = TaskHistoryManager()
    
    # 1. Create a fake session
    session_id = mgr.create_session("Test Session", 1, ["Test Company Ltd"])
    print(f"Created Session {session_id}")
    
    # 2. Simulate saving a result with CONFIDENCE
    data_payload = {
        "name": "Test Company Ltd",
        "reg_no": "12345",
        "confidence": "High (Auto Test)" 
    }
    
    print("Updating result with confidence data...")
    mgr.update_result(session_id, "Test Company Ltd", "Found", "12345", data_payload)
    
    # 3. Retrieve and check
    print("Retrieving results...")
    results = mgr.get_session_results(session_id)
    
    if not results:
        print("FAILED: No results result.")
        return
        
    item = results[0]
    confidence = item['data'].get('confidence')
    
    print(f"Retrieved Confidence: '{confidence}'")
    
    if confidence == "High (Auto Test)":
        print("SUCCESS: Confidence value was preserved!")
    else:
        print(f"FAILED: Expected 'High (Auto Test)', got '{confidence}'")

    mgr.close()

if __name__ == "__main__":
    verify_persistence()
