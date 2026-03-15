
import sys
import os
import time

try:
    print("Importing ctos_app...")
    from app.gui import CTOSApp
    print("Import successful.")
except ImportError as e:
    print(f"Import failed: {e}")
    sys.exit(1)
except Exception as e:
    print(f"Error during import: {e}")
    sys.exit(1)

try:
    print("Testing TaskHistoryManager...")
    from app.history import TaskHistoryManager
    history = TaskHistoryManager()
    print("History Manager initialized.")
    
    session_id = history.create_session("Test Session", 1, ["Test Company"])
    print(f"Created session with ID: {session_id}")
    
    sessions = history.get_all_sessions()
    print(f"Sessions found: {len(sessions)}")
    
    results = history.get_session_results(session_id)
    print(f"Results for session: {len(results)}")
    print(f"First result status: {results[0]['status']}")
    
    print("Verification successful.")
except Exception as e:
    print(f"History Manager test failed: {e}")
    sys.exit(1)
