import sqlite3
import datetime
import json
import threading

import os
import sys


def _get_app_dir():
    """Get the directory for app data files (works both in dev and as bundled exe)."""
    if getattr(sys, 'frozen', False):
        # Running as bundled exe - store DB next to the exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


DB_FILE = os.path.join(_get_app_dir(), "scrape_history.db")


class TaskHistoryManager:
    def __init__(self):
        self.conn = None
        self._lock = threading.Lock()
        self._init_db()

    def _init_db(self):
        self.conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        c = self.conn.cursor()

        c.execute(
            """CREATE TABLE IF NOT EXISTS sessions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT,
                        created_at TEXT,
                        total_items INTEGER,
                        status TEXT
                    )"""
        )

        c.execute(
            """CREATE TABLE IF NOT EXISTS results (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        session_id INTEGER,
                        company_input TEXT,
                        status TEXT,
                        reg_no TEXT,
                        result_json TEXT,
                        updated_at TEXT,
                        FOREIGN KEY(session_id) REFERENCES sessions(id)
                    )"""
        )
        self.conn.commit()

    def create_session(self, name, total_items, item_list):
        """Creates a new session and pre-populates the results table with Pending items."""
        with self._lock:
            c = self.conn.cursor()
            created_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            c.execute(
                "INSERT INTO sessions (name, created_at, total_items, status) VALUES (?, ?, ?, ?)",
                (name, created_at, total_items, "IN_PROGRESS"),
            )
            session_id = c.lastrowid

            data_to_insert = [
                (session_id, company, "Pending", "", "{}") for company in item_list
            ]
            c.executemany(
                "INSERT INTO results (session_id, company_input, status, reg_no, result_json) VALUES (?, ?, ?, ?, ?)",
                data_to_insert,
            )

            self.conn.commit()
            return session_id

    def update_result(self, session_id, company_input, status, reg_no, result_data):
        """Updates a specific result in the database."""
        with self._lock:
            c = self.conn.cursor()
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            json_str = json.dumps(result_data)

            c.execute(
                """UPDATE results
                         SET status=?, reg_no=?, result_json=?, updated_at=?
                         WHERE session_id=? AND company_input=?""",
                (status, reg_no, json_str, now, session_id, company_input),
            )
            self.conn.commit()

    def get_all_sessions(self):
        """Returns list of all sessions."""
        with self._lock:
            c = self.conn.cursor()
            c.execute(
                "SELECT id, name, created_at, total_items, status FROM sessions ORDER BY id DESC"
            )
            return c.fetchall()

    def get_session_results(self, session_id):
        """Returns all results for a session, useful for resuming or exporting."""
        with self._lock:
            c = self.conn.cursor()
            c.execute(
                "SELECT company_input, status, reg_no, result_json FROM results WHERE session_id=?",
                (session_id,),
            )
            rows = c.fetchall()

        data = []
        for r in rows:
            data.append(
                {
                    "company_input": r[0],
                    "status": r[1],
                    "reg_no": r[2],
                    "data": json.loads(r[3]) if r[3] else {},
                }
            )
        return data

    def close(self):
        if self.conn:
            self.conn.close()
