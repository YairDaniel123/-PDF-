import sqlite3
import contextlib

class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_db()

    @contextlib.contextmanager
    def get_conn(self):
        # שימוש נכון ב-contextmanager כדי לסגור את החיבור בסיום
        conn = sqlite3.connect(self.db_path, timeout=30, check_same_thread=False)
        conn.execute('PRAGMA journal_mode=WAL')
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def _init_db(self):
        with self.get_conn() as conn:
            conn.executescript('''
                CREATE TABLE IF NOT EXISTS files (id INTEGER PRIMARY KEY AUTOINCREMENT, path TEXT UNIQUE, filename TEXT, mod_time REAL, scanned INTEGER DEFAULT 0);
                CREATE VIRTUAL TABLE IF NOT EXISTS pages USING fts5(file_id UNINDEXED, page_num UNINDEXED, content);
                CREATE TABLE IF NOT EXISTS history (path TEXT UNIQUE, last_access REAL);
                CREATE TABLE IF NOT EXISTS favorites (path TEXT UNIQUE);
            ''')