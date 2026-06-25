import sqlite3
import os
from datetime import datetime, date
import pytz

TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")
DB_PATH = os.getenv("DB_PATH", "attendance.db")


class Database:
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()

    def _init_tables(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS employees (
                telegram_id INTEGER PRIMARY KEY,
                name        TEXT NOT NULL,
                registered_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS attendance (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER NOT NULL,
                date        TEXT NOT NULL,
                entry_time  TEXT,
                exit_time   TEXT,
                total_hours REAL,
                FOREIGN KEY (telegram_id) REFERENCES employees(telegram_id)
            );
        """)
        self.conn.commit()

    # ---------- employees ----------

    def register_employee(self, telegram_id: int, name: str):
        if self.get_employee(telegram_id):
            self.conn.execute(
                "UPDATE employees SET name = ? WHERE telegram_id = ?",
                (name, telegram_id),
            )
        else:
            self.conn.execute(
                "INSERT INTO employees (telegram_id, name, registered_at) VALUES (?, ?, ?)",
                (telegram_id, name, datetime.now(TIMEZONE).isoformat()),
            )
        self.conn.commit()

    def get_employee(self, telegram_id: int):
        row = self.conn.execute(
            "SELECT * FROM employees WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        return dict(row) if row else None

    def list_employees(self):
        rows = self.conn.execute(
            "SELECT * FROM employees ORDER BY name"
        ).fetchall()
        return [dict(r) for r in rows]

    # ---------- attendance ----------

    def register_entry(self, telegram_id: int, timestamp: datetime) -> str:
        """Returns 'ok' or 'already_in'."""
        today = timestamp.date().isoformat()
        existing = self.conn.execute(
            "SELECT id FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, today),
        ).fetchone()
        if existing:
            return "already_in"
        self.conn.execute(
            "INSERT INTO attendance (telegram_id, date, entry_time) VALUES (?, ?, ?)",
            (telegram_id, today, timestamp.isoformat()),
        )
        self.conn.commit()
        return "ok"

    def register_exit(self, telegram_id: int, timestamp: datetime):
        """Returns dict with total_hours, or None if no open entry today."""
        today = timestamp.date().isoformat()
        record = self.conn.execute(
            """SELECT * FROM attendance
               WHERE telegram_id = ? AND date = ?
                 AND entry_time IS NOT NULL AND exit_time IS NULL""",
            (telegram_id, today),
        ).fetchone()
        if not record:
            return None
        entry_time = datetime.fromisoformat(record["entry_time"])
        if entry_time.tzinfo is None:
            entry_time = TIMEZONE.localize(entry_time)
        total_hours = (timestamp - entry_time).total_seconds() / 3600
        self.conn.execute(
            "UPDATE attendance SET exit_time = ?, total_hours = ? WHERE id = ?",
            (timestamp.isoformat(), total_hours, record["id"]),
        )
        self.conn.commit()
        return {"total_hours": total_hours}

    def get_today_status(self, telegram_id: int, today: date):
        record = self.conn.execute(
            "SELECT * FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, today.isoformat()),
        ).fetchone()
        if not record:
            return None
        r = dict(record)
        if r["entry_time"]:
            dt = datetime.fromisoformat(r["entry_time"])
            r["entry_time"] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
        if r["exit_time"]:
            dt = datetime.fromisoformat(r["exit_time"])
            r["exit_time"] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
        return r

    def get_all_records(self):
        rows = self.conn.execute("""
            SELECT e.name, a.date, a.entry_time, a.exit_time, a.total_hours
            FROM attendance a
            JOIN employees e ON a.telegram_id = e.telegram_id
            ORDER BY a.date DESC, e.name ASC
        """).fetchall()
        result = []
        for row in rows:
            r = dict(row)
            for field in ("entry_time", "exit_time"):
                if r[field]:
                    dt = datetime.fromisoformat(r[field])
                    r[field] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
            result.append(r)
        return result
