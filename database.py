import sqlite3
import os
from datetime import datetime, date
import pytz

TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")
DB_PATH = os.getenv("DB_PATH", "attendance.db")

FERIADOS_2026 = [
    ("2026-01-01", "Año Nuevo"),
    ("2026-02-16", "Carnaval"),
    ("2026-02-17", "Carnaval"),
    ("2026-03-24", "Día de la Memoria"),
    ("2026-04-02", "Día del Veterano de Malvinas"),
    ("2026-04-03", "Viernes Santo"),
    ("2026-05-01", "Día del Trabajador"),
    ("2026-05-25", "Revolución de Mayo"),
    ("2026-06-20", "Paso a la Inmortalidad del Gral. Güemes"),
    ("2026-07-09", "Día de la Independencia"),
    ("2026-08-17", "Paso a la Inmortalidad del Gral. San Martín"),
    ("2026-10-12", "Día de la Diversidad Cultural"),
    ("2026-11-20", "Día de la Soberanía Nacional"),
    ("2026-12-08", "Inmaculada Concepción"),
    ("2026-12-25", "Navidad"),
]

ABSENCE_LABELS = {
    "vacacion":      "Vacación",
    "vacación":      "Vacación",
    "enfermedad":    "Enfermedad",
    "licencia":      "Licencia",
    "justificada":   "Justificada",
    "injustificada": "Injustificada",
}


class Database:
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()
        self._seed_holidays()

    def _init_tables(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS employees (
                telegram_id   INTEGER PRIMARY KEY,
                name          TEXT NOT NULL,
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

            CREATE TABLE IF NOT EXISTS holidays (
                date TEXT PRIMARY KEY,
                name TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS absences (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER NOT NULL,
                date        TEXT NOT NULL,
                type        TEXT NOT NULL,
                notes       TEXT DEFAULT '',
                UNIQUE(telegram_id, date),
                FOREIGN KEY (telegram_id) REFERENCES employees(telegram_id)
            );

            CREATE TABLE IF NOT EXISTS shifts (
                telegram_id  INTEGER PRIMARY KEY,
                entry_hour   INTEGER DEFAULT 9,
                entry_minute INTEGER DEFAULT 0,
                exit_hour    INTEGER DEFAULT 18,
                exit_minute  INTEGER DEFAULT 0,
                FOREIGN KEY (telegram_id) REFERENCES employees(telegram_id)
            );
        """)
        self.conn.commit()

    def _seed_holidays(self):
        count = self.conn.execute("SELECT COUNT(*) FROM holidays").fetchone()[0]
        if count == 0:
            self.conn.executemany(
                "INSERT OR IGNORE INTO holidays (date, name) VALUES (?, ?)",
                FERIADOS_2026,
            )
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
        rows = self.conn.execute("SELECT * FROM employees ORDER BY name").fetchall()
        return [dict(r) for r in rows]

    def find_employee_by_name(self, fragment: str):
        fragment = fragment.lower()
        rows = self.conn.execute("SELECT * FROM employees").fetchall()
        return [dict(r) for r in rows if fragment in r["name"].lower()]

    # ---------- attendance ----------

    def register_entry(self, telegram_id: int, timestamp: datetime) -> str:
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
        for field in ("entry_time", "exit_time"):
            if r[field]:
                dt = datetime.fromisoformat(r[field])
                r[field] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
        return r

    def get_records_by_period(self, start: date, end: date):
        rows = self.conn.execute("""
            SELECT e.telegram_id, e.name, a.date, a.entry_time, a.exit_time, a.total_hours
            FROM attendance a
            JOIN employees e ON a.telegram_id = e.telegram_id
            WHERE a.date BETWEEN ? AND ?
            ORDER BY a.date ASC, e.name ASC
        """, (start.isoformat(), end.isoformat())).fetchall()
        return [self._parse_times(dict(r)) for r in rows]

    def get_all_records(self):
        rows = self.conn.execute("""
            SELECT e.telegram_id, e.name, a.date, a.entry_time, a.exit_time, a.total_hours
            FROM attendance a
            JOIN employees e ON a.telegram_id = e.telegram_id
            ORDER BY a.date DESC, e.name ASC
        """).fetchall()
        return [self._parse_times(dict(r)) for r in rows]

    def get_employees_without_entry(self, today: date) -> list:
        rows = self.conn.execute("""
            SELECT e.telegram_id, e.name
            FROM employees e
            LEFT JOIN attendance a
                   ON a.telegram_id = e.telegram_id AND a.date = ?
            WHERE a.id IS NULL
        """, (today.isoformat(),)).fetchall()
        return [dict(r) for r in rows]

    def get_employees_with_open_entry(self, today: date) -> list:
        rows = self.conn.execute("""
            SELECT e.telegram_id, e.name, a.entry_time
            FROM employees e
            JOIN attendance a ON a.telegram_id = e.telegram_id AND a.date = ?
            WHERE a.exit_time IS NULL AND a.entry_time IS NOT NULL
        """, (today.isoformat(),)).fetchall()
        result = []
        for row in rows:
            r = dict(row)
            dt = datetime.fromisoformat(r["entry_time"])
            r["entry_time"] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
            result.append(r)
        return result

    def _parse_times(self, r: dict) -> dict:
        for field in ("entry_time", "exit_time"):
            if r.get(field):
                dt = datetime.fromisoformat(r[field])
                r[field] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
        return r

    # ---------- shifts ----------

    def set_shift(self, telegram_id: int, entry_hour: int, entry_minute: int,
                  exit_hour: int, exit_minute: int):
        self.conn.execute("""
            INSERT OR REPLACE INTO shifts
                (telegram_id, entry_hour, entry_minute, exit_hour, exit_minute)
            VALUES (?, ?, ?, ?, ?)
        """, (telegram_id, entry_hour, entry_minute, exit_hour, exit_minute))
        self.conn.commit()

    def get_shift(self, telegram_id: int) -> tuple:
        """Returns (entry_hour, entry_minute, exit_hour, exit_minute). Default 09:00-18:00."""
        row = self.conn.execute(
            "SELECT * FROM shifts WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        if row:
            return (row["entry_hour"], row["entry_minute"],
                    row["exit_hour"], row["exit_minute"])
        return (9, 0, 18, 0)

    # ---------- attendance corrections ----------

    def update_attendance_field(self, telegram_id: int, d: date,
                                field: str, new_dt: datetime) -> bool:
        """Corrige entry_time o exit_time de un día y recalcula total_hours."""
        record = self.conn.execute(
            "SELECT * FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, d.isoformat()),
        ).fetchone()
        if not record:
            return False
        self.conn.execute(
            f"UPDATE attendance SET {field} = ? WHERE telegram_id = ? AND date = ?",
            (new_dt.isoformat(), telegram_id, d.isoformat()),
        )
        updated = dict(self.conn.execute(
            "SELECT * FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, d.isoformat()),
        ).fetchone())
        if updated["entry_time"] and updated["exit_time"]:
            ent = datetime.fromisoformat(updated["entry_time"])
            ext = datetime.fromisoformat(updated["exit_time"])
            if ent.tzinfo is None: ent = TIMEZONE.localize(ent)
            if ext.tzinfo is None: ext = TIMEZONE.localize(ext)
            total = (ext - ent).total_seconds() / 3600
            self.conn.execute(
                "UPDATE attendance SET total_hours = ? WHERE telegram_id = ? AND date = ?",
                (total, telegram_id, d.isoformat()),
            )
        self.conn.commit()
        return True

    def delete_attendance(self, telegram_id: int, d: date) -> bool:
        cur = self.conn.execute(
            "DELETE FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, d.isoformat()),
        )
        self.conn.commit()
        return cur.rowcount > 0

    def get_employee_records(self, telegram_id: int, limit: int = 10):
        rows = self.conn.execute("""
            SELECT date, entry_time, exit_time, total_hours FROM attendance
            WHERE telegram_id = ?
            ORDER BY date DESC LIMIT ?
        """, (telegram_id, limit)).fetchall()
        return [self._parse_times(dict(r)) for r in rows]

    # ---------- holidays ----------

    def add_holiday(self, d: date, name: str):
        self.conn.execute(
            "INSERT OR REPLACE INTO holidays (date, name) VALUES (?, ?)",
            (d.isoformat(), name),
        )
        self.conn.commit()

    def remove_holiday(self, d: date) -> bool:
        cur = self.conn.execute(
            "DELETE FROM holidays WHERE date = ?", (d.isoformat(),)
        )
        self.conn.commit()
        return cur.rowcount > 0

    def get_holidays(self, start: date, end: date) -> dict:
        """Returns {date_str: name}."""
        rows = self.conn.execute(
            "SELECT date, name FROM holidays WHERE date BETWEEN ? AND ?",
            (start.isoformat(), end.isoformat()),
        ).fetchall()
        return {r["date"]: r["name"] for r in rows}

    def get_holidays_month(self, year: int, month: int) -> dict:
        from calendar import monthrange
        last = monthrange(year, month)[1]
        start = date(year, month, 1)
        end = date(year, month, last)
        return self.get_holidays(start, end)

    def is_holiday(self, d: date) -> str | None:
        row = self.conn.execute(
            "SELECT name FROM holidays WHERE date = ?", (d.isoformat(),)
        ).fetchone()
        return row["name"] if row else None

    # ---------- absences ----------

    def add_absence(self, telegram_id: int, d: date, absence_type: str, notes: str = ""):
        self.conn.execute(
            """INSERT OR REPLACE INTO absences (telegram_id, date, type, notes)
               VALUES (?, ?, ?, ?)""",
            (telegram_id, d.isoformat(), absence_type, notes),
        )
        self.conn.commit()

    def remove_absence(self, telegram_id: int, d: date) -> bool:
        cur = self.conn.execute(
            "DELETE FROM absences WHERE telegram_id = ? AND date = ?",
            (telegram_id, d.isoformat()),
        )
        self.conn.commit()
        return cur.rowcount > 0

    def get_absences(self, start: date, end: date) -> dict:
        """Returns {telegram_id: {date_str: type}}."""
        rows = self.conn.execute("""
            SELECT telegram_id, date, type FROM absences
            WHERE date BETWEEN ? AND ?
        """, (start.isoformat(), end.isoformat())).fetchall()
        result: dict = {}
        for r in rows:
            tid = r["telegram_id"]
            if tid not in result:
                result[tid] = {}
            result[tid][r["date"]] = r["type"]
        return result
