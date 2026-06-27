import sqlite3
import os
import shutil
import threading
from datetime import datetime, date
import pytz

TIMEZONE  = pytz.timezone("America/Argentina/Buenos_Aires")
DB_PATH   = os.getenv("DB_PATH", "attendance.db")

# Destinos de copia en tiempo real (carpetas locales sincronizadas a la nube)
_LIVE_DESTINATIONS = [
    os.path.join(os.path.expanduser("~"), "OneDrive", "FichaYA", "attendance_live.db"),
    r"G:\Mi unidad\FichaYA\attendance_live.db",
]


def _backup_live():
    """Copia la DB a OneDrive y Google Drive en un hilo secundario. Silencioso si falla."""
    def _copy():
        for dest in _LIVE_DESTINATIONS:
            try:
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copy2(DB_PATH, dest)
            except Exception:
                pass
    threading.Thread(target=_copy, daemon=True).start()

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
        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()
        self._seed_holidays()
        _backup_live()   # copia inicial al arrancar

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
        # Migraciones seguras
        emp_cols = [r[1] for r in self.conn.execute("PRAGMA table_info(employees)").fetchall()]
        if "categoria" not in emp_cols:
            self.conn.execute("ALTER TABLE employees ADD COLUMN categoria TEXT DEFAULT 'empleado'")

        att_cols = [r[1] for r in self.conn.execute("PRAGMA table_info(attendance)").fetchall()]
        if "obra_id" not in att_cols:
            self.conn.execute("ALTER TABLE attendance ADD COLUMN obra_id TEXT DEFAULT NULL")

        # Tabla de obras
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS obras (
                id         TEXT PRIMARY KEY,
                name       TEXT NOT NULL UNIQUE,
                active     INTEGER DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS obra_sessions (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER NOT NULL,
                obra_id     TEXT NOT NULL,
                date        TEXT NOT NULL,
                entry_time  TEXT,
                exit_time   TEXT,
                total_hours REAL,
                FOREIGN KEY (telegram_id) REFERENCES employees(telegram_id),
                FOREIGN KEY (obra_id) REFERENCES obras(id)
            )
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
        _backup_live()

    def get_employee(self, telegram_id: int):
        row = self.conn.execute(
            "SELECT * FROM employees WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        return dict(row) if row else None

    def list_employees(self):
        rows = self.conn.execute("SELECT * FROM employees ORDER BY name").fetchall()
        return [dict(r) for r in rows]

    def set_categoria(self, telegram_id: int, categoria: str):
        self.conn.execute(
            "UPDATE employees SET categoria = ? WHERE telegram_id = ?",
            (categoria, telegram_id),
        )
        self.conn.commit()

    def find_employee_by_name(self, fragment: str):
        fragment = fragment.lower()
        rows = self.conn.execute("SELECT * FROM employees").fetchall()
        return [dict(r) for r in rows if fragment in r["name"].lower()]

    # ---------- attendance ----------

    # ---------- obras ----------

    def create_obra(self, name: str) -> str:
        import re
        obra_id = re.sub(r'[^a-z0-9_]', '', name.lower().strip().replace(" ", "_"))
        if not obra_id:
            obra_id = f"obra_{int(datetime.now(TIMEZONE).timestamp())}"
        existing = self.conn.execute("SELECT id FROM obras WHERE id = ?", (obra_id,)).fetchone()
        if existing:
            self.conn.execute("UPDATE obras SET active = 1, name = ? WHERE id = ?", (name.strip(), obra_id))
        else:
            self.conn.execute(
                "INSERT INTO obras (id, name, active, created_at) VALUES (?, ?, 1, ?)",
                (obra_id, name.strip(), datetime.now(TIMEZONE).isoformat()),
            )
        self.conn.commit()
        return obra_id

    def list_obras(self, active_only: bool = True) -> list:
        q = "SELECT * FROM obras WHERE active = 1 ORDER BY name" if active_only \
            else "SELECT * FROM obras ORDER BY name"
        return [dict(r) for r in self.conn.execute(q).fetchall()]

    def get_obra(self, obra_id: str):
        row = self.conn.execute("SELECT * FROM obras WHERE id = ?", (obra_id,)).fetchone()
        return dict(row) if row else None

    def close_obra(self, name_fragment: str) -> list:
        matches = self.conn.execute(
            "SELECT * FROM obras WHERE LOWER(name) LIKE ? AND active = 1",
            (f"%{name_fragment.lower()}%",),
        ).fetchall()
        for m in matches:
            self.conn.execute("UPDATE obras SET active = 0 WHERE id = ?", (m["id"],))
        self.conn.commit()
        return [dict(m) for m in matches]

    def get_obra_hours(self, obra_id: str, start: date, end: date) -> list:
        rows = self.conn.execute("""
            SELECT e.name, o.entry_time, o.exit_time, o.total_hours, o.date
            FROM obra_sessions o
            JOIN employees e ON o.telegram_id = e.telegram_id
            WHERE o.obra_id = ? AND o.date BETWEEN ? AND ?
            ORDER BY o.date ASC, e.name ASC
        """, (obra_id, start.isoformat(), end.isoformat())).fetchall()
        return [self._parse_times(dict(r)) for r in rows]

    def register_obra_entry(self, telegram_id: int, obra_id: str, timestamp: datetime) -> str:
        """Registra llegada a obra. Solo bloquea si hay sesión de obra abierta."""
        today = timestamp.date().isoformat()
        open_s = self.conn.execute(
            """SELECT id FROM obra_sessions
               WHERE telegram_id = ? AND exit_time IS NULL""",
            (telegram_id,),
        ).fetchone()
        if open_s:
            return "already_in_obra"
        self.conn.execute(
            """INSERT INTO obra_sessions (telegram_id, obra_id, date, entry_time)
               VALUES (?, ?, ?, ?)""",
            (telegram_id, obra_id, today, timestamp.isoformat()),
        )
        self.conn.commit()
        _backup_live()
        return "ok"

    def register_obra_exit(self, telegram_id: int, timestamp: datetime):
        """Cierra la sesión de obra abierta."""
        record = self.conn.execute(
            """SELECT * FROM obra_sessions
               WHERE telegram_id = ? AND exit_time IS NULL
               ORDER BY entry_time DESC LIMIT 1""",
            (telegram_id,),
        ).fetchone()
        if not record:
            return None
        entry = datetime.fromisoformat(record["entry_time"])
        if entry.tzinfo is None:
            entry = TIMEZONE.localize(entry)
        total = (timestamp - entry).total_seconds() / 3600
        self.conn.execute(
            "UPDATE obra_sessions SET exit_time = ?, total_hours = ? WHERE id = ?",
            (timestamp.isoformat(), total, record["id"]),
        )
        self.conn.commit()
        _backup_live()
        return {"obra_id": record["obra_id"], "total_hours": total}

    def get_open_obra_session(self, telegram_id: int):
        row = self.conn.execute(
            """SELECT * FROM obra_sessions
               WHERE telegram_id = ? AND exit_time IS NULL
               ORDER BY entry_time DESC LIMIT 1""",
            (telegram_id,),
        ).fetchone()
        if not row:
            return None
        r = dict(row)
        if r["entry_time"]:
            dt = datetime.fromisoformat(r["entry_time"])
            r["entry_time"] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
        return r

    def get_all_obra_sessions(self, start: date, end: date) -> list:
        rows = self.conn.execute("""
            SELECT e.name emp_name, ob.name obra_name,
                   os.date, os.entry_time, os.exit_time, os.total_hours
            FROM obra_sessions os
            JOIN employees e  ON os.telegram_id = e.telegram_id
            JOIN obras ob ON os.obra_id = ob.id
            WHERE os.date BETWEEN ? AND ?
            ORDER BY ob.name, os.date, e.name
        """, (start.isoformat(), end.isoformat())).fetchall()
        return [self._parse_times(dict(r)) for r in rows]

    # ---------- attendance ----------

    def register_entry(self, telegram_id: int, timestamp: datetime, obra_id: str = None) -> str:
        """Permite múltiples entradas por día. Solo bloquea si hay entrada abierta."""
        today = timestamp.date().isoformat()
        open_entry = self.conn.execute(
            """SELECT id FROM attendance
               WHERE telegram_id = ? AND date = ? AND exit_time IS NULL""",
            (telegram_id, today),
        ).fetchone()
        if open_entry:
            return "already_in"
        count = self.conn.execute(
            "SELECT COUNT(*) FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, today),
        ).fetchone()[0]
        self.conn.execute(
            "INSERT INTO attendance (telegram_id, date, entry_time, obra_id) VALUES (?, ?, ?, ?)",
            (telegram_id, today, timestamp.isoformat(), obra_id),
        )
        self.conn.commit()
        _backup_live()
        return f"ok_{count + 1}"

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
        raw_hours   = (timestamp - entry_time).total_seconds() / 3600
        total_hours = round(raw_hours * 4) / 4  # redondeo a cuarto de hora
        self.conn.execute(
            "UPDATE attendance SET exit_time = ?, total_hours = ? WHERE id = ?",
            (timestamp.isoformat(), total_hours, record["id"]),
        )
        self.conn.commit()
        _backup_live()
        return {"total_hours": total_hours}

    def get_today_status(self, telegram_id: int, today: date):
        """Devuelve la entrada abierta actual (si existe) más el total acumulado del día."""
        open_rec = self.conn.execute(
            """SELECT * FROM attendance
               WHERE telegram_id = ? AND date = ? AND exit_time IS NULL
               ORDER BY entry_time DESC LIMIT 1""",
            (telegram_id, today.isoformat()),
        ).fetchone()

        # Total horas de sesiones cerradas
        daily_total = self.conn.execute(
            """SELECT COALESCE(SUM(total_hours), 0) FROM attendance
               WHERE telegram_id = ? AND date = ? AND exit_time IS NOT NULL""",
            (telegram_id, today.isoformat()),
        ).fetchone()[0]

        # Cantidad de sesiones
        session_count = self.conn.execute(
            "SELECT COUNT(*) FROM attendance WHERE telegram_id = ? AND date = ?",
            (telegram_id, today.isoformat()),
        ).fetchone()[0]

        if open_rec:
            r = dict(open_rec)
            for field in ("entry_time", "exit_time"):
                if r[field]:
                    dt = datetime.fromisoformat(r[field])
                    r[field] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
            r["daily_total"] = daily_total
            r["session_count"] = session_count
            return r

        # Sin entrada abierta → devolver último registro cerrado
        last = self.conn.execute(
            """SELECT * FROM attendance
               WHERE telegram_id = ? AND date = ?
               ORDER BY exit_time DESC LIMIT 1""",
            (telegram_id, today.isoformat()),
        ).fetchone()
        if not last:
            return None
        r = dict(last)
        for field in ("entry_time", "exit_time"):
            if r[field]:
                dt = datetime.fromisoformat(r[field])
                r[field] = TIMEZONE.localize(dt) if dt.tzinfo is None else dt.astimezone(TIMEZONE)
        r["daily_total"]   = daily_total
        r["session_count"] = session_count
        return r

    def get_daily_sessions(self, telegram_id: int, today: date) -> list:
        """Devuelve todas las sesiones del día para calcular horas totales."""
        rows = self.conn.execute(
            """SELECT * FROM attendance
               WHERE telegram_id = ? AND date = ?
               ORDER BY entry_time ASC""",
            (telegram_id, today.isoformat()),
        ).fetchall()
        return [self._parse_times(dict(r)) for r in rows]

    def _aggregate_sessions(self, raw_rows: list) -> list:
        """Agrupa múltiples sesiones por (empleado, día) en un solo registro."""
        from collections import OrderedDict
        grouped = OrderedDict()
        for r in raw_rows:
            key = (r["telegram_id"], r["name"], r["date"])
            if key not in grouped:
                grouped[key] = {
                    "telegram_id": r["telegram_id"],
                    "name":        r["name"],
                    "date":        r["date"],
                    "sessions":    [],
                }
            grouped[key]["sessions"].append(r)

        result = []
        for g in grouped.values():
            sessions    = g["sessions"]
            closed      = [s for s in sessions if s.get("exit_time")]
            has_open    = any(s for s in sessions if not s.get("exit_time"))
            total_hours = sum(s["total_hours"] or 0 for s in closed)

            entry_times = [s["entry_time"] for s in sessions if s.get("entry_time")]
            exit_times  = [s["exit_time"]  for s in closed]

            result.append({
                "telegram_id":   g["telegram_id"],
                "name":          g["name"],
                "date":          g["date"],
                "entry_time":    entry_times[0] if entry_times else None,
                "exit_time":     exit_times[-1] if exit_times and not has_open else None,
                "total_hours":   total_hours if total_hours > 0 else None,
                "sessions":      sessions,
                "session_count": len(sessions),
                "has_open":      has_open,
            })
        return result

    def get_records_by_period(self, start: date, end: date):
        rows = self.conn.execute("""
            SELECT e.telegram_id, e.name, a.date, a.entry_time, a.exit_time, a.total_hours
            FROM attendance a
            JOIN employees e ON a.telegram_id = e.telegram_id
            WHERE a.date BETWEEN ? AND ?
            ORDER BY a.date ASC, e.name ASC, a.entry_time ASC
        """, (start.isoformat(), end.isoformat())).fetchall()
        parsed = [self._parse_times(dict(r)) for r in rows]
        return self._aggregate_sessions(parsed)

    def get_all_records(self):
        rows = self.conn.execute("""
            SELECT e.telegram_id, e.name, a.date, a.entry_time, a.exit_time, a.total_hours
            FROM attendance a
            JOIN employees e ON a.telegram_id = e.telegram_id
            ORDER BY a.date DESC, e.name ASC, a.entry_time ASC
        """).fetchall()
        parsed = [self._parse_times(dict(r)) for r in rows]
        return self._aggregate_sessions(parsed)

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
