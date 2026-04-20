import datetime as dt
import difflib
import json
import os
import re
import shutil
import sqlite3
import urllib.parse
import zipfile
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree as ET


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DB_PATH = BASE_DIR / "planner.db"
BACKUP_DIR = BASE_DIR / "backups"
PORT = 8765
BASE_YEAR = 2026
BOOTSTRAP_XLSX = Path(r"C:\Users\UTENTE\Desktop\PROJ\09032026\ERP_RESTART_V13_MACRO.xlsm")
DEFAULT_DEMAND_XLSX = Path(r"\\SPECSTATION\0301_Commesse\12 Varie\FABBISOGNO RISORSE 2026\Fabbisogno Risorse al 13-03-2026.xlsx")
DEFAULT_RESOURCES_XLSX = Path(r"C:\Users\UTENTE\Desktop\risorse.xlsx")
DEFAULT_PLAN_XLSX = Path(r"\\SPECSTATION\0301_Commesse\12 Varie\PLAN 2026_rev.02.xlsx")

NS_MAIN = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

STANDARD_ROLES = [
    "CAPO CANTIERE",
    "QUALITY CONTROL / WELDING INSPECTOR",
    "ASPP",
    "CAPO SQUADRA",
    "TUBISTA",
    "CARPENTIERE",
    "SOLLEVAMENTI",
    "AUTISTA",
    "SALDATORE TIG-ELETTRODO",
    "SALDATORE FILO",
    "PWHT",
    "GENERICO",
    "MAGAZZINIERE",
    "MECCANICO",
    "MECCANICO SERVICE",
    "MONTATORE",
    "MANDRINATORE",
    "ELETTRICISTA",
    "PONTEGGIATORE",
    "COIBENTATORE",
    "VERNICIATORE",
]
EXT_SUFFIX = "-EXT"
EXT_EMPLOYER = "EXT"
EXTRA_ROLES_CACHE = set()
CURRENT_DEMAND_META_KEY = "current_demand_path"


def refresh_custom_roles(conn):
    global EXTRA_ROLES_CACHE
    rows = conn.execute("SELECT name FROM roles").fetchall()
    EXTRA_ROLES_CACHE = {str(r["name"]).strip().upper() for r in rows if str(r["name"]).strip()}


def all_standard_roles():
    return sorted(set(STANDARD_ROLES) | set(EXTRA_ROLES_CACHE))


def db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def get_meta_value(conn, key, default=""):
    try:
        row = conn.execute("SELECT value FROM meta WHERE key=?", (normalize_text(key),)).fetchone()
    except sqlite3.Error:
        return default
    if not row:
        return default
    return normalize_text(row["value"]) or default


def set_meta_value(conn, key, value):
    conn.execute(
        """
        INSERT INTO meta(key, value)
        VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
        """,
        (normalize_text(key), normalize_text(value)),
    )


def get_current_demand_source_path(conn):
    meta_path = get_meta_value(conn, CURRENT_DEMAND_META_KEY, "")
    if meta_path:
        return meta_path
    # Fallback robusto: recupera ultimo path importato dal log attività.
    try:
        rows = conn.execute(
            """
            SELECT action, detail
            FROM activity_log
            WHERE action IN ('Import fabbisogno', 'Import iniziale')
            ORDER BY id DESC
            LIMIT 50
            """
        ).fetchall()
    except sqlite3.Error:
        return str(DEFAULT_DEMAND_XLSX)
    for r in rows:
        detail = normalize_text(r["detail"])
        if not detail:
            continue
        # Import iniziale: "<demand_path> | <resources_path>"
        candidate = normalize_text(detail.split("|")[0])
        if candidate:
            return candidate
    return str(DEFAULT_DEMAND_XLSX)


def ensure_backup_dir():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    return BACKUP_DIR


def backup_timestamp():
    return dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def latest_backup_info():
    ensure_backup_dir()
    latest_link = BACKUP_DIR / "planner_last_backup.db"
    if latest_link.exists():
        stat = latest_link.stat()
        return {
            "path": str(latest_link),
            "name": latest_link.name,
            "created_at": dt.datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
            "size": stat.st_size,
        }
    backups = sorted(BACKUP_DIR.glob("planner_backup_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not backups:
        return None
    backup = backups[0]
    stat = backup.stat()
    return {
        "path": str(backup),
        "name": backup.name,
        "created_at": dt.datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
        "size": stat.st_size,
    }


def create_db_backup(reason="manual"):
    ensure_backup_dir()
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database non trovato: {DB_PATH}")
    stamp = backup_timestamp()
    safe_reason = re.sub(r"[^A-Za-z0-9_-]+", "-", normalize_text(reason) or "manual").strip("-") or "manual"
    backup_path = BACKUP_DIR / f"planner_backup_{stamp}_{safe_reason}.db"
    shutil.copy2(DB_PATH, backup_path)
    latest_path = BACKUP_DIR / "planner_last_backup.db"
    shutil.copy2(backup_path, latest_path)
    stat = backup_path.stat()
    return {
        "path": str(backup_path),
        "name": backup_path.name,
        "created_at": dt.datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
        "size": stat.st_size,
        "reason": safe_reason,
    }


def restore_latest_backup():
    info = latest_backup_info()
    if not info:
        raise FileNotFoundError("Nessun salvataggio disponibile.")
    source = Path(info["path"])
    if not source.exists():
        raise FileNotFoundError(f"Salvataggio non trovato: {source}")
    shutil.copy2(source, DB_PATH)
    return info


def init_db():
    with db_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS resources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                employee_code TEXT NOT NULL DEFAULT '',
                role1 TEXT NOT NULL DEFAULT '',
                role2 TEXT NOT NULL DEFAULT '',
                hire_date TEXT DEFAULT '',
                end_date TEXT DEFAULT '',
                birth_date TEXT DEFAULT '',
                phone TEXT DEFAULT '',
                email TEXT DEFAULT '',
                residence_city TEXT DEFAULT '',
                employer TEXT NOT NULL DEFAULT 'SPEC',
                hire_level TEXT DEFAULT '',
                base_location TEXT DEFAULT '',
                level TEXT DEFAULT '',
                contract_type TEXT DEFAULT '',
                pb_overtime_hourly REAL NOT NULL DEFAULT 0,
                pb_day_office REAL NOT NULL DEFAULT 0,
                pb_day_site REAL NOT NULL DEFAULT 0,
                pb_fixed_plus REAL NOT NULL DEFAULT 0,
                glob_hour_office REAL NOT NULL DEFAULT 0,
                glob_hour_site REAL NOT NULL DEFAULT 0,
                glob_fixed_plus REAL NOT NULL DEFAULT 0,
                doc_type TEXT DEFAULT '',
                doc_number TEXT DEFAULT '',
                doc_expiry TEXT DEFAULT '',
                certifications TEXT DEFAULT '',
                note TEXT DEFAULT '',
                no_travel INTEGER NOT NULL DEFAULT 0,
                cost REAL NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL UNIQUE,
                activity TEXT NOT NULL DEFAULT '',
                type TEXT NOT NULL DEFAULT 'SITE',
                closed INTEGER NOT NULL DEFAULT 0,
                workshop_rollup INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS demands (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                week INTEGER NOT NULL,
                role TEXT NOT NULL,
                qty INTEGER NOT NULL DEFAULT 0,
                UNIQUE(project_id, week, role)
            );

            CREATE TABLE IF NOT EXISTS allocations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                resource_id INTEGER NOT NULL REFERENCES resources(id) ON DELETE CASCADE,
                role TEXT NOT NULL,
                week_from INTEGER NOT NULL,
                week_to INTEGER NOT NULL,
                weight REAL NOT NULL DEFAULT 1,
                note TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS unavailability (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                resource_id INTEGER NOT NULL REFERENCES resources(id) ON DELETE CASCADE,
                week_from INTEGER NOT NULL,
                week_to INTEGER NOT NULL,
                reason TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS activity_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT NOT NULL,
                actor TEXT NOT NULL,
                action TEXT NOT NULL,
                detail TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS project_templates (
                project_id INTEGER PRIMARY KEY REFERENCES projects(id) ON DELETE CASCADE,
                week_from INTEGER NOT NULL,
                week_to INTEGER NOT NULL,
                use_standard_roles INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS roles (
                name TEXT PRIMARY KEY
            );

            CREATE TABLE IF NOT EXISTS meta (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            );
            """
        )

        # Backward-compatible migrations
        proj_cols = [r["name"] for r in conn.execute("PRAGMA table_info(projects)")]
        if "workshop_rollup" not in proj_cols:
            conn.execute("ALTER TABLE projects ADD COLUMN workshop_rollup INTEGER NOT NULL DEFAULT 0")
        resource_cols = [r["name"] for r in conn.execute("PRAGMA table_info(resources)")]
        allocation_cols = [r["name"] for r in conn.execute("PRAGMA table_info(allocations)")]
        if "weight" not in allocation_cols:
            conn.execute("ALTER TABLE allocations ADD COLUMN weight REAL NOT NULL DEFAULT 1")
        resource_migrations = {
            "employee_code": "TEXT NOT NULL DEFAULT ''",
            "birth_date": "TEXT DEFAULT ''",
            "phone": "TEXT DEFAULT ''",
            "email": "TEXT DEFAULT ''",
            "residence_city": "TEXT DEFAULT ''",
            "employer": "TEXT NOT NULL DEFAULT 'SPEC'",
            "hire_level": "TEXT DEFAULT ''",
            "base_location": "TEXT DEFAULT ''",
            "level": "TEXT DEFAULT ''",
            "contract_type": "TEXT DEFAULT ''",
            "pb_overtime_hourly": "REAL NOT NULL DEFAULT 0",
            "pb_day_office": "REAL NOT NULL DEFAULT 0",
            "pb_day_site": "REAL NOT NULL DEFAULT 0",
            "pb_fixed_plus": "REAL NOT NULL DEFAULT 0",
            "glob_hour_office": "REAL NOT NULL DEFAULT 0",
            "glob_hour_site": "REAL NOT NULL DEFAULT 0",
            "glob_fixed_plus": "REAL NOT NULL DEFAULT 0",
            "doc_type": "TEXT DEFAULT ''",
            "doc_number": "TEXT DEFAULT ''",
            "doc_expiry": "TEXT DEFAULT ''",
            "certifications": "TEXT DEFAULT ''",
            "note": "TEXT DEFAULT ''",
            "no_travel": "INTEGER NOT NULL DEFAULT 0",
            "cost": "REAL NOT NULL DEFAULT 0",
        }
        for col_name, col_def in resource_migrations.items():
            if col_name not in resource_cols:
                conn.execute(f"ALTER TABLE resources ADD COLUMN {col_name} {col_def}")
        conn.execute("UPDATE resources SET role1='CAPO CANTIERE' WHERE UPPER(role1)='CAPO SITE'")
        conn.execute("UPDATE resources SET role2='CAPO CANTIERE' WHERE UPPER(role2)='CAPO SITE'")
        conn.execute("UPDATE demands SET role='CAPO CANTIERE' WHERE UPPER(role)='CAPO SITE'")
        conn.execute("UPDATE allocations SET role='CAPO CANTIERE' WHERE UPPER(role)='CAPO SITE'")
        conn.execute("UPDATE projects SET closed=1 WHERE UPPER(code)='WS OVERALL'")
        set_meta_value(conn, CURRENT_DEMAND_META_KEY, get_current_demand_source_path(conn))
        refresh_custom_roles(conn)
        ensure_external_resources_pool(conn)
        conn.commit()


def log_event(conn, actor, action, detail=""):
    conn.execute(
        "INSERT INTO activity_log(ts, actor, action, detail) VALUES (?, ?, ?, ?)",
        (dt.datetime.now().isoformat(timespec="seconds"), actor, action, detail or ""),
    )


def trim_allocation_range(conn, allocation_id, cut_from, cut_to):
    row = conn.execute("SELECT * FROM allocations WHERE id=?", (int(allocation_id),)).fetchone()
    if not row:
        return {"deleted": 0, "updated": 0, "inserted": 0}
    a_from = int(row["week_from"])
    a_to = int(row["week_to"])
    cut_from = int(cut_from)
    cut_to = int(cut_to)
    if cut_to < a_from or cut_from > a_to:
        return {"deleted": 0, "updated": 0, "inserted": 0}
    if cut_from <= a_from and cut_to >= a_to:
        conn.execute("DELETE FROM allocations WHERE id=?", (int(allocation_id),))
        return {"deleted": 1, "updated": 0, "inserted": 0}
    if cut_from > a_from and cut_to < a_to:
        left_to = cut_from - 1
        right_from = cut_to + 1
        conn.execute("UPDATE allocations SET week_to=? WHERE id=?", (int(left_to), int(allocation_id)))
        conn.execute(
            """
            INSERT INTO allocations(project_id, resource_id, role, week_from, week_to, weight, note)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(row["project_id"]),
                int(row["resource_id"]),
                row["role"],
                int(right_from),
                int(a_to),
                float(row["weight"] if row["weight"] is not None else 1),
                row["note"] or "",
            ),
        )
        return {"deleted": 0, "updated": 1, "inserted": 1}
    if cut_from <= a_from <= cut_to < a_to:
        conn.execute("UPDATE allocations SET week_from=? WHERE id=?", (int(cut_to + 1), int(allocation_id)))
        return {"deleted": 0, "updated": 1, "inserted": 0}
    if a_from < cut_from <= a_to <= cut_to:
        conn.execute("UPDATE allocations SET week_to=? WHERE id=?", (int(cut_from - 1), int(allocation_id)))
        return {"deleted": 0, "updated": 1, "inserted": 0}
    return {"deleted": 0, "updated": 0, "inserted": 0}


def find_allocation_overlaps(conn, resource_id, week_from, week_to, current_allocation_id=None):
    rid = int(resource_id)
    wf = int(week_from)
    wt = int(week_to)
    if wt < wf:
        wf, wt = wt, wf

    if current_allocation_id is None:
        return conn.execute(
            """
            SELECT a.*, p.code AS project
            FROM allocations a
            JOIN projects p ON p.id = a.project_id
            WHERE a.resource_id = ?
              AND a.week_from <= ?
              AND a.week_to >= ?
            ORDER BY a.week_from, a.week_to, a.id
            """,
            (rid, wt, wf),
        ).fetchall()

    return conn.execute(
        """
        SELECT a.*, p.code AS project
        FROM allocations a
        JOIN projects p ON p.id = a.project_id
        WHERE a.resource_id = ?
          AND a.id <> ?
          AND a.week_from <= ?
          AND a.week_to >= ?
        ORDER BY a.week_from, a.week_to, a.id
        """,
        (rid, int(current_allocation_id), wt, wf),
    ).fetchall()


def parse_week(value):
    if value is None:
        return None
    s = str(value).strip().upper()
    if not s:
        return None
    if not re.fullmatch(r"W?\d{1,2}", s):
        return None
    if s.startswith("W"):
        s = s[1:]
    try:
        n = int(float(s))
    except ValueError:
        return None
    if 1 <= n <= 53:
        return n
    return None


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def parse_decimal(value, default=0.0):
    if value is None:
        return float(default)
    raw = str(value).strip()
    if not raw:
        return float(default)
    raw = raw.replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        raise ValueError(f"Valore numerico non valido: {value}")


def normalize_label(value):
    return " ".join(normalize_text(value).replace("\n", " ").replace("\r", " ").split())


def canonicalize_project_code(value):
    code = normalize_label(value)
    if not code:
        return ""
    code = re.sub(r"\(\s*CANTIERE\s*\)", "(SITE)", code, flags=re.IGNORECASE)
    code = re.sub(r"\(\s*SITE\s*\)", "(SITE)", code, flags=re.IGNORECASE)
    code = re.sub(r"\(\s*SITE SERVICE\s*\)", "(SITE SERVICE)", code, flags=re.IGNORECASE)
    return code


def normalize_person_name_key(value):
    text = normalize_label(value).upper()
    if not text:
        return ""
    return re.sub(r"[^A-Z0-9 ]+", " ", text).strip()


def find_matching_resource(resource_index, name):
    direct = resource_index.get(normalize_person_name_key(name))
    if direct:
        return direct
    tokens = normalize_person_name_key(name).split()
    if len(tokens) < 2:
        return None
    best = None
    best_score = 0.0
    for key, resource in resource_index.items():
        other = key.split()
        if len(other) < 2:
            continue
        if tokens[0] != other[0]:
            continue
        score = difflib.SequenceMatcher(None, " ".join(tokens[1:]), " ".join(other[1:])).ratio()
        if score > best_score:
            best = resource
            best_score = score
    if best_score >= 0.72:
        return best
    return None


def canonicalize_role(value):
    role = normalize_label(value)
    if not role:
        return ""
    role = " ".join(role.split())
    return role.upper()


def role_key(value):
    role = canonicalize_role(value)
    if not role:
        return ""
    return re.sub(r"[^A-Z0-9]+", " ", role).strip()


def is_external_role_name(name_text):
    name = normalize_label(name_text).upper()
    return bool(name) and name.endswith(EXT_SUFFIX)


def is_external_resource_row(res):
    if not res:
        return False
    employer = normalize_label(res.get("employer", "") if isinstance(res, dict) else res["employer"]).upper()
    name = normalize_label(res.get("name", "") if isinstance(res, dict) else res["name"]).upper()
    return employer == EXT_EMPLOYER or is_external_role_name(name)


def ensure_external_resources_pool(conn):
    for role in all_standard_roles():
        role_up = canonicalize_role(role)
        if not role_up or role_up.endswith(EXT_SUFFIX):
            continue
        ext_name = f"{role_up}{EXT_SUFFIX}"
        conn.execute(
            """
            INSERT OR IGNORE INTO resources(
                name, employee_code, role1, role2, hire_date, end_date,
                employer, contract_type, note
            )
            VALUES (?, ?, ?, '', 'W1', 'INDET.', ?, ?, ?)
            """,
            (
                ext_name,
                ext_name,
                role_up,
                EXT_EMPLOYER,
                EXT_EMPLOYER,
                "Risorsa esterna virtuale per copertura subappalto",
            ),
        )


def is_standard_role(role):
    if not role:
        return False
    upper = str(role).strip().upper()
    return upper in STANDARD_ROLES or upper in EXTRA_ROLES_CACHE


def parse_int(value, default=0):
    try:
        if value in (None, ""):
            return default
        return int(float(str(value).replace(",", ".")))
    except Exception:
        return default


def date_to_week(date_text):
    s = normalize_text(date_text)
    if not s:
        return 999
    s_upper = s.upper()
    if s_upper in {"N.D.", "ND", "N/D"}:
        return 0
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            d = dt.datetime.strptime(s, fmt).date()
            if d.year > BASE_YEAR:
                return 999
            if d.year < BASE_YEAR:
                return 0
            return int(d.isocalendar().week)
        except ValueError:
            pass
    return 999


def parse_date_flexible(date_text):
    s = normalize_text(date_text)
    if not s:
        return None
    s_upper = s.upper()
    if s_upper in {"INDET.", "INDET", "31/12/2099"}:
        return None
    if s_upper in {"N.D.", "ND", "N/D"}:
        return dt.date(1900, 1, 1)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def week_start_date(week):
    return dt.date.fromisocalendar(BASE_YEAR, int(week), 1)


def week_end_date(week):
    return week_start_date(week) + dt.timedelta(days=6)


def xlsx_col_index(cell_ref):
    letters = ""
    for ch in cell_ref:
        if ch.isalpha():
            letters += ch
        else:
            break
    idx = 0
    for c in letters:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx


def parse_shared_strings(zf):
    shared = []
    if "xl/sharedStrings.xml" not in zf.namelist():
        return shared
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    for si in root.findall(".//x:si", NS_MAIN):
        texts = []
        for t in si.findall(".//x:t", NS_MAIN):
            texts.append(t.text or "")
        shared.append("".join(texts))
    return shared


def get_sheet_path(zf, sheet_name):
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {}
    for rel in rel_root.findall(".//r:Relationship", NS_REL):
        rel_map[rel.attrib.get("Id")] = rel.attrib.get("Target")
    for sheet in wb_root.findall(".//x:sheets/x:sheet", NS_MAIN):
        name = sheet.attrib.get("name", "")
        if name == sheet_name:
            rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = rel_map.get(rel_id, "")
            if target and not target.startswith("/"):
                target = "xl/" + target
            return target
    return None


def read_sheet_rows(xlsx_path, sheet_name):
    rows = []
    with zipfile.ZipFile(xlsx_path) as zf:
        shared = parse_shared_strings(zf)
        sheet_path = get_sheet_path(zf, sheet_name)
        if not sheet_path or sheet_path not in zf.namelist():
            return rows
        root = ET.fromstring(zf.read(sheet_path))
        for row in root.findall(".//x:sheetData/x:row", NS_MAIN):
            row_map = {}
            for c in row.findall("x:c", NS_MAIN):
                rref = c.attrib.get("r", "")
                idx = xlsx_col_index(rref)
                typ = c.attrib.get("t")
                value = ""
                if typ == "inlineStr":
                    t = c.find("x:is/x:t", NS_MAIN)
                    value = t.text if t is not None and t.text is not None else ""
                else:
                    v = c.find("x:v", NS_MAIN)
                    if v is not None and v.text is not None:
                        value = v.text
                if typ == "s":
                    try:
                        value = shared[int(value)]
                    except Exception:
                        value = ""
                row_map[idx] = value
            rows.append(row_map)
    return rows


def read_sheet_rows_with_styles(xlsx_path, sheet_name):
    rows = []
    with zipfile.ZipFile(xlsx_path) as zf:
        shared = parse_shared_strings(zf)
        sheet_path = get_sheet_path(zf, sheet_name)
        if not sheet_path or sheet_path not in zf.namelist():
            return rows
        root = ET.fromstring(zf.read(sheet_path))
        for row in root.findall(".//x:sheetData/x:row", NS_MAIN):
            row_map = {}
            style_map = {}
            for c in row.findall("x:c", NS_MAIN):
                rref = c.attrib.get("r", "")
                idx = xlsx_col_index(rref)
                typ = c.attrib.get("t")
                value = ""
                if typ == "inlineStr":
                    t = c.find("x:is/x:t", NS_MAIN)
                    value = t.text if t is not None and t.text is not None else ""
                else:
                    v = c.find("x:v", NS_MAIN)
                    if v is not None and v.text is not None:
                        value = v.text
                if typ == "s":
                    try:
                        value = shared[int(value)]
                    except Exception:
                        value = ""
                row_map[idx] = value
                style_map[idx] = int(c.attrib.get("s", 0) or 0)
            rows.append({"values": row_map, "styles": style_map})
    return rows


def parse_style_fill_colors(xlsx_path):
    with zipfile.ZipFile(xlsx_path) as zf:
        if "xl/styles.xml" not in zf.namelist():
            return {}
        root = ET.fromstring(zf.read("xl/styles.xml"))
        fills = root.find("x:fills", NS_MAIN)
        cell_xfs = root.find("x:cellXfs", NS_MAIN)
        fill_colors = {}
        if fills is None or cell_xfs is None:
            return fill_colors
        for idx, xf in enumerate(cell_xfs):
            try:
                fill_id = int(xf.attrib.get("fillId", 0))
            except Exception:
                fill_id = 0
            color = ""
            if 0 <= fill_id < len(fills):
                pf = fills[fill_id].find("x:patternFill", NS_MAIN)
                if pf is not None:
                    fg = pf.find("x:fgColor", NS_MAIN)
                    if fg is not None:
                        color = normalize_text(fg.attrib.get("rgb") or fg.attrib.get("indexed") or fg.attrib.get("theme"))
            fill_colors[idx] = color.upper()
        return fill_colors


def find_header_row(rows, expected):
    for i, row in enumerate(rows[:20], start=1):
        vals = {normalize_text(v).upper() for v in row.values()}
        if all(e in vals for e in expected):
            return i, row
    return None, None


def detect_week_columns(rows, scan_rows=30):
    best = {}
    for row in rows[:scan_rows]:
        candidate = {}
        for col, value in row.items():
            txt = normalize_text(value).upper()
            if not txt:
                continue
            week = parse_week(txt)
            if week is None:
                continue
            candidate[col] = week
        if len(candidate) > len(best):
            best = candidate
    return best


def excel_serial_to_date(value):
    try:
        serial = int(float(str(value).strip()))
    except Exception:
        return None
    if serial <= 0:
        return None
    return dt.date(1899, 12, 30) + dt.timedelta(days=serial)


def infer_project_type(cms_code):
    c = normalize_label(cms_code).upper()
    if "WS" in c or "OFFICINA" in c or "WORKSHOP" in c:
        return "WS"
    return "SITE"


def classify_planning_project(project_text):
    code = canonicalize_project_code(project_text)
    if not code:
        return None
    upper = code.upper()
    if upper == "OVERALL COMMESSE OFFICINA":
        return {
            "code": "OVERALL OFFICINA",
            "activity": "OVERALL OFFICINA",
            "type": "WS",
            "workshop_rollup": 1,
        }
    if upper == "WS OVERALL":
        return None
    if upper.startswith("OVERALL COMMESSE"):
        return None
    if "OFFICINA" in upper or "WS" in upper or "WORKSHOP" in upper:
        return {
            "code": code,
            "activity": code,
            "type": "WS",
            "workshop_rollup": 1,
        }
    return {
        "code": code,
        "activity": code,
        "type": infer_project_type(code),
        "workshop_rollup": 0,
    }


OFFICINA_DETAIL_CODES = {
    "388-25 (OFFICINA)",
    "401_25 (OFFICINA)",
    "403-25 (OFFICINA)",
    "414_25 (OFFICINA)",
    "415-25 (OFFICINA)",
    "416-25 (OFFICINA)",
    "422-26 (OFFICINA)",
    "423_26 (OFFICINA)",
    "424_26 (OFFICINA)",
}


def import_selected_workshop_subprojects(demand_xlsx_path, selected_codes=None, replace_existing_demands=True):
    demand_xlsx = Path(demand_xlsx_path)
    if not demand_xlsx.exists():
        raise FileNotFoundError(f"File fabbisogno non trovato: {demand_xlsx}")

    demand_rows = read_sheet_rows(demand_xlsx, "MHRS & RISORSE SCHEDULE")
    week_cols = detect_week_columns(demand_rows)
    if not week_cols:
        raise ValueError("Impossibile leggere le colonne settimana dal file fabbisogno")

    target_codes = {canonicalize_project_code(c) for c in (selected_codes or OFFICINA_DETAIL_CODES)}
    parsed_rows = []
    parsed_projects = {}
    current_project = ""
    for row in demand_rows:
        project_text = normalize_label(row.get(8))
        role = canonicalize_role(row.get(9))
        project_upper = project_text.upper()
        if project_text and not project_upper.startswith("TOT."):
            current_project = project_text
        if not current_project or not role:
            continue
        role_upper = role.upper()
        if role_upper in {"MANSIONE", "TOT. RISORSE PLANNED"}:
            continue
        code = canonicalize_project_code(current_project)
        if code not in target_codes:
            continue
        project_meta = {
            "code": code,
            "activity": code,
            "type": "WS",
            "workshop_rollup": 1,
        }
        parsed_projects[code] = project_meta
        for col, week in week_cols.items():
            qty = parse_int(row.get(col), 0)
            if qty <= 0:
                continue
            parsed_rows.append((code, role, week, qty))

    stats = {
        "projects": len(parsed_projects),
        "demands": 0,
        "demand_source_path": str(demand_xlsx),
    }
    if not parsed_projects:
        return stats

    with db_conn() as conn:
        seen_projects = set()
        for project_meta in parsed_projects.values():
            conn.execute(
                """
                INSERT INTO projects(code, activity, type, workshop_rollup, closed)
                VALUES (?, ?, ?, ?, 0)
                ON CONFLICT(code) DO UPDATE SET
                    activity=excluded.activity,
                    type=excluded.type,
                    workshop_rollup=excluded.workshop_rollup,
                    closed=0
                """,
                (
                    project_meta["code"],
                    project_meta["activity"],
                    project_meta["type"],
                    project_meta["workshop_rollup"],
                ),
            )
            pid = conn.execute("SELECT id FROM projects WHERE code=?", (project_meta["code"],)).fetchone()["id"]
            seen_projects.add(pid)

        if replace_existing_demands and seen_projects:
            placeholders = ",".join("?" for _ in seen_projects)
            conn.execute(f"DELETE FROM demands WHERE project_id IN ({placeholders})", tuple(seen_projects))

        for project_code, role, week, qty in parsed_rows:
            pid = conn.execute("SELECT id FROM projects WHERE code=?", (project_code,)).fetchone()["id"]
            conn.execute(
                "INSERT OR REPLACE INTO demands(project_id, week, role, qty) VALUES (?, ?, ?, ?)",
                (pid, week, role, qty),
            )
            stats["demands"] += 1
    return stats


def import_from_xlsx(xlsx_path, replace_all=True):
    xlsx = Path(xlsx_path)
    if not xlsx.exists():
        raise FileNotFoundError(f"File non trovato: {xlsx}")

    resources_rows = read_sheet_rows(xlsx, "RISORSE")
    demands_rows = read_sheet_rows(xlsx, "FABBISOGNI_LONG")
    alloc_rows = read_sheet_rows(xlsx, "ALLOCAZIONE_TEMPLATE")

    stats = {
        "resources": 0,
        "projects": 0,
        "demands": 0,
        "allocations": 0,
        "source_path": str(xlsx),
    }

    with db_conn() as conn:
        if replace_all:
            conn.execute("DELETE FROM allocations")
            conn.execute("DELETE FROM demands")
            conn.execute("DELETE FROM project_templates")
            conn.execute("DELETE FROM unavailability")
            conn.execute("DELETE FROM projects")
            conn.execute("DELETE FROM resources")

        # Resources (sheet RISORSE)
        h_idx, h_row = find_header_row(resources_rows, ["RISORSA", "MANSIONE1"])
        res_cols = {}
        if h_idx:
            for col, val in h_row.items():
                res_cols[normalize_text(val).upper()] = col
            for row in resources_rows[h_idx:]:
                name = normalize_text(row.get(res_cols.get("RISORSA", -1)))
                if not name:
                    continue
                role1 = canonicalize_role(row.get(res_cols.get("MANSIONE1", -1)))
                role2 = canonicalize_role(row.get(res_cols.get("MANSIONE2", -1)))
                hire = normalize_text(row.get(res_cols.get("ASSUNZIONE", -1)))
                endd = normalize_text(row.get(res_cols.get("FINE", -1)))
                note = normalize_text(row.get(res_cols.get("NOTE", -1)))
                no_travel_txt = normalize_text(row.get(res_cols.get("NO TRASFERTA", -1))).upper()
                no_travel = 1 if no_travel_txt in {"SI", "S", "YES", "Y", "1", "TRUE"} else 0
                conn.execute(
                    """
                    INSERT INTO resources(name, employee_code, role1, role2, hire_date, end_date, note, no_travel)
                    VALUES (?, '', ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(name) DO UPDATE SET
                        role1=excluded.role1,
                        role2=excluded.role2,
                        hire_date=excluded.hire_date,
                        end_date=excluded.end_date,
                        note=excluded.note,
                        no_travel=excluded.no_travel
                    """,
                (name, role1, role2, hire, endd, note, no_travel),
                )
                stats["resources"] += 1

        # Projects + Demands (sheet FABBISOGNI_LONG)
        h_idx, h_row = find_header_row(demands_rows, ["CMS", "WEEK", "MANSIONE"])
        dem_cols = {}
        if h_idx:
            for col, val in h_row.items():
                dem_cols[normalize_text(val).upper()] = col
            for row in demands_rows[h_idx:]:
                cms = normalize_text(row.get(dem_cols.get("CMS", -1)))
                week = parse_week(row.get(dem_cols.get("WEEK", -1)))
                role = canonicalize_role(row.get(dem_cols.get("MANSIONE", -1)))
                qty_raw = normalize_text(
                    row.get(
                        dem_cols.get(
                            "QTA",
                            dem_cols.get("QTY", dem_cols.get("FABBISOGNO", -1)),
                        )
                    )
                )
                if not cms or week is None or not role:
                    continue
                try:
                    qty = int(float(qty_raw))
                except Exception:
                    qty = 0
                ptype = infer_project_type(cms)
                conn.execute(
                    """
                    INSERT INTO projects(code, activity, type, workshop_rollup)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(code) DO UPDATE SET
                        activity=excluded.activity,
                        type=excluded.type,
                        workshop_rollup=excluded.workshop_rollup
                    """,
                    (cms, cms, ptype, 1 if ptype == "WS" else 0),
                )
                pid = conn.execute("SELECT id FROM projects WHERE code=?", (cms,)).fetchone()["id"]
                if qty > 0:
                    conn.execute(
                        "INSERT OR REPLACE INTO demands(project_id, week, role, qty) VALUES (?, ?, ?, ?)",
                        (pid, week, role, qty),
                    )
                    stats["demands"] += 1

        # Allocations (sheet ALLOCAZIONE_TEMPLATE)
        # Struttura attesa: col1=COMMESSA, col2=MANSIONE, col3=NOME RISORSA, col4=WEEK_DA, col5=WEEK_A
        alloc_header_row = None
        for i, row in enumerate(alloc_rows, start=1):
            c1 = normalize_text(row.get(1, "")).upper()
            c4 = normalize_text(row.get(4, "")).upper()
            c5 = normalize_text(row.get(5, "")).upper()
            if c1 in {"COMMESSA", "CMS"} and ("WEEK_DA" in c4 or "WEEK DA" in c4) and ("WEEK_A" in c5 or "WEEK A" in c5):
                alloc_header_row = i
                break
        start_idx = alloc_header_row + 1 if alloc_header_row else 1
        for row in alloc_rows[start_idx:]:
            cms = normalize_text(row.get(1))
            role = canonicalize_role(row.get(2))
            rname = normalize_text(row.get(3))
            wd = parse_week(row.get(4))
            wa = parse_week(row.get(5))
            if not cms or not role or not rname or wd is None or wa is None:
                continue
            rname = normalize_text(rname.split("|")[0])
            rid_row = conn.execute("SELECT id FROM resources WHERE name=?", (rname,)).fetchone()
            if not rid_row:
                continue
            rid = rid_row["id"]
            ptype = infer_project_type(cms)
            conn.execute(
                """
                INSERT INTO projects(code, activity, type, workshop_rollup)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                    activity=excluded.activity,
                    type=excluded.type,
                    workshop_rollup=excluded.workshop_rollup
                """,
                (cms, cms, ptype, 1 if ptype == "WS" else 0),
            )
            pid = conn.execute("SELECT id FROM projects WHERE code=?", (cms,)).fetchone()["id"]
            conn.execute(
                """
                INSERT INTO allocations(project_id, resource_id, role, week_from, week_to, note)
                VALUES (?, ?, ?, ?, ?, '')
                """,
                (pid, rid, role, wd, wa),
            )
            stats["allocations"] += 1

        stats["projects"] = conn.execute("SELECT COUNT(*) n FROM projects").fetchone()["n"]
        conn.commit()

    return stats


def import_from_resource_planning_files(demand_xlsx_path, resources_xlsx_path, replace_all=True):
    demand_xlsx = Path(demand_xlsx_path)
    resources_xlsx = Path(resources_xlsx_path)
    if not demand_xlsx.exists():
        raise FileNotFoundError(f"File fabbisogno non trovato: {demand_xlsx}")
    if not resources_xlsx.exists():
        raise FileNotFoundError(f"File risorse non trovato: {resources_xlsx}")

    demand_rows = read_sheet_rows(demand_xlsx, "MHRS & RISORSE SCHEDULE")
    resources_rows = read_sheet_rows(resources_xlsx, "Foglio1")

    week_cols = detect_week_columns(demand_rows)
    if not week_cols:
        raise ValueError("Impossibile leggere le colonne settimana dal file fabbisogno")

    stats = {
        "resources": 0,
        "projects": 0,
        "demands": 0,
        "allocations": 0,
        "demand_source_path": str(demand_xlsx),
        "resources_source_path": str(resources_xlsx),
    }

    with db_conn() as conn:
        if replace_all:
            conn.execute("DELETE FROM allocations")
            conn.execute("DELETE FROM demands")
            conn.execute("DELETE FROM project_templates")
            conn.execute("DELETE FROM unavailability")
            conn.execute("DELETE FROM projects")
            conn.execute("DELETE FROM resources")

        res_header_idx, res_header = find_header_row(resources_rows, ["NOME RIS.", "MANSIONE RIS."])
        if not res_header_idx:
            raise ValueError("Intestazioni risorse non trovate nel file risorse.xlsx")
        res_cols = {}
        for col, val in res_header.items():
            res_cols[normalize_label(val).upper()] = col
        for row in resources_rows[res_header_idx:]:
            name = normalize_label(row.get(res_cols.get("NOME RIS.", -1)))
            role1 = canonicalize_role(row.get(res_cols.get("MANSIONE RIS.", -1)))
            hire = normalize_label(row.get(res_cols.get("ASSUNZ.", -1)))
            if not name or not role1:
                continue
            conn.execute(
                """
                INSERT INTO resources(name, employee_code, role1, role2, hire_date, end_date, note, no_travel)
                VALUES (?, '', ?, '', ?, '', '', 0)
                ON CONFLICT(name) DO UPDATE SET
                    role1=excluded.role1,
                    hire_date=excluded.hire_date
                """,
                (name, role1, hire),
            )
            stats["resources"] += 1

        current_project = ""
        for row in demand_rows:
            project_text = normalize_label(row.get(8))
            role = canonicalize_role(row.get(9))

            project_upper = project_text.upper()
            if project_text and not project_upper.startswith("TOT."):
                current_project = project_text
            if not current_project or not role:
                continue
            role_upper = role.upper()
            if role_upper in {"MANSIONE", "TOT. RISORSE PLANNED"}:
                continue

            project_meta = classify_planning_project(current_project)
            if not project_meta:
                continue
            conn.execute(
                """
                INSERT INTO projects(code, activity, type, workshop_rollup)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                    activity=excluded.activity,
                    type=excluded.type,
                    workshop_rollup=excluded.workshop_rollup
                """,
                (
                    project_meta["code"],
                    project_meta["activity"],
                    project_meta["type"],
                    project_meta["workshop_rollup"],
                ),
            )
            pid = conn.execute("SELECT id FROM projects WHERE code=?", (project_meta["code"],)).fetchone()["id"]

            for col, week in week_cols.items():
                qty = parse_int(row.get(col), 0)
                if qty <= 0:
                    continue
                conn.execute(
                    "INSERT OR REPLACE INTO demands(project_id, week, role, qty) VALUES (?, ?, ?, ?)",
                    (pid, week, role, qty),
                )
                stats["demands"] += 1

        stats["projects"] = conn.execute("SELECT COUNT(*) n FROM projects").fetchone()["n"]
        conn.commit()

    return stats


def import_demands_only_from_planning_file(
    demand_xlsx_path,
    replace_existing_demands=True,
    role_mapping=None,
    create_roles=None,
):
    demand_xlsx = Path(demand_xlsx_path)
    if not demand_xlsx.exists():
        raise FileNotFoundError(f"File fabbisogno non trovato: {demand_xlsx}")

    demand_rows = read_sheet_rows(demand_xlsx, "MHRS & RISORSE SCHEDULE")
    week_cols = detect_week_columns(demand_rows)
    if not week_cols:
        raise ValueError("Impossibile leggere le colonne settimana dal file fabbisogno")

    stats = {
        "projects": 0,
        "demands": 0,
        "demand_source_path": str(demand_xlsx),
    }

    role_mapping = role_mapping or {}
    create_roles = create_roles or []
    normalized_mapping = {}
    ignored_roles = set()
    for key, value in role_mapping.items():
        src = canonicalize_role(key)
        if not src:
            continue
        if str(value).strip() == "__ignore__":
            ignored_roles.add(src)
            continue
        dest = canonicalize_role(value)
        if dest:
            normalized_mapping[src] = dest
    create_roles = {canonicalize_role(r) for r in create_roles if canonicalize_role(r)}

    parsed_rows = []
    parsed_projects = {}
    unknown_roles = set()
    current_project = ""
    for row in demand_rows:
        project_text = normalize_label(row.get(8))
        role = canonicalize_role(row.get(9))
        if role in normalized_mapping:
            role = normalized_mapping[role]
        if role in ignored_roles:
            continue

        project_upper = project_text.upper()
        if project_text and not project_upper.startswith("TOT."):
            current_project = project_text
        if not current_project or not role:
            continue
        role_upper = role.upper()
        if role_upper in {"MANSIONE", "TOT. RISORSE PLANNED"}:
            continue
        if not is_standard_role(role) and role not in create_roles:
            unknown_roles.add(role)

        project_meta = classify_planning_project(current_project)
        if not project_meta:
            continue
        parsed_projects[project_meta["code"]] = project_meta
        for col, week in week_cols.items():
            qty = parse_int(row.get(col), 0)
            if qty <= 0:
                continue
            parsed_rows.append((project_meta["code"], role, week, qty))

    with db_conn() as conn:
        seen_projects = set()
        for project_meta in parsed_projects.values():
            conn.execute(
                """
                INSERT INTO projects(code, activity, type, workshop_rollup)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                    activity=excluded.activity,
                    type=excluded.type,
                    workshop_rollup=excluded.workshop_rollup
                """,
                (
                    project_meta["code"],
                    project_meta["activity"],
                    project_meta["type"],
                    project_meta["workshop_rollup"],
                ),
            )
            pid = conn.execute("SELECT id FROM projects WHERE code=?", (project_meta["code"],)).fetchone()["id"]
            seen_projects.add(pid)

        if replace_existing_demands and seen_projects:
            placeholders = ",".join("?" for _ in seen_projects)
            conn.execute(f"DELETE FROM demands WHERE project_id IN ({placeholders})", tuple(seen_projects))

        if create_roles:
            for role in create_roles:
                conn.execute("INSERT OR IGNORE INTO roles(name) VALUES (?)", (role,))
            refresh_custom_roles(conn)

        for project_code, role, week, qty in parsed_rows:
            if not is_standard_role(role):
                continue
            pid = conn.execute("SELECT id FROM projects WHERE code=?", (project_code,)).fetchone()["id"]
            conn.execute(
                "INSERT OR REPLACE INTO demands(project_id, week, role, qty) VALUES (?, ?, ?, ?)",
                (pid, week, role, qty),
            )
            stats["demands"] += 1

        stats["projects"] = len(seen_projects)
        conn.commit()

    stats["unknown_roles"] = sorted(unknown_roles)
    stats["unknown_roles_skipped"] = bool(unknown_roles)
    return stats


def _parse_demands_from_planning_file(demand_xlsx_path):
    demand_xlsx = Path(demand_xlsx_path)
    if not demand_xlsx.exists():
        raise FileNotFoundError(f"File fabbisogno non trovato: {demand_xlsx}")

    demand_rows = read_sheet_rows(demand_xlsx, "MHRS & RISORSE SCHEDULE")
    week_cols = detect_week_columns(demand_rows)
    if not week_cols:
        raise ValueError("Impossibile leggere le colonne settimana dal file fabbisogno")

    parsed_rows = []
    parsed_projects = {}
    unknown_roles = set()
    current_project = ""
    for row in demand_rows:
        project_text = normalize_label(row.get(8))
        role = canonicalize_role(row.get(9))

        project_upper = project_text.upper()
        if project_text and not project_upper.startswith("TOT."):
            current_project = project_text
        if not current_project or not role:
            continue
        role_upper = role.upper()
        if role_upper in {"MANSIONE", "TOT. RISORSE PLANNED"}:
            continue
        if not is_standard_role(role):
            unknown_roles.add(role)

        project_meta = classify_planning_project(current_project)
        if not project_meta:
            continue
        parsed_projects[project_meta["code"]] = project_meta
        for col, week in week_cols.items():
            qty = parse_int(row.get(col), 0)
            if qty <= 0:
                continue
            parsed_rows.append((project_meta["code"], role, week, qty))

    return week_cols, parsed_projects, parsed_rows, sorted(unknown_roles)


def _build_demand_map_from_rows(rows):
    demand_map = {}
    for project_code, role, week, qty in rows:
        key = (project_code, role, int(week))
        demand_map[key] = demand_map.get(key, 0) + int(qty)
    return demand_map


def _fetch_existing_demands(conn, project_codes):
    if not project_codes:
        return {}, {}
    placeholders = ",".join("?" for _ in project_codes)
    rows = conn.execute(
        f"""
        SELECT p.code AS project_code, d.role, d.week, d.qty
        FROM demands d
        JOIN projects p ON p.id = d.project_id
        WHERE p.code IN ({placeholders})
        """,
        tuple(project_codes),
    ).fetchall()
    demand_map = {}
    by_project = {}
    for row in rows:
        key = (row["project_code"], row["role"], int(row["week"]))
        demand_map[key] = int(row["qty"])
        proj = by_project.setdefault(row["project_code"], {"weeks": set(), "total": 0})
        proj["weeks"].add(int(row["week"]))
        proj["total"] += int(row["qty"])
    return demand_map, by_project


def _build_allocation_map(conn, project_codes):
    if not project_codes:
        return {}

    placeholders = ",".join("?" for _ in project_codes)

    rows = conn.execute(
        f"""
        SELECT a.id, a.resource_id, p.code AS project_code, a.role, a.week_from, a.week_to
        FROM allocations a
        JOIN projects p ON p.id = a.project_id
        WHERE p.code IN ({placeholders})
        """,
        tuple(project_codes),
    ).fetchall()

    from collections import defaultdict

    grouped = defaultdict(list)

    # Raggruppo per RISORSA + SETTIMANA
    for r in rows:
        for week in range(int(r["week_from"]), int(r["week_to"]) + 1):
            grouped[(r["resource_id"], week)].append(r)

    alloc_map = {}

    for (res_id, week), items in grouped.items():
        n = len(items)
        if n == 0:
            continue

        weight = 1.0 / n

        for r in items:
            key = (r["project_code"], r["role"], week)
            alloc_map[key] = alloc_map.get(key, 0) + weight

    return alloc_map


def analyze_demands_import(demand_xlsx_path):
    week_cols, parsed_projects, parsed_rows, unknown_roles = _parse_demands_from_planning_file(demand_xlsx_path)
    project_codes = sorted(parsed_projects.keys())
    new_map = _build_demand_map_from_rows(parsed_rows)
    current_week = int(dt.date.today().isocalendar().week)

    with db_conn() as conn:
        prev_map, prev_by_project = _fetch_existing_demands(conn, project_codes)
        alloc_map = _build_allocation_map(conn, project_codes)

    diffs = []
    all_keys = set(prev_map.keys()) | set(new_map.keys())
    for key in all_keys:
        if int(key[2]) < current_week:
            continue
        prev_qty = prev_map.get(key, 0)
        new_qty = new_map.get(key, 0)
        if prev_qty == new_qty:
            continue
        pct = None
        if prev_qty > 0:
            pct = round(((new_qty - prev_qty) / prev_qty) * 100, 1)
        diffs.append(
            {
                "project": key[0],
                "role": key[1],
                "week": key[2],
                "prev_qty": prev_qty,
                "new_qty": new_qty,
                "pct_change": pct,
            }
        )

    project_summaries = []
    for code in project_codes:
        prev_info = prev_by_project.get(code, {"weeks": set(), "total": 0})
        new_weeks = {k[2] for k in new_map.keys() if k[0] == code and int(k[2]) >= current_week}
        new_total = sum(v for k, v in new_map.items() if k[0] == code and int(k[2]) >= current_week)
        prev_total = sum(v for k, v in prev_map.items() if k[0] == code and int(k[2]) >= current_week)
        prev_weeks = {w for w in prev_info["weeks"] if int(w) >= current_week}
        prev_min = min(prev_weeks) if prev_weeks else None
        prev_max = max(prev_weeks) if prev_weeks else None
        new_min = min(new_weeks) if new_weeks else None
        new_max = max(new_weeks) if new_weeks else None
        change_type = "invariata"
        if prev_max and new_max:
            if new_max > prev_max:
                change_type = "allungata"
            elif new_max < prev_max:
                change_type = "accorciata"
        if new_total > prev_total:
            change_type = "aumentata" if change_type == "invariata" else change_type
        elif new_total < prev_total:
            change_type = "ridotta" if change_type == "invariata" else change_type
        pct_total = None
        if prev_total > 0:
            pct_total = round(((new_total - prev_total) / prev_total) * 100, 1)
        project_summaries.append(
            {
                "project": code,
                "prev_total": prev_total,
                "new_total": new_total,
                "pct_change": pct_total,
                "prev_week_min": prev_min,
                "prev_week_max": prev_max,
                "new_week_min": new_min,
                "new_week_max": new_max,
                "change_type": change_type,
            }
        )

    conflicts = []
    for key in set(new_map.keys()) | set(prev_map.keys()):
        if int(key[2]) < current_week:
            continue
        new_qty = new_map.get(key, 0)
        allocated = alloc_map.get(key, 0)
        if allocated == new_qty:
            continue
        pct = None
        if new_qty > 0:
            pct = round(((allocated - new_qty) / new_qty) * 100, 1)
        conflicts.append(
            {
                "project": key[0],
                "role": key[1],
                "week": key[2],
                "demand": new_qty,
                "allocated": allocated,
                "delta": allocated - new_qty,
                "pct_delta": pct,
            }
        )

    return {
        "projects": project_codes,
        "diffs": diffs,
        "project_summaries": project_summaries,
        "allocation_conflicts": conflicts,
        "weeks": sorted({k[2] for k in new_map.keys() if int(k[2]) >= current_week}),
        "current_week": current_week,
        "unknown_roles": unknown_roles,
    }


def bootstrap_from_xlsx_if_empty():
    if not BOOTSTRAP_XLSX.exists():
        return
    with db_conn() as conn:
        n_res = conn.execute("SELECT COUNT(*) n FROM resources").fetchone()["n"]
        n_proj = conn.execute("SELECT COUNT(*) n FROM projects").fetchone()["n"]
        n_dem = conn.execute("SELECT COUNT(*) n FROM demands").fetchone()["n"]
        if n_res > 0 and n_proj > 0 and n_dem > 0:
            return
    try:
        import_from_xlsx(BOOTSTRAP_XLSX, replace_all=False)
    except Exception:
        return


PLAN_IGNORE_TOKENS = {
    "",
    "X",
    "CIG",
    "FE",
    "F",
    "M",
    "FERIE",
    "MAL",
    "MALATTIA",
    "INF",
    "VAL",
    "!!!",
}


def canonicalize_plan_token(value):
    token = normalize_label(value).upper()
    if token in PLAN_IGNORE_TOKENS:
        return ""
    return token


def map_plan_token_to_project(project_index, token):
    if not token:
        return None
    if token in {"OFF", "OFFICINA", "OVERALL OFFICINA"}:
        return "OVERALL OFFICINA"
    if token in project_index:
        return project_index[token]
    if token.isdigit():
        return project_index.get(token)
    return None


def source_role_suggestion(role_text):
    role = normalize_label(role_text).upper()
    if not role:
        return ""
    if role in all_standard_roles():
        return role
    aliases = {
        "CAPO CANT": "CAPO CANTIERE",
        "CAPO SQ": "CAPO SQUADRA",
        "PONT": "PONTEGGIATORE",
        "MONTATORE CARP": "MONTATORE",
    }
    return aliases.get(role, "")


def clamp_proposal_weeks(week_from, week_to, hire_date_text="", end_date_text=""):
    current_week = int(dt.date.today().isocalendar().week)
    wf = max(int(week_from), current_week)
    wt = int(week_to)

    hire_date = parse_date_flexible(hire_date_text)
    if hire_date and hire_date.year == BASE_YEAR:
        hire_week = int(hire_date.isocalendar().week)
        wf = max(wf, hire_week)

    end_date = parse_date_flexible(end_date_text)
    if end_date:
        if end_date.year < BASE_YEAR or (end_date.year == BASE_YEAR and int(end_date.isocalendar().week) < current_week):
            return None
        if end_date.year == BASE_YEAR:
            wt = min(wt, int(end_date.isocalendar().week))

    if wt < wf:
        return None
    return wf, wt


def analyze_foglio2_plan(plan_xlsx_path):
    plan_xlsx = Path(plan_xlsx_path)
    if not plan_xlsx.exists():
        raise FileNotFoundError(f"File planning non trovato: {plan_xlsx}")

    styled_rows = read_sheet_rows_with_styles(plan_xlsx, "Foglio2")
    if not styled_rows:
        raise ValueError("Foglio2 non trovato o vuoto")
    fill_colors = parse_style_fill_colors(plan_xlsx)

    header_values = styled_rows[0]["values"]
    date_cols = {}
    for col, value in header_values.items():
        d = excel_serial_to_date(value)
        if d:
            date_cols[col] = d
    if not date_cols:
        raise ValueError("Impossibile leggere le colonne data da Foglio2")

    with db_conn() as conn:
        resources = [row_to_dict(r) for r in conn.execute("SELECT * FROM resources ORDER BY name")]
        projects = [
            row_to_dict(r)
            for r in conn.execute("SELECT id, code, closed FROM projects WHERE closed=0 ORDER BY code")
        ]

    resource_index = {normalize_person_name_key(r["name"]): r for r in resources}
    project_index = {}
    for p in projects:
        code = normalize_label(p["code"])
        upper = code.upper()
        if upper == "OVERALL OFFICINA":
            project_index["OVERALL OFFICINA"] = code
            project_index["OFF"] = code
            continue
        short = code.split("-", 1)[0].split("_", 1)[0].strip().upper()
        if short and short not in project_index:
            project_index[short] = code

    matched_resources = []
    new_resources = []
    ceased_resources = []
    unknown_project_tokens = {}
    proposals = []

    for styled_row in styled_rows[1:]:
        row = styled_row["values"]
        styles = styled_row["styles"]
        name = normalize_label(row.get(2))
        if not name or name.upper() == "CENTRO DI COSTO":
            continue
        role_source = normalize_label(row.get(3))
        note = normalize_label(row.get(29))
        end_date = normalize_label(row.get(27))
        name_fill = fill_colors.get(int(styles.get(2, 0) or 0), "")
        is_ceased = name_fill == "FFFF0000"
        suggested_role = source_role_suggestion(role_source)
        master = find_matching_resource(resource_index, name)
        effective_hire_date = master.get("hire_date", "") if master else normalize_label(row.get(26))
        effective_end_date = master.get("end_date", "") if master else end_date

        week_project_counts = {}
        for col, day in date_cols.items():
            token = canonicalize_plan_token(row.get(col))
            if not token:
                continue
            week = int(day.isocalendar().week)
            mapped_project = map_plan_token_to_project(project_index, token)
            if not mapped_project:
                unknown_project_tokens[token] = unknown_project_tokens.get(token, 0) + 1
                continue
            bucket = week_project_counts.setdefault(week, {})
            bucket[mapped_project] = bucket.get(mapped_project, 0) + 1

        if is_ceased:
            ceased_resources.append(
                {
                    "name": name,
                    "source_role": role_source,
                    "suggested_role": suggested_role,
                    "end_date": end_date,
                    "note": note,
                }
            )
            continue

        target_list = matched_resources if master else new_resources
        target_list.append(
                {
                    "name": name,
                    "source_role": role_source,
                    "suggested_role": suggested_role,
                    "master_role1": master.get("role1", "") if master else "",
                    "master_role2": master.get("role2", "") if master else "",
                    "hire_date": effective_hire_date if not master else "",
                    "end_date": effective_end_date if not master else "",
                    "note": note,
                }
            )

        if not week_project_counts:
            continue

        weekly_projects = []
        for week in sorted(week_project_counts):
            best_project = sorted(
                week_project_counts[week].items(),
                key=lambda item: (-item[1], item[0]),
            )[0][0]
            weekly_projects.append((week, best_project))

        start_week = None
        end_week = None
        current_project = None
        for week, project_code in weekly_projects:
            if current_project is None:
                current_project = project_code
                start_week = week
                end_week = week
                continue
            if project_code == current_project and week == end_week + 1:
                end_week = week
                continue
            clamped = clamp_proposal_weeks(start_week, end_week, effective_hire_date, effective_end_date)
            if clamped:
                wf, wt = clamped
                proposals.append(
                    {
                        "name": name,
                        "source_role": role_source,
                        "suggested_role": suggested_role,
                        "project": current_project,
                        "week_from": wf,
                        "week_to": wt,
                        "matched_master": bool(master),
                    }
                )
            current_project = project_code
            start_week = week
            end_week = week
        if current_project is not None:
            clamped = clamp_proposal_weeks(start_week, end_week, effective_hire_date, effective_end_date)
            if clamped:
                wf, wt = clamped
                proposals.append(
                    {
                        "name": name,
                        "source_role": role_source,
                        "suggested_role": suggested_role,
                        "project": current_project,
                        "week_from": wf,
                        "week_to": wt,
                        "matched_master": bool(master),
                    }
                )

    matched_resources.sort(key=lambda r: r["name"])
    new_resources.sort(key=lambda r: r["name"])
    ceased_resources.sort(key=lambda r: r["name"])
    proposals.sort(key=lambda r: (0 if r["matched_master"] else 1, r["project"], r["name"], r["week_from"]))
    unknown_roles = sorted(
        [r for r in new_resources + matched_resources if r.get("source_role") and not r.get("suggested_role")],
        key=lambda r: (r["source_role"], r["name"]),
    )
    unknown_projects = [
        {"token": token, "count": count}
        for token, count in sorted(unknown_project_tokens.items(), key=lambda item: (-item[1], item[0]))
    ]

    return {
        "source_path": str(plan_xlsx),
        "summary": {
            "matched_resources": len(matched_resources),
            "new_resources": len(new_resources),
            "ceased_excluded": len(ceased_resources),
            "unknown_roles": len(unknown_roles),
            "proposal_rows": len(proposals),
        },
        "matched_resources": matched_resources,
        "new_resources": new_resources,
        "ceased_resources": ceased_resources,
        "unknown_roles": unknown_roles,
        "unknown_projects": unknown_projects,
        "proposals": proposals,
    }


def row_to_dict(row):
    return {k: row[k] for k in row.keys()}


def resource_status(res):
    end_date = normalize_text(res["end_date"])
    if end_date:
        if end_date.upper() in {"N.D.", "ND", "N/D"}:
            return "CESSATO"
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                if dt.datetime.strptime(end_date, fmt).date() < dt.date.today():
                    return "CESSATO"
            except ValueError:
                pass
    return "ATTIVO"


def ensure_project_template(conn, project_id, week_from, week_to, use_standard_roles=True):
    wf = int(week_from)
    wt = int(week_to)
    if wf < 1:
        wf = 1
    if wt > 52:
        wt = 52
    if wt < wf:
        wt = wf
    conn.execute(
        """
        INSERT INTO project_templates(project_id, week_from, week_to, use_standard_roles)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(project_id) DO UPDATE SET
            week_from=excluded.week_from,
            week_to=excluded.week_to,
            use_standard_roles=excluded.use_standard_roles
        """,
        (project_id, wf, wt, 1 if use_standard_roles else 0),
    )


def generate_demands_template(conn, project_id, week_from, week_to, replace_existing=False):
    wf = int(week_from)
    wt = int(week_to)
    if wf < 1:
        wf = 1
    if wt > 52:
        wt = 52
    if wt < wf:
        wt = wf
    if replace_existing:
        conn.execute("DELETE FROM demands WHERE project_id=? AND week BETWEEN ? AND ?", (project_id, wf, wt))
    for w in range(wf, wt + 1):
        for role in all_standard_roles():
            conn.execute(
                """
                INSERT OR IGNORE INTO demands(project_id, week, role, qty)
                VALUES (?, ?, ?, 0)
                """,
                (project_id, w, role),
            )


def get_projects_with_template(conn):
    return [
        row_to_dict(r)
        for r in conn.execute(
            """
            SELECT p.id, p.code, p.activity, p.type, p.closed, p.workshop_rollup,
                   t.week_from, t.week_to, t.use_standard_roles
            FROM projects p
            LEFT JOIN project_templates t ON t.project_id = p.id
            ORDER BY p.code
            """
        )
    ]


def allocation_warning_segments(conn, alloc_row):
    rid = alloc_row["resource_id"]
    aid = alloc_row["id"]
    wf = int(alloc_row["week_from"])
    wt = int(alloc_row["week_to"])
    segments = []

    res = conn.execute("SELECT * FROM resources WHERE id=?", (rid,)).fetchone()
    if not res:
        segments.append({"type": "RISORSA MANCANTE", "week_from": wf, "week_to": wt})
        return segments

    alloc_role = role_key(alloc_row["role"])
    res_roles = {role_key(res["role1"]), role_key(res["role2"])}
    res_roles.discard("")
    if alloc_role and alloc_role not in res_roles:
        segments.append({"type": "MANSIONE NON COERENTE", "week_from": wf, "week_to": wt})

    status = resource_status(res)
    if status == "CESSATO":
        segments.append({"type": "CESSATO", "week_from": wf, "week_to": wt})
    end_w = date_to_week(res["end_date"])
    if end_w > 0 and end_w < 999 and wt > end_w:
        segments.append({"type": f"CONTRATTO W{end_w}", "week_from": max(wf, end_w + 1), "week_to": wt})

    if not is_external_resource_row(res):
        overlap_rows = conn.execute(
            """
            SELECT week_from, week_to
            FROM allocations a
            WHERE a.resource_id = ?
              AND a.id <> ?
              AND a.week_from <= ?
              AND a.week_to >= ?
            """,
            (rid, aid, wt, wf),
        ).fetchall()
        for ov in overlap_rows:
            ov_wf = max(wf, int(ov["week_from"]))
            ov_wt = min(wt, int(ov["week_to"]))
            if ov_wt >= ov_wf:
                segments.append({"type": "SOVRAPPOSIZIONE", "week_from": ov_wf, "week_to": ov_wt})

    unav_rows = conn.execute(
        """
        SELECT week_from, week_to
        FROM unavailability u
        WHERE u.resource_id = ?
          AND u.week_from <= ?
          AND u.week_to >= ?
        """,
        (rid, wt, wf),
    ).fetchall()
    for unav in unav_rows:
        unav_wf = max(wf, int(unav["week_from"]))
        unav_wt = min(wt, int(unav["week_to"]))
        if unav_wt >= unav_wf:
            segments.append({"type": "INDISPONIBILE", "week_from": unav_wf, "week_to": unav_wt})
    return segments


def allocation_warnings(conn, alloc_row):
    warnings = []
    for seg in allocation_warning_segments(conn, alloc_row):
        label = seg["type"]
        if label not in warnings:
            warnings.append(label)
    return warnings


def compute_analysis(conn):
    coverage_rows = []
    demand_rows = conn.execute(
        """
        SELECT d.id, d.week, d.role, d.qty, p.code AS project, p.type, p.workshop_rollup
        FROM demands d
        JOIN projects p ON p.id = d.project_id
        WHERE p.closed = 0
          AND d.qty > 0
        ORDER BY d.week, p.code, d.role
        """
    ).fetchall()
    alloc_map = _build_allocation_map(conn, [r["project"] for r in demand_rows])
    for d in demand_rows:
        if d["type"] == "WS" and d["workshop_rollup"] == 1 and d["project"] != "OVERALL OFFICINA":
            continue
        allocated = alloc_map.get((d["project"], d["role"], d["week"]), 0)
        gap = max(int(d["qty"]) - float(allocated), 0)
        coverage_rows.append(
            {
                "week": int(d["week"]),
                "project": d["project"],
                "role": d["role"],
                "demand": int(d["qty"]),
                "allocated": float(allocated),
                "gap": float(gap),
                "scope": "PROJECT",
            }
        )

    # Rollup officina: usa le sottocommesse WS quando presenti, altrimenti il riepilogo diretto OVERALL OFFICINA.
    ws_overall_direct = {}
    for wr in conn.execute(
        """
        SELECT d.week, d.role, d.qty AS demand
        FROM demands d
        JOIN projects p ON p.id = d.project_id
        WHERE p.closed = 0
          AND UPPER(p.code) = 'OVERALL OFFICINA'
          AND d.qty > 0
        """
    ):
        ws_overall_direct[(int(wr["week"]), normalize_text(wr["role"]))] = int(wr["demand"])

    ws_child_rollup = {}
    for wr in conn.execute(
        """
        SELECT d.week, d.role, SUM(d.qty) AS demand
        FROM demands d
        JOIN projects p ON p.id = d.project_id
        WHERE p.closed = 0
          AND p.type = 'WS'
          AND p.workshop_rollup = 1
          AND UPPER(p.code) NOT IN ('OVERALL OFFICINA', 'WS OVERALL')
          AND d.qty > 0
        GROUP BY d.week, d.role
        """
    ):
        ws_child_rollup[(int(wr["week"]), normalize_text(wr["role"]))] = int(wr["demand"])

    ws_effective = dict(ws_overall_direct)
    ws_effective.update(ws_child_rollup)

    for (week, role), demand in sorted(ws_effective.items()):
        allocated = alloc_map.get(("OVERALL OFFICINA", role, week), 0)
        gap = max(demand - float(allocated), 0)
        coverage_rows.append(
            {
                "week": week,
                "project": "OVERALL OFFICINA",
                "role": role,
                "demand": demand,
                "allocated": float(allocated),
                "gap": float(gap),
                "scope": "WS_OVERALL",
            }
        )

    roles = sorted(
        {
            canonicalize_role(r["role1"])
            for r in conn.execute("SELECT role1 FROM resources")
            if canonicalize_role(r["role1"])
        }
        | {
            canonicalize_role(r["role2"])
            for r in conn.execute("SELECT role2 FROM resources")
            if canonicalize_role(r["role2"])
        }
        | {canonicalize_role(r["role"]) for r in conn.execute("SELECT role FROM demands") if canonicalize_role(r["role"])}
    )

    free_rows = []
    all_resources = [row_to_dict(r) for r in conn.execute("SELECT * FROM resources")]
    unav_rows = [row_to_dict(u) for u in conn.execute("SELECT * FROM unavailability")]
    alloc_rows = [row_to_dict(a) for a in conn.execute("SELECT * FROM allocations")]
    for week in range(1, 53):
        for role in roles:
            free = 0
            for res in all_resources:
                if is_external_resource_row(res):
                    continue
                if resource_status(res) == "CESSATO":
                    continue
                hw = parse_week(res["hire_date"]) or 1
                ew = date_to_week(res["end_date"])
                if week < hw:
                    continue
                if ew < 999 and week > ew:
                    continue
                if role_key(role) not in {role_key(res["role1"]), role_key(res["role2"])}:
                    continue
                blocked = False
                for u in unav_rows:
                    if u["resource_id"] == res["id"] and int(u["week_from"]) <= week <= int(u["week_to"]):
                        blocked = True
                        break
                if blocked:
                    continue
                assigned = False
                for a in alloc_rows:
                    if a["resource_id"] == res["id"] and int(a["week_from"]) <= week <= int(a["week_to"]):
                        assigned = True
                        break
                if not assigned:
                    free += 1
            free_rows.append({"week": week, "role": role, "free": free})

    return {"coverage": coverage_rows, "free_by_role": free_rows}


def compute_v13_template(conn):
    project_map = {
        int(r["id"]): normalize_text(r["code"])
        for r in conn.execute("SELECT id, code FROM projects WHERE closed=0")
    }

    demand_map = {}
    combos = set()
    for r in conn.execute(
        """
        SELECT d.project_id, d.week, d.role, d.qty
        FROM demands d
        JOIN projects p ON p.id=d.project_id
        WHERE p.closed=0
        """
    ):
        pid = int(r["project_id"])
        role = normalize_text(r["role"])
        week = int(r["week"])
        qty = int(r["qty"])
        demand_map[(pid, role, week)] = qty
        if qty > 0:
            combos.add((pid, role))

    alloc_by_combo = {}
    for a in conn.execute(
        """
        SELECT a.*, p.code AS project, rs.name AS resource
        FROM allocations a
        JOIN projects p ON p.id = a.project_id
        JOIN resources rs ON rs.id = a.resource_id
        WHERE p.closed=0
        ORDER BY p.code, a.role, a.week_from, a.id
        """
    ):
        pid = int(a["project_id"])
        role = normalize_text(a["role"])
        combos.add((pid, role))
        alloc_by_combo.setdefault((pid, role), []).append(a)

    rows = []
    for pid, role in sorted(combos, key=lambda x: (project_map.get(x[0], ""), x[1])):
        project = project_map.get(pid, "")
        allocs = alloc_by_combo.get((pid, role), [])

        week_cells = []
        tot_scop = 0
        for week in range(1, 53):
            f = int(demand_map.get((pid, role, week), 0))
            a = 0
            for al in allocs:
                if int(al["week_from"]) <= week <= int(al["week_to"]):
                    a += 1
            d = max(f - a, 0)
            tot_scop += d
            week_cells.append({"f": f, "a": a, "d": d})

        rows.append(
            {
                "row_type": "GROUP",
                "project_id": pid,
                "project": project,
                "role": role,
                "resource": "",
                "resource_id": None,
                "week_from": None,
                "week_to": None,
                "check": "",
                "decision": "",
                "action": "",
                "tot_scop": int(tot_scop),
                "weeks": week_cells,
            }
        )

        demand_weeks = [ix + 1 for ix, c in enumerate(week_cells) if int(c["f"]) > 0]
        default_wf = min(demand_weeks) if demand_weeks else ""
        default_wt = max(demand_weeks) if demand_weeks else ""
        base_slots = max([c["f"] for c in week_cells] + [0])
        # Elasticita': lascia sempre almeno uno slot extra per spezzare periodi.
        total_slots = max(base_slots + 1, len(allocs))
        for idx in range(total_slots):
            if idx < len(allocs):
                al = allocs[idx]
                warnings = allocation_warnings(conn, al)
                check = "OK" if len(warnings) == 0 else " | ".join(warnings)
                wf = int(al["week_from"])
                wt = int(al["week_to"])
                slot_cells = []
                for week in range(1, 53):
                    in_range = wf <= week <= wt
                    slot_cells.append({"f": 1 if in_range else 0, "a": 1 if in_range else 0, "d": 0})
                rows.append(
                    {
                        "row_type": "SLOT",
                        "slot_index": idx + 1,
                        "allocation_id": int(al["id"]),
                        "project_id": pid,
                        "project": project,
                        "role": role,
                        "resource": normalize_text(al["resource"]),
                        "resource_id": int(al["resource_id"]),
                        "week_from": wf,
                        "week_to": wt,
                        "check": check,
                        "decision": "",
                        "action": "",
                        "tot_scop": "",
                        "weeks": slot_cells,
                    }
                )
            else:
                rows.append(
                    {
                        "row_type": "SLOT",
                        "slot_index": idx + 1,
                        "allocation_id": None,
                        "project_id": pid,
                        "project": project,
                        "role": role,
                        "resource": "",
                        "resource_id": None,
                        "week_from": default_wf,
                        "week_to": default_wt,
                        "check": "",
                        "decision": "",
                        "action": "",
                        "tot_scop": "",
                        "weeks": [{"f": 0, "a": 0, "d": 0} for _ in range(52)],
                    }
                )

    return rows


def compute_report_weekly(conn):
    out = []
    all_resources = [row_to_dict(r) for r in conn.execute("SELECT * FROM resources") if not is_external_resource_row(row_to_dict(r))]
    for week in range(1, 53):
        fabbisogno = conn.execute("SELECT COALESCE(SUM(qty),0) AS n FROM demands WHERE week=?", (week,)).fetchone()["n"]
        allocato = conn.execute(
            "SELECT COALESCE(SUM(COALESCE(weight,1)),0) AS n FROM allocations WHERE week_from<=? AND week_to>=?",
            (week, week),
        ).fetchone()["n"]
        scopertura = max(float(fabbisogno) - float(allocato), 0)

        in_forza = 0
        for r in all_resources:
            hw = date_to_week(r.get("hire_date", ""))
            ew = date_to_week(r.get("end_date", ""))
            if hw == 999:
                hw = 1
            if week < hw:
                continue
            if ew < 999 and week > ew:
                continue
            in_forza += 1
        non_allocati = max(float(in_forza) - float(allocato), 0)

        out.append(
            {
                "week": week,
                "required": int(fabbisogno),
                "allocated": float(allocato),
                "residual": float(scopertura),
                "available": int(in_forza),
                "surplus": float(non_allocati),
            }
        )
    return out


def compute_report_role_weekly(conn):
    roles = sorted(
        {
            canonicalize_role(r["role1"])
            for r in conn.execute("SELECT role1 FROM resources")
            if canonicalize_role(r["role1"])
        }
        | {
            canonicalize_role(r["role2"])
            for r in conn.execute("SELECT role2 FROM resources")
            if canonicalize_role(r["role2"])
        }
        | {
            canonicalize_role(r["role"])
            for r in conn.execute("SELECT role FROM demands")
            if canonicalize_role(r["role"])
        }
    )
    all_resources = [row_to_dict(r) for r in conn.execute("SELECT * FROM resources") if not is_external_resource_row(row_to_dict(r))]
    out = []
    for week in range(1, 53):
        for role in roles:
            available = 0
            for r in all_resources:
                hw = date_to_week(r.get("hire_date", ""))
                ew = date_to_week(r.get("end_date", ""))
                if hw == 999:
                    hw = 1
                if week < hw:
                    continue
                if ew < 999 and week > ew:
                    continue
                if role_key(role) not in {role_key(r.get("role1")), role_key(r.get("role2"))}:
                    continue
                available += 1
            required = conn.execute(
                "SELECT COALESCE(SUM(qty),0) AS n FROM demands WHERE week=? AND UPPER(role)=UPPER(?)",
                (week, role),
            ).fetchone()["n"]
            allocated = conn.execute(
                "SELECT COALESCE(SUM(COALESCE(weight,1)),0) AS n FROM allocations WHERE week_from<=? AND week_to>=? AND UPPER(role)=UPPER(?)",
                (week, week, role),
            ).fetchone()["n"]
            surplus = max(float(available) - float(allocated), 0)
            residual = max(float(required) - float(allocated), 0)
            if required or available or allocated:
                out.append(
                    {
                        "week": week,
                        "role": role,
                        "required": int(required),
                        "available": int(available),
                        "allocated": float(allocated),
                        "surplus": float(surplus),
                        "residual": float(residual),
                    }
                )
    return out


class AppHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def log_message(self, fmt, *args):
        return

    def _send_json(self, data, status=HTTPStatus.OK):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length else b"{}"
        return json.loads(raw.decode("utf-8"))

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path == "/api/resources":
            with db_conn() as conn:
                rows = []
                for r in conn.execute("SELECT * FROM resources ORDER BY name"):
                    d = row_to_dict(r)
                    d["status"] = resource_status(r)
                    d["is_external"] = is_external_resource_row(d)
                    rows.append(d)
                return self._send_json(rows)
        if path == "/api/projects":
            with db_conn() as conn:
                rows = get_projects_with_template(conn)
                return self._send_json(rows)
        if path == "/api/demands":
            with db_conn() as conn:
                rows = [
                    row_to_dict(r)
                    for r in conn.execute(
                        """
                        SELECT d.id, d.project_id, p.code AS project, d.week, d.role, d.qty
                        FROM demands d
                        JOIN projects p ON p.id = d.project_id
                        ORDER BY d.week, p.code, d.role
                        """
                    )
                ]
                return self._send_json(rows)
        if path == "/api/allocations":
            with db_conn() as conn:
                rows = []
                for r in conn.execute(
                    """
                    SELECT a.*, p.code AS project, rs.name AS resource
                    FROM allocations a
                    JOIN projects p ON p.id = a.project_id
                    JOIN resources rs ON rs.id = a.resource_id
                    ORDER BY rs.name, a.week_from, p.code
                    """
                ):
                    d = row_to_dict(r)
                    d["warnings"] = allocation_warnings(conn, r)
                    d["warning_segments"] = allocation_warning_segments(conn, r)
                    rows.append(d)
                return self._send_json(rows)
        if path == "/api/logs":
            limit = int(qs.get("limit", [200])[0])
            with db_conn() as conn:
                rows = [
                    row_to_dict(r)
                    for r in conn.execute(
                        "SELECT * FROM activity_log ORDER BY id DESC LIMIT ?",
                        (limit,),
                    )
                ]
            return self._send_json(rows)
        if path == "/api/unavailability":
            with db_conn() as conn:
                rows = [
                    row_to_dict(r)
                    for r in conn.execute(
                        """
                        SELECT u.*, rs.name AS resource
                        FROM unavailability u
                        JOIN resources rs ON rs.id = u.resource_id
                        ORDER BY rs.name, u.week_from, u.week_to
                        """
                    )
                ]
                return self._send_json(rows)
        if path == "/api/workshop-breakdown":
            params = urllib.parse.parse_qs(parsed.query or "")
            week = parse_week((params.get("week") or [""])[0])
            role = canonicalize_role((params.get("role") or [""])[0])
            if week is None or not role:
                return self._send_json([])
            with db_conn() as conn:
                rows = []
                child_rows = conn.execute(
                    """
                    SELECT p.id AS project_id, p.code AS project, d.role, d.week, d.qty
                    FROM demands d
                    JOIN projects p ON p.id = d.project_id
                    WHERE p.closed = 0
                      AND p.type = 'WS'
                      AND p.workshop_rollup = 1
                      AND UPPER(p.code) <> 'OVERALL OFFICINA'
                      AND d.week = ?
                      AND UPPER(d.role) = UPPER(?)
                      AND d.qty > 0
                    ORDER BY d.qty DESC, p.code
                    """,
                    (int(week), role),
                ).fetchall()
                if child_rows:
                    rows = [row_to_dict(r) for r in child_rows]
                else:
                    overall = conn.execute(
                        """
                        SELECT p.id AS project_id, p.code AS project, d.role, d.week, d.qty
                        FROM demands d
                        JOIN projects p ON p.id = d.project_id
                        WHERE p.closed = 0
                          AND UPPER(p.code) = 'OVERALL OFFICINA'
                          AND d.week = ?
                          AND UPPER(d.role) = UPPER(?)
                          AND d.qty > 0
                        LIMIT 1
                        """,
                        (int(week), role),
                    ).fetchone()
                    if overall:
                        rows = [row_to_dict(overall)]
                return self._send_json(rows)
        if path == "/api/analysis":
            with db_conn() as conn:
                return self._send_json(compute_analysis(conn))
        if path == "/api/all":
            with db_conn() as conn:
                current_demand_path = get_current_demand_source_path(conn)
                resources = []
                for r in conn.execute("SELECT * FROM resources ORDER BY name"):
                    d = row_to_dict(r)
                    d["status"] = resource_status(r)
                    d["is_external"] = is_external_resource_row(d)
                    resources.append(d)
                projects = get_projects_with_template(conn)
                demands = [
                    row_to_dict(r)
                    for r in conn.execute(
                        """
                        SELECT d.id, d.project_id, p.code AS project, d.week, d.role, d.qty
                        FROM demands d
                        JOIN projects p ON p.id = d.project_id
                        ORDER BY d.week, p.code, d.role
                        """
                    )
                ]
                allocations = []
                for r in conn.execute(
                    """
                    SELECT a.*, p.code AS project, rs.name AS resource
                    FROM allocations a
                    JOIN projects p ON p.id = a.project_id
                    JOIN resources rs ON rs.id = a.resource_id
                    ORDER BY rs.name, a.week_from, p.code
                    """
                ):
                    d = row_to_dict(r)
                    d["warnings"] = allocation_warnings(conn, r)
                    d["warning_segments"] = allocation_warning_segments(conn, r)
                    allocations.append(d)
                unavailability = [
                    row_to_dict(r)
                    for r in conn.execute(
                        """
                        SELECT u.*, rs.name AS resource
                        FROM unavailability u
                        JOIN resources rs ON rs.id = u.resource_id
                        ORDER BY rs.name, u.week_from, u.week_to
                        """
                    )
                ]
                report_weekly = compute_report_weekly(conn)
                report_role_weekly = compute_report_role_weekly(conn)
                return self._send_json(
                    {
                        "resources": resources,
                        "projects": projects,
                        "demands": demands,
                        "allocations": allocations,
                        "unavailability": unavailability,
                        "report_weekly": report_weekly,
                        "report_role_weekly": report_role_weekly,
                        "standard_roles": all_standard_roles(),
                        "bootstrap_path": str(BOOTSTRAP_XLSX),
                        "default_demand_path": current_demand_path,
                        "current_demand_path": current_demand_path,
                        "default_resources_path": str(DEFAULT_RESOURCES_XLSX),
                        "default_plan_path": str(DEFAULT_PLAN_XLSX),
                    }
                )
        if path.startswith("/api/project-template/"):
            pid = int(path.split("/")[-1])
            with db_conn() as conn:
                p = conn.execute(
                    """
                    SELECT p.id, p.code, t.week_from, t.week_to
                    FROM projects p
                    LEFT JOIN project_templates t ON t.project_id=p.id
                    WHERE p.id=?
                    """,
                    (pid,),
                ).fetchone()
                if not p:
                    return self._send_json({"error": "Commessa non trovata"}, HTTPStatus.NOT_FOUND)
                week_from = int(p["week_from"] or 1)
                week_to = int(p["week_to"] or 52)
                rows = []
                for w in range(week_from, week_to + 1):
                    for role in all_standard_roles():
                        qrow = conn.execute(
                            "SELECT qty FROM demands WHERE project_id=? AND week=? AND role=?",
                            (pid, w, role),
                        ).fetchone()
                        rows.append(
                            {
                                "project_id": pid,
                                "project": p["code"],
                                "week": w,
                                "role": role,
                                "qty": int(qrow["qty"]) if qrow else 0,
                            }
                        )
                return self._send_json(
                    {
                        "project_id": pid,
                        "project": p["code"],
                        "week_from": week_from,
                        "week_to": week_to,
                        "rows": rows,
                    }
                )

        if path == "/" or path == "":
            self.path = "/index.html"
            return super().do_GET()
        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        body = self._read_json()
        if path == "/api/log":
            actor = normalize_text(body.get("actor")) or "user"
            action = normalize_text(body.get("action")) or "Azione"
            detail = normalize_text(body.get("detail")) or ""
            with db_conn() as conn:
                log_event(conn, actor, action, detail)
            return self._send_json({"ok": True}, HTTPStatus.CREATED)
        if path == "/api/backup/create":
            try:
                info = create_db_backup(body.get("reason") or "manual")
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Salvataggio fallito: {ex}"}, HTTPStatus.BAD_REQUEST)
            return self._send_json({"ok": True, "backup": info}, HTTPStatus.CREATED)
        if path == "/api/backup/restore-latest":
            try:
                info = restore_latest_backup()
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Ripristino fallito: {ex}"}, HTTPStatus.BAD_REQUEST)
            with db_conn() as conn:
                log_event(conn, "system", "Ripristino ultimo", "Ripristino da backup")
            return self._send_json({"ok": True, "backup": info}, HTTPStatus.CREATED)
        if path == "/api/import-excel":
            src = normalize_text(body.get("path")) or str(BOOTSTRAP_XLSX)
            replace_all = bool(body.get("replace_all", True))
            try:
                stats = import_from_xlsx(src, replace_all=replace_all)
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Import Excel fallito: {ex}"}, HTTPStatus.BAD_REQUEST)
            return self._send_json({"ok": True, "stats": stats}, HTTPStatus.CREATED)
        if path == "/api/import-resource-planning":
            demand_src = normalize_text(body.get("demand_path")) or str(DEFAULT_DEMAND_XLSX)
            resources_src = normalize_text(body.get("resources_path")) or str(DEFAULT_RESOURCES_XLSX)
            replace_all = bool(body.get("replace_all", True))
            try:
                stats = import_from_resource_planning_files(demand_src, resources_src, replace_all=replace_all)
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Import planning fallito: {ex}"}, HTTPStatus.BAD_REQUEST)
            with db_conn() as conn:
                set_meta_value(conn, CURRENT_DEMAND_META_KEY, demand_src)
                log_event(conn, "system", "Import iniziale", f"{demand_src} | {resources_src}")
                conn.commit()
            return self._send_json({"ok": True, "stats": stats}, HTTPStatus.CREATED)
        if path == "/api/import-demands-only":
            demand_src = normalize_text(body.get("demand_path")) or str(DEFAULT_DEMAND_XLSX)
            replace_existing = bool(body.get("replace_existing_demands", True))
            role_mapping = body.get("role_mapping") or {}
            create_roles = body.get("create_roles") or []
            try:
                analysis = analyze_demands_import(demand_src)
                backup = create_db_backup("pre-import-demands")
                stats = import_demands_only_from_planning_file(
                    demand_src,
                    replace_existing_demands=replace_existing,
                    role_mapping=role_mapping,
                    create_roles=create_roles,
                )
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Import solo fabbisogno fallito: {ex}"}, HTTPStatus.BAD_REQUEST)
            with db_conn() as conn:
                set_meta_value(conn, CURRENT_DEMAND_META_KEY, demand_src)
                log_event(conn, "system", "Import fabbisogno", f"{demand_src}")
                conn.commit()
            return self._send_json({"ok": True, "stats": stats, "backup": backup, "analysis": analysis}, HTTPStatus.CREATED)
        if path == "/api/import-demands-preview":
            demand_src = normalize_text(body.get("demand_path")) or str(DEFAULT_DEMAND_XLSX)
            try:
                analysis = analyze_demands_import(demand_src)
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Analisi import fabbisogno fallita: {ex}"}, HTTPStatus.BAD_REQUEST)
            return self._send_json({"ok": True, "analysis": analysis}, HTTPStatus.OK)
        if path == "/api/allocations/unassign":
            resource_id = body.get("resource_id")
            week_from = parse_week(body.get("week_from"))
            week_to = parse_week(body.get("week_to"))
            if resource_id is None or week_from is None or week_to is None:
                return self._send_json({"error": "Compila risorsa e settimane."}, HTTPStatus.BAD_REQUEST)
            if week_to < week_from:
                week_from, week_to = week_to, week_from
            with db_conn() as conn:
                rows = conn.execute(
                    """
                    SELECT id FROM allocations
                    WHERE resource_id=? AND week_from <= ? AND week_to >= ?
                    """,
                    (int(resource_id), int(week_to), int(week_from)),
                ).fetchall()
                ids = [r["id"] for r in rows]
                if ids:
                    placeholders = ",".join("?" for _ in ids)
                    conn.execute(f"DELETE FROM allocations WHERE id IN ({placeholders})", tuple(ids))
                log_event(conn, "system", "Allocazione svincolata", f"Risorsa {resource_id} | W{week_from}-W{week_to} | rimosse {len(ids)} allocazioni")
                conn.commit()
            return self._send_json({"ok": True, "deleted": len(ids)}, HTTPStatus.CREATED)
        if path == "/api/allocations/unassign-scope":
            project_id = body.get("project_id")
            role = canonicalize_role(body.get("role"))
            week_from = parse_week(body.get("week_from"))
            week_to = parse_week(body.get("week_to"))
            ext_only = bool(body.get("ext_only", False))
            if project_id is None or not role or week_from is None or week_to is None:
                return self._send_json({"error": "Compila commessa, mansione e settimane."}, HTTPStatus.BAD_REQUEST)
            if week_to < week_from:
                week_from, week_to = week_to, week_from
            with db_conn() as conn:
                query = """
                    SELECT a.id
                    FROM allocations a
                    JOIN resources r ON r.id = a.resource_id
                    WHERE a.project_id=?
                      AND UPPER(a.role)=UPPER(?)
                      AND a.week_from <= ?
                      AND a.week_to >= ?
                """
                params = [int(project_id), role, int(week_to), int(week_from)]
                if ext_only:
                    query += " AND (UPPER(COALESCE(r.employer,''))='EXT' OR UPPER(COALESCE(r.name,'')) LIKE '%-EXT%')"
                rows = conn.execute(query, tuple(params)).fetchall()
                stats = {"deleted": 0, "updated": 0, "inserted": 0}
                for rr in rows:
                    out = trim_allocation_range(conn, rr["id"], int(week_from), int(week_to))
                    stats["deleted"] += out["deleted"]
                    stats["updated"] += out["updated"]
                    stats["inserted"] += out["inserted"]
                log_event(
                    conn,
                    "system",
                    "Svincolo per ambito",
                    f"Commessa {project_id} | {role} | W{week_from}-W{week_to} | ext_only={int(ext_only)} | del {stats['deleted']} upd {stats['updated']} ins {stats['inserted']}",
                )
                conn.commit()
            return self._send_json({"ok": True, **stats}, HTTPStatus.CREATED)
        if path == "/api/open-demand-source":
            demand_src = normalize_text(body.get("demand_path"))
            if not demand_src:
                with db_conn() as conn:
                    demand_src = get_current_demand_source_path(conn)
            src = Path(demand_src)
            if not src.exists():
                return self._send_json({"error": f"File non trovato: {src}"}, HTTPStatus.BAD_REQUEST)
            try:
                os.startfile(str(src))
            except Exception as ex:
                return self._send_json({"error": f"Impossibile aprire il file: {ex}"}, HTTPStatus.BAD_REQUEST)
            return self._send_json({"ok": True}, HTTPStatus.CREATED)
        if path == "/api/analyze-foglio2":
            plan_src = normalize_text(body.get("plan_path")) or str(DEFAULT_PLAN_XLSX)
            try:
                result = analyze_foglio2_plan(plan_src)
            except FileNotFoundError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": f"Analisi Foglio2 fallita: {ex}"}, HTTPStatus.BAD_REQUEST)
            return self._send_json({"ok": True, "analysis": result}, HTTPStatus.CREATED)
        if path == "/api/resources":
            name = normalize_text(body.get("name"))
            role1 = canonicalize_role(body.get("role1"))
            role2 = canonicalize_role(body.get("role2"))
            if not name:
                return self._send_json({"error": "Il nome risorsa e' obbligatorio."}, HTTPStatus.BAD_REQUEST)
            if not role1 and not role2:
                return self._send_json({"error": "Inserisci almeno una mansione per la risorsa."}, HTTPStatus.BAD_REQUEST)
            try:
                with db_conn() as conn:
                    cur = conn.execute(
                        """
                        INSERT INTO resources(
                            name, employee_code, role1, role2, hire_date, end_date, birth_date,
                            phone, email, residence_city, employer, hire_level, base_location, level, contract_type,
                            pb_overtime_hourly, pb_day_office, pb_day_site, pb_fixed_plus,
                            glob_hour_office, glob_hour_site, glob_fixed_plus,
                            doc_type, doc_number, doc_expiry,
                            certifications, note, no_travel, cost
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            name,
                            normalize_text(body.get("employee_code")),
                            role1,
                            role2,
                            normalize_text(body.get("hire_date")),
                            normalize_text(body.get("end_date")),
                            normalize_text(body.get("birth_date")),
                            normalize_text(body.get("phone")),
                            normalize_text(body.get("email")),
                            normalize_text(body.get("residence_city")),
                            normalize_text(body.get("employer")) or "SPEC",
                            normalize_text(body.get("hire_level")),
                            normalize_text(body.get("base_location")),
                            normalize_text(body.get("level")),
                            normalize_text(body.get("contract_type")),
                            parse_decimal(body.get("pb_overtime_hourly"), 0),
                            parse_decimal(body.get("pb_day_office"), 0),
                            parse_decimal(body.get("pb_day_site"), 0),
                            parse_decimal(body.get("pb_fixed_plus"), 0),
                            parse_decimal(body.get("glob_hour_office"), 0),
                            parse_decimal(body.get("glob_hour_site"), 0),
                            parse_decimal(body.get("glob_fixed_plus"), 0),
                            normalize_text(body.get("doc_type")),
                            normalize_text(body.get("doc_number")),
                            normalize_text(body.get("doc_expiry")),
                            normalize_text(body.get("certifications")),
                            normalize_text(body.get("note")),
                            int(bool(body.get("no_travel"))),
                            parse_decimal(body.get("cost"), 0),
                        ),
                    )
                    conn.commit()
                    return self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
            except ValueError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except sqlite3.IntegrityError:
                return self._send_json({"error": "Risorsa gia esistente: usa un nome diverso."}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
        if path == "/api/resources/bulk":
            rows = body.get("rows", [])
            if not isinstance(rows, list):
                return self._send_json({"error": "rows deve essere una lista"}, HTTPStatus.BAD_REQUEST)
            inserted = 0
            try:
                with db_conn() as conn:
                    for r in rows:
                        name = normalize_text(r.get("name"))
                        if not name:
                            continue
                        conn.execute(
                        """
                        INSERT INTO resources(
                            name, employee_code, role1, role2, hire_date, end_date, birth_date,
                            phone, email, residence_city, employer, hire_level, base_location, level, contract_type,
                            pb_overtime_hourly, pb_day_office, pb_day_site, pb_fixed_plus,
                            glob_hour_office, glob_hour_site, glob_fixed_plus,
                            doc_type, doc_number, doc_expiry,
                            certifications, note, no_travel, cost
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(name) DO UPDATE SET
                            employee_code=excluded.employee_code,
                            role1=excluded.role1,
                            role2=excluded.role2,
                            hire_date=excluded.hire_date,
                            end_date=excluded.end_date,
                            birth_date=excluded.birth_date,
                            phone=excluded.phone,
                            email=excluded.email,
                            residence_city=excluded.residence_city,
                            employer=excluded.employer,
                            hire_level=excluded.hire_level,
                            base_location=excluded.base_location,
                            level=excluded.level,
                            contract_type=excluded.contract_type,
                            pb_overtime_hourly=excluded.pb_overtime_hourly,
                            pb_day_office=excluded.pb_day_office,
                            pb_day_site=excluded.pb_day_site,
                            pb_fixed_plus=excluded.pb_fixed_plus,
                            glob_hour_office=excluded.glob_hour_office,
                            glob_hour_site=excluded.glob_hour_site,
                            glob_fixed_plus=excluded.glob_fixed_plus,
                            doc_type=excluded.doc_type,
                            doc_number=excluded.doc_number,
                            doc_expiry=excluded.doc_expiry,
                            certifications=excluded.certifications,
                            note=excluded.note,
                            no_travel=excluded.no_travel,
                            cost=excluded.cost
                        """,
                        (
                            name,
                            normalize_text(r.get("employee_code")),
                            canonicalize_role(r.get("role1")),
                            canonicalize_role(r.get("role2")),
                            normalize_text(r.get("hire_date")),
                            normalize_text(r.get("end_date")),
                            normalize_text(r.get("birth_date")),
                            normalize_text(r.get("phone")),
                            normalize_text(r.get("email")),
                            normalize_text(r.get("residence_city")),
                            normalize_text(r.get("employer")) or "SPEC",
                            normalize_text(r.get("hire_level")),
                            normalize_text(r.get("base_location")),
                            normalize_text(r.get("level")),
                            normalize_text(r.get("contract_type")),
                            parse_decimal(r.get("pb_overtime_hourly"), 0),
                            parse_decimal(r.get("pb_day_office"), 0),
                            parse_decimal(r.get("pb_day_site"), 0),
                            parse_decimal(r.get("pb_fixed_plus"), 0),
                            parse_decimal(r.get("glob_hour_office"), 0),
                            parse_decimal(r.get("glob_hour_site"), 0),
                            parse_decimal(r.get("glob_fixed_plus"), 0),
                            normalize_text(r.get("doc_type")),
                            normalize_text(r.get("doc_number")),
                            normalize_text(r.get("doc_expiry")),
                            normalize_text(r.get("certifications")),
                            normalize_text(r.get("note")),
                            int(bool(r.get("no_travel"))),
                            parse_decimal(r.get("cost"), 0),
                        ),
                        )
                        inserted += 1
                    conn.commit()
            except ValueError as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            except sqlite3.IntegrityError:
                return self._send_json({"error": "Conflitto nomi risorse: una o piu risorse esistono gia."}, HTTPStatus.BAD_REQUEST)
            except Exception as ex:
                return self._send_json({"error": str(ex)}, HTTPStatus.BAD_REQUEST)
            return self._send_json({"ok": True, "rows": inserted}, HTTPStatus.CREATED)
        if path == "/api/projects":
            with db_conn() as conn:
                ptype = normalize_text(body.get("type")) or "SITE"
                rollup = int(bool(body.get("workshop_rollup")))
                if ptype == "WS" and "workshop_rollup" not in body:
                    rollup = 1
                cur = conn.execute(
                    """
                    INSERT INTO projects(code, activity, type, closed, workshop_rollup)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (
                        normalize_text(body.get("code")),
                        normalize_text(body.get("activity")) or normalize_text(body.get("code")),
                        ptype,
                        int(bool(body.get("closed"))),
                        rollup,
                    ),
                )
                pid = cur.lastrowid
                week_from = parse_week(body.get("week_from"))
                week_to = parse_week(body.get("week_to"))
                create_template = bool(body.get("create_template"))
                if week_from and week_to:
                    ensure_project_template(conn, pid, week_from, week_to, True)
                    if create_template:
                        generate_demands_template(conn, pid, week_from, week_to, replace_existing=False)
                conn.commit()
                return self._send_json({"id": pid}, HTTPStatus.CREATED)
        if path == "/api/demands":
            with db_conn() as conn:
                project_id = body.get("project_id")
                if project_id is None or str(project_id).strip() == "":
                    pcode = normalize_text(body.get("project"))
                    if not pcode:
                        return self._send_json({"error": "project_id o project obbligatorio"}, HTTPStatus.BAD_REQUEST)
                    prow = conn.execute("SELECT id FROM projects WHERE code=?", (pcode,)).fetchone()
                    if not prow:
                        ptype = infer_project_type(pcode)
                        conn.execute(
                            "INSERT INTO projects(code, activity, type, workshop_rollup) VALUES (?, ?, ?, ?)",
                            (pcode, pcode, ptype, 1 if ptype == "WS" else 0),
                        )
                        prow = conn.execute("SELECT id FROM projects WHERE code=?", (pcode,)).fetchone()
                    project_id = int(prow["id"])
                else:
                    project_id = int(project_id)
                week = parse_week(body.get("week"))
                role = canonicalize_role(body.get("role"))
                qty = int(body.get("qty") or 0)
                if week is None:
                    return self._send_json({"error": "Week non valida"}, HTTPStatus.BAD_REQUEST)
                conn.execute(
                    "INSERT OR REPLACE INTO demands(project_id, week, role, qty) VALUES (?, ?, ?, ?)",
                    (project_id, week, role, qty),
                )
                conn.commit()
                return self._send_json({"ok": True}, HTTPStatus.CREATED)
        if path == "/api/demands/bulk":
            rows = body.get("rows", [])
            if not isinstance(rows, list):
                return self._send_json({"error": "rows deve essere una lista"}, HTTPStatus.BAD_REQUEST)
            updated = 0
            skipped = 0
            with db_conn() as conn:
                for rr in rows:
                    pid = rr.get("project_id")
                    if pid is None:
                        pcode = normalize_text(rr.get("project"))
                        if pcode:
                            prow = conn.execute("SELECT id FROM projects WHERE code=?", (pcode,)).fetchone()
                            if not prow:
                                conn.execute(
                                    "INSERT INTO projects(code, activity, type, workshop_rollup) VALUES (?, ?, 'SITE', 0)",
                                    (pcode, pcode),
                                )
                                prow = conn.execute("SELECT id FROM projects WHERE code=?", (pcode,)).fetchone()
                            pid = prow["id"]
                    if pid is None:
                        skipped += 1
                        continue
                    week = parse_week(rr.get("week"))
                    role = canonicalize_role(rr.get("role"))
                    if week is None or not role:
                        skipped += 1
                        continue
                    try:
                        qty = int(float(rr.get("qty", 0)))
                    except Exception:
                        qty = 0
                    conn.execute(
                        """
                        INSERT OR REPLACE INTO demands(project_id, week, role, qty)
                        VALUES (?, ?, ?, ?)
                        """,
                        (int(pid), int(week), role, qty),
                    )
                    updated += 1
                conn.commit()
            return self._send_json({"ok": True, "updated": updated, "skipped": skipped}, HTTPStatus.CREATED)
        if path == "/api/unavailability":
            resource_id = int(body.get("resource_id") or 0)
            week_from = parse_week(body.get("week_from"))
            week_to = parse_week(body.get("week_to"))
            if resource_id <= 0 or week_from is None or week_to is None:
                return self._send_json({"error": "Risorsa e settimane sono obbligatorie"}, HTTPStatus.BAD_REQUEST)
            if week_to < week_from:
                week_from, week_to = week_to, week_from
            with db_conn() as conn:
                cur = conn.execute(
                    """
                    INSERT INTO unavailability(resource_id, week_from, week_to, reason)
                    VALUES (?, ?, ?, ?)
                    """,
                    (resource_id, int(week_from), int(week_to), normalize_text(body.get("reason")) or "INDISP"),
                )
                conn.commit()
                return self._send_json({"id": cur.lastrowid}, HTTPStatus.CREATED)
        if path.startswith("/api/projects/") and path.endswith("/template"):
            # POST /api/projects/{id}/template
            parts = [p for p in path.split("/") if p]
            if len(parts) >= 3:
                pid = int(parts[2])
            else:
                return self._send_json({"error": "Path non valido"}, HTTPStatus.BAD_REQUEST)
            week_from = parse_week(body.get("week_from"))
            week_to = parse_week(body.get("week_to"))
            if week_from is None or week_to is None:
                return self._send_json({"error": "week_from/week_to obbligatori"}, HTTPStatus.BAD_REQUEST)
            replace_existing = bool(body.get("replace_existing"))
            with db_conn() as conn:
                ensure_project_template(conn, pid, week_from, week_to, True)
                generate_demands_template(conn, pid, week_from, week_to, replace_existing=replace_existing)
                conn.commit()
            return self._send_json({"ok": True}, HTTPStatus.CREATED)
        if path == "/api/allocations":
            with db_conn() as conn:
                project_id = int(body.get("project_id"))
                resource_id = int(body.get("resource_id"))
                role = canonicalize_role(body.get("role"))
                week_from = int(body.get("week_from"))
                week_to = int(body.get("week_to"))
                weight = float(body.get("weight", 1) or 1)
                if weight <= 0:
                    weight = 1.0
                if week_to < week_from:
                    week_from, week_to = week_to, week_from

                overlaps = find_allocation_overlaps(conn, resource_id, week_from, week_to)

                cur = conn.execute(
                    """
                    INSERT INTO allocations(project_id, resource_id, role, week_from, week_to, weight, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        project_id,
                        resource_id,
                        role,
                        week_from,
                        week_to,
                        weight,
                        normalize_text(body.get("note")),
                    ),
                )
                aid = cur.lastrowid
                conn.commit()
                row = conn.execute(
                    """
                    SELECT a.*, p.code AS project, rs.name AS resource
                    FROM allocations a
                    JOIN projects p ON p.id = a.project_id
                    JOIN resources rs ON rs.id = a.resource_id
                    WHERE a.id=?
                    """,
                    (aid,),
                ).fetchone()
                log_event(conn, "system", "Allocazione creata", f"{row['resource']} -> {row['project']} | {row['role']} | W{row['week_from']}-W{row['week_to']}")
                d = row_to_dict(row)
                d["warnings"] = allocation_warnings(conn, row)
                d["warning_segments"] = allocation_warning_segments(conn, row)
                if overlaps:
                    if "SOVRAPPOSIZIONE" not in d["warnings"]:
                        d["warnings"].append("SOVRAPPOSIZIONE")
                    d["overlaps"] = [row_to_dict(r) for r in overlaps]
                return self._send_json(d, HTTPStatus.CREATED)
        return self._send_json({"error": "Endpoint non trovato"}, HTTPStatus.NOT_FOUND)

    def do_PUT(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        body = self._read_json()
        if path.startswith("/api/resources/"):
            rid = int(path.split("/")[-1])
            name = normalize_text(body.get("name"))
            role1 = canonicalize_role(body.get("role1"))
            role2 = canonicalize_role(body.get("role2"))
            if not name:
                return self._send_json({"error": "Il nome risorsa e' obbligatorio."}, HTTPStatus.BAD_REQUEST)
            if not role1 and not role2:
                return self._send_json({"error": "Inserisci almeno una mansione per la risorsa."}, HTTPStatus.BAD_REQUEST)
            with db_conn() as conn:
                conn.execute(
                    """
                    UPDATE resources
                    SET name=?, employee_code=?, role1=?, role2=?, hire_date=?, end_date=?, birth_date=?, phone=?, email=?,
                        residence_city=?, employer=?, hire_level=?, base_location=?, level=?, contract_type=?,
                        pb_overtime_hourly=?, pb_day_office=?, pb_day_site=?, pb_fixed_plus=?,
                        glob_hour_office=?, glob_hour_site=?, glob_fixed_plus=?,
                        doc_type=?, doc_number=?, doc_expiry=?,
                        certifications=?, note=?, no_travel=?, cost=?
                    WHERE id=?
                    """,
                    (
                        name,
                        normalize_text(body.get("employee_code")),
                        role1,
                        role2,
                        normalize_text(body.get("hire_date")),
                        normalize_text(body.get("end_date")),
                        normalize_text(body.get("birth_date")),
                        normalize_text(body.get("phone")),
                        normalize_text(body.get("email")),
                        normalize_text(body.get("residence_city")),
                        normalize_text(body.get("employer")) or "SPEC",
                        normalize_text(body.get("hire_level")),
                        normalize_text(body.get("base_location")),
                        normalize_text(body.get("level")),
                        normalize_text(body.get("contract_type")),
                        parse_decimal(body.get("pb_overtime_hourly"), 0),
                        parse_decimal(body.get("pb_day_office"), 0),
                        parse_decimal(body.get("pb_day_site"), 0),
                        parse_decimal(body.get("pb_fixed_plus"), 0),
                        parse_decimal(body.get("glob_hour_office"), 0),
                        parse_decimal(body.get("glob_hour_site"), 0),
                        parse_decimal(body.get("glob_fixed_plus"), 0),
                        normalize_text(body.get("doc_type")),
                        normalize_text(body.get("doc_number")),
                        normalize_text(body.get("doc_expiry")),
                        normalize_text(body.get("certifications")),
                        normalize_text(body.get("note")),
                        int(bool(body.get("no_travel"))),
                        parse_decimal(body.get("cost"), 0),
                        rid,
                    ),
                )
                conn.commit()
                return self._send_json({"ok": True})
        if path.startswith("/api/projects/"):
            pid = int(path.split("/")[-1])
            with db_conn() as conn:
                conn.execute(
                    """
                    UPDATE projects
                    SET code=?, activity=?, type=?, closed=?, workshop_rollup=?
                    WHERE id=?
                    """,
                    (
                        normalize_text(body.get("code")),
                        normalize_text(body.get("activity")),
                        normalize_text(body.get("type")) or "SITE",
                        int(bool(body.get("closed"))),
                        int(bool(body.get("workshop_rollup"))),
                        pid,
                    ),
                )
                wf = parse_week(body.get("week_from"))
                wt = parse_week(body.get("week_to"))
                if wf is not None and wt is not None:
                    ensure_project_template(conn, pid, wf, wt, True)
                conn.commit()
                return self._send_json({"ok": True})
        return self._send_json({"error": "Endpoint non trovato"}, HTTPStatus.NOT_FOUND)

    def do_PATCH(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        body = self._read_json()
        if path.startswith("/api/allocations/"):
            aid = int(path.split("/")[-1])
            with db_conn() as conn:
                current = conn.execute("SELECT * FROM allocations WHERE id=?", (aid,)).fetchone()
                if not current:
                    return self._send_json({"error": "Allocazione non trovata"}, HTTPStatus.NOT_FOUND)

                project_id = int(body.get("project_id", current["project_id"]))
                resource_id = int(body.get("resource_id", current["resource_id"]))
                role = canonicalize_role(body.get("role", current["role"]))
                week_from = int(body.get("week_from", current["week_from"]))
                week_to = int(body.get("week_to", current["week_to"]))
                weight = float(body.get("weight", current["weight"] if "weight" in current.keys() else 1) or 1)
                if weight <= 0:
                    weight = 1.0
                if week_to < week_from:
                    week_from, week_to = week_to, week_from

                overlaps = find_allocation_overlaps(conn, resource_id, week_from, week_to, current_allocation_id=aid)

                conn.execute(
                    """
                    UPDATE allocations
                    SET project_id=?, resource_id=?, role=?, week_from=?, week_to=?, weight=?, note=?
                    WHERE id=?
                    """,
                    (
                        project_id,
                        resource_id,
                        role,
                        week_from,
                        week_to,
                        weight,
                        normalize_text(body.get("note", current["note"])),
                        aid,
                    ),
                )
                conn.commit()
                row = conn.execute(
                    """
                    SELECT a.*, p.code AS project, rs.name AS resource
                    FROM allocations a
                    JOIN projects p ON p.id = a.project_id
                    JOIN resources rs ON rs.id = a.resource_id
                    WHERE a.id=?
                    """,
                    (aid,),
                ).fetchone()
                log_event(conn, "system", "Allocazione aggiornata", f"{row['resource']} -> {row['project']} | {row['role']} | W{row['week_from']}-W{row['week_to']}")
                d = row_to_dict(row)
                d["warnings"] = allocation_warnings(conn, row)
                d["warning_segments"] = allocation_warning_segments(conn, row)
                if overlaps:
                    if "SOVRAPPOSIZIONE" not in d["warnings"]:
                        d["warnings"].append("SOVRAPPOSIZIONE")
                    d["overlaps"] = [row_to_dict(r) for r in overlaps]
                return self._send_json(d)
        return self._send_json({"error": "Endpoint non trovato"}, HTTPStatus.NOT_FOUND)

    def do_DELETE(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path.startswith("/api/resources/"):
            rid = int(path.split("/")[-1])
            with db_conn() as conn:
                conn.execute("DELETE FROM resources WHERE id=?", (rid,))
                conn.commit()
                return self._send_json({"ok": True})
        if path.startswith("/api/allocations/"):
            aid = int(path.split("/")[-1])
            with db_conn() as conn:
                cur = conn.execute(
                    """
                    SELECT a.*, p.code AS project, rs.name AS resource
                    FROM allocations a
                    JOIN projects p ON p.id = a.project_id
                    JOIN resources rs ON rs.id = a.resource_id
                    WHERE a.id=?
                    """,
                    (aid,),
                ).fetchone()
                conn.execute("DELETE FROM allocations WHERE id=?", (aid,))
                if cur:
                    log_event(conn, "system", "Allocazione rimossa", f"{cur['resource']} -> {cur['project']} | {cur['role']} | W{cur['week_from']}-W{cur['week_to']}")
                conn.commit()
                return self._send_json({"ok": True})
        if path.startswith("/api/demands/"):
            did = int(path.split("/")[-1])
            with db_conn() as conn:
                conn.execute("DELETE FROM demands WHERE id=?", (did,))
                conn.commit()
                return self._send_json({"ok": True})
        if path.startswith("/api/unavailability/"):
            uid = int(path.split("/")[-1])
            with db_conn() as conn:
                conn.execute("DELETE FROM unavailability WHERE id=?", (uid,))
                conn.commit()
                return self._send_json({"ok": True})
        return self._send_json({"error": "Endpoint non trovato"}, HTTPStatus.NOT_FOUND)


def start_server():
    init_db()
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), AppHandler)
    print(f"Planner live su http://127.0.0.1:{PORT}")
    httpd.serve_forever()


if __name__ == "__main__":
    start_server()
