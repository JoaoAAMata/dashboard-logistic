import sqlite3
import hashlib
from datetime import datetime

DB_PATH = "logistics.db"

DEFAULT_PIN = "1234"

STORES = [
    ("M1UT",  "M1UT Store MY 1Utama Blue & One Blue",       "m1ut",           "Lebuh Bandar Utama",                                  "Petaling Jaya",    "Malaysia", 0),
    ("MBAN",  "MBAN Store MY Bangsar Village Brothers",      "mban",           "Jalan Telawi Satu",                                   "Kuala Lumpur",     "Malaysia", 0),
    ("MDVI",  "MDVI Store MY Design Village Outlet",         "mdvi",           "Jalan Cassia Barat 2",                                "Penang",           "Malaysia", 0),
    ("MMAL",  "MMAL Store MY Freeport Malaca Outlet",        "mmal",           "Jalan Kemus/Simpang Ampat",                           "Melaka",           "Malaysia", 0),
    ("MGAR",  "MGAR Store MY Gardens Brothers",              "mgar",           "Lingkaran Syed Putra",                                "Kuala Lumpur",     "Malaysia", 0),
    ("MGEN",  "MGEN Store MY Genting Blue BlueOut",          "mgen_blue",      "Genting Highlands Premium Outlet, Unit 544",          "Genting Highlands","Malaysia", 0),
    ("MGEN",  "MGEN Store MY Genting Highlands Outlet",      "mgen_highlands", "Genting Highlands Premium Outlet, Unit 126",          "Genting Highlands","Malaysia", 0),
    ("MGUR",  "MGUR Store MY Gurney Plaza Brothers",         "mgur",           "Persiaran Gurney",                                    "Penang",           "Malaysia", 0),
    ("MIMA",  "MIMA Store MY Imago Brothers",                "mima",           "Off Coastal Highway",                                 "Kota Kinabalu",    "Malaysia", 0),
    ("MJOH",  "MJOH Store MY Johor Bahru Outlet",            "mjoh_outlet",    "Johor Bahru Premium Outlet",                          "Johor",            "Malaysia", 0),
    ("MJOH",  "MJOH Store MY Johor Bahru Blue BlueOut",      "mjoh_blue",      "Johor Bahru Premium Outlet",                          "Johor",            "Malaysia", 0),
    ("MJOH",  "MJOH Store MY Johor Bahru Classic Outlet",    "mjoh_classic",   "Johor Bahru Premium Outlet",                          "Johor",            "Malaysia", 0),
    ("MLAL",  "MLAL Store MY Lalaport Brothers",             "mlal_brothers",  "Jln Hang Tuah",                                       "Kuala Lumpur",     "Malaysia", 0),
    ("MLAL",  "MLAL Store MY Lalaport Blue Blue",            "mlal_blue",      "Mitsui Shopping Park LaLaport, G-15D",                "Kuala Lumpur",     "Malaysia", 0),
    ("MLAL",  "MLAL Store MY Lalaport One One",              "mlal_one",       "Mitsui Shopping Park LaLaport, G-53",                 "Kuala Lumpur",     "Malaysia", 0),
    ("MMIT",  "MMIT Store MY Mitsui Outlet",                 "mmit_outlet",    "Persiaran Komersial",                                 "Sepang",           "Malaysia", 0),
    ("MMIT",  "MMIT Store MY Mitsui Blue BlueOut",           "mmit_blue",      "Persiaran Komersial",                                 "Sepang",           "Malaysia", 0),
    ("MMIT",  "MMIT Store MY Mitsui Classic Outlet",         "mmit_classic",   "Persiaran Komersial",                                 "Sepang",           "Malaysia", 0),
    ("MMIT",  "MMIT Store MY Mitsui Women Outlet",           "mmit_women",     "Persiaran Komersial",                                 "Sepang",           "Malaysia", 0),
    ("MPAV",  "MPAV Store MY Pavilion Brothers",             "mpav",           "Jalan Bukit Bintang",                                 "Kuala Lumpur",     "Malaysia", 0),
    ("MQUE",  "MQUE Store MY Queensbay Brothers",            "mque",           "Persiaran Bayan Indah",                               "Penang",           "Malaysia", 0),
    ("MSUN",  "MSUN Store MY Sunway Brothers",               "msun",           "Sunway Pyramid, Unit: G1.06",                         "Petaling Jaya",    "Malaysia", 0),
    ("MKLC",  "MKLC Store MY Suria KLCC Brothers",           "mklc",           "Suria KLCC, Lot No. 231B/232, Level 2",               "Kuala Lumpur",     "Malaysia", 0),
    ("MTRX",  "MTRX Store MY TRX Brothers",                  "mtrx",           "Persiaran TRX",                                       "Kuala Lumpur",     "Malaysia", 0),
    ("DHL",   "SACOOR DHL Warehouse",                        "dhl_sacoor",     "50, Persiaran Perusahaan, Kawasan Miel, Shah Alam",    "Selangor",         "Malaysia", 0),
    # Admin account — logistics team
    ("PVOFF", "SACOOR HQ - Logistics",                       "logistics",      "Jalan Bukit Bintang",                                 "Kuala Lumpur",     "Malaysia", 1),
]


def hash_pin(pin: str) -> str:
    return hashlib.sha256(pin.encode()).hexdigest()


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS stores (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            store_code  TEXT NOT NULL,
            store_name  TEXT NOT NULL UNIQUE,
            username    TEXT NOT NULL UNIQUE,
            pin_hash    TEXT NOT NULL,
            address     TEXT,
            city        TEXT,
            country     TEXT,
            is_admin    INTEGER DEFAULT 0,
            is_active   INTEGER DEFAULT 1
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS transfers (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            collect_no      TEXT NOT NULL UNIQUE,
            form_type       TEXT DEFAULT 'commercial',
            from_store_id   INTEGER NOT NULL,
            to_store_id     INTEGER NOT NULL,
            collection_date TEXT NOT NULL,
            delivery_date   TEXT NOT NULL,
            total_pcs       INTEGER DEFAULT 0,
            total_ctn       INTEGER DEFAULT 0,
            total_rln       INTEGER DEFAULT 0,
            status          TEXT DEFAULT 'pending',
            rejection_reason TEXT,
            submitted_at    TEXT NOT NULL,
            updated_at      TEXT,
            FOREIGN KEY (from_store_id) REFERENCES stores(id),
            FOREIGN KEY (to_store_id)   REFERENCES stores(id)
        )
    """)

    # Add form_type column if upgrading from older DB
    try:
        c.execute("ALTER TABLE transfers ADD COLUMN form_type TEXT DEFAULT 'commercial'")
    except Exception:
        pass

    # Add total_rln column if upgrading from older DB
    try:
        c.execute("ALTER TABLE transfers ADD COLUMN total_rln INTEGER DEFAULT 0")
    except Exception:
        pass

    c.execute("""
        CREATE TABLE IF NOT EXISTS transfer_lines (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            transfer_id INTEGER NOT NULL,
            tg_number   TEXT,
            description TEXT DEFAULT 'Stock Rotation by Email',
            uom         TEXT DEFAULT 'Pcs',
            qty         INTEGER NOT NULL,
            picture_ref TEXT,
            FOREIGN KEY (transfer_id) REFERENCES transfers(id)
        )
    """)

    # Add picture_ref column if upgrading from older DB
    try:
        c.execute("ALTER TABLE transfer_lines ADD COLUMN picture_ref TEXT")
    except Exception:
        pass

    # Seed stores (skip if already seeded)
    for store_code, store_name, username, address, city, country, is_admin in STORES:
        c.execute(
            "INSERT OR IGNORE INTO stores (store_code, store_name, username, pin_hash, address, city, country, is_admin) VALUES (?,?,?,?,?,?,?,?)",
            (store_code, store_name, username, hash_pin(DEFAULT_PIN), address, city, country, is_admin)
        )

    conn.commit()
    conn.close()


# ── Store queries ─────────────────────────────────────────────────────────────

def get_store_by_username(username: str):
    conn = get_conn()
    row = conn.execute("SELECT * FROM stores WHERE username = ? AND is_active = 1", (username,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_store_by_id(store_id: int):
    conn = get_conn()
    row = conn.execute("SELECT * FROM stores WHERE id = ?", (store_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_stores(exclude_admin=True):
    conn = get_conn()
    if exclude_admin:
        rows = conn.execute("SELECT * FROM stores WHERE is_admin = 0 AND is_active = 1 ORDER BY store_name").fetchall()
    else:
        rows = conn.execute("SELECT * FROM stores WHERE is_active = 1 ORDER BY store_name").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def change_pin(store_id: int, new_pin: str):
    conn = get_conn()
    conn.execute("UPDATE stores SET pin_hash = ? WHERE id = ?", (hash_pin(new_pin), store_id))
    conn.commit()
    conn.close()


# ── Transfer queries ───────────────────────────────────────────────────────────

def _next_collection_no(conn) -> str:
    year = datetime.utcnow().year
    rows = conn.execute(
        "SELECT collect_no FROM transfers WHERE collect_no LIKE ?",
        (f"RET-%-{year}",)
    ).fetchall()
    max_seq = 0
    for row in rows:
        try:
            seq = int(row[0].split("-")[1])
            max_seq = max(max_seq, seq)
        except (IndexError, ValueError):
            pass
    return f"RET-{max_seq + 1:03d}-{year}"


def create_transfer(from_store_id, to_store_id, collection_date, delivery_date,
                    total_pcs, total_ctn, total_rln, lines: list, form_type: str = "commercial"):
    conn = get_conn()
    collect_no = _next_collection_no(conn)

    now = datetime.utcnow().isoformat()
    cur = conn.execute(
        """INSERT INTO transfers (collect_no, form_type, from_store_id, to_store_id,
           collection_date, delivery_date, total_pcs, total_ctn, total_rln, status, submitted_at)
           VALUES (?,?,?,?,?,?,?,?,?,'pending',?)""",
        (collect_no, form_type, from_store_id, to_store_id,
         collection_date, delivery_date, total_pcs, total_ctn, total_rln, now)
    )
    transfer_id = cur.lastrowid

    for line in lines:
        conn.execute(
            """INSERT INTO transfer_lines (transfer_id, tg_number, description, uom, qty, picture_ref)
               VALUES (?,?,?,?,?,?)""",
            (transfer_id,
             line.get("tg_number", ""),
             line.get("description", "Stock Rotation by Email"),
             line.get("uom", "Pcs"),
             line["qty"],
             line.get("picture_ref", ""))
        )

    conn.commit()
    conn.close()
    return transfer_id


def get_transfers_by_store(store_id: int):
    conn = get_conn()
    rows = conn.execute("""
        SELECT t.*, s.store_name as to_store_name
        FROM transfers t
        JOIN stores s ON s.id = t.to_store_id
        WHERE t.from_store_id = ?
        ORDER BY t.submitted_at DESC
    """, (store_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_transfers():
    conn = get_conn()
    rows = conn.execute("""
        SELECT t.*,
               sf.store_name as from_store_name,
               st.store_name as to_store_name
        FROM transfers t
        JOIN stores sf ON sf.id = t.from_store_id
        JOIN stores st ON st.id = t.to_store_id
        ORDER BY t.submitted_at DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_transfer_detail(transfer_id: int):
    conn = get_conn()
    row = conn.execute("""
        SELECT t.*,
               sf.store_name as from_store_name, sf.store_code as from_store_code,
               sf.address as from_address, sf.city as from_city, sf.country as from_country,
               st.store_name as to_store_name, st.store_code as to_store_code,
               st.address as to_address, st.city as to_city, st.country as to_country
        FROM transfers t
        JOIN stores sf ON sf.id = t.from_store_id
        JOIN stores st ON st.id = t.to_store_id
        WHERE t.id = ?
    """, (transfer_id,)).fetchone()

    if not row:
        conn.close()
        return None

    transfer = dict(row)
    lines = conn.execute(
        "SELECT * FROM transfer_lines WHERE transfer_id = ? ORDER BY id", (transfer_id,)
    ).fetchall()
    transfer["lines"] = [dict(l) for l in lines]
    conn.close()
    return transfer


def delete_transfer(transfer_id: int):
    conn = get_conn()
    conn.execute("DELETE FROM transfer_lines WHERE transfer_id = ?", (transfer_id,))
    conn.execute("DELETE FROM transfers WHERE id = ?", (transfer_id,))
    conn.commit()
    conn.close()


def update_transfer_status(transfer_id: int, status: str, reason: str = None):
    conn = get_conn()
    conn.execute(
        "UPDATE transfers SET status = ?, rejection_reason = ?, updated_at = ? WHERE id = ?",
        (status, reason, datetime.utcnow().isoformat(), transfer_id)
    )
    conn.commit()
    conn.close()


def update_transfer(transfer_id: int, to_store_id: int, collection_date: str,
                    delivery_date: str, total_ctn: int, total_rln: int, lines: list):
    conn = get_conn()
    total_pcs = sum(l["qty"] for l in lines)
    now = datetime.utcnow().isoformat()

    conn.execute("""
        UPDATE transfers SET to_store_id=?, collection_date=?, delivery_date=?,
        total_pcs=?, total_ctn=?, total_rln=?, updated_at=? WHERE id=?
    """, (to_store_id, collection_date, delivery_date, total_pcs, total_ctn, total_rln, now, transfer_id))

    conn.execute("DELETE FROM transfer_lines WHERE transfer_id = ?", (transfer_id,))
    for line in lines:
        conn.execute(
            """INSERT INTO transfer_lines (transfer_id, tg_number, description, uom, qty, picture_ref)
               VALUES (?,?,?,?,?,?)""",
            (transfer_id,
             line.get("tg_number", ""),
             line.get("description", "Stock Rotation by Email"),
             line.get("uom", "Pcs"),
             line["qty"],
             line.get("picture_ref", ""))
        )

    conn.commit()
    conn.close()
