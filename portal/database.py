import sqlite3
import hashlib
from datetime import datetime

import os as _os
DB_PATH = _os.environ.get("DB_PATH", "logistics.db")
# Ensure the parent directory exists (needed when using a Railway Volume path like /data/logistics.db)
_db_dir = _os.path.dirname(DB_PATH)
if _db_dir:
    _os.makedirs(_db_dir, exist_ok=True)

DEFAULT_PIN = "1234"

# Tuple format: (store_code, store_name, username, address, city, country, is_admin)
STORES = [
    # ── Warehouse ─────────────────────────────────────────────────────────────
    ("DHL-SACOOR", "SACOOR DHL Warehouse",                   "dhl_sacoor",
     "50, Persiaran Perusahaan, Kawasan Miel, 40300 Shah Alam",               "Selangor",      "Malaysia", 0),

    # ── Stores ────────────────────────────────────────────────────────────────
    ("IMAG",    "IMAG - Imago Brothers",                     "imag",
     "Imago Shopping Mall, Unit G-25, KK Times Square Phase 2, Off Coastal Highway, 88100 Kota Kinabalu",
                                                                               "Sabah",         "Malaysia", 0),

    ("M1UTAB",  "M1UTAB - 1 Utama Brothers",                "m1utab",
     "1 Utama Shopping Centre, Unit G-129, Old Wing, 1, Lebuh Bandar Utama, Bandar Utama, 47800 Petaling Jaya",
                                                                               "Selangor",      "Malaysia", 0),

    ("MALA",    "MALA - Freeport A'Famosa Outlet",           "mala",
     "Freeport A'Famosa Outlet, Alor Gajah, 78000 Melaka",                    "Malacca",       "Malaysia", 0),

    ("MBGV",    "MBGV - Bangsar Village Brothers",           "mbgv",
     "Bangsar Village 2, Jalan Telawi 1, Bangsar, 59100 Kuala Lumpur",        "Kuala Lumpur",  "Malaysia", 0),

    ("MDSV",    "MDSV - Design Village Outlet",              "mdsv",
     "Design Village Outlet Mall, Bandar Cassia, 14110 Pulau Pinang",         "Penang",        "Malaysia", 0),

    ("MGAR",    "MGAR - Gardens Brothers",                   "mgar",
     "The Boulevard, F-221, 1st Floor, 59, Lingkaran Syed Putra, Mid Valley City, 59200 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MGPP",    "MGPP - Gurney Plaza Brothers",              "mgpp",
     "Gurney Plaza, Georgetown, 10250 Penang",                                 "Penang",        "Malaysia", 0),

    ("MGTI",    "MGTI - Genting Highlands Outlet",           "mgti",
     "Genting Highlands Premium Outlets, Genting Highlands, 69000 Pahang",    "Pahang",        "Malaysia", 0),

    ("MGTIB",   "MGTIB - Genting Blue Outlet",               "mgtib",
     "Genting Highlands Premium Outlets, Suite 544, KM13, Genting Highlands Resorts, 69000 Genting Highlands",
                                                                               "Pahang",        "Malaysia", 0),

    ("MITS",    "MITS - Mitsui Outlet Brothers",             "mits",
     "Mitsui Outlet Park KLIA, Lot G51, Persiaran Komersial, 64000 KLIA Sepang",
                                                                               "Selangor",      "Malaysia", 0),

    ("MITSB",   "MITSB - Mitsui Blue",                       "mitsb",
     "Mitsui Outlet Park KLIA, Unit G26, Ground Floor, Persiaran Komersial, 64000 KLIA Sepang",
                                                                               "Selangor",      "Malaysia", 0),

    ("MITSO",   "MITSO - Mitsui Classic Outlet",             "mitso",
     "Mitsui Outlet Park KLIA, Lot G96, Persiaran Komersial, 64000 KLIA Sepang",
                                                                               "Selangor",      "Malaysia", 0),

    ("MITSP",   "MITSP - Mitsui Women",                      "mitsp",
     "Mitsui Outlet Park KLIA, Unit G84, Ground Floor, Persiaran Komersial, 64000 KLIA Sepang",
                                                                               "Selangor",      "Malaysia", 0),

    ("MITSW",   "MITSW - Mitsui Women Outlet",               "mitsw",
     "Mitsui Outlet Park, Lot G-40, Ground Floor, Persiaran Komersial, 64000 KLIA Sepang",
                                                                               "Selangor",      "Malaysia", 0),

    ("MJOBB",   "MJOBB - Johor Blue",                        "mjobb",
     "Johor Bahru Premium Outlet, Unit 530, Indahpura, Kulaijaya, 81000 Johor",
                                                                               "Johor",         "Malaysia", 0),

    ("MJPO",    "MJPO - Johor Premium Outlets",              "mjpo",
     "Jalan Premium Outlets Indahpura, 81000 Kulai",                          "Johor",         "Malaysia", 0),

    ("MKLC",    "MKLC - Suria KLCC Brothers",                "mklc",
     "Suria KLCC, 231B-232, Persiaran Petronas, Kuala Lumpur City Centre, 50088 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MLAL",    "MLAL - Lalaport Brothers",                  "mlal",
     "Mitsui Lalaport, Unit G-59, 2, Jalan Hang Tuah, Bukit Bintang, 55100 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MLALB",   "MLALB - Lalaport Blue",                     "mlalb",
     "Mitsui Lalaport, Unit G-15D, 2, Jalan Hang Tuah, Bukit Bintang, 55100 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MLALO",   "MLALO - Lalaport One",                      "mlalo",
     "Mitsui Lalaport, Unit G-53, 2, Jalan Hang Tuah, Bukit Bintang, 55100 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MPAV",    "MPAV - Pavilion Brothers",                  "mpav",
     "Pavilion Kuala Lumpur, 168, Jalan Raja Chulan, Bukit Bintang, 55100 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MPAVB",   "MPAVB - Pavilion Blue",                     "mpavb",
     "Pavilion Kuala Lumpur Mall, Lot No. 4.01.03 & 4.01.04, Level 4, 55100 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 0),

    ("MSUN",    "MSUN - Sunway Pyramid Brothers",             "msun",
     "Sunway Pyramid, Unit G1.06, Bandar Sunway, 47500 Petaling Jaya",        "Selangor",      "Malaysia", 0),

    # ── Admin — Logistics HQ ──────────────────────────────────────────────────
    ("PVOFF",   "SACOOR HQ - Logistics",                     "logistics",
     "Pavilion Retail Office Block, Level 9, Unit 9.03.00, 55100 Kuala Lumpur",
                                                                               "Kuala Lumpur",  "Malaysia", 1),
]


# Transporter accounts (seeded separately from stores)
TRANSPORTERS = [
    ("DHL", "DHL Malaysia - Transporter", "dhl_transport",
     "50, Persiaran Perusahaan, Kawasan Miel, 40300 Shah Alam", "Selangor", "Malaysia"),
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

    # Add receipt_note column if upgrading from older DB
    try:
        c.execute("ALTER TABLE transfers ADD COLUMN receipt_note TEXT")
    except Exception:
        pass

    # Add warehouse_date column if upgrading from older DB
    try:
        c.execute("ALTER TABLE transfers ADD COLUMN warehouse_date TEXT")
    except Exception:
        pass

    # Add receipt_date column if upgrading from older DB
    try:
        c.execute("ALTER TABLE transfers ADD COLUMN receipt_date TEXT")
    except Exception:
        pass

    # Add is_transporter column if upgrading from older DB
    try:
        c.execute("ALTER TABLE stores ADD COLUMN is_transporter INTEGER DEFAULT 0")
    except Exception:
        pass

    # Sync stores: update existing by username, insert new ones, deactivate removed ones
    new_usernames = [s[2] for s in STORES]
    # Deactivate any store whose username is no longer in the master list
    placeholders = ",".join("?" * len(new_usernames))
    c.execute(f"UPDATE stores SET is_active = 0 WHERE username NOT IN ({placeholders})", new_usernames)

    for store_code, store_name, username, address, city, country, is_admin in STORES:
        # Try to update an existing row (preserves pin_hash)
        updated = c.execute(
            """UPDATE stores
               SET store_code=?, store_name=?, address=?, city=?, country=?, is_admin=?, is_active=1
               WHERE username=?""",
            (store_code, store_name, address, city, country, is_admin, username)
        ).rowcount
        if not updated:
            # Brand-new store — insert with default PIN
            c.execute(
                """INSERT INTO stores (store_code, store_name, username, pin_hash, address, city, country, is_admin)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (store_code, store_name, username, hash_pin(DEFAULT_PIN), address, city, country, is_admin)
            )

    # Seed transporter accounts
    for store_code, store_name, username, address, city, country in TRANSPORTERS:
        updated = c.execute(
            """UPDATE stores
               SET store_code=?, store_name=?, address=?, city=?, country=?,
                   is_admin=0, is_transporter=1, is_active=1
               WHERE username=?""",
            (store_code, store_name, address, city, country, username)
        ).rowcount
        if not updated:
            c.execute(
                """INSERT INTO stores
                   (store_code, store_name, username, pin_hash, address, city, country, is_admin, is_transporter)
                   VALUES (?,?,?,?,?,?,?,0,1)""",
                (store_code, store_name, username, hash_pin(DEFAULT_PIN), address, city, country)
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


def get_all_transfers_with_lines():
    """Like get_all_transfers() but also attaches transfer_lines to each record."""
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
    transfers = []
    for row in rows:
        t = dict(row)
        lines = conn.execute(
            "SELECT * FROM transfer_lines WHERE transfer_id = ? ORDER BY id",
            (t["id"],)
        ).fetchall()
        t["lines"] = [dict(l) for l in lines]
        transfers.append(t)
    conn.close()
    return transfers


def get_incoming_transfers(store_id: int):
    """Returns approved transfers whose destination is this store."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT t.*, s.store_name as from_store_name
        FROM transfers t
        JOIN stores s ON s.id = t.from_store_id
        WHERE t.to_store_id = ? AND t.status IN ('approved', 'warehouse', 'completed', 'incorrect')
        ORDER BY t.delivery_date ASC
    """, (store_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_receipt_status(transfer_id: int, status: str, note: str = ""):
    """Mark a transfer as completed or incorrect (receipt confirmation by store or logistics)."""
    conn = get_conn()
    receipt_date = datetime.utcnow().strftime("%Y-%m-%d")
    conn.execute(
        "UPDATE transfers SET status = ?, receipt_note = ?, receipt_date = ?, updated_at = ? WHERE id = ?",
        (status, note or "", receipt_date, datetime.utcnow().isoformat(), transfer_id)
    )
    conn.commit()
    conn.close()


def get_transfers_for_transporter():
    """Returns all approved/warehouse transfers for the transporter to act on."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT t.*,
               sf.store_name as from_store_name,
               st.store_name as to_store_name
        FROM transfers t
        JOIN stores sf ON sf.id = t.from_store_id
        JOIN stores st ON st.id = t.to_store_id
        WHERE t.status IN ('approved', 'warehouse', 'completed', 'incorrect')
        ORDER BY t.collection_date ASC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_transporter_status(transfer_id: int, status: str, warehouse_date: str = None):
    """Transporter marks transfer as arrived at warehouse hub."""
    conn = get_conn()
    if warehouse_date:
        conn.execute(
            "UPDATE transfers SET status = ?, warehouse_date = ?, updated_at = ? WHERE id = ?",
            (status, warehouse_date, datetime.utcnow().isoformat(), transfer_id)
        )
    else:
        conn.execute(
            "UPDATE transfers SET status = ?, updated_at = ? WHERE id = ?",
            (status, datetime.utcnow().isoformat(), transfer_id)
        )
    conn.commit()
    conn.close()
