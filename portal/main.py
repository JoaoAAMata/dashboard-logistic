from fastapi import FastAPI, Request, Form, Response, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List, Optional
import secrets
import uuid
import os
import shutil
import database
import pdf_generator
from datetime import date, datetime

# ── Delivery code mapping (Excel code → system store_code) ───────────────────
# Keys are UPPERCASE Excel STORE codes that don't match system codes directly.
DELIVERY_CODE_MAP = {
    "1UTAB":        "M1UTAB",
    "IMAGO":        "IMAG",
    "JOHORCLASSIC": "MJOCL",
    "LALAONE":      "MLALO",
    "MITSC":        "MITSO",
    "MJOB":         "MJPO",
    "MQUEEN":       "MQBB",
    "SUNWAY":       "MSUN",
}

# Excel codes that represent non-store entities — silently ignored
DELIVERY_IGNORE = {"ONLINE", "ONLINE-SEA", "UNIF", "UNIF-MY", "MEMP", "MIOI", "MJOBBLUE"}

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "static", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = FastAPI(title="Sacoor Logistics Portal")
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ── In-memory session store ───────────────────────────────────────────────────
SESSIONS: dict = {}


def get_session(request: Request) -> Optional[dict]:
    sid = request.cookies.get("sid")
    return SESSIONS.get(sid)


def redirect(url: str, **cookie_args):
    r = RedirectResponse(url, status_code=302)
    for k, v in cookie_args.items():
        r.set_cookie(k, v, httponly=True)
    return r


# ── Startup ───────────────────────────────────────────────────────────────────

@app.on_event("startup")
async def startup():
    database.init_db()
    database.auto_archive_old_transfers()


# ── Root ──────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    return RedirectResponse("/home")


# ── Home hub ──────────────────────────────────────────────────────────────────

@app.get("/home", response_class=HTMLResponse)
async def home(request: Request):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    return templates.TemplateResponse("home.html", {"request": request, "session": s})


# ── Auth ──────────────────────────────────────────────────────────────────────

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    s = get_session(request)
    if s:
        return RedirectResponse("/home")
    return templates.TemplateResponse("login.html", {"request": request})


@app.post("/login")
async def login(request: Request, username: str = Form(...), pin: str = Form(...)):
    store = database.get_store_by_username(username.strip().lower())
    if not store or store["pin_hash"] != database.hash_pin(pin):
        return templates.TemplateResponse("login.html", {
            "request": request,
            "error": "Invalid username or PIN. Please try again."
        })
    sid = secrets.token_hex(32)
    SESSIONS[sid] = {
        "store_id":       store["id"],
        "store_name":     store["store_name"],
        "store_code":     store["store_code"],
        "username":       store["username"],
        "is_admin":       bool(store["is_admin"]),
        "is_transporter": bool(store.get("is_transporter", 0)),
    }
    url = "/home"
    r = RedirectResponse(url, status_code=302)
    r.set_cookie("sid", sid, httponly=True)
    return r


@app.get("/logout")
async def logout(request: Request):
    sid = request.cookies.get("sid")
    SESSIONS.pop(sid, None)
    r = RedirectResponse("/login", status_code=302)
    r.delete_cookie("sid")
    return r


# ── Change PIN ────────────────────────────────────────────────────────────────

@app.get("/change-pin", response_class=HTMLResponse)
async def change_pin_page(request: Request):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    return templates.TemplateResponse("change_pin.html", {"request": request, "session": s})


@app.post("/change-pin")
async def change_pin(request: Request, current_pin: str = Form(...),
                     new_pin: str = Form(...), confirm_pin: str = Form(...)):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")

    store = database.get_store_by_id(s["store_id"])
    error = None
    if store["pin_hash"] != database.hash_pin(current_pin):
        error = "Current PIN is incorrect."
    elif len(new_pin) < 4:
        error = "New PIN must be at least 4 characters."
    elif new_pin != confirm_pin:
        error = "PINs do not match."

    if error:
        return templates.TemplateResponse("change_pin.html", {"request": request, "session": s, "error": error})

    database.change_pin(s["store_id"], new_pin)
    return templates.TemplateResponse("change_pin.html", {
        "request": request, "session": s,
        "success": "PIN changed successfully!"
    })


# ── Store portal ──────────────────────────────────────────────────────────────

@app.get("/store", response_class=HTMLResponse)
async def store_dashboard(request: Request):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    if s["is_admin"]:
        return RedirectResponse("/logistics")
    if s.get("is_transporter"):
        return RedirectResponse("/transporter")
    database.auto_archive_old_transfers()
    transfers = database.get_transfers_by_store(s["store_id"])
    incoming  = database.get_incoming_transfers(s["store_id"])
    return templates.TemplateResponse("store_dashboard.html", {
        "request": request, "session": s,
        "transfers": transfers, "incoming": incoming,
    })


@app.get("/store/archived", response_class=HTMLResponse)
async def store_archived_page(request: Request):
    s = get_session(request)
    if not s or s["is_admin"] or s.get("is_transporter"):
        return RedirectResponse("/login")
    archived = database.get_archived_transfers(s["store_id"])
    return templates.TemplateResponse("store_archived.html", {
        "request": request, "session": s, "archived": archived,
    })


@app.get("/store/new", response_class=HTMLResponse)
async def new_transfer_page(request: Request):
    s = get_session(request)
    if not s or s["is_admin"]:
        return RedirectResponse("/login")
    stores = [st for st in database.get_all_stores() if st["id"] != s["store_id"]]
    return templates.TemplateResponse("store_form.html", {
        "request": request, "session": s,
        "stores": stores,
        "today": date.today().isoformat(),
        "error": None,
    })


@app.post("/store/new")
async def submit_transfer(request: Request):
    s = get_session(request)
    if not s or s["is_admin"]:
        return RedirectResponse("/login")

    form = await request.form()
    stores = [st for st in database.get_all_stores() if st["id"] != s["store_id"]]

    to_store_id     = int(form.get("to_store_id", 0))
    collection_date = form.get("collection_date", "")
    delivery_date   = form.get("delivery_date", "")
    total_ctn       = int(form.get("total_ctn", 0) or 0)
    total_rln       = int(form.get("total_rln", 0) or 0)
    form_type       = form.get("form_type", "commercial")

    tg_numbers   = form.getlist("tg_numbers")
    descriptions = form.getlist("descriptions")
    uoms         = form.getlist("uoms")
    quantities   = form.getlist("quantities")
    photos       = form.getlist("photos")  # UploadFile objects or empty strings

    # Save uploaded photos
    saved_photos = []
    for photo in photos:
        if hasattr(photo, "filename") and photo.filename:
            ext = os.path.splitext(photo.filename)[1].lower()
            filename = f"{uuid.uuid4().hex}{ext}"
            dest = os.path.join(UPLOAD_DIR, filename)
            with open(dest, "wb") as f:
                shutil.copyfileobj(photo.file, f)
            saved_photos.append(filename)
        else:
            saved_photos.append("")

    lines = []
    for i, qty in enumerate(quantities):
        desc = descriptions[i] if i < len(descriptions) else "Stock Rotation by Email"
        desc = desc.strip() or "Stock Rotation by Email"
        try:
            lines.append({
                "tg_number":   tg_numbers[i].strip() if i < len(tg_numbers) else "",
                "description": desc,
                "uom":         uoms[i].strip() if i < len(uoms) else "Pcs",
                "qty":         int(qty),
                "picture_ref": saved_photos[i] if i < len(saved_photos) else "",
            })
        except (ValueError, IndexError):
            pass

    if not lines:
        return templates.TemplateResponse("store_form.html", {
            "request": request, "session": s, "stores": stores,
            "today": date.today().isoformat(), "form_type": form_type,
            "error": "Please add at least one line.",
        })

    database.create_transfer(
        from_store_id=s["store_id"],
        to_store_id=to_store_id,
        collection_date=collection_date,
        delivery_date=delivery_date,
        total_pcs=sum(l["qty"] for l in lines),
        total_ctn=total_ctn,
        total_rln=total_rln,
        lines=lines,
        form_type=form_type,
    )
    return RedirectResponse("/store?submitted=1", status_code=302)


# ── Logistics dashboard ───────────────────────────────────────────────────────

@app.get("/logistics", response_class=HTMLResponse)
async def logistics_dashboard(request: Request, status: str = ""):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    all_transfers = database.get_all_transfers()
    transfers = [t for t in all_transfers if t["status"] == status] if status else all_transfers
    counts = {
        "all":       len(all_transfers),
        "pending":   sum(1 for t in all_transfers if t["status"] == "pending"),
        "approved":  sum(1 for t in all_transfers if t["status"] == "approved"),
        "warehouse": sum(1 for t in all_transfers if t["status"] == "warehouse"),
        "rejected":  sum(1 for t in all_transfers if t["status"] == "rejected"),
        "completed": sum(1 for t in all_transfers if t["status"] == "completed"),
        "incorrect": sum(1 for t in all_transfers if t["status"] == "incorrect"),
    }
    stores = database.get_all_stores(exclude_admin=False)
    return templates.TemplateResponse("logistics.html", {
        "request": request, "session": s,
        "transfers": transfers,
        "active_filter": status or "all",
        "counts": counts,
        "stores": stores,
    })


@app.post("/logistics/transfer/{tid}/delete")
async def delete_transfer(request: Request, tid: int):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    database.delete_transfer(tid)
    return RedirectResponse("/logistics", status_code=302)


@app.get("/logistics/transfer/{tid}", response_class=HTMLResponse)
async def transfer_detail(request: Request, tid: int):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    transfer = database.get_transfer_detail(tid)
    if not transfer:
        return RedirectResponse("/logistics")
    stores = database.get_all_stores()
    return templates.TemplateResponse("transfer_detail.html", {
        "request": request, "session": s,
        "transfer": transfer, "stores": stores,
    })


@app.get("/logistics/export.csv")
async def export_csv(request: Request):
    import csv, io
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")

    transfers = database.get_all_transfers_with_lines()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Collection No", "Type", "Origin", "Destination",
        "Collection Date", "Delivery Date",
        "TG Number", "Description", "UOM", "Qty",
        "Total Pcs", "Total Ctn", "Total Rln",
        "Status", "Submitted",
    ])
    for t in transfers:
        common = [
            t["collect_no"],
            t.get("form_type", "commercial"),
            t["from_store_name"],
            t["to_store_name"],
            t["collection_date"],
            t["delivery_date"],
        ]
        tail = [
            t["total_pcs"],
            t["total_ctn"],
            t.get("total_rln", 0),
            t["status"],
            t["submitted_at"][:10],
        ]
        if t["lines"]:
            for line in t["lines"]:
                writer.writerow(common + [
                    line.get("tg_number", ""),
                    line.get("description", ""),
                    line.get("uom", ""),
                    line.get("qty", ""),
                ] + tail)
        else:
            # transfer with no lines — still export one row
            writer.writerow(common + ["", "", "", ""] + tail)

    filename = f"sacoor-logistics-{date.today().isoformat()}.csv"
    return Response(
        content=output.getvalue().encode("utf-8-sig"),  # BOM for Excel
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/logistics/transfer/{tid}/approve")
async def approve(request: Request, tid: int):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    database.update_transfer_status(tid, "approved")
    return RedirectResponse(f"/logistics/transfer/{tid}?approved=1", status_code=302)


@app.post("/logistics/transfer/{tid}/reject")
async def reject(request: Request, tid: int, reason: str = Form("")):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    database.update_transfer_status(tid, "rejected", reason)
    return RedirectResponse(f"/logistics/transfer/{tid}?rejected=1", status_code=302)


@app.post("/store/transfer/{tid}/receipt")
async def store_receipt(request: Request, tid: int):
    """Destination store confirms receipt as completed or incorrect."""
    s = get_session(request)
    if not s or s["is_admin"]:
        return RedirectResponse("/login")
    form = await request.form()
    receipt_status = form.get("receipt_status", "")
    receipt_note   = form.get("receipt_note", "").strip()
    if receipt_status in ("completed", "incorrect"):
        database.update_receipt_status(tid, receipt_status, receipt_note)
    return RedirectResponse("/store", status_code=302)


@app.post("/logistics/transfer/{tid}/receipt")
async def logistics_receipt(request: Request, tid: int):
    """Logistics team can also update receipt status."""
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    form = await request.form()
    receipt_status = form.get("receipt_status", "")
    receipt_note   = form.get("receipt_note", "").strip()
    if receipt_status in ("completed", "incorrect"):
        database.update_receipt_status(tid, receipt_status, receipt_note)
    return RedirectResponse(f"/logistics/transfer/{tid}?updated=1", status_code=302)


# ── Transporter portal ────────────────────────────────────────────────────────

@app.get("/transporter", response_class=HTMLResponse)
async def transporter_dashboard(request: Request):
    s = get_session(request)
    if not s or not s.get("is_transporter"):
        return RedirectResponse("/login")
    all_transfers = database.get_transfers_for_transporter()
    # DHL-SACOOR is the warehouse store — transporter can complete those directly
    dhl_store = database.get_store_by_username("dhl_sacoor")
    dhl_store_id = dhl_store["id"] if dhl_store else None
    in_transit   = [t for t in all_transfers if t["to_store_id"] != dhl_store_id]
    wh_receipts  = [t for t in all_transfers if t["to_store_id"] == dhl_store_id]
    counts = {
        "awaiting":  sum(1 for t in in_transit  if t["status"] == "approved"),
        "warehouse": sum(1 for t in in_transit  if t["status"] == "warehouse"),
        "delivered": sum(1 for t in in_transit  if t["status"] in ("completed", "incorrect")),
        "wh_pending":   sum(1 for t in wh_receipts if t["status"] in ("approved", "warehouse")),
        "wh_completed": sum(1 for t in wh_receipts if t["status"] in ("completed", "incorrect")),
    }
    return templates.TemplateResponse("transporter_dashboard.html", {
        "request": request, "session": s,
        "in_transit": in_transit, "wh_receipts": wh_receipts,
        "counts": counts, "today": date.today().isoformat(),
    })


@app.post("/transporter/transfer/{tid}/warehouse")
async def mark_warehouse(request: Request, tid: int):
    s = get_session(request)
    if not s or not s.get("is_transporter"):
        return RedirectResponse("/login")
    form = await request.form()
    warehouse_date = form.get("warehouse_date", date.today().isoformat())
    database.update_transporter_status(tid, "warehouse", warehouse_date)
    return RedirectResponse("/transporter?warehouse=1", status_code=302)


@app.post("/transporter/transfer/{tid}/complete")
async def transporter_complete(request: Request, tid: int):
    """Transporter marks a DHL-SACOOR destined transfer as completed."""
    s = get_session(request)
    if not s or not s.get("is_transporter"):
        return RedirectResponse("/login")
    database.update_receipt_status(tid, "completed")
    return RedirectResponse("/transporter?completed=1", status_code=302)


@app.post("/transporter/transfer/{tid}/out-for-delivery")
async def mark_out_for_delivery(request: Request, tid: int):
    s = get_session(request)
    if not s or not s.get("is_transporter"):
        return RedirectResponse("/login")
    database.update_transporter_status(tid, "approved")
    return RedirectResponse("/transporter", status_code=302)


@app.post("/logistics/bulk-approve")
async def bulk_approve(request: Request):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    form = await request.form()
    ids = form.getlist("transfer_ids")
    for tid in ids:
        try:
            database.update_transfer_status(int(tid), "approved")
        except Exception:
            pass
    approved_count = len(ids)
    return RedirectResponse(f"/logistics?bulk_approved={approved_count}", status_code=302)


@app.post("/logistics/transfer/{tid}/edit")
async def edit_transfer(request: Request, tid: int):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")

    form = await request.form()
    to_store_id     = int(form.get("to_store_id", 0))
    collection_date = form.get("collection_date", "")
    delivery_date   = form.get("delivery_date", "")
    total_ctn       = int(form.get("total_ctn", 0) or 0)
    total_rln       = int(form.get("total_rln", 0) or 0)

    tg_numbers   = form.getlist("tg_numbers")
    descriptions = form.getlist("descriptions")
    uoms         = form.getlist("uoms")
    quantities   = form.getlist("quantities")
    picture_refs = form.getlist("picture_refs")  # existing filenames
    photos       = form.getlist("photos")         # new uploads

    # Save any new photos; keep existing ref if no new file uploaded
    final_refs = []
    for i in range(len(quantities)):
        photo = photos[i] if i < len(photos) else None
        if hasattr(photo, "filename") and photo.filename:
            ext = os.path.splitext(photo.filename)[1].lower()
            filename = f"{uuid.uuid4().hex}{ext}"
            dest = os.path.join(UPLOAD_DIR, filename)
            with open(dest, "wb") as f:
                shutil.copyfileobj(photo.file, f)
            final_refs.append(filename)
        else:
            final_refs.append(picture_refs[i].strip() if i < len(picture_refs) else "")

    lines = []
    for i, qty in enumerate(quantities):
        try:
            lines.append({
                "tg_number":   tg_numbers[i].strip() if i < len(tg_numbers) else "",
                "description": descriptions[i].strip() if i < len(descriptions) else "Stock Rotation by Email",
                "uom":         uoms[i].strip() if i < len(uoms) else "Pcs",
                "qty":         int(qty),
                "picture_ref": final_refs[i] if i < len(final_refs) else "",
            })
        except (ValueError, IndexError):
            pass

    database.update_transfer(tid, to_store_id, collection_date, delivery_date, total_ctn, total_rln, lines)
    return RedirectResponse(f"/logistics/transfer/{tid}?edited=1", status_code=302)


# ── Stock Reports (dashboards) ────────────────────────────────────────────────

# Stock data file lives next to the DB on the Railway volume
_STOCK_DATA_PATH = os.path.join(
    os.path.dirname(os.path.abspath(database.DB_PATH)), "stock_data.xlsx"
)


@app.get("/reports", response_class=HTMLResponse)
async def reports_hub(request: Request, success: str = "", error: str = ""):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    import datetime as _dt
    stock_updated = None
    if os.path.exists(_STOCK_DATA_PATH):
        ts = os.path.getmtime(_STOCK_DATA_PATH)
        stock_updated = _dt.datetime.fromtimestamp(ts).strftime("%d %b %Y  %H:%M")
    return templates.TemplateResponse("reports.html", {
        "request": request, "session": s,
        "stock_updated": stock_updated,
        "success": success, "error": error,
    })


@app.post("/reports/upload-stock")
async def upload_stock_data(request: Request, file: UploadFile = File(...)):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    filename = file.filename or ""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        return RedirectResponse("/reports?error=Only+.xlsx+files+accepted", status_code=302)
    content = await file.read()
    with open(_STOCK_DATA_PATH, "wb") as f:
        f.write(content)
    return RedirectResponse("/reports?success=Stock+data+updated+successfully", status_code=302)


@app.get("/reports/stock-data")
async def serve_stock_data(request: Request):
    """Serve the stored stock Excel file to dashboards."""
    s = get_session(request)
    if not s:
        from fastapi.responses import Response as _Resp
        return _Resp(status_code=401)
    if not os.path.exists(_STOCK_DATA_PATH):
        from fastapi.responses import Response as _Resp
        return _Resp(status_code=404)
    with open(_STOCK_DATA_PATH, "rb") as f:
        data = f.read()
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-cache"},
    )


# ── Monthly file helper ───────────────────────────────────────────────────────

_DATA_DIR = os.path.dirname(os.path.abspath(database.DB_PATH))

_MONTHS = [
    ("01","January"),("02","February"),("03","March"),("04","April"),
    ("05","May"),("06","June"),("07","July"),("08","August"),
    ("09","September"),("10","October"),("11","November"),("12","December"),
]

def _monthly_context(prefix: str, upload_url: str, download_url: str,
                     page_title: str, page_subtitle: str, page_icon: str,
                     success: str = "", error: str = "") -> dict:
    import datetime as _dt
    months = []
    for val, label in _MONTHS:
        path = os.path.join(_DATA_DIR, f"{prefix}_{val}.xlsx")
        exists = os.path.exists(path)
        updated = ""
        if exists:
            ts = os.path.getmtime(path)
            updated = _dt.datetime.fromtimestamp(ts).strftime("%d %b %Y")
        months.append({"value": val, "label": label, "available": exists, "updated": updated})
    return {
        "months": months,
        "upload_url": upload_url,
        "download_url": download_url,
        "page_title": page_title,
        "page_subtitle": page_subtitle,
        "page_icon": page_icon,
        "success": success,
        "error": error,
    }


async def _handle_monthly_upload(request: Request, prefix: str, file: UploadFile) -> str:
    """Save uploaded file; return error string or empty string on success."""
    form = await request.form()
    month = form.get("month", "")
    if month not in [m[0] for m in _MONTHS]:
        return "Invalid month selected."
    filename = file.filename or ""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        return "Only .xlsx files are accepted."
    content = await file.read()
    dest = os.path.join(_DATA_DIR, f"{prefix}_{month}.xlsx")
    with open(dest, "wb") as f:
        f.write(content)
    return ""


# ── Inventory Results ─────────────────────────────────────────────────────────

@app.get("/inventory-results", response_class=HTMLResponse)
async def inventory_results(request: Request, success: str = "", error: str = ""):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    ctx = _monthly_context(
        prefix="stock_count_2026",
        upload_url="/inventory-results/upload",
        download_url="/inventory-results/download",
        page_title="Inventory Results — Stock Count 2026",
        page_subtitle="Monthly stock count files · Upload by logistics, download by all stores",
        page_icon="📋",
        success=success, error=error,
    )
    return templates.TemplateResponse("monthly_files.html", {"request": request, "session": s, **ctx,
                                                              "dashboard_url": "/static/dashboards/stock_count.html"})


@app.post("/inventory-results/upload")
async def inventory_upload(request: Request, file: UploadFile = File(...)):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    # re-read form for month (file already consumed by FastAPI)
    form = await request.form()
    month = str(form.get("month", ""))
    if month not in [m[0] for m in _MONTHS]:
        return RedirectResponse("/inventory-results?error=Invalid+month", status_code=302)
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in (".xlsx", ".xls"):
        return RedirectResponse("/inventory-results?error=Only+.xlsx+accepted", status_code=302)
    content = await file.read()
    dest = os.path.join(_DATA_DIR, f"stock_count_2026_{month}.xlsx")
    with open(dest, "wb") as f:
        f.write(content)
    month_name = dict(_MONTHS)[month]
    return RedirectResponse(f"/inventory-results?success={month_name}+uploaded+successfully", status_code=302)


@app.get("/inventory-results/download/{month}")
async def inventory_download(request: Request, month: str):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    path = os.path.join(_DATA_DIR, f"stock_count_2026_{month}.xlsx")
    if not os.path.exists(path):
        return RedirectResponse("/inventory-results")
    with open(path, "rb") as f:
        data = f.read()
    month_name = dict(_MONTHS).get(month, month)
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="StockCount_2026_{month_name}.xlsx"'},
    )


@app.get("/inventory-results/available-months")
async def inventory_available_months(request: Request):
    from fastapi.responses import JSONResponse
    s = get_session(request)
    if not s:
        return JSONResponse({"error": "Not authenticated"}, status_code=401)
    months = []
    for val, label in _MONTHS:
        path = os.path.join(_DATA_DIR, f"stock_count_2026_{val}.xlsx")
        if os.path.exists(path):
            months.append({"value": val, "label": f"{label} 2026"})
    return JSONResponse(months)


# ── NOOS Sales (single file) ──────────────────────────────────────────────────

def _noos_path():
    return os.path.join(_DATA_DIR, "noos_sales_all.xlsx")

@app.get("/noos-sales", response_class=HTMLResponse)
async def noos_sales(request: Request, success: str = "", error: str = ""):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    path = _noos_path()
    file_info = None
    if os.path.exists(path):
        stat = os.stat(path)
        import datetime as _dt2
        file_info = {
            "size_mb": round(stat.st_size / 1024 / 1024, 1),
            "updated": _dt2.datetime.fromtimestamp(stat.st_mtime).strftime("%d %b %Y %H:%M")
        }
    return templates.TemplateResponse("noos_portal.html", {
        "request": request, "session": s,
        "file_info": file_info, "success": success, "error": error,
        "dashboard_url": "/static/dashboards/noos_sales.html"
    })

@app.post("/noos-sales/upload")
async def noos_upload(request: Request, file: UploadFile = File(...)):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in (".xlsx", ".xls"):
        return RedirectResponse("/noos-sales?error=Only+.xlsx+accepted", status_code=302)
    content = await file.read()
    with open(_noos_path(), "wb") as f:
        f.write(content)
    return RedirectResponse("/noos-sales?success=NOOS+sales+data+uploaded+successfully", status_code=302)

@app.get("/noos-sales/data")
async def noos_data(request: Request):
    s = get_session(request)
    if not s:
        return Response(status_code=401)
    path = _noos_path()
    if not os.path.exists(path):
        return Response(status_code=404)
    with open(path, "rb") as f:
        data = f.read()
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'inline; filename="noos_sales.xlsx"'}
    )


# ── Stock Delivery ────────────────────────────────────────────────────────────

def _parse_date_str(val) -> str:
    """Convert Excel date object or string to YYYY-MM-DD string."""
    if val is None:
        return ""
    from datetime import date as _date, datetime as _datetime
    if isinstance(val, (_date, _datetime)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    # try common formats
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            from datetime import datetime as _dt
            return _dt.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s  # return as-is if can't parse


def _col(headers: list, *names) -> int | None:
    """Find column index (0-based) by trying several name variants."""
    lower_headers = [h.lower().strip() for h in headers]
    for name in names:
        try:
            return lower_headers.index(name.lower().strip())
        except ValueError:
            pass
    return None


def _v(row, idx):
    """Safely get a cell value by index, return empty string if out of range."""
    if idx is None or idx >= len(row):
        return ""
    val = row[idx]
    return "" if val is None else str(val).strip()


@app.get("/stock-delivery", response_class=HTMLResponse)
async def store_stock_delivery(request: Request):
    s = get_session(request)
    if not s or s["is_admin"] or s.get("is_transporter"):
        return RedirectResponse("/login")
    from datetime import date as _date, timedelta as _td
    today = _date.today()
    today_str = today.isoformat()
    # Monday of current week
    week_start = (today - _td(days=today.weekday())).isoformat()
    week_num   = today.isocalendar()[1]
    raw = database.get_deliveries_for_store(s["store_code"])
    upcoming = 0
    deliveries = []
    for d in raw:
        dd = d.get("delivery_date", "")
        d["is_today"] = dd == today_str
        d["is_past"]  = dd < today_str if dd else False
        if not d["is_past"]:
            upcoming += 1
        try:
            d["total_ctn_int"] = int(d.get("total_ctn") or 0)
        except ValueError:
            d["total_ctn_int"] = 0
        deliveries.append(d)
    upload_info = database.get_delivery_upload_info()
    return templates.TemplateResponse("stock_delivery.html", {
        "request": request, "session": s,
        "deliveries": deliveries,
        "upcoming_count": upcoming,
        "upload_info": upload_info,
        "week_start": week_start,
        "week_num": week_num,
        "today": today_str,
    })


@app.get("/logistics/stock-delivery", response_class=HTMLResponse)
async def logistics_stock_delivery(request: Request, error: str = "", success: str = ""):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")
    from datetime import date as _date, timedelta as _td
    today      = _date.today()
    week_start = (today - _td(days=today.weekday())).isoformat()
    deliveries  = database.get_all_deliveries()
    upload_info = database.get_delivery_upload_info()
    store_codes = sorted({d["store_code"] for d in deliveries if d["store_code"]})
    # Find codes that don't match any system store (excluding ignored ones)
    all_stores  = database.get_all_stores(exclude_admin=False)
    sys_codes   = {st["store_code"].upper() for st in all_stores}
    unmatched   = sorted({
        c for c in store_codes
        if c.upper() not in sys_codes and c.upper() not in DELIVERY_IGNORE
    })
    return templates.TemplateResponse("logistics_stock_delivery.html", {
        "request": request, "session": s,
        "deliveries": deliveries,
        "upload_info": upload_info,
        "store_codes": store_codes,
        "unmatched": unmatched,
        "week_start": week_start,
        "today": today.isoformat(),
        "error": error, "success": success,
    })


@app.post("/logistics/stock-delivery/upload")
async def upload_stock_delivery(request: Request, file: UploadFile = File(...)):
    s = get_session(request)
    if not s or not s["is_admin"]:
        return RedirectResponse("/login")

    filename = file.filename or ""
    ext = os.path.splitext(filename)[1].lower()

    rows = []
    try:
        content = await file.read()

        if ext == ".xlsx":
            import openpyxl, io as _io
            wb = openpyxl.load_workbook(_io.BytesIO(content), data_only=True)
            ws = wb.active
            all_rows = list(ws.iter_rows(values_only=True))
            if not all_rows:
                raise ValueError("Empty file")
            headers = [str(h or "").strip() for h in all_rows[0]]
            data_rows = all_rows[1:]

        elif ext == ".csv":
            import csv as _csv, io as _io
            text = content.decode("utf-8-sig")
            reader = list(_csv.reader(_io.StringIO(text)))
            if not reader:
                raise ValueError("Empty file")
            headers = [h.strip() for h in reader[0]]
            data_rows = reader[1:]
        else:
            raise ValueError("Only .xlsx or .csv files are accepted.")

        # Map header names → column indices
        ci = {
            "no":               _col(headers, "no"),
            "mth":              _col(headers, "mth", "mth (auto)", "month"),
            "order_rtn_no":     _col(headers, "order/rtn no", "order rtn no", "order no"),
            "store_code":       _col(headers, "store"),
            "store_name_excel": _col(headers, "storename", "store name"),
            "order_drop_date":  _col(headers, "sacoor order drop", "order drop"),
            "picking_complete": _col(headers, "picking complete", "picking completed"),
            "cargo_readiness":  _col(headers, "cargo readiness", "cargo ready"),
            "delivery_date":    _col(headers, "delivery date"),
            "delivery_day":     _col(headers, "delivery day"),
            "delivery_time":    _col(headers, "delivery time"),
            "order_type":       _col(headers, "order type"),
            "qty":              _col(headers, "qty", "quantity"),
            "vol":              _col(headers, "vol", "volume"),
            "ncg":              _col(headers, "ncg"),
            "total_ctn":        _col(headers, "total ctn", "total cartons"),
            "total_rln":        _col(headers, "total railings", "total rln", "total railing"),
        }

        for raw in data_rows:
            raw = list(raw)
            # Skip completely empty rows
            if not any(v for v in raw if v is not None and str(v).strip()):
                continue
            # Normalise store code: uppercase → apply mapping → skip ignored
            raw_code = _v(raw, ci["store_code"]).strip().upper()
            norm_code = DELIVERY_CODE_MAP.get(raw_code, raw_code)
            if norm_code in DELIVERY_IGNORE:
                continue
            rows.append({
                "no":               _v(raw, ci["no"]),
                "mth":              _v(raw, ci["mth"]),
                "order_rtn_no":     _v(raw, ci["order_rtn_no"]),
                "store_code":       norm_code,
                "store_name_excel": _v(raw, ci["store_name_excel"]),
                "order_drop_date":  _parse_date_str(raw[ci["order_drop_date"]] if ci["order_drop_date"] is not None and ci["order_drop_date"] < len(raw) else ""),
                "picking_complete": _parse_date_str(raw[ci["picking_complete"]] if ci["picking_complete"] is not None and ci["picking_complete"] < len(raw) else ""),
                "cargo_readiness":  _parse_date_str(raw[ci["cargo_readiness"]] if ci["cargo_readiness"] is not None and ci["cargo_readiness"] < len(raw) else ""),
                "delivery_date":    _parse_date_str(raw[ci["delivery_date"]] if ci["delivery_date"] is not None and ci["delivery_date"] < len(raw) else ""),
                "delivery_day":     _v(raw, ci["delivery_day"]),
                "delivery_time":    _v(raw, ci["delivery_time"]),
                "order_type":       _v(raw, ci["order_type"]),
                "qty":              _v(raw, ci["qty"]),
                "vol":              _v(raw, ci["vol"]),
                "ncg":              _v(raw, ci["ncg"]),
                "total_ctn":        _v(raw, ci["total_ctn"]),
                "total_rln":        _v(raw, ci["total_rln"]),
            })

        import uuid as _uuid
        batch_id = _uuid.uuid4().hex[:8]
        database.upload_deliveries(rows, batch_id)

    except Exception as exc:
        deliveries  = database.get_all_deliveries()
        upload_info = database.get_delivery_upload_info()
        store_codes = sorted({d["store_code"] for d in deliveries if d["store_code"]})
        return templates.TemplateResponse("logistics_stock_delivery.html", {
            "request": request, "session": s,
            "deliveries": deliveries,
            "upload_info": upload_info,
            "store_codes": store_codes,
            "error": f"Upload failed: {exc}",
            "success": "",
        })

    return RedirectResponse(
        f"/logistics/stock-delivery?success=Uploaded+{len(rows)}+rows+successfully",
        status_code=302
    )


# ── PDF download ──────────────────────────────────────────────────────────────

@app.get("/transfer/{tid}/pdf")
async def download_pdf(request: Request, tid: int):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    transfer = database.get_transfer_detail(tid)
    if not transfer:
        return RedirectResponse("/")
    # Transporters can download any transfer; stores only their own
    if not s["is_admin"] and not s.get("is_transporter"):
        if transfer["from_store_id"] != s["store_id"] and transfer["to_store_id"] != s["store_id"]:
            return RedirectResponse("/store")

    if transfer.get("form_type") == "decoration":
        pdf_bytes = pdf_generator.generate_decoration_pdf(transfer)
    else:
        pdf_bytes = pdf_generator.generate_pdf(transfer)
    filename = f"{transfer['collect_no']}.pdf"
    return StreamingResponse(
        iter([pdf_bytes]),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
