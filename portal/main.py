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


# ── Root ──────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    return RedirectResponse("/logistics" if s["is_admin"] else "/store")


# ── Auth ──────────────────────────────────────────────────────────────────────

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    s = get_session(request)
    if s:
        return RedirectResponse("/logistics" if s["is_admin"] else "/store")
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
        "store_id":   store["id"],
        "store_name": store["store_name"],
        "store_code": store["store_code"],
        "username":   store["username"],
        "is_admin":   bool(store["is_admin"]),
    }
    url = "/logistics" if store["is_admin"] else "/store"
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
    transfers = database.get_transfers_by_store(s["store_id"])
    incoming  = database.get_incoming_transfers(s["store_id"])
    return templates.TemplateResponse("store_dashboard.html", {
        "request": request, "session": s,
        "transfers": transfers, "incoming": incoming,
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
        "all":      len(all_transfers),
        "pending":  sum(1 for t in all_transfers if t["status"] == "pending"),
        "approved": sum(1 for t in all_transfers if t["status"] == "approved"),
        "rejected": sum(1 for t in all_transfers if t["status"] == "rejected"),
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


# ── PDF download ──────────────────────────────────────────────────────────────

@app.get("/transfer/{tid}/pdf")
async def download_pdf(request: Request, tid: int):
    s = get_session(request)
    if not s:
        return RedirectResponse("/login")
    transfer = database.get_transfer_detail(tid)
    if not transfer:
        return RedirectResponse("/")
    # Stores can only download their own approved transfers
    if not s["is_admin"] and transfer["from_store_id"] != s["store_id"]:
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
