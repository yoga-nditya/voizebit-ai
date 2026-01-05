import uuid
import json
import re
import platform
from datetime import datetime
from pathlib import Path
from urllib.parse import quote  # ✅ untuk encode filename (spasi, dll)

from flask import Flask, request, jsonify, render_template, send_from_directory, session, send_file

from config_new import *
from limbah_database import LIMBAH_B3_DB
from utils import (
    init_db, load_counter,
    db_insert_history, db_list_histories, db_get_history_detail,
    db_update_title, db_delete_history, db_append_message, db_update_state,
    call_ai,
    PDF_AVAILABLE, PDF_METHOD
)

from quotation import handle_quotation_flow
from invoice import handle_invoice_flow
from mou import handle_mou_flow

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = FLASK_SECRET_KEY

# =========================
# ✅ Pastikan folder file ada: static/files
# =========================
# FILES_DIR dari config_new.py = Path(BASE_DIR/static/files)
# Di sini kita pakai Path object konsisten.
try:
    FILES_DIR_PATH = Path(FILES_DIR)  # type: ignore
except Exception:
    FILES_DIR_PATH = Path("static") / "files"

FILES_DIR_PATH.mkdir(parents=True, exist_ok=True)

# state memory per session
conversations = {}
init_db()

@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    return resp

@app.route("/")
def index():
    return render_template("index.html")

# =========================
# HISTORY (DB: chat_history)
# =========================
@app.route("/api/history", methods=["GET"])
def api_history_list():
    try:
        q = (request.args.get("q") or "").strip()
        items = db_list_histories(limit=200, q=q if q else None)
        return jsonify({"items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<int:history_id>", methods=["GET"])
def api_history_detail(history_id):
    try:
        detail = db_get_history_detail(history_id)
        if not detail:
            return jsonify({"error": "history tidak ditemukan"}), 404

        return jsonify({
            "id": detail["id"],
            "title": detail["title"],
            "task_type": detail["task_type"],
            "created_at": detail["created_at"],
            "data": json.loads(detail.get("data_json") or "{}"),
            "files": json.loads(detail.get("files_json") or "[]"),
            "messages": json.loads(detail.get("messages_json") or "[]"),
            "state": json.loads(detail.get("state_json") or "{}"),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<int:history_id>", methods=["PUT"])
def api_history_update(history_id):
    try:
        body = request.get_json() or {}
        new_title = (body.get("title") or "").strip()
        if not new_title:
            return jsonify({"error": "title wajib diisi"}), 400
        ok = db_update_title(history_id, new_title)
        if not ok:
            return jsonify({"error": "history tidak ditemukan"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<int:history_id>", methods=["DELETE"])
def api_history_delete(history_id):
    try:
        ok = db_delete_history(history_id)
        if not ok:
            return jsonify({"error": "history tidak ditemukan"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ==========================================
# DOCUMENTS (DB-based) - endpoint lama
# ==========================================
@app.route("/api/documents", methods=["GET"])
def api_documents():
    """
    Endpoint lama: list dokumen berdasarkan DB history (files_json).
    Tetap ada kalau masih dipakai.
    """
    try:
        q = (request.args.get("q") or "").strip().lower()
        items = db_list_histories(limit=500)

        docs = []
        for h in items:
            detail = db_get_history_detail(int(h["id"]))
            if not detail:
                continue
            try:
                files = json.loads(detail.get("files_json") or "[]")
            except:
                files = []

            for f in files:
                filename = (f.get("filename") or "").strip()
                if not filename:
                    continue

                title = detail.get("title") or ""
                task_type = detail.get("task_type") or ""
                created_at = detail.get("created_at") or ""

                url = (f.get("url") or "").strip()
                if not url:
                    url = f"/download/{quote(filename)}"

                row = {
                    "history_id": int(detail["id"]),
                    "history_title": title,
                    "task_type": task_type,
                    "created_at": created_at,
                    "type": (f.get("type") or Path(filename).suffix.lower().lstrip(".")) or "file",
                    "filename": filename,
                    "url": url,
                }

                if q:
                    hay = f"{title} {filename} {task_type}".lower()
                    if q not in hay:
                        continue

                docs.append(row)

        docs.sort(key=lambda x: x.get("created_at") or "", reverse=True)
        return jsonify({"items": docs})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =========================
# Helper: reset / cancel flow
# =========================
def is_cancel_cmd(lower: str) -> bool:
    keys = ["batal", "cancel", "reset", "ulang", "start over", "keluar"]
    return any(k in lower for k in keys)

def reset_state(sid: str):
    conversations[sid] = {"step": "idle", "data": {}}

# =========================
# ✅ Helper: PDF thumbnail generator (page 1 -> PNG)
# =========================
THUMB_DIR = Path("static") / "thumbs"
THUMB_DIR.mkdir(parents=True, exist_ok=True)

def _safe_thumb_name(filename: str) -> str:
    safe = re.sub(r'[^a-zA-Z0-9._-]+', '_', filename)
    return f"{safe}.png"

def generate_pdf_thumbnail(pdf_path: Path, out_path: Path) -> bool:
    """
    Return True kalau thumbnail berhasil dibuat.
    Butuh PyMuPDF (fitz).
    """
    try:
        import fitz  # PyMuPDF
    except Exception:
        return False

    try:
        doc = fitz.open(str(pdf_path))
        if doc.page_count == 0:
            return False

        page = doc.load_page(0)
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(out_path))
        doc.close()
        return True
    except Exception:
        return False

# =========================
# CHAT
# =========================
@app.route("/api/chat", methods=["POST"])
def chat():
    try:
        data = request.get_json() or {}
        text = (data.get("message", "") or "").strip()
        history_id_in = data.get("history_id")

        if not text:
            return jsonify({"error": "Pesan kosong"}), 400

        sid = request.headers.get("X-Session-ID") or session.get("sid")
        if not sid:
            sid = str(uuid.uuid4())
            session["sid"] = sid

        state = conversations.get(sid, {"step": "idle", "data": {}})
        conversations[sid] = state
        lower = text.lower().strip()

        if is_cancel_cmd(lower):
            reset_state(sid)
            out = "✅ Flow dibatalkan. Kamu mau buat <b>invoice</b>, <b>mou</b>, atau <b>quotation</b>?"
            if history_id_in:
                try:
                    db_append_message(int(history_id_in), "user", text, files=[])
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out), files=[])
                    db_update_state(int(history_id_in), conversations[sid])
                except:
                    pass
            return jsonify({"text": out, "history_id": history_id_in})

        if history_id_in:
            try:
                db_append_message(int(history_id_in), "user", text, files=[])
                db_update_state(int(history_id_in), state)
            except:
                pass

        # 1) Invoice
        resp = handle_invoice_flow(data, text, lower, sid, state, conversations, history_id_in)
        if resp is not None:
            return jsonify(resp)

        # 2) MoU
        resp = handle_mou_flow(data, text, lower, sid, state, conversations, history_id_in)
        if resp is not None:
            return jsonify(resp)

        # 3) Quotation
        resp = handle_quotation_flow(data, text, lower, sid, state, conversations, history_id_in)
        if resp is not None:
            return jsonify(resp)

        ai_out = call_ai(text)
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", ai_out, files=[])
        return jsonify({"text": ai_out, "history_id": history_id_in})

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# =========================
# DOWNLOAD / PREVIEW FILE
# =========================
@app.route("/download/<path:filename>")
def download(filename):
    """
    ✅ Default: inline preview (as_attachment=False) => cocok untuk preview PDF di app
    ✅ Paksa download: /download/<filename>?download=1
    ✅ Thumbnail: /download/<filename>?thumbnail=1
       - untuk PDF: return PNG preview (inline)
       - untuk non-PDF: 404 (biar app fallback ke icon)
    """
    file_path = FILES_DIR_PATH / filename

    if not file_path.exists():
        return jsonify({"error": "file tidak ditemukan"}), 404

    # thumbnail mode
    thumb = (request.args.get("thumbnail") or "").strip().lower()
    if thumb in ("1", "true", "yes"):
        if str(file_path).lower().endswith(".pdf"):
            thumb_path = THUMB_DIR / _safe_thumb_name(filename)

            if (not thumb_path.exists()) or (thumb_path.stat().st_mtime < file_path.stat().st_mtime):
                ok = generate_pdf_thumbnail(file_path, thumb_path)
                if not ok:
                    return jsonify({"error": "thumbnail generator not available"}), 404

            return send_file(str(thumb_path), mimetype="image/png", as_attachment=False)

        return jsonify({"error": "thumbnail hanya untuk pdf"}), 404

    dl = (request.args.get("download") or "").strip().lower()
    as_attachment = dl in ("1", "true", "yes")
    return send_from_directory(str(FILES_DIR_PATH), filename, as_attachment=as_attachment)

# =========================================================
# ✅ COMPANY DOCUMENTS - LIST LANGSUNG DARI static/files
# =========================================================
@app.route("/api/company-documents", methods=["GET"])
def api_company_documents():
    """
    Ambil dokumen langsung dari folder FILES_DIR_PATH (static/files).
    Support query: ?q= untuk filter nama file.

    Output:
    { "items": [ { filename, type, url, created_at, title } ] }
    """
    try:
        q = (request.args.get("q") or "").strip().lower()

        if not FILES_DIR_PATH.exists():
            return jsonify({"items": []})

        items = []
        for p in FILES_DIR_PATH.iterdir():
            if not p.is_file():
                continue

            filename = p.name
            if filename.startswith("."):
                continue

            if q and q not in filename.lower():
                continue

            ext = p.suffix.lower().lstrip(".")
            file_type = ext if ext else "file"
            url = f"/download/{quote(filename)}"
            created_at = datetime.fromtimestamp(p.stat().st_mtime).isoformat()

            items.append({
                "history_id": 0,
                "history_title": "",
                "task_type": "company",
                "created_at": created_at,
                "type": file_type,
                "filename": filename,
                "url": url,
                "title": p.stem,
            })

        items.sort(key=lambda x: x.get("created_at") or "", reverse=True)
        return jsonify({"items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =========================================================
# ✅ ADMIN PURGE - HAPUS CHAT DB + FILE static/files
# =========================================================
@app.route("/api/admin/purge", methods=["POST"])
def api_admin_purge():
    """
    Hapus SEMUA data chat & dokumen.
    Body JSON:
    {
      "confirm": "DELETE_ALL",
      "delete_files": true
    }
    """
    try:
        body = request.get_json() or {}
        confirm = (body.get("confirm") or "").strip()
        delete_files = bool(body.get("delete_files", False))

        if confirm != "DELETE_ALL":
            return jsonify({"error": "confirm invalid. set confirm=DELETE_ALL"}), 400

        # ======================
        # 🗄️ HAPUS DATABASE (TABLE YANG BENAR: chat_history)
        # ======================
        import sqlite3

        conn = sqlite3.connect(str(DB_FILE))
        cur = conn.cursor()
        cur.execute("DELETE FROM chat_history")  # ✅ FIX: bukan histories
        conn.commit()
        conn.close()

        deleted_files = 0

        # ======================
        # 📄 HAPUS FILE DOKUMEN
        # ======================
        if delete_files:
            for p in FILES_DIR_PATH.iterdir():
                if p.is_file():
                    try:
                        p.unlink()
                        deleted_files += 1
                    except Exception as e:
                        print("Gagal hapus file:", p.name, e)

            # hapus thumbnail PDF juga
            if THUMB_DIR.exists():
                for p in THUMB_DIR.iterdir():
                    if p.is_file():
                        try:
                            p.unlink()
                        except:
                            pass

        return jsonify({
            "ok": True,
            "deleted_table": "chat_history",
            "db": str(DB_FILE),
            "files_dir": str(FILES_DIR_PATH),
            "deleted_files": deleted_files
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = FLASK_PORT
    debug_mode = FLASK_DEBUG

    print("\n" + "=" * 60)
    print("🚀 DOCUMENT GENERATOR")
    print("=" * 60)
    try:
        from config_new import TEMPLATE_FILE
        print(f"📁 Template: {TEMPLATE_FILE.exists() and '✅ Found' or '❌ Missing'}")
    except Exception:
        pass
    print(f"📂 FILES_DIR: {FILES_DIR_PATH} (exists={FILES_DIR_PATH.exists()})")
    print(f"🔎 Serper: {SERPER_API_KEY and '✅' or '❌'}")
    print(f"📄 PDF: {PDF_AVAILABLE and f'✅ {PDF_METHOD}' or '❌ Disabled'}")
    print(f"🗄️  Database: {len(LIMBAH_B3_DB)} jenis limbah")
    print(f"🔢 Counter: {load_counter()}")
    print(f"🌐 Port: {port}")
    print(f"💻 Platform: {platform.system()}")
    print("=" * 60 + "\n")

    app.run(host="0.0.0.0", port=port, debug=debug_mode)
