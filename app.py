"""tex2word web app (v4).

- Upload: whole LaTeX project zip (tex + bib + figures + compiled PDF),
  or a scanned PDF for OCR mode.
- Pipeline: pipeline.convert() / ocr_pipeline.ocr_convert() -> output.docx
- Identity: Firebase email/password login (same x-planner-99dd3 project as
  X Planner). The owner is the account whose email matches OWNER_EMAIL.
  The browser sends the Firebase ID
  token as `Authorization: Bearer <token>`; the server verifies it with
  firebase-admin.
- Storage:
    * Owner -> persistent backend: FirebaseStore (Cloud Storage +
      Firestore) when Firebase is configured, else local disk. Files stay
      until manually deleted.
    * Everyone else (guests, incl. other signed-in accounts) -> ephemeral
      guest backend: local temp dir only, NEVER touches Firebase. The
      project is deleted right after the Word file is downloaded; a
      sweeper also removes guest projects idle for more than an hour.
- Run: ./venv/bin/uvicorn app:app --port 8471
"""
import json
import os
import re
import shutil
import tempfile
import threading
import time
import uuid

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from urllib.parse import quote as _urlquote

import pipeline
import ocr_pipeline
import store as store_mod

BASE = os.path.dirname(os.path.abspath(__file__))

# Owner identity: the Firebase account with this email. Fail closed: only a
# cryptographically verified ID token with this email counts as owner.
OWNER_EMAIL = os.environ.get("OWNER_EMAIL", "").lower()
OWNER_UID = "owner"
FIREBASE_PROJECT = "x-planner-99dd3"
# Web API key is a public client identifier (not a secret); the user pastes it
# into the host's env dashboard herself. Injected into the page at serve time.
FIREBASE_WEB_API_KEY = os.environ.get("FIREBASE_WEB_API_KEY", "")

GUEST_TTL_SEC = 3600      # sweeper: guest projects idle this long are wiped
GUEST_SWEEP_EVERY = 600   # sweeper interval

app = FastAPI(title="tex2word")

owner_store = store_mod.get_store()  # firebase if configured, else disk
guest_store = store_mod.LocalStore(tempfile.mkdtemp(prefix="tex2word-guests-"))
print(f"tex2word owner backend: {owner_store.kind}; "
      f"guest backend: ephemeral local", flush=True)


_fb_app = None
_fb_ok = None


def _fb_app_init():
    """Lazily init firebase-admin. Returns the app or None (auth disabled)."""
    global _fb_app, _fb_ok
    if _fb_ok is not None:
        return _fb_app
    try:
        import firebase_admin
        from firebase_admin import credentials
        try:
            # The store module may have initialized the default app already;
            # initialize_app() raises if called twice.
            _fb_app = firebase_admin.get_app()
        except ValueError:
            sa = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON") or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
            opts = {"projectId": FIREBASE_PROJECT}
            if sa:
                _fb_app = firebase_admin.initialize_app(
                    credentials.Certificate(json.loads(sa)), opts)
            else:
                # No service account (e.g. local dev): token verification will
                # fail closed; everyone is treated as guest.
                _fb_app = firebase_admin.initialize_app(options=opts)
        _fb_ok = True
    except Exception as e:
        print(f"tex2word: firebase-admin unavailable ({e}); "
              f"owner login disabled", flush=True)
        _fb_app, _fb_ok = None, False
    return _fb_app


def _owner_email(request: Request):
    """Verified owner email from the Authorization header, else None."""
    app = _fb_app_init()
    if not app:
        return None
    h = request.headers.get("authorization", "")
    if not h.lower().startswith("bearer "):
        return None
    try:
        from firebase_admin import auth as fb_auth
        decoded = fb_auth.verify_id_token(h[7:].strip(), app=app)
        email = (decoded.get("email") or "").lower()
        if email == OWNER_EMAIL:
            return email
    except Exception:
        pass
    return None


def _route(request: Request, user_id: str):
    """Return (store, effective_uid, is_owner)."""
    if _owner_email(request):
        return owner_store, OWNER_UID, True
    return guest_store, (user_id or "guest")[:64], False


def _sweep_guests():
    root = guest_store.root
    while True:
        time.sleep(GUEST_SWEEP_EVERY)
        now = time.time()
        try:
            for uid in os.listdir(root):
                ud = os.path.join(root, uid)
                if not os.path.isdir(ud):
                    continue
                for pid in os.listdir(ud):
                    pd = os.path.join(ud, pid)
                    if (os.path.isdir(pd)
                            and now - os.path.getmtime(pd) > GUEST_TTL_SEC):
                        shutil.rmtree(pd, ignore_errors=True)
                if not os.listdir(ud):
                    os.rmdir(ud)
        except Exception:
            pass


threading.Thread(target=_sweep_guests, daemon=True).start()


@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(BASE, "templates", "index.html")) as fh:
        html = fh.read()
    return html.replace("__FIREBASE_API_KEY__", FIREBASE_WEB_API_KEY)


@app.get("/api/health")
def health():
    return {"ok": True, "storage": owner_store.kind,
            "auth": "firebase" if _fb_app_init() else "disabled"}


@app.get("/api/me")
def me(request: Request):
    email = _owner_email(request)
    return {"is_owner": bool(email), "email": email or ""}


@app.get("/api/projects")
def list_projects(request: Request, user_id: str = ""):
    st, uid, _ = _route(request, user_id)
    return st.list_projects(uid)


def _new_project(st, uid, filename, name, fileobj, mode=None):
    pid = time.strftime("%Y%m%d-%H%M%S") + "_" + uuid.uuid4().hex[:6]
    blob_name = "scan.pdf" if filename.lower().endswith(".pdf") \
        else "upload.zip"
    st.save_upload(uid, pid, blob_name, fileobj)
    meta = {"id": pid, "name": name or filename, "status": "uploaded",
            "created": pid.split("_")[0]}
    if mode:
        meta["mode"] = mode
    st.set_meta(uid, pid, meta)
    return {"project_id": pid}


@app.post("/api/upload")
async def upload(request: Request, user_id: str = Form(""),
                 name: str = Form(""),
                 file: UploadFile = File(...)):
    """Single upload entry: .zip -> LaTeX pipeline, .pdf -> OCR pipeline."""
    fn = (file.filename or "").lower()
    if fn.endswith(".zip"):
        mode = None
    elif fn.endswith(".pdf"):
        mode = "ocr"
    else:
        raise HTTPException(
            400, "please upload a .zip of your LaTeX project or a scanned .pdf")
    st, uid, _ = _route(request, user_id)
    return _new_project(st, uid, file.filename, name, file.file, mode=mode)


def _run(st, uid, project_id, profile, kind):
    wd = st.work_dir(uid, project_id)
    src = os.path.join(wd, "scan.pdf" if kind == "ocr" else "upload.zip")
    if not os.path.exists(src):
        raise HTTPException(404, "project not found")
    meta = st.get_meta(uid, project_id)
    meta["status"] = "converting"
    st.set_meta(uid, project_id, meta)
    try:
        if kind == "ocr":
            result = ocr_pipeline.ocr_convert(src, wd, profile_name=profile)
        else:
            result = pipeline.convert(wd, profile_name=profile)
    except Exception as e:
        # Never a bare 500: capture the real reason so the UI can show it
        # (and the project meta keeps it for the Log view).
        import traceback
        traceback.print_exc()
        result = {"ok": False, "log": [],
                  "error": f"{type(e).__name__}: {e}"}
    meta["status"] = "done" if result["ok"] else "failed"
    meta["log"] = result.get("log", [])
    if result.get("error"):
        meta["error"] = result["error"]
    st.sync_back(uid, project_id, meta)
    return {"ok": result["ok"], "log": result.get("log", []),
            "error": result.get("error")}


@app.post("/api/ocr/{project_id}")
def ocr(project_id: str, request: Request, user_id: str = "",
        profile: str = "default"):
    st, uid, _ = _route(request, user_id)
    return _run(st, uid, project_id, profile, "ocr")


@app.post("/api/convert/{project_id}")
def convert(project_id: str, request: Request, user_id: str = "",
            profile: str = "default"):
    st, uid, _ = _route(request, user_id)
    return _run(st, uid, project_id, profile, "tex")


def _safe_docx_name(raw, fallback):
    """Download filename: original project name with .docx.

    Strips the uploaded extension (.zip/.pdf), removes path separators
    and control characters, keeps everything else (incl. non-ASCII).
    """
    base = (raw or "").strip() or fallback
    low = base.lower()
    for ext in (".zip", ".pdf", ".tex"):
        if low.endswith(ext):
            base = base[: -len(ext)]
            break
    base = re.sub(r'[\\/:"*?<>|\x00-\x1f]', "_", base).strip(" .") or fallback
    if len(base) > 120:
        base = base[:120].rstrip()
    return base + ".docx"


def _rfc5987(name):
    """Percent-encode a filename for the filename*=UTF-8'' parameter."""
    return _urlquote(name, safe="")


@app.get("/api/download/{project_id}")
def download(project_id: str, request: Request,
             background_tasks: BackgroundTasks, user_id: str = ""):
    st, uid, is_owner = _route(request, user_id)
    meta = st.get_meta(uid, project_id) or {}
    filename = _safe_docx_name(meta.get("name"), project_id)
    url = st.download_url(uid, project_id, response_disposition=(
        "attachment; filename*=UTF-8''" + _rfc5987(filename)))
    if url:
        # owner on Firebase: redirect to a signed URL (no deletion)
        return RedirectResponse(url)
    fp = st.download_path(uid, project_id)
    if not fp:
        raise HTTPException(404, "no Word file yet — convert first")
    if not is_owner:
        # guest policy: wipe the project right after the file is served
        background_tasks.add_task(st.delete_project, uid, project_id)
    return FileResponse(fp, filename=filename,
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.delete("/api/projects/{project_id}")
def delete_project(project_id: str, request: Request, user_id: str = ""):
    """Owner-only manual delete. Guest files auto-wipe on pagehide
    and right after a guest download is served."""
    st, uid, is_owner = _route(request, user_id)
    if not is_owner:
        raise HTTPException(403, "only the owner can delete projects")
    st.delete_project(uid, project_id)
    return {"deleted": project_id}


@app.post("/api/logout-cleanup")
def logout_cleanup(request: Request, user_id: str = ""):
    """Guest policy: on logout, wipe all of the user's data."""
    st, uid, is_owner = _route(request, user_id)
    if is_owner:
        return {"cleaned": 0, "note": "owner files are kept"}
    n = st.logout_cleanup(uid)
    return {"cleaned": n}
