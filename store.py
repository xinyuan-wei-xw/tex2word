"""Persistence backends for tex2word.

- LocalStore: plain disk under ./storage (dev / no Firebase configured).
- FirebaseStore: Cloud Storage (files) + Firestore (project metadata).
  Enabled when FIREBASE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS
  is set and firebase-admin is installed.

The pipelines (pipeline.py, ocr_pipeline.py) always work on a local
directory; FirebaseStore downloads to a temp work dir and syncs back.
"""
import json
import os
import shutil
import tempfile

try:
    import firebase_admin
    from firebase_admin import credentials as fb_credentials
    from firebase_admin import storage as fb_storage
    from firebase_admin import firestore as fb_firestore
    _FB_AVAILABLE = True
except ImportError:  # local dev without firebase-admin
    _FB_AVAILABLE = False


def _service_account_json():
    sa = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
    if sa:
        return sa
    p = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if p and os.path.exists(p):
        with open(p) as fh:
            return fh.read()
    return None


USE_FIREBASE = bool(_service_account_json()) and _FB_AVAILABLE


# ---------------------------------------------------------------- local

class LocalStore:
    kind = "local"

    def __init__(self, root):
        self.root = root
        os.makedirs(root, exist_ok=True)

    def _user_dir(self, user_id):
        d = os.path.join(self.root, user_id)
        os.makedirs(d, exist_ok=True)
        return d

    def _project_dir(self, user_id, project_id):
        return os.path.join(self._user_dir(user_id), project_id)

    # -- API used by app.py -------------------------------------------
    def save_upload(self, user_id, project_id, filename, fileobj):
        pd = self._project_dir(user_id, project_id)
        os.makedirs(pd, exist_ok=True)
        with open(os.path.join(pd, filename), "wb") as fh:
            shutil.copyfileobj(fileobj, fh)

    def get_meta(self, user_id, project_id):
        mf = os.path.join(self._project_dir(user_id, project_id), "meta.json")
        if os.path.exists(mf):
            try:
                return json.load(open(mf))
            except Exception:
                pass
        return {}

    def set_meta(self, user_id, project_id, meta):
        pd = self._project_dir(user_id, project_id)
        os.makedirs(pd, exist_ok=True)
        json.dump(meta, open(os.path.join(pd, "meta.json"), "w"))

    def work_dir(self, user_id, project_id):
        return self._project_dir(user_id, project_id)

    def sync_back(self, user_id, project_id, meta):
        self.set_meta(user_id, project_id, meta)

    def list_projects(self, user_id):
        out = []
        ud = self._user_dir(user_id)
        for pid in sorted(os.listdir(ud), reverse=True):
            pd = os.path.join(ud, pid)
            if not os.path.isdir(pd):
                continue
            meta = {"id": pid, "name": pid, "status": "uploaded",
                    "has_docx": False, "created": pid.split("_")[0]}
            meta.update(self.get_meta(user_id, pid))
            meta["has_docx"] = os.path.exists(os.path.join(pd, "output.docx"))
            out.append(meta)
        return out

    def download_path(self, user_id, project_id):
        fp = os.path.join(self._project_dir(user_id, project_id),
                          "output.docx")
        return fp if os.path.exists(fp) else None

    def download_url(self, user_id, project_id, response_disposition=None):
        return None  # local mode serves the file directly

    def delete_project(self, user_id, project_id):
        pd = self._project_dir(user_id, project_id)
        if os.path.isdir(pd):
            shutil.rmtree(pd)

    def logout_cleanup(self, user_id):
        ud = self._user_dir(user_id)
        n = 0
        for pid in os.listdir(ud):
            pd = os.path.join(ud, pid)
            if os.path.isdir(pd):
                shutil.rmtree(pd)
                n += 1
        return n


# ------------------------------------------------------------- firebase

class FirebaseStore:
    kind = "firebase"

    def __init__(self):
        info = json.loads(_service_account_json())
        try:
            firebase_admin.get_app()
        except ValueError:
            bucket_name = os.environ.get(
                "FIREBASE_STORAGE_BUCKET",
                f"{info.get('project_id')}.appspot.com")
            firebase_admin.initialize_app(
                fb_credentials.Certificate(info),
                {"storageBucket": bucket_name})
        self.bucket = fb_storage.bucket()
        self.db = fb_firestore.client()
        self._tmp = tempfile.mkdtemp(prefix="tex2word-")

    # -- helpers -------------------------------------------------------
    def _projects(self, user_id):
        return (self.db.collection("tex2word").document(user_id)
                .collection("projects"))

    def _blob(self, user_id, project_id, name):
        return self.bucket.blob(
            f"tex2word/{user_id}/{project_id}/{name}")

    def _touch_user(self, user_id):
        (self.db.collection("tex2word").document(user_id)
         .set({"updated": fb_firestore.SERVER_TIMESTAMP}, merge=True))

    # -- API used by app.py -------------------------------------------
    def save_upload(self, user_id, project_id, filename, fileobj):
        self._blob(user_id, project_id, filename).upload_from_file(
            fileobj, content_type="application/octet-stream")
        self._touch_user(user_id)

    def get_meta(self, user_id, project_id):
        snap = self._projects(user_id).document(project_id).get()
        return snap.to_dict() or {} if snap.exists else {}

    def set_meta(self, user_id, project_id, meta):
        self._projects(user_id).document(project_id).set(meta, merge=True)
        self._touch_user(user_id)

    def work_dir(self, user_id, project_id):
        d = os.path.join(self._tmp, user_id, project_id)
        os.makedirs(d, exist_ok=True)
        for name in ("upload.zip", "scan.pdf"):
            blob = self._blob(user_id, project_id, name)
            if blob.exists():
                blob.download_to_filename(os.path.join(d, name))
        return d

    def sync_back(self, user_id, project_id, meta):
        d = os.path.join(self._tmp, user_id, project_id)
        docx = os.path.join(d, "output.docx")
        if os.path.exists(docx):
            self._blob(user_id, project_id, "output.docx") \
                .upload_from_filename(
                    docx,
                    content_type="application/vnd.openxmlformats-officedocument"
                                 ".wordprocessingml.document")
            meta["has_docx"] = True
        else:
            meta["has_docx"] = False
        self.set_meta(user_id, project_id, meta)

    def list_projects(self, user_id):
        out = []
        for snap in self._projects(user_id).stream():
            meta = snap.to_dict() or {}
            meta.setdefault("id", snap.id)
            meta.setdefault("name", snap.id)
            out.append(meta)
        out.sort(key=lambda m: m.get("id", ""), reverse=True)
        return out

    def download_path(self, user_id, project_id):
        return None  # firebase mode redirects to a signed URL

    def download_url(self, user_id, project_id, response_disposition=None):
        from datetime import timedelta
        blob = self._blob(user_id, project_id, "output.docx")
        if not blob.exists():
            return None
        kw = {"expiration": timedelta(hours=1)}
        if response_disposition:
            kw["response_disposition"] = response_disposition
        return blob.generate_signed_url(**kw)

    def delete_project(self, user_id, project_id):
        prefix = f"tex2word/{user_id}/{project_id}/"
        for blob in self.bucket.list_blobs(prefix=prefix):
            blob.delete()
        self._projects(user_id).document(project_id).delete()
        shutil.rmtree(os.path.join(self._tmp, user_id, project_id),
                      ignore_errors=True)

    def logout_cleanup(self, user_id):
        ids = [s.id for s in self._projects(user_id).stream()]
        for pid in ids:
            self.delete_project(user_id, pid)
        return len(ids)


def get_store():
    if USE_FIREBASE:
        return FirebaseStore()
    base = os.path.dirname(os.path.abspath(__file__))
    return LocalStore(os.path.join(base, "storage"))
