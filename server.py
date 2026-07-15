#!/usr/bin/env python3
"""LAN Share Hub — Cross-platform LAN file sharing & messaging with terminal aesthetic.
   Run: python3 server.py [port] [serve_dir]
   Open http://<your-ip>:8888 on any device on the same WiFi."""

import http.server, socketserver, os, sys, json, sqlite3, urllib.parse, mimetypes, socket, subprocess, zipfile, tempfile, email.utils, hashlib, html, xml.etree.ElementTree as ET
from pathlib import Path

def _hub_path(env_key, default):
    v = os.environ.get(env_key)
    p = v if v else default
    return Path(os.path.expanduser(p)).resolve()

PORT            = int(os.environ.get("HUB_PORT") or (sys.argv[1] if len(sys.argv) > 1 else 8888))
BIND            = os.environ.get("HUB_BIND", "0.0.0.0")
SERVE_DIR       = _hub_path("HUB_SERVE_DIR", sys.argv[2] if len(sys.argv) > 2 else "~/")
UPLOAD_DIR      = _hub_path("HUB_UPLOAD_DIR", "~/lan-share/inbox/")
DB_PATH         = _hub_path("HUB_DB_PATH", "~/.config/lan-share/messages.db")
HUB_PUBLIC_URL  = os.environ.get("HUB_PUBLIC_URL", "").rstrip("/")
HUB_MAX_UPLOAD  = int(os.environ.get("HUB_MAX_UPLOAD_MB", "0")) * 1024 * 1024
HUB_MODE        = "cloud" if HUB_PUBLIC_URL else "lan"
DB_DIR          = DB_PATH.parent
DB_DIR.mkdir(parents=True, exist_ok=True)
THUMB_CACHE     = Path(os.environ.get("HUB_THUMB_CACHE", str(DB_DIR / "thumb-cache"))).resolve()
THUMB_CACHE.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
SERVE_DIR.mkdir(parents=True, exist_ok=True)
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
try:
    UPLOAD_REL = UPLOAD_DIR.resolve().relative_to(SERVE_DIR.resolve()).as_posix()
except ValueError:
    UPLOAD_REL = "from-phone"

def dl_url(relpath):
    rel = relpath.replace("\\", "/").strip("/")
    return "/api/download/" + "/".join(urllib.parse.quote(p) for p in rel.split("/") if p)

def resolve_download(rel):
    rel = urllib.parse.unquote(rel).lstrip("/")
    if not rel:
        return None
    if rel.startswith("downloaddev/"):
        rel = "download/dev/" + rel[len("downloaddev/"):]
    fp = (SERVE_DIR / rel).resolve()
    try:
        fp.relative_to(SERVE_DIR)
        if fp.is_file():
            return fp
    except ValueError:
        pass
    if rel.startswith("from-phone/"):
        fp = (UPLOAD_DIR / rel[11:]).resolve()
        try:
            fp.relative_to(UPLOAD_DIR)
            if fp.is_file():
                return fp
        except ValueError:
            pass
    # Legacy local paths (下载/from-phone) and cross-mode aliases
    for prefix in ("下载/from-phone/", "from-phone/"):
        if rel.startswith(prefix):
            sub = rel[len(prefix):]
            fp = (UPLOAD_DIR / sub).resolve()
            try:
                fp.relative_to(UPLOAD_DIR)
                if fp.is_file():
                    return fp
            except ValueError:
                pass
    return None

def file_etag(fp):
    st = fp.stat()
    return '"{:x}-{:x}"'.format(int(st.st_mtime), st.st_size)

def content_disposition(disp, filename):
    safe = filename.replace("\\", "_").replace('"', "'")
    ascii_name = safe.encode("ascii", "ignore").decode() or "download"
    utf8_name = urllib.parse.quote(safe)
    return '{}; filename="{}"; filename*=UTF-8\'\'{}'.format(disp, ascii_name, utf8_name)

def is_raster_image(fp):
    return fp.suffix.lower() in (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp")

def thumb_path_for(fp, width):
    key = hashlib.sha1("{}:{}:{}".format(fp.resolve(), fp.stat().st_mtime, width).encode()).hexdigest()
    return THUMB_CACHE / "{}_{}.jpg".format(key, width)

def generate_thumb(src, dst, width):
    try:
        from PIL import Image
    except ImportError:
        raise RuntimeError("Pillow not installed")
    with Image.open(src) as im:
        if im.mode in ("RGBA", "P"):
            im = im.convert("RGB")
        elif im.mode != "RGB":
            im = im.convert("RGB")
        im.thumbnail((width, max(width * 3, 480)), Image.Resampling.LANCZOS)
        dst.parent.mkdir(parents=True, exist_ok=True)
        im.save(dst, "JPEG", quality=82, optimize=True)

_WNS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_VNS = "{urn:schemas-microsoft-com:vml}"
_RELS_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"

def _load_docx_image_rels(zf):
    rels = {}
    try:
        root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        for rel in root:
            rid = rel.get("Id")
            target = rel.get("Target")
            typ = rel.get("Type") or ""
            if rid and target and "image" in typ:
                path = target if target.startswith("word/") else "word/" + target.lstrip("/")
                rels[rid] = path
    except KeyError:
        pass
    return rels

def _docx_img_html(rid, rels, media_base):
    path = rels.get(rid)
    if not path:
        return ""
    name = path.split("/")[-1]
    src = media_base + urllib.parse.quote(name)
    return '<figure class="docx-img"><img src="{}" loading="lazy" decoding="async" alt=""></figure>'.format(src)

def _docx_blip_html(blip, rels, media_base):
    rid = blip.get(_RNS + "embed")
    if rid:
        return _docx_img_html(rid, rels, media_base)
    return ""

def _render_docx_para(p, rels, media_base):
    chunks = []
    for node in p:
        if node.tag == _WNS + "r":
            for blip in node.iter(_A_NS + "blip"):
                tag = _docx_blip_html(blip, rels, media_base)
                if tag:
                    chunks.append(("img", tag))
            for im in node.iter(_VNS + "imagedata"):
                rid = im.get(_RNS + "id") or im.get("id")
                if rid:
                    tag = _docx_img_html(rid, rels, media_base)
                    if tag:
                        chunks.append(("img", tag))
            texts = []
            bold = False
            rpr = node.find(_WNS + "rPr")
            if rpr is not None and rpr.find(_WNS + "b") is not None:
                bold = True
            t = node.find(_WNS + "t")
            if t is not None:
                if t.text:
                    texts.append(t.text)
                if t.tail:
                    texts.append(t.tail)
            line = "".join(texts)
            if line:
                chunks.append(("text", line, bold))
        elif node.tag in (_WNS + "drawing", _WNS + "pict"):
            for blip in node.iter(_A_NS + "blip"):
                tag = _docx_blip_html(blip, rels, media_base)
                if tag:
                    chunks.append(("img", tag))
            for im in node.iter(_VNS + "imagedata"):
                rid = im.get(_RNS + "id") or im.get("id")
                if rid:
                    tag = _docx_img_html(rid, rels, media_base)
                    if tag:
                        chunks.append(("img", tag))
    if not chunks:
        return "<p>&nbsp;</p>"
    tag = "p"
    ppr = p.find(_WNS + "pPr")
    if ppr is not None:
        ps = ppr.find(_WNS + "pStyle")
        if ps is not None:
            val = ps.get(_WNS + "val") or ps.get("val") or ""
            if "Heading" in val or val.startswith("heading") or val.startswith("Title"):
                tag = "h2"
    out = []
    text_buf = []
    bold_any = False
    def flush_text():
        nonlocal text_buf, bold_any
        line = "".join(text_buf).strip()
        text_buf = []
        if not line:
            return
        esc = html.escape(line)
        if tag == "h2":
            out.append("<h2>" + esc + "</h2>")
        elif bold_any:
            out.append("<p><strong>" + esc + "</strong></p>")
        else:
            out.append("<p>" + esc + "</p>")
        bold_any = False
    for ch in chunks:
        if ch[0] == "img":
            flush_text()
            out.append(ch[1])
        else:
            text_buf.append(ch[1])
            if ch[2]:
                bold_any = True
    flush_text()
    return "".join(out) if out else "<p>&nbsp;</p>"

def docx_to_html(fp, media_base):
    with zipfile.ZipFile(str(fp)) as zf:
        try:
            xml = zf.read("word/document.xml")
        except KeyError:
            raise ValueError("invalid docx")
        rels = _load_docx_image_rels(zf)
        root = ET.fromstring(xml)
        body_el = root.find(_WNS + "body")
        if body_el is None:
            raise ValueError("invalid docx")
        out = []
        for p in body_el.findall(_WNS + "p"):
            out.append(_render_docx_para(p, rels, media_base))
    body = "".join(out)
    if not body.strip():
        raise ValueError("empty document")
    return body

def docx_media_base(rel):
    rel = rel.strip().lstrip("/")
    return "/api/preview/docx-media/" + "/".join(urllib.parse.quote(p) for p in rel.split("/") if p) + "/"

mimetypes.add_type("image/jpeg", ".jpg")
mimetypes.add_type("image/heic", ".heic")
mimetypes.add_type("image/heif", ".heif")
mimetypes.add_type("video/mp4", ".mp4")
mimetypes.add_type("video/quicktime", ".mov")
mimetypes.add_type("video/x-m4v", ".m4v")
mimetypes.add_type("video/webm", ".webm")
mimetypes.add_type("video/3gpp", ".3gp")

import time as _time
from datetime import datetime

def zip_paths_to_buffer(paths, max_files=200):
    paths = paths if isinstance(paths, list) else []
    buf = tempfile.SpooledTemporaryFile(max_size=128 * 1024 * 1024)
    added = 0
    seen = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in paths[:max_files]:
            if not isinstance(rel, str):
                continue
            rel = rel.strip().lstrip("/")
            if not rel or rel in seen:
                continue
            fp = resolve_download(rel)
            if not fp:
                continue
            seen.add(rel)
            added += 1
            zf.write(fp, fp.name)
    if not added:
        return None, 0
    buf.seek(0)
    return buf.read(), added
HUB_BUILD = str(int(_time.time()))
HUB_BRAND = {
    "lan": {"title": "Relay Local", "sub": "局域网传输站", "badge": "LOCAL", "page": "Relay Local · 局域网传输站"},
    "cloud": {"title": "Relay Cloud", "sub": "云传输站", "badge": "CLOUD", "page": "Relay Cloud · 云传输站"},
}
_brand = HUB_BRAND[HUB_MODE]

db = sqlite3.connect(str(DB_PATH), check_same_thread=False)
db.execute("CREATE TABLE IF NOT EXISTS messages(id INTEGER PRIMARY KEY AUTOINCREMENT,sender TEXT NOT NULL,sender_name TEXT DEFAULT'',content TEXT NOT NULL,created_at TEXT DEFAULT(datetime('now','localtime')))")
db.commit()
msg_epoch = 0

def bump_msg_epoch():
    global msg_epoch
    msg_epoch += 1

def db_msgs(since=None, before=None, limit=500):
    limit = max(1, min(int(limit), 2000))
    if since is not None:
        rows = db.execute("SELECT id,sender,sender_name,content,created_at FROM messages WHERE id>? ORDER BY id LIMIT ?",(int(since),limit)).fetchall()
    elif before is not None:
        rows = db.execute("SELECT id,sender,sender_name,content,created_at FROM messages WHERE id<? ORDER BY id DESC LIMIT ?",(int(before),limit)).fetchall()[::-1]
    else:
        rows = db.execute("SELECT id,sender,sender_name,content,created_at FROM messages ORDER BY id DESC LIMIT ?",(limit,)).fetchall()[::-1]
    return [{"id":r[0],"sender":r[1],"sender_name":r[2],"content":r[3],"time":r[4]} for r in rows]

def db_msgs_anchor(anchor_id, radius=120):
    aid = int(anchor_id)
    radius = max(20, min(int(radius), 400))
    lo = max(1, aid - radius)
    hi = aid + radius
    rows = db.execute(
        "SELECT id,sender,sender_name,content,created_at FROM messages WHERE id BETWEEN ? AND ? ORDER BY id",
        (lo, hi),
    ).fetchall()
    has_more_before = lo > 1
    has_more_after = db.execute("SELECT 1 FROM messages WHERE id>? LIMIT 1", (hi,)).fetchone() is not None
    oldest = rows[0][0] if rows else 0
    msgs = [{"id":r[0],"sender":r[1],"sender_name":r[2],"content":r[3],"time":r[4]} for r in rows]
    return msgs, has_more_before, has_more_after, oldest

def db_msg_count():
    return db.execute("SELECT COUNT(*) FROM messages").fetchone()[0]

def _like_pattern(q):
    q = (q or "").strip()
    if not q:
        return ""
    esc = q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return "%{}%".format(esc)

def _msg_snippet(content, q, max_len=140):
    content = content or ""
    q = (q or "").strip()
    if not content:
        return ""
    low = content.lower()
    idx = low.find(q.lower()) if q else -1
    if idx < 0:
        s = content.replace("\n", " ")
        return (s[:max_len] + "…") if len(s) > max_len else s
    start = max(0, idx - 50)
    end = min(len(content), idx + len(q) + 70)
    s = content[start:end].replace("\n", " ")
    if start > 0:
        s = "…" + s
    if end < len(content):
        s = s + "…"
    return s

def db_search(q, limit=80):
    pat = _like_pattern(q)
    if not pat:
        return []
    limit = max(1, min(int(limit), 200))
    rows = db.execute(
        "SELECT id,sender,sender_name,content,created_at FROM messages WHERE content LIKE ? ESCAPE '\\' ORDER BY id DESC LIMIT ?",
        (pat, limit),
    ).fetchall()
    out = []
    for r in rows:
        content = r[3] or ""
        if content.startswith("[["):
            kind = "file"
        else:
            kind = "text"
        out.append({
            "id": r[0],
            "sender": r[1],
            "sender_name": r[2],
            "content": content,
            "time": r[4],
            "kind": kind,
            "snippet": _msg_snippet(content, q),
        })
    return out

_SEARCH_SKIP_DIRS = {
    ".git", "node_modules", "__pycache__", ".cache", "thumb-cache", ".hermes",
    ".venv", "venv", ".npm", ".cargo", "target", ".local/share/Trash",
}

def search_files(q, limit=80, max_scan=80000):
    q = (q or "").strip()
    if not q:
        return [], False
    q_lower = q.lower()
    limit = max(1, min(int(limit), 200))
    root = UPLOAD_DIR if HUB_MODE == "cloud" else SERVE_DIR
    results = []
    scanned = 0
    truncated = False

    def add_file(fp):
        try:
            st = fp.stat()
            if HUB_MODE == "cloud":
                try:
                    rp = fp.resolve().relative_to(UPLOAD_DIR.resolve()).as_posix()
                    rp = UPLOAD_REL + "/" + rp
                except ValueError:
                    return
            else:
                try:
                    rp = fp.resolve().relative_to(SERVE_DIR.resolve()).as_posix()
                except ValueError:
                    return
            results.append({
                "name": fp.name,
                "relpath": rp,
                "size": st.st_size,
                "mtime": int(st.st_mtime),
            })
        except OSError:
            pass

    def walk_dir(d, depth=0):
        nonlocal scanned, truncated
        if len(results) >= limit or scanned >= max_scan:
            truncated = True
            return
        try:
            entries = list(d.iterdir())
        except OSError:
            return
        entries.sort(key=lambda x: (not x.is_dir(), x.name.lower()))
        for f in entries:
            if len(results) >= limit or scanned >= max_scan:
                truncated = True
                return
            name = f.name
            if name.startswith(".") and f.is_dir():
                continue
            if f.is_dir():
                if name in _SEARCH_SKIP_DIRS:
                    continue
                if depth >= 14:
                    continue
                walk_dir(f, depth + 1)
            elif f.is_file():
                scanned += 1
                if q_lower in name.lower():
                    add_file(f)

    if HUB_MODE == "cloud":
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        walk_dir(UPLOAD_DIR, 0)
    else:
        walk_dir(SERVE_DIR, 0)
    return results, truncated

def db_add(sender, content, sender_name=""):
    cur = db.execute("INSERT INTO messages(sender,sender_name,content)VALUES(?,?,?)",(sender,sender_name,content))
    db.commit()
    bump_msg_epoch()
    return cur.lastrowid

def db_delete_messages(ids):
    clean = []
    for i in ids:
        try:
            clean.append(int(i))
        except (TypeError, ValueError):
            continue
    if not clean:
        return 0
    q = "DELETE FROM messages WHERE id IN ({})".format(",".join("?" * len(clean)))
    cur = db.execute(q, clean)
    db.commit()
    if cur.rowcount:
        bump_msg_epoch()
    return cur.rowcount

def is_in_upload_dir(fp):
    try:
        fp.resolve().relative_to(UPLOAD_DIR.resolve())
        return True
    except ValueError:
        return False

def deletable_file(rel):
    fp = resolve_download(rel)
    if not fp or not fp.is_file():
        return None
    if HUB_MODE == "cloud":
        return fp if is_in_upload_dir(fp) else None
    try:
        fp.resolve().relative_to(SERVE_DIR.resolve())
    except ValueError:
        return None
    return fp

def related_message_ids_for_file(fp):
    if not is_in_upload_dir(fp):
        return []
    name = fp.name
    ids = []
    for rid, content in db.execute("SELECT id, content FROM messages"):
        head = (content or "").split("\n", 1)[0]
        if head.startswith("[[") and name in head:
            ids.append(rid)
    return ids

def upload_path_for_name(name):
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    sn = os.path.basename(name or "")
    if not sn:
        return None
    fp = UPLOAD_DIR / sn
    base, ext = os.path.splitext(sn)
    c = 1
    while fp.exists():
        fp = UPLOAD_DIR / "{}_{}{}".format(base, c, ext)
        c += 1
    return fp

def parse_multipart_filename(header_bytes):
    hdr = header_bytes.decode("utf-8", errors="ignore")
    for ln in hdr.split("\r\n"):
        if "filename=" in ln:
            fn = ln.split("filename=", 1)[1].strip().strip('"')
            return os.path.basename(fn) if fn else None
    return None

def stream_multipart_upload(rfile, boundary, content_length):
    bd = boundary.encode("ascii", errors="ignore")
    delim = b"\r\n--" + bd
    keep = len(delim) + 8
    read = 0
    pending = b""
    out = None
    out_path = None
    in_body = False
    written = 0

    def abort():
        if out:
            try:
                out.close()
            except OSError:
                pass
        if out_path and out_path.exists():
            try:
                out_path.unlink()
            except OSError:
                pass

    try:
        while read < content_length:
            chunk = rfile.read(min(1048576, content_length - read))
            if not chunk:
                break
            read += len(chunk)
            data = pending + chunk
            pending = b""

            if not in_body:
                sep = data.find(b"\r\n\r\n")
                if sep < 0:
                    pending = data[-32768:] if len(data) > 32768 else data
                    continue
                fn = parse_multipart_filename(data[:sep])
                if not fn or fn.startswith("."):
                    abort()
                    return None, 0, read
                out_path = upload_path_for_name(fn)
                if not out_path:
                    abort()
                    return None, 0, read
                out = open(out_path, "wb")
                in_body = True
                data = data[sep + 4:]

            if in_body and out is not None:
                idx = data.find(delim)
                if idx >= 0:
                    out.write(data[:idx])
                    written += idx
                    out.close()
                    out = None
                    return out_path, written, read
                if len(data) > keep:
                    out.write(data[:-keep])
                    written += len(data) - keep
                    pending = data[-keep:]
                else:
                    pending = data

        if in_body and out is not None and pending:
            idx = pending.find(delim)
            if idx >= 0:
                out.write(pending[:idx])
                written += idx
                out.close()
                out = None
                return out_path, written, read

        abort()
        return None, written, read
    except Exception:
        abort()
        raise

def lan_ips():
    ips = []
    try:
        out = subprocess.check_output(["hostname", "-I"], text=True, timeout=2)
        for ip in out.strip().split():
            if ip.startswith("192.168.") or ip.startswith("10."):
                ips.append(ip)
    except Exception:
        pass
    if not ips:
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.settimeout(1)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            if ip and not ip.startswith(("127.", "172.", "198.")):
                ips.append(ip)
        except Exception:
            pass
    return ips or ["127.0.0.1"]

def lan_ip():
    return lan_ips()[0]

def client_is_hub_host(peer):
    if not peer:
        return False
    if peer in ("127.0.0.1", "::1"):
        return True
    if peer.startswith("::ffff:"):
        peer = peer[7:]
    if peer == lan_ip():
        return True
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            if info[4][0] == peer:
                return True
    except Exception:
        pass
    return False

SVG = {
    "files":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',
    "upload":  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>',
    "chat":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>',
    "folder":  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    "download":'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>',
    "image":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>',
    "video":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><polygon points="23 7 16 12 23 17 23 7"/><rect x="1" y="5" width="15" height="14" rx="2"/></svg>',
    "audio":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/></svg>',
    "file":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><polyline points="13 2 13 9 20 9"/></svg>',
    "home":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>',
    "attach":  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg>',
    "send":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>',
    "copy":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>',
    "x":       '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>',
    "check":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><polyline points="20 6 9 17 4 12"/></svg>',
    "inbox":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/></svg>',
    "photo":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>',
    "speaker": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><polygon points="11 5 6 9 2 9 2 15 6 15 11 19 11 5"/><path d="M19.07 4.93a10 10 0 0 1 0 14.14M15.54 8.46a5 5 0 0 1 0 7.07"/></svg>',
    "link":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>',
    "grid":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>',
    "logo":    '<svg viewBox="0 0 24 24" fill="none"><rect x="1" y="1" width="22" height="22" rx="5" fill="#0a0a0c" stroke="#d4b84a" stroke-width="1.2"/><path d="M12 5.5l5.5 3.2v6.6L12 18.5 6.5 15.3V8.7L12 5.5z" stroke="#d4b84a" stroke-width="1.1" fill="none"/><circle cx="12" cy="12" r="2.2" fill="#d4b84a"/></svg>',
    "bell":    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>',
    "batch":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/></svg>',
    "archive": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M4 4h16v4H4z"/><path d="M6 8v12h12V8"/><path d="M10 12h4"/><path d="M12 12v4"/></svg>',
    "trash":   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>',
    "search":  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
}

def si(n, s=18, c=""):
    return '<span class="svg-icon %s" style="width:%dpx;height:%dpx">%s</span>' % (c, s, s, SVG.get(n, SVG["file"]))

# ════════════════ HTML ════════════════
T = r'''<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="theme-color" content="#0a0a0c">
<link rel="icon" href="/icon.svg" type="image/svg+xml">
<link rel="apple-touch-icon" href="/icon.svg">
<title>__PAGE_TITLE__</title>
<style>
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#070709;--s1:#0c0c10;--s2:#121218;--s3:#181820;--panel:#0f0f14;
  --bd:#3d3a22;--bd2:#2a2818;--am:#d4b84a;--am-dim:#7a7038;--ad:#9a9450;
  --ag:rgba(212,184,74,.06);--gn:#52b046;--gn-glow:rgba(82,176,70,.5);--rd:#d85646;
  --tx:#e8dc9a;--t2:#8a8440;--t3:#5a5828;--steel:#222228;
  --fm:"JetBrains Mono","Cascadia Code","SF Mono",Menlo,ui-monospace,monospace;
  --fs:-apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif;
  --r:0px;--inset:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -2px 0 rgba(0,0,0,.55);
  --bevel:0 1px 0 rgba(255,255,255,.07),0 3px 0 rgba(0,0,0,.45);
  --press-ease:cubic-bezier(.34,1.45,.64,1);
  --act-glow:rgba(212,184,74,.55);
}
html,body{height:100%;overflow:hidden}
body{
  background:var(--bg);color:var(--ad);font-family:var(--fs);min-height:100dvh;display:flex;flex-direction:column;
  background-image:linear-gradient(rgba(212,184,74,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(212,184,74,.025) 1px,transparent 1px);
  background-size:20px 20px;
}
body::after{content:'';pointer-events:none;position:fixed;inset:0;z-index:9998;
  background:repeating-linear-gradient(0deg,transparent,transparent 2px,rgba(0,0,0,.07) 2px,rgba(0,0,0,.07) 4px);opacity:.12}
@media(max-width:767px){body::after{display:none}}
.svg-icon{display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;vertical-align:middle}
.svg-icon svg{width:100%;height:100%}

/* ── Press / tap feedback (all interactive controls) ── */
button,.lan-bar-url,.fr:not(.dir),.yg-item,.sel-chip{
  -webkit-tap-highlight-color:transparent;
  touch-action:manipulation;
}
button,.lan-bar-url,.deck-btn,.tab-btn,.cat-btn,.sort-btn,.qb,.btn,.ibtn,.cs,.sel-x,.sel-dl,.sel-del,.batch-btn,.jump-bottom,.pv-arrow,.pv-edge,.expand-btn,.mcp,.preview-actions button,#mdl button{
  transition:transform .07s var(--press-ease),box-shadow .07s,border-color .1s,background .1s,color .1s,filter .08s;
  will-change:transform;
}
@keyframes act-ring{
  0%{box-shadow:0 0 0 0 var(--act-glow),var(--bevel)}
  65%{box-shadow:0 0 0 11px rgba(212,184,74,0),var(--bevel)}
  100%{box-shadow:0 0 0 0 rgba(212,184,74,0),var(--bevel)}
}
@keyframes act-ring-cloud{
  0%{box-shadow:0 0 0 0 rgba(110,181,255,.5),var(--bevel)}
  65%{box-shadow:0 0 0 11px rgba(110,181,255,0),var(--bevel)}
  100%{box-shadow:0 0 0 0 rgba(110,181,255,0),var(--bevel)}
}
.act-hit{animation:act-ring .42s ease-out}
body.hub-cloud .act-hit{animation:act-ring-cloud .42s ease-out}
.act-press,.deck-btn:active,.cp-btn:active,.sfx-btn:active,.tab-btn:active,.cat-btn:active,.sort-btn:active,.qb:active,.btn:active:not(:disabled),.ibtn:active,.sel-x:active,.sel-dl:not(:disabled):active,.sel-del:not(:disabled):active,.lan-bar-url:active,.batch-btn:active,.jump-bottom:active,.pv-arrow:active,.pv-edge:active,.expand-btn:active,.mcp:active,.preview-actions button:active,#mdl button:active,.mo-nav button:active{
  transform:scale(.93) translateY(1px)!important;
  filter:brightness(1.14);
}
.tab-btn.act-press,.tab-btn:active{background:var(--s2)!important}
.tab-btn.active.act-press,.tab-btn.active:active{transform:scale(.97) translateY(1px)!important;background:var(--bg)!important}
.cs.act-press,.cs:active,.btn-pri.act-press,.btn-pri:active{
  transform:scale(.9) translateY(2px)!important;
  filter:brightness(1.2);
  box-shadow:inset 0 3px 10px rgba(0,0,0,.5),0 0 18px rgba(212,184,74,.35)!important;
}
.fr:not(.dir).act-press,.fr:not(.dir):active{background:var(--s2)!important;box-shadow:inset 3px 0 0 var(--am)!important;transform:scale(.992)!important}
.yg-item.act-press,.yg-item:active{transform:scale(.94)!important;border-color:var(--am)!important;box-shadow:0 0 14px rgba(212,184,74,.2)!important}
.sel-chip.act-press,.sel-chip:active{transform:scale(.88)!important;border-color:var(--am)!important}
@media (prefers-reduced-motion:reduce){
  .act-hit{animation:none!important}
  button:active,.act-press{transform:none!important;filter:none!important}
}

/* Scrollbar — visible on mobile */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--s1)}
::-webkit-scrollbar-thumb{background:var(--am-dim);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:var(--am)}
*{scrollbar-width:thin;scrollbar-color:var(--am-dim) var(--s1)}
@media(max-width:767px){
  ::-webkit-scrollbar{width:8px;height:8px}
  ::-webkit-scrollbar-thumb{background:var(--am-dim);min-height:40px}
  .ca,.fl,.year-gallery,.year-strip{scrollbar-width:auto}
}

/* Header — minimal */
header.deck-h{
  padding:0;border-bottom:1px solid var(--bd2);display:flex;justify-content:space-between;align-items:center;
  font-family:var(--fs);font-size:13px;color:var(--tx);
  background:var(--s1);position:sticky;top:0;z-index:50;
}
.deck-left,.deck-right{display:flex;align-items:center;gap:10px;padding:14px 16px}
.deck-right{margin-left:auto;gap:8px}
header .tid{color:var(--am);font-size:15px;font-weight:600;letter-spacing:.04em}
header .brand-mark{display:inline-flex;align-items:center;margin-right:6px}
header .brand-mark .svg-icon{width:28px!important;height:28px!important}
.deck-btn,.cp-btn,.sfx-btn{background:transparent;border:1px solid var(--bd2);color:var(--ad);font-size:11px;padding:7px 10px;cursor:pointer;font-family:var(--fs);transition:all .15s;display:inline-flex;align-items:center;gap:4px;min-width:36px;min-height:36px;justify-content:center;-webkit-tap-highlight-color:transparent}
.deck-btn .svg-icon,.cp-btn .svg-icon,.sfx-btn .svg-icon{width:18px!important;height:18px!important}
.deck-btn:hover,.cp-btn:hover,.sfx-btn:hover{border-color:var(--am);color:var(--am);box-shadow:0 0 12px rgba(212,184,74,.15)}
.sfx-btn.on{color:var(--gn);border-color:rgba(82,176,70,.45)}
.sel-toggle-btn.on{color:var(--am);border-color:var(--am);background:rgba(212,184,74,.1);box-shadow:0 0 16px rgba(212,184,74,.22)}
body.sel-active header.deck-h{border-bottom-color:var(--am);box-shadow:0 2px 20px rgba(212,184,74,.12)}
body.hub-cloud .lan-bar{display:none!important}
body.hub-lan .cloud-bar{display:none!important}
header .hub-en{font-size:11px;color:var(--t3);font-weight:400;margin-left:6px;letter-spacing:.02em}
.mode-badge{font-size:9px;font-family:var(--fm);padding:2px 7px;border:1px solid;margin-left:8px;letter-spacing:.1em;vertical-align:middle}
.mode-badge.mode-lan{color:var(--am);border-color:var(--am-dim);background:rgba(212,184,74,.06)}
.mode-badge.mode-cloud{color:#6eb5ff;border-color:rgba(110,181,255,.35);background:rgba(110,181,255,.06)}
.cloud-bar{display:none;align-items:center;gap:10px;padding:10px 16px;background:rgba(110,181,255,.06);border-bottom:1px solid rgba(110,181,255,.2);font-family:var(--fs);flex-shrink:0}
.cloud-bar.show{display:flex}
.cloud-bar-label{font-size:12px;color:var(--ad);white-space:nowrap;flex-shrink:0}
.cloud-bar-url{font-family:var(--fm);font-size:13px;color:#6eb5ff;flex-shrink:0}
.cloud-bar-hint{font-size:11px;color:var(--ad);white-space:nowrap;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis}
.url-hidden{position:absolute;left:-9999px;opacity:0;pointer-events:none}
.lan-bar{display:none;align-items:center;gap:10px;padding:10px 16px;background:rgba(212,184,74,.07);border-bottom:1px solid var(--bd2);font-family:var(--fs);flex-shrink:0}
.lan-bar.show{display:flex}
.lan-bar-label{font-size:12px;color:var(--ad);white-space:nowrap;flex-shrink:0}
.lan-bar-url{flex:1;min-width:0;text-align:left;background:var(--s2);border:1px solid rgba(212,184,74,.35);color:var(--am);padding:8px 12px;font-family:var(--fm);font-size:13px;cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.lan-bar-url:hover{border-color:var(--am);box-shadow:0 0 14px rgba(212,184,74,.12)}
.lan-bar-hint{font-size:11px;color:var(--ad);white-space:nowrap;flex-shrink:0}

/* Rack chassis */
.chassis{flex:1;min-height:0;display:flex;flex-direction:column;margin:0;border:none;background:var(--bg);box-shadow:none;overflow:hidden}
.desktop .chassis{margin:0;border:none}

/* Tabs */
nav.tabs{display:flex;background:var(--s1);border-bottom:1px solid var(--bd2);font-family:var(--fs);gap:0;flex-shrink:0;margin:0}
.tab-btn{flex:1;padding:14px 8px;border:none;border-right:1px solid var(--bd2);background:var(--s1);color:var(--ad);cursor:pointer;display:flex;align-items:center;justify-content:center;gap:7px;transition:all .12s;font-size:14px;font-weight:500;min-height:50px}
.tab-btn .svg-icon{width:18px!important;height:18px!important;color:var(--ad)}
.tab-btn:last-child{border-right:none}
.tab-btn.active{background:var(--bg);color:var(--am);box-shadow:inset 0 -2px 0 var(--am);font-weight:600}
.tab-btn.active .svg-icon{color:var(--am)}
.tab-btn:hover:not(.active){background:var(--s2);color:var(--tx)}
.tab-badge{background:var(--gn);color:#070709;font-size:9px;min-width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-weight:700;font-family:var(--fm);margin-left:4px;box-shadow:0 0 6px var(--gn-glow)}
.tab-badge:empty{display:none}

/* Desktop drag-drop upload */
.drop-overlay{display:none;position:fixed;inset:0;z-index:150;background:rgba(10,10,12,.88);border:3px dashed var(--am);align-items:center;justify-content:center;flex-direction:column;gap:10px;pointer-events:none}
.drop-overlay.active{display:flex}
.drop-overlay .drop-title{color:var(--am);font-size:20px;font-family:var(--fs);font-weight:600;letter-spacing:.02em}
.drop-overlay .drop-sub{color:var(--ad);font-size:13px;font-family:var(--fs)}

/* Main — desktop: side-by-side */
main{flex:1;min-height:0;display:flex;overflow:hidden;background:var(--bg)}
.tab-panel{display:none;flex:1;min-height:0;overflow:hidden}
.tab-panel.active{display:flex;flex-direction:column}
.main-layout{display:flex;flex:1;min-height:0;overflow:hidden}
.main-left{width:300px;min-width:200px;max-width:45%;border-right:1px solid var(--bd);display:flex;flex-direction:column;overflow:hidden;min-height:0;transition:width .2s}
.main-right{flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden;background:var(--bg)}
.preview-pane{flex:1;min-height:0;display:flex;align-items:flex-start;justify-content:flex-start;overflow:auto;padding:16px}
.preview-wrap{flex:1;min-height:0;position:relative;display:flex;align-items:stretch}
.pv-arrow{position:absolute;top:50%;transform:translateY(-50%);z-index:20;background:rgba(10,10,12,.85);border:1px solid var(--bd);color:var(--am);width:32px;height:44px;cursor:pointer;font-size:22px;line-height:1;font-family:var(--fm);display:flex;align-items:center;justify-content:center}
.pv-arrow:active{background:var(--s2);border-color:var(--am)}
.pv-prev{left:4px}.pv-next{right:4px}
.pv-count{position:absolute;bottom:8px;left:50%;transform:translateX(-50%);font-size:9px;color:var(--ad);font-family:var(--fm);z-index:20;background:rgba(10,10,12,.7);padding:2px 8px;border:1px solid var(--bd)}
.mo-nav{display:flex;align-items:center;gap:6px;flex:1;justify-content:center}
.mo-nav button{background:none;border:1px solid var(--bd);color:var(--ad);width:32px;height:28px;cursor:pointer;font-size:18px;font-family:var(--fm)}
.mo-nav button:active{border-color:var(--am);color:var(--am)}
.preview-pane .txt-preview{max-width:700px;width:100%;font-family:var(--fm);font-size:11px;line-height:1.6;color:var(--tx);white-space:pre-wrap;word-break:break-word}
.preview-pane .empty{color:var(--t3);font-family:var(--fs);font-size:13px;text-align:center;letter-spacing:.02em;padding:24px;line-height:1.6}
.preview-pane .empty b{display:block;color:var(--am);font-size:15px;margin-bottom:8px}
.preview-pane img{max-width:100%;max-height:70vh;object-fit:contain;border:1px solid var(--bd)}
.preview-pane video{max-width:100%;max-height:70vh;outline:none}
.preview-pane audio{width:300px}
.preview-pane .md-preview{max-width:600px;width:100%;font-family:var(--fs);font-size:13px;line-height:1.7;color:var(--tx);padding:16px}
.preview-pane .md-preview h1,.preview-pane .md-preview h2,.preview-pane .md-preview h3{color:var(--am);font-family:var(--fm);margin:12px 0 6px;letter-spacing:.03em}
.preview-pane .md-preview h1{font-size:18px}.preview-pane .md-preview h2{font-size:15px}.preview-pane .md-preview h3{font-size:13px}
.preview-pane .md-preview code{background:var(--s2);padding:1px 5px;font-family:var(--fm);font-size:11px;color:var(--am)}
.preview-pane .md-preview pre{background:var(--s1);border:1px solid var(--bd);padding:10px;overflow-x:auto;font-family:var(--fm);font-size:10px;line-height:1.5;margin:8px 0}
.preview-pane .md-preview a{color:var(--am);text-decoration:underline}
.preview-pane .md-preview li{margin:2px 0 2px 18px}
.preview-pane .md-preview b,.preview-pane .md-preview strong{color:var(--am);font-weight:600}
.preview-actions{display:flex;gap:8px;padding:8px 16px;border-top:1px solid var(--bd);justify-content:center}
.preview-actions button{
  background:var(--s1);border:1px solid var(--bd);color:var(--ad);padding:8px 16px;
  font-family:var(--fm);font-size:10px;cursor:pointer;transition:all .15s;letter-spacing:.05em;
}
.preview-actions button:hover{border-color:var(--am);color:var(--am);box-shadow:0 0 10px rgba(212,184,74,.12)}
.preview-actions button.primary{background:var(--am);color:#0a0a0c;border-color:var(--am);font-weight:600}
.preview-actions button.primary:hover{box-shadow:0 0 16px rgba(212,184,74,.3)}
@media(max-width:767px){
  .main-layout{flex-direction:column;flex:1;min-height:0}
  .main-left{width:100%!important;max-width:100%!important;border-right:none;flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden}
  .main-right{display:none}
  .tab-panel#panel-chat.active,.tab-panel#panel-files.active{flex:1;min-height:0;overflow:hidden}
  .tab-panel#panel-upload.active{flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column}
  #panel-upload #ufl{flex:1;min-height:0;overflow-y:auto}
  .ca{flex:1;min-height:0;overflow-y:auto;-webkit-overflow-scrolling:touch}
  .chat-wrap{flex:1;min-height:0}
  .fl{flex:1;min-height:0;overflow-y:auto;-webkit-overflow-scrolling:touch}
  .fl.img-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:3px;align-content:start;padding:4px}
  .fl.img-grid .fr{flex-direction:column;min-height:0;padding:0;aspect-ratio:1;position:relative;border:none}
  .fl.img-grid .fr.active{outline:2px solid var(--am);outline-offset:-2px}
  .fl.img-grid .fr .fp{width:100%;height:100%;object-fit:cover;display:block!important}
  .fl.img-grid .fr .fi,.fl.img-grid .fr .fn,.fl.img-grid .fr .fm{display:none!important}
  .cib{flex-shrink:0;background:var(--bg);z-index:10}
  .ipb{flex-shrink:0}
  nav.tabs,header,.chassis{flex-shrink:0}
  .ph{display:none}
  .sort-bar{display:none}
  .br{display:none}
  .cat-bar{padding:10px 12px;gap:6px;border-bottom:none}
  #pv-prev-m,#pv-next-m{display:none!important}
}
@media(min-width:768px){
  .preview-overlay.show-desktop{display:flex!important}
}
@media(min-width:1024px){
  header{padding:12px 24px;font-size:13px}
  header .tid{font-size:15px}
  .url-row{padding:8px 24px;font-size:12px}
  nav.tabs{margin:0}
  .tab-btn{font-size:16px;padding:14px 20px;min-height:52px;gap:8px}
  .tab-btn .svg-icon{width:26px!important;height:26px!important}
  .ph{padding:10px 14px;font-size:12px}
  .cat-btn{padding:8px 14px;font-size:13px;min-height:38px}
  .cat-btn .svg-icon{width:18px!important;height:18px!important}
  .sort-bar{padding:6px 14px;font-size:11px}
  .sort-btn{font-size:11px;padding:4px 8px}
  .br{padding:8px 14px;font-size:12px}
  .main-left{width:440px;min-width:320px;max-width:50%}
  .fr{padding:12px 14px;min-height:52px;gap:10px}
  .fr .fi{width:28px;height:28px}
  .fr .fp{width:44px;height:44px}
  .fr .fn{font-size:15px}
  .fr .fm{font-size:11px}
  .preview-pane{padding:24px}
  .preview-pane .empty{font-size:15px}
  .preview-actions button{padding:10px 24px;font-size:13px}
  .ca{padding:12px 16px}
  .mb{font-size:15px;padding:10px 14px}
  .mib img{max-width:min(360px,100%);max-height:400px}
  .mvid video{max-width:min(480px,100%);max-height:360px}
  .cib{padding:12px 16px;gap:10px}
  .upload-banner{padding:14px 20px}
  .ub-name{font-size:15px}
  .ub-pct{font-size:16px}
  .ub-track{height:10px}
  .ub-status{font-size:12px}
  .ibtn{width:48px;height:48px;min-width:48px}
  .ibtn .svg-icon{width:26px!important;height:26px!important}
  .cib-input{font-size:16px;padding:12px 16px}
  .cs{height:48px;min-width:72px;font-size:15px;padding:0 20px}
  .toast{font-size:13px;padding:12px 20px}
  .pv-arrow{width:44px;height:56px;font-size:28px}
}
body.preview-open{overscroll-behavior-x:none}
.preview-overlay{touch-action:none;overscroll-behavior:contain}
#mbd{flex:1;position:relative;display:flex;min-height:0;overflow:hidden;touch-action:pan-y pinch-zoom}
.preview-overlay.show-desktop #mbd{cursor:pointer}
.preview-overlay.show-desktop #mbd-inner img,.preview-overlay.show-desktop #mbd-inner video,.preview-overlay.show-desktop #mbd-inner iframe{cursor:default}
#mbd-inner{flex:1;display:flex;align-items:center;justify-content:center;overflow:auto;padding:8px 56px;width:100%;min-height:0}
#mbd-inner img,#mbd-inner video{max-width:100%;max-height:calc(100dvh - 150px);object-fit:contain}
.pv-edge{position:absolute;top:50%;transform:translateY(-50%);z-index:25;width:48px;height:64px;background:rgba(0,0,0,.4);border:1px solid rgba(212,184,74,.2);color:var(--am);font-size:28px;display:none;align-items:center;justify-content:center;cursor:pointer;font-family:var(--fm);padding:0}
.pv-edge:active{background:rgba(212,184,74,.15);border-color:var(--am)}
.pv-edge-l{left:0}.pv-edge-r{right:0}
.file-inputs-hidden{position:fixed;left:-9999px;top:0;width:0;height:0;opacity:0;overflow:hidden;z-index:-1}

/* Panel header — desktop only detail */
.ph{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;border-bottom:1px solid var(--bd2);font-family:var(--fs);font-size:12px;flex-shrink:0;background:var(--s1)}
.ph::before,.ph::after{display:none}
.ph-label{color:var(--ad);display:flex;align-items:center;gap:6px;font-weight:500}
.ph-num{display:none}

/* Year gallery — mobile timeline rows */
.year-gallery{flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:20px;padding:12px 12px 28px;-webkit-overflow-scrolling:touch}
.year-section{display:flex;flex-direction:column;gap:10px}
.year-label{font-size:18px;font-weight:600;color:var(--am);padding:4px 2px 2px;position:sticky;top:0;background:var(--bg);z-index:2}
.year-strip{display:flex;flex-direction:row;gap:8px;overflow-x:auto;padding:4px 2px 10px;-webkit-overflow-scrolling:touch;scroll-snap-type:x proximity;min-height:112px;align-items:flex-start}
.yg-item{flex-shrink:0;width:108px;height:108px;cursor:pointer;border:1px solid var(--bd2);background:var(--s1);scroll-snap-align:start;overflow:hidden;position:relative}
.yg-item:active{border-color:var(--am);opacity:.92}
.yg-item img{width:100%;height:100%;object-fit:cover;display:block}
.yg-item .yg-vid{width:100%;height:100%;display:flex;align-items:center;justify-content:center;background:var(--s2);color:var(--am)}
.yg-item.selected,.fr.selected,.mr.sel-row{outline:2px solid var(--am);outline-offset:-2px}
.yg-item.selected::after,.fr.selected::after,.mr.sel-row::after{content:'✓';position:absolute;top:6px;right:6px;width:18px;height:18px;background:var(--am);border:2px solid var(--bg);border-radius:50%;z-index:4;font-size:11px;line-height:14px;text-align:center;color:#070709;font-weight:700;animation:sel-pop .15s ease-out}
body.sel-active #ca .mr{position:relative;cursor:pointer;box-shadow:inset 0 0 0 1px rgba(212,184,74,.18)}
body.sel-active #ca .mr.sel-row{box-shadow:inset 0 0 0 2px var(--am)}
.sel-preview{display:none;position:fixed;bottom:calc(58px + env(safe-area-inset-bottom));left:0;right:0;z-index:189;background:rgba(12,12,16,.94);border-top:1px solid var(--bd2);padding:6px 10px;backdrop-filter:blur(10px)}
.sel-pv-list{display:flex;gap:8px;overflow-x:auto;-webkit-overflow-scrolling:touch;padding:2px 0}
.sel-chip{flex-shrink:0;width:68px;height:68px;border:1px solid var(--bd2);background:var(--bg);padding:3px;font-size:10px;color:var(--tx);cursor:pointer;overflow:hidden;display:flex;align-items:center;justify-content:center;text-align:center;line-height:1.3;font-family:var(--fs);transition:transform .12s,border-color .12s;animation:sel-pop .18s ease-out}
.sel-chip:active{transform:scale(.92);border-color:var(--am)}
.sel-chip img{width:100%;height:100%;object-fit:cover;display:block}
.sel-bar{position:fixed;bottom:0;left:0;right:0;z-index:190;display:none;align-items:center;gap:10px;padding:8px 12px calc(8px + env(safe-area-inset-bottom));background:rgba(10,10,12,.96);border-top:1px solid var(--am-dim);box-shadow:0 -4px 28px rgba(0,0,0,.5);backdrop-filter:blur(12px)}
.sel-count{flex:1;text-align:center;color:var(--am);font-size:22px;font-weight:700;font-family:var(--fm);min-width:0;text-shadow:0 0 14px rgba(212,184,74,.35);letter-spacing:.04em}
.sel-count:empty{opacity:0}
.sel-x,.sel-dl,.sel-del{padding:0;border:1px solid var(--bd2);background:var(--bg);color:var(--ad);cursor:pointer;min-width:44px;min-height:44px;display:inline-flex;align-items:center;justify-content:center;transition:all .12s;-webkit-tap-highlight-color:transparent}
.sel-dl{color:var(--am);border-color:var(--am)}
.sel-del{color:var(--rd);border-color:rgba(216,86,70,.45)}
.sel-del:disabled{opacity:.3;cursor:not-allowed}
.sel-dl:disabled{opacity:.3;cursor:not-allowed}
.sel-x:active,.sel-dl:not(:disabled):active,.sel-del:not(:disabled):active{transform:scale(.92);border-color:var(--am)}
body.sel-active .cib,body.sel-active nav.tabs{display:none!important}
body.sel-active .ca,body.sel-active .fl,body.sel-active .year-gallery{padding-bottom:calc(132px + env(safe-area-inset-bottom))}
@keyframes sel-pop{from{transform:scale(.82);opacity:.5}to{transform:scale(1);opacity:1}}
body.sel-active [data-sel-key]{cursor:pointer}
body.sel-active #ca .mr .md-body,body.sel-active #ca .mr .mb-txt{cursor:pointer;user-select:none;-webkit-user-select:none}
.batch-bar{display:none!important}
.batch-btn{background:var(--s1);border:1px solid var(--am-dim);color:var(--am);padding:8px 14px;font-family:var(--fm);font-size:10px;cursor:pointer;letter-spacing:.05em;box-shadow:var(--bevel)}
.batch-btn:hover{border-color:var(--am);box-shadow:0 0 12px rgba(212,184,74,.15)}
.file-src-bar{display:none;gap:6px;padding:8px 12px 0;flex-shrink:0;flex-wrap:wrap}
body.hub-lan .file-src-bar{display:flex}
.src-btn{background:var(--s1);border:1px solid var(--bd2);color:var(--ad);padding:7px 12px;font-size:11px;cursor:pointer;font-family:var(--fs);transition:all .12s;min-height:34px}
.src-btn.active{color:var(--am);border-color:var(--am);background:rgba(212,184,74,.08);font-weight:600}
.src-btn#src-inbox.active{color:var(--gn);border-color:var(--gn);background:rgba(82,176,70,.08)}
.cat-bar{display:flex;gap:4px;padding:6px 8px;flex-shrink:0;border-bottom:1px solid var(--bd);flex-wrap:wrap}
.cat-btn{background:var(--s1);border:1px solid var(--bd2);color:var(--ad);padding:6px 11px;font-size:10px;cursor:pointer;font-family:var(--fm);transition:all .12s;letter-spacing:.06em;display:inline-flex;align-items:center;gap:4px;min-height:32px;text-transform:uppercase;box-shadow:var(--bevel)}
.cat-btn .svg-icon{width:14px!important;height:14px!important}
.cat-btn.active{color:var(--am);border-color:var(--am);background:var(--s2);box-shadow:var(--inset),0 0 8px rgba(212,184,74,.12)}
.cat-btn:active:not(.active){background:var(--s2);border-color:var(--am-dim)}
.cat-btn.rx-btn{color:var(--gn);border-color:rgba(74,140,63,.5);font-weight:600}
.cat-btn.rx-btn:hover,.cat-btn.rx-btn.active{border-color:var(--gn);background:rgba(74,140,63,.08)}
.cat-count{color:var(--t3);font-size:7px;margin-left:2px}

/* Sort — text links, visually distinct from category chips */
.sort-bar{display:flex;align-items:center;gap:6px;padding:3px 8px;flex-shrink:0;border-bottom:1px solid var(--bd);font-family:var(--fm);font-size:8px;color:var(--t3)}
.sort-lbl{color:var(--t3);letter-spacing:.1em;margin-right:2px}
.sort-btn{background:none;border:none;border-bottom:1px solid transparent;color:var(--ad);padding:2px 4px;cursor:pointer;font-family:var(--fm);font-size:8px;letter-spacing:.05em;transition:all .15s}
.sort-btn:hover{color:var(--am)}
.sort-btn.active{color:var(--am);border-bottom-color:var(--am)}
.sort-arr{font-size:7px;color:var(--am);margin-left:1px}

/* Breadcrumb */
.br{display:flex;align-items:center;gap:1px;flex-wrap:wrap;padding:4px 8px;font-family:var(--fm);font-size:9px;color:var(--t2);flex-shrink:0;border-bottom:1px solid var(--bd)}
.br a{color:var(--ad);text-decoration:none;padding:1px 4px;transition:color .15s}.br a:hover{color:var(--am)}
.br .s{color:var(--t3)}

/* File list */
.fl{flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:1px}
.fr{display:flex;align-items:center;padding:8px 10px;color:var(--tx);gap:8px;cursor:pointer;border:1px solid transparent;border-bottom:1px solid rgba(61,58,34,.35);transition:all .1s;min-height:44px;font-family:var(--fm);font-size:11px}
.fr:hover{background:var(--s1);border-color:var(--bd2)}
.fr.active{background:var(--s2);border-color:var(--am);box-shadow:inset 3px 0 0 var(--am)}
.fr .fi{width:22px;height:22px;flex-shrink:0;color:var(--ad)}.fr.dir .fi{color:var(--am)}
.fr .fp{width:36px;height:36px;object-fit:cover;flex-shrink:0;display:none;border:1px solid var(--bd);background:var(--s2)}.fr .fp.show{display:block}
.fr .fp.show+.fi{display:none}
.fr .fn{flex:1;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;font-family:var(--fs)}
.fr .fm{font-size:9px;color:var(--ad);flex-shrink:0;font-family:var(--fm);text-align:right}
.fr .fs{min-width:42px}.fr .fd{min-width:52px;display:none}
@media(min-width:500px){.fr .fd{display:block}}

/* Upload */
.uqb{display:flex;gap:6px;padding:8px;flex-shrink:0}
.qb{flex:1;display:flex;flex-direction:column;align-items:center;gap:6px;padding:16px 10px;cursor:pointer;background:var(--s1);border:1px solid var(--bd);color:var(--ad);font-size:10px;font-family:var(--fm);transition:all .15s;letter-spacing:.05em}
.qb:hover,.qb:active{border-color:var(--am);color:var(--am);background:var(--s2)}
#fi,#pi{position:absolute!important;left:-9999px!important;width:0!important;height:0!important;opacity:0!important;pointer-events:none!important}
#ci,#cg{position:fixed!important;left:-100vw!important;top:0!important;width:1px!important;height:1px!important;opacity:0!important;overflow:hidden!important}
#ufl{flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:3px;padding:0 8px}
.uc{background:var(--s1);border:1px solid var(--bd);padding:8px;display:flex;flex-direction:column;gap:5px;transition:all .3s}
.uc.up{border-color:var(--am);background:var(--s2)}.uc.ok{border-color:var(--gn)}.uc.er{border-color:var(--rd)}
.ui{display:flex;justify-content:space-between;align-items:center;gap:6px}
.un{font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0;font-family:var(--fs)}
.um{font-size:9px;color:var(--ad);white-space:nowrap;font-family:var(--fm)}
.us{font-size:8px;font-family:var(--fm);text-align:right;min-width:44px;letter-spacing:.05em}
.us.wt{color:var(--ad)}.us.up{color:var(--am)}.us.ok{color:var(--gn)}.us.er{color:var(--rd)}
.pt{height:2px;background:var(--bd);overflow:hidden;opacity:0;transition:opacity .2s}
.uc.up .pt,.uc.ok .pt,.uc.er .pt{opacity:1}
.pf{height:100%;background:var(--am);width:0%;transition:width .15s linear}
.uc.ok .pf{background:var(--gn);width:100%!important}.uc.er .pf{background:var(--rd);width:100%!important}
.ua{display:flex;gap:6px;padding:6px 8px;flex-shrink:0}
.btn{flex:1;padding:10px 14px;border:none;font-size:11px;cursor:pointer;transition:all .15s;font-family:var(--fm);letter-spacing:.05em}
.btn:active:not(:disabled){opacity:1}
.btn-pri{background:var(--am);color:#0a0a0c;font-weight:600}
.btn-pri:hover:not(:disabled){box-shadow:0 0 16px rgba(212,184,74,.25)}
.btn-pri:disabled{background:var(--s2);color:var(--ad);cursor:not-allowed}
.btn-sec{background:var(--s1);color:var(--ad);border:1px solid var(--bd)}.btn-sec:disabled{opacity:.3;cursor:not-allowed}
.usm{padding:8px;font-size:10px;text-align:center;display:none;font-family:var(--fm);border:1px solid var(--bd);margin:6px 8px}
.usm.show{display:block;animation:fu .3s ease}
.usm.ok{color:var(--gn);border-color:rgba(74,140,63,.2)}.usm.er{color:var(--rd);border-color:rgba(196,74,58,.2)}

/* Chat */
.chat-wrap{position:relative;flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden}
.chat-filter{display:flex;gap:6px;padding:8px 14px 0;flex-shrink:0}
.chat-filter button{padding:6px 12px;border:1px solid var(--bd2);background:var(--s1);color:var(--ad);font-size:12px;cursor:pointer;font-family:var(--fs)}
.chat-filter button.on{border-color:var(--am);color:var(--am);background:rgba(212,184,74,.06)}
.mr.hide-filter{display:none!important}
.ca{flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:14px;padding:16px 14px 20px;background:var(--bg);-webkit-overflow-scrolling:touch}
.jump-bottom{display:none;position:absolute;bottom:14px;right:14px;z-index:20;padding:9px 14px;border:1px solid var(--am);background:var(--s1);color:var(--am);font-size:13px;font-family:var(--fs);cursor:pointer;box-shadow:0 4px 18px rgba(0,0,0,.45);align-items:center;gap:4px}
.jump-bottom.show{display:flex}
.load-older{display:none;width:100%;padding:10px 14px;margin:0 0 6px;border:1px dashed var(--bd2);background:var(--s1);color:var(--ad);font-size:13px;font-family:var(--fs);cursor:pointer;text-align:center;flex-shrink:0}
.load-older.show{display:block}
.load-older:disabled{opacity:.45;cursor:wait}
.mb .md-body{font-family:var(--fs);font-size:14px;line-height:1.65;word-break:break-word}
.mb .md-body h1,.mb .md-body h2,.mb .md-body h3{color:var(--am);margin:10px 0 6px;font-weight:600;line-height:1.35}
.mb .md-body h1{font-size:17px}.mb .md-body h2{font-size:15px}.mb .md-body h3{font-size:14px}
.mb .md-body pre{background:var(--s1);border:1px solid var(--bd2);padding:10px;overflow-x:auto;font-family:var(--fm);font-size:12px;line-height:1.5;margin:8px 0;white-space:pre-wrap}
.mb .md-body code{background:var(--s1);padding:1px 5px;font-family:var(--fm);font-size:12px;color:var(--am)}
.mb .md-body pre code{padding:0;background:none;color:var(--tx)}
.mb .md-body ul,.mb .md-body ol{margin:6px 0 6px 20px}
.mb .md-body li{margin:3px 0}
.mb .md-body blockquote{border-left:3px solid var(--am);padding:4px 0 4px 12px;margin:8px 0;color:var(--ad)}
.mb .md-body a{color:var(--am);text-decoration:underline}
.mb .md-body p{margin:0 0 8px}
.mr{display:flex;flex-direction:column;animation:mi .2s ease;max-width:88%}
.mr.pc{align-self:flex-start;align-items:flex-start}
.mr.ph{align-self:flex-end;align-items:flex-end}
.mb{padding:10px 13px;font-size:14px;line-height:1.6;word-break:break-word;white-space:pre-wrap;position:relative;font-family:var(--fs);max-width:100%}
.mb-txt{word-break:break-word;white-space:pre-wrap}
.mb .md-body,.mb .mb-txt{user-select:text;-webkit-user-select:text;cursor:text}
.expand-btn{display:block;margin-top:8px;background:none;border:none;color:var(--am);font-size:13px;padding:6px 0;cursor:pointer;font-family:var(--fs);text-align:left;width:100%;touch-action:manipulation;-webkit-tap-highlight-color:transparent;position:relative;z-index:2}
.expand-top{margin:0 0 8px;padding:6px 0 8px;border-bottom:1px dashed var(--bd2)}
.mb-expanded .expand-btn:not(.expand-top){position:sticky;bottom:0;margin-top:12px;padding:10px 0 8px;background:linear-gradient(to bottom,rgba(0,0,0,0),var(--s2) 35%);box-shadow:0 -8px 16px rgba(0,0,0,.25)}
.mr.ph .mb.mb-expanded .expand-btn:not(.expand-top){background:linear-gradient(to bottom,rgba(0,0,0,0),rgba(212,184,74,.12) 35%)}
.mb-collapsed{display:-webkit-box;-webkit-line-clamp:5;-webkit-box-orient:vertical;overflow:hidden;white-space:normal}
.mcp{position:absolute;top:4px;right:4px;background:rgba(10,10,12,.9);color:var(--ad);border:1px solid var(--bd2);padding:0;cursor:pointer;opacity:.35;transition:opacity .15s;z-index:3;width:28px;height:28px;display:flex;align-items:center;justify-content:center}
.mb:hover .mcp,.mb:active .mcp{opacity:1;pointer-events:auto}
@media(hover:none){.mcp{opacity:.7;pointer-events:auto}}
.mr.pc .mb{background:var(--s2);color:var(--tx);border:1px solid var(--bd2);border-left:3px solid var(--am)}
.mr.ph .mb{background:rgba(212,184,74,.04);color:#f0e4a8;border:1px solid rgba(212,184,74,.15);border-right:3px solid var(--gn)}
.mb b,.mb strong{color:var(--am)}
.mb code{background:var(--s1);padding:1px 4px;font-family:var(--fm);font-size:11px;color:var(--am)}
.mb a,.mb-txt a{color:var(--am);text-decoration:underline;cursor:pointer;word-break:break-all}
.mb a:active,.mb-txt a:active{opacity:.75}
.mib{padding:0!important;background:transparent!important;border:none!important;overflow:visible}
.mib img{display:block;max-width:min(280px,100%);width:auto;height:auto;max-height:min(420px,65vh);object-fit:contain;border:1px solid var(--bd2);border-radius:2px;background:var(--s2);transition:opacity .2s}
.mib img.img-loading{opacity:.35;min-height:100px}
.mib img.img-ready{opacity:1}
.preview-pane .docx-body,.mbd-inner .docx-body{font-family:var(--fs);font-size:14px;line-height:1.7;color:var(--tx)}
.preview-pane .docx-body p,.mbd-inner .docx-body p{margin:0 0 10px}
.preview-pane .docx-body h1,.preview-pane .docx-body h2,.mbd-inner .docx-body h1,.mbd-inner .docx-body h2{color:var(--am);margin:12px 0 8px}
.docx-body figure.docx-img{margin:14px 0;text-align:center}
.docx-body figure.docx-img img{max-width:100%;height:auto;border:1px solid var(--bd2);background:var(--s1)}
.mb.mdoc{padding:12px!important;cursor:pointer;border:1px solid var(--bd2)!important;background:var(--s1)!important}
.mb.mdoc .docx-title{color:var(--am);font-weight:600;margin-bottom:4px}
.mb.mdoc .docx-hint{font-size:12px;color:var(--ad)}
.mb.mzip{padding:12px!important;cursor:pointer;border:1px solid rgba(212,184,74,.35)!important;background:linear-gradient(135deg,var(--s1),rgba(212,184,74,.06))!important;display:flex!important;gap:12px;align-items:center;max-width:min(340px,100%);min-width:220px}
.mb.mzip .zip-icon{flex-shrink:0;width:48px;height:48px;display:flex;align-items:center;justify-content:center;background:rgba(212,184,74,.14);border:1px solid rgba(212,184,74,.4);border-radius:6px;color:var(--am)}
.mb.mzip .zip-body{flex:1;min-width:0}
.mb.mzip .zip-title{color:var(--am);font-weight:600;margin-bottom:4px;word-break:break-all;line-height:1.35;font-size:14px}
.mb.mzip .zip-hint{font-size:12px;color:var(--ad)}
.mb.mzip .zip-tag{display:inline-block;margin-top:5px;padding:2px 7px;font-size:10px;font-family:var(--fm);letter-spacing:.06em;color:var(--am);border:1px solid rgba(212,184,74,.35);background:rgba(212,184,74,.08)}
.file-path-hint{font-size:11px;color:var(--ad);margin-top:6px;font-family:var(--fm);word-break:break-all;line-height:1.4;opacity:.85}
.mvid,.maud{padding:6px!important;background:var(--s1)!important;border:1px solid var(--bd2)!important}
.maud .aud-name{font-size:12px;color:var(--am);margin-bottom:6px;word-break:break-all;line-height:1.4;display:flex;align-items:center;gap:6px}
.maud audio{display:block;width:min(280px,100%);min-height:36px}
.mvid video{display:block;max-width:min(280px,100%);max-height:240px;border:1px solid var(--bd);background:#000}
.mvid{position:relative}
.mvid .v-exp{position:absolute;top:6px;right:6px;z-index:4;background:rgba(10,10,12,.85);border:1px solid var(--bd2);color:var(--am);width:32px;height:32px;padding:0;cursor:pointer;font-size:14px;line-height:1}
.mvid .v-exp:active{border-color:var(--am)}
.maud audio{display:block;width:min(260px,100%)}
.mm{font-size:8px;color:var(--t3);margin-top:3px;padding:0 2px;font-family:var(--fm);letter-spacing:.05em;line-height:1.2;width:100%}
.mr.pc .mm{text-align:left}
.mr.ph .mm{text-align:right}
.mm-nm{display:inline;margin-right:6px;color:var(--ad);font-weight:600}
.mr.ph .mm-nm{color:var(--gn)}
.mr.pc .mm-nm{color:var(--am)}
.cib{display:flex;gap:8px;padding:12px 12px calc(12px + env(safe-area-inset-bottom));align-items:stretch;flex-shrink:0;background:var(--s1);border-top:1px solid var(--bd2)}
.cib-tools{display:flex;gap:8px;flex-shrink:0;align-items:stretch}
.ibtn{width:44px;min-width:44px;height:44px;padding:0;border:1px solid var(--bd2);background:var(--bg);color:var(--ad);cursor:pointer;display:flex;align-items:center;justify-content:center;transition:border-color .12s;-webkit-tap-highlight-color:transparent;box-shadow:none;flex-shrink:0}
.ibtn:hover{border-color:var(--am);color:var(--am);box-shadow:0 0 10px rgba(212,184,74,.1)}
.ibtn .svg-icon{width:22px!important;height:22px!important}
.cib-input::placeholder{color:var(--t3)}
.cib-input{flex:1;min-width:0;background:var(--bg);border:1px solid var(--bd2);color:var(--tx);padding:0 12px;font-size:15px;font-family:var(--fs);outline:none;transition:border .12s;line-height:1.4;resize:none;overflow-y:auto;max-height:120px;min-height:44px;height:44px;box-shadow:none;align-self:stretch}
.cib-input:focus{border-color:var(--am)}
.cs{background:var(--am);color:#070709;border:1px solid var(--am);padding:0 16px;min-width:60px;min-height:44px;height:44px;font-size:15px;cursor:pointer;font-family:var(--fs);font-weight:600;transition:opacity .12s;flex-shrink:0;align-self:stretch;letter-spacing:0;text-transform:none;box-shadow:none}
.cs:hover:not(:disabled){box-shadow:0 0 14px rgba(212,184,74,.28)}
.file-inputs-hidden input{position:absolute;width:0;height:0;opacity:0;font-size:0;border:0;padding:0;margin:0}
.cib select#ss{display:none}
.ipb{display:flex;align-items:center;gap:8px;padding:6px 10px;border-top:1px solid var(--bd);background:var(--s2)}.ipb img{max-height:48px;max-width:48px;object-fit:cover;border:1px solid var(--bd);cursor:pointer;flex-shrink:0}
.ipb .ipn{font-size:10px;color:var(--ad);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-family:var(--fm)}
.ipx{background:none;border:1px solid var(--bd);color:var(--ad);padding:4px 8px;font-size:9px;cursor:pointer;font-family:var(--fm);flex-shrink:0}.ipx:hover{border-color:var(--rd);color:var(--rd)}
.attach-queue{border-top:1px solid var(--bd);background:var(--s2);padding:6px 10px;flex-shrink:0}
.aq-hd{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px;font-size:9px;font-family:var(--fm);color:var(--ad);margin-bottom:6px}
.aq-hd button{background:none;border:1px solid var(--bd);color:var(--ad);padding:2px 8px;font-size:8px;cursor:pointer;font-family:var(--fm)}
.aq-hd button:active{border-color:var(--rd);color:var(--rd)}
.pack-choice-overlay{display:none;position:fixed;inset:0;z-index:350;background:rgba(0,0,0,.78);align-items:center;justify-content:center;padding:16px;padding-bottom:calc(16px + env(safe-area-inset-bottom))}
.pack-choice-overlay.show{display:flex}
.pack-choice-panel{width:min(360px,100%);background:var(--bg);border:1px solid var(--bd2);box-shadow:0 16px 48px rgba(0,0,0,.55);padding:20px 18px 16px;text-align:center}
.pack-choice-title{font-size:17px;color:var(--tx);font-family:var(--fs);font-weight:600;margin-bottom:6px}
.pack-choice-sub{font-size:13px;color:var(--ad);font-family:var(--fs);margin-bottom:18px;line-height:1.5}
.pack-choice-actions{display:flex;flex-direction:column;gap:10px}
.pack-choice-btn{width:100%;min-height:48px;border:1px solid var(--bd2);background:var(--s1);color:var(--tx);font-size:15px;font-family:var(--fs);cursor:pointer;padding:12px 16px}
.pack-choice-btn.primary{border-color:var(--am);color:var(--am);background:rgba(212,184,74,.08);font-weight:600}
.pack-choice-btn:active{transform:scale(.98);opacity:.9}
.pack-choice-cancel{margin-top:4px;background:none;border:none;color:var(--ad);font-size:13px;font-family:var(--fs);padding:10px;cursor:pointer;min-height:44px}
.aq-list{display:flex;gap:8px;overflow-x:auto;padding:2px 0;-webkit-overflow-scrolling:touch;scrollbar-width:none}
.aq-list::-webkit-scrollbar{display:none}
.aq-chip{flex-shrink:0;width:58px;position:relative;text-align:center}
.aq-chip .aq-thumb{width:52px;height:52px;border:1px solid var(--bd);object-fit:cover;cursor:pointer;background:var(--s1);display:flex;align-items:center;justify-content:center;margin:0 auto}
.aq-chip .aq-thumb img,.aq-chip .aq-thumb video{width:100%;height:100%;object-fit:cover}
.aq-chip .aq-nm{font-size:7px;color:var(--t3);max-width:58px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-top:3px;font-family:var(--fm)}
.aq-chip .aq-rm{position:absolute;top:-5px;right:0;width:18px;height:18px;border:1px solid var(--bd);background:var(--bg);color:var(--rd);font-size:11px;cursor:pointer;line-height:16px;padding:0;border-radius:0}
.aq-chip.up .aq-thumb{border-color:var(--am);opacity:.6}
.aq-chip.ok .aq-thumb{border-color:var(--gn)}
.aq-chip.er .aq-thumb{border-color:var(--rd)}
.aq-progress{height:3px;background:var(--bd);margin-top:4px;overflow:hidden;width:52px;margin-left:auto;margin-right:auto}
.aq-progress .aq-pf{height:100%;background:var(--am);width:0;transition:width .12s linear}
.aq-chip.ok .aq-progress .aq-pf{background:var(--gn);width:100%!important}
.aq-chip.er .aq-progress .aq-pf{background:var(--rd);width:100%!important}
.aq-pct{font-size:7px;color:var(--am);font-family:var(--fm);margin-top:2px}

/* Upload progress banner — full width above input */
.upload-banner{display:none;border-top:1px solid var(--bd);background:var(--s1);padding:10px 14px;flex-shrink:0}
.upload-banner.show{display:block}
.ub-row{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:8px}
.ub-name{font-size:13px;color:var(--tx);font-family:var(--fs);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0}
.ub-pct{font-size:14px;color:var(--am);font-family:var(--fm);font-weight:600;flex-shrink:0;min-width:42px;text-align:right}
.ub-track{height:8px;background:var(--steel);overflow:hidden;border:1px solid var(--bd2);box-shadow:var(--inset)}
.ub-fill{height:100%;background:linear-gradient(90deg,var(--am-dim),var(--am));width:0;transition:width .12s linear;box-shadow:0 0 8px rgba(212,184,74,.3)}
.ub-fill.done{background:var(--gn);width:100%!important}
.ub-fill.er{background:var(--rd);width:100%!important}
.ub-status{font-size:11px;color:var(--ad);font-family:var(--fm);margin-top:6px;letter-spacing:.04em}
.ub-status.er{color:var(--rd)}.ub-status.ok{color:var(--gn)}
.download-banner{display:none;border-top:1px solid var(--bd);background:rgba(110,181,255,.06);padding:10px 14px;flex-shrink:0}
.download-banner.show{display:block}
.ub-fill.indeterminate{width:40%!important;animation:ub-slide 1.1s ease-in-out infinite alternate}
@keyframes ub-slide{from{margin-left:0}to{margin-left:60%}}

.es{text-align:center;padding:24px 12px;color:var(--ad);font-size:10px;display:flex;flex-direction:column;align-items:center;gap:5px;font-family:var(--fm);letter-spacing:.05em}
.es .svg-icon{opacity:.25;margin-bottom:3px}
.toast{position:fixed;bottom:50px;left:50%;transform:translateX(-50%);background:var(--s3);border:1px solid var(--am);color:var(--am);padding:7px 14px;font-size:9px;opacity:0;pointer-events:none;transition:opacity .3s;font-family:var(--fm);z-index:99;letter-spacing:.05em}
.toast.show{opacity:1}

/* Search overlay */
.search-overlay{display:none;position:fixed;inset:0;z-index:300;background:rgba(0,0,0,.72);align-items:flex-start;justify-content:center;padding:12px 12px calc(12px + env(safe-area-inset-bottom))}
.search-overlay.show{display:flex}
.search-panel{width:min(720px,100%);max-height:min(88vh,900px);background:var(--bg);border:1px solid var(--bd2);box-shadow:0 16px 48px rgba(0,0,0,.55);display:flex;flex-direction:column;overflow:hidden}
.search-hd{display:flex;gap:8px;padding:12px 12px 8px;border-bottom:1px solid var(--bd);align-items:center}
.search-hd input{flex:1;min-width:0;padding:12px 14px;border:1px solid var(--bd2);background:var(--s1);color:var(--tx);font-size:16px;font-family:var(--fs);outline:none}
.search-hd input:focus{border-color:var(--am);box-shadow:0 0 0 1px rgba(212,184,74,.25)}
.search-close{width:44px;height:44px;border:1px solid var(--bd2);background:var(--s1);color:var(--ad);font-size:20px;cursor:pointer;flex-shrink:0;font-family:var(--fm)}
.search-scope{display:flex;gap:6px;padding:8px 12px;border-bottom:1px solid var(--bd);flex-wrap:wrap}
.search-scope button{padding:6px 12px;border:1px solid var(--bd2);background:var(--s1);color:var(--ad);font-size:12px;cursor:pointer;font-family:var(--fs)}
.search-scope button.on{border-color:var(--am);color:var(--am)}
.search-hint{padding:8px 14px;font-size:11px;color:var(--ad);font-family:var(--fm);border-bottom:1px solid var(--bd)}
.search-results{flex:1;overflow-y:auto;padding:8px 0;-webkit-overflow-scrolling:touch}
.search-group{padding:6px 12px 2px;font-size:10px;color:var(--ad);font-family:var(--fm);letter-spacing:.08em}
.search-item{display:block;width:100%;text-align:left;padding:10px 14px;border:none;border-bottom:1px solid var(--bd);background:transparent;color:var(--tx);cursor:pointer;font-family:var(--fs);font-size:13px;line-height:1.5;-webkit-tap-highlight-color:transparent;touch-action:manipulation}
.search-item:active,.search-item:hover{background:var(--s1)}
.search-item .sr-snippet{color:var(--tx);word-break:break-word;white-space:pre-wrap}
.search-item .sr-meta{font-size:11px;color:var(--ad);margin-top:4px;font-family:var(--fm)}
.search-item mark{background:rgba(212,184,74,.25);color:var(--am);padding:0 1px}
.search-empty{padding:24px 14px;text-align:center;color:var(--ad);font-size:13px}
.mr.msg-flash{animation:msgflash 1.2s ease 2}
@keyframes msgflash{0%,100%{box-shadow:none}50%{box-shadow:0 0 0 2px var(--am),0 0 24px rgba(212,184,74,.35)}}

@keyframes fu{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}
@keyframes mi{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:translateY(0)}}
@keyframes spin{to{transform:rotate(360deg)}}
.spinner{display:inline-block;width:14px;height:14px;border:1px solid var(--bd);border-top-color:var(--am);animation:spin .6s linear infinite}
</style></head><body class="hub-__HUB_MODE__">
<header class="deck-h">
<div class="deck-left"><span class="brand-mark">__LOGO__</span><span class="tid" id="hub-title-zh">__HUB_SUB__</span><span class="hub-en" id="hub-title-en">__HUB_TITLE__</span><span class="mode-badge mode-__HUB_MODE__" id="hub-badge">__HUB_BADGE__</span></div>
<div class="deck-right"><button type="button" class="deck-btn" id="search-btn" title="搜索 (Ctrl+K)" aria-label="搜索">__SEARCH__</button><button type="button" class="deck-btn sel-toggle-btn" id="sel-toggle" title="多选" aria-label="多选">__GRID__</button><button class="deck-btn sfx-btn on" id="sfx-btn" onclick="toggleSfx()" title="音效" aria-label="音效">__SPK__</button><button class="deck-btn sfx-btn" id="notify-btn" onclick="requestNotify()" title="桌面通知" aria-label="通知" style="display:none">__BELL__</button></div>
</header>
<div class="lan-bar" id="lan-bar"><span class="lan-bar-label">手机访问</span><button type="button" class="lan-bar-url" id="lan-bar-url" onclick="copyLanUrl()" title="点击复制到剪贴板">http://---</button><span class="lan-bar-hint">须同一 WiFi · Relay Local</span></div>
<div class="cloud-bar" id="cloud-bar"><span class="cloud-bar-label">云收件箱</span><span class="cloud-bar-url" id="cloud-bar-url">—</span><span class="cloud-bar-hint">数据在服务器 · 回家 PC 自动同步 · Relay Cloud</span></div>
<span class="url-hidden" id="lu">http://---</span>
<div class="chassis">
<nav class="tabs">
<button class="tab-btn" data-tab="files">__I1__ 文件</button>
<button class="tab-btn" data-tab="upload" style="display:none">__I2__ 上传</button>
<button class="tab-btn active" data-tab="chat">__I3__ 聊天<span class="tab-badge" id="ub"></span></button>
</nav>
<main>
<div class="tab-panel" id="panel-files"><div class="main-layout">
<div class="main-left">
<div class="ph"><span class="ph-label" id="files-ph-label">文件</span></div>
<div class="file-src-bar" id="file-src-bar">
<button type="button" class="src-btn active" id="src-home" data-src="home">本机全部</button>
<button type="button" class="src-btn" id="src-inbox" data-src="inbox">手机传来</button>
</div>
<div class="cat-bar" id="cat-bar">
<button class="cat-btn active" data-cat="all">全部</button>
<button class="cat-btn" data-cat="image">图片</button>
<button class="cat-btn" data-cat="video">视频</button>
<button class="cat-btn" data-cat="audio">音频</button>
<button class="cat-btn" data-cat="doc">文档</button>
</div>
<div class="batch-bar" id="batch-bar" style="display:none"><button type="button" class="batch-btn" id="batch-dl-btn">__BATCH__ 批量下载</button></div>
<div class="br" id="br"></div>
<div class="sort-bar" id="sort-bar">
<span class="sort-lbl">SORT</span>
<button class="sort-btn" data-sort="name">NAME</button>
<button class="sort-btn" data-sort="size">SIZE</button>
<button class="sort-btn active" data-sort="mtime">DATE</button>
</div>
<div class="fl" id="fl"><div class="es"><div class="spinner"></div>LOADING</div></div>
</div>
<div class="main-right"><div class="ph"><span class="ph-label">预览</span></div><div class="preview-wrap" id="pw"><button type="button" class="pv-arrow pv-prev" id="pv-prev-d" style="display:none">‹</button><div class="preview-pane" id="pp"><div class="empty">选择左侧文件预览</div></div><button type="button" class="pv-arrow pv-next" id="pv-next-d" style="display:none">›</button><span class="pv-count" id="pv-count-d" style="display:none"></span></div><div class="preview-actions" id="pa" style="display:none"><button onclick="dlPreview()" data-dl-btn>下载</button><button onclick="dlPreviewCopy()" data-dl-copy-btn style="display:none">另存副本</button><button id="stop-preview-btn" style="display:none" onclick="stopMedia()">停止</button><button class="primary" id="send-preview-btn" style="display:none" onclick="sendPreviewToChat()">发到聊天</button></div></div>
</div></div>

<div class="tab-panel" id="panel-upload">
<div class="ph"><span class="ph-label"><span class="ph-num">02.</span>UPLOAD TO SERVER</span><span style="font-size:8px;color:var(--ad)">&rarr; Downloads/from-phone/</span></div>
<div class="uqb"><button class="qb" id="pbt">__I6__<span>PHOTO/VIDEO</span></button><button class="qb" id="fbt">__I7__<span>FILES</span></button></div>
<input type="file" id="fi" multiple><input type="file" id="pi" accept="image/*,video/*" multiple>
<div class="es" id="ue">__I8__NO FILES SELECTED</div>
<div id="ufl"></div>
<div class="ua"><button class="btn btn-sec" id="cb" disabled>CLEAR</button><button class="btn btn-pri" id="ubt" disabled>UPLOAD</button></div>
<div class="usm" id="us"></div>
</div>

<div class="tab-panel active" id="panel-chat">
<div class="chat-wrap">
<div class="chat-filter" id="chat-filter">
<button type="button" class="on" id="cf-all">全部</button>
<button type="button" id="cf-text">仅文字</button>
</div>
<div class="ca" id="ca"><button type="button" class="load-older" id="load-older">↑ 加载更早消息</button><div class="es">__I9__NO MESSAGES YET</div></div>
<button type="button" class="jump-bottom" id="jump-bottom">↓ 最新消息</button>
</div>
<div class="attach-queue" id="aqb" style="display:none">
<div class="aq-hd"><span id="aq-count">0 个文件</span><button type="button" id="aq-clear">清空</button></div>
<div class="aq-list" id="aq-list"></div>
</div>
<div class="upload-banner" id="upload-banner">
<div class="ub-row"><span class="ub-name" id="ub-name">—</span><span class="ub-pct" id="ub-pct">0%</span></div>
<div class="ub-track"><div class="ub-fill" id="ub-fill"></div></div>
<div class="ub-status" id="ub-status">上传中</div>
</div>
<div class="download-banner" id="download-banner">
<div class="ub-row"><span class="ub-name" id="db-name">—</span><span class="ub-pct" id="db-pct">0%</span></div>
<div class="ub-track"><div class="ub-fill" id="db-fill"></div></div>
<div class="ub-status" id="db-status">下载中</div>
</div>
<div class="cib">
<div class="cib-tools">
<label class="ibtn" id="cfile" for="ci" aria-label="文件">__I10b__</label>
<label class="ibtn" id="cgallery" for="cg" aria-label="相册">__I6b__</label>
</div>
<textarea class="cib-input" id="ci2" placeholder="输入消息…" rows="1" autocomplete="off"></textarea>
<button type="button" class="cs" id="csb">发送</button>
</div>
<div class="file-inputs-hidden">
<input type="file" id="ci" accept="*/*" multiple>
<input type="file" id="cg" accept="image/*,video/*" multiple>
</div>
</div></main></div>
<div class="search-overlay" id="search-overlay">
<div class="search-panel">
<div class="search-hd">
<input type="search" id="search-input" placeholder="搜索账号、密码、聊天记录、文件名…" autocomplete="off" enterkeyhint="search">
<button type="button" class="search-close" id="search-close" aria-label="关闭">×</button>
</div>
<div class="search-scope" id="search-scope">
<button type="button" class="on" data-scope="all">全部</button>
<button type="button" data-scope="chat">聊天</button>
<button type="button" data-scope="files" id="search-scope-files">文件</button>
</div>
<div class="search-hint" id="search-hint">搜索全部聊天记录 · 云端永久保存</div>
<div class="search-results" id="search-results"><div class="search-empty">输入关键词开始搜索</div></div>
</div>
</div>
<!-- Mobile preview overlay -->
<div class="preview-overlay" id="mo" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.95);z-index:200;flex-direction:column">
<div class="ph" style="border-bottom:1px solid var(--bd)">
<button type="button" class="pv-arrow pv-prev" id="pv-prev-m" style="display:none;position:static;transform:none;width:36px;height:32px">‹</button>
<span id="mfn" style="color:var(--am);font-family:var(--fm);font-size:11px;flex:1;text-align:center;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"></span>
<span id="pv-count-m" class="pv-count" style="display:none;position:static;transform:none"></span>
<button type="button" class="pv-arrow pv-next" id="pv-next-m" style="display:none;position:static;transform:none;width:36px;height:32px">›</button>
<button class="preview-close" onclick="closePreviewBtn()" style="background:none;border:1px solid var(--bd);color:var(--ad);padding:6px 10px;font-family:var(--fm);font-size:10px;cursor:pointer;margin-left:4px">×</button>
</div>
<div id="mbd" style="flex:1;position:relative;display:flex;min-height:0;overflow:hidden">
<button type="button" class="pv-edge pv-edge-l" id="pv-edge-l">‹</button>
<div id="mbd-inner"></div>
<button type="button" class="pv-edge pv-edge-r" id="pv-edge-r">›</button>
</div>
<div id="mdl" style="display:none;padding:8px;border-top:1px solid var(--bd);gap:8px;justify-content:center;flex-wrap:wrap">
<button onclick="dlPreview()" data-dl-btn style="background:var(--s1);border:1px solid var(--bd);color:var(--ad);padding:8px 16px;font-family:var(--fm);font-size:10px;cursor:pointer">下载</button>
<button onclick="dlPreviewCopy()" data-dl-copy-btn style="display:none;background:var(--s1);border:1px solid var(--bd);color:var(--ad);padding:8px 16px;font-family:var(--fm);font-size:10px;cursor:pointer">另存副本</button>
<button id="send-preview-btn-m" style="display:none" class="primary" onclick="sendPreviewToChat()" style="background:var(--am);color:#0a0a0c;border:none;padding:8px 16px;font-family:var(--fm);font-size:10px;cursor:pointer;font-weight:600">发到聊天</button>
</div></div>
<div class="drop-overlay" id="drop-ov"><span class="drop-title">松手上传</span><span class="drop-sub">图片 · 视频 · 文档 · 多文件可选打包</span></div>
<div class="pack-choice-overlay" id="pack-choice-overlay" role="dialog" aria-modal="true" aria-labelledby="pack-choice-title">
<div class="pack-choice-panel">
<div class="pack-choice-title" id="pack-choice-title">已选 <span id="pack-choice-count">0</span> 个文件</div>
<div class="pack-choice-sub">怎么发送到聊天？</div>
<div class="pack-choice-actions">
<button type="button" class="pack-choice-btn" id="pack-choice-separate">逐条发送</button>
<button type="button" class="pack-choice-btn primary" id="pack-choice-zip">打包成一个 zip</button>
<button type="button" class="pack-choice-cancel" id="pack-choice-cancel">取消</button>
</div></div></div>
<div class="toast" id="toast"></div>
<div class="sel-preview" id="sel-preview"><div class="sel-pv-list" id="sel-pv-list"></div></div>
<div class="sel-bar" id="sel-bar" style="display:none">
<button type="button" class="sel-x" id="sel-cancel" aria-label="取消">__SELX__</button>
<span class="sel-count" id="sel-count"></span>
<button type="button" class="sel-dl" id="sel-dl" disabled aria-label="下载">__BATCH__</button>
<button type="button" class="sel-del" id="sel-del" disabled aria-label="删除">__TRASH__</button>
</div>
<select id="ss" style="display:none"><option value="phone">phone</option><option value="pc">pc</option></select>

<script src="https://cdn.jsdelivr.net/npm/heic2any@0.0.4/dist/heic2any.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js"></script>
<script>
var ICONS=__JSON__;
var UPLOAD_REL=__UPLOAD_REL__;
var HUB_MODE=__HUB_MODE__;
var BUILD=__BUILD__;
function si(n,s,c){return'<span class="svg-icon'+(c?' '+c:'')+'" style="width:'+s+'px;height:'+s+'px">'+(ICONS[n]||ICONS.file)+'</span>'}
function dlPath(rel){if(!rel)return'';return'/api/download/'+rel.split('/').map(function(p){return encodeURIComponent(p)}).join('/')}
function thumbPath(rel,w){
  if(!rel)return'';
  w=w||480;
  return'/api/thumb/'+rel.split('/').map(function(p){return encodeURIComponent(p)}).join('/')+'?w='+w;
}
function chatImgUrl(url,name){
  url=fixImgUrl(url);
  if(isHeic(name||''))return url;
  var fn=name||'';
  try{if(!fn)fn=decodeURIComponent((url.split('/').pop()||'').split('?')[0])}catch(_){fn=(url.split('/').pop()||'').split('?')[0]}
  if(!isImg(fn))return url;
  if(url.indexOf('/api/download/')!==0)return url;
  return thumbPath(decodeURIComponent(url.slice('/api/download/'.length)),480);
}
function imgThumbFallback(el){
  if(!el)return;
  if(el.dataset.fallback==='1'){el.classList.remove('show');return}
  el.dataset.fallback='1';
  var full=el.getAttribute('data-full');
  if(full&&el.src!==full){el.src=full;return}
  el.classList.remove('show');
}
function listImgHtml(rel,du){
  if(!rel||!du)return'';
  var t=thumbPath(rel,120);
  return '<img class="fp show" src="'+t+'" loading="lazy" decoding="async" data-full="'+du+'" onerror="imgThumbFallback(this)">';
}
function showDocxError(el,name,msg){
  el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><p style="color:var(--rd);margin:8px 0;line-height:1.6">'+
    eh(msg||'无法预览此文档')+'</p><p style="color:var(--ad);font-size:12px;margin:8px 0">常见原因：上传中断、文件损坏。请重新发送一份。</p>'+
    '<button onclick="dlPreview()" style="background:var(--am);color:#0a0a0c;border:none;padding:8px 16px;font-family:var(--fm);cursor:pointer;margin-top:4px">另存副本</button></div>';
}
function fixImgUrl(u){
  if(!u)return'';
  u=u.trim();
  if(u.indexOf('/api/downloaddev/')===0)u='/api/download/dev/'+u.slice(16);
  if(u.indexOf('/api/download')===0){
    try{
      var dec=decodeURIComponent(u);
      var m=dec.match(/\/api\/download\/(?:下载|\u4e0b\u8f7d)\/from-phone\/(.+)$/i);
      if(m)return dlPath(UPLOAD_REL+'/'+m[1]);
      m=dec.match(/\/api\/download\/from-phone\/(.+)$/i);
      if(m&&UPLOAD_REL!=='from-phone')return dlPath(UPLOAD_REL+'/'+m[1]);
    }catch(e){}
    return u;
  }
  if(u.indexOf('from-phone/')===0)return dlPath(UPLOAD_REL+'/'+u.slice(11));
  if(u.indexOf('下载/from-phone/')===0||u.indexOf('\u4e0b\u8f7d/from-phone/')===0){
    var sub=u.split('from-phone/')[1]||'';
    if(sub)return dlPath(UPLOAD_REL+'/'+sub);
  }
  if(u.charAt(0)!=='/'&&u.indexOf('download/')<0)return dlPath(u);
  return u;
}
function attachBasename(u){
  if(!u)return'';
  u=fixImgUrl(u);
  try{
    var p=decodeURIComponent((u.split('?')[0]||''));
    return(p.split('/').pop()||'').toLowerCase();
  }catch(e){return((u.split('/').pop()||'').split('?')[0]).toLowerCase()}
}
function imgFromContent(c){
  if(!c)return null;
  if(c.indexOf('[[IMG]]')===0)return fixImgUrl(c.slice(7).split('\n')[0]);
  if(c.indexOf('[[VID]]')===0)return fixImgUrl(c.slice(7).split('\n')[0]);
  if(c.indexOf('[[FILE]]')===0)return fixImgUrl(c.slice(8).split('\n')[0]);
  return null;
}
function isAttachMsg(c){return c&&(c.indexOf('[[IMG]]')===0||c.indexOf('[[VID]]')===0||c.indexOf('[[FILE]]')===0)}
var cp='',cr=false,sk='mtime',sa=false,ct='chat',uc=0,ci=false,pq=[],pf=null,uploadPipeline=0;
var pvGallery=[],pvIdx=0,pvTouchX=0,pvTouchY=0,pvHistory=false,galleryImageEntries=[];
var sfxOn=true,notifyOk=false;
var selMode=false,selMap={};

// ═══ Sound FX (Web Audio API beeps) ═══
var ac=null;
function _ac(){if(!ac)try{ac=new(window.AudioContext||window.webkitAudioContext)()}catch(e){}return ac}
function beep(freq,dur,vol,type){
  type=type||'square';if(!sfxOn||!_ac())return;
  var o=_ac().createOscillator(),g=_ac().createGain(),f=_ac().createBiquadFilter();
  o.type=type;o.frequency.value=freq;g.gain.setValueAtTime(vol*.12,_ac().currentTime);
  g.gain.exponentialRampToValueAtTime(.001,_ac().currentTime+dur);
  f.type='lowpass';f.frequency.value=freq*2;
  o.connect(f);f.connect(g);g.connect(_ac().destination);o.start();o.stop(_ac().currentTime+dur);
}
function noise(dur,vol){
  if(!sfxOn||!_ac())return;var bs=_ac().sampleRate*dur,bf=_ac().createBuffer(1,bs,_ac().sampleRate),d=bf.getChannelData(0);
  for(var i=0;i<bs;i++)d[i]=Math.random()*2-1;
  var s=_ac().createBufferSource(),g=_ac().createGain(),f=_ac().createBiquadFilter();
  s.buffer=bf;f.type='bandpass';f.frequency.value=800;f.Q.value=.5;
  g.gain.setValueAtTime(vol*.08,_ac().currentTime);g.gain.exponentialRampToValueAtTime(.001,_ac().currentTime+dur);
  s.connect(f);f.connect(g);g.connect(_ac().destination);s.start();s.stop(_ac().currentTime+dur);
}
function sfxClick(){beep(200,.06,.25,'triangle');noise(.03,.15);setTimeout(function(){beep(600,.04,.2,'square')},30)}
function sfxOpen(){beep(150,.08,.3,'triangle');noise(.04,.2);setTimeout(function(){beep(400,.06,.25,'square');noise(.02,.15)},40);setTimeout(function(){beep(900,.04,.2,'square')},70)}
function sfxSend(){noise(.05,.25);beep(120,.08,.3,'triangle');setTimeout(function(){beep(300,.05,.2,'square')},40);setTimeout(function(){beep(600,.04,.15,'square')},70);setTimeout(function(){beep(1000,.03,.1,'triangle')},100)}
function toggleSfx(){sfxOn=!sfxOn;var b=document.getElementById('sfx-btn');b.classList.toggle('on',sfxOn);if(sfxOn){beep(600,.03,.15,'triangle');beep(1000,.04,.2,'square')}}

function initActFeedback(){
  var ACT='button,.lan-bar-url,.fr:not(.dir),.yg-item,.sel-chip';
  var pressed=null;
  document.addEventListener('pointerdown',function(e){
    if(e.button>0)return;
    var el=e.target.closest(ACT);
    if(!el||el.disabled||el.getAttribute('aria-disabled')==='true')return;
    pressed=el;
    el.classList.add('act-press');
  },{passive:true});
  function release(e){
    if(!pressed)return;
    var el=pressed;
    pressed=null;
    el.classList.remove('act-press');
    if(!e||e.type==='pointerup'){
      el.classList.remove('act-hit');
      void el.offsetWidth;
      el.classList.add('act-hit');
      setTimeout(function(){el.classList.remove('act-hit')},450);
      if(sfxOn&&!el.classList.contains('tab-btn')&&el.id!=='csb'&&!el.classList.contains('sel-x')&&!el.classList.contains('sel-dl'))
        beep(620,.022,.16,'triangle');
    }
  }
  document.addEventListener('pointerup',release,{passive:true});
  document.addEventListener('pointercancel',release,{passive:true});
}

// ═══ Utils ═══
function $(s){return document.getElementById(s)}
function fmt(b){if(b<1024)return b+' B';if(b<1048576)return(b/1024).toFixed(1)+' KB';if(b<1073741824)return(b/1048576).toFixed(1)+' MB';return(b/1073741824).toFixed(2)+' GB'}
function eh(s){var d=document.createElement('div');d.textContent=s;return d.innerHTML}
var tt;function toast(m){var t=$('toast');t.textContent=m;t.classList.add('show');clearTimeout(tt);tt=setTimeout(function(){t.classList.remove('show')},1800)}
function copyUrl(){copyLanUrl()}
function copyLanUrl(){
  sfxClick();
  var url=(($('lan-bar-url')&&$('lan-bar-url').textContent)||($('lu')&&$('lu').textContent)||location.origin).replace(/\/+$/,'');
  cpTx(url);
  toast('已复制 · 手机浏览器粘贴此地址');
}
function applyHubInfo(d){
  if(!d)return;
  var url=(d.lan_url||location.origin).replace(/\/+$/,'')+'/';
  $('lu').textContent=url;
  if(d.mode)HUB_MODE=d.mode;
  if(d.upload_dir)HUB_UPLOAD_DIR=d.upload_dir;
  if(d.local_hub_client)HUB_LOCAL_CLIENT=!!d.local_hub_client;
  document.body.classList.remove('hub-lan','hub-cloud');
  document.body.classList.add(HUB_MODE==='cloud'?'hub-cloud':'hub-lan');
  var zh=$('hub-title-zh'),en=$('hub-title-en'),badge=$('hub-badge');
  if(HUB_MODE==='cloud'){
    if(zh)zh.textContent='云传输站';
    if(en)en.textContent='Relay Cloud';
    if(badge){badge.textContent='CLOUD';badge.className='mode-badge mode-cloud'}
    document.title='Relay Cloud · 云传输站';
  }else{
    if(zh)zh.textContent='局域网传输站';
    if(en)en.textContent='Relay Local';
    if(badge){badge.textContent='LOCAL';badge.className='mode-badge mode-lan'}
    document.title='Relay Local · 局域网传输站';
  }
  var lanBar=$('lan-bar'),lanBtn=$('lan-bar-url'),cloudBar=$('cloud-bar'),cloudUrl=$('cloud-bar-url');
  if(HUB_MODE==='lan'&&lanBar&&!im){
    lanBar.classList.add('show');
    if(lanBtn)lanBtn.textContent=url.replace(/\/$/,'');
  }else if(lanBar){lanBar.classList.remove('show')}
  if(HUB_MODE==='cloud'&&cloudBar){
    cloudBar.classList.add('show');
    if(cloudUrl){
      try{cloudUrl.textContent=new URL(url).host}catch(_){cloudUrl.textContent='cloud hub'}
    }
  }else if(cloudBar){cloudBar.classList.remove('show')}
  updFileSrcUI();
  updSearchHint();
  if(typeof d.msg_epoch==='number')applyMsgEpoch(d.msg_epoch);
}
function updSearchHint(){
  var h=$('search-hint'),sf=$('search-scope-files');
  if(!h)return;
  if(HUB_MODE==='cloud'){
    h.textContent='搜索全部聊天记录（云端永久保存）· 云收件箱文件名';
    if(sf)sf.textContent='云文件';
  }else{
    h.textContent='搜索聊天记录 + 本机 ~/ 下文件名（跳过 .git 等目录）';
    if(sf)sf.textContent='本机文件';
  }
}
function cpTxFallback(t){
  var ta=document.createElement('textarea');
  ta.value=t;
  ta.setAttribute('readonly','');
  ta.style.cssText='position:fixed;left:-9999px;top:0;opacity:0';
  document.body.appendChild(ta);
  ta.focus();ta.select();ta.setSelectionRange(0,t.length);
  var ok=false;
  try{ok=document.execCommand('copy')}catch(e){}
  ta.remove();
  toast(ok?'已复制':'复制失败，请手动选中文字');
  return ok;
}
function cpTx(t){
  if(!t&&t!==0){toast('无内容');return}
  t=String(t);
  if(navigator.clipboard&&window.isSecureContext){
    navigator.clipboard.writeText(t).then(function(){toast('已复制')}).catch(function(){cpTxFallback(t)});
  }else{
    cpTxFallback(t);
  }
}
function isImg(n){return/\.(jpg|jpeg|png|gif|webp|bmp|svg|heic|heif)$/i.test(n)}
function isVid(n){return/\.(mp4|webm|mov|avi|mkv|m4v|3gp|hevc|mpeg|mpg)$/i.test(n)}
function isVideoFile(f){
  if(!f)return false;
  if(f.type&&f.type.indexOf('video/')===0)return true;
  return isVid(f.name||'');
}
function fmtSize(b){
  if(!b||b<1024)return'';
  if(b<1048576)return(b/1024).toFixed(0)+'KB';
  return(b/1048576).toFixed(1)+'MB';
}
function isHeic(n){return/\.(heic|heif)$/i.test(n)}
function absUrl(u){
  u=fixImgUrl(u||'').split('?')[0];
  if(!u)return location.origin;
  if(u.indexOf('http')===0)return u;
  return location.origin+(u.charAt(0)==='/'?u:'/'+u);
}
function fileNameFrom(url,name){
  if(name)return name;
  try{return decodeURIComponent((fixImgUrl(url).split('/').pop()||'file').split('?')[0])}catch(_){return(fixImgUrl(url).split('/').pop()||'file').split('?')[0]}
}
async function blobToPng(blob){
  if(blob.type==='image/png')return blob;
  return new Promise(function(res,rej){
    var img=new Image(),obj=URL.createObjectURL(blob);
    img.onload=function(){
      URL.revokeObjectURL(obj);
      var c=document.createElement('canvas');
      c.width=img.naturalWidth||img.width;c.height=img.naturalHeight||img.height;
      c.getContext('2d').drawImage(img,0,0);
      c.toBlob(function(b){b?res(b):rej(new Error('toBlob'))},'image/png');
    };
    img.onerror=function(){URL.revokeObjectURL(obj);rej(new Error('img'))};
    img.src=obj;
  });
}
async function fetchImageBlob(url,name){
  url=fixImgUrl(url);
  var r=await fetch(url);
  if(!r.ok)throw new Error('HTTP '+r.status);
  var blob=await r.blob();
  name=name||fileNameFrom(url,'');
  if(isHeic(name)||blob.type==='image/heic'||blob.type==='image/heif'){
    if(typeof heic2any==='undefined')return blob;
    var out=await heic2any({blob:blob,toType:'image/jpeg',quality:0.92});
    blob=Array.isArray(out)?out[0]:out;
  }
  return blob;
}
async function cpImage(url,name){
  toast('复制图片中…');
  try{
    var blob=await fetchImageBlob(url,name);
    var png=await blobToPng(blob);
    if(navigator.clipboard&&window.ClipboardItem&&window.isSecureContext){
      await navigator.clipboard.write([new ClipboardItem({'image/png':png})]);
      toast('已复制图片 · 可直接粘贴到微信/文档');
      return true;
    }
  }catch(e){console.error('cpImage',e)}
  var p=hubFilePath(url,name);
  if(p&&isLocalHub()){
    cpTx(p);
    toast('已复制图片路径 · 可在文件管理器打开');
    return true;
  }
  return false;
}
function triggerDownload(url,name){
  url=absUrl(fixImgUrl(url));
  name=name||fileNameFrom(url,'');
  var a=document.createElement('a');
  a.href=url;a.download=name;a.rel='noopener';
  document.body.appendChild(a);a.click();a.remove();
}
async function cpFile(url,name){
  url=fixImgUrl(url);
  name=name||fileNameFrom(url,'');
  var p=hubFilePath(url,name);
  if(p&&isLocalHub()){
    cpTx(p);
    toast('已复制文件路径 · 可在文件管理器打开');
    return true;
  }
  toast('复制文件中…');
  try{
    var r=await fetch(url,{credentials:'same-origin'});
    if(!r.ok)throw new Error('HTTP '+r.status);
    var blob=await r.blob();
    if(navigator.clipboard&&window.ClipboardItem&&window.isSecureContext){
      var mime=blob.type||'application/octet-stream';
      var item={};item[mime]=blob;
      await navigator.clipboard.write([new ClipboardItem(item)]);
      toast('已复制文件 · 可直接粘贴');
      return true;
    }
  }catch(e){console.error('cpFile',e)}
  triggerDownload(url,name);
  toast('已开始下载文件');
  return true;
}
async function cpSmart(url,name,text){
  url=url?fixImgUrl(url):'';
  name=name||fileNameFrom(url,'');
  if(text&&!url){
    cpTx(text);
    toast('已复制文字');
    return;
  }
  if(url&&isImg(name||url)){
    if(await cpImage(url,name))return;
  }
  if(url){
    if(await cpFile(url,name))return;
  }
  if(text){
    cpTx(text);
    toast('已复制文字');
    return;
  }
}
function copyBtnTitle(url,name,text){
  if(text&&!url)return'复制文字';
  name=name||fileNameFrom(url,'');
  if(url&&isImg(name||url))return'复制图片';
  if(isAud(name))return'复制音频';
  if(isVid(name))return'复制视频';
  if(hubFilePath(url,name)&&isLocalHub())return'复制文件路径';
  return'复制文件';
}
function autoSender(){return im?'phone':'pc'}
function autoSenderName(){return im?'手机':'电脑'}
function queueBusy(){return uploadPipeline>0}
async function loadHeicInto(img,url){
  if(typeof heic2any==='undefined'){img.alt='HEIC';return}
  try{
    var r=await fetch(url);if(!r.ok)throw new Error('HTTP '+r.status);
    var blob=await r.blob();
    var out=await heic2any({blob:blob,toType:'image/jpeg',quality:0.85});
    var b=Array.isArray(out)?out[0]:out;
    img.src=URL.createObjectURL(b);
  }catch(e){img.style.display='none';var p=img.parentElement;if(p)p.innerHTML='<div class="mb" style="cursor:pointer;color:var(--am)">'+si('photo',14)+' HEIC · 点击预览</div>'}
}
function isAud(n){return/\.(mp3|wav|flac|aac|ogg)$/i.test(n)}
function isDoc(n){var l=n.toLowerCase();return/\.(pdf|doc|docx|xls|xlsx|ppt|pptx|md|txt|csv|json|xml|yml|yaml|log)$/i.test(l)||/\.(py|js|ts|html|css|java|c|cpp|rs|go)$/i.test(l)}
function isMd(n){return/\.md$/i.test(n)}
function isHtml(n){return/\.html?$/i.test(n)}
function isTxt(n){return/\.(txt|log|json|xml|csv|yml|yaml)$/i.test(n)||/\.(py|js|ts|css|java|c|cpp|rs|go)$/i.test(n)}
function catOf(n){if(isImg(n))return'image';if(isVid(n))return'video';if(isAud(n))return'audio';if(isDoc(n))return'doc';return'other'}
function isPdf(n){return/\.pdf$/i.test(n)}
function isDocx(n){return/\.docx$/i.test(n)}
function isZip(n){return/\.(zip|rar|7z|tar|gz|tgz|bz2)$/i.test(n||'')}
function zipBody(url,fn){
  var body=document.createElement('div');
  var ext=(fn.split('.').pop()||'zip').toUpperCase();
  body.className='mb mzip';
  body.innerHTML='<div class="zip-icon">'+si('archive',30)+'</div><div class="zip-body"><div class="zip-title">'+eh(fn)+'</div><div class="zip-hint">压缩包 · 点击下载</div><span class="zip-tag">'+eh(ext)+'</span></div>';
  body.addEventListener('click',function(e){if(selMode)return;dlOne(url,fn)});
  addCopyBtn(body,url,fn);
  return body;
}
function fileKind(n){
  if(isImg(n))return'img';if(isVid(n))return'vid';if(isAud(n))return'aud';
  if(isMd(n))return'md';if(isHtml(n))return'html';if(isTxt(n))return'txt';if(isPdf(n))return'pdf';
  if(isDocx(n))return'docx';
  return'other';
}
function previewFile(url,name,gallery,idx){
  url=fixImgUrl(url);
  name=name||decodeURIComponent((url.split('/').pop()||'file').split('?')[0]);
  var k=fileKind(name);
  if(k==='img'){
    if(!gallery||!gallery.length){
      gallery=ct==='chat'?collectChatImages():collectFolderImages();
      idx=-1;
      for(var i=0;i<gallery.length;i++){if(gallery[i].url===url||fixImgUrl(gallery[i].url)===url){idx=i;break}}
      if(idx<0){gallery=gallery.concat([{url:url,name:name}]);idx=gallery.length-1}
    }
    pvGallery=gallery;pvIdx=idx>=0?idx:0;
  }else{pvGallery=[];pvIdx=0}
  openFilePreview(url,name,k==='img',k==='vid',k==='aud',k==='md',k==='html',k==='txt',k==='pdf',k==='docx');
}
function collectChatImages(){
  var list=[],seen={};
  document.querySelectorAll('#ca .mr[data-attach-url]').forEach(function(row){
    var u=fixImgUrl(row.getAttribute('data-attach-url')||'');
    if(!u||seen[u])return;
    var fn='';
    try{fn=decodeURIComponent((u.split('/').pop()||'img').split('?')[0])}catch(_){fn=(u.split('/').pop()||'img').split('?')[0]}
    if(isImg(fn)){seen[u]=1;list.push({url:u,name:fn})}
  });
  return list;
}
function collectFolderImages(){
  if(galleryImageEntries.length)return galleryImageEntries.slice();
  var list=[];
  document.querySelectorAll('#fl .fr:not(.dir)').forEach(function(row){
    var u=row.getAttribute('data-url'),n=row.getAttribute('data-name');
    if(u&&n&&isImg(n))list.push({url:fixImgUrl(u),name:n});
  });
  return list;
}
function rebuildGalleryEntries(es,isRx){
  galleryImageEntries=es.filter(function(e){
    return !e.is_dir&&isImg(e.name);
  }).map(function(e){
    var rel=e.relpath||(e.is_dir?'':((cp?cp+'/':'')+e.name));
    return {url:fixImgUrl(dlPath(rel)),name:e.name};
  });
}
function galleryNav(dir){
  if(!pvGallery.length)return;
  stopMedia();
  pvIdx=(pvIdx+dir+pvGallery.length)%pvGallery.length;
  var it=pvGallery[pvIdx];
  var k=fileKind(it.name);
  openFilePreview(it.url,it.name,k==='img',k==='vid',k==='aud',false,false,false,false,false);
}
function updGalleryNav(kind){
  var show=(kind==='img'||kind==='vid')&&pvGallery.length>1;
  ['pv-prev-d','pv-next-d'].forEach(function(id){var e=$(id);if(e)e.style.display=show&&!im?'flex':'none'});
  ['pv-edge-l','pv-edge-r'].forEach(function(id){var e=$(id);if(e)e.style.display=show?'flex':'none'});
  if(!im){['pv-prev-m','pv-next-m'].forEach(function(id){var e=$(id);if(e)e.style.display='none'})}
  var cnt=(pvIdx+1)+' / '+pvGallery.length;
  var cd=$('pv-count-d'),cm=$('pv-count-m');
  if(cd){cd.style.display=show&&!im?'block':'none';cd.textContent=cnt}
  if(cm){cm.style.display=show?'inline':'none';cm.textContent=cnt}
}

function fi(n){if(isVid(n))return si('video',14);if(isImg(n))return si('image',14);if(isAud(n))return si('audio',14);if(isZip(n))return si('archive',14);if(/\.pdf$/i.test(n))return si('file',14);if(/\.(py|js|ts|html|css|json|md)$/i.test(n))return si('file',14);return si('file',14)}

// ═══ Tabs ═══
function sw(t){
  if(selMode)exitSelMode();
  ct=t;document.querySelectorAll('.tab-btn').forEach(function(b){b.classList.toggle('active',b.dataset.tab===t)});
  document.querySelectorAll('.tab-panel').forEach(function(p){p.classList.toggle('active',p.id==='panel-'+t)});
  if(t==='chat'){
    uc=0;ubd();
    if(chatNeedsRefresh){
      chatNeedsRefresh=false;
      fm(null).then(function(){scb(true)});
    }else requestAnimationFrame(function(){requestAnimationFrame(function(){scb(true)})});
  }
  if(t==='files'&&!$('fl').querySelector('.fr,.yg-item'))loadFilesDefault();
}
document.querySelectorAll('.tab-btn').forEach(function(b){b.addEventListener('click',function(){sfxClick();sw(b.dataset.tab)})});
function ubd(){var b=$('ub');b.textContent=uc>0?uc:''}

// ═══ Search ═══
var searchScope='all',searchTimer=null,searchBusy=false;
var lastSearchFiles=[],lastSearchMsgs=[];
function openSearch(){
  var ov=$('search-overlay');if(!ov)return;
  ov.classList.add('show');
  document.body.classList.add('search-open');
  var inp=$('search-input');
  if(inp){inp.focus();if(inp.value.trim())runSearch(inp.value.trim())}
  sfxOpen();
}
function closeSearch(){
  var ov=$('search-overlay');if(!ov)return;
  ov.classList.remove('show');
  document.body.classList.remove('search-open');
}
function hlSnippet(snippet,q){
  if(!snippet||!q)return eh(snippet||'');
  var low=snippet.toLowerCase(),ql=q.toLowerCase(),idx=low.indexOf(ql);
  if(idx<0)return eh(snippet);
  return eh(snippet.slice(0,idx))+'<mark>'+eh(snippet.slice(idx,idx+q.length))+'</mark>'+eh(snippet.slice(idx+q.length));
}
function renderSearchResults(d){
  var box=$('search-results');if(!box)return;
  if(!d||!d.ok){box.innerHTML='<div class="search-empty">搜索失败</div>';return}
  var q=d.q||'',ms=d.messages||[],fs=d.files||[],html='';
  if(!ms.length&&!fs.length){
    box.innerHTML='<div class="search-empty">没有找到「'+eh(q)+'」</div>';
    return;
  }
  if(ms.length){
    html+='<div class="search-group">聊天 · '+ms.length+' 条</div>';
    ms.forEach(function(m){
      var who=m.sender==='pc'?'电脑':'手机';
      var label=m.kind==='file'?'[文件] ':'';
      html+='<button type="button" class="search-item" data-kind="msg" data-mid="'+m.id+'">'+
        '<div class="sr-snippet">'+label+hlSnippet(m.snippet,q)+'</div>'+
        '<div class="sr-meta">'+eh(who)+' · '+eh(m.time||'')+'</div></button>';
    });
  }
  if(fs.length){
    html+='<div class="search-group">文件 · '+fs.length+' 个'+(d.truncated?'（部分）':'')+'</div>';
    fs.forEach(function(f,i){
      html+='<button type="button" class="search-item" data-kind="file" data-fi="'+i+'">'+
        '<div class="sr-snippet">'+eh(f.name||'')+'</div>'+
        '<div class="sr-meta">'+eh(f.relpath||'')+' · '+fmtSize(f.size)+'</div></button>';
    });
  }
  box.innerHTML=html;
  lastSearchFiles=fs;
  lastSearchMsgs=ms;
}
function scrollToMsg(row){
  if(!row)return;
  row.classList.add('msg-flash');
  row.scrollIntoView({behavior:'smooth',block:'center'});
  setTimeout(function(){row.classList.remove('msg-flash')},2400);
}
function expandMsgRow(row){
  if(!row)return;
  var inner=row.querySelector('.mb-collapsed');
  if(inner){
    inner.classList.remove('mb-collapsed');
    var wrap=inner.parentElement;
    if(wrap)wrap.classList.add('mb-expanded');
    var btns=wrap?wrap.querySelectorAll('.expand-btn'):[];
    btns.forEach(function(b){b.textContent=b.classList.contains('expand-top')?'↑ 收起':'收起'});
  }
}
async function jumpToMessage(mid){
  if(!mid||isNaN(mid))return;
  closeSearch();
  sw('chat');
  chatFilter='all';
  var cfAll=$('cf-all'),cfText=$('cf-text');
  if(cfAll)cfAll.classList.add('on');
  if(cfText)cfText.classList.remove('on');
  applyChatFilter();
  toast('定位消息中…');
  await new Promise(function(r){requestAnimationFrame(function(){requestAnimationFrame(r)})});
  var row=document.querySelector('#ca .mr[data-mid="'+mid+'"]');
  if(!row){
    try{
      var wait=0;
      while(fetching&&wait<40){await new Promise(function(r){setTimeout(r,50)});wait++}
      var r=await fetch('/api/messages?meta=1&anchor='+mid);
      var d=await r.json();
      var ms=(d&&d.messages)?d.messages:[];
      if(!ms.length){toast('消息不存在或已删除');return}
      ram(ms);
      if(d&&typeof d.total==='number'){msgTotal=d.total;msgHasMore=!!d.has_more;msgOldest=d.oldest_id||0}
      if(d&&typeof d.epoch==='number')msgEpoch=d.epoch;
      if(ms.length)lid=ms[ms.length-1].id;
      updLoadOlder();
      applyChatFilter();
      await new Promise(function(r){requestAnimationFrame(r)});
      row=document.querySelector('#ca .mr[data-mid="'+mid+'"]');
    }catch(e){
      console.error('jumpToMessage',e);
      toast('定位失败，请重试');
      return;
    }
  }
  if(row){
    expandMsgRow(row);
    scrollToMsg(row);
    toast('已定位到消息');
  }else toast('无法定位到该消息');
}
async function runSearch(q){
  if(!q||q.length<1){renderSearchResults(null);return}
  if(searchBusy)return;
  searchBusy=true;
  var box=$('search-results');
  if(box)box.innerHTML='<div class="search-empty"><div class="spinner"></div> 搜索中…</div>';
  try{
    var r=await fetch('/api/search?q='+encodeURIComponent(q)+'&scope='+encodeURIComponent(searchScope));
    var d=await r.json();
    renderSearchResults(d);
  }catch(e){
    if(box)box.innerHTML='<div class="search-empty">搜索失败</div>';
  }
  searchBusy=false;
}
function openSearchFile(rel,name){
  if(!rel)return;
  closeSearch();
  sw('files');
  var parts=rel.split('/').filter(Boolean);
  parts.pop();
  var dir=parts.join('/');
  var isInbox=rel.indexOf((UPLOAD_REL||'from-phone')+'/')===0||rel.indexOf('from-phone/')===0;
  if(isInbox){fileSrc='inbox';updFileSrcUI();bf('',true)}
  else{fileSrc='home';updFileSrcUI();bf(dir,false)}
  setTimeout(function(){previewFile(dlPath(rel),name||rel.split('/').pop())},400);
}
function initSearchUI(){
  var btn=$('search-btn'),close=$('search-close'),inp=$('search-input'),ov=$('search-overlay');
  var results=$('search-results');
  if(results&&!results._jumpBound){
    results._jumpBound=true;
    function onPick(e){
      var item=e.target.closest('.search-item[data-kind]');
      if(!item)return;
      e.preventDefault();
      e.stopPropagation();
      sfxClick();
      if(item.getAttribute('data-kind')==='msg'){
        jumpToMessage(parseInt(item.getAttribute('data-mid'),10));
      }else{
        var fi=parseInt(item.getAttribute('data-fi'),10);
        var f=lastSearchFiles[fi];
        if(f)openSearchFile(f.relpath,f.name);
      }
    }
    results.addEventListener('click',onPick);
    results.addEventListener('touchend',function(e){onPick(e)},{passive:false});
  }
  if(btn)btn.addEventListener('click',function(){sfxClick();openSearch()});
  if(close)close.addEventListener('click',function(){sfxClick();closeSearch()});
  if(ov)ov.addEventListener('click',function(e){if(e.target===ov)closeSearch()});
  if(inp){
    inp.addEventListener('input',function(){
      var q=inp.value.trim();
      if(searchTimer)clearTimeout(searchTimer);
      searchTimer=setTimeout(function(){runSearch(q)},280);
    });
    inp.addEventListener('keydown',function(e){
      if(e.key==='Enter'){e.preventDefault();runSearch(inp.value.trim())}
      if(e.key==='Escape')closeSearch();
    });
  }
  document.querySelectorAll('#search-scope button').forEach(function(b){
    b.addEventListener('click',function(){
      sfxClick();
      searchScope=b.getAttribute('data-scope')||'all';
      document.querySelectorAll('#search-scope button').forEach(function(x){x.classList.toggle('on',x===b)});
      var q=$('search-input');if(q&&q.value.trim())runSearch(q.value.trim());
    });
  });
  document.addEventListener('keydown',function(e){
    if((e.ctrlKey||e.metaKey)&&e.key==='k'){e.preventDefault();openSearch()}
    if(e.key==='Escape'&&$('search-overlay')&&$('search-overlay').classList.contains('show'))closeSearch();
  });
  updSearchHint();
}

// ═══ File preview pane ═══
var currentCat='all';
function renderPreview(el,url,name,isI,isV,isA,isM,isH,isT,isP,isDocx){
  el.innerHTML='';
  var k=fileKind(name);
  if(k==='docx')isDocx=true;
  if(k==='pdf')isP=true;
  if(k==='md')isM=true;
  if(k==='html')isH=true;
  if(k==='txt')isT=true;
  if(k==='img')isI=true;
  if(k==='vid')isV=true;
  if(k==='aud')isA=true;
  var heic=isHeic(name);
  if(isI||heic){
    var img=document.createElement('img');
    img.style.cssText='max-width:100%;max-height:70vh;object-fit:contain';
    if(heic)loadHeicInto(img,url);
    else{
      var rel=relFromUrl(url),pvSrc=rel?thumbPath(rel,960):url;
      img.src=pvSrc;
      img.onerror=function(){
        if(!img.dataset.retry&&url){img.dataset.retry='1';img.src=url;return}
        el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><p style="color:var(--rd)">图片加载失败</p></div>';
      };
    }
    el.appendChild(img);return'img';
  }
  if(isV){
    var pv=document.createElement('video');
    pv.id='preview-video';pv.src=url;pv.controls=true;pv.playsInline=true;pv.muted=false;
    pv.setAttribute('playsinline','');pv.setAttribute('webkit-playsinline','');
    pv.style.cssText='max-width:100%;max-height:70vh;background:#000';
    el.appendChild(pv);return'vid';
  }
  if(isA){el.innerHTML='<audio id="preview-audio" src="'+url+'" controls style="width:100%;max-width:320px"></audio>';return'aud'}
  if(isM){loadMdToEl(el,url,name);return'md'}
  if(isH){el.innerHTML='<iframe src="'+url+'" sandbox="allow-scripts allow-same-origin allow-forms" style="width:100%;min-height:65vh;border:1px solid var(--bd);background:#fff"></iframe>';return'html'}
  if(isP){el.innerHTML='<iframe src="'+url+'" style="width:100%;min-height:65vh;border:1px solid var(--bd);background:#fff"></iframe>';return'pdf'}
  if(isDocx){loadDocxToEl(el,url,name);return'docx'}
  if(isT){loadTxtToEl(el,url,name);return'txt'}
  /* fallback: try as image, then offer download */
  var img2=document.createElement('img');
  img2.src=url;img2.style.cssText='max-width:100%;max-height:70vh;object-fit:contain';
  img2.onload=function(){};
  img2.onerror=function(){
    el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><p style="color:var(--ad);margin:8px 0">此格式暂不支持内嵌预览</p><button onclick="dlPreview()" style="background:var(--am);color:#0a0a0c;border:none;padding:8px 16px;font-family:var(--fm);cursor:pointer">下载</button></div>';
  };
  el.appendChild(img2);return'other';
}
function openFilePreview(url,name,isI,isV,isA,isM,isH,isT,isP,isDocx){
  url=fixImgUrl(url);
  pf={url:url,name:name};
  if(isHeic(name))isI=true;
  if(!pvHistory && (im||ct==='chat')){history.pushState({pv:1},'',location.href);pvHistory=true;document.body.classList.add('preview-open')}
  var useOverlay=im||ct==='chat';
  if(useOverlay){
    var mo=$('mo');
    mo.style.display='flex';
    if(!im)mo.classList.add('show-desktop');
    $('mfn').textContent=name;
    var inner=$('mbd-inner');if(inner)inner.innerHTML='';
    var k=renderPreview(inner||$('mbd'),url,name,isI,isV,isA,isM,isH,isT,isP,isDocx);
    $('mdl').style.display='flex';
    var pbb=$('pv-batch-dl');
    if(pbb)pbb.style.display='none';
    $('send-preview-btn-m').style.display=(k==='img')?'inline-block':'none';
    updDlButtons(url,name);
    updGalleryNav(k==='img'||k==='vid'?k:'other');return;
  }
  $('pp').innerHTML='';$('pa').style.display='flex';
  var k=renderPreview($('pp'),url,name,isI,isV,isA,isM,isH,isT,isP,isDocx);
  $('stop-preview-btn').style.display=(k==='vid'||k==='aud')?'inline-block':'none';
  $('send-preview-btn').style.display=(k==='img')?'inline-block':'none';
  updDlButtons(url,name);
  updGalleryNav(k);
}
function closeMobilePreview(){
  $('mo').style.display='none';
  $('mo').classList.remove('show-desktop');
  stopMedia();
  document.body.classList.remove('preview-open');
  pvHistory=false;
}
function closePreviewBtn(){
  if($('mo').style.display==='flex'&&pvHistory)history.back();
  else closeMobilePreview();
}
function bindPreviewBackdrop(){
  var mo=$('mo');
  if(!mo||mo._bdBound)return;
  mo._bdBound=true;
  mo.addEventListener('click',function(e){
    if(mo.style.display!=='flex')return;
    if(e.target.closest('.ph,#mdl,.pv-edge,.pv-arrow,.preview-close,button,a'))return;
    if(e.target.closest('#mbd-inner img,#mbd-inner video,#mbd-inner iframe,.docx-body,.md-preview,.md-body'))return;
    closePreviewBtn();
  });
}
window.addEventListener('popstate',function(){
  if($('mo').style.display==='flex')closeMobilePreview();
});
function bindGalleryNav(){
  ['pv-prev-d','pv-prev-m','pv-edge-l'].forEach(function(id){var el=$(id);if(el)el.addEventListener('click',function(e){e.stopPropagation();galleryNav(-1)})});
  ['pv-next-d','pv-next-m','pv-edge-r'].forEach(function(id){var el=$(id);if(el)el.addEventListener('click',function(e){e.stopPropagation();galleryNav(1)})});
  var swipe=function(el){
    if(!el)return;
    el.addEventListener('touchstart',function(e){pvTouchX=e.changedTouches[0].clientX;pvTouchY=e.changedTouches[0].clientY},{passive:true});
    el.addEventListener('touchend',function(e){
      if(!pvGallery.length||selMode)return;
      var t=e.changedTouches[0];
      var dx=t.clientX-pvTouchX,dy=t.clientY-pvTouchY;
      if(Math.abs(dx)<40||Math.abs(dx)<Math.abs(dy))return;
      galleryNav(dx<0?1:-1);
    },{passive:true});
  };
  swipe($('mbd'));swipe($('mbd-inner'));swipe($('pp'));
  document.addEventListener('keydown',function(e){
    if($('mo').style.display==='flex'&&(e.key==='Escape'||e.key==='Esc')){
      e.preventDefault();closePreviewBtn();return;
    }
    if(!pf||!pvGallery.length)return;
    if(e.key==='ArrowLeft')galleryNav(-1);
    if(e.key==='ArrowRight')galleryNav(1);
  });
}
async function loadMdToEl(el,url,name){
  try{var r=await fetch(url);if(!r.ok)throw new Error('HTTP '+r.status);var t=await r.text();el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2>'+mdToHtml(t)+'</div>'}
  catch(e){el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><p style="color:var(--rd)">Failed to load</p></div>'}
}
async function loadTxtToEl(el,url,name){
  try{var r=await fetch(url);if(!r.ok)throw new Error('HTTP '+r.status);var t=await r.text();el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><pre class="txt-preview">'+eh(t)+'</pre></div>'}
  catch(e){el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><p style="color:var(--rd)">Failed to load</p></div>'}
}
async function loadDocxToEl(el,url,name){
  el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><p style="color:var(--ad)">加载 Word 文档…</p></div>';
  var pvUrl='';
  url=fixImgUrl(url).split('?')[0];
  if(url.indexOf('/api/download/')===0)pvUrl='/api/preview/docx/'+url.slice('/api/download/'.length);
  try{
    if(pvUrl){
      var pr=await fetch(pvUrl);
      var pd={};
      try{pd=await pr.json()}catch(_){}
      if(pr.ok&&pd.ok&&pd.html){
        el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><div class="docx-body">'+pd.html+'</div></div>';
        return;
      }
      if(pd.error==='corrupt'||pd.message){
        showDocxError(el,name,pd.message||'文件已损坏，不是有效的 Word 文档');
        return;
      }
    }
    if(typeof mammoth!=='undefined'){
      var r=await fetch(url);if(!r.ok)throw new Error('HTTP '+r.status);
      var buf=await r.arrayBuffer();
      if(buf.byteLength<4||new Uint8Array(buf,0,2)[0]!==0x50||new Uint8Array(buf,0,2)[1]!==0x4B)
        throw new Error('corrupt');
      var out=await mammoth.convertToHtml({arrayBuffer:buf});
      el.innerHTML='<div class="md-preview"><h2>'+eh(name)+'</h2><div class="docx-body">'+out.value+'</div></div>';
      return;
    }
    throw new Error('no preview');
  }catch(e){
    console.error('docx preview',e);
    var msg=(e&&e.message==='corrupt')?'文件已损坏（上传可能中断），请重新发送':null;
    showDocxError(el,name,msg);
  }
}
function stopMedia(){
  var v=$('preview-video'),a=$('preview-audio');
  if(v){v.pause();v.currentTime=0;v.src=''}
  if(a){a.pause();a.currentTime=0;a.src=''}
}
var HUB_UPLOAD_DIR='';
var HUB_LOCAL_CLIENT=false;
function isHubHost(){
  var h=location.hostname;
  return h==='localhost'||h==='127.0.0.1'||h==='[::1]';
}
function isLocalHub(){
  return HUB_MODE==='lan'&&(HUB_LOCAL_CLIENT||isHubHost());
}
function hubFilePath(url,name){
  if(HUB_MODE!=='lan'||!HUB_UPLOAD_DIR)return '';
  var fn=name||'';
  if(!fn&&url){
    try{fn=decodeURIComponent((fixImgUrl(url).split('/').pop()||'').split('?')[0])}catch(_){fn=(fixImgUrl(url).split('/').pop()||'').split('?')[0]}
  }
  if(!fn)return '';
  return HUB_UPLOAD_DIR.replace(/\/$/,'')+'/'+fn;
}
function appendPathHint(body,url,name){
  var p=hubFilePath(url,name);
  if(!p||!isLocalHub())return;
  var hint=document.createElement('div');
  hint.className='file-path-hint';
  hint.textContent=p;
  hint.title='点击复制路径';
  hint.style.cursor='pointer';
  hint.addEventListener('click',function(e){e.stopPropagation();cpTx(p);toast('已复制路径')});
  body.appendChild(hint);
}
function updDlButtons(url,name){
  var local=isLocalHub()&&hubFilePath(url,name);
  document.querySelectorAll('[data-dl-btn]').forEach(function(b){
    b.textContent=local?'复制路径':'下载';
    b.title=local?(hubFilePath(url,name)||''):'';
  });
  document.querySelectorAll('[data-dl-copy-btn]').forEach(function(b){
    b.style.display=local?'inline-block':'none';
  });
}
function hideDownloadBanner(){
  var b=$('download-banner');if(!b)return;
  b.classList.remove('show');
  if(dlBannerTimer){clearTimeout(dlBannerTimer);dlBannerTimer=null}
  var f=$('db-fill');if(f){f.className='ub-fill';f.style.width='0%'}
}
var dlBannerTimer=null;
var dlBusy=false;
function startNativeDownload(dlUrl,name){
  showDownloadBanner(name,-1,'大文件 · 已交给浏览器下载，请查看通知栏/下载管理',null);
  var a=document.createElement('a');
  a.href=dlUrl;
  a.download=name;
  a.target='_blank';
  a.rel='noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function(){showDownloadBanner(name,100,'系统正在下载 · 请看通知栏','ok');dlBusy=false},1200);
}
function dlOne(url,name){
  url=fixImgUrl(url);
  var dlUrl=url+(url.indexOf('?')>=0?'&':'?')+'dl=1';
  try{name=name||decodeURIComponent((url.split('/').pop()||'download').split('?')[0])}catch(_){name=(url.split('/').pop()||'download').split('?')[0]}
  if(dlBusy){toast('已有下载进行中，请稍候');return}
  dlBusy=true;
  fetch(dlUrl,{method:'HEAD'}).then(function(r){
    if(!r.ok)throw new Error('HTTP '+r.status);
    var len=parseInt(r.headers.get('Content-Length')||'0',10)||0;
    var big=len>50*1024*1024;
    if(im||big){
      startNativeDownload(dlUrl,name);
      toast('大文件 '+fmt(len)+' · 请看浏览器下载管理');
      return;
    }
    return xhrDownloadBlob({url:dlUrl,name:name}).finally(function(){dlBusy=false});
  }).catch(function(e){
    console.error('dlOne',e);
    dlBusy=false;
    xhrDownloadBlob({url:dlUrl,name:name}).catch(function(e2){
      console.error('dlOne fallback',e2);
      var msg=e2&&e2.message==='network'?'网络中断，请重试':((e2&&e2.message)||'下载失败');
      showDownloadBanner(name,0,msg,'er');
      toast('下载失败，尝试浏览器打开');
      startNativeDownload(dlUrl,name);
    }).finally(function(){dlBusy=false});
  });
}
function showDownloadBanner(name,pct,status,phase){
  var b=$('download-banner');if(!b)return;
  b.classList.add('show');
  if($('db-name'))$('db-name').textContent=name||'下载中';
  var fill=$('db-fill'),st=$('db-status'),pc=$('db-pct');
  if(phase==='ok'){
    if(pc)pc.textContent='100%';
    if(fill){fill.className='ub-fill done';fill.style.width='100%'}
    if(st){st.textContent=status||'下载完成';st.className='ub-status ok'}
    if(dlBannerTimer)clearTimeout(dlBannerTimer);
    dlBannerTimer=setTimeout(hideDownloadBanner,2800);
  }else if(phase==='er'){
    if(pc)pc.textContent='!';
    if(fill){fill.className='ub-fill er';fill.style.width='100%'}
    if(st){st.textContent=status||'下载失败';st.className='ub-status er'}
    if(dlBannerTimer)clearTimeout(dlBannerTimer);
    dlBannerTimer=setTimeout(hideDownloadBanner,4500);
  }else if(pct==null||pct<0){
    if(pc)pc.textContent='…';
    if(fill){fill.className='ub-fill indeterminate'}
    if(st){st.textContent=status||'下载中';st.className='ub-status'}
  }else{
    if(pc)pc.textContent=pct+'%';
    if(fill){fill.className='ub-fill';fill.style.width=Math.min(100,pct)+'%'}
    if(st){st.textContent=status||('下载中 '+pct+'%');st.className='ub-status'}
  }
}
function xhrDownloadBlob(opts){
  opts=opts||{};
  var url=opts.url,name=opts.name||'download',method=opts.method||'GET',body=opts.body||null;
  return new Promise(function(resolve,reject){
    var x=new XMLHttpRequest();
    x.open(method,url);
    if(body)x.setRequestHeader('Content-Type','application/json');
    x.responseType='blob';
    var t0=Date.now(),lastLoaded=0,lastT=t0;
    x.onprogress=function(e){
      var now=Date.now(),spdTxt='';
      if(now-lastT>350){
        var dLoaded=e.loaded-lastLoaded,dT=(now-lastT)/1000;
        if(dT>0&&dLoaded>0){
          var spd=dLoaded/1024/1024/dT;
          if(spd>=0.05)spdTxt=' · '+spd.toFixed(1)+' MB/s';
        }
        lastLoaded=e.loaded;lastT=now;
      }
      if(e.lengthComputable&&e.total>0){
        var pct=Math.min(100,Math.round(e.loaded/e.total*100));
        showDownloadBanner(name,pct,'下载中 '+fmt(e.loaded)+' / '+fmt(e.total)+spdTxt,null);
      }else{
        showDownloadBanner(name,-1,'已接收 '+fmt(e.loaded)+spdTxt,null);
      }
    };
    x.onload=function(){
      if(x.status!==200){reject(new Error('HTTP '+x.status));return}
      showDownloadBanner(name,100,'正在保存到设备…',null);
      try{
        var a=document.createElement('a');
        a.href=URL.createObjectURL(x.response);
        a.download=name;
        document.body.appendChild(a);a.click();a.remove();
        showDownloadBanner(name,100,'已开始下载','ok');
        resolve();
      }catch(err){reject(err)}
    };
    x.onerror=function(){reject(new Error('network'))};
    x.ontimeout=function(){reject(new Error('下载超时'))};
    x.timeout=7200000;
    showDownloadBanner(name,0,'连接服务器…',null);
    x.send(body||null);
  });
}
var batchPaths=[];
async function batchDownload(paths,label){
  paths=(paths||[]).filter(Boolean);
  if(!paths.length){toast('没有可下载的文件');return}
  var zipName=(label||'Relay-batch')+'-'+new Date().toISOString().slice(0,16).replace(/[T:]/g,'-')+'.zip';
  try{
    await xhrDownloadBlob({
      url:'/api/batch-download',
      method:'POST',
      body:JSON.stringify({paths:paths}),
      name:zipName
    });
  }catch(e){
    console.error('batchDownload',e);
    showDownloadBanner('打包下载',0,'打包下载失败','er');
    toast('打包失败');
  }
}
async function postPackedBatch(items){
  var paths=items.map(function(it){return it.relpath||relFromUrl(it.url)}).filter(Boolean);
  if(!paths.length)throw new Error('no paths');
  var r=await fetch('/api/pack-save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({paths:paths})});
  var d=await r.json();
  if(!r.ok||!d.ok)throw new Error((d&&d.error)||'pack failed');
  var cap=(d.count||paths.length)+' 个文件 · 一次下载全部';
  await postMsg(autoSender(),autoSenderName(),'[[FILE]]'+fixImgUrl(d.url)+'\n'+cap);
  return d;
}
var packChoiceResolve=null;
function askPackSendChoice(count){
  if(!count||count<2)return Promise.resolve('separate');
  return new Promise(function(resolve){
    var ov=$('pack-choice-overlay'),n=$('pack-choice-count');
    if(!ov){resolve('separate');return}
    if(n)n.textContent=String(count);
    packChoiceResolve=resolve;
    ov.classList.add('show');
  });
}
function closePackSendChoice(choice){
  var ov=$('pack-choice-overlay');
  if(ov)ov.classList.remove('show');
  if(packChoiceResolve){packChoiceResolve(choice);packChoiceResolve=null}
}
function initPackChoiceUI(){
  var sep=$('pack-choice-separate'),zip=$('pack-choice-zip'),cancel=$('pack-choice-cancel');
  if(sep)sep.addEventListener('click',function(){sfxClick();closePackSendChoice('separate')});
  if(zip)zip.addEventListener('click',function(){sfxClick();closePackSendChoice('pack')});
  if(cancel)cancel.addEventListener('click',function(){sfxClick();closePackSendChoice(null)});
}
async function sendUploadedItems(uploaded,packed){
  if(packed&&uploaded.length>=2){
    await postPackedBatch(uploaded);
    toast('已打包发出 '+uploaded.length+' 个文件（1 个 zip）');
    return;
  }
  for(var i=0;i<uploaded.length;i++){await postMsgAuto(uploaded[i].url,uploaded[i].name)}
  toast(uploaded.length>1?('已发出 '+uploaded.length+' 个文件'):'已发送');
}
function relFromUrl(u){
  u=fixImgUrl(u||'');
  if(!u)return'';
  if(u.indexOf('http')===0){
    try{var a=document.createElement('a');a.href=u;u=a.pathname+(a.search||'')}catch(e){}
  }
  try{
    var dec=decodeURIComponent(u);
    var m=dec.match(/\/api\/download\/(.+?)(?:\?|$)/);
    if(m)return m[1];
  }catch(e){}
  if(u.indexOf('/api/download/')===0)return u.slice(14).split('?')[0];
  return'';
}
function addCopyBtn(body,url,name,text){
  if(arguments.length===2&&typeof url==='string'&&!name&&!text){
    if(url.indexOf('/')>=0||url.indexOf('http')===0)name=fileNameFrom(url,'');
    else{text=url;url=''}
  }
  var cpb=document.createElement('button');cpb.type='button';cpb.className='mcp';cpb.innerHTML=si('copy',12);
  cpb.title=copyBtnTitle(url,name,text);
  cpb.addEventListener('click',function(e){
    e.stopPropagation();e.preventDefault();
    sfxClick();
    cpSmart(url,name,text);
  });
  body.appendChild(cpb);
}
function setSelToggle(on){
  var tb=$('sel-toggle');if(!tb)return;
  tb.classList.toggle('on',!!on);
  tb.innerHTML=si(on?'check':'grid',18);
  tb.title=on?'完成':'多选';
  tb.setAttribute('aria-label',on?'完成多选':'多选');
}
function toggleSelMode(){
  if(selMode){exitSelMode();return}
  selMode=true;document.body.classList.add('sel-active');
  setSelToggle(true);updSelBar();sfxClick();
}
function resolveSelMeta(el){
  if(!el)return null;
  var key=el.getAttribute('data-sel-key');
  if(!key)return null;
  var type=el.getAttribute('data-sel-type')||'file';
  var rel=el.getAttribute('data-sel-rel')||'';
  var url=el.getAttribute('data-sel-url')||'';
  var name=el.getAttribute('data-sel-name')||'';
  var text=el.getAttribute('data-sel-text')||el.getAttribute('data-raw-text')||'';
  if(type==='file'){
    if(!rel)rel=relFromUrl(url)||'';
    if(!rel&&url){
      try{name=name||decodeURIComponent((url.split('/').pop()||'').split('?')[0])}catch(_){name=(url.split('/').pop()||'').split('?')[0]}
      if(name)rel=UPLOAD_REL+'/'+name;
    }
    if(!rel)return null;
    if(!name)name=rel.split('/').pop()||'file';
    return{key:key,type:'file',rel:rel,url:fixImgUrl(url),name:name,text:'',mid:el.getAttribute('data-mid')||''};
  }
  if(!text.trim())return null;
  return{key:key,type:'text',rel:'',url:'',name:'',text:text,mid:el.getAttribute('data-mid')||''};
}
function isInboxRel(rel){
  if(!rel)return false;
  rel=String(rel).replace(/\\/g,'/');
  var ul=(UPLOAD_REL||'from-phone').replace(/\\/g,'/');
  return rel.indexOf('from-phone/')>=0||rel.indexOf(ul)===0||rel.indexOf('下载/from-phone')>=0;
}
function updSelBar(){
  var n=Object.keys(selMap).length;
  var c=$('sel-count'),d=$('sel-dl'),del=$('sel-del'),bar=$('sel-bar');
  if(c)c.textContent=n?('×'+n):'';
  if(d)d.disabled=n===0;
  if(del)del.disabled=n===0;
  if(bar)bar.style.display=selMode?'flex':'none';
  if(!selMode)setSelToggle(false);
  renderSelPreview();
}
function exitSelMode(){
  selMode=false;
  document.body.classList.remove('sel-active');
  Object.keys(selMap).forEach(function(k){if(selMap[k].el)selMap[k].el.classList.remove('selected','sel-row')});
  selMap={};
  var sp=$('sel-preview');if(sp)sp.style.display='none';
  updSelBar();
}
function toggleSel(key,el){
  if(!key||!el)return;
  var meta=resolveSelMeta(el);
  if(!meta){toast('无法选中此项');return}
  if(selMap[key]){delete selMap[key];el.classList.remove('selected','sel-row')}
  else{selMap[key]=Object.assign({el:el},meta);el.classList.add(el.classList.contains('yg-item')||el.classList.contains('fr')?'selected':'sel-row')}
  updSelBar();sfxClick();
}
function enterSelMode(key,el){
  if(!selMode)toggleSelMode();
  toggleSel(key,el);
}
var selTapAt=0;
function handleSelTarget(e){
  if(!selMode)return false;
  if(window.getSelection&&window.getSelection().toString())return false;
  if(e.target.closest('.sel-bar,.sel-preview,.sel-toggle-btn,.mcp,.expand-btn,.expand-top,.cib,nav.tabs,header.deck-h,.sel-del'))return false;
  if(e.target.closest('.maud,.mvid'))return false;
  var now=Date.now();
  if(now-selTapAt<280)return false;
  var el=e.target.closest('#ca .mr[data-sel-key],#fl .yg-item[data-sel-key],#fl .fr[data-sel-key]:not(.dir)');
  if(!el)return false;
  selTapAt=now;
  e.preventDefault();e.stopPropagation();
  toggleSel(el.getAttribute('data-sel-key'),el);
  return true;
}
function renderSelPreview(){
  var bar=$('sel-preview'),list=$('sel-pv-list');
  if(!bar||!list)return;
  var keys=Object.keys(selMap);
  if(!selMode||!keys.length){bar.style.display='none';return}
  bar.style.display='block';
  list.innerHTML='';
  keys.forEach(function(k){
    var it=selMap[k],chip=document.createElement('button');
    chip.type='button';chip.className='sel-chip';
    chip.title=it.type==='text'?it.text:(it.name||'');
    if(it.type==='text'){
      chip.textContent=(it.text.replace(/\s+/g,' ').trim().slice(0,28)||'文字')+(it.text.length>28?'…':'');
    }else if(it.url&&isImg(it.name||it.url)){
      var im=document.createElement('img');im.src=fixImgUrl(it.url);im.alt='';chip.appendChild(im);
    }else{chip.textContent=(it.name||'文件').slice(0,12)}
    chip.addEventListener('click',function(ev){ev.stopPropagation();previewSelItem(it)});
    list.appendChild(chip);
  });
}
function previewSelItem(it){
  if(!it)return;
  if(it.type==='text'){
    var mo=$('mo');if(!mo)return;
    mo.style.display='flex';if(!im)mo.classList.add('show-desktop');
    $('mfn').textContent='文字消息';
    var inner=$('mbd-inner');if(inner){
      inner.innerHTML='<div class="md-body" style="padding:12px;max-width:100%;text-align:left;user-select:text">'+mdToHtml(it.text)+'</div>';
      addCopyBtn(inner,'','',it.text);
    }
    $('mdl').style.display='flex';pvGallery=[];updGalleryNav('txt');
    return;
  }
  previewFile(it.url,it.name||'file');
}
function dlTextBlob(text,filename){
  var blob=new Blob([text],{type:'text/markdown;charset=utf-8'});
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename;
  document.body.appendChild(a);a.click();a.remove();
}
async function downloadSelected(){
  var items=Object.keys(selMap).map(function(k){return selMap[k]});
  if(!items.length){toast('请先选择');return}
  var files=items.filter(function(i){return i.type==='file'&&i.rel});
  var texts=items.filter(function(i){return i.type==='text'&&i.text});
  try{
    if(files.length)await batchDownload(files.map(function(f){return f.rel}),'files');
    if(texts.length){
      var md=texts.map(function(t,i){
        var row=t.el,nm=row&&row.querySelector('.mm-nm'),tm=row&&row.querySelector('.mm-time');
        return '## '+(nm?nm.textContent:'消息')+' · '+(tm?tm.textContent:('#'+(i+1)))+'\n\n'+t.text;
      }).join('\n\n---\n\n');
      dlTextBlob(md,'lan-hub-'+new Date().toISOString().slice(0,10)+'.md');
    }
    toast('已开始下载');exitSelMode();
  }catch(e){toast('下载失败')}
}
async function deleteSelected(){
  var items=Object.keys(selMap).map(function(k){return selMap[k]});
  if(!items.length){toast('请先选择');return}
  var files=items.filter(function(i){return i.type==='file'&&i.rel});
  var msgIds=[];
  items.forEach(function(i){
    if(i.mid){var m=parseInt(i.mid,10);if(m&&msgIds.indexOf(m)<0)msgIds.push(m)}
  });
  var outside=files.filter(function(f){return !isInboxRel(f.rel)});
  var n=files.length+msgIds.length;
  var msg='确定删除选中的 '+n+' 项？';
  if(outside.length){
    msg+='\n\n其中 '+outside.length+' 个文件不在「手机传来」，删除后无法恢复。';
  }else if(HUB_MODE==='cloud'){
    msg+='\n\n将从云收件箱永久删除，释放服务器空间。';
  }else{
    msg+='\n\n文件删除后无法恢复。';
  }
  if(!confirm(msg))return;
  toast('删除中…');
  try{
    var r=await fetch('/api/batch-delete',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({paths:files.map(function(f){return f.rel}),message_ids:msgIds})
    });
    var d=await r.json();
    if(!r.ok||!d.ok)throw new Error((d&&d.error)||'HTTP '+r.status);
    var parts=[];
    if(d.deleted_files)parts.push(d.deleted_files+' 个文件');
    if(d.deleted_messages)parts.push(d.deleted_messages+' 条消息');
    if(parts.length)toast('已删除 '+parts.join('、'));
    else toast('所选内容在服务器上已不存在，正在刷新');
    if(d.failed&&d.failed.length)toast('部分失败：'+d.failed.length+' 项');
    exitSelMode();
    if(ct==='chat'){await fm(null);scb(true)}
    else bf(cp,cr);
  }catch(e){
    console.error('deleteSelected',e);
    toast('删除失败');
  }
}
function bindLongPress(el,key,rel){
  if(!el||el._lpBound)return;el._lpBound=true;
  key=key||el.getAttribute('data-sel-key')||'';
  var timer=null,sx=0,sy=0;
  function cancel(){if(timer){clearTimeout(timer);timer=null}}
  el.addEventListener('touchstart',function(e){
    if(selMode||e.touches.length>1)return cancel();
    sx=e.touches[0].clientX;sy=e.touches[0].clientY;
    timer=setTimeout(function(){timer=null;sfxClick();enterSelMode(key,el)},450);
  },{passive:true});
  el.addEventListener('touchmove',function(e){
    if(!timer)return;
    if(Math.abs(e.touches[0].clientX-sx)>12||Math.abs(e.touches[0].clientY-sy)>12)cancel();
  },{passive:true});
  el.addEventListener('touchend',cancel);
  el.addEventListener('touchcancel',cancel);
  el.addEventListener('contextmenu',function(e){
    e.preventDefault();e.stopPropagation();
    enterSelMode(key,el);
  });
}
function bindSelectable(root){
  if(!root)return;
  root.querySelectorAll('[data-sel-key]').forEach(function(el){
    bindLongPress(el,el.getAttribute('data-sel-key'));
  });
}
function updBatchBar(){}
function renderYearGallery(es,l){
  var items=es.filter(function(e){return !e.is_dir&&(isImg(e.name)||isVid(e.name))});
  items.sort(function(a,b){return (b.mtime||0)-(a.mtime||0)});
  if(currentCat==='image')items=items.filter(function(e){return isImg(e.name)});
  else if(currentCat==='video')items=items.filter(function(e){return isVid(e.name)});
  else if(currentCat!=='all')items=items.filter(function(e){return catOf(e.name)===currentCat});
  if(!items.length){l.className='fl';l.innerHTML='<div class="es">'+si('photo',20)+'暂无文件</div>';return}
  var byYear={};
  items.forEach(function(e){
    var y=new Date((e.mtime||0)*1000).getFullYear();
    if(!byYear[y])byYear[y]=[];
    var rel=e.relpath||((cp?cp+'/':'')+e.name);
    var du=dlPath(rel);
    byYear[y].push({e:e,du:du,rel:rel});
  });
  var years=Object.keys(byYear).sort(function(a,b){return b-a});
  l.className='fl year-gallery';
  l.innerHTML=years.map(function(y){
    var cells=byYear[y].map(function(it){
      var e=it.e,du=it.du,n=e.name;
      if(isImg(n))return '<div class="yg-item" data-url="'+du+'" data-name="'+eh(n)+'" data-sel-key="'+eh(it.rel)+'" data-sel-rel="'+eh(it.rel)+'"><img src="'+du+'" loading="lazy" alt=""></div>';
      return '<div class="yg-item" data-url="'+du+'" data-name="'+eh(n)+'" data-sel-key="'+eh(it.rel)+'" data-sel-rel="'+eh(it.rel)+'"><div class="yg-vid">'+si('video',28)+'</div></div>';
    }).join('');
    return '<div class="year-section"><div class="year-label">'+y+'</div><div class="year-strip">'+cells+'</div></div>';
  }).join('');
  rebuildGalleryEntries(items.map(function(e){return Object.assign({},e,{is_dir:false})}),true);
  bindSelectable(l);
}
function dlPreview(){
  if(!pf)return;
  if(HUB_MODE==='lan'&&isLocalHub()&&hubFilePath(pf.url,pf.name)){
    cpTx(hubFilePath(pf.url,pf.name));
    toast('已复制路径\n文件已在 Hub 目录，无需重复下载');
    return;
  }
  dlOne(pf.url,pf.name);
}
function dlPreviewCopy(){if(!pf)return;dlOne(pf.url,pf.name)}
function sendPreviewToChat(){if(pf){addToQueue(pf.url,pf.name);sw('chat');toast('已加入待发')}}

// ═══ Safe Markdown → HTML ═══
function mdToHtml(t){
  if(!t)return'';
  var lines=t.split('\n'),out=[],inPre=false,preBuf=[],list=[];
  function flushList(){
    if(!list.length)return;
    out.push('<ul>'+list.map(function(li){return'<li>'+li+'</li>'}).join('')+'</ul>');
    list=[];
  }
  for(var i=0;i<lines.length;i++){
    var line=lines[i];
    if(line.indexOf('```')===0){
      flushList();
      if(!inPre){inPre=true;preBuf=[]}
      else{out.push('<pre><code>'+eh(preBuf.join('\n'))+'</code></pre>');inPre=false;preBuf=[]}
      continue;
    }
    if(inPre){preBuf.push(line);continue}
    var hm=line.match(/^(#{1,3})\s+(.+)$/);
    if(hm){flushList();out.push('<h'+hm[1].length+'>'+inlineMd(hm[2])+'</h'+hm[1].length+'>');continue}
    var bq=line.match(/^>\s?(.*)$/);
    if(bq){flushList();out.push('<blockquote>'+inlineMd(bq[1])+'</blockquote>');continue}
    var lm=line.match(/^[\-\*]\s+(.+)$/);
    if(lm){list.push(inlineMd(lm[1]));continue}
    flushList();
    if(line.trim()===''){out.push('');continue}
    out.push('<p>'+inlineMd(line)+'</p>');
  }
  flushList();
  if(inPre&&preBuf.length)out.push('<pre><code>'+eh(preBuf.join('\n'))+'</code></pre>');
  return out.join('');
}
function inlineMd(s){
  s=eh(s||'');
  s=linkifyUrls(s);
  s=s.replace(/`([^`]+)`/g,'<code>$1</code>');
  s=s.replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>');
  s=s.replace(/\*([^*]+)\*/g,'<em>$1</em>');
  s=s.replace(/\[([^\]]+)\]\(([^)]+)\)/g,'<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>');
  return s;
}
function linkifyUrls(s){
  return s.replace(/(https?:\/\/[^\s<>"']+)/gi,function(m){
    var url=m.replace(/[.,;:!?)>]+$/,'');
    var trail=m.slice(url.length);
    return '<a href="'+url+'" target="_blank" rel="noopener noreferrer">'+url+'</a>'+trail;
  });
}

// ═══ Files ═══
async function bf(path,isRx){
  cp=path||'';cr=!!isRx;var l=$('fl');l.innerHTML='<div class="es"><div class="spinner"></div></div>';
  var ep=isRx?'/api/received':'/api/browse?path='+encodeURIComponent(cp);
  try{
    var r=await fetch(ep),d=await r.json();
    var bc=$('br');
    if(isRx)bc.innerHTML='<span style="color:var(--gn)">'+si('inbox',11)+' 手机传来 · from-phone</span>';
    else{var h='<a href="#" onclick="browseHome();return false">'+si('home',11,'folder')+' 本机 ~</a>';
      if(path){var ps=path.split('/').filter(Boolean),ac='';ps.forEach(function(p,i){ac+='/'+p;h+=' <span class="s">/</span> ';if(i<ps.length-1)h+='<a href="#" onclick="bf(\''+eh(ac)+'\');return false">'+eh(p)+'</a>';else h+='<span style="color:var(--am)">'+eh(p)+'</span>'})}
      bc.innerHTML=h}
    var es=d.entries||[];
    rebuildGalleryEntries(es,isRx);
    // Sort — dirs+files together by chosen key (DATE shows true recency)
    es.sort(function(a,b){
      var va=a[sk],vb=b[sk];
      if(sk==='name'){va=(a.name||'').toLowerCase();vb=(b.name||'').toLowerCase()}
      else{va=Number(va)||0;vb=Number(vb)||0}
      if(va<vb)return sa?-1:1;
      if(va>vb)return sa?1:-1;
      return(a.name||'').localeCompare(b.name||'');
    });
    // Category counts (files only)
    var counts={all:0,image:0,video:0,audio:0,doc:0};
    es.forEach(function(e){if(!e.is_dir){counts.all++;var c=catOf(e.name);if(counts[c]!==undefined)counts[c]++}});
    document.querySelectorAll('.cat-btn[data-cat]').forEach(function(b){
      var c=b.dataset.cat;b.querySelector('.cat-count')&&b.querySelector('.cat-count').remove();
      if(counts[c]!==undefined){var s=document.createElement('span');s.className='cat-count';s.textContent=counts[c];b.appendChild(s)}
    });
    // Filter
    if(currentCat!=='all')es=es.filter(function(e){return e.is_dir||catOf(e.name)===currentCat});
    if(im&&isRx&&(currentCat==='all'||currentCat==='image'||currentCat==='video')){
      renderYearGallery(d.entries||[],l);
      return;
    }
    if(!es.length){l.className='fl';l.innerHTML='<div class="es">'+si('folder',20)+'暂无文件</div>';updBatchBar([]);return}
    var useGrid=im&&currentCat==='image';
    l.className=useGrid?'fl img-grid':'fl';
    l.innerHTML=es.map(function(e){
      var cls=e.is_dir?' dir':'',icn=e.is_dir?si('folder',14,'folder'):fi(e.name);
      var up=e.is_dir?((cp?cp+'/':'')+e.name):'';
      var rel=e.relpath||(e.is_dir?up:((cp?cp+'/':'')+e.name));
      var du=dlPath(rel);
      var sz=e.is_dir?'':fmt(e.size),dd=new Date(e.mtime*1000);
      var ds=(dd.getMonth()+1)+'-'+dd.getDate()+' '+String(dd.getHours()).padStart(2,'0')+':'+String(dd.getMinutes()).padStart(2,'0');
      var sp=!e.is_dir&&isImg(e.name)&&!isHeic(e.name);
      var attrs=e.is_dir?' data-path="'+eh(up)+'"':' data-url="'+du+'" data-name="'+eh(e.name)+'" data-sel-key="'+eh(rel)+'" data-sel-rel="'+eh(rel)+'"';
      return'<div class="fr'+cls+'"'+attrs+' style="cursor:pointer">'+
        (sp?listImgHtml(rel,du):'')+
        '<span class="fi">'+icn+'</span><span class="fn">'+eh(e.name)+'</span>'+
        '<span class="fm fs">'+sz+'</span><span class="fm fd">'+ds+'</span></div>';
    }).join('');
    bindSelectable(l);
  }catch(err){l.innerHTML='<div class="es">'+si('x',20)+'LOAD FAILED</div>'}
}
$('fl').addEventListener('click',function(e){
  if(selMode)return;
  var yg=e.target.closest('.yg-item');
  if(yg){
    var u=yg.getAttribute('data-url'),n=yg.getAttribute('data-name');
    if(u){sfxOpen();previewFile(u,n);}
    return;
  }
  var row=e.target.closest('.fr');if(!row)return;
  if(row.classList.contains('dir')){
    sfxOpen();bf(row.getAttribute('data-path')||'',cr);
  }else{
    var u=row.getAttribute('data-url'),n=row.getAttribute('data-name');
    if(u){
      document.querySelectorAll('#fl .fr.active').forEach(function(r){r.classList.remove('active')});
      row.classList.add('active');
      sfxOpen();previewFile(u,n);
    }
  }
});
var fileSrc='home';
function updFileSrcUI(){
  var bar=$('file-src-bar');
  if(bar)bar.style.display=HUB_MODE==='lan'?'flex':'none';
  var h=$('src-home'),i=$('src-inbox');
  if(h)h.classList.toggle('active',fileSrc==='home');
  if(i)i.classList.toggle('active',fileSrc==='inbox');
  var ph=$('files-ph-label');
  if(ph){
    if(HUB_MODE==='cloud')ph.textContent='云收件箱';
    else ph.textContent=fileSrc==='home'?'本机文件':'手机传来';
  }
}
function browseHome(){fileSrc='home';updFileSrcUI();if(selMode)exitSelMode();sfxClick();bf('',false)}
function browseReceived(){fileSrc='inbox';updFileSrcUI();if(selMode)exitSelMode();sfxClick();bf('',true)}
function loadFilesDefault(){
  if(HUB_MODE==='cloud')browseReceived();
  else browseHome();
}
// Category filter
document.getElementById('cat-bar').addEventListener('click',function(e){
  var btn=e.target.closest('.cat-btn[data-cat]');if(!btn)return;
  sfxClick();
  currentCat=btn.dataset.cat;sfxClick();
  document.querySelectorAll('.cat-btn[data-cat]').forEach(function(b){b.classList.toggle('active',b.dataset.cat===currentCat)});
  bf(cp,cr);
});

// Sort buttons
function updSortUI(){
  document.querySelectorAll('.sort-btn[data-sort]').forEach(function(b){
    var on=b.dataset.sort===sk;
    b.classList.toggle('active',on);
    var old=b.querySelector('.sort-arr');if(old)old.remove();
    if(on){var s=document.createElement('span');s.className='sort-arr';s.textContent=sa?'↑':'↓';b.appendChild(s)}
  });
}
document.getElementById('sort-bar').addEventListener('click',function(e){
  var btn=e.target.closest('.sort-btn[data-sort]');if(!btn)return;
  sfxClick();
  var key=btn.dataset.sort;
  if(sk===key)sa=!sa;else{sk=key;sa=(key==='name')}
  updSortUI();sfxClick();bf(cp,cr);
});
// ═══ Upload ═══
var uf=[],fli=$('fi'),pii=$('pi'),ufl=$('ufl'),ubt=$('ubt'),cb=$('cb'),us=$('us'),ue=$('ue');
function rul(){
  ufl.innerHTML='';uf.forEach(function(f,i){var c=document.createElement('div');c.className='uc';c.dataset.idx=i;c.innerHTML='<div class="ui"><span class="un">'+eh(f.name)+'</span><span class="um">'+fmt(f.size)+'</span></div><div class="us wt">WAITING</div><div class="pt"><div class="pf"></div></div>';ufl.appendChild(c)});
  var e=uf.length===0;ue.style.display=e?'flex':'none';ufl.style.display=e?'none':'';ubt.disabled=e;cb.disabled=e;us.classList.remove('show');
}
$('pbt').addEventListener('click',function(){sfxClick();pii.click()});
pii.addEventListener('change',function(){var p=Array.from(pii.files);if(p.length){uf=uf.concat(p);rul();pii.value=''}});
$('fbt').addEventListener('click',function(){sfxClick();fli.click()});
fli.addEventListener('change',function(){var p=Array.from(fli.files);if(p.length){uf=uf.concat(p);rul();fli.value=''}});
cb.addEventListener('click',function(){uf=[];rul()});
function sus(idx,s,t){var c=ufl.querySelector('[data-idx="'+idx+'"]');if(!c)return;c.className='uc '+s;c.querySelector('.us').className='us '+s;c.querySelector('.us').textContent=t}
function sup(idx,pct){var c=ufl.querySelector('[data-idx="'+idx+'"]');if(!c)return;c.querySelector('.pf').style.width=pct+'%'}
ubt.addEventListener('click',async function(){
  if(!uf.length)return;ubt.disabled=true;cb.disabled=true;us.classList.remove('show');us.className='usm';var ok=0,er=0;
  for(var i=0;i<uf.length;i++){sus(i,'up','UPLOADING');try{await dou(uf[i],i);sus(i,'ok','OK');sup(i,100);ok++}catch(e){sus(i,'er','FAIL');sup(i,100);er++}}
  us.classList.add('show');var t=uf.length;
  if(er===0){us.className='usm show ok';us.innerHTML=t+' FILE(S) UPLOADED'}
  else if(ok===0){us.className='usm show er';us.innerHTML='ALL FAILED'}
  else{us.className='usm show er';us.innerHTML=ok+'/'+t+' OK, '+er+' FAILED'}
  uf=[];rul();sfxSend();
});
function dou(file,idx){return new Promise(function(res,rej){var x=new XMLHttpRequest(),fd=new FormData();fd.append('file',file);x.upload.addEventListener('progress',function(e){if(e.lengthComputable)sup(idx,Math.round(e.loaded/e.total*100))});x.addEventListener('load',function(){if(x.status===200){try{JSON.parse(x.responseText);res()}catch(_){rej(new Error('bad response'))}}else rej(new Error('HTTP '+x.status))});x.addEventListener('error',function(){rej()});x.open('POST','/api/upload');x.timeout=600000;x.send(fd)})}

// ═══ Chat ═══
function chatAtBottom(){var a=$('ca');if(!a)return true;return a.scrollHeight-a.scrollTop-a.clientHeight<80}
function updJumpBtn(){var b=$('jump-bottom');if(b)b.classList.toggle('show',!chatAtBottom())}
function scb(force){
  if(!force&&!chatAtBottom()){updJumpBtn();return}
  var a=$('ca');if(!a)return;
  requestAnimationFrame(function(){a.scrollTop=a.scrollHeight;updJumpBtn()});
}
function ic(){return fm(null).then(function(){scb(true);sp()})}
var icDone=false;
async function initChat(){
  if(icDone)return;icDone=true;
  try{await fm(null);scb(true);sp();loadDraft();applyChatFilter();}catch(e){}
}
var chatFilter='all';
function applyChatFilter(){
  var hide=chatFilter==='text';
  document.querySelectorAll('#ca .mr').forEach(function(r){
    var isFile=r.getAttribute('data-sel-type')==='file'||!!r.getAttribute('data-attach-url');
    r.classList.toggle('hide-filter',hide&&isFile);
  });
}
var DRAFT_KEY='ls-chat-draft-'+HUB_MODE;
function saveDraft(){try{localStorage.setItem(DRAFT_KEY,$('ci2').value)}catch(e){}}
function loadDraft(){try{var v=localStorage.getItem(DRAFT_KEY);if(v&&v.trim()){$('ci2').value=v;resizeInput()}}catch(e){}}
function clearDraft(){try{localStorage.removeItem(DRAFT_KEY)}catch(e){}}
function notifyMsg(m){
  if(!('Notification' in window)||!notifyOk||!m)return;
  if(im&&m.sender==='phone')return;       // 手机端不发自己消息的通知
  if(!im&&m.sender==='pc')return;          // 电脑端不发自己消息的通知
  var body=m.content||'';
  if(isAttachMsg(body)){
    var u=imgFromContent(body)||'';
    body='[附件] '+(decodeURIComponent((u.split('/').pop()||'file').split('?')[0]));
  }else if(body.length>100)body=body.slice(0,100)+'…';
  try{
    var who=m.sender==='pc'?'电脑':'手机';
    var n=new Notification(who+' · 新消息',{body:body,tag:'ls-'+m.id});
    n.onclick=function(){window.focus();sw('chat');setTimeout(function(){scb(true)},50);n.close()};
  }catch(e){}
}
function requestNotify(){
  if(!('Notification' in window)){toast('浏览器不支持通知');return}
  Notification.requestPermission().then(function(p){
    notifyOk=(p==='granted');
    var b=$('notify-btn');if(b){b.classList.toggle('on',notifyOk);b.innerHTML=notifyOk?si('bell',16):si('bell',16)}
    toast(notifyOk?'通知已开启':'通知被拒绝');
  });
}
function initNotify(){
  if(im||!('Notification' in window))return;
  var nb=$('notify-btn');if(nb)nb.style.display='flex';
  if(Notification.permission==='granted'){notifyOk=true;if(nb){nb.classList.add('on');nb.innerHTML=si('bell',16)}}
  else if(Notification.permission==='default')setTimeout(function(){toast('点击右上角铃铛开启通知')},1500);
}
function onIncoming(m){
  // 微信式双向：自己发的也立即渲染，对方发的给通知
  var mine=(m.sender==='pc'&&!im)||(m.sender==='phone'&&im);
  if(!mine){
    notifyMsg(m);
    if(ct!=='chat'){sfxSend();toast('新消息')}
    if(document.hidden&&document.title.indexOf('●')!==0)document.title='● '+document.title;
  }
}
var lid=0,fetching=false,msgTotal=0,msgHasMore=false,msgOldest=0,loadingOlder=false;
var msgEpoch=0,chatNeedsRefresh=false;
function applyMsgEpoch(epoch){
  if(typeof epoch!=='number')return Promise.resolve();
  if(msgEpoch&&epoch!==msgEpoch){
    msgEpoch=epoch;
    if(ct==='chat'){chatNeedsRefresh=false;return fm(null)}
    chatNeedsRefresh=true;
    return Promise.resolve();
  }
  msgEpoch=epoch;
  return Promise.resolve();
}
function lastSeenId(){var a=$('ca');var rows=a.querySelectorAll('.mr');if(!rows.length)return 0;var last=rows[rows.length-1];var m=last.getAttribute('data-mid');return m?parseInt(m,10)||0:0}
function firstSeenId(){var a=$('ca');var row=a.querySelector('.mr');if(!row)return 0;var m=row.getAttribute('data-mid');return m?parseInt(m,10)||0:0}
function updLoadOlder(){
  var b=$('load-older');if(!b)return;
  var hidden=msgTotal>0?Math.max(0,msgTotal-document.querySelectorAll('#ca .mr').length):0;
  var show=msgHasMore||hidden>0;
  b.classList.toggle('show',show);
  if(show){
    b.disabled=loadingOlder;
    b.textContent=loadingOlder?'加载中…':('↑ 加载更早消息'+(hidden>0?'（还有 '+hidden+' 条）':''));
  }
}
async function fm(si){
  if(fetching)return;
  fetching=true;
  var u='/api/messages?meta=1';if(si!==null)u+='&since='+si;
  try{
    var r=await fetch(u),d=await r.json();
    var ms=(d&&d.messages)?d.messages:(Array.isArray(d)?d:[]);
    if(d&&typeof d.epoch==='number'){
      if(msgEpoch&&d.epoch!==msgEpoch&&si!==null){
        msgEpoch=d.epoch;
        fetching=false;
        chatNeedsRefresh=false;
        return fm(null);
      }
      msgEpoch=d.epoch;
    }
    if(d&&typeof d.total==='number'){msgTotal=d.total;msgHasMore=!!d.has_more;msgOldest=d.oldest_id||0}
    if(ms&&ms.length){
      if(si===null)ram(ms);
      else{
        appendNew(ms);
        var domCount=document.querySelectorAll('#ca .mr').length;
        if(domCount===0){ram(ms)}
      }
      if(ms.length)lid=ms[ms.length-1].id;
      if(!msgOldest)msgOldest=firstSeenId();
    }else if(si===null){
      lid=0;msgTotal=0;msgHasMore=false;msgOldest=0;
    }
    updLoadOlder();
  }catch(e){console.error('fm',e)}
  fetching=false;
}
function prependMsgs(ms){
  var a=$('ca'),btn=$('load-older'),es=a.querySelector('.es');
  if(es)es.remove();
  var seen={};
  a.querySelectorAll('.mr').forEach(function(r){var mid=r.getAttribute('data-mid');if(mid)seen[mid]=1});
  var anchor=a.querySelector('.mr')||(btn?btn.nextSibling:null);
  var h=a.scrollHeight,st=a.scrollTop,added=0,frag=document.createDocumentFragment();
  ms.forEach(function(m){
    if(seen[m.id])return;
    frag.appendChild(cem(m));
    seen[m.id]=1;added++;
  });
  if(added){
    if(anchor)a.insertBefore(frag,anchor);
    else a.appendChild(frag);
    bindSelectable(a);
    a.scrollTop=st+(a.scrollHeight-h);
    updLoadOlder();
  }
}
async function loadOlderMsgs(){
  if(loadingOlder||fetching)return;
  var before=firstSeenId();if(!before||before<=1)return;
  loadingOlder=true;updLoadOlder();
  try{
    var r=await fetch('/api/messages?meta=1&before='+before);
    var d=await r.json();
    var ms=(d&&d.messages)?d.messages:[];
    if(d&&typeof d.total==='number'){msgTotal=d.total;msgHasMore=!!d.has_more;msgOldest=d.oldest_id||before}
    if(ms.length)prependMsgs(ms);
    else msgHasMore=false;
    updLoadOlder();
  }catch(e){toast('加载失败');console.error('loadOlder',e)}
  loadingOlder=false;updLoadOlder();
}
function ram(ms){
  var a=$('ca');
  a.querySelectorAll('.mr,.es').forEach(function(n){n.remove()});
  if(!ms.length){var es=document.createElement('div');es.className='es';es.innerHTML=si('chat',20)+'暂无消息';a.appendChild(es);lid=0;msgOldest=0;updLoadOlder();return}
  ms.forEach(function(m){a.appendChild(cem(m))});
  bindSelectable(a);
  lid=ms[ms.length-1].id;
  msgOldest=ms[0].id;
  scb(true);
  updLoadOlder();
  applyChatFilter();
}
function makeTextBody(text){
  var wrap=document.createElement('div');wrap.className='mb';
  var long=text.length>220||text.split('\n').length>6;
  if(long){
    var inner=document.createElement('div');inner.className='mb-txt mb-collapsed';inner.innerHTML=mdToHtml(text);
    wrap.appendChild(inner);
    var topBtn=null;
    function setExpanded(exp){
      if(exp){
        inner.classList.remove('mb-collapsed');
        wrap.classList.add('mb-expanded');
        btn.textContent='收起';
        if(!topBtn){
          topBtn=document.createElement('button');
          topBtn.type='button';topBtn.className='expand-btn expand-top';
          topBtn.textContent='↑ 收起';
          topBtn.addEventListener('click',toggleExpand);
          topBtn.addEventListener('touchend',function(e){e.stopPropagation()},{passive:true});
          wrap.insertBefore(topBtn,inner);
        }
      }else{
        inner.classList.add('mb-collapsed');
        wrap.classList.remove('mb-expanded');
        btn.textContent='展开全文';
        if(topBtn){topBtn.remove();topBtn=null}
      }
      updJumpBtn();
    }
    function toggleExpand(e){
      if(e){e.stopPropagation();e.preventDefault()}
      sfxClick();
      setExpanded(inner.classList.contains('mb-collapsed'));
    }
    var btn=document.createElement('button');btn.type='button';btn.className='expand-btn';btn.textContent='展开全文';
    btn.addEventListener('click',toggleExpand);
    btn.addEventListener('touchend',function(e){e.stopPropagation()},{passive:true});
    wrap.appendChild(btn);
  }else{var inner2=document.createElement('div');inner2.className='md-body';inner2.innerHTML=mdToHtml(text);wrap.appendChild(inner2)}
  addCopyBtn(wrap,'','',text);
  return wrap;
}
function appendNew(ms){
  var a=$('ca'),e=a.querySelector('.es');if(e)e.remove();
  var seen={};
  a.querySelectorAll('.mr').forEach(function(r){var mid=r.getAttribute('data-mid');if(mid)seen[mid]=1});
  var wasBottom=chatAtBottom(),added=false;
  ms.forEach(function(m){
    if(seen[m.id])return;
    try{
      onIncoming(m);
      a.appendChild(cem(m));
      seen[m.id]=1;added=true;
    }catch(ex){console.error('render msg',m.id,ex)}
  });
  if(added){scb(wasBottom);bindSelectable(a);applyChatFilter();if(a.scrollHeight<50)a.style.minHeight='60vh'}
}
function addToQueue(url,name){url=fixImgUrl(url);pq.push({url:url,name:name,ok:true});renderAttachQueue()}
function hideUploadBanner(){var b=$('upload-banner');if(b){b.classList.remove('show');var f=$('ub-fill');if(f){f.className='ub-fill';f.style.width='0%'}}}
function showUploadBanner(slot,index,total,phase){
  var b=$('upload-banner');if(!b||!slot)return;
  b.classList.add('show');
  var pct=slot.progress||0;
  $('ub-name').textContent=(total>1?'('+(index+1)+'/'+total+') ':'')+slot.name;
  $('ub-pct').textContent=phase==='ok'||phase==='er'?'100%':pct+'%';
  var fill=$('ub-fill');fill.style.width=(phase==='ok'||phase==='er'||pct>=100?'100':pct)+'%';fill.className='ub-fill';
  var st=$('ub-status');
  if(phase==='ok'){st.textContent='已发出';st.className='ub-status ok';fill.classList.add('done')}
  else if(phase==='er'){st.textContent=slot.error||'上传失败';st.className='ub-status er';fill.classList.add('er')}
  else if(phase==='post'){st.textContent='发送到聊天…';st.className='ub-status'}
  else if(pct>=100){st.textContent='等待服务器…';st.className='ub-status'}
  else{st.textContent='上传中 '+pct+'%';st.className='ub-status'}
}
function attachQueueLabel(){
  if(pq.some(function(it){return it.uploading&&(it.progress||0)<100}))return ' · 上传中';
  if(pq.some(function(it){return it.uploading&&(it.progress||0)>=100}))return ' · 等待服务器';
  if(pq.some(function(it){return it.posting}))return ' · 发送中';
  return '';
}
function renderAttachQueue(){
  var bar=$('aqb'),list=$('aq-list');
  if(!pq.length){bar.style.display='none';list.innerHTML='';return}
  bar.style.display='block';
  var ready=pq.filter(function(it){return it.ok&&it.url&&!it.uploading&&!it.posting}).length;
  $('aq-count').textContent=pq.length+' 个文件'+attachQueueLabel();
  list.innerHTML='';
  pq.forEach(function(it,i){
    var chip=document.createElement('div');chip.className='aq-chip'+(it.ok===false?' er':it.uploading||it.posting?' up':' ok');
    var thumb=document.createElement('div');thumb.className='aq-thumb';
    if(it.uploading||it.posting){thumb.innerHTML=si(it.posting?'send':'upload',22)}
    else if(isImg(it.name)&&it.url){var im=document.createElement('img');im.src=it.url;im.onclick=function(){previewFile(it.url,it.name)};thumb.appendChild(im)}
    else if(isVid(it.name)&&it.url){var vd=document.createElement('video');vd.src=it.url;vd.muted=true;vd.preload='metadata';vd.onclick=function(){previewFile(it.url,it.name)};thumb.appendChild(vd)}
    else{thumb.innerHTML=si('file',22);if(it.url){thumb.style.cursor='pointer';thumb.onclick=function(){previewFile(it.url,it.name)}}}
    var nm=document.createElement('div');nm.className='aq-nm';nm.textContent=it.name;nm.title=it.name;
    var rm=document.createElement('button');rm.className='aq-rm';rm.textContent='×';
    rm.onclick=function(e){e.stopPropagation();pq.splice(i,1);renderAttachQueue()};
    chip.appendChild(thumb);chip.appendChild(nm);
    if(it.uploading||it.posting){
      var pg=document.createElement('div');pg.className='aq-progress';
      var pf=document.createElement('div');pf.className='aq-pf';pf.style.width=(it.posting?100:(it.progress||0))+'%';
      pg.appendChild(pf);chip.appendChild(pg);
      var pc=document.createElement('div');pc.className='aq-pct';
      pc.textContent=it.posting?'发送中':((it.progress||0)+'%');chip.appendChild(pc);
    }
    chip.appendChild(rm);list.appendChild(chip);
  });
}
function clearQueue(){pq=[];renderAttachQueue();hideUploadBanner()}
function resizeInput(){
  var el=$('ci2');if(!el)return;
  el.style.height='44px';
  el.style.height=Math.min(Math.max(el.scrollHeight,44),120)+'px';
}
function uploadErrMsg(e){
  var m=String(e&&e.message?e.message:e||'').trim();
  if(!m||m==='失败'||m==='fail')return '上传失败，请重试';
  if(m==='network')return '网络中断，请保持 WiFi 连接并重试';
  if(m.indexOf('upload incomplete')>=0){
    var mm=m.match(/got\s+([\d.]+)MB\s*\/\s*([\d.]+)MB/i);
    if(mm)return '上传不完整（只收到 '+mm[1]+'MB / '+mm[2]+'MB），请重试';
    return '上传不完整，请重试';
  }
  if(m.indexOf('zip file incomplete')>=0||m.indexOf('corrupted')>=0)return '压缩包不完整或已损坏，请重新下载后再传';
  if(m.indexOf('服务器处理超时')>=0)return '服务器处理超时，文件过大请用电脑上传';
  if(m.indexOf('上传超时')>=0)return '上传超时，请用电脑浏览器拖文件上传';
  if(m.indexOf('HTTP 413')>=0)return '文件超过服务器大小限制';
  if(m.indexOf('server error during upload')>=0)return '服务器处理上传时出错，请重试或用电脑上传';
  if(m.indexOf('HTTP 400')>=0&&m.indexOf('error')>=0){
    try{var j=JSON.parse(m.slice(m.indexOf('{')));if(j.error)return uploadErrMsg(j.error)}catch(_){}
  }
  if(/^[a-zA-Z0-9\s,.:()\/\-]+$/.test(m))return '上传失败：'+m;
  return m;
}
async function uploadOne(file,onProgress){
  return new Promise(function(res,rej){
    var x=new XMLHttpRequest(),fd=new FormData();fd.append('file',file);
    var bytesDone=false,t0=Date.now(),lastPct=-1;
    var size=file.size||0;
    var timeoutMs=size>100*1048576?7200000:(size>20*1048576?5400000:3600000);
    x.upload.addEventListener('progress',function(e){
      if(!e.lengthComputable)return;
      var pct=Math.round(e.loaded/e.total*100);
      if(onProgress)onProgress(pct);
      if(pct!==lastPct&&pct>0&&pct<100){
        var sec=(Date.now()-t0)/1000;
        if(sec>0.3){
          var spd=(e.loaded/1024/1024/sec).toFixed(1);
          var st=$('ub-status');if(st&&$('upload-banner').classList.contains('show')){
            st.textContent='上传中 '+pct+'% · '+spd+' MB/s'+(size?' · '+fmtSize(size):'');
          }
        }
      }
      lastPct=pct;
      if(e.loaded>=e.total)bytesDone=true;
    });
    x.addEventListener('load',function(){
      if(x.status===200){
        try{
          var d=JSON.parse(x.responseText);
          if(!d.ok){rej(new Error(d.error||'fail'));return}
          res({url:d.url||dlPath(d.relpath||UPLOAD_REL+'/'+d.name),name:d.name,relpath:d.relpath});
        }catch(_){rej(new Error('bad response'))}
      }else{
        var errTxt=(x.responseText||'').slice(0,200);
        try{var jd=JSON.parse(x.responseText);if(jd&&jd.error){rej(new Error(jd.error));return}}catch(_){}
        rej(new Error('HTTP '+x.status+' '+errTxt));
      }
    });
    x.addEventListener('error',function(){rej(new Error('network'))});
    x.addEventListener('timeout',function(){rej(new Error(bytesDone?'服务器处理超时':'上传超时，文件过大或网络太慢'))});
    x.open('POST','/api/upload');
    x.timeout=timeoutMs;
    x.upload.addEventListener('loadend',function(){x.timeout=600000});
    x.send(fd);
  });
}
async function processIncomingFiles(files){
  if(!files.length)return;
  if(ct!=='chat')sw('chat');
  var sendMode='separate';
  if(files.length>=2){
    sendMode=await askPackSendChoice(files.length);
    if(!sendMode)return;
  }
  var doPack=(sendMode==='pack');
  var big=files.filter(function(f){return (f.size||0)>20*1048576});
  if(HUB_MODE==='cloud'&&big.length){
    toast('大文件走云端较慢（受宽带上行限制），在家请用局域网地址');
  }
  uploadPipeline++;
  var ok=0,fail=0,total=files.length,uploaded=[];
  try{
  for(var i=0;i<files.length;i++){
    var slot={url:'',name:files[i].name,uploading:true,ok:false,progress:0,posting:false};
    pq.push(slot);renderAttachQueue();
    showUploadBanner(slot,i,total);
    try{
      var d=await uploadOne(files[i],function(pct){slot.progress=pct;showUploadBanner(slot,i,total);renderAttachQueue()});
      slot.url=d.url;slot.name=d.name;slot.uploading=false;slot.progress=100;
      uploaded.push({url:d.url,name:d.name,relpath:d.relpath||relFromUrl(d.url)});
      showUploadBanner(slot,i,total);
      ok++;
    }catch(e){
      slot.uploading=false;slot.posting=false;slot.ok=false;slot.progress=100;
      var errMsg=uploadErrMsg(e);
      slot.error=errMsg;
      showUploadBanner(slot,i,total,'er');renderAttachQueue();
      fail++;
      console.error('upload',files[i].name,e);
      toast('失败: '+files[i].name+' · '+errMsg);
    }
  }
  if(uploaded.length){
    try{
      showUploadBanner({name:uploaded.length+' 个文件',progress:100},0,uploaded.length,'post');
      try{
        await sendUploadedItems(uploaded,doPack);
      }catch(e){
        if(doPack){
          console.error('pack',e);
          toast('打包失败，改为逐条发出');
          await sendUploadedItems(uploaded,false);
        }else throw e;
      }
    }catch(e){toast('发送失败')}
    await fm(lid);scb(true);
  }
  pq=[];renderAttachQueue();
  setTimeout(hideUploadBanner,1500);
  if(!uploaded.length&&fail)toast('上传全部失败');
  }finally{uploadPipeline=Math.max(0,uploadPipeline-1)}
}
async function pickAndAttach(input){
  var files=Array.from(input.files||[]);input.value='';
  await processIncomingFiles(files);
}
async function postMsgAuto(url,name){
  var tag=isVid(name)?'[[VID]]':isImg(name)?'[[IMG]]':'[[FILE]]';
  await postMsg(autoSender(),autoSenderName(),tag+fixImgUrl(url));
}
function vidBody(url,fn){
  url=fixImgUrl(url);
  var body=document.createElement('div');
  body.className='mb mvid';
  var vid=document.createElement('video');
  vid.src=url;vid.controls=true;vid.playsInline=true;vid.preload='metadata';
  vid.muted=false;vid.setAttribute('playsinline','');vid.setAttribute('webkit-playsinline','');
  var exp=document.createElement('button');
  exp.type='button';exp.className='v-exp';exp.title='全屏';exp.textContent='⛶';
  exp.addEventListener('click',function(e){e.stopPropagation();previewFile(url,fn)});
  body.appendChild(vid);body.appendChild(exp);
  addCopyBtn(body,url,fn);
  return body;
}
async function postMsg(snd,snm,content){
  var r=await fetch('/api/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({sender:snd,sender_name:snm,content:content})});
  if(!r.ok)throw new Error('post failed');
}
async function loadMdInline(body,url,name){
  body.className='mb md-inline';
  body.innerHTML='<div class="md-body" style="color:var(--ad)">加载中…</div>';
  try{
    var r=await fetch(url);if(!r.ok)throw new Error('HTTP '+r.status);
    var t=await r.text();
    var inner=body.querySelector('.md-body')||document.createElement('div');
    inner.className='md-body';inner.innerHTML=mdToHtml(t);
    body.innerHTML='';body.appendChild(inner);
    addCopyBtn(body,url,name);
  }catch(e){
    body.className='mb';body.style.cursor='pointer';
    body.innerHTML=si('file',14)+' '+eh(name)+' <span style="color:var(--rd);font-size:12px">(加载失败，点击预览)</span>';
    body.onclick=function(){previewFile(url,name)};
    addCopyBtn(body,url,name);
  }
}
function attachBody(url,fn,row){
  url=fixImgUrl(url);
  var body=document.createElement('div');
  if(isImg(fn)){
    body.className='mb mib';
    var img=document.createElement('img');
    img.loading='lazy';img.decoding='async';
    img.className='img-loading';
    img.setAttribute('data-full',url);
    img.addEventListener('click',function(e){if(selMode)return;previewFile(url,fn)});
    img.addEventListener('load',function(){img.classList.remove('img-loading');img.classList.add('img-ready')});
    img.addEventListener('error',function(){
      img.onerror=null;
      if(img.src!==url){img.src=url;return}
      img.classList.remove('img-loading');
    });
    if(isHeic(fn))loadHeicInto(img,url);
    else{img.src=chatImgUrl(url,fn)}
    body.appendChild(img);
    addCopyBtn(body,url,fn);
  }else if(isVid(fn)){
    body=vidBody(url,fn);
  }else if(isAud(fn)){
    body.className='mb maud';
    var audNm=document.createElement('div');audNm.className='aud-name';
    audNm.innerHTML=si('audio',14)+' '+eh(fn);
    body.appendChild(audNm);
    var aud=document.createElement('audio');
    aud.src=url;aud.controls=true;aud.preload='metadata';
    aud.setAttribute('controlsList','nodownload');
    body.appendChild(aud);
    addCopyBtn(body,url,fn);
  }else if(isMd(fn)){
    body=document.createElement('div');
    loadMdInline(body,url,fn);
  }else if(isDocx(fn)){
    body.className='mb mdoc';
    body.innerHTML='<div class="docx-title">'+si('file',16)+' '+eh(fn)+'</div><div class="docx-hint">Word · 点击预览正文</div>';
    body.addEventListener('click',function(e){if(selMode)return;previewFile(url,fn)});
    appendPathHint(body,url,fn);
    addCopyBtn(body,url,fn);
  }else if(isZip(fn)){
    body=zipBody(url,fn);
  }else{
    body.className='mb';body.style.cursor='pointer';
    body.innerHTML=si('file',14)+' '+eh(fn);
    body.addEventListener('click',function(e){if(selMode)return;previewFile(url,fn)});
    appendPathHint(body,url,fn);
    addCopyBtn(body,url,fn);
  }
  return body;
}
function cem(m){
  try{
  var row=document.createElement('div');row.className='mr '+(m.sender==='pc'?'pc':'ph');
  row.setAttribute('data-mid',m.id);
  row.setAttribute('data-sel-key','msg-'+m.id);
  var nm=m.sender==='pc'?'电脑':'手机';
  var au=imgFromContent(m.content);
  if(au){
    var fn='file';
    try{fn=decodeURIComponent((au.split('/').pop()||'file').split('?')[0])}catch(_){fn=(au.split('/').pop()||'file').split('?')[0]}
    var rel=relFromUrl(au);
    if(!rel&&fn)rel=UPLOAD_REL+'/'+fn;
    row.setAttribute('data-sel-type','file');
    row.setAttribute('data-sel-rel',rel||'');
    row.setAttribute('data-sel-url',fixImgUrl(au));
    row.setAttribute('data-sel-name',fn);
    row.setAttribute('data-attach-url',fixImgUrl(au));
    var cap='';
    if(m.content.indexOf('\n')>=0)cap=m.content.slice(m.content.indexOf('\n')+1).trim();
    var body=attachBody(au,fn,row);
    row.appendChild(body);
    if(cap){var cb=makeTextBody(cap);cb.style.marginTop='4px';row.appendChild(cb)}
  }else{
    row.setAttribute('data-sel-type','text');
    row.setAttribute('data-sel-text',m.content);
    row.setAttribute('data-raw-text',m.content);
    row.appendChild(makeTextBody(m.content));
  }
  var meta=document.createElement('div');meta.className='mm';
  meta.innerHTML='<span class="mm-nm">'+eh(nm)+'</span><span class="mm-time">'+eh(m.time)+'</span>';
  row.appendChild(meta);
  return row;
  }catch(ex){console.error('cem',m.id,ex);var d=document.createElement('div');d.className='mr ph';d.setAttribute('data-mid',m.id);d.innerHTML='<div class="mb">[渲染失败] '+eh(m.content.slice(0,50))+'</div>';return d}
}
var pt=null;function sp(){if(pt)clearInterval(pt);pt=setInterval(async function(){
  if(!lid)lid=lastSeenId();
  if(ct==='chat'){await fm(lid)}
  else{try{
    var r=await fetch('/api/messages?since='+lid+'&meta=1');
    var d=await r.json();
    if(d&&typeof d.epoch==='number'){
      if(msgEpoch&&d.epoch!==msgEpoch){
        msgEpoch=d.epoch;
        chatNeedsRefresh=true;
      }else msgEpoch=d.epoch;
    }
    var ms=(d&&d.messages)?d.messages:[];
    if(ms.length){ms.forEach(function(m){onIncoming(m);uc+=1});ubd();lid=ms[ms.length-1].id}
  }catch(e){}}
},1500)}
document.addEventListener('visibilitychange',function(){
  if(!document.hidden){
    if(document.title.indexOf('●')===0)document.title=document.title.replace(/^●\s*/,'');
    fetch('/api/info').then(function(r){return r.json()}).then(function(d){applyMsgEpoch(d.msg_epoch)}).catch(function(){});
  }
});

function uploadActive(){
  return uploadPipeline>0||pq.some(function(it){return it.uploading||it.posting});
}
async function sm(){
  var inp=$('ci2'),cnt=inp.value.trim();
  var items=pq.filter(function(it){return it.ok&&it.url});
  if(!items.length&&!cnt)return;
  if(items.length&&uploadActive()){toast('附件还在上传，请稍候');return}
  $('csb').disabled=true;
  try{
    if(cnt){
      inp.value='';resizeInput();clearDraft();
      await postMsg(autoSender(),autoSenderName(),cnt);
      await fm(lid);scb(true);sfxSend();toast('已发送');
    }
    if(items.length){
      var sendMode='separate';
      if(items.length>=2){
        sendMode=await askPackSendChoice(items.length);
        if(!sendMode)return;
      }
      var doPack=(sendMode==='pack');
      pq=[];renderAttachQueue();
      try{
        await sendUploadedItems(items.map(function(it){return {url:it.url,name:it.name,relpath:relFromUrl(it.url)}}),doPack);
      }catch(ex){
        toast('发送失败');
      }
      await fm(lid);scb(true);
      if(!cnt){sfxSend()}
    }
  }catch(e){toast('发送失败');if(cnt){inp.value=cnt;saveDraft()}}
  $('csb').disabled=false;
}
$('csb').addEventListener('click',sm);
$('ci2').addEventListener('keydown',function(e){
  if(e.key!=='Enter'||e.shiftKey)return;
  if(im)return;
  e.preventDefault();
  sm();
});
$('ci2').addEventListener('input',function(){resizeInput();saveDraft()});
var cfAll=$('cf-all'),cfText=$('cf-text');
if(cfAll)cfAll.addEventListener('click',function(){sfxClick();chatFilter='all';cfAll.classList.add('on');if(cfText)cfText.classList.remove('on');applyChatFilter()});
if(cfText)cfText.addEventListener('click',function(){sfxClick();chatFilter='text';cfText.classList.add('on');if(cfAll)cfAll.classList.remove('on');applyChatFilter();toast('已隐藏文件消息，文字都在')});
if($('cfile'))$('cfile').addEventListener('click',function(){sfxClick()});
if($('cgallery'))$('cgallery').addEventListener('click',function(){sfxClick()});
$('ci').addEventListener('change',function(){pickAndAttach($('ci'))});
$('cg').addEventListener('change',function(){pickAndAttach($('cg'))});
$('aq-clear').addEventListener('click',function(){clearQueue();toast('已清空')});
var selCancel=$('sel-cancel'),selDl=$('sel-dl'),selDel=$('sel-del'),selToggle=$('sel-toggle');
if(selCancel)selCancel.addEventListener('click',function(){sfxClick();exitSelMode()});
if(selToggle)selToggle.addEventListener('click',function(){sfxClick();toggleSelMode()});
if(selDl)selDl.addEventListener('click',function(){sfxClick();downloadSelected()});
if(selDel)selDel.addEventListener('click',function(){sfxClick();deleteSelected()});
var caEl=$('ca'),jumpBtn=$('jump-bottom'),flEl=$('fl');
if(caEl)caEl.addEventListener('scroll',updJumpBtn,{passive:true});
if(jumpBtn)jumpBtn.addEventListener('click',function(){sfxClick();scb(true)});
var loadOlderBtn=$('load-older');
if(loadOlderBtn)loadOlderBtn.addEventListener('click',function(){sfxClick();loadOlderMsgs()});
function bindSelEvents(el){
  if(!el)return;
  el.addEventListener('touchend',function(e){handleSelTarget(e)},{passive:false});
  if(!im)el.addEventListener('click',function(e){handleSelTarget(e)},true);
}
bindSelEvents(caEl);
bindSelEvents(flEl);

// Ctrl+V paste (desktop only)
document.addEventListener('paste',async function(e){
  if(im||ct!=='chat')return;
  var is=e.clipboardData&&e.clipboardData.items;if(!is)return;
  var imgs=[];
  for(var i=0;i<is.length;i++){if(is[i].type.indexOf('image')===0){var f=is[i].getAsFile();if(f)imgs.push(f)}}
  if(!imgs.length)return;
  e.preventDefault();
  await processIncomingFiles(imgs);
});

// Init
initActFeedback();
initSearchUI();
initPackChoiceUI();
var im=/Mobi|Android|iPhone|iPad/i.test(navigator.userAgent)||('ontouchstart' in window&&window.innerWidth<768);
fetch('/api/info').then(function(r){return r.json()}).then(applyHubInfo).catch(function(){});
setInterval(function(){fetch('/api/info').then(function(r){return r.json()}).then(applyHubInfo).catch(function(){})},30000);
$('ss').value=im?'phone':'pc';
if(im){
  $('ci2').placeholder='Enter 换行 · 必须点「发送」才会发出';
  currentCat='image';
}else{
  $('ci2').placeholder='输入文字 · 拖文件到页面上传\nShift+Enter 换行';
  initNotify();
  document.documentElement.classList.add('desktop');
  sw('files');
  (function(){
    var dragDepth=0,dov=$('drop-ov');
    function showDrop(v){if(dov)dov.classList.toggle('active',v)}
    document.addEventListener('dragenter',function(e){
      if(!e.dataTransfer||!e.dataTransfer.types||!Array.from(e.dataTransfer.types).includes('Files'))return;
      e.preventDefault();dragDepth++;showDrop(true);
    });
    document.addEventListener('dragleave',function(e){e.preventDefault();dragDepth=Math.max(0,dragDepth-1);if(dragDepth===0)showDrop(false)});
    document.addEventListener('dragover',function(e){
      if(!e.dataTransfer||!e.dataTransfer.types||!Array.from(e.dataTransfer.types).includes('Files'))return;
      e.preventDefault();e.dataTransfer.dropEffect='copy';
    });
    document.addEventListener('drop',function(e){
      if(!e.dataTransfer||!e.dataTransfer.files||!e.dataTransfer.files.length)return;
      e.preventDefault();dragDepth=0;showDrop(false);
      processIncomingFiles(Array.from(e.dataTransfer.files));
    });
  })();
}
sw('files');
loadFilesDefault();
var srcHome=$('src-home'),srcInbox=$('src-inbox');
if(srcHome)srcHome.addEventListener('click',function(){if(fileSrc!=='home'){sfxClick();browseHome()}});
if(srcInbox)srcInbox.addEventListener('click',function(){if(fileSrc!=='inbox'){sfxClick();browseReceived()}});
if(im){
  document.querySelectorAll('.cat-btn[data-cat]').forEach(function(b){b.classList.toggle('active',b.dataset.cat===currentCat)});
}
updSortUI();
bindGalleryNav();
bindPreviewBackdrop();
initChat();
</script></body></html>'''

# Inject icons
HTML = T
for k, v in {
    "__I1__":si("files",18),"__I2__":si("upload",18),"__I3__":si("chat",18),
    "__I4__":si("inbox",16),"__I5__":si("home",9),"__I6__":si("photo",22),
    "__I7__":si("file",22),"__I8__":si("upload",22),"__I9__":si("chat",22),
    "__I10__":si("attach",14),"__SPK__":si("speaker",11),
    "__I10b__":si("attach",20),"__I6b__":si("photo",20),
    "__LOGO__":si("logo",28),"__COPY__":si("copy",16),"__BATCH__":si("batch",18),"__TRASH__":si("trash",18),"__BELL__":si("bell",16),
    "__GRID__":si("grid",18),"__SELX__":si("x",18),"__SEARCH__":si("search",18),
    "__JSON__":json.dumps(SVG, ensure_ascii=False),
    "__UPLOAD_REL__":json.dumps(UPLOAD_REL, ensure_ascii=False),
    "__HUB_MODE__":json.dumps(HUB_MODE),
    "__HUB_TITLE__":_brand["title"],
    "__HUB_SUB__":_brand["sub"],
    "__HUB_BADGE__":_brand["badge"],
    "__PAGE_TITLE__":_brand["page"],
    "__BUILD__":json.dumps(HUB_BUILD),
}.items():
    HTML = HTML.replace(k, v)

# ══════ Handler ══════
class H(http.server.BaseHTTPRequestHandler):
    def do_HEAD(self):
        p=self.path.split("?")[0]
        if p.startswith("/api/download"):
            self._dl(p, head_only=True)
        elif p.startswith("/api/thumb/"):
            self._thumb(p, head_only=True)
        elif p=="/" or p=="/index.html":
            self.send_response(200)
            self.send_header("Content-Type","text/html; charset=utf-8")
            self.send_header("Cache-Control","no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma","no-cache")
            self.end_headers()
        else:
            self.send_error(404)

    def do_GET(self):
        p=self.path.split("?")[0]
        if p=="/" or p=="/index.html":self._html()
        elif p=="/api/info":
            url = (HUB_PUBLIC_URL + "/") if HUB_PUBLIC_URL else "http://{}:{}/".format(lan_ip(), PORT)
            self._j(200, {
                "lan_url": url,
                "lan_ips": lan_ips(),
                "port": PORT,
                "mode": HUB_MODE,
                "upload_dir": str(UPLOAD_DIR),
                "local_hub_client": client_is_hub_host(self.client_address[0]),
                "msg_epoch": msg_epoch,
            })
        elif p=="/api/browse":self._browse()
        elif p=="/api/received":self._received()
        elif p.startswith("/api/preview/docx-media/"):self._preview_docx_media(p)
        elif p.startswith("/api/preview/docx/"):self._preview_docx(p)
        elif p.startswith("/api/download"):self._dl(p)
        elif p.startswith("/api/thumb/"):self._thumb(p)
        elif p=="/api/messages":self._gmsg()
        elif p=="/api/search":self._search()
        elif p=="/icon.svg":self._icon()
        else:self.send_error(404)

    def do_POST(self):
        p=self.path.split("?")[0]
        if p=="/api/upload":self._up()
        elif p=="/api/messages":self._pmsg()
        elif p=="/api/batch-download":self._batch_dl()
        elif p=="/api/pack-save":self._pack_save()
        elif p=="/api/batch-delete":self._batch_delete()
        else:self.send_error(404)

    def _batch_delete(self):
        cl = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(cl)
        try:
            d = json.loads(body)
        except json.JSONDecodeError:
            self._j(400, {"ok": False, "error": "invalid json"}); return
        paths = d.get("paths", [])
        msg_ids = d.get("message_ids", [])
        if not isinstance(paths, list):
            paths = []
        if not isinstance(msg_ids, list):
            msg_ids = []
        if not paths and not msg_ids:
            self._j(400, {"ok": False, "error": "nothing to delete"}); return
        deleted_files = 0
        failed = []
        msg_id_set = set()
        for mid in msg_ids:
            try:
                msg_id_set.add(int(mid))
            except (TypeError, ValueError):
                pass
        seen = set()
        for rel in paths[:200]:
            if not isinstance(rel, str):
                continue
            rel = rel.strip().lstrip("/")
            if not rel or rel in seen:
                continue
            seen.add(rel)
            fp = deletable_file(rel)
            if not fp:
                failed.append(rel)
                continue
            for mid in related_message_ids_for_file(fp):
                msg_id_set.add(mid)
            try:
                caches = []
                for w in (120, 480, 960):
                    try:
                        caches.append(thumb_path_for(fp, w))
                    except Exception:
                        pass
                fp.unlink()
                deleted_files += 1
                for cache in caches:
                    if cache.is_file():
                        cache.unlink()
            except OSError as e:
                failed.append(rel)
                print("[delete] {} -> {}".format(rel, e))
        deleted_msgs = db_delete_messages(list(msg_id_set)) if msg_id_set else 0
        self._j(200, {
            "ok": True,
            "deleted_files": deleted_files,
            "deleted_messages": deleted_msgs,
            "failed": failed,
        })

    def _batch_dl(self):
        cl=int(self.headers.get("Content-Length",0));body=self.rfile.read(cl)
        try:d=json.loads(body)
        except json.JSONDecodeError:self._j(400,{"ok":False,"error":"invalid json"});return
        paths=d.get("paths",[])
        if not isinstance(paths,list) or not paths or len(paths)>200:
            self._j(400,{"ok":False,"error":"invalid paths"});return
        data, added = zip_paths_to_buffer(paths)
        if not data:
            self._j(404,{"ok":False,"error":"no files found"});return
        self.send_response(200)
        self.send_header("Content-Type","application/zip")
        self.send_header("Content-Disposition",'attachment; filename="lan-hub-batch.zip"')
        self.send_header("Content-Length",str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _pack_save(self):
        cl = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(cl)
        try:
            d = json.loads(body)
        except json.JSONDecodeError:
            self._j(400, {"ok": False, "error": "invalid json"}); return
        paths = d.get("paths", [])
        if not isinstance(paths, list) or len(paths) < 1 or len(paths) > 200:
            self._j(400, {"ok": False, "error": "invalid paths"}); return
        data, added = zip_paths_to_buffer(paths)
        if not data:
            self._j(404, {"ok": False, "error": "no files found"}); return
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        zip_name = "Relay-{}files-{}.zip".format(added, ts)
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        zip_path = UPLOAD_DIR / zip_name
        n = 1
        while zip_path.exists():
            zip_name = "Relay-{}files-{}-{}.zip".format(added, ts, n)
            zip_path = UPLOAD_DIR / zip_name
            n += 1
        zip_path.write_bytes(data)
        try:
            rp = zip_path.resolve().relative_to(SERVE_DIR.resolve()).as_posix()
        except ValueError:
            rp = UPLOAD_REL + "/" + zip_path.name
        self._j(200, {
            "ok": True,
            "name": zip_path.name,
            "relpath": rp,
            "url": dl_url(rp),
            "count": added,
            "size": len(data),
        })

    def _icon(self):
        fp = ASSETS_DIR / "hub-icon.svg"
        if not fp.is_file():
            self.send_error(404); return
        b = fp.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "image/svg+xml")
        self.send_header("Cache-Control", "public, max-age=86400")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _html(self):
        b=HTML.encode();self.send_response(200)
        self.send_header("Content-Type","text/html; charset=utf-8")
        self.send_header("Cache-Control","no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma","no-cache")
        self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b)

    def _ls(self,target):
        e=[]
        for f in sorted(target.iterdir(),key=lambda x:(not x.is_dir(),x.name.lower())):
            try:
                s=f.stat()
                try:rp=f.resolve().relative_to(SERVE_DIR.resolve()).as_posix()
                except ValueError:rp=f.name
                e.append({"name":f.name,"is_dir":f.is_dir(),"size":s.st_size if f.is_file() else 0,"mtime":int(s.st_mtime),"relpath":rp})
            except OSError:pass
        return e

    def _browse(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        dp=q.get("path",[""])[0];t=(SERVE_DIR/dp).resolve() if dp else SERVE_DIR
        try:t.relative_to(SERVE_DIR)
        except ValueError:self._j(403,{"error":"forbidden"});return
        if not t.is_dir():self._j(404,{});return
        self._j(200,{"entries":self._ls(t),"path":dp})

    def _received(self):
        UPLOAD_DIR.mkdir(parents=True,exist_ok=True)
        e=self._ls(UPLOAD_DIR)
        for x in e:
            x["is_dir"]=False
            x["relpath"]=UPLOAD_REL+"/"+x["name"]
        self._j(200,{"entries":e,"path":UPLOAD_REL})

    def _search(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        query=(q.get("q",[""])[0] or "").strip()
        scope=(q.get("scope",["all"])[0] or "all").strip().lower()
        try:
            limit=int(q.get("limit",["80"])[0])
        except ValueError:
            limit=80
        if not query:
            self._j(400, {"ok": False, "error": "empty query"}); return
        if len(query) > 200:
            query = query[:200]
        if scope not in ("all", "chat", "files"):
            scope = "all"
        msgs = []
        files = []
        truncated = False
        if scope in ("all", "chat"):
            msgs = db_search(query, limit)
        if scope in ("all", "files"):
            files, truncated = search_files(query, limit)
        self._j(200, {
            "ok": True,
            "q": query,
            "scope": scope,
            "mode": HUB_MODE,
            "messages": msgs,
            "files": files,
            "truncated": truncated,
        })

    def _preview_docx(self, p):
        rel = urllib.parse.unquote(p[len("/api/preview/docx"):])
        if not rel or rel == "/":
            self._j(400, {"ok": False, "error": "missing path"}); return
        fp = resolve_download(rel)
        if not fp or not fp.is_file():
            self._j(404, {"ok": False, "error": "not found"}); return
        if fp.suffix.lower() != ".docx":
            self._j(400, {"ok": False, "error": "not docx"}); return
        try:
            with open(fp, "rb") as f:
                head = f.read(4)
            if len(head) < 4 or head[:2] != b"PK":
                self._j(422, {
                    "ok": False,
                    "error": "corrupt",
                    "message": "文件已损坏（上传可能中断），不是有效的 Word 文档",
                }); return
            body = docx_to_html(fp, docx_media_base(rel))
            self._j(200, {"ok": True, "html": body})
        except zipfile.BadZipFile:
            self._j(422, {
                "ok": False,
                "error": "corrupt",
                "message": "文件已损坏（上传可能中断），不是有效的 Word 文档",
            })
        except ValueError as e:
            print("[docx-preview] {} -> {}".format(fp.name, e))
            self._j(422, {"ok": False, "error": "invalid", "message": str(e)})
        except Exception as e:
            print("[docx-preview] {} -> {}".format(fp.name, e))
            self._j(500, {"ok": False, "error": "preview failed", "message": "预览失败，请尝试重新上传"})

    def _preview_docx_media(self, p=None):
        path = (p or self.path).split("?")[0]
        prefix = "/api/preview/docx-media/"
        if not path.startswith(prefix):
            self.send_error(404); return
        rest = urllib.parse.unquote(path[len(prefix):])
        if "/" not in rest or ".." in rest:
            self.send_error(400); return
        doc_rel, media_name = rest.rsplit("/", 1)
        media_name = os.path.basename(media_name)
        if not media_name:
            self.send_error(400); return
        fp = resolve_download(doc_rel)
        if not fp or not fp.is_file() or fp.suffix.lower() != ".docx":
            self.send_error(404); return
        target = "word/media/" + media_name
        try:
            with zipfile.ZipFile(str(fp)) as zf:
                if target not in zf.namelist():
                    self.send_error(404); return
                data = zf.read(target)
        except (zipfile.BadZipFile, KeyError, OSError):
            self.send_error(404); return
        ct, _ = mimetypes.guess_type(media_name)
        ct = ct or "application/octet-stream"
        etag = '"{:x}-{:x}"'.format(int(fp.stat().st_mtime), len(data))
        if self.headers.get("If-None-Match") == etag:
            self.send_response(304); self.end_headers(); return
        self.send_response(200)
        self.send_header("Content-Type", ct)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("ETag", etag)
        self.send_header("Cache-Control", "public, max-age=604800, immutable")
        self.end_headers()
        self.wfile.write(data)

    def _send_file(self, fp, ct, force_dl=False, head_only=False):
        st = fp.stat()
        etag = file_etag(fp)
        if self.headers.get("If-None-Match") == etag:
            self.send_response(304)
            self.end_headers()
            return
        fs = st.st_size
        self.send_response(200)
        self.send_header("Content-Type", ct)
        self.send_header("Content-Length", str(fs))
        self.send_header("ETag", etag)
        self.send_header("Last-Modified", email.utils.formatdate(st.st_mtime, usegmt=True))
        if not force_dl:
            self.send_header("Cache-Control", "public, max-age=604800, immutable")
        disp = "attachment" if force_dl else "inline"
        self.send_header("Content-Disposition", content_disposition(disp, fp.name.split("?")[0]))
        self.end_headers()
        if head_only:
            return
        with open(fp, "rb") as f:
            while True:
                ck = f.read(65536)
                if not ck:
                    break
                self.wfile.write(ck)

    def _thumb(self, p, head_only=False):
        rel = urllib.parse.unquote(p[len("/api/thumb"):])
        if not rel or rel == "/":
            self.send_error(400); return
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        try:
            width = min(max(int(q.get("w", ["480"])[0]), 120), 1200)
        except (TypeError, ValueError):
            width = 480
        fp = resolve_download(rel)
        if not fp or not fp.is_file():
            self.send_error(404); return
        if not is_raster_image(fp):
            self._dl("/api/download/" + rel.lstrip("/"), head_only=head_only); return
        try:
            from PIL import Image  # noqa: F401
        except ImportError:
            self._dl("/api/download/" + rel.lstrip("/"), head_only=head_only); return
        cache = thumb_path_for(fp, width)
        try:
            if not cache.is_file() or cache.stat().st_mtime < fp.stat().st_mtime:
                generate_thumb(fp, cache, width)
        except Exception as e:
            print("[thumb] {} -> {}".format(fp.name, e))
            self._dl("/api/download/" + rel.lstrip("/"), head_only=head_only); return
        self._send_file(cache, "image/jpeg", head_only=head_only)

    def _dl(self,p=None, head_only=False):
        parsed=urllib.parse.urlparse(p or self.path)
        rel=urllib.parse.unquote(parsed.path[len("/api/download"):])
        force_dl=urllib.parse.parse_qs(parsed.query).get("dl",["0"])[0]=="1"
        if not rel or rel=="/":self.send_error(400);return
        fp=resolve_download(rel)
        if not fp:self.send_error(404);return
        fs=fp.stat().st_size;ct,_=mimetypes.guess_type(str(fp));ct=ct or "application/octet-stream"
        rh=self.headers.get("Range","")
        if rh.startswith("bytes=") and not head_only:
            try:
                r=rh[6:];s,e=r.split("-");start=int(s) if s else 0;end=int(e) if e else fs-1;end=min(end,fs-1)
                if start>end or start>=fs:self.send_error(416);return
                cl=end-start+1;self.send_response(206)
                self.send_header("Content-Range","bytes {}-{}/{}".format(start,end,fs))
                self.send_header("Content-Length",str(cl));self.send_header("Content-Type",ct)
                self.send_header("Accept-Ranges","bytes");self.end_headers()
                with open(fp,"rb") as f:
                    f.seek(start)
                    rem = cl
                    while rem:
                        ck = f.read(min(65536, rem))
                        if ck: self.wfile.write(ck); rem -= len(ck)
                return
            except(ValueError,IndexError):pass
        etag = file_etag(fp)
        if not force_dl and self.headers.get("If-None-Match") == etag:
            self.send_response(304); self.end_headers(); return
        self.send_response(200);self.send_header("Content-Length",str(fs))
        self.send_header("Content-Type",ct);self.send_header("Accept-Ranges","bytes")
        if not force_dl:
            self.send_header("ETag", etag)
            self.send_header("Last-Modified", email.utils.formatdate(fp.stat().st_mtime, usegmt=True))
            self.send_header("Cache-Control", "public, max-age=604800, immutable")
        disp="attachment" if force_dl else "inline"
        self.send_header("Content-Disposition", content_disposition(disp, fp.name.split("?")[0]));self.end_headers()
        if head_only:return
        with open(fp,"rb") as f:
            while True:
                ck = f.read(65536)
                if not ck: break
                self.wfile.write(ck)

    def _up(self):
        import time
        t0 = time.monotonic()
        try:
            ct = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in ct:
                self._j(400, {"ok": False, "error": "expected multipart"}); return
            bd = None
            for part in ct.split(";"):
                part = part.strip()
                if part.startswith("boundary="):
                    bd = part[9:].strip('"'); break
            if not bd:
                self._j(400, {"ok": False, "error": "no boundary"}); return
            cl = int(self.headers.get("Content-Length", 0))
            if cl:
                print("[upload] start {:.1f}MB from {}".format(cl / 1024 / 1024, self.client_address[0]))
            if HUB_MAX_UPLOAD and cl > HUB_MAX_UPLOAD:
                self._j(413, {"ok": False, "error": "file too large (max {}MB)".format(HUB_MAX_UPLOAD // (1024 * 1024))}); return
            fp, written, got = stream_multipart_upload(self.rfile, bd, cl)
            if not fp or not written:
                got_mb = got / 1024 / 1024 if got else 0
                expect_mb = cl / 1024 / 1024 if cl else 0
                print("[upload] incomplete {:.1f}/{:.1f}MB from {}".format(got_mb, expect_mb, self.client_address[0]))
                self._j(400, {"ok": False, "error": "upload incomplete, please retry (got {:.0f}MB / {:.0f}MB)".format(got_mb, expect_mb)}); return
            if got < cl:
                got_mb = got / 1024 / 1024
                expect_mb = cl / 1024 / 1024
                try:
                    fp.unlink()
                except OSError:
                    pass
                print("[upload] incomplete {:.1f}/{:.1f}MB from {}".format(got_mb, expect_mb, self.client_address[0]))
                self._j(400, {"ok": False, "error": "upload incomplete, please retry (got {:.0f}MB / {:.0f}MB)".format(got_mb, expect_mb)}); return
            if HUB_MAX_UPLOAD and written > HUB_MAX_UPLOAD:
                try:
                    fp.unlink()
                except OSError:
                    pass
                self._j(413, {"ok": False, "error": "file too large (max {}MB)".format(HUB_MAX_UPLOAD // (1024 * 1024))}); return
            sn = fp.name
            if sn.lower().endswith(".zip") and not zipfile.is_zipfile(str(fp)):
                try:
                    fp.unlink()
                except OSError:
                    pass
                print("[upload] reject bad zip {} {:.1f}MB from {}".format(sn, written / 1024 / 1024, self.client_address[0]))
                self._j(400, {"ok": False, "error": "zip file incomplete or corrupted, please retry"}); return
            elapsed = time.monotonic() - t0
            speed = (written / 1024 / 1024) / elapsed if elapsed > 0 else 0
            print("[upload] {} {:.0f}KB {:.2f}s {:.1f}MB/s recv from {}".format(
                sn, written / 1024, elapsed, speed, self.client_address[0]))
            try:
                rp = fp.resolve().relative_to(SERVE_DIR.resolve()).as_posix()
            except ValueError:
                rp = UPLOAD_REL + "/" + fp.name
            self._j(200, {"ok": True, "name": fp.name, "relpath": rp, "url": dl_url(rp)})
        except Exception as e:
            print("[upload] error {} from {}: {}".format(type(e).__name__, self.client_address[0], e))
            self._j(500, {"ok": False, "error": "server error during upload"})

    def handle(self):
        try:
            self.request.settimeout(7200)
        except OSError:
            pass
        super().handle()

    def _gmsg(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        anchor=q.get("anchor",[None])[0]
        since=q.get("since",[None])[0]
        before=q.get("before",[None])[0]
        meta=q.get("meta",[None])[0]
        try:
            limit=int(q.get("limit",["500"])[0])
        except ValueError:
            limit=500
        if anchor is not None:
            try:
                msgs, has_more_before, has_more_after, oldest = db_msgs_anchor(anchor)
            except (TypeError, ValueError):
                self._j(400, {"ok": False, "error": "invalid anchor"}); return
            self._j(200, {
                "messages": msgs,
                "total": db_msg_count(),
                "oldest_id": oldest,
                "has_more": has_more_before,
                "has_more_after": has_more_after,
                "epoch": msg_epoch,
                "anchor": int(anchor),
            })
            return
        msgs=db_msgs(since=since,before=before,limit=limit)
        if meta:
            total=db_msg_count()
            oldest=msgs[0]["id"] if msgs else 0
            self._j(200,{
                "messages": msgs,
                "total": total,
                "oldest_id": oldest,
                "has_more": oldest > 1,
                "epoch": msg_epoch,
            })
        else:
            self._j(200, msgs)

    def _pmsg(self):
        cl=int(self.headers.get("Content-Length",0));body=self.rfile.read(cl)
        try:d=json.loads(body)
        except json.JSONDecodeError:self._j(400,{"ok":False,"error":"invalid json"});return
        s=d.get("sender","phone");c=d.get("content","").strip();sn=d.get("sender_name","")
        if not c:self._j(400,{"ok":False,"error":"empty message"});return
        mid=db_add(s,c,sn);self._j(200,{"ok":True,"id":mid})

    def _j(self,code,data):
        b=json.dumps(data,ensure_ascii=False).encode()
        self.send_response(code);self.send_header("Content-Type","application/json; charset=utf-8")
        self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b)

    def log_message(self,fmt,*args):
        f=args[0] if args else""
        if isinstance(f,str)and f.startswith("/api/"):return
        print("[{}] {}".format(self.client_address[0],f))

class TS(socketserver.ThreadingMixIn,http.server.HTTPServer):daemon_threads=True;allow_reuse_address=True

if __name__=="__main__":
    ip=lan_ip()
    print("\n  NODE:Relay {} v4".format("Cloud" if HUB_MODE == "cloud" else "Local"))
    if HUB_PUBLIC_URL:
        print("  PUBLIC:   {}/".format(HUB_PUBLIC_URL))
    print("  TERMINAL: http://localhost:{}/".format(PORT))
    if BIND != "127.0.0.1":
        print("  MOBILE:   http://{}:{}/".format(ip, PORT))
    print("  SERVE:    {}".format(SERVE_DIR))
    print("  UPLOAD:   {}\n".format(UPLOAD_DIR))
    with TS((BIND, PORT), H) as srv:
        try:srv.serve_forever()
        except KeyboardInterrupt:print("\n  [SHUTDOWN]\n")
