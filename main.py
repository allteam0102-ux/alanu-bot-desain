# -*- coding: utf-8 -*-
"""
ALANU BOT DESAIN  (gaya @ala_nu — foto hitam-putih + panel teks serif)
======================================================================
Notion (Status "Siap Desain") -> gambar dari link tempel (Pinterest) ->
hitam-putih + grain -> rakit desain -> Telegram.
Output: .pptx (Canva) + .svg gabungan + .svg per slide.
Selesai: Status -> Terkirim & ikon halaman -> ✅.
"""

import os
import re
import io
import json
import html
import time
import base64
import random
import secrets
import shutil
import tempfile
import traceback

import requests
from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ======================================================================
# >>> PENGATURAN BRAND <<<
# ======================================================================
HANDLE        = os.environ.get("HANDLE")        or "@ala_nu"
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "E3B23C").lstrip("#")   # kuning/emas
LOGO_URL      = os.environ.get("LOGO_URL")      or ""
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Vremena Grotesk"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Vremena Grotesk"
ARROW_TEXT    = os.environ.get("ARROW_TEXT")    or "→"
DEFAULT_THEME = (os.environ.get("DEFAULT_THEME") or "Putih").strip().lower()
PHOTO_FRAC    = float(os.environ.get("PHOTO_FRAC") or 0.30)   # tinggi band foto (0.30 = foto lebih keatas/ringkas)

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG") or "false").lower() == "true"
OUTPUT_SVG_PER_SLIDE = (os.environ.get("OUTPUT_SVG_PER_SLIDE") or "false").lower() == "true"
OUTPUT_HTML = (os.environ.get("OUTPUT_HTML") or "true").lower() == "true"

# ----------------------------------------------------------------------
# SISTEM BELAJAR (memory + feedback). File memori disimpan di repo.
# ----------------------------------------------------------------------
MEMORY_PATH = os.environ.get("MEMORY_FILE") or "memory.json"
LEARN_ENABLED = (os.environ.get("LEARN_ENABLED") or "true").lower() == "true"

# "knob" yang bisa digeser oleh feedback (nilai awal = netral)
G_HEAD_SCALE = 1.0   # pengali ukuran judul
G_BODY_SCALE = 1.0   # pengali ukuran body
G_SCRIM      = 1.0   # pengali kegelapan overlay (inert di gaya B&W Ala NU)
G_SOFT       = 0.0   # tambahan "film/soft" (inert di gaya B&W Ala NU)
G_HEAD_BOLD  = True  # judul tebal (True) / tipis (False) -> digeser lewat revisi "jangan bold"

# Foto yg udah kepakai nggak boleh dipakai lagi sampai lewat sekian hari (default 6 bulan).
PHOTO_COOLDOWN_DAYS = int(os.environ.get("PHOTO_COOLDOWN_DAYS") or 180)

def _clamp(v, lo, hi):
    return max(lo, min(hi, v))

# ----------------------------------------------------------------------
# CACHE DESAIN (buat revisi per-slide: foto yg udah di-approve dipakai lagi)
# ----------------------------------------------------------------------
CACHE_DIR = os.environ.get("CACHE_DIR") or "design_cache"
def _pid_norm(pid): return (pid or "").replace("-", "")
def _page_cache_dir(pid):
    d = os.path.join(CACHE_DIR, _pid_norm(pid)); os.makedirs(d, exist_ok=True); return d
def _content_hash(s):
    import hashlib; return hashlib.md5((s or "").encode("utf-8")).hexdigest()[:12]
def _slide_nums_in(text):
    # ambil nomor slide dari teks revisi, mis. "slide 2 dan 4" -> [2,4]
    import re as _re
    low = (text or "").lower(); nums = []
    for mt in _re.finditer(r"slide\s*(\d{1,2})((?:\s*(?:,|dan|serta|&|/|\+)\s*\d{1,2})*)", low):
        nums.append(int(mt.group(1))); nums += [int(x) for x in _re.findall(r"\d{1,2}", mt.group(2))]
    seen, out = set(), []
    for n in nums:
        if n not in seen: seen.add(n); out.append(n)
    return out

# ukuran teks DASAR (identik dgn versi lama) lalu dikali knob sistem belajar.
# min ~10pt (pptx) / ~16px (svg) biar tetap kebaca walau knob mengecil.
HEAD_PT_BASE = 33    # judul (pptx)
BODY_PT_BASE = 18    # body  (pptx)
HEAD_PX_BASE = 52    # judul (svg)
BODY_PX_BASE = 30    # body  (svg)
def head_pt(): return max(10, int(round(HEAD_PT_BASE * G_HEAD_SCALE)))
def body_pt(): return max(10, int(round(BODY_PT_BASE * G_BODY_SCALE)))
def head_px(): return max(16, int(round(HEAD_PX_BASE * G_HEAD_SCALE)))
def body_px(): return max(16, int(round(BODY_PX_BASE * G_BODY_SCALE)))

# ---------- AUTO-FIT: body selalu di BAWAH judul + muat sebelum footer ----------
def _est_lines(text, width_px, font_px, bold):
    if not text or not text.strip():
        return 0
    cf = 0.58 if bold else 0.55
    max_chars = max(6, int(width_px / max(1.0, font_px * cf)))
    lines = 0
    for src in (text or "").split("\n"):
        words = []
        for seg, _hl in parse_highlights(src):
            for w in seg.split(" "):
                if w != "":
                    words.append(w)
        if not words:
            lines += 1; continue
        cur_len = 0
        for w in words:
            add = len(w) + (1 if cur_len > 0 else 0)
            if cur_len > 0 and cur_len + add > max_chars:
                lines += 1; cur_len = len(w)
            else:
                cur_len += add
        lines += 1
    return lines

def fit_text_layout(headline, body):
    """Ukuran font & posisi (fraksi) judul+body: body di BAWAH judul, muat sebelum footer.
    Font auto-kecil kalau kepanjangan."""
    W, H = CANVAS_W, CANVAS_H
    tw = 0.86 * W
    hpx = head_px(); bpx = body_px()
    HEAD_LH, BODY_LH, GAP = 1.30, 1.40, 0.02
    top0 = PHOTO_FRAC + 0.05
    body_min = top0 + 0.20
    bottom = 0.90
    for _ in range(24):
        hl = _est_lines(headline, tw, hpx, G_HEAD_BOLD)
        bl = _est_lines(body, tw, bpx, False)
        hh = hl * hpx * HEAD_LH / H
        bh = bl * bpx * BODY_LH / H
        btop = max(body_min, top0 + hh + (GAP if hl else 0.0))
        if btop + bh <= bottom:
            break
        if hpx > 26 and hh >= bh:
            hpx = max(26, int(hpx * 0.94))
        elif bpx > 16:
            bpx = max(16, int(bpx * 0.95))
        elif hpx > 18:
            hpx = max(18, int(hpx * 0.94))
        else:
            break
    hl = _est_lines(headline, tw, hpx, G_HEAD_BOLD)
    hh = hl * hpx * HEAD_LH / H
    btop = max(body_min, top0 + hh + (GAP if hl else 0.0))
    return {"head_px": hpx, "body_px": bpx, "head_top": top0, "body_top": btop}


# ======================================================================
# 1. TEKNIS (Secrets)
# ======================================================================
def env(name, default=None, required=False):
    val = os.environ.get(name, default)
    if required and (val is None or str(val).strip() == ""):
        raise SystemExit(f"[SETUP ERROR] Environment variable '{name}' belum diisi. Cek README.")
    return val

NOTION_TOKEN        = env("NOTION_TOKEN", required=True)
NOTION_DATABASE_ID  = env("NOTION_DATABASE_ID", required=True)
TELEGRAM_BOT_TOKEN  = env("TELEGRAM_BOT_TOKEN", required=True)
TELEGRAM_CHAT_ID    = env("TELEGRAM_CHAT_ID", required=True)
PEXELS_API_KEY      = os.environ.get("PEXELS_API_KEY", "")        # buat auto (kalau tak tempel link)
UNSPLASH_ACCESS_KEY = os.environ.get("UNSPLASH_ACCESS_KEY", "")   # buat auto
PIXABAY_API_KEY     = os.environ.get("PIXABAY_API_KEY", "")       # buat auto (sumber tambahan)
# PINTEREST (board pribadi) — daftar URL RSS board, dipisah koma. Diutamakan kalau diisi.
# contoh: https://www.pinterest.com/USERNAME/NAMA-BOARD.rss  (atau /feed.rss utk semua pin)
PINTEREST_RSS       = [u.strip() for u in (os.environ.get("PINTEREST_RSS") or "").split(",") if u.strip()]
PINTEREST_FIRST     = (os.environ.get("PINTEREST_FIRST") or "true").lower() == "true"
STYLE_HINT          = os.environ.get("STYLE_HINT") or "cinematic moody aesthetic silhouette dramatic light"
# bias Islami: kata yg ditambahkan biar hasil foto tetap bernuansa muslim/islami
ISLAMIC_HINT        = os.environ.get("ISLAMIC_HINT") or "muslim islamic"

TITLE_PROPERTY   = env("TITLE_PROPERTY", "Judul")
STATUS_PROPERTY  = env("STATUS_PROPERTY", "Status")
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Format")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Konten")
IMGURL_PROPERTY  = env("IMGURL_PROPERTY", "Gambar URL")
KEYWORD_PROPERTY = env("KEYWORD_PROPERTY", "Kata Kunci Gambar")
THEME_PROPERTY   = env("THEME_PROPERTY", "Tema")

STATUS_READY = env("STATUS_READY", "Siap Desain")
STATUS_DONE  = env("STATUS_DONE",  "Terkirim")
STATUS_ERROR = env("STATUS_ERROR", "Gagal")
STATUS_TYPE  = env("STATUS_TYPE", "select").strip().lower()
DONE_EMOJI   = env("DONE_EMOJI", "✅")

PEXELS_ORIENTATION = "portrait"

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

ACCENT = hex_rgb(ACCENT_COLOR)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

EMU_W = CANVAS_W * 9525
EMU_H = CANVAS_H * 9525
def fx(f): return Emu(int(f * EMU_W))
def fy(f): return Emu(int(f * EMU_H))

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"}


# ======================================================================
# 2. NOTION
# ======================================================================
NOTION_BASE = "https://api.notion.com/v1"
NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

def notion_find_ready():
    url = f"{NOTION_BASE}/databases/{NOTION_DATABASE_ID}/query"
    if STATUS_TYPE == "status":
        flt = {"property": STATUS_PROPERTY, "status": {"equals": STATUS_READY}}
    else:
        flt = {"property": STATUS_PROPERTY, "select": {"equals": STATUS_READY}}
    resp = requests.post(url, headers=NOTION_HEADERS, json={"filter": flt}, timeout=60)
    if resp.status_code != 200:
        raise SystemExit(f"[NOTION ERROR] {resp.status_code}: {resp.text}")
    return resp.json().get("results", [])

def notion_get_page(page_id):
    """Ambil ulang 1 halaman Notion by ID (buat auto-revisi)."""
    try:
        r = requests.get(f"{NOTION_BASE}/pages/{page_id}", headers=NOTION_HEADERS, timeout=60)
        if r.status_code == 200:
            return r.json()
    except Exception as e:
        print("  ! notion_get_page gagal:", e)
    return None

def notion_set_status(page_id, status_value, note=None, icon_emoji=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    body = {"properties": props_with_note}
    if icon_emoji:
        body["icon"] = {"type": "emoji", "emoji": icon_emoji}
    r = requests.patch(url, headers=NOTION_HEADERS, json=body, timeout=60)
    if r.status_code != 200:
        body2 = {"properties": props}
        if icon_emoji:
            body2["icon"] = {"type": "emoji", "emoji": icon_emoji}
        requests.patch(url, headers=NOTION_HEADERS, json=body2, timeout=60)

def _plain_text(rich_list):
    return "".join(part.get("plain_text", "") for part in (rich_list or []))

def read_property(props, name, kind):
    p = props.get(name)
    if not p:
        return ""
    if kind == "title":
        return _plain_text(p.get("title", [])).strip()
    if kind == "rich_text":
        return _plain_text(p.get("rich_text", [])).strip()
    if kind == "url":
        return (p.get("url") or "").strip()
    if kind == "select":
        return (p.get("select") or {}).get("name", "").strip()
    if kind == "status":
        return (p.get("status") or {}).get("name", "").strip()
    return ""

def get_title(props):
    t = read_property(props, TITLE_PROPERTY, "title")
    if t:
        return t
    for _n, p in props.items():
        if isinstance(p, dict) and p.get("type") == "title":
            return _plain_text(p.get("title", [])).strip()
    return ""


# ======================================================================
# 3. PARSING
# ======================================================================
def split_slides(content_text):
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    # Pemisah slide: baris 'Slide 1/2/3...' ATAU baris '---' / dash.
    blocks = re.split(r"(?mi)^\s*(?:slide\s*\d+\s*[:.)\-]?|[-–—]{2,}|[–—])\s*$", raw)
    return [b.strip() for b in blocks if b.strip()]

def split_urls(text):
    raw = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub(r"(?m)^\s*(?:[-–—]{2,}|[–—])\s*$", "\n", raw)
    return [u.strip() for u in raw.split("\n") if u.strip()]

URL_RE = re.compile(r"https?://\S+")

def extract_url_from_block(block):
    url = None
    kept = []
    for line in block.split("\n"):
        m = URL_RE.search(line)
        if m and url is None:
            url = m.group(0).rstrip(").,")
            sisa = URL_RE.sub("", line).strip()
            if sisa:
                kept.append(sisa)
        else:
            kept.append(line)
    return url, "\n".join(kept).strip()

def headline_and_body(block):
    lines = block.split("\n")
    idx = next((i for i, l in enumerate(lines) if l.strip()), None)
    if idx is None:
        return "", ""
    return lines[idx].strip(), "\n".join(lines[idx + 1:]).strip()

def parse_highlights(text):
    out = []
    for part in re.split(r"(==.+?==)", text):
        if len(part) >= 4 and part.startswith("==") and part.endswith("=="):
            out.append((part[2:-2], True))
        elif part != "":
            out.append((part, False))
    return out


# ======================================================================
# 4. GAMBAR -> hitam-putih + grain
# ======================================================================
def fetch_image_bytes(url):
    r = requests.get(url, headers=UA, timeout=60)
    r.raise_for_status()
    ct = r.headers.get("Content-Type", "").lower()
    if "image" not in ct:
        raise ValueError(
            f"URL bukan gambar langsung (Content-Type: {ct or 'tidak diketahui'}). "
            f"Di Pinterest: klik kanan gambar -> 'Copy image address' (link diakhiri .jpg/.png)."
        )
    return r.content

def keyword_for(index, keyword_blocks, headline, body):
    if index < len(keyword_blocks) and keyword_blocks[index].strip():
        q = keyword_blocks[index].strip()
    elif headline:
        q = headline
    elif body:
        q = body
    else:
        q = "cinematic silhouette"
    q = re.sub(r"==", "", q)
    words = re.sub(r"[^\w\s]", " ", q).split()
    return " ".join(words[:4]) if words else "cinematic silhouette"

def _short_q(query):
    words = (query or "").split()[:3]
    return " ".join(words) if words else "aesthetic portrait"

def _islamic_q(query):
    """Query pendek + dipastikan ada nuansa Islami (muslim/hijab/mosque/dll),
    tapi tetap bawa konteks dari isi konten."""
    base = _short_q(query)
    low = base.lower()
    keys = ("muslim", "islam", "hijab", "mosque", "masjid", "quran", "pray", "ramadan", "ummah", "sholat")
    if not any(k in low for k in keys):
        base = (base + " " + ISLAMIC_HINT).strip()
    return base

def pexels_pick(query, used_ids):
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 15, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    q = _islamic_q(query)
    photos = _search(q) or _search(q + " aesthetic") or _search("muslim praying silhouette")
    if not photos:
        return None
    fresh = [p for p in photos if str(p.get("id")) not in used_ids]
    pool = fresh if fresh else photos
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(str(photo["id"]))
    u = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(u, timeout=60).content

def pixabay_pick(query, used_ids):
    r = requests.get("https://pixabay.com/api/",
                     params={"key": PIXABAY_API_KEY, "q": _islamic_q(query),
                             "image_type": "photo", "orientation": "vertical",
                             "safesearch": "true", "per_page": 20}, timeout=60)
    r.raise_for_status()
    hits = r.json().get("hits", [])
    if not hits:
        return None
    fresh = [p for p in hits if str(p.get("id")) not in used_ids]
    pool = fresh if fresh else hits
    photo = random.choice(pool[:10])
    if photo.get("id"):
        used_ids.add(str(photo["id"]))
    u = photo.get("largeImageURL") or photo.get("webformatURL")
    return requests.get(u, headers=UA, timeout=60).content

def unsplash_pick(query, used_ids):
    r = requests.get("https://api.unsplash.com/search/photos",
                     headers={"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"},
                     params={"query": _islamic_q(query), "per_page": 15,
                             "orientation": PEXELS_ORIENTATION}, timeout=60)
    r.raise_for_status()
    results = r.json().get("results", [])
    if not results:
        return None
    fresh = [p for p in results if str(p.get("id")) not in used_ids]
    pool = fresh if fresh else results
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(str(photo["id"]))
    urls = photo.get("urls", {})
    return requests.get(urls.get("regular") or urls.get("full") or urls.get("raw"), timeout=60).content

# ---------- RIWAYAT FOTO: biar foto yg udah kepakai nggak muncul lagi < 6 bulan ----------
def photo_block_set(mem):
    """Kumpulan id foto (str) yg masih 'cooldown' + prune yg udah lewat 6 bulan (boleh lagi)."""
    hist = mem.get("photo_history") or {}
    now = int(time.time()); cutoff = PHOTO_COOLDOWN_DAYS * 86400
    block, keep = set(), {}
    for pid, ts in hist.items():
        try:
            ts = int(ts)
        except Exception:
            continue
        if now - ts < cutoff:
            block.add(str(pid)); keep[str(pid)] = ts
    mem["photo_history"] = keep
    return block

def record_photos_used(mem, ids):
    """Catat id foto (str) yg baru kepakai + timestamp sekarang."""
    if not ids:
        return
    hist = mem.setdefault("photo_history", {})
    now = int(time.time())
    for pid in ids:
        if pid:
            hist[str(pid)] = now

def _pin_sizes(url):
    """Daftar kandidat URL dari yg resolusi GEDE ke kecil: originals -> 736x -> aslinya (236x)."""
    import re as _re
    if not url:
        return []
    cands = []
    for size in ("originals", "736x"):
        c = _re.sub(r"/\d+x\d*/", f"/{size}/", url)
        if c not in cands:
            cands.append(c)
    if url not in cands:
        cands.append(url)
    return cands

def pinterest_pick(used_ids):
    """Ambil foto dari RSS board Pinterest pribadi (PINTEREST_RSS). None kalau mati/kosong.
    Aman: kalau gagal, pemanggilnya lanjut ke sumber stok."""
    if not PINTEREST_RSS:
        return None
    import re as _re
    feeds = list(PINTEREST_RSS); random.shuffle(feeds)
    for feed in feeds:
        try:
            r = requests.get(feed, headers=UA, timeout=60)
            if r.status_code != 200 or "<item" not in r.text:
                print(f"    ! Pinterest RSS '{feed}' status {r.status_code} / bukan feed valid")
                continue
            txt = r.text
            imgs  = _re.findall(r'<media:content[^>]+url="([^"]+)"', txt)
            imgs += _re.findall(r'<enclosure[^>]+url="([^"]+)"', txt)
            imgs += _re.findall(r'&lt;img[^&]*?src=&quot;([^&]+?)&quot;', txt)   # <img> di description yg ter-escape
            imgs += _re.findall(r'<img[^>]+src="([^"]+)"', txt)
            imgs = [u for u in imgs if "pinimg.com" in u]
            if not imgs:
                print(f"    ! Pinterest RSS '{feed}' nggak ada gambar pinimg")
                continue
            seen, uniq = set(), []
            for u in imgs:
                key = u.split("/")[-1].split("?")[0]
                if key in seen:
                    continue
                seen.add(key); uniq.append((key, u))
            fresh = [(k, u) for k, u in uniq if str(k) not in used_ids] or uniq
            key, u = random.choice(fresh)
            used_ids.add(str(key))
            for cand in _pin_sizes(u):
                try:
                    data = requests.get(cand, headers=UA, timeout=60).content
                    Image.open(io.BytesIO(data)).verify()
                    return data
                except Exception:
                    continue
        except Exception as e:
            print(f"    ! Pinterest RSS '{feed}' gagal: {e}")
    return None

def auto_image_bytes(query, used_ids):
    """Ambil otomatis (tanpa link). URUTAN: Pinterest board pribadi dulu (kalau diisi),
    baru Pixabay + Pexels + Unsplash. Semua dibias Islami + sesuai konteks."""
    # 0) PINTEREST (board pribadi) diutamakan
    if PINTEREST_RSS and PINTEREST_FIRST:
        try:
            data = pinterest_pick(used_ids)
            if data:
                print("    [sumber foto: Pinterest board kamu]")
                return data
            print("    (Pinterest kosong/mati -> lanjut ke stok)")
        except Exception as e:
            print("    ! Pinterest gagal, lanjut ke stok:", e)
    picker = {"unsplash": unsplash_pick, "pexels": pexels_pick, "pixabay": pixabay_pick}
    sources = []
    if PIXABAY_API_KEY:     sources.append("pixabay")
    if UNSPLASH_ACCESS_KEY: sources.append("unsplash")
    if PEXELS_API_KEY:      sources.append("pexels")
    # query utama (konteks konten) -> kalau mentok, fallback islami umum
    for q in [query, "muslim praying silhouette", "woman hijab contemplative", "mosque dramatic light"]:
        for src in sources:
            try:
                data = picker[src](q, used_ids)
                if data:
                    return data
            except Exception as e:
                print(f"    ! {src} gagal ('{_islamic_q(q)}'): {e}")
    return None

# ======================================================================
# AUTO-CROP CERDAS (3 lapis): AI focus -> deteksi wajah -> saliency -> crop tengah
# ======================================================================
SMART_CROP_ENABLED = (os.environ.get("SMART_CROP") or "true").lower() == "true"

def _box_crop_wh(im, box_px, w, h):
    """Crop im jadi rasio w:h, se-ketat mungkin TAPI wajib memuat box_px (x0,y0,x1,y1).
    Balikin (cropped_resized, rect_fraksi)."""
    target = w / h
    W, H = im.size
    bx0, by0, bx1, by1 = box_px
    bw = max(1.0, bx1 - bx0); bh = max(1.0, by1 - by0)
    cx = (bx0 + bx1) / 2.0; cy = (by0 + by1) / 2.0
    cw = max(bw, bh * target); ch = cw / target
    if ch < bh:
        ch = bh; cw = ch * target
    cw = min(cw, W); ch = min(ch, H)
    if cw / ch > target:
        cw = ch * target
    else:
        ch = cw / target
    x0 = cx - cw / 2.0; y0 = cy - ch / 2.0
    x0 = max(0, min(W - cw, x0)); y0 = max(0, min(H - ch, y0))
    im2 = im.crop((int(round(x0)), int(round(y0)), int(round(x0 + cw)), int(round(y0 + ch))))
    rect = (x0 / W, y0 / H, (x0 + cw) / W, (y0 + ch) / H)
    return im2.resize((w, h), Image.LANCZOS), rect

def ai_focus_box(image_path, context):
    """Tanya MATA AI (Gemini): bagian foto mana yg PALING nyambung konteks. (x,y,w,h) fraksi atau None."""
    if not GEMINI_API_KEY:
        return None
    prompt = (
        "Kamu fotografer/editor. Ini satu FOTO untuk konten Instagram.\n"
        f"Konteks/topik slide: \"{(context or '').strip()[:200]}\".\n"
        "Tentukan KOTAK bagian foto yang PALING relevan & penting untuk konteks ini "
        "(boleh objek/subjek yang paling nyambung dengan topik, BUKAN selalu yang paling menonjol). "
        "Jaga wajah/objek penting tetap utuh (jangan kepotong).\n"
        "Nilai dalam FRAKSI 0..1 relatif ukuran foto: x,y = pojok kiri-atas kotak; w,h = lebar,tinggi.\n"
        'Balas HANYA JSON: {"x":<0..1>,"y":<0..1>,"w":<0..1>,"h":<0..1>}'
    )
    d = _gemini_vision_json(prompt, image_path)
    try:
        x = float(d["x"]); y = float(d["y"]); w = float(d["w"]); h = float(d["h"])
    except Exception:
        return None
    if not (0 <= x <= 1 and 0 <= y <= 1 and 0 < w <= 1 and 0 < h <= 1):
        return None
    if x + w > 1: w = 1 - x
    if y + h > 1: h = 1 - y
    return (x, y, w, h)

def _face_box(im):
    """Deteksi wajah (depan + samping) pakai OpenCV. None kalau nggak ada / cv2 belum keinstall."""
    try:
        import cv2, numpy as np
    except Exception:
        return None
    try:
        arr = np.array(im.convert("RGB"))[:, :, ::-1]
        gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)
        boxes = []
        for xml in ("haarcascade_frontalface_default.xml", "haarcascade_profileface.xml"):
            c = cv2.CascadeClassifier(cv2.data.haarcascades + xml)
            if c.empty():
                continue
            for (x, y, w, h) in c.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(40, 40)):
                boxes.append((int(x), int(y), int(w), int(h)))
        if not boxes:
            return None
        H, W = gray.shape[:2]
        x0 = min(b[0] for b in boxes); y0 = min(b[1] for b in boxes)
        x1 = max(b[0] + b[2] for b in boxes); y1 = max(b[1] + b[3] for b in boxes)
        return (x0 / W, y0 / H, (x1 - x0) / W, (y1 - y0) / H)
    except Exception as e:
        print("    ! deteksi wajah gagal:", e)
        return None

def _saliency_box(im):
    """Fokus berbasis 'keramaian detail' (edge energy) pakai numpy saja — jalan walau opencv nggak ada."""
    try:
        import numpy as np
        ww = 200; hh = max(1, int(ww * im.size[1] / max(1, im.size[0])))
        g = np.asarray(im.convert("L").resize((ww, hh)), dtype=np.float32)
        gx = np.pad(np.abs(np.diff(g, axis=1)), ((0, 0), (0, 1)))
        gy = np.pad(np.abs(np.diff(g, axis=0)), ((0, 1), (0, 0)))
        energy = gx + gy
        total = float(energy.sum())
        if total < 1.0:
            return None
        H, W = energy.shape
        yy, xx = np.indices((H, W))
        cx = float((xx * energy).sum() / total) / W
        cy = float((yy * energy).sum() / total) / H
        sx = (float(((xx / W - cx) ** 2 * energy).sum() / total)) ** 0.5
        sy = (float(((yy / H - cy) ** 2 * energy).sum() / total)) ** 0.5
        hw = min(0.5, max(0.18, 2.0 * sx)); hh2 = min(0.5, max(0.18, 2.0 * sy))
        x0 = max(0.0, cx - hw); y0 = max(0.0, cy - hh2)
        x1 = min(1.0, cx + hw); y1 = min(1.0, cy + hh2)
        if (x1 - x0) >= 0.96 and (y1 - y0) >= 0.96:
            return None
        return (x0, y0, x1 - x0, y1 - y0)
    except Exception as e:
        print("    ! saliency gagal:", e)
        return None

PERSON_DETECT = (os.environ.get("PERSON_DETECT") or "true").lower() == "true"
_LAST_LOCAL = None

def _person_box(im):
    """Deteksi ORANG/badan pakai HOG+SVM bawaan OpenCV (nggak perlu download model)."""
    if not PERSON_DETECT:
        return None
    try:
        import cv2, numpy as np
    except Exception:
        return None
    try:
        arr = np.array(im.convert("RGB"))[:, :, ::-1]
        H0, W0 = arr.shape[:2]
        scale = 720.0 / max(1, W0)
        small = cv2.resize(arr, (max(1, int(W0 * scale)), max(1, int(H0 * scale)))) if scale < 1 else arr
        hog = cv2.HOGDescriptor()
        hog.setSVMDetector(cv2.HOGDescriptor_getDefaultPeopleDetector())
        found = hog.detectMultiScale(small, winStride=(8, 8), padding=(8, 8), scale=1.05)
        rects = list(found[0]) if (found is not None and len(found) and found[0] is not None) else []
        weights = found[1] if (found is not None and len(found) > 1) else None
        if not rects:
            return None
        if weights is not None and len(weights):
            keep = [r for r, w in zip(rects, weights) if float(w) >= 0.5]
            rects = keep or rects
        Hs, Ws = small.shape[:2]
        x0 = min(int(r[0]) for r in rects); y0 = min(int(r[1]) for r in rects)
        x1 = max(int(r[0]) + int(r[2]) for r in rects); y1 = max(int(r[1]) + int(r[3]) for r in rects)
        x0 = max(0, x0); y0 = max(0, y0); x1 = min(Ws, x1); y1 = min(Hs, y1)
        if x1 - x0 < 8 or y1 - y0 < 8:
            return None
        return (x0 / Ws, y0 / Hs, (x1 - x0) / Ws, (y1 - y0) / Hs)
    except Exception as e:
        print("    ! deteksi orang gagal:", e)
        return None

def _local_focus(im):
    """WAJAH -> ORANG -> area paling ramai detail. Catat metode yg kena ke _LAST_LOCAL."""
    global _LAST_LOCAL
    b = _face_box(im)
    if b: _LAST_LOCAL = "wajah"; return b
    b = _person_box(im)
    if b: _LAST_LOCAL = "orang"; return b
    b = _saliency_box(im)
    if b: _LAST_LOCAL = "saliency"; return b
    _LAST_LOCAL = None
    return None

def local_focus_box(im):
    """Fokus LOKAL (gratis): utamakan WAJAH; kalau nggak ada, pakai area paling ramai detail."""
    return _local_focus(im)

def smart_crop_to(im, w, h, context=""):
    """Crop ke (w,h), 3 lapis: 1) AI (konteks) kalau nyala & kuota; 2) WAJAH lokal (gratis);
    3) crop tengah. Balikin (image, metode 'ai'/'lokal'/None, rect_fraksi atau None)."""
    if SMART_CROP_ENABLED and context and GEMINI_API_KEY:
        tmp = None
        try:
            fd, tmp = tempfile.mkstemp(suffix=".jpg"); os.close(fd)
            im.convert("RGB").save(tmp, "JPEG", quality=88)
            box = ai_focus_box(tmp, context)
            if box:
                W, H = im.size; x, y, bw, bh = box; pad = 0.10
                bx0 = max(0, (x - pad) * W); by0 = max(0, (y - pad) * H)
                bx1 = min(W, (x + bw + pad) * W); by1 = min(H, (y + bh + pad) * H)
                if bx1 - bx0 >= 8 and by1 - by0 >= 8:
                    print(f"    [smart-crop] AI fokus box={tuple(round(v, 2) for v in box)}")
                    cropped, rect = _box_crop_wh(im, (bx0, by0, bx1, by1), w, h)
                    return cropped, "ai", rect
        except Exception as e:
            print("    ! smart-crop AI gagal:", e)
        finally:
            try:
                if tmp and os.path.exists(tmp): os.remove(tmp)
            except Exception:
                pass
    if SMART_CROP_ENABLED:
        try:
            fb = local_focus_box(im)
            if fb:
                W, H = im.size; x, y, bw, bh = fb; pad = 0.18
                bx0 = max(0, (x - pad) * W); by0 = max(0, (y - pad) * H)
                bx1 = min(W, (x + bw + pad) * W); by1 = min(H, (y + bh + pad) * H)
                if bx1 - bx0 >= 8 and by1 - by0 >= 8:
                    print("    [smart-crop] pakai fokus LOKAL (wajah/saliency)")
                    cropped, rect = _box_crop_wh(im, (bx0, by0, bx1, by1), w, h)
                    return cropped, (_LAST_LOCAL or "lokal"), rect
        except Exception as e:
            print("    ! crop lokal gagal:", e)
    # crop tengah biasa
    W, H = im.size; target = w / h
    if W / H > target:
        nw = int(H * target); x = (W - nw) // 2; im2 = im.crop((x, 0, x + nw, H))
        rect = (x / W, 0.0, (x + nw) / W, 1.0)
    else:
        nh = int(W / target); y = (H - nh) // 2; im2 = im.crop((0, y, W, y + nh))
        rect = (0.0, y / H, 1.0, (y + nh) / H)
    return im2.resize((w, h), Image.LANCZOS), None, rect

def make_bw_photo(img_bytes, out_path, clean_out=None, full_out=None, context=""):
    src = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    PH = int(CANVAS_H * PHOTO_FRAC)
    # FOTO UTUH B&W (nggak di-crop ke band) buat crop editable di Figma lewat HTML
    if full_out:
        try:
            fw = 1200; fh = max(1, int(fw * src.size[1] / max(1, src.size[0])))
            full_bw = ImageOps.autocontrast(ImageOps.grayscale(src.resize((fw, fh), Image.LANCZOS)), cutoff=1)
            full_bw.convert("RGB").save(full_out, "PNG")
        except Exception as _e:
            print("    ! simpan foto utuh gagal:", _e)
    # AUTO-CROP cerdas ke band (ganti crop tengah). rect = posisi crop (buat HTML editable).
    im, crop_method, crop_rect = smart_crop_to(src, CANVAS_W, PH, context)
    gray = ImageOps.autocontrast(ImageOps.grayscale(im), cutoff=1)
    try:
        grain = Image.effect_noise((CANVAS_W, PH), 22)
        gray = Image.blend(gray, grain, 0.08)
    except Exception:
        pass
    photo = gray.convert("RGB")
    # simpan versi BERSIH (tanpa gradasi) buat HTML -> masking jadi layer terpisah
    if clean_out:
        try:
            photo.save(clean_out, "PNG")
        except Exception as _e:
            print("    ! simpan foto bersih gagal:", _e)
    overlay = Image.new("RGBA", (CANVAS_W, PH), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    top_h = int(PH * 0.28)
    for y in range(0, top_h):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(110 * (1 - y / top_h))))
    photo = Image.alpha_composite(photo.convert("RGBA"), overlay).convert("RGB")
    photo.save(out_path, "PNG")
    return crop_rect, crop_method   # rect (crop editable Figma) + metode auto-crop ('ai'/'lokal'/None)


# ======================================================================
# 5. TEMA
# ======================================================================
def theme_colors(theme):
    if str(theme).strip().lower().startswith("hitam"):
        return "0E0E0E", "FFFFFF", "CFCFCF"
    return "FFFFFF", "141414", "4A4A4A"

def download_logo_path():
    if not LOGO_URL:
        return None
    try:
        data = requests.get(LOGO_URL, headers=UA, timeout=30).content
        path = os.path.join(tempfile.gettempdir(), f"logo_{int(time.time()*1000)}.png")
        with open(path, "wb") as fp:
            fp.write(data)
        Image.open(path).verify()
        return path
    except Exception as e:
        print("  ! logo gagal diambil:", e)
        return None


# ======================================================================
# 6. PPTX
# ======================================================================
def _add_text(slide, left, top, width, height, text, size_pt, bold, font, color_hex,
              align=PP_ALIGN.LEFT, highlight=False, italic=False):
    box = slide.shapes.add_textbox(fx(left), fy(top), fx(width), fy(height))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in text.split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        segs = parse_highlights(line) if highlight else [(line, False)]
        for seg, is_hl in (segs or [("", False)]):
            run = p.add_run()
            run.text = seg
            f = run.font
            f.size = Pt(size_pt); f.bold = bold; f.italic = italic; f.name = font
            f.color.rgb = ACCENT if is_hl else hex_rgb(color_hex)
    return box

def build_pptx(slides_data, out_path, logo_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W); prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]
    for s in slides_data:
        panel_bg, head_c, body_c = theme_colors(s["theme"])
        slide = prs.slides.add_slide(blank)
        panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        panel.fill.solid(); panel.fill.fore_color.rgb = hex_rgb(panel_bg)
        panel.line.fill.background(); panel.shadow.inherit = False
        slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=fy(PHOTO_FRAC))
        if logo_path:
            try:
                slide.shapes.add_picture(logo_path, fx(0.06), fy(0.045), height=fy(0.045))
            except Exception:
                pass
        _add_text(slide, 0.52, 0.045, 0.42, 0.08, HANDLE, 14, True, BODY_FONT, "FFFFFF", align=PP_ALIGN.RIGHT)
        fit = fit_text_layout(s.get("headline", ""), s.get("body", ""))
        hr = fit["head_px"] / max(1, head_px()); br = fit["body_px"] / max(1, body_px())
        if s["headline"]:
            _add_text(slide, 0.07, fit["head_top"], 0.86, 0.22, s["headline"],
                      max(8, int(round(head_pt() * hr))), G_HEAD_BOLD, HEADLINE_FONT, head_c, highlight=True)
        if s["body"]:
            _add_text(slide, 0.07, fit["body_top"], 0.86, 0.24, s["body"],
                      max(7, int(round(body_pt() * br))), False, BODY_FONT, body_c, highlight=True)
        _add_text(slide, 0.06, 0.935, 0.5, 0.05, HANDLE, 12, False, BODY_FONT, body_c)
        if s["total"] > 1:
            _add_text(slide, 0.80, 0.925, 0.14, 0.06, ARROW_TEXT, 24, True, BODY_FONT, head_c, align=PP_ALIGN.RIGHT)
    prs.save(out_path)
    return out_path


# ======================================================================
# 7. SVG
# ======================================================================
def _img_data_uri(path):
    im = Image.open(path).convert("RGB")
    buf = io.BytesIO(); im.save(buf, "JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()

def _svg_escape(t):
    return html.escape(t or "", quote=True)

def _wrap(text, max_chars):
    lines = []
    for para in (text or "").split("\n"):
        words = para.split()
        if not words:
            continue
        cur = ""
        for w in words:
            if len(cur) + len(w) + 1 <= max_chars:
                cur = (cur + " " + w).strip()
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
    return lines

def _svg_text(text, x, y, width, size, bold, font, color_hex, anchor="start", highlight=False, gap=1.3):
    # SATU elemen <text> saja (tiap baris = 1 <tspan> berposisi), biar di editor
    # kebaca sebagai satu blok teks — bukan satu layer per baris.
    weight = "700" if bold else "400"
    max_chars = max(6, int(width / (size * 0.55)))
    def words_of(line):
        res = []
        for seg, is_hl in (parse_highlights(line) if highlight else [(line, False)]):
            for w in seg.split(" "):
                if w != "":
                    res.append((w, is_hl))
        return res
    line_tspans = []
    yy = y
    for src in text.split("\n"):
        ws = words_of(src)
        if not ws:
            yy += int(size * gap); continue
        vlines, cur, cur_len = [], [], 0
        for w, is_hl in ws:
            add = len(w) + (1 if cur else 0)
            if cur and cur_len + add > max_chars:
                vlines.append(cur); cur, cur_len = [], 0; add = len(w)
            cur.append((w, is_hl)); cur_len += add
        if cur:
            vlines.append(cur)
        for vl in vlines:
            base = yy + size
            frag = ""
            for k, (w, is_hl) in enumerate(vl):
                prefix = " " if k > 0 else ""
                if is_hl:   # kata di-highlight -> warna aksen (emas)
                    frag += f'<tspan fill="#{ACCENT_COLOR}">{_svg_escape(prefix + w)}</tspan>'
                else:
                    frag += _svg_escape(prefix + w)
            line_tspans.append(f'<tspan x="{x}" y="{int(base)}" xml:space="preserve">{frag}</tspan>')
            yy += int(size * gap)
    if not line_tspans:
        return ""
    return (f'<text text-anchor="{anchor}" fill="#{color_hex}" '
            f'font-family="{_svg_escape(font)}, Georgia, serif" '
            f'font-size="{size}" font-weight="{weight}">' + "".join(line_tspans) + '</text>')

def build_svg(slides_data, out_path, logo_path):
    total = len(slides_data)
    W = CANVAS_W * total
    H = CANVAS_H
    PH = int(CANVAS_H * PHOTO_FRAC)
    logo_uri = None
    if logo_path:
        try:
            im = Image.open(logo_path).convert("RGBA")
            b = io.BytesIO(); im.save(b, "PNG")
            logo_uri = "data:image/png;base64," + base64.b64encode(b.getvalue()).decode()
        except Exception:
            pass
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">']
    for idx, s in enumerate(slides_data):
        xo = idx * CANVAS_W
        panel_bg, head_c, body_c = theme_colors(s["theme"])
        parts.append(f'<rect x="{xo}" y="0" width="{CANVAS_W}" height="{H}" fill="#{panel_bg}"/>')
        parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{PH}" '
                     f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(s["image_path"])}"/>')
        if logo_uri:
            parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*CANVAS_H)}" '
                         f'height="{int(0.05*CANVAS_H)}" href="{logo_uri}"/>')
        parts.append(_svg_text(HANDLE, xo + int(0.94*CANVAS_W), int(0.05*CANVAS_H),
                               int(0.4*CANVAS_W), 24, True, BODY_FONT, "FFFFFF", anchor="end"))
        fit = fit_text_layout(s.get("headline", ""), s.get("body", ""))
        if s["headline"]:
            parts.append(_svg_text(s["headline"], xo + int(0.07*CANVAS_W), int(fit["head_top"]*CANVAS_H),
                                   int(0.86*CANVAS_W), fit["head_px"], G_HEAD_BOLD, HEADLINE_FONT, head_c, highlight=True))
        if s["body"]:
            parts.append(_svg_text(s["body"], xo + int(0.07*CANVAS_W), int(fit["body_top"]*CANVAS_H),
                                   int(0.86*CANVAS_W), fit["body_px"], False, BODY_FONT, body_c, highlight=True))
        parts.append(_svg_text(HANDLE, xo + int(0.06*CANVAS_W), int(0.93*CANVAS_H),
                               int(0.4*CANVAS_W), 22, False, BODY_FONT, body_c))
        if total > 1:
            parts.append(_svg_text(ARROW_TEXT, xo + int(0.93*CANVAS_W), int(0.925*CANVAS_H),
                                   int(0.1*CANVAS_W), 40, True, BODY_FONT, head_c, anchor="end"))
    parts.append('</svg>')
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return out_path


# ======================================================================
# 7b. HTML  (buat diedit lagi di Figma via plugin html.to.design -> text box asli)
# ======================================================================
def _html_rich(text):
    """Ubah teks ber-marker ==emas== jadi HTML. Baris baru -> <br>. Aman di-escape."""
    out = []
    for i, line in enumerate(text.split("\n")):
        if i > 0:
            out.append("<br>")
        for seg, is_hl in parse_highlights(line):
            esc = html.escape(seg, quote=True)
            if is_hl:
                out.append(f'<span class="hl">{esc}</span>')
            else:
                out.append(esc)
    return "".join(out)

def build_html(slides_data, out_path, logo_path):
    total = len(slides_data)
    PH = int(CANVAS_H * PHOTO_FRAC)
    ty = int((PHOTO_FRAC + 0.05) * CANVAS_H)
    body_ty = ty + int(0.20 * CANVAS_H)
    logo_uri = None
    if logo_path:
        try:
            im = Image.open(logo_path).convert("RGBA")
            b = io.BytesIO(); im.save(b, "PNG")
            logo_uri = "data:image/png;base64," + base64.b64encode(b.getvalue()).decode()
        except Exception:
            pass

    css = f"""
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    body {{ background:#555; font-family:'Vremena Grotesk','Lora',Georgia,serif; }}
    .wrap {{ display:flex; flex-direction:column; align-items:center; gap:28px; padding:28px; }}
    .slide {{ position:relative; width:{CANVAS_W}px; height:{CANVAS_H}px; overflow:hidden; }}
    /* FOTO (bersih) dan MASKING dipisah jadi layer sendiri-sendiri */
    .photo {{ position:absolute; top:0; left:0; width:{CANVAS_W}px; height:{PH}px; overflow:hidden; }}
    .photo .bg {{ width:100%; height:100%; object-fit:cover; display:block; filter:grayscale(1); }}
    .photo .ov {{ position:absolute; top:0; left:0; width:100%; height:28%;
                  background:linear-gradient(to bottom, rgba(0,0,0,.43), rgba(0,0,0,0)); }}
    .logo {{ position:absolute; left:{int(0.06*CANVAS_W)}px; top:{int(0.05*CANVAS_H)}px;
             height:{int(0.05*CANVAS_H)}px; }}
    .handle-top {{ position:absolute; right:{int(0.06*CANVAS_W)}px; top:{int(0.05*CANVAS_H)}px;
                   color:#fff; font-weight:700; font-size:24px; }}
    .headline {{ position:absolute; left:{int(0.07*CANVAS_W)}px; width:{int(0.86*CANVAS_W)}px; line-height:1.3; }}
    .body {{ position:absolute; left:{int(0.07*CANVAS_W)}px; width:{int(0.86*CANVAS_W)}px;
             font-weight:400; line-height:1.4; }}
    .hl {{ color:#{ACCENT_COLOR}; }}
    .handle-bottom {{ position:absolute; left:{int(0.06*CANVAS_W)}px; top:{int(0.92*CANVAS_H)}px;
                      font-size:22px; }}
    .arrow {{ position:absolute; right:{int(0.06*CANVAS_W)}px; top:{int(0.9*CANVAS_H)}px;
              font-weight:700; font-size:40px; }}
    """

    body_parts = ['<div class="wrap">']
    for s in slides_data:
        panel_bg, head_c, body_c = theme_colors(s["theme"])
        full_p = s.get("full_path"); rect = s.get("crop_rect")
        body_parts.append(f'<div class="slide" style="background:#{panel_bg}">')
        body_parts.append('<div class="photo">')
        if full_p and rect:
            # FOTO UTUH + crop sbg frame -> tampil sama kaya hasil auto-crop, tapi di Figma bisa di-recrop
            x0, y0, x1, y1 = rect
            cw = max(1e-3, x1 - x0); ch = max(1e-3, y1 - y0)
            st = (f"position:absolute;left:{-x0/cw*100:.3f}%;top:{-y0/ch*100:.3f}%;"
                  f"width:{100/cw:.3f}%;height:{100/ch:.3f}%;max-width:none;object-fit:fill;filter:grayscale(1);")
            body_parts.append(f'<img src="{_img_data_uri(full_p)}" alt="" style="{st}">')
        else:
            bgp = full_p or s.get("clean_path") or s.get("image_path")
            body_parts.append(f'<img class="bg" src="{_img_data_uri(bgp)}" alt="">')
        body_parts.append('<div class="ov"></div>')   # masking terpisah dari foto
        body_parts.append('</div>')
        if logo_uri:
            body_parts.append(f'<img class="logo" src="{logo_uri}" alt="">')
        body_parts.append(f'<div class="handle-top">{html.escape(HANDLE)}</div>')
        fit = fit_text_layout(s.get("headline", ""), s.get("body", ""))
        hweight = 700 if G_HEAD_BOLD else 400
        if s["headline"]:
            body_parts.append(f'<div class="headline" style="color:#{head_c};top:{int(fit["head_top"]*CANVAS_H)}px;'
                              f'font-size:{fit["head_px"]}px;font-weight:{hweight}">{_html_rich(s["headline"])}</div>')
        if s["body"]:
            body_parts.append(f'<div class="body" style="color:#{body_c};top:{int(fit["body_top"]*CANVAS_H)}px;'
                              f'font-size:{fit["body_px"]}px">{_html_rich(s["body"])}</div>')
        body_parts.append(f'<div class="handle-bottom" style="color:#{body_c}">{html.escape(HANDLE)}</div>')
        if total > 1:
            body_parts.append(f'<div class="arrow" style="color:#{head_c}">{html.escape(ARROW_TEXT)}</div>')
        body_parts.append('</div>')
    body_parts.append('</div>')

    doc = (
        "<!doctype html><html lang=\"id\"><head><meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
        "<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">"
        "<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>"
        "<link href=\"https://fonts.googleapis.com/css2?family=Lora:wght@400;500;600;700&display=swap\" rel=\"stylesheet\">"
        f"<title>{html.escape((slides_data[0]['headline'] if slides_data else 'Desain') or 'Desain')}</title>"
        f"<style>{css}</style></head><body>" + "".join(body_parts) + "</body></html>"
    )
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)
    return out_path


# ======================================================================
# 8. TELEGRAM
# ======================================================================
TG_BASE = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

def _tg_call(method, data, files=None, label=""):
    try:
        r = requests.post(f"{TG_BASE}/{method}", data=data, files=files, timeout=300)
        try:
            j = r.json()
        except Exception:
            j = {}
        time.sleep(0.5)
        if not r.ok or not j.get("ok", False):
            print(f"    ! Telegram {method} {label} GAGAL (HTTP {r.status_code}): {r.text[:300]}")
            return False
        return True
    except Exception as e:
        print(f"    ! Telegram {method} {label} ERROR: {e}")
        return False

def tg_message(text):
    return _tg_call("sendMessage", {"chat_id": TELEGRAM_CHAT_ID, "text": text[:4000]}, label="msg")

def tg_photo(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendPhoto", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]}, {"photo": f}, label="photo")

def tg_document(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendDocument", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]}, {"document": f}, label="doc")

def tg_buttons(text, keyboard):
    """Kirim pesan + tombol inline. keyboard = list of rows, tiap tombol (label, callback_data)."""
    kb = {"inline_keyboard": [[{"text": t, "callback_data": d} for (t, d) in row] for row in keyboard]}
    return _tg_call("sendMessage", {"chat_id": TELEGRAM_CHAT_ID, "text": text[:3500],
                                    "reply_markup": json.dumps(kb)}, label="btn")

def tg_delete_webhook():
    """Pastikan nggak ada webhook nyangkut. Webhook bikin getUpdates MANDEK (balik kosong),
    jadi tombol/ketikan user nggak pernah kebaca walau bot tetap bisa KIRIM pesan.
    drop_pending_updates=false -> pesan yg ketahan TIDAK dibuang (langsung kebaca sesudahnya)."""
    try:
        r = requests.post(f"{TG_BASE}/deleteWebhook",
                          data={"drop_pending_updates": "false"}, timeout=20)
        j = r.json()
        print(f"    deleteWebhook: ok={j.get('ok')} {j.get('description','')}")
    except Exception as e:
        print("    ! deleteWebhook error:", e)

def tg_get_updates(offset):
    try:
        r = requests.get(f"{TG_BASE}/getUpdates",
                         params={"offset": offset, "timeout": 0, "allowed_updates": json.dumps(["callback_query", "message"])},
                         timeout=40)
        j = r.json()
        if not j.get("ok"):
            print(f"    ! getUpdates NOT ok: {str(j)[:200]}")
        return j.get("result", []) if j.get("ok") else []
    except Exception as e:
        print("    ! getUpdates error:", e)
        return []

def tg_answer_callback(cb_id, text=""):
    # notif kecil di tombol; kalau gagal (query kedaluwarsa krn batch) abaikan diam-diam
    try:
        requests.post(f"{TG_BASE}/answerCallbackQuery",
                      data={"callback_query_id": cb_id, "text": text[:180]}, timeout=20)
    except Exception:
        pass

def tg_get_file_path(file_id):
    """Ambil path internal file Telegram dari file_id (buat download foto yg dikirim user)."""
    try:
        r = requests.get(f"{TG_BASE}/getFile", params={"file_id": file_id}, timeout=30)
        j = r.json()
        return j["result"]["file_path"] if j.get("ok") else None
    except Exception as e:
        print("    ! getFile error:", e); return None

def tg_download_file(file_path, out_path):
    """Download file dari server Telegram ke lokal."""
    try:
        url = f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"
        r = requests.get(url, timeout=60); r.raise_for_status()
        with open(out_path, "wb") as f:
            f.write(r.content)
        return out_path
    except Exception as e:
        print("    ! download file error:", e); return None


# ======================================================================
# 8b. SISTEM BELAJAR — memori + feedback (approve/reject + alasan)
# ======================================================================
DEFAULT_MEMORY = {
    "preferences": {"head_scale": 1.0, "body_scale": 1.0, "scrim": 1.0, "soft": 0.0, "head_bold": True},
    "photo_history": {},  # id_foto (str) -> unix_ts terakhir dipakai (anti pakai-ulang < 6 bulan)
    "pending": {},        # token -> {title, page_id, ts, revision}
    "resolved": [],       # token yang sudah dinilai (anti dobel)
    "awaiting_text": None,  # token yang menunggu alasan ketik
    "revise_queue": {},   # page_id -> {title, revision, reasons[]} konten yg minta direvisi
    "log": [],            # riwayat {ts, title, status, reason}
    "stats": {"approved": 0, "rejected": 0, "weekly": []},
    "tg_offset": 0,
    "design_cache": {},   # page_id_norm -> {hash, n, title, slides:{i:{...}}, ts}
}

MAX_REVISI = int(os.environ.get("MAX_REVISI") or 10)   # batas aman auto-revisi per konten

# tombol alasan saat Reject: (label, kode)
REASON_BUTTONS = [
    ("Teks kegedean", "tbig"),
    ("Teks kekecilan", "tsmall"),
    ("Teks susah kebaca", "hard"),
    ("Overlay kegelapan", "dark"),
    ("Gambar kurang aesthetic", "img"),
]
REASON_LABEL = {k: v for v, k in REASON_BUTTONS}

def load_memory():
    try:
        with open(MEMORY_PATH, "r", encoding="utf-8") as f:
            mem = json.load(f)
        for k, v in DEFAULT_MEMORY.items():
            mem.setdefault(k, v if not isinstance(v, (dict, list)) else (dict(v) if isinstance(v, dict) else list(v)))
        for k, v in DEFAULT_MEMORY["preferences"].items():
            mem["preferences"].setdefault(k, v)
        return mem
    except Exception:
        import copy
        return copy.deepcopy(DEFAULT_MEMORY)

def save_memory(mem):
    try:
        with open(MEMORY_PATH, "w", encoding="utf-8") as f:
            json.dump(mem, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("    ! gagal simpan memori:", e)

def apply_reason(prefs, code):
    """Geser knob sesuai alasan. Mengembalikan deskripsi singkat perubahan."""
    if code == "tbig":
        prefs["head_scale"] = _clamp(prefs["head_scale"] - 0.06, 0.7, 1.3)
        prefs["body_scale"] = _clamp(prefs["body_scale"] - 0.05, 0.7, 1.3)
        return "teks dikecilin"
    if code == "tsmall":
        prefs["head_scale"] = _clamp(prefs["head_scale"] + 0.06, 0.7, 1.3)
        prefs["body_scale"] = _clamp(prefs["body_scale"] + 0.05, 0.7, 1.3)
        return "teks digedein"
    if code == "hard":
        prefs["scrim"] = _clamp(prefs["scrim"] + 0.12, 0.6, 1.6)
        return "overlay dipergelap (biar teks kebaca)"
    if code == "dark":
        prefs["scrim"] = _clamp(prefs["scrim"] - 0.12, 0.6, 1.6)
        return "overlay dipertipis"
    if code == "img":
        prefs["soft"] = _clamp(prefs["soft"] + 0.08, 0.0, 0.4)
        return "gambar dibikin lebih soft/film"
    return "dicatat (tanpa ubah setelan)"

def prefs_summary(prefs):
    return (f"Setelan belajar sekarang: judul {int(prefs['head_scale']*100)}%, "
            f"body {int(prefs['body_scale']*100)}%, overlay {int(prefs['scrim']*100)}%, "
            f"soft +{int(prefs['soft']*100)}%, judul {'tebal' if prefs.get('head_bold', True) else 'tipis'}.")


# ---------- OTAK AI: Groq (utama) + Gemini (cadangan). Dua-duanya gratis ----------
GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL    = os.environ.get("GEMINI_MODEL") or "gemini-3.5-flash"
# model cadangan buat 'mata' (vision) kalau model utama lagi overload (503)
GEMINI_VISION_FALLBACKS = [m.strip() for m in (os.environ.get("GEMINI_VISION_FALLBACKS") or "gemini-3.5-flash,gemini-3.6-flash").split(",") if m.strip()]
GROQ_API_KEY    = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL      = os.environ.get("GROQ_MODEL") or "openai/gpt-oss-20b"
AI_CHAT_ENABLED = ((os.environ.get("AI_CHAT_ENABLED") or "true").lower() == "true"
                   and (bool(GEMINI_API_KEY) or bool(GROQ_API_KEY)))
# cek mandiri pakai 'mata' AI (Gemini vision). gratis (gambar sbg INPUT), cuma butuh Gemini key.
VISION_CHECK_ENABLED = ((os.environ.get("VISION_CHECK_ENABLED") or "true").lower() == "true"
                        and bool(GEMINI_API_KEY))

def _gemini_json(prompt):
    if not GEMINI_API_KEY:
        return None
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {"contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.3, "responseMimeType": "application/json"}}
    try:
        r = requests.post(url, json=body, timeout=45)
        if r.status_code != 200:
            print(f"    ! Gemini {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
    except Exception as e:
        print("    ! Gemini gagal:", e); return None

def _groq_json(prompt):
    if not GROQ_API_KEY:
        return None
    try:
        r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                          headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                          json={"model": GROQ_MODEL, "temperature": 0.3,
                                "response_format": {"type": "json_object"},
                                "messages": [{"role": "user", "content": prompt}]},
                          timeout=45)
        if r.status_code != 200:
            print(f"    ! Groq {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["choices"][0]["message"]["content"])
    except Exception as e:
        print("    ! Groq gagal:", e); return None

AI_PRIMARY = (os.environ.get("AI_PRIMARY") or "groq").lower()   # 'groq' (stabil) atau 'gemini'

def ai_json(prompt):
    """Minta jawaban JSON ke AI. Default: Groq dulu (lebih stabil), Gemini cadangan.
    Bisa dibalik lewat secret AI_PRIMARY=gemini."""
    if not AI_CHAT_ENABLED:
        return None
    if AI_PRIMARY == "gemini":
        first, second, second_name = _gemini_json, _groq_json, "Groq"
    else:
        first, second, second_name = _groq_json, _gemini_json, "Gemini"
    data = first(prompt)
    if data is None:
        data = second(prompt)   # cadangan kalau yang utama ngadat
        if data is not None:
            print(f"    (pakai {second_name} sebagai cadangan)")
    return data

AI_PROMPT = """Kamu asisten desain untuk bot konten Instagram (brand Ala NU — konseling spiritual Nahdiyyin).
User memberi perintah/feedback (bahasa Indonesia atau Inggris) untuk merevisi sebuah desain.
Terjemahkan jadi penyesuaian angka. Setelan sekarang (1.0 = normal):
- head_scale: ukuran JUDUL (batas 0.7-1.3)
- body_scale: ukuran teks BODY (batas 0.7-1.3)
- scrim: kegelapan overlay di belakang teks (batas 0.6-1.6; naik = lebih gelap = teks lebih terbaca)
- soft: kesan film/lembut pada gambar (batas 0.0-0.4; naik = lebih lembut/pudar; turun = lebih cerah/tajam)
- head_bold: JUDUL tebal/bold (true) atau tipis/normal (false). Isi true/false HANYA kalau user
  nyinggung ketebalan judul ("judul jangan bold" -> false; "tebelin judulnya" -> true). Kalau nggak -> null.
PENTING: hampir SEMUA masukan user = permintaan revisi -> set action "revise".
Keluhan/komentar sekecil apapun soal desain (foto/teks/warna/overlay) = "revise".
Contoh yang HARUS "revise": "gambar kurang bagus", "kurang aesthetic", "ganti foto",
"fotonya jelek", "judul kegedean", "kurang jelas", "overlay kurang gelap".
Kalau user komentarin/minta ganti foto, set action "revise" (foto otomatis diganti baru saat revisi).
Pakai action "none" HANYA kalau pesan jelas cuma sapaan/basa-basi/terima kasih/pertanyaan umum
(mis. "halo", "makasih", "lagi apa") — BUKAN komentar soal desain.
PENTING soal FOTO: set "new_photo": true HANYA kalau user minta ganti/komentar jelek soal FOTO/GAMBAR
(mis. "ganti foto", "fotonya jelek", "gambar kurang aesthetic", "cari gambar lain").
Kalau user cuma soal TEKS/judul/overlay/warna (bukan foto), "new_photo": false (foto lama dipakai lagi).
PENTING soal SLIDE: kalau user nyebut nomor slide (mis. "slide 2", "slide 1 dan 3"),
isi "slides" dengan daftar nomor itu. Kalau nggak nyebut nomor, "slides": [] (artinya semua slide).
Balas HANYA JSON valid, tanpa teks lain:
{"head_scale_delta": <angka -0.12..0.12>, "body_scale_delta": <angka>, "scrim_delta": <angka>, "soft_delta": <angka>, "head_bold": true/false/null, "note": "<ringkasan singkat dalam bahasa Indonesia, maks 12 kata>", "action": "revise" atau "none", "new_photo": true/false, "slides": [<nomor slide>]}
Pesan user: "%s" """

def ai_parse_command(text):
    """Ubah kalimat bebas jadi penyesuaian angka (Groq/Gemini). None kalau gagal/mati."""
    return ai_json(AI_PROMPT % text[:400])

def apply_ai(prefs, data):
    """Terapkan delta dari AI (dibatasi aman). Kembalikan daftar yg berubah."""
    changed = []
    plan = [("head_scale_delta", "head_scale", 0.7, 1.3),
            ("body_scale_delta", "body_scale", 0.7, 1.3),
            ("scrim_delta", "scrim", 0.6, 1.6),
            ("soft_delta", "soft", 0.0, 0.4)]
    for key, knob, lo, hi in plan:
        try:
            d = float(data.get(key, 0) or 0)
        except Exception:
            d = 0.0
        d = _clamp(d, -0.15, 0.15)   # batasi per-perintah biar nggak loncat jauh
        if abs(d) >= 0.01:
            new = _clamp(prefs[knob] + d, lo, hi)
            if abs(new - float(prefs[knob])) >= 1e-6:   # cuma catat kalau beneran berubah
                prefs[knob] = new
                changed.append(knob)
    hb = data.get("head_bold", None)
    if isinstance(hb, str):
        hb = {"true": True, "false": False, "ya": True, "iya": True, "tidak": False,
              "no": False, "yes": True, "1": True, "0": False}.get(hb.strip().lower(), None)
    if isinstance(hb, bool) and bool(prefs.get("head_bold", True)) != hb:
        prefs["head_bold"] = hb
        changed.append("head_bold")
    return changed

def local_parse_command(text):
    """Penerjemah revisi LOKAL (tanpa AI) -> ketikan SELALU ngefek walau AI mati."""
    low = (text or "").lower()
    def has(*kw): return any(k in low for k in kw)
    d = {"head_scale_delta": 0.0, "body_scale_delta": 0.0, "scrim_delta": 0.0,
         "soft_delta": 0.0, "head_bold": None, "new_photo": False,
         "slides": _slide_nums_in(text), "note": "", "action": "none"}
    notes = []
    hit_judul = has("judul", "headline", "title", "tulisan besar",
                    "atas", "yang atas", "teks atas", "bagian atas", "tulisan atas")
    hit_body  = has("body", "isi", "caption", "deskripsi", "paragraf", "teks kecil",
                    "bawah", "yang bawah", "teks bawah", "bagian bawah", "tulisan bawah")
    both = (not hit_judul and not hit_body)
    if has("kegedean", "kegede", "kebesaran", "gede banget", "terlalu besar", "terlalu gede", "kecilin", "perkecil", "kecilkan"):
        if hit_judul or both: d["head_scale_delta"] -= 0.12
        if hit_body  or both: d["body_scale_delta"] -= 0.10
        notes.append("teks dikecilin")
    if has("kekecilan", "kecil banget", "terlalu kecil", "gedein", "perbesar", "besarin", "lebih besar", "lebih gede",
           "kurang besar", "kurang gede", "kurang gedhe", "krg besar"):
        if hit_judul or both: d["head_scale_delta"] += 0.12
        if hit_body  or both: d["body_scale_delta"] += 0.10
        notes.append("teks digedein")
    # "samain / ukuran sama" -> dekatkan: judul diturunin, body dinaikin (ulang biar makin dekat)
    if has("samain", "ukuran sama", "ukurannya sama", "sama besar", "sama gede", "seragam", "disamain", "biar sama", "samakan"):
        d["head_scale_delta"] -= 0.10
        d["body_scale_delta"] += 0.10
        notes.append("ukuran judul & body didekatin")
    if has("lebih soft", "lebih lembut", "lebih pudar", "lembutin", "vintage"):
        d["soft_delta"] += 0.10; notes.append("foto dibikin lembut")
    if has("lebih cerah", "lebih terang", "lebih tajam", "cerahin", "terangin", "tajamin", "jangan pudar",
           "kurang terang", "kurang cerah", "kurang nyala"):
        d["soft_delta"] -= 0.10; notes.append("foto dibikin cerah")
    if has("jangan bold", "ga bold", "gak bold", "nggak bold", "tanpa bold", "jangan tebal", "jangan tebel",
           "nggak tebal", "tipisin", "tipis", "jangan ada yang bold", "nggak ada yang bold"):
        d["head_bold"] = False; notes.append("judul dibikin tipis")
    elif has("dibold", "di bold", "pakai bold", "tebelin", "tebalin", "dipertebal", "pertebal", "judul tebal", "bold lagi",
             "kurang tebal", "kurang tebel", "kurang bold"):
        d["head_bold"] = True; notes.append("judul dibikin tebal")
    if has("ganti foto", "ganti gambar", "foto baru", "gambar baru", "foto jelek", "gambar jelek",
           "foto kurang", "gambar kurang", "foto nggak", "gambar nggak", "aesthetic", "estetik",
           "foto lain", "gambar lain", "fotonya jelek", "gambarnya jelek"):
        d["new_photo"] = True; notes.append("foto diganti baru")
    _rev_kata = ("ganti", "ubah", "revisi", "perbaiki", "kurang", "jelek", "buram", "norak",
                 "kegedean", "kekecilan", "lebih ", "jangan", "terlalu", "bold", "tebal",
                 "tipis", "tebel", "numpuk", "tumpang", "ketimpa", "nabrak")
    any_change = (abs(d["head_scale_delta"]) + abs(d["body_scale_delta"]) + abs(d["scrim_delta"])
                  + abs(d["soft_delta"]) >= 0.01) or (d["head_bold"] is not None) or d["new_photo"]
    if any_change or has(*_rev_kata):
        d["action"] = "revise"
    d["note"] = (", ".join(notes) or (text or "")[:60])
    return d

def _merge_cmd(ai_data, local):
    """Gabung AI + lokal. Lokal nambal yg AI lewatin -> ketikan selalu ngefek."""
    if not ai_data:
        return local
    for k in ("head_scale_delta", "body_scale_delta", "scrim_delta", "soft_delta"):
        try:
            if abs(float(local.get(k, 0) or 0)) >= 0.01:
                ai_data[k] = local[k]          # lokal override (lebih bisa dipercaya drpd AI)
        except Exception:
            pass
    local_size = (abs(local.get("head_scale_delta", 0) or 0) >= 0.01
                  or abs(local.get("body_scale_delta", 0) or 0) >= 0.01)
    if local_size:
        if abs(local.get("head_scale_delta", 0) or 0) < 0.01:
            ai_data["head_scale_delta"] = 0.0
        if abs(local.get("body_scale_delta", 0) or 0) < 0.01:
            ai_data["body_scale_delta"] = 0.0
    if local.get("head_bold", None) is not None:
        ai_data["head_bold"] = local["head_bold"]
    if not ai_data.get("new_photo") and local.get("new_photo"):
        ai_data["new_photo"] = True
    if not ai_data.get("slides") and local.get("slides"):
        ai_data["slides"] = local["slides"]
    if local.get("action") == "revise" and ai_data.get("action") != "revise":
        ai_data["action"] = "revise"
    return ai_data

def apply_example(prefs, data):
    """Terapkan setelan HASIL BACA foto contoh (nilai ABSOLUT, bukan delta)."""
    changed = []
    plan = [("head_scale", 0.7, 1.3), ("body_scale", 0.7, 1.3),
            ("scrim", 0.6, 1.6), ("soft", 0.0, 0.4)]
    for knob, lo, hi in plan:
        v = (data or {}).get(knob)
        if v is None:
            continue
        try:
            v = _clamp(float(v), lo, hi)
        except Exception:
            continue
        if abs(v - float(prefs.get(knob, 1.0))) >= 0.02:
            prefs[knob] = v
            changed.append(knob)
    return changed

def _latest_pending(mem):
    best = None
    for tok, rec in mem.get("pending", {}).items():
        if best is None or rec.get("ts", 0) > best[1].get("ts", 0):
            best = (tok, rec)
    return best

def ai_image_keywords(slide_texts):
    """1 panggilan AI (Groq/Gemini) -> kata kunci FOTO (Inggris, aesthetic) per slide. None kalau gagal."""
    if not AI_CHAT_ENABLED or not slide_texts:
        return None
    joined = "\n".join(f"{i+1}. {(t or '')[:160]}" for i, t in enumerate(slide_texts))
    prompt = (
        "Kamu art director untuk brand spiritual Nahdiyyin (Ala NU) — gaya sinematik, moody, "
        "dramatic light, foto hitam-putih, nuansa tenang & khusyuk.\n"
        "Untuk TIAP slide di bawah, buat 1 kata kunci pencarian FOTO STOK dalam BAHASA INGGRIS (3-5 kata).\n"
        "WAJIB: setiap kata kunci bernuansa ISLAMI/MUSLIM (mis. 'muslim', 'hijab', 'mosque', 'praying', "
        "'quran', 'ramadan') DAN tetap nyambung sama isi/suasana slide-nya. "
        "Contoh: konten soal sabar -> 'muslim man praying patience'; soal keluarga -> 'muslim family warmth'.\n"
        "Fokus ke suasana/visual (bukan terjemahan harfiah), hindari tulisan/logo, utamakan natural light & tone lembut.\n"
        'Balas HANYA JSON object: {"keywords": ["...", "..."]} dengan panjang array PERSIS sama dengan jumlah slide.\n\n' + joined)
    data = ai_json(prompt)
    try:
        arr = (data or {}).get("keywords")
        if isinstance(arr, list) and arr:
            return [str(x) for x in arr]
    except Exception:
        pass
    return None

def _extract_json(resp):
    """Ambil JSON dari balasan Gemini walau dibungkus teks/markdown (```json ... ```)."""
    try:
        raw = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
    except Exception:
        return None
    raw = (raw or "").strip()
    if raw.startswith("```"):
        raw = raw.strip("`")
        if raw[:4].lower() == "json":
            raw = raw[4:]
        raw = raw.strip()
    try:
        return json.loads(raw)
    except Exception:
        pass
    import re as _re
    mobj = _re.search(r"\{.*\}", raw, _re.S)
    if mobj:
        try:
            return json.loads(mobj.group(0))
        except Exception:
            return None
    return None

def _gemini_vision_json(prompt, image_path):
    """Kirim gambar + pertanyaan ke Gemini (mode 'mata'/vision), minta jawaban JSON.
    Gambar dikecilin dulu (max 640px, JPEG) biar payload kecil & cepat."""
    if not GEMINI_API_KEY:
        return None
    try:
        im = Image.open(image_path).convert("RGB")
        im.thumbnail((640, 640))
        buf = io.BytesIO(); im.save(buf, "JPEG", quality=80)
        b64 = base64.b64encode(buf.getvalue()).decode()
    except Exception as e:
        print("    ! vision: gagal siapin gambar:", e); return None
    body = {"contents": [{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}],
            "generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"}}
    # coba model utama, lalu cadangan; tiap model diulang kalau overload sesaat (5xx/429)
    tried = []
    for mdl in [GEMINI_MODEL] + [mm for mm in GEMINI_VISION_FALLBACKS if mm != GEMINI_MODEL]:
        tried.append(mdl)
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{mdl}:generateContent?key={GEMINI_API_KEY}"
        for attempt in range(3):
            try:
                r = requests.post(url, json=body, timeout=60)
            except Exception as e:
                print(f"    ! Gemini vision {mdl} error:", e); break
            if r.status_code == 200:
                parsed = _extract_json(r)
                if parsed is not None:
                    return parsed
                print(f"    ! Gemini vision {mdl}: 200 tapi JSON gagal diparse"); break
            if r.status_code in (429, 500, 502, 503, 504):
                print(f"    ! Gemini vision {mdl} {r.status_code} (overload, coba lagi {attempt+1}/3)")
                time.sleep(1.5 * (attempt + 1)); continue
            print(f"    ! Gemini vision {mdl} {r.status_code}: {r.text[:100]}"); break
    print("    ! Gemini vision nyerah; model dicoba:", tried)
    return None

def ai_vision_check(image_path, zone, theme):
    """Cek mandiri: kirim foto COVER slide ke Gemini, tanya apakah teks bakal mudah dibaca.
    Kembalikan {ok: bool, issue: str} atau None kalau mati/gagal."""
    if not VISION_CHECK_ENABLED:
        return None
    zmap = {"top": "atas", "bottom": "bawah", "left": "kiri", "right": "kanan", "center": "tengah"}
    warna = "putih" if theme == "dark" else "gelap (hampir hitam)"
    prompt = (
        "Kamu reviewer desain konten Instagram. Ini FOTO cover sebuah slide (foto hitam-putih).\n"
        f"Rencana: handle/logo + teks diletakkan di area {zmap.get(zone, zone)}, warna teks {warna}.\n"
        "Nilai apakah di area itu teks akan MUDAH DIBACA: kontras cukup, area tidak terlalu ramai/berpola, "
        "dan tidak menutupi wajah/objek penting.\n"
        'Balas HANYA JSON: {"ok": true atau false, "issue": "<masalah singkat bahasa Indonesia maks 12 kata; '
        'kosongkan kalau sudah ok>"}'
    )
    data = _gemini_vision_json(prompt, image_path)   # 1x aja (retry malah boros kuota)
    if not isinstance(data, dict):
        return None
    return {"ok": bool(data.get("ok", True)), "issue": str(data.get("issue") or "").strip()}

def ai_learn_from_example(image_path):
    """Baca FOTO CONTOH yg dikirim user -> tangkap 'vibe'-nya jadi setelan (knob).
    Kembalikan {head_scale, body_scale, scrim, soft, note} atau None kalau gagal."""
    if not GEMINI_API_KEY:
        return None
    prompt = (
        "Kamu art director. Ini FOTO CONTOH desain konten Instagram yang DISUKAI user; "
        "dia mau desain berikutnya bergaya mirip ini.\n"
        "Bandingkan dengan desain 'normal' (semua nilai 1.0 = normal), lalu perkirakan setelannya dari yg kamu LIHAT:\n"
        "- head_scale: ukuran JUDUL (0.7 kecil .. 1.3 besar)\n"
        "- body_scale: ukuran teks ISI (0.7 .. 1.3)\n"
        "- scrim: kegelapan lapisan gelap di belakang teks (0.6 terang .. 1.6 gelap pekat)\n"
        "- soft: kesan lembut/pudar foto (0.0 tajam/kontras .. 0.4 lembut/film)\n"
        "Kasih angka sewajarnya (jangan semua 1.0 kalau memang keliatan beda).\n"
        'Balas HANYA JSON: {"head_scale": <angka>, "body_scale": <angka>, "scrim": <angka>, '
        '"soft": <angka>, "note": "<ciri gaya foto ini, bahasa Indonesia, maks 15 kata>"}'
    )
    data = _gemini_vision_json(prompt, image_path)
    return data if isinstance(data, dict) else None

def _queue_revision(mem, rec, reason_label, new_photo=False, slides=None):
    """Masukkan konten ke antrian revisi (biar run berikutnya di-generate ulang).
    new_photo=True -> foto dicari baru; False -> foto lama (cache) dipakai lagi.
    slides=[nomor] -> hanya slide itu yg direvisi; []/None = semua slide."""
    page_id = ((rec or {}).get("page_id") or "").replace("-", "")   # normalisasi kunci
    if not page_id:
        return False
    rev = (rec or {}).get("revision", 0)
    q = mem.setdefault("revise_queue", {})
    entry = q.get(page_id) or {"title": (rec or {}).get("title", "desain"), "revision": rev, "reasons": []}
    entry["title"] = (rec or {}).get("title", entry.get("title", "desain"))
    entry["revision"] = rev + 1
    entry["reasons"].append(reason_label)
    entry["reasons"] = entry["reasons"][-8:]
    # gabung flag "foto baru": sekali minta foto baru -> tetap foto baru
    entry["new_photo"] = bool(entry.get("new_photo", False)) or bool(new_photo)
    # gabung target slide: kalau salah satu minta "semua" ([]) -> jadi semua
    new = [int(x) for x in (slides or [])]
    if "slides" not in entry:
        entry["slides"] = new
    else:
        old = entry.get("slides") or []
        if old == [] or new == []:
            entry["slides"] = []
        else:
            entry["slides"] = sorted(set(old) | set(new))
    q[page_id] = entry
    return True

def _mark_resolved(mem, tok):
    res = mem.setdefault("resolved", [])
    if tok not in res:
        res.append(tok)
    if len(res) > 500:
        del res[:len(res) - 500]

def process_feedback(mem):
    """Baca pencetan tombol sejak run terakhir, update stats + knob.
    Knob diterapkan BERDASARKAN alasan yg dipencet (tidak tergantung 'pending'),
    biar tetap jalan walau catatan pending hilang. Anti-dobel pakai daftar 'resolved'."""
    prefs = mem["preferences"]
    tg_delete_webhook()   # jaga-jaga: webhook nyangkut bikin tombol/ketikan nggak kebaca
    updates = tg_get_updates(mem.get("tg_offset", 0) + 1)
    changed = []
    processed = 0
    for up in updates:
        mem["tg_offset"] = max(mem.get("tg_offset", 0), up.get("update_id", 0))

        # ---- pesan teks: perintah stats / reset / revisi bebas (AI) ----
        msg = up.get("message")
        if msg:
            # ---- FOTO CONTOH: bot belajar "vibe" dari gambar yg dikirim user ----
            photo = msg.get("photo")
            doc = msg.get("document")
            is_img_doc = bool(doc) and str(doc.get("mime_type", "")).startswith("image/")
            if photo or is_img_doc:
                file_id = (photo[-1]["file_id"] if photo else doc["file_id"])
                if not GEMINI_API_KEY:
                    tg_message("Buat belajar dari foto contoh, bot butuh 'mata' (Gemini). "
                               "Setel dulu GEMINI_API_KEY-nya ya 🙏")
                    processed += 1; continue
                ok = False
                fp = tg_get_file_path(file_id)
                if fp:
                    tmp = os.path.join(tempfile.gettempdir(), f"contoh_{int(time.time()*1000)}.jpg")
                    if tg_download_file(fp, tmp):
                        tg_message("👀 Oke, gw lagi 'lihat' foto contohnya bentar…")
                        res = ai_learn_from_example(tmp)
                        if res:
                            ch = apply_example(prefs, res)
                            note = (res.get("note") or "").strip()
                            # cari desain buat LANGSUNG direvisi: yg lagi nunggu dinilai, atau yg terakhir dibuat
                            rec = None
                            latest = _latest_pending(mem)
                            if latest:
                                tok, rec = latest
                                mem["pending"].pop(tok, None); _mark_resolved(mem, tok)
                            else:
                                lp = mem.get("last_page") or {}
                                if lp.get("page_id"):
                                    rec = {"title": lp.get("title", "desain"),
                                           "page_id": lp.get("page_id"), "revision": 0}
                            if rec and rec.get("page_id"):
                                _queue_revision(mem, rec, "contoh foto: " + (note or "niru gaya"),
                                                new_photo=False, slides=[])
                                tg_message(f"🧠 Dapet gayanya! {note}\n"
                                           f"Setelan: {prefs_summary(prefs)}\n"
                                           f"Gw langsung revisi '{rec.get('title','desain')}' pake gaya ini ya 🎨\n"
                                           f"(Kalau yg mau diubah konten LAIN, set statusnya ke 'Siap Desain' di Notion.)")
                            else:
                                tg_message(f"🧠 Dapet gayanya! {note}\n"
                                           f"Setelan gw simpen → {prefs_summary(prefs)}\n"
                                           f"Tapi belum ada desain yg bisa gw revisi. Aktifin 1 konten dulu "
                                           f"(set status 'Siap Desain' di Notion), nanti gayanya kepake 🙂")
                            ok = True
                if not ok:
                    tg_message("Hmm, gw gagal baca foto contohnya (mungkin kuota AI penuh / foto rusak). "
                               "Coba kirim ulang bentar lagi ya 🙏")
                processed += 1
                continue

            text = (msg.get("text") or "").strip()
            low = text.lower()
            mem["awaiting_text"] = None   # fitur lama nggak dipakai lagi
            if low in ("stats", "/stats", "statistik"):
                a = mem["stats"]["approved"]; r = mem["stats"]["rejected"]; tot = a + r
                rate = int(100 * a / tot) if tot else 0
                tg_message(f"📊 Statistik belajar Ala NU:\n"
                           f"Approve {a} / Reject {r}  (approval rate {rate}%).\n"
                           f"{prefs_summary(mem['preferences'])}")
                continue
            if low in ("/reset", "reset belajar"):
                mem["preferences"] = dict(DEFAULT_MEMORY["preferences"])
                prefs = mem["preferences"]
                tg_message("🔄 Setelan belajar direset ke awal (netral). Statistik tetap tersimpan.")
                continue
            # ---- CHAT AI: perintah revisi bebas (ketik kalimat biasa) ----
            if text:
                # Penerjemah LOKAL dulu (tanpa AI) -> ketikan SELALU ngefek walau AI mati.
                local_cmd = local_parse_command(text)
                data = ai_parse_command(text)
                data = _merge_cmd(data, local_cmd)
                # PENGAMAN: kalau AI bilang "none" TAPI teksnya jelas keluhan/minta ubah
                # dan ada desain yang nunggu dinilai -> tetap anggap REVISI (biar AI nggak kelewat cuek)
                _rev_kata = ("ganti", "ubah", "revisi", "perbaiki", "kurang", "jelek", "buram",
                             "norak", "kegedean", "kekecilan", "lebih ", "jangan", "terlalu",
                             "bold", "tebal", "tipis", "tebel", "numpuk", "tumpang", "ketimpa", "nabrak")
                if (data and data.get("action") == "none"
                        and any(k in low for k in _rev_kata) and _latest_pending(mem)):
                    data["action"] = "revise"
                if data and data.get("action") == "none":
                    tg_message("Oke, dicatat 🙂 Kalau mau gw revisi, kasih tau yang perlu diubah ya "
                               "(misal: “judul kekecilan”, “overlay kurang gelap”, “ganti foto lebih cerah”).")
                    continue
                note = (data or {}).get("note") or text[:80]
                ch = apply_ai(prefs, data) if data else []
                # foto baru? -> kalau AI bilang new_photo, atau teksnya jelas soal ganti/komentar foto
                _photo_kata = ("foto", "gambar", "fotonya", "gambarnya", "aesthetic",
                               "jelek", "buram", "norak", "cari gambar", "ganti foto")
                _said_photo = any(k in low for k in _photo_kata)
                new_photo = bool((data or {}).get("new_photo")) or (data is None and _said_photo) or (_said_photo and ("ganti" in low or "jelek" in low or "aesthetic" in low or "lain" in low))
                # kalau AI nggak aktif & user nggak nyebut foto -> aman pakai foto lama (cache)
                slides = (data or {}).get("slides") or _slide_nums_in(text)
                latest = _latest_pending(mem)
                if latest:
                    tok, rec = latest
                    _queue_revision(mem, rec, "chat: " + note, new_photo=new_photo, slides=slides)
                    mem["pending"].pop(tok, None); _mark_resolved(mem, tok)
                    mem["stats"]["rejected"] += 1
                    extra = ("\n" + prefs_summary(prefs)) if ch else ""
                    tail = "" if data else " (AI lagi nggak aktif)"
                    seg = ("foto baru" if new_photo else "foto lama dipakai lagi")
                    if slides:
                        seg += f", cuma slide {', '.join(str(x) for x in slides)}"
                    tg_message(f"🧠 Paham: {note}. Gw revisi '{rec.get('title','desain')}' ya ({seg}){tail}.{extra}")
                else:
                    if ch:
                        tg_message(f"🧠 Paham: {note}. Setelan disesuaikan, kepakai di desain berikutnya.\n{prefs_summary(prefs)}")
                    else:
                        tg_message("Belum ada desain yang bisa direvisi. Kirim konten dulu ya 🙂")
                processed += 1
                continue
            continue

        # ---- pencetan tombol (callback) ----
        cb = up.get("callback_query")
        if not cb:
            continue
        data = cb.get("data", "")
        cb_id = cb.get("id", "")
        parts = data.split(":")
        kind = parts[0]
        tok = parts[1] if len(parts) > 1 else ""
        if tok and tok in mem.get("resolved", []):
            tg_answer_callback(cb_id, "Sudah dinilai sebelumnya 👍"); continue

        # page_id & code diambil dari DATA TOMBOL (nggak gantung ke pending yg bisa hilang)
        if kind == "rs":
            code = parts[2] if len(parts) > 2 else ""
            pid_cb = parts[3] if len(parts) > 3 else ""
        else:
            code = ""
            pid_cb = parts[2] if len(parts) > 2 else ""
        pend = mem["pending"].get(tok, {})
        rec = {"title": pend.get("title", "desain"),
               "page_id": pid_cb or pend.get("page_id"),
               "revision": pend.get("revision", 0)}

        if kind == "a":                      # approve (1x pencet) -> FINAL
            mem["pending"].pop(tok, None)
            mem["stats"]["approved"] += 1
            _mark_resolved(mem, tok); processed += 1
            pid_norm = (rec.get("page_id") or "").replace("-", "")
            if pid_norm:
                mem.get("revise_queue", {}).pop(pid_norm, None)   # batal revisi, udah oke
            mem["log"].append({"ts": int(time.time()), "title": rec["title"],
                               "status": "approved", "reason": ""})
            tg_answer_callback(cb_id, "✅ Disimpan sebagai contoh bagus!")
        elif kind == "rs":                   # reject + alasan (1x pencet) -> knob + antri revisi
            if code == "other":   # tombol lama: sekarang cukup KETIK langsung
                tg_answer_callback(cb_id, "Ketik aja perintah revisimu langsung ya 🙂")
                tg_message("💬 Ketik aja apa yang mau diubah (misal: “judul kekecilan, foto lebih cerah”) — "
                           "langsung gw proses, nggak usah pencet tombol ini.")
            else:
                mem["stats"]["rejected"] += 1
                desc = apply_reason(prefs, code)
                mem["pending"].pop(tok, None)
                _mark_resolved(mem, tok); processed += 1
                ok = _queue_revision(mem, rec, REASON_LABEL.get(code, code),
                                     new_photo=(code == "img"), slides=[])
                print(f"  reject({code}) pid={rec.get('page_id')} queued={ok}")
                mem["log"].append({"ts": int(time.time()), "title": rec["title"],
                                   "status": "rejected", "reason": REASON_LABEL.get(code, code)})
                changed.append(desc)
                tg_answer_callback(cb_id, f"Paham. {desc}. Gw revisi ya.")
        elif kind == "r":                    # kompat tombol lama (Reject 2-langkah) -> abaikan halus
            tg_answer_callback(cb_id, "Pakai tombol alasan di desain baru ya 🙏")

    # log diagnostik di console (biar kelihatan apa yg kebaca)
    print(f"  feedback: {len(updates)} update, {processed} diproses, "
          f"revise_queue={len(mem.get('revise_queue', {}))}, offset={mem.get('tg_offset')}")

    # laporan ke Telegram biar KELIHATAN bot baca feedback
    if processed > 0:
        head = f"🔎 Feedback terbaca: {processed} pencetan diproses."
        if changed:
            head += "\n🧠 Bot menyesuaikan diri: " + "; ".join(changed) + "."
        head += "\n" + prefs_summary(prefs)
        tg_message(head)
    mem["_last_processed"] = processed
    return mem

def apply_prefs_to_globals(mem):
    global G_HEAD_SCALE, G_BODY_SCALE, G_SCRIM, G_SOFT, G_HEAD_BOLD
    p = mem.get("preferences", {})
    G_HEAD_SCALE = float(p.get("head_scale", 1.0))
    G_BODY_SCALE = float(p.get("body_scale", 1.0))
    G_SCRIM      = float(p.get("scrim", 1.0))
    G_SOFT       = float(p.get("soft", 0.0))
    G_HEAD_BOLD  = bool(p.get("head_bold", True))


# ======================================================================
# 9. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir, mem=None, revision=None):
    props = page["properties"]
    page_id = page.get("id", "")
    title = get_title(props) or "Tanpa Judul"
    fmt = read_property(props, FORMAT_PROPERTY, "select") or "Single Post"
    content = read_property(props, CONTENT_PROPERTY, "rich_text")
    img_field = read_property(props, IMGURL_PROPERTY, "url") or read_property(props, IMGURL_PROPERTY, "rich_text")
    theme = read_property(props, THEME_PROPERTY, "select") or read_property(props, THEME_PROPERTY, "status") or DEFAULT_THEME

    if not content.strip():
        raise ValueError("Kolom 'Konten' kosong.")

    slide_blocks = split_slides(content)
    if "single" in fmt.lower():
        slide_blocks = slide_blocks[:1]
    col_urls = split_urls(img_field)
    keyword_blocks = split_slides(read_property(props, KEYWORD_PROPERTY, "rich_text"))
    # kata kunci gambar pintar (AI) — cuma kepakai di jalur AUTO (kalau tak tempel link).
    # 1 panggilan utk semua slide; None kalau AI mati/gagal.
    ai_kws = ai_image_keywords(slide_blocks)
    if ai_kws:
        print(f"  [AI keywords: {ai_kws}]")

    print(f"  -> '{title}' | {fmt} | {len(slide_blocks)} slide | tema={theme}")

    slides_data, preview_paths = [], []
    vision_notes = []           # catatan cek mandiri (AI vision) — cuma cover
    smart_crop_slides = []      # (nomor slide, metode) yg berhasil auto-crop -> buat lapor Telegram
    # mulai dari daftar foto yg masih 'cooldown' (dipakai < 6 bln) -> otomatis dihindari
    used_ids = set(photo_block_set(mem)) if mem is not None else set()
    total = len(slide_blocks)

    # --- REUSE CACHE: pas revisi, foto yg udah di-approve dipakai lagi (per slide) ---
    pid_norm = _pid_norm(page_id)
    content_hash = _content_hash(content)
    rev_reuse = None
    if revision is not None and mem is not None:
        cached = (mem.get("design_cache") or {}).get(pid_norm)
        if cached and cached.get("hash") == content_hash and int(cached.get("n", 0)) == total:
            rev_reuse = {
                "new_photo": bool(revision.get("new_photo", False)),
                "slides": [int(x) for x in (revision.get("slides") or [])],
                "dir": _page_cache_dir(page_id),
            }
            tgt = rev_reuse["slides"] or "semua"
            print(f"  [reuse] cache cocok (hash={content_hash}) foto_baru={rev_reuse['new_photo']} target={tgt}")
    slide_cache = {}

    for i, block in enumerate(slide_blocks, start=1):
        inline_url, clean_block = extract_url_from_block(block)
        headline, body = headline_and_body(clean_block)
        url = inline_url or (col_urls[i - 1] if i - 1 < len(col_urls) else "")
        img_path = os.path.join(workdir, f"slide_{i}.png")
        clean_path = os.path.join(workdir, f"slide_{i}_clean.png")
        full_path = os.path.join(workdir, f"slide_{i}_full.png") if OUTPUT_HTML else None   # foto UTUH (crop editable Figma)
        crop_rect = None       # posisi auto-crop (fraksi) di foto utuh

        # pakai foto lama (cache)? ya, kalau slide ini BUKAN target foto baru & file cache ada
        reused = False
        if rev_reuse is not None:
            cpath = os.path.join(rev_reuse["dir"], f"slide_{i}.png")
            cpath_clean = os.path.join(rev_reuse["dir"], f"slide_{i}_clean.png")
            if os.path.exists(cpath):
                targeted = (not rev_reuse["slides"]) or (i in rev_reuse["slides"])
                if not (rev_reuse["new_photo"] and targeted):
                    shutil.copyfile(cpath, img_path)   # foto lama dipakai lagi
                    if os.path.exists(cpath_clean):
                        shutil.copyfile(cpath_clean, clean_path)
                    reused = True
                    print(f"    slide {i}: foto lama (cache) dipakai lagi")

        if not reused:
            if url:                                   # ada link (Pinterest) -> pakai
                crop_rect, _cm = make_bw_photo(fetch_image_bytes(url), img_path, clean_out=clean_path,
                                               full_out=full_path, context=(headline + " " + body))
                if _cm: smart_crop_slides.append((i, _cm))
            else:                                     # tak ada link -> auto Unsplash/Pexels
                explicit = keyword_blocks[i - 1].strip() if (i - 1) < len(keyword_blocks) else ""
                if explicit:
                    q = keyword_for(i - 1, keyword_blocks, headline, body)
                elif ai_kws and (i - 1) < len(ai_kws) and ai_kws[i - 1].strip():
                    q = ai_kws[i - 1].strip()          # kata kunci FOTO dari AI (Inggris)
                else:
                    q = keyword_for(i - 1, keyword_blocks, headline, body)
                _before = set(used_ids)
                data = auto_image_bytes(q, used_ids)
                if data and mem is not None:
                    record_photos_used(mem, used_ids - _before)   # catat anti pakai-ulang
                if not data:
                    raise ValueError(
                        f"Slide {i}: tak ada link gambar DAN auto gagal. "
                        f"Tempel link di slide, isi 'Kata Kunci Gambar' (Inggris), "
                        f"atau isi secret UNSPLASH_ACCESS_KEY / PEXELS_API_KEY."
                    )
                crop_rect, _cm = make_bw_photo(data, img_path, clean_out=clean_path,
                                               full_out=full_path, context=(headline + " " + body))
                if _cm: smart_crop_slides.append((i, _cm))

        has_clean = os.path.exists(clean_path)

        # simpan foto final slide ini ke cache (buat revisi berikutnya)
        try:
            cdir = _page_cache_dir(page_id)
            shutil.copyfile(img_path, os.path.join(cdir, f"slide_{i}.png"))
            if has_clean:
                shutil.copyfile(clean_path, os.path.join(cdir, f"slide_{i}_clean.png"))
            slide_cache[str(i)] = {"theme": theme}
        except Exception as _e:
            print("    ! cache simpan gagal:", _e)

        preview_paths.append(img_path)
        slides_data.append({"headline": headline, "body": body, "image_path": img_path,
                            "clean_path": (clean_path if has_clean else None),
                            "full_path": (full_path if (full_path and os.path.exists(full_path)) else None),
                            "crop_rect": crop_rect,
                            "index": i, "total": total, "theme": theme})

    # catat metadata cache (buat cek kecocokan di revisi berikutnya)
    if mem is not None:
        mem.setdefault("design_cache", {})[pid_norm] = {
            "hash": content_hash, "n": total, "title": title,
            "slides": slide_cache, "ts": int(time.time())}
        mem["last_page"] = {"page_id": page_id, "title": title, "ts": int(time.time())}

    # --- CEK MANDIRI (AI vision) — CUKUP slide COVER biar hemat kuota gratis Gemini ---
    if slides_data:
        # tema teks di foto cover: panel "Hitam" -> handle putih (dark); panel "Putih" -> area teks kebaca (light)
        vis_theme = "dark" if str(theme).strip().lower().startswith("hitam") else "light"
        chk = ai_vision_check(slides_data[0]["image_path"], "bottom", vis_theme)
        if chk is None:
            print("    [vision cover] dilewati/gagal (kuota/AI nggak jawab)")
        elif not chk["ok"]:
            vision_notes.append(f"Cover: ⚠️ {chk['issue'] or 'teks mungkin kurang terbaca'}")
            print(f"    [vision cover] ⚠️ {chk['issue']}")
        else:
            vision_notes.append("Cover: ✅ aman")
            print("    [vision cover] aman")

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()

    lines = []
    if revision:
        rev_n = revision.get("revision", 0)
        prev = ", ".join(revision.get("reasons", [])[-4:]) or "-"
        lines += [f"🔄 Merevisi '{title}' (revisi ke-{rev_n})",
                  f"⚠️ Catatan dari sebelumnya: {prev}", ""]
    lines += [f"🕌 {title}  ({fmt}, {total} slide, tema {theme})", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:150] + "…") if len(s["body"]) > 150 else s["body"]
        lines.append(f"— Slide {i}: {h}")
        if b:
            lines.append(f"  {b}")
    if VISION_CHECK_ENABLED:
        warn = [n for n in vision_notes if "⚠️" in n]
        if warn:
            lines += ["", "🔎 Cek mandiri (AI vision):", *warn,
                      "   (kalau mau diperbaiki, pencet Reject atau ketik revisimu)"]
        elif vision_notes:
            lines += ["", "🔎 Cek mandiri (cover): aman ✅"]
        else:
            lines += ["", "🔎 Cek mandiri: dilewati — kuota AI lagi penuh, nanti dicoba lagi"]
    # laporan auto-crop (biar keliatan jalan apa nggak + metodenya apa)
    if smart_crop_slides:
        from collections import OrderedDict
        _lbl = {"ai": "AI", "wajah": "wajah", "orang": "orang", "saliency": "saliency", "lokal": "lokal"}
        groups = OrderedDict()
        for i, mth in smart_crop_slides:
            groups.setdefault(_lbl.get(mth, mth), []).append(str(i))
        seg = [f"{k}: slide {', '.join(v)}" for k, v in groups.items()]
        lines += ["", "✂️ Auto-crop jalan — " + "; ".join(seg)]
    else:
        lines += ["", "✂️ Auto-crop: crop tengah (subjek nggak kedeteksi / foto polos)"]
    tg_message("\n".join(lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    tg_message(f"📎 ====================\nFILE DESAIN: {title}\n====================")

    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path)
        if not tg_document(pptx_path, caption=f"{title} — .pptx (semua slide): import ke Canva ✨"):
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg(slides_data, svg_path, logo_path)
        if not tg_document(svg_path, caption=f"{title} — .svg (semua slide): tarik ke Figma ✨"):
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    if OUTPUT_SVG_PER_SLIDE and total > 1:
        for s in slides_data:
            sp = os.path.join(workdir, f"{safe_name}_slide{s['index']}.svg")
            build_svg([s], sp, logo_path)
            if not tg_document(sp, caption=f"{title} — slide {s['index']}/{total} (.svg per slide)"):
                tg_message(f"⚠️ SVG slide {s['index']} gagal dikirim (cek log).")

    if OUTPUT_HTML:
        html_path = os.path.join(workdir, f"{safe_name}.html")
        build_html(slides_data, html_path, logo_path)
        cap = (f"{title} — .html (teks jadi satu paragraf rapih, masking foto terpisah).\n"
               "Edit di Figma: pakai plugin html.to.design → Import → upload file ini. "
               "Teks jadi text box asli (bukan per-baris).")
        if not tg_document(html_path, caption=cap):
            tg_message("⚠️ File .html gagal dikirim (cek log).")

    # ---- tombol penilaian (1x pencet langsung selesai) utk sistem belajar ----
    if LEARN_ENABLED and mem is not None:
        rev_n = (revision or {}).get("revision", 0)
        token = secrets.token_hex(4)
        pid_cb = (page_id or "").replace("-", "")     # page_id ditaruh di tombol (anti-hilang)
        mem["pending"][token] = {"title": title, "page_id": page_id,
                                 "ts": int(time.time()), "revision": rev_n}
        rows = [[("✅ Approve (oke!)", f"a:{token}:{pid_cb}")]]
        rr = []
        for label, code in REASON_BUTTONS:
            rr.append(("❌ " + label, f"rs:{token}:{code}:{pid_cb}"))
            if len(rr) == 2:
                rows.append(rr); rr = []
        if rr:
            rows.append(rr)
        head = f"Nilai desain '{title}' 👇"
        if rev_n:
            prev = ", ".join((revision or {}).get("reasons", [])[-4:]) or "-"
            head = (f"🔄 REVISI ke-{rev_n} dari '{title}' 👇\n"
                    f"⚠️ Catatan dari sebelumnya: {prev}")
        tip = ("\n💬 atau ketik aja perintahmu (misal: “judul kekecilan, foto lebih cerah”)"
               if AI_CHAT_ENABLED else "")
        tg_buttons(head + "\n✅ kalau udah oke — atau pencet alasannya kalau masih kurang pas:" + tip, rows)

    return total


# ======================================================================
# 10. MAIN
# ======================================================================
def main():
    print(f"== ALANU BOT DESAIN ({CANVAS_W}x{CANVAS_H}) | unsplash={'on' if UNSPLASH_ACCESS_KEY else 'off'} | belajar={'on' if LEARN_ENABLED else 'off'} ==")

    # 1) muat memori + 2) proses feedback (pencetan tombol) sejak run lalu
    mem = load_memory()
    if LEARN_ENABLED:
        try:
            mem = process_feedback(mem)
        except Exception as e:
            print("  ! proses feedback gagal:", e)
        apply_prefs_to_globals(mem)
        print("  " + prefs_summary(mem["preferences"]))

    # 3) AUTO-REVISI: generate ulang konten yg tadi di-reject (sampai di-approve)
    if LEARN_ENABLED and mem.get("revise_queue"):
        for pid in list(mem["revise_queue"].keys()):
            info = mem["revise_queue"].pop(pid)   # pindah dari antrian -> jadi pending lagi
            if info.get("revision", 1) > MAX_REVISI:
                tg_message(f"🛑 '{info.get('title','desain')}' udah direvisi {MAX_REVISI}x tapi belum pas. "
                           f"Mungkin lebih enak kamu kasih contoh manual biar gw niru. Gw stop auto-revisi yg ini dulu.")
                continue
            page = notion_get_page(pid)
            if not page:
                tg_message(f"⚠️ Gagal ambil ulang '{info.get('title','desain')}' dari Notion buat revisi.")
                continue
            wd = tempfile.mkdtemp()
            try:
                tg_message(f"🔄 Merevisi '{info.get('title','desain')}' (revisi ke-{info.get('revision',1)})…")
                process_page(page, wd, mem, revision=info)
            except Exception as e:
                print("  x revisi gagal:", e); traceback.print_exc()
                tg_message(f"⚠️ Revisi '{info.get('title','desain')}' gagal: {type(e).__name__}. Gw coba lagi next.")
                mem["revise_queue"][pid] = info   # balikin ke antrian biar dicoba lagi

    # 4) konten BARU berstatus "Siap Desain"
    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        # tidak ada kerjaan -> jangan spam Telegram tiap run, cukup simpan memori
        print("Tidak ada konten baru. (Feedback tetap diproses.)")
        save_memory(mem)
        return

    tg_message(f"✅ Robot Ala NU jalan — ada {len(pages)} konten baru.")
    for page in pages:
        page_id = page["id"]
        workdir = tempfile.mkdtemp()
        try:
            process_page(page, workdir, mem)
            notion_set_status(page_id, STATUS_DONE, icon_emoji=DONE_EMOJI)
            print("  v Sukses & status ->", STATUS_DONE, "| ikon ->", DONE_EMOJI)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            print("  x GAGAL:", err)
            traceback.print_exc()
            try:
                tg_message(f"⚠️ Gagal memproses '{get_title(page['properties'])}'.\n{err}")
            except Exception:
                pass
            notion_set_status(page_id, STATUS_ERROR, note=err)

    save_memory(mem)
    print("== Selesai ==")


if __name__ == "__main__":
    main()
