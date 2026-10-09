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
PHOTO_FRAC    = float(os.environ.get("PHOTO_FRAC") or 0.46)

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"
OUTPUT_SVG_PER_SLIDE = (os.environ.get("OUTPUT_SVG_PER_SLIDE") or "true").lower() == "true"

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

def _clamp(v, lo, hi):
    return max(lo, min(hi, v))

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
    blocks = re.split(r"(?m)^\s*(?:[-–—]{2,}|[–—])\s*$", raw)
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
    fresh = [p for p in photos if p.get("id") not in used_ids]
    pool = fresh if fresh else photos
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
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
    fresh = [p for p in hits if p.get("id") not in used_ids]
    pool = fresh if fresh else hits
    photo = random.choice(pool[:10])
    if photo.get("id"):
        used_ids.add(photo["id"])
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
    fresh = [p for p in results if p.get("id") not in used_ids]
    pool = fresh if fresh else results
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
    urls = photo.get("urls", {})
    return requests.get(urls.get("regular") or urls.get("full") or urls.get("raw"), timeout=60).content

def auto_image_bytes(query, used_ids):
    """Ambil otomatis (tanpa link): Pixabay + Pexels + Unsplash. Semua dibias Islami + sesuai konteks."""
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

def make_bw_photo(img_bytes, out_path):
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    PH = int(CANVAS_H * PHOTO_FRAC)
    target = CANVAS_W / PH
    w, h = im.size
    if w / h > target:
        nw = int(h * target); x = (w - nw) // 2; im = im.crop((x, 0, x + nw, h))
    else:
        nh = int(w / target); y = (h - nh) // 2; im = im.crop((0, y, w, y + nh))
    im = im.resize((CANVAS_W, PH), Image.LANCZOS)
    gray = ImageOps.autocontrast(ImageOps.grayscale(im), cutoff=1)
    try:
        grain = Image.effect_noise((CANVAS_W, PH), 22)
        gray = Image.blend(gray, grain, 0.08)
    except Exception:
        pass
    photo = gray.convert("RGB")
    overlay = Image.new("RGBA", (CANVAS_W, PH), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    top_h = int(PH * 0.28)
    for y in range(0, top_h):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(110 * (1 - y / top_h))))
    photo = Image.alpha_composite(photo.convert("RGBA"), overlay).convert("RGB")
    photo.save(out_path, "PNG")
    return out_path


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
        top = PHOTO_FRAC + 0.05
        if s["headline"]:
            _add_text(slide, 0.07, top, 0.86, 0.20, s["headline"], head_pt(), True, HEADLINE_FONT, head_c, highlight=True)
        if s["body"]:
            _add_text(slide, 0.07, top + 0.20, 0.86, 0.22, s["body"], body_pt(), False, BODY_FONT, body_c, highlight=True)
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
    weight = "700" if bold else "400"
    max_chars = max(6, int(width / (size * 0.55)))
    def words_of(line):
        res = []
        for seg, is_hl in (parse_highlights(line) if highlight else [(line, False)]):
            for w in seg.split(" "):
                if w != "":
                    res.append((w, is_hl))
        return res
    out = []
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
            spans = ""
            for k, (w, is_hl) in enumerate(vl):
                prefix = " " if k > 0 else ""
                fill = ACCENT_COLOR if is_hl else color_hex
                spans += f'<tspan fill="#{fill}">{_svg_escape(prefix + w)}</tspan>'
            out.append(f'<text x="{x}" y="{int(base)}" text-anchor="{anchor}" xml:space="preserve" '
                       f'font-family="{_svg_escape(font)}, Georgia, serif" '
                       f'font-size="{size}" font-weight="{weight}">{spans}</text>')
            yy += int(size * gap)
    return "\n".join(out)

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
        ty = int((PHOTO_FRAC + 0.05) * CANVAS_H)
        if s["headline"]:
            parts.append(_svg_text(s["headline"], xo + int(0.07*CANVAS_W), ty,
                                   int(0.86*CANVAS_W), head_px(), True, HEADLINE_FONT, head_c, highlight=True))
        if s["body"]:
            parts.append(_svg_text(s["body"], xo + int(0.07*CANVAS_W), ty + int(0.20*CANVAS_H),
                                   int(0.86*CANVAS_W), body_px(), False, BODY_FONT, body_c, highlight=True))
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


# ======================================================================
# 8b. SISTEM BELAJAR — memori + feedback (approve/reject + alasan)
# ======================================================================
DEFAULT_MEMORY = {
    "preferences": {"head_scale": 1.0, "body_scale": 1.0, "scrim": 1.0, "soft": 0.0},
    "pending": {},        # token -> {title, page_id, ts, revision}
    "resolved": [],       # token yang sudah dinilai (anti dobel)
    "awaiting_text": None,  # token yang menunggu alasan ketik
    "revise_queue": {},   # page_id -> {title, revision, reasons[]} konten yg minta direvisi
    "log": [],            # riwayat {ts, title, status, reason}
    "stats": {"approved": 0, "rejected": 0, "weekly": []},
    "tg_offset": 0,
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
            f"soft +{int(prefs['soft']*100)}%.")


# ---------- OTAK AI: Groq (utama) + Gemini (cadangan). Dua-duanya gratis ----------
GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL    = os.environ.get("GEMINI_MODEL") or "gemini-3.8-flash"
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
PENTING: hampir SEMUA masukan user = permintaan revisi -> set action "revise".
Keluhan/komentar sekecil apapun soal desain (foto/teks/warna/overlay) = "revise".
Contoh yang HARUS "revise": "gambar kurang bagus", "kurang aesthetic", "ganti foto",
"fotonya jelek", "judul kegedean", "kurang jelas", "overlay kurang gelap".
Kalau user komentarin/minta ganti foto, set action "revise" (foto otomatis diganti baru saat revisi).
Pakai action "none" HANYA kalau pesan jelas cuma sapaan/basa-basi/terima kasih/pertanyaan umum
(mis. "halo", "makasih", "lagi apa") — BUKAN komentar soal desain.
Balas HANYA JSON valid, tanpa teks lain:
{"head_scale_delta": <angka -0.12..0.12>, "body_scale_delta": <angka>, "scrim_delta": <angka>, "soft_delta": <angka>, "note": "<ringkasan singkat dalam bahasa Indonesia, maks 12 kata>", "action": "revise" atau "none"}
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
            prefs[knob] = _clamp(prefs[knob] + d, lo, hi)
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
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {"contents": [{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}],
            "generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"}}
    try:
        r = requests.post(url, json=body, timeout=60)
        if r.status_code != 200:
            print(f"    ! Gemini vision {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
    except Exception as e:
        print("    ! Gemini vision gagal:", e); return None

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

def _queue_revision(mem, rec, reason_label):
    """Masukkan konten ke antrian revisi (biar run berikutnya di-generate ulang)."""
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
                data = ai_parse_command(text)
                # PENGAMAN: kalau AI bilang "none" TAPI teksnya jelas keluhan/minta ubah
                # dan ada desain yang nunggu dinilai -> tetap anggap REVISI (biar AI nggak kelewat cuek)
                _rev_kata = ("ganti", "ubah", "revisi", "perbaiki", "kurang", "jelek", "buram",
                             "norak", "kegedean", "kekecilan", "lebih ", "jangan", "terlalu")
                if (data and data.get("action") == "none"
                        and any(k in low for k in _rev_kata) and _latest_pending(mem)):
                    data["action"] = "revise"
                if data and data.get("action") == "none":
                    tg_message("Oke, dicatat 🙂 Kalau mau gw revisi, kasih tau yang perlu diubah ya "
                               "(misal: “judul kekecilan”, “overlay kurang gelap”, “ganti foto lebih cerah”).")
                    continue
                note = (data or {}).get("note") or text[:80]
                ch = apply_ai(prefs, data) if data else []
                latest = _latest_pending(mem)
                if latest:
                    tok, rec = latest
                    _queue_revision(mem, rec, "chat: " + note)
                    mem["pending"].pop(tok, None); _mark_resolved(mem, tok)
                    mem["stats"]["rejected"] += 1
                    extra = ("\n" + prefs_summary(prefs)) if ch else ""
                    tail = "" if data else " (AI lagi nggak aktif, jadi gw revisi dgn foto baru aja)"
                    tg_message(f"🧠 Paham: {note}. Gw revisi '{rec.get('title','desain')}' ya{tail}.{extra}")
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
                ok = _queue_revision(mem, rec, REASON_LABEL.get(code, code))
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
    global G_HEAD_SCALE, G_BODY_SCALE, G_SCRIM, G_SOFT
    p = mem.get("preferences", {})
    G_HEAD_SCALE = float(p.get("head_scale", 1.0))
    G_BODY_SCALE = float(p.get("body_scale", 1.0))
    G_SCRIM      = float(p.get("scrim", 1.0))
    G_SOFT       = float(p.get("soft", 0.0))


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
    used_ids = set()
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        inline_url, clean_block = extract_url_from_block(block)
        headline, body = headline_and_body(clean_block)
        url = inline_url or (col_urls[i - 1] if i - 1 < len(col_urls) else "")
        img_path = os.path.join(workdir, f"slide_{i}.png")
        if url:                                   # ada link (Pinterest) -> pakai
            make_bw_photo(fetch_image_bytes(url), img_path)
        else:                                     # tak ada link -> auto Unsplash/Pexels
            explicit = keyword_blocks[i - 1].strip() if (i - 1) < len(keyword_blocks) else ""
            if explicit:
                q = keyword_for(i - 1, keyword_blocks, headline, body)
            elif ai_kws and (i - 1) < len(ai_kws) and ai_kws[i - 1].strip():
                q = ai_kws[i - 1].strip()          # kata kunci FOTO dari AI (Inggris)
            else:
                q = keyword_for(i - 1, keyword_blocks, headline, body)
            data = auto_image_bytes(q, used_ids)
            if not data:
                raise ValueError(
                    f"Slide {i}: tak ada link gambar DAN auto gagal. "
                    f"Tempel link di slide, isi 'Kata Kunci Gambar' (Inggris), "
                    f"atau isi secret UNSPLASH_ACCESS_KEY / PEXELS_API_KEY."
                )
            make_bw_photo(data, img_path)
        preview_paths.append(img_path)
        slides_data.append({"headline": headline, "body": body, "image_path": img_path,
                            "index": i, "total": total, "theme": theme})

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
