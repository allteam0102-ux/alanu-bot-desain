# -*- coding: utf-8 -*-
"""
ALANU BOT DESAIN  (gaya @ala_nu — foto hitam-putih + panel teks serif)
======================================================================
Alur: Notion (Status "Siap Desain") -> ambil gambar dari kolom 'Gambar URL'
(tempel link gambar per slide) -> jadikan hitam-putih + grain -> rakit desain
gaya ala_nu -> kirim ke Telegram. Menghasilkan:
  1. .pptx -> import ke Canva (teks & gambar bisa diedit)
  2. .svg  -> tarik ke Figma / Illustrator (teks bisa diedit)

TIDAK pakai Pexels & TIDAK pakai AI. Gambar 100% dari URL yang kamu tempel.
"""

import os
import re
import io
import html
import time
import base64
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
# >>> PENGATURAN BRAND (boleh kamu ubah) <<<
# ======================================================================
HANDLE        = os.environ.get("HANDLE")        or "@ala_nu"
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "E3B23C").lstrip("#")   # kuning/emas
LOGO_URL      = os.environ.get("LOGO_URL")      or ""                       # logo NU (PNG transparan), opsional
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Playfair Display"       # serif elegan
BODY_FONT     = os.environ.get("BODY_FONT")     or "Poppins"
ARROW_TEXT    = os.environ.get("ARROW_TEXT")    or "→"                 # panah geser
DEFAULT_THEME = (os.environ.get("DEFAULT_THEME") or "Putih").strip().lower()  # putih / hitam
PHOTO_FRAC    = float(os.environ.get("PHOTO_FRAC") or 0.46)                 # tinggi foto (bagian atas)

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"


# ======================================================================
# 1. PENGATURAN TEKNIS (dari Secrets)
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

TITLE_PROPERTY   = env("TITLE_PROPERTY", "Judul")
STATUS_PROPERTY  = env("STATUS_PROPERTY", "Status")
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Format")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Konten")
IMGURL_PROPERTY  = env("IMGURL_PROPERTY", "Gambar URL")
THEME_PROPERTY   = env("THEME_PROPERTY", "Tema")

STATUS_READY = env("STATUS_READY", "Siap Desain")
STATUS_DONE  = env("STATUS_DONE",  "Terkirim")
STATUS_ERROR = env("STATUS_ERROR", "Gagal")
STATUS_TYPE  = env("STATUS_TYPE", "select").strip().lower()

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
# 2. HELPER NOTION
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

def notion_set_status(page_id, status_value, note=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    r = requests.patch(url, headers=NOTION_HEADERS, json={"properties": props_with_note}, timeout=60)
    if r.status_code != 200 and note:
        requests.patch(url, headers=NOTION_HEADERS, json={"properties": props}, timeout=60)

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


# ======================================================================
# 3. PARSING ISI KONTEN
# ======================================================================
def split_slides(content_text):
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"(?m)^\s*---\s*$", raw)
    return [b.strip() for b in blocks if b.strip()]

def split_urls(text):
    """URL dipisah baris ATAU '---'. Kembalikan list bersih."""
    raw = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub(r"(?m)^\s*---\s*$", "\n", raw)
    return [u.strip() for u in raw.split("\n") if u.strip()]

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
# 4. GAMBAR (dari URL) -> hitam-putih + grain
# ======================================================================
def fetch_image_bytes(url):
    r = requests.get(url, headers=UA, timeout=60)
    r.raise_for_status()
    ct = r.headers.get("Content-Type", "").lower()
    if "image" not in ct:
        raise ValueError(
            f"URL bukan gambar langsung (Content-Type: {ct or 'tidak diketahui'}). "
            f"Di Pinterest: klik kanan gambarnya -> 'Copy image address' (link diakhiri .jpg/.png), "
            f"bukan menyalin link halaman pin."
        )
    return r.content

def make_bw_photo(img_bytes, out_path):
    """Potong ke area foto (atas), jadikan grayscale + grain halus."""
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    PH = int(CANVAS_H * PHOTO_FRAC)
    target = CANVAS_W / PH
    w, h = im.size
    if w / h > target:
        nw = int(h * target); x = (w - nw) // 2; im = im.crop((x, 0, x + nw, h))
    else:
        nh = int(w / target); y = (h - nh) // 2; im = im.crop((0, y, w, y + nh))
    im = im.resize((CANVAS_W, PH), Image.LANCZOS)

    gray = ImageOps.grayscale(im)                       # hitam-putih
    gray = ImageOps.autocontrast(gray, cutoff=1)        # kontras rapi
    try:
        grain = Image.effect_noise((CANVAS_W, PH), 22)  # tekstur film
        gray = Image.blend(gray, grain, 0.08)
    except Exception:
        pass
    photo = gray.convert("RGB")

    # sedikit gelap di atas biar logo & handle terbaca
    overlay = Image.new("RGBA", (CANVAS_W, PH), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    top_h = int(PH * 0.28)
    for y in range(0, top_h):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(110 * (1 - y / top_h))))
    photo = Image.alpha_composite(photo.convert("RGBA"), overlay).convert("RGB")
    photo.save(out_path, "PNG")
    return out_path


# ======================================================================
# 5. WARNA TEMA
# ======================================================================
def theme_colors(theme):
    """Kembalikan (panel_bg_hex, headline_hex, body_hex) sesuai tema."""
    if str(theme).strip().lower().startswith("hitam"):
        return "0E0E0E", "FFFFFF", "CFCFCF"
    return "FFFFFF", "141414", "4A4A4A"   # Putih (default)


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
# 6. RAKIT .PPTX (untuk Canva)
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
            f.size = Pt(size_pt)
            f.bold = bold
            f.italic = italic
            f.name = font
            f.color.rgb = ACCENT if is_hl else hex_rgb(color_hex)
    return box

def build_pptx(slides_data, out_path, logo_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W)
    prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]

    for s in slides_data:
        panel_bg, head_c, body_c = theme_colors(s["theme"])
        slide = prs.slides.add_slide(blank)

        # panel bawah (latar penuh dulu, lalu foto di atas)
        panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        panel.fill.solid(); panel.fill.fore_color.rgb = hex_rgb(panel_bg)
        panel.line.fill.background(); panel.shadow.inherit = False

        # foto (bagian atas)
        slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=fy(PHOTO_FRAC))

        # logo + handle di atas foto
        if logo_path:
            try:
                slide.shapes.add_picture(logo_path, fx(0.06), fy(0.05), height=fy(0.05))
            except Exception:
                pass
        _add_text(slide, 0.50, 0.05, 0.44, 0.06, HANDLE, 13, True, BODY_FONT, "FFFFFF", align=PP_ALIGN.RIGHT)

        # headline (serif) + body
        top = PHOTO_FRAC + 0.05
        if s["headline"]:
            _add_text(slide, 0.07, top, 0.86, 0.20, s["headline"], 33, True, HEADLINE_FONT, head_c, highlight=True)
        if s["body"]:
            _add_text(slide, 0.07, top + 0.20, 0.86, 0.22, s["body"], 18, False, BODY_FONT, body_c, highlight=True)

        # footer handle + panah
        _add_text(slide, 0.06, 0.935, 0.5, 0.05, HANDLE, 12, False, BODY_FONT, body_c)
        if s["total"] > 1:
            _add_text(slide, 0.80, 0.925, 0.14, 0.06, ARROW_TEXT, 24, True, BODY_FONT, head_c, align=PP_ALIGN.RIGHT)

    prs.save(out_path)
    return out_path


# ======================================================================
# 7. RAKIT .SVG (untuk Figma) — SATU file, semua slide berjejer
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
    out = []
    for i, line in enumerate(_wrap(text, max_chars)):
        base = y + size + int(i * size * gap)
        spans = ""
        segs = parse_highlights(line) if highlight else [(line, False)]
        for seg, is_hl in segs:
            fill = ACCENT_COLOR if is_hl else color_hex
            spans += f'<tspan fill="#{fill}">{_svg_escape(seg)}</tspan>'
        out.append(
            f'<text x="{x}" y="{base}" text-anchor="{anchor}" '
            f'font-family="{_svg_escape(font)}, Georgia, serif" '
            f'font-size="{size}" font-weight="{weight}">{spans}</text>'
        )
    return "\n".join(out)

def build_svg_single(slides_data, out_path, logo_path):
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
        # panel penuh
        parts.append(f'<rect x="{xo}" y="0" width="{CANVAS_W}" height="{H}" fill="#{panel_bg}"/>')
        # foto atas
        parts.append(
            f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{PH}" '
            f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(s["image_path"])}"/>'
        )
        # logo + handle
        if logo_uri:
            parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*CANVAS_H)}" '
                         f'height="{int(0.05*CANVAS_H)}" href="{logo_uri}"/>')
        parts.append(_svg_text(HANDLE, xo + int(0.94*CANVAS_W), int(0.05*CANVAS_H),
                               int(0.4*CANVAS_W), 24, True, BODY_FONT, "FFFFFF", anchor="end"))
        # headline serif + body
        ty = int((PHOTO_FRAC + 0.05) * CANVAS_H)
        if s["headline"]:
            parts.append(_svg_text(s["headline"], xo + int(0.07*CANVAS_W), ty,
                                   int(0.86*CANVAS_W), 52, True, HEADLINE_FONT, head_c, highlight=True))
        if s["body"]:
            parts.append(_svg_text(s["body"], xo + int(0.07*CANVAS_W), ty + int(0.20*CANVAS_H),
                                   int(0.86*CANVAS_W), 30, False, BODY_FONT, body_c, highlight=True))
        # footer handle + panah
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
        return _tg_call("sendPhoto", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                        {"photo": f}, label="photo")

def tg_document(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendDocument", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                        {"document": f}, label="doc")


# ======================================================================
# 9. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir):
    props = page["properties"]
    title = read_property(props, TITLE_PROPERTY, "title") or "Tanpa Judul"
    fmt = read_property(props, FORMAT_PROPERTY, "select") or "Single Post"
    content = read_property(props, CONTENT_PROPERTY, "rich_text")
    # kolom Gambar URL bisa bertipe URL atau Text -> coba dua-duanya
    img_field = read_property(props, IMGURL_PROPERTY, "url") or read_property(props, IMGURL_PROPERTY, "rich_text")
    theme = read_property(props, THEME_PROPERTY, "select") or DEFAULT_THEME

    if not content.strip():
        raise ValueError("Kolom 'Konten' kosong.")

    slide_blocks = split_slides(content)
    if "single" in fmt.lower():
        slide_blocks = slide_blocks[:1]
    urls = split_urls(img_field)
    if not urls:
        raise ValueError("Kolom 'Gambar URL' kosong. Tempel minimal 1 link gambar (per slide dipisah baris/---).")

    print(f"  -> '{title}' | {fmt} | {len(slide_blocks)} slide | tema={theme}")

    slides_data, preview_paths = [], []
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        headline, body = headline_and_body(block)
        url = urls[i - 1] if i - 1 < len(urls) else urls[-1]   # kalau URL kurang, pakai yang terakhir
        img_path = os.path.join(workdir, f"slide_{i}.png")
        make_bw_photo(fetch_image_bytes(url), img_path)
        preview_paths.append(img_path)
        slides_data.append({"headline": headline, "body": body, "image_path": img_path,
                            "index": i, "total": total, "theme": theme})

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()

    # ringkasan caption
    lines = [f"🕌 {title}  ({fmt}, {total} slide, tema {theme})", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:150] + "…") if len(s["body"]) > 150 else s["body"]
        lines.append(f"— Slide {i}: {h}")
        if b:
            lines.append(f"  {b}")
    tg_message("\n".join(lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path)
        if not tg_document(pptx_path, caption=f"{title} — .pptx: import ke Canva ✨"):
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg_single(slides_data, svg_path, logo_path)
        if not tg_document(svg_path, caption=f"{title} — .svg: tarik ke Figma (teks bisa diedit) ✨"):
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    return total


# ======================================================================
# 10. MAIN
# ======================================================================
def main():
    print(f"== ALANU BOT DESAIN ({CANVAS_W}x{CANVAS_H}) | pptx={OUTPUT_PPTX} svg={OUTPUT_SVG} ==")
    if tg_message("✅ Alanu bot desain — mulai jalan."):
        print("  Telegram OK.")
    else:
        print("  !! Telegram BERMASALAH — cek TELEGRAM_BOT_TOKEN & TELEGRAM_CHAT_ID.")
    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        print("Tidak ada yang perlu diproses. Selesai.")
        return

    for page in pages:
        page_id = page["id"]
        workdir = tempfile.mkdtemp()
        try:
            process_page(page, workdir)
            notion_set_status(page_id, STATUS_DONE)
            print("  v Sukses & status diubah jadi:", STATUS_DONE)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            print("  x GAGAL:", err)
            traceback.print_exc()
            try:
                tg_message(f"⚠️ Gagal memproses '{read_property(page['properties'], TITLE_PROPERTY, 'title')}'.\n{err}")
            except Exception:
                pass
            notion_set_status(page_id, STATUS_ERROR, note=err)

    print("== Selesai ==")


if __name__ == "__main__":
    main()
