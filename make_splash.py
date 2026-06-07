"""Generate splash.png for VisualStimEdger — transparent card concept."""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import re, pathlib, sys

# ── Read version from main.py ──────────────────────────────────────────────
try:
    src = pathlib.Path("VSE.py").read_text(encoding="utf-8")
    VERSION = re.search(r'VERSION\s*=\s*"([^"]+)"', src).group(1)
except Exception:
    VERSION = "1.3.6"

# ── Canvas / palette ───────────────────────────────────────────────────────
W, H       = 700, 640
CARD_X     = 55
CARD_Y     = 28
CARD_W     = 590
CARD_H     = 584

CARD_COL   = (17,  17,  17, 255)   # #111111
RED        = (204, 34,   0, 255)   # #cc2200
RED_GLOW   = (204, 34,   0, 180)
YELLOW     = (255, 204,  0, 255)   # #ffcc00
WHITE      = (238, 238, 238, 255)  # #eeeeee
GREY       = (100, 100, 100, 255)  # #646464
SUBDIM     = (130, 130, 130, 255)  # #828282
DISC_COL   = ( 80,  80,  80, 255)  # #505050
BTN_FG     = (255, 255, 255, 255)

# ── Fonts (Segoe UI on Windows) ────────────────────────────────────────────
FONT_DIR = pathlib.Path("C:/Windows/Fonts")

def font(size, bold=False):
    name = "segoeuib.ttf" if bold else "segoeui.ttf"
    try:
        return ImageFont.truetype(str(FONT_DIR / name), size)
    except Exception:
        return ImageFont.load_default()

f_title   = font(27, bold=True)
f_sub     = font(11)
f_section = font(13, bold=True)
f_step_h  = font(12, bold=True)
f_step_b  = font(10)
f_num     = font(11, bold=True)
f_disc    = font(10)
f_btn     = font(14, bold=True)

# ── Draw card ─────────────────────────────────────────────────────────────
img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)
draw.rounded_rectangle(
    [CARD_X, CARD_Y, CARD_X + CARD_W, CARD_Y + CARD_H],
    radius=18, fill=CARD_COL
)

# ── Red top glow ──────────────────────────────────────────────────────────
glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
gd   = ImageDraw.Draw(glow)
gd.rounded_rectangle(
    [CARD_X, CARD_Y, CARD_X + CARD_W, CARD_Y + 10],
    radius=18, fill=RED_GLOW
)
glow = glow.filter(ImageFilter.GaussianBlur(radius=7))
img  = Image.alpha_composite(img, glow)

# Solid 3-px bar on top of glow
draw = ImageDraw.Draw(img)
draw.rounded_rectangle(
    [CARD_X, CARD_Y, CARD_X + CARD_W, CARD_Y + 3],
    radius=18, fill=RED
)

# ── Icon ──────────────────────────────────────────────────────────────────
ICX, ICY, ICR = W // 2, 71, 34
draw.ellipse([ICX - ICR, ICY - ICR, ICX + ICR, ICY + ICR], fill=RED)
icon_path = pathlib.Path("icon.ico")
if icon_path.exists():
    try:
        ico = Image.open(icon_path).convert("RGBA")
        ico = ico.resize((50, 50), Image.LANCZOS)
        img.paste(ico, (ICX - 25, ICY - 25), ico)
    except Exception:
        pass

# ── Helper: centered text ──────────────────────────────────────────────────
draw = ImageDraw.Draw(img)

def text_c(y, txt, f, col):
    bb = draw.textbbox((0, 0), txt, font=f)
    x  = (W - (bb[2] - bb[0])) // 2
    draw.text((x, y), txt, fill=col, font=f)

def text_l(x, y, txt, f, col):
    draw.text((x, y), txt, fill=col, font=f)

def wrap(txt, f, max_w):
    words, lines, line = txt.split(), [], ""
    for w in words:
        test = (line + " " + w).strip()
        bb   = draw.textbbox((0, 0), test, font=f)
        if bb[2] - bb[0] > max_w and line:
            lines.append(line)
            line = w
        else:
            line = test
    if line:
        lines.append(line)
    return lines

# ── Title ─────────────────────────────────────────────────────────────────
text_c(117, "VisualStimEdger", f_title, WHITE)
text_c(155, f"v{VERSION}  ·  edge smarter", f_sub, GREY)

# ── Section header ────────────────────────────────────────────────────────
text_l(88, 187, "Before you hit Start", f_section, YELLOW)

# ── Steps ─────────────────────────────────────────────────────────────────
STEPS = [
    ("Open your camera feed",
     "OBS, webcam app, browser stream — anything showing your cock in a window. Full-screen not needed."),
    ("Keep it visible and not minimised",
     "You'll draw a box around that window. Bring it to the front before clicking Start."),
    ("Mark the tip of your cock",
     "Draw a small box around the head. An electrode or ring on the head tracks better than skin."),
    ("Calibrate Flaccid / Erect / Edging heights",
     "Do this during your session. AUTO mode can handle it automatically. Re-calibrate any time."),
]

sy = 217
NR = 12

for i, (title, body) in enumerate(STEPS):
    # Number badge
    nx, ny = 88 + NR, sy + NR
    draw.ellipse([nx - NR, ny - NR, nx + NR, ny + NR], fill=RED)
    nb = draw.textbbox((0, 0), str(i + 1), font=f_num)
    nw = nb[2] - nb[0]
    nh = nb[3] - nb[1]
    draw.text((nx - nw // 2, ny - nh // 2 - 1), str(i + 1), fill=BTN_FG, font=f_num)

    # Step title
    text_l(114, sy, title, f_step_h, WHITE)

    # Body lines
    body_lines = wrap(body, f_step_b, 476)
    by = sy + 18
    for ln in body_lines:
        text_l(114, by, ln, f_step_b, SUBDIM)
        by += 15

    sy += 66

# ── Disclaimer ────────────────────────────────────────────────────────────
text_c(488, "Controls volume only — does not generate e-stim signals.", f_disc, DISC_COL)
text_c(503, "You need Restim, xToys, electron-redrive, an .mp3, etc. already running.", f_disc, DISC_COL)

# ── Start button ──────────────────────────────────────────────────────────
BX1, BY1 = 88, 525
BX2, BY2 = 612, 581
draw.rounded_rectangle([BX1, BY1, BX2, BY2], radius=9, fill=RED)
btn_txt = "I'm ready \u2014 select my camera feed  \u2192"
bb  = draw.textbbox((0, 0), btn_txt, font=f_btn)
bw  = bb[2] - bb[0]
bh  = bb[3] - bb[1]
draw.text(
    ((BX1 + BX2) // 2 - bw // 2, (BY1 + BY2) // 2 - bh // 2),
    btn_txt, fill=BTN_FG, font=f_btn
)

# ── Ko-fi footer ──────────────────────────────────────────────────────────
f_kofi = font(16, bold=True)
_kofi_txt = "enjoyed VSE?  support at ko-fi.com/stimstation"
_bb_txt = draw.textbbox((0, 0), _kofi_txt, font=f_kofi)
_txt_w = _bb_txt[2] - _bb_txt[0]
_txt_h = _bb_txt[3] - _bb_txt[1]

# Load and resize ko-fi logo to match text height
_logo_h = 22
_kofi_logo_path = pathlib.Path("ko-fi.png")
if _kofi_logo_path.exists():
    _logo = Image.open(_kofi_logo_path).convert("RGBA")
    _logo_w = int(_logo.width * _logo_h / _logo.height)
    _logo = _logo.resize((_logo_w, _logo_h), Image.LANCZOS)
    _gap = 8
    _total_w = _logo_w + _gap + _txt_w
    _row_x = (W - _total_w) // 2
    _row_y = 612
    img.paste(_logo, (_row_x, _row_y + (_txt_h - _logo_h) // 2), _logo)
    draw = ImageDraw.Draw(img)
    draw.text((_row_x + _logo_w + _gap, _row_y), _kofi_txt, fill=YELLOW, font=f_kofi)
else:
    _bb_full = draw.textbbox((0, 0), _kofi_txt, font=f_kofi)
    _kofi_x = (W - (_bb_full[2] - _bb_full[0])) // 2
    draw.text((_kofi_x, 612), _kofi_txt, fill=YELLOW, font=f_kofi)

# ── Save ──────────────────────────────────────────────────────────────────
out = pathlib.Path("splash.png")
img.save(out, "PNG")
print(f"Saved {out}  ({W}x{H}px, {out.stat().st_size // 1024} KB)")
