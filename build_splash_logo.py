"""Pre-render the ascii-motion logo (splash_logo.json) into a bundled sprite sheet
using Cascadia Mono (full glyph coverage). The splash cycles the frames as images,
so there's no runtime font dependency on any platform. The 'v1.0' version cells are
left BLANK — the splash stamps the real version live over that spot so it never
goes stale. Re-run only if the source logo changes; version is NOT baked here.

Outputs (bundled):  splash_logo_sheet.png  +  splash_logo_meta.json
"""
import json
import pathlib
from PIL import Image, ImageDraw, ImageFont

BASE = pathlib.Path(__file__).parent
FONT = r"C:\Windows\Fonts\CascadiaMono.ttf"
TARGET_W = 600            # logo width in px at 1x
VER_ROW, VER_X = 1, (69, 74)   # blank the baked "v1.0" cells

data = json.loads((BASE / "splash_logo.json").read_text(encoding="utf-8"))
COLORS, FRAMES = data["colors"], data["frames"]

xs = [c[0] for f in FRAMES for c in f["cells"]]
ys = [c[1] for f in FRAMES for c in f["cells"]]
minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
gw, gh = maxx - minx + 1, maxy - miny + 1

size = 8                  # grow until the grid hits the target width
while ImageFont.truetype(FONT, size + 1).getlength("█") * gw <= TARGET_W:
    size += 1
font = ImageFont.truetype(FONT, size)
cw = font.getlength("█")
asc, desc = font.getmetrics()
ch = asc + desc
FW, FH = int(round(cw * gw)), ch * gh


def render(frame):
    img = Image.new("RGB", (FW, FH), (13, 13, 13))
    d = ImageDraw.Draw(img)
    for c in frame["cells"]:
        x, y, char, ci = c[0], c[1], c[2], c[3]
        if y == VER_ROW and VER_X[0] <= x <= VER_X[1]:
            continue      # leave the version slot blank for the live overlay
        d.text(((x - minx) * cw, (y - miny) * ch), char, font=font, fill=COLORS[ci])
    return img


frames = [render(f) for f in FRAMES]
sheet = Image.new("RGB", (FW, FH * len(frames)), (13, 13, 13))
for i, im in enumerate(frames):
    sheet.paste(im, (0, i * FH))
sheet.save(BASE / "splash_logo_sheet.png")

meta = {
    "n": len(frames), "fw": FW, "fh": FH,
    "durations": [f["duration"] for f in FRAMES],
    "ver_px": [int(round((69 - minx) * cw)), int(round((VER_ROW - miny) * ch))],
    "ver_color": COLORS[7],
}
(BASE / "splash_logo_meta.json").write_text(json.dumps(meta), encoding="utf-8")

frames[0].save(BASE / "_logo_preview_f0.png")
frames[8].save(BASE / "_logo_preview_f8.png")
print(f"font size {size}  cell {cw:.1f}x{ch}  frame {FW}x{FH}  sheet {sheet.size}")
print(f"meta: durations sum {sum(meta['durations'])}ms, ver_px {meta['ver_px']}")
