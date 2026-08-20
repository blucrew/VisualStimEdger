"""Pull COLORS + FRAMES out of the ascii-motion .tsx export, report stats, and
render the first/last frame as plain text so we can see the logo before wiring it
into the splash. raw_decode handles brackets-that-are-chars inside the cell data."""
import json
import pathlib
import sys

SRC = pathlib.Path(r"G:\Downloads\vse-splash2.tsx")
for _a in sys.argv[1:]:
    if not _a.startswith("--"):
        SRC = pathlib.Path(_a)
txt = SRC.read_text(encoding="utf-8")
dec = json.JSONDecoder()


def grab(marker):
    eq = txt.index("=", txt.index(marker))   # past the `string[]` / `Frame[]` annotation
    i = txt.index("[", eq)
    obj, _ = dec.raw_decode(txt, i)
    return obj


COLORS = grab("const COLORS")
FRAMES = grab("const FRAMES")

W = max(c[0] for f in FRAMES for c in f["cells"]) + 1
H = max(c[1] for f in FRAMES for c in f["cells"]) + 1
used = sorted({c[3] for f in FRAMES for c in f["cells"]})
has_bg = any(len(c) > 4 for f in FRAMES for c in f["cells"])

print(f"palette: {len(COLORS)} colors")
print(f"frames: {len(FRAMES)}   durations(ms): {[f['duration'] for f in FRAMES]}")
print(f"grid: {W} wide x {H} tall   cells/frame: {[len(f['cells']) for f in FRAMES]}")
print(f"fg color indices used: {used}")
print(f"per-cell background colors present: {has_bg}")
print("colors used ->", {i: COLORS[i] for i in used})


def render(frame):
    g = [[" "] * W for _ in range(H)]
    for c in frame["cells"]:
        x, y, ch = c[0], c[1], c[2]
        if 0 <= y < H and 0 <= x < W:
            g[y][x] = ch
    return "\n".join("".join(r).rstrip() for r in g)


for idx in ([0] if len(FRAMES) == 1 else [0, len(FRAMES) // 2, len(FRAMES) - 1]):
    print(f"\n===== FRAME {idx} (dur {FRAMES[idx]['duration']}ms) =====")
    print(render(FRAMES[idx]))

if "--save" in sys.argv:
    out = pathlib.Path(__file__).parent / "splash_logo.json"
    out.write_text(json.dumps({"colors": COLORS, "frames": FRAMES},
                              ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"\nsaved {out.name} ({out.stat().st_size} bytes)")
