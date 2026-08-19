"""i18n sync / drift report. Extracts every tr("literal") from VSE.py (the code's
canonical English key set), uses the zh catalog as the reference, and diffs every
other language file against it. Run before a release to catch untranslated strings.

Usage:  python i18n_sync.py            # report
        python i18n_sync.py --gaps     # also write i18n_gaps_<lang>.txt per language
"""
import ast
import pathlib
import sys

BASE = pathlib.Path(__file__).parent
LANGS = ["zh", "es", "de", "fr", "ru", "pt"]
WRITE_GAPS = "--gaps" in sys.argv


def norm(s):
    return s.replace("\\n", "\n")   # match the loader (literal \n -> real newline)


def load_keys(path):
    """Return {en_key: target} from a two-column TSV (target may be '')."""
    out = {}
    if not path.exists():
        return None
    for i, line in enumerate(path.read_text(encoding="utf-8").splitlines()):
        if i == 0 or "\t" not in line:
            continue
        en, tgt = line.split("\t", 1)
        if en:
            out[norm(en)] = norm(tgt)
    return out


def catalog_path(lang):
    p = BASE / "i18n" / f"vse_ui_strings_{lang}.tsv"        # bundled catalogs
    return p if p.exists() else BASE / f"vse_ui_strings_{lang}.tsv"  # root drafts


# 1. code literal tr() keys (the source of truth) via AST
tree = ast.parse((BASE / "VSE.py").read_text(encoding="utf-8"))
code_keys, dynamic = set(), 0
for node in ast.walk(tree):
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "tr":
        a = node.args[0] if node.args else None
        if isinstance(a, ast.Constant) and isinstance(a.value, str):
            code_keys.add(a.value)
        else:
            dynamic += 1

# 2. reference = zh catalog (the most complete language)
zh = load_keys(catalog_path("zh")) or {}
master = set(zh)

print(f"code: {len(code_keys)} literal tr() keys  (+{dynamic} dynamic tr(var) calls — "
      f"their values, e.g. Easy/Erect/mode labels, live in the catalog)")
missing_in_zh = code_keys - master
print(f"zh master: {len(master)} keys, {sum(1 for v in zh.values() if v)} translated, "
      f"{sum(1 for v in zh.values() if not v)} blank")
print(f"code literals NOT in zh master (would fall back to English): {len(missing_in_zh)}")
for k in sorted(missing_in_zh)[:15]:
    print("   -", repr(k)[:90])

print()
print(f"{'lang':<5}{'file':>6}{'covered':>9}{'missing':>9}{'stale':>7}{'blank':>7}{'  %':>6}")
for lang in LANGS:
    lk = load_keys(catalog_path(lang))
    if lk is None:
        print(f"{lang:<5}  (no file)")
        continue
    have = set(lk)
    covered = master & have
    missing = master - have
    stale = have - master
    blank = {k for k in covered if not lk[k]}
    pct = 100 * (len(covered) - len(blank)) / max(1, len(master))
    print(f"{lang:<5}{len(have):>6}{len(covered):>9}{len(missing):>9}{len(stale):>7}"
          f"{len(blank):>7}{pct:>5.0f}%")
    if WRITE_GAPS and (missing or blank):
        need = sorted(missing | blank)
        (BASE / f"i18n_gaps_{lang}.txt").write_text(
            "\n".join(need), encoding="utf-8")
        print(f"       -> wrote i18n_gaps_{lang}.txt ({len(need)} keys need translation)")
