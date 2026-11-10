# -*- coding: utf-8 -*-
"""
Regressionstjek af slide-systemet
=================================
De faste mål kan skride uden at nogen ser det. Dette script er værnet: det
bygger galleriet og kontrollerer de invarianter der ellers skal rettes i
hånden deck efter deck.

    python verify.py                # byg example.pptx og tjek alt
    python verify.py mit-deck.pptx  # tjek en bestemt fil (gruppe 1, 2 og 5)

Exit-kode 0 = alt bestået, 1 = mindst én fejl.

Syv grupper:
  0. Referencemål       — at konstanterne stadig er dem systemet er bygget på
  1. Geometri i .pptx   — titelblok, logo, kilde-linje, den grønne flade bag billeder
  2. Font-indlejring    — at IBM Plex faktisk ligger i filen og er hægtet rigtigt op
  3. HTML/PPTX-paritet  — at slides.css og build_pptx.py taler om samme mål og farver
  4. Galleri-dækning    — at hver slide-type vises i begge gallerier
  5. Lag og kontrast    — ingen tekst dækket af en form, ingen tekst under WCAG AA
  6. Tekst og oprydning — ingen tekst flyder ud af sin boks, ingen rester fra den gamle stil
"""
from __future__ import annotations

import os
import re
import runpy
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(HERE, "assets", "fonts")
EMU_PER_PX = 9525.0          # 1280 px = 13,333"

_failures: list[str] = []


def check(label, cond, detail=""):
    if not cond:
        _failures.append(label)
    print("  %s %-56s %s" % ("OK  " if cond else "FEJL", label, detail))
    return bool(cond)


def near(a, b, tol):
    return abs(a - b) <= tol


def head(title):
    print("\n%s\n%s" % (title, "-" * len(title)))


# =================================================================== 0. REFERENCEMÅL
# Systemets sandhed, i px på 1280 x 720-canvas'et. Tallene står HER — ikke som
# import fra build_pptx — netop så et uheldigt tryk på en konstant bliver fanget
# i stedet for stiltiende at flytte forventningen med sig. Ændrer du bevidst et
# mål, skal det ændres her, i build_pptx.py/oes_pptx.py OG i :root i slides.css.
REFERENCE_PX = {
    "W": 1280, "H": 720,
    "PAD_X": 88,             # sikker margin i siderne
    "TITLE_Y": 64,           # titelblokkens top
    "TITLE_SIZE": 31,        # tekst i titelblokken
    "TITLE_SHADOW": 7,       # titelblokkens grønne skygge
    "CONTENT_TOP": 176,      # indholdet starter
    "CONTENT_BOT": 640,      # indholdet slutter
    "SOURCE_BOTTOM": 682,    # kilde-linjens underkant
    "LOGO_W": 124,           # logo, bredde
    "LOGO_INSET": 32,        # logo, ens luft til højre og bund
    "OFFSET": 24,            # den grønne flade bag billeder
    "MEDIA": 470,            # billedrammen i billede+tekst
    "MEDIA_Y": 125,          # billedrammens top
    "BORDER": 3,             # kantbredde
    "SHADOW": 8,             # hård skygge
    "GAP": 40,               # luft mellem spalter
}
REFERENCE_COLORS = {
    "GROUND": "F4F4F2", "BLACK": "000000", "BRAND": "066B43", "HILITE": "B3E8CA",
    "LAKS": "ED5E66", "LAKS_TEXT": "C93A44", "DIM": "6E6E6E", "INK": "141414",
}


def check_reference_measurements():
    import build_pptx as bp
    for name, want in sorted(REFERENCE_PX.items()):
        got = getattr(bp, name, None)
        check("%-14s = %s px" % (name, want), got == want, "" if got == want else "er %r" % (got,))
    for name, want in sorted(REFERENCE_COLORS.items()):
        got = getattr(bp, name, None)
        check("%-14s = #%s" % (name, want), got == want, "" if got == want else "er %r" % (got,))


# ======================================================================= 1. GEOMETRI
def check_pptx_geometry(path):
    from pptx import Presentation
    R = REFERENCE_PX

    def box(sh):
        return tuple(round(v / EMU_PER_PX, 1) for v in (sh.left, sh.top, sh.width, sh.height))

    prs = Presentation(path)
    titles = logos = sources = planes = 0
    bad = []
    for i, s in enumerate(prs.slides, 1):
        shapes = list(s.shapes)
        by_name = {}
        for sh in shapes:
            by_name.setdefault(sh.name, []).append(box(sh))

        # titelblok: sort skilt øverst til venstre + grøn skygge forskudt TITLE_SHADOW
        for (x, y, w, h) in by_name.get("Titelblok", []):
            if h > 90:                      # de store blokke (forside, kapitel, tak) står midt på
                continue
            titles += 1
            if not (near(x, R["PAD_X"], 0.6) and near(y, R["TITLE_Y"], 0.6)):
                bad.append("slide %d: titelblok på (%s, %s), ventet (%s, %s)"
                           % (i, x, y, R["PAD_X"], R["TITLE_Y"]))
            sh_ok = any(near(sx - x, R["TITLE_SHADOW"], 0.6) and near(sy - y, R["TITLE_SHADOW"], 0.6)
                        and near(sw, w, 0.6) for (sx, sy, sw, _) in by_name.get("Titelblok, skygge", []))
            if not sh_ok:
                bad.append("slide %d: titelblok uden grøn skygge på %s px" % (i, R["TITLE_SHADOW"]))

        # logo: ens luft til højre og bund
        for (x, y, w, h) in by_name.get("ØS-logo", []):
            logos += 1
            r, b = R["W"] - (x + w), R["H"] - (y + h)
            if not (near(r, R["LOGO_INSET"], 0.6) and near(b, R["LOGO_INSET"], 0.6)
                    and near(w, R["LOGO_W"], 0.6)):
                bad.append("slide %d: logo-luft h=%.1f b=%.1f, bredde %.1f" % (i, r, b, w))

        # kilde-linje
        for (x, y, w, h) in by_name.get("Kilde", []):
            sources += 1
            if not (near(x, R["PAD_X"], 0.6) and near(y + h, R["SOURCE_BOTTOM"], 0.6)):
                bad.append("slide %d: kilde-linje på (%s, %s)" % (i, x, y + h))

        # grøn flade: præcis OFFSET forskudt bag et billede/en pladsholder i samme størrelse
        others = [box(sh) for sh in shapes if sh.name != "Grøn flade"]
        for (x, y, w, h) in by_name.get("Grøn flade", []):
            planes += 1
            ok = any(near(abs(x - ox), R["OFFSET"], 0.6) and near(y - oy, R["OFFSET"], 0.6)
                     and near(w, ow, 0.6) and near(h, oh, 0.6) for (ox, oy, ow, oh) in others)
            if not ok:
                bad.append("slide %d: grøn flade uden billede forskudt %s px" % (i, R["OFFSET"]))

    check("titelblokke på (%s, %s) px med grøn skygge" % (R["PAD_X"], R["TITLE_Y"]),
          titles > 0 and not any("titelblok" in m for m in bad), "%d fundet" % titles)
    check("logo i hjørnet med ens luft (%s px)" % R["LOGO_INSET"],
          logos > 0 and not any("logo" in m for m in bad), "%d slides" % logos)
    check("kilde-linjer slutter på %s px" % R["SOURCE_BOTTOM"], not any("kilde" in m for m in bad),
          "%d fundet" % sources)
    check("grøn flade forskudt %s px bag billeder" % R["OFFSET"], not any("flade" in m for m in bad),
          "%d fundet" % planes)
    check("ingen geometri-afvigelser i alt", not bad, "%d afvigelser" % len(bad))
    for m in bad:
        print("       - " + m)
    return len(prs.slides._sldIdLst)


# ================================================================ 2. FONT-INDLEJRING
def check_font_embedding(path):
    from lxml import etree

    NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
          "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
    # CT_Presentation-rækkefølgen ifølge ECMA-376 — embeddedFontLst SKAL ligge her
    ORDER = ["sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "sldIdLst",
             "sldSz", "notesSz", "smartTags", "embeddedFontLst", "custShowLst",
             "photoAlbum", "custDataLst", "kinsoku", "defaultTextStyle",
             "modifyVerifier", "extLst"]

    z = zipfile.ZipFile(path)
    names = z.namelist()

    ct = etree.fromstring(z.read("[Content_Types].xml"))
    defaults = {d.get("Extension").lower(): d.get("ContentType")
                for d in ct.findall("ct:Default", NS)}
    check("[Content_Types].xml: Default for fntdata",
          defaults.get("fntdata") == "application/x-fontdata",
          defaults.get("fntdata") or "mangler")

    parts = sorted(n for n in names if n.startswith("ppt/fonts/"))
    src = {}
    if os.path.isdir(FONTS_DIR):
        for f in sorted(os.listdir(FONTS_DIR)):
            if f.lower().endswith((".ttf", ".otf")):
                blob = open(os.path.join(FONTS_DIR, f), "rb").read()
                src[f] = blob
    check("fontfiler i assets/fonts/", len(src) >= 4, "%d filer" % len(src))
    check("OFL-licensen ligger hos fonten",
          os.path.exists(os.path.join(FONTS_DIR, "LICENSE.txt")),
          "assets/fonts/LICENSE.txt")
    check("font-dele indlejret i pakken", len(parts) == len(src) and parts,
          "%d dele" % len(parts))
    # PowerPoint gemmer indlejrede fonte som EOT: en header efterfulgt af den rå TTF.
    # (En rå TTF som .fntdata giver "Generel fejl" i PowerPoints font-dialog.)
    def source_of(blob):
        for f, ttf in src.items():
            if blob.endswith(ttf) and len(blob) > len(ttf) and blob[8:12] == bytes([1, 0, 2, 0]):
                return f
        return "?"
    matched = [source_of(z.read(n)) for n in parts]
    check("hver del er en EOT-pakket kopi af sin kildefil", "?" not in matched and len(set(matched)) == len(parts),
          ", ".join(sorted(matched)))

    rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    FONT_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
    font_rels = {r.get("Id"): r.get("Target")
                 for r in rels.findall("pr:Relationship", NS) if r.get("Type") == FONT_RT}
    check("font-relationships", len(font_rels) == len(parts), " ".join(sorted(font_rels)))

    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    kids = [etree.QName(k).localname for k in pres if isinstance(k.tag, str)]
    check("p:embeddedFontLst findes", "embeddedFontLst" in kids)
    idx = [ORDER.index(k) for k in kids if k in ORDER]
    check("elementrækkefølge følger ECMA-376", idx == sorted(idx),
          " -> ".join(k for k in kids if k in ORDER))

    lst = pres.find("p:embeddedFontLst", NS)
    for ef in (lst.findall("p:embeddedFont", NS) if lst is not None else []):
        font = ef.find("p:font", NS)
        tf = font.get("typeface")
        sub = [etree.QName(c).localname for c in ef]
        check("%s: barn-rækkefølge font,regular,bold" % tf,
              sub[:1] == ["font"] and sub[1:] == sorted(sub[1:], key=
                  lambda s: ["regular", "bold", "italic", "boldItalic"].index(s)),
              ",".join(sub))
        check("%s: panose + pitchFamily" % tf,
              bool(font.get("panose")) and bool(font.get("pitchFamily")),
              "panose=%s pitch=%s" % (font.get("panose"), font.get("pitchFamily")))
        for slot in sub[1:]:
            rid = ef.find("p:" + slot, NS).get("{%s}id" % NS["r"])
            tgt = font_rels.get(rid, "")
            check("%s: %s -> %s" % (tf, slot, tgt),
                  rid in font_rels and ("ppt/" + tgt.lstrip("/")) in names, rid)


# ============================================================== 3. HTML/PPTX-PARITET
def _strip_comments(css):
    """Fjern /* ... */ — ellers bliver kommentartekst læst som selector."""
    return re.sub(r"/\*.*?\*/", "", css, flags=re.S)


def _css_block(css, selector):
    """Returnér erklæringerne i den FØRSTE regel med præcis denne selector."""
    css = _strip_comments(css)
    for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", css):
        # alt før sidste ';' er @import/@charset-sætninger, ikke selector
        raw = m.group(1).split(";")[-1]
        sels = [" ".join(s.split()) for s in raw.split(",")]
        if selector in sels:
            return m.group(2)
    return None


def _css_px(css, selector, prop):
    block = _css_block(css, selector)
    if block is None:
        return None
    m = re.search(r"(?m)(?:^|;)\s*%s\s*:\s*([^;]+)" % re.escape(prop), block)
    if not m:
        return None
    nums = re.findall(r"(-?\d+(?:\.\d+)?)px", m.group(1))
    return float(nums[0]) if nums else None


def _css_num(css, selector, prop):
    """Som _css_px, men til enhedsløse værdier (fx line-height: 1.05)."""
    block = _css_block(css, selector)
    if block is None:
        return None
    m = re.search(r"(?m)(?:^|;)\s*%s\s*:\s*([^;]+)" % re.escape(prop), block)
    if not m:
        return None
    n = re.findall(r"-?\d+(?:\.\d+)?", m.group(1))
    return float(n[0]) if n else None


def _css_color(css, name):
    m = re.search(r"--%s\s*:\s*#([0-9a-fA-F]{6})" % re.escape(name), _strip_comments(css))
    return m.group(1).upper() if m else None


def check_parity():
    import build_pptx as bp
    css = open(os.path.join(HERE, "slides.css"), encoding="utf-8").read()

    pairs = [
        # (label, css-værdi, pptx-px)
        ("sidemargin (--pad-x)", _css_px(css, ":root", "--pad-x"), bp.PAD_X),
        ("titelblokkens top (--title-y)", _css_px(css, ":root", "--title-y"), bp.TITLE_Y),
        ("indhold starter (--content-top)", _css_px(css, ":root", "--content-top"), bp.CONTENT_TOP),
        ("indhold slutter (--pad-bottom)", _css_px(css, ":root", "--pad-bottom"), bp.H - bp.CONTENT_BOT),
        ("kantbredde (--border)", _css_px(css, ":root", "--border"), bp.BORDER),
        ("hård skygge (--shadow)", _css_px(css, ":root", "--shadow"), bp.SHADOW),
        ("grøn flade (--offset)", _css_px(css, ":root", "--offset"), bp.OFFSET),
        ("billedramme (--media)", _css_px(css, ":root", "--media"), bp.MEDIA),
        ("spalteluft (--gap)", _css_px(css, ":root", "--gap"), bp.GAP),
        ("titelblok: skriftstørrelse", _css_px(css, ".slide-title", "font-size"), bp.TITLE_SIZE),
        ("titelblok: grøn skygge", _css_px(css, ".slide-title", "box-shadow"), bp.TITLE_SHADOW),
        ("logo: luft til højre", _css_px(css, ".slide::after", "right"), bp.LOGO_INSET),
        ("logo: luft til bunden", _css_px(css, ".slide::after", "bottom"), bp.LOGO_INSET),
        ("logo: bredde", _css_px(css, ".slide::after", "width"), bp.LOGO_W),
        ("billede+tekst: rammens top", _css_px(css, ".slide--image-text", "padding-top"), bp.MEDIA_Y),
    ]
    for label, got, want in pairs:
        if got is None:
            check("paritet: %s" % label, False, "kunne ikke læses i slides.css")
            continue
        check("paritet: %s" % label, near(got, want, 0.5), "css %gpx  vs  pptx %gpx" % (got, want))

    # Titelblokkens højde + luften under den skal lande præcis på --content-top.
    fs = _css_px(css, ".slide-title", "font-size")
    lh = _css_num(css, ".slide-title", "line-height")
    pad = _css_px(css, ".slide-title", "padding")
    mb = _css_px(css, ".slide-title", "margin-bottom")
    ty, ct = _css_px(css, ":root", "--title-y"), _css_px(css, ":root", "--content-top")
    if None in (fs, lh, pad, mb, ty, ct):
        check("paritet: titelblok + luft = --content-top", False, "kunne ikke læse målene i slides.css")
    else:
        got = ty + fs * lh + 2 * pad + mb
        check("paritet: titelblok + luft = --content-top", near(got, ct, 1), "%.1fpx  vs  %gpx" % (got, ct))
        check("paritet: titelblokkens luft (lodret)", near(pad, bp.TITLE_PAD[1], 0.5),
              "css %gpx  vs  pptx %gpx" % (pad, bp.TITLE_PAD[1]))

    for css_name, py_name in (("ground", "GROUND"), ("brand", "BRAND"), ("hilite", "HILITE"), ("laks", "LAKS"),
                              ("laks-text", "LAKS_TEXT"), ("dim", "DIM"), ("ink", "INK"), ("sticky", "STICKY"),
                              ("empty", "EMPTY"), ("rule", "RULE")):
        got = _css_color(css, css_name)
        check("paritet: farven --%s" % css_name, got == getattr(bp, py_name),
              "css #%s  vs  pptx #%s" % (got, getattr(bp, py_name)))


# ============================================================== 4. GALLERI-DÆKNING
# Deck-metode -> den slide-klasse der viser samme type i template.html
HTML_CLASS = {
    "title": "slide--title", "hero": "slide--hero", "agenda": "slide--agenda", "section": "slide--section",
    "content": "slide--content", "split": "slide--split", "cards": "slide--cards",
    "comparison": "slide--compare", "timeline": "slide--timeline", "process": "slide--process",
    "statement": "slide--statement", "testimonial": "slide--testimonial", "image_text": "slide--image-text",
    "media_text": "slide--image-text", "image_grid": "slide--image-grid", "cover": "slide--cover",
    "image_full": "slide--image-full", "bignum": "slide--bignum", "stats": "slide--stats",
    "bars": "slide--bars", "table": "slide--table", "closing": "slide--closing",
}


def check_gallery_coverage(n_pptx):
    src = open(os.path.join(HERE, "build_pptx.py"), encoding="utf-8").read()
    example = open(os.path.join(HERE, "deck.example.py"), encoding="utf-8").read()
    tpl = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    css = open(os.path.join(HERE, "slides.css"), encoding="utf-8").read()
    methods = set(re.findall(r"(?m)^    def ([a-z][a-z_0-9]*)\(self", src)) - {"save", "blank"}
    used = set(re.findall(r"deck\.([a-z_0-9]+)\(", example))
    missing = sorted(methods - used)
    check("PPTX-galleriet viser hver slide-type", not missing,
          "%d typer%s" % (len(methods), "" if not missing else " — mangler: " + ", ".join(missing)))
    unmapped = sorted(methods - set(HTML_CLASS))
    check("hver slide-type har en HTML-klasse i verify.py", not unmapped, ", ".join(unmapped))
    no_html = sorted(m for m in methods & set(HTML_CLASS) if HTML_CLASS[m] not in tpl)
    check("HTML-galleriet viser hver slide-type", not no_html, ", ".join(no_html))
    no_css = sorted(c for c in set(HTML_CLASS.values()) if c != "slide--content" and "." + c not in css)
    check("slides.css har regler for hver slide-type", not no_css, ", ".join(no_css))
    n_html = len(re.findall(r'<section class="slide ', tpl))
    check("de to gallerier har lige mange slides", n_html == n_pptx, "html %d  vs  pptx %d" % (n_html, n_pptx))


# ================================================================ 5. LAG OG KONTRAST
def check_layers_and_contrast(path):
    import tjek_pptx
    probs = tjek_pptx.check(path)
    covered = [p for p in probs if "dækket" in p]
    contrast = [p for p in probs if "kontrast" in p]
    check("ingen tekst er dækket af en form ovenpå", not covered, "%d fundet" % len(covered))
    check("al tekst holder WCAG AA-kontrast", not contrast, "%d fundet" % len(contrast))
    for p in probs:
        print("       - " + p)


# ============================================================ 6. TEKST OG OPRYDNING
def check_text_and_cleanup(warnings):
    check("ingen tekst flyder ud af sin boks i galleriet", not warnings, "%d advarsler" % len(warnings))
    for w in warnings:
        print("       - " + w)
    tpl = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    css = _strip_comments(open(os.path.join(HERE, "slides.css"), encoding="utf-8").read())
    example = open(os.path.join(HERE, "deck.example.py"), encoding="utf-8").read()
    over_title = re.findall(r'<p class="(?:kicker|eyebrow)">[^<]*</p>\s*<h[12] class="slide-(?:display|title)"', tpl)
    check("ingen label over en titelblok i HTML-galleriet", not over_title, "%d fundet" % len(over_title))
    check("galleriet linker slides.css (ikke det udfasede shim)",
          'href="slides.css"' in tpl and "slides-oes.css" not in tpl)
    check("eksemplet bruger ikke kicker= eller theme=", "kicker=" not in example and "theme=" not in example)
    check("slides.css henter ikke fonte udefra", "fonts.googleapis.com" not in css)
    # hårde skygger: tredje længde i box-shadow (blur) skal være 0
    blurs = [m for m in re.findall(r"box-shadow:\s*([^;]+);", css)
             if len(m.split()) >= 3 and re.match(r"^-?[\d.]+px$", m.split()[2]) and m.split()[2] != "0px"]
    check("kun hårde skygger i slides.css (ingen blur)", not blurs, "; ".join(blurs))
    many = [m for m in re.findall(r'<section class="slide [^"]*"[^>]*>(.*?)</section>', tpl, flags=re.S)
            if len(re.findall(r'class="slide-title"', m)) > 1]
    check("højst én titelblok pr. slide i HTML-galleriet", not many, "%d slides" % len(many))


# ========================================================================== MAIN
def main():
    sys.stdout.reconfigure(encoding="utf-8")
    print(__doc__.strip().split("\n\n")[0])
    sys.path.insert(0, HERE)
    own_file = len(sys.argv) > 1
    path = sys.argv[1] if own_file else os.path.join(HERE, "example.pptx")
    warnings = []

    if not own_file:
        head("Bygger galleriet")
        try:
            ns = runpy.run_path(os.path.join(HERE, "deck.example.py"), run_name="verify")
            deck = ns["deck"]
            deck.save(path, preview_dir=os.path.join(HERE, "preview"))
            warnings = deck.warnings
            print("  example.pptx — %d slides" % len(deck.slides))
        except Exception as e:                       # noqa: BLE001
            print("\nFEJL: deck.example.py fejlede — kan ikke fortsætte: %r" % (e,))
            return 1
    elif not os.path.exists(path):
        print("\nFEJL: %s findes ikke." % path)
        return 1

    if not own_file:
        head("0. Referencemål")
        check_reference_measurements()

    head("1. Geometri i %s" % os.path.basename(path))
    n = check_pptx_geometry(path)
    print("     (%d slides)" % n)

    head("2. Font-indlejring")
    check_font_embedding(path)

    if not own_file:
        head("3. HTML/PPTX-paritet")
        check_parity()

        head("4. Galleri-dækning")
        check_gallery_coverage(n)

    head("5. Lag og kontrast")
    check_layers_and_contrast(path)

    if not own_file:
        head("6. Tekst og oprydning")
        check_text_and_cleanup(warnings)

    print()
    if _failures:
        print("%d TJEK FEJLEDE:" % len(_failures))
        for f in _failures:
            print("  - " + f)
        return 1
    print("ALT BESTÅET")
    return 0


if __name__ == "__main__":
    sys.exit(main())
