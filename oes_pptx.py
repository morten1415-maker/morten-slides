"""
Økonomistyrelsen — slide-systemets motor (PowerPoint + PNG-preview).

Denne fil kender INGEN stil. Den kan tegne simple elementer (Rect, Ellipse,
Line, Path, Text, Image) på et 1280 x 720-canvas og skrive dem til .pptx OG
til PNG-preview, så layoutet kan tjekkes uden PowerPoint. Stilen — farver,
titelblokke, slide-typer — ligger i build_pptx.py.

Alt måles i px på et 1280 x 720-canvas (præcis som slides.css).
1 px = 9525 EMU = 0,75 pt, så 1280 px = 13,333".

Tekst ombrydes med de rigtige IBM Plex-metrikker, og tekst der ikke kan være
i sin boks giver en advarsel fra save().
"""
import copy
import math
import os
import re

from PIL import Image as PILImage, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

# ------------------------------------------------------------------ TOKENS
W, H = 1280, 720
EMU_PER_PX = 9525

PAPER = "FFFFFF"
INK = "141414"
INK_SOFT = "262626"     # brødtekst — næsten sort, som på Designmuseet
MUTED = "767676"
LINE = "D6D6D6"
PLACEHOLDER = "ECEBE7"
BRAND = "066B43"        # ØS-grøn — den eneste accent
BRAND_LIGHT = "B3E8CA"  # grøn på mørk flade
SIGNAL = "D9434E"       # kun fald/negative tal

PAD_X = 88
PAD_TOP = 72
OFFSET = 24             # signaturens forskydning
LOGO_W, LOGO_INSET = 124, 32

# Typeskala (px). Brug KUN disse størrelser — så bliver slides ens.
S_DISPLAY = 64     # forside, kapitler, hero, afslutning (grøn versal)
S_TITLE = 36       # slide-overskrift (sort versal)
S_STATEMENT = 44   # et udsagn alene på en slide
S_QUOTE = 40       # citat under en overskrift
S_LEAD = 24        # manchet
S_BODY = 22        # brødtekst og lister
S_SMALL = 18       # tæt tekst i smalle spalter, etiketter i diagrammer
S_CAPTION = 15     # meta-linje, akse-etiketter
S_LABEL = 14       # grøn versal-label (eyebrow)

# Faste mål
CONTENT_TOP = 176  # indhold under en overskrift starter altid her
CONTENT_BOT = 640  # ... og slutter senest her
MEDIA = 470        # billedrammen i billed+tekst-slides (kvadrat)
MEDIA_Y = 125      # rammens top — billedet + flade står lodret centreret

FONTS_DIR = os.path.join(ASSETS, "fonts")
# rolle -> (PowerPoint-familie, fil regular, fil bold)
#   body  = brødtekst (sans)          head = overskrifter (sans, altid fed)
#   label = mono-labels (altid fed)   mono = tal i tabeller (regular/fed)
FONTS = {"body":  ("IBM Plex Sans", "IBMPlexSans-Regular.ttf", "IBMPlexSans-Bold.ttf"),
         "head":  ("IBM Plex Sans", "IBMPlexSans-Bold.ttf", "IBMPlexSans-Bold.ttf"),
         "label": ("IBM Plex Mono", "IBMPlexMono-Bold.ttf", "IBMPlexMono-Bold.ttf"),
         "mono":  ("IBM Plex Mono", "IBMPlexMono-Regular.ttf", "IBMPlexMono-Bold.ttf")}
FORCE_BOLD = {"label", "head"}
# Indlejres i .pptx (OFL — assets/fonts/LICENSE.txt skal følge med)
EMBED = (("IBM Plex Sans", 34, (("regular", "IBMPlexSans-Regular.ttf"), ("bold", "IBMPlexSans-Bold.ttf"))),
         ("IBM Plex Mono", 49, (("regular", "IBMPlexMono-Regular.ttf"), ("bold", "IBMPlexMono-Bold.ttf"))))



def px(v):
    return Emu(int(round(v * EMU_PER_PX)))


# ------------------------------------------------------------------ TEKST-HJÆLPERE
def run(text, f="body", size=S_BODY, color=INK_SOFT, bold=False, caps=False, sp=0.0, italic=False, hl=None):
    """Et tekst-løb. f: 'body' | 'label' | 'head'. size i px. sp = sporing i em."""
    return {"text": text, "f": f, "size": size, "color": color, "bold": bold or f in FORCE_BOLD,
            "caps": caps, "sp": sp, "italic": italic, "hl": hl}   # hl = overstregningsfarve (hex)


def para(*runs, align="l", lh=1.4, before=0):
    """Et afsnit. lh = linjehøjde som faktor af største skriftstørrelse. before i px."""
    return {"runs": list(runs), "align": align, "lh": lh, "before": before}


# ------------------------------------------------------------------ ELEMENTER
class El:
    name = None


class Rect(El):
    def __init__(self, x, y, w, h, fill=None, line=None, lw=1, shadow=False, name=None):
        self.x, self.y, self.w, self.h = x, y, w, h
        self.fill, self.line, self.lw, self.shadow, self.name = fill, line, lw, shadow, name


class Ellipse(Rect):
    pass


class Line(El):
    def __init__(self, x1, y1, x2, y2, color=INK, lw=1, dash=None, head=None, tail=None, name=None):
        self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2
        self.color, self.lw, self.dash, self.head, self.tail, self.name = color, lw, dash, head, tail, name


class Path(El):
    """cmds: [('M',x,y), ('L',x,y), ('C',x1,y1,x2,y2,x,y), ('Z',)] i px."""
    def __init__(self, cmds, color=INK, lw=2, fill=None, dash=None, name=None):
        self.cmds, self.color, self.lw, self.fill, self.dash, self.name = cmds, color, lw, fill, dash, name


class Text(El):
    def __init__(self, x, y, w, h, paras, anchor="t", name=None, title=False):
        self.x, self.y, self.w, self.h = x, y, w, h
        self.paras, self.anchor, self.name, self.title = paras, anchor, name, title


class Image(El):
    """fit: 'cover' (beskær til boksen) eller 'contain' (hele billedet i boksen)."""
    def __init__(self, path, x, y, w, h, fit="cover", focus=(0.5, 0.5), name=None, alt=""):
        self.path, self.x, self.y, self.w, self.h = path, x, y, w, h
        self.fit, self.focus, self.name, self.alt = fit, focus, name, alt


# ------------------------------------------------------------------ TEKST-LAYOUT
_font_cache = {}


def _pil_font(r):
    fam = FONTS[r["f"]]
    fn = fam[2] if r["bold"] else fam[1]
    key = (fn, round(r["size"] * 4))
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(os.path.join(FONTS_DIR, fn), size=max(1, round(r["size"] * 4)))
    return _font_cache[key]


def _measure(text, r):
    t = text.upper() if r["caps"] else text
    f = _pil_font(r)
    w = f.getlength(t) / 4.0
    if r["sp"]:
        w += len(t) * r["sp"] * r["size"]
    return w


def layout(paras, width):
    """Ombryd afsnit til linjer. Returnerer [(para, [lines]), ...] og samlet højde.
    Hver linje: (højde, [(tekst, run, bredde)])"""
    out, total = [], 0.0
    for i, p in enumerate(paras):
        size = max(r["size"] for r in p["runs"]) if p["runs"] else 21
        lh = size * p["lh"]
        tokens = []
        for r in p["runs"]:
            if r["text"] == "\n":
                tokens.append(("\n", r))
                continue
            for tok in re.findall(r"[^ \t\n]+|[ \t\n]+", r["text"]):   # hårdt mellemrum ( ) knækker ikke
                tokens.append((tok, r))
        lines, cur, cur_w = [], [], 0.0
        for tok, r in tokens:
            if tok == "\n":
                lines.append(cur)
                cur, cur_w = [], 0.0
                continue
            tw = _measure(tok, r)
            if tok.isspace():
                if cur:
                    cur.append((tok, r, tw))
                    cur_w += tw
                continue
            if cur and cur_w + tw > width + 0.5:
                while cur and cur[-1][0].isspace():
                    cur_w -= cur[-1][2]
                    cur.pop()
                lines.append(cur)
                cur, cur_w = [], 0.0
            cur.append((tok, r, tw))
            cur_w += tw
        while cur and cur[-1][0].isspace():
            cur.pop()
        lines.append(cur)
        before = p["before"] if i > 0 else 0
        total += before + lh * len(lines)
        out.append((p, lines, lh, before))
    return out, total


# ------------------------------------------------------------------ SLIDE
class Slide:
    def __init__(self, deck, bg=PAPER, logo="green", kind="blank"):
        self.deck, self.bg, self.logo, self.kind = deck, bg, logo, kind
        self.els = []
        self.notes = ""

    def add(self, el):
        self.els.append(el)
        return el

    # genveje
    def rect(self, *a, **k): return self.add(Rect(*a, **k))
    def ellipse(self, *a, **k): return self.add(Ellipse(*a, **k))
    def line(self, *a, **k): return self.add(Line(*a, **k))
    def path(self, *a, **k): return self.add(Path(*a, **k))
    def text(self, *a, **k): return self.add(Text(*a, **k))
    def image(self, *a, **k): return self.add(Image(*a, **k))

    def media(self, path, x, y, w, h, side="right", fit="cover", focus=(0.5, 0.5), alt=""):
        """Billede med den grønne flade forskudt bagved (systemets signatur)."""
        dx = OFFSET if side == "right" else -OFFSET
        self.rect(x + dx, y + OFFSET, w, h, fill=BRAND, name="Grøn flade")
        return self.image(path, x, y, w, h, fit=fit, focus=focus, alt=alt)


# ------------------------------------------------------------------ DECK
class Deck:
    """Grund-deck: en liste af slides + save(). Slide-typerne ligger i build_pptx.Deck."""

    def __init__(self):
        self.slides = []
        self.deck_title = None                  # filens titel (standard: filnavnet)

    def blank(self, bg=PAPER, logo="green"):
        s = Slide(self, bg=bg, logo=logo)
        self.slides.append(s)
        return s

    # ------------------------------------------------------------------ GEM
    def save(self, path, preview_dir=None, embed_fonts=True):
        warnings = []
        prs = Presentation()
        prs.slide_width, prs.slide_height = px(W), px(H)
        _apply_theme(prs)
        for i, s in enumerate(self.slides, 1):
            _write_slide(prs, s, i, warnings)
        if embed_fonts:
            _embed_fonts(prs)
        # filegenskaber: ellers arver filen python-pptx' skabelon (forfatter "Steve Canny", 2013)
        import datetime
        cp = prs.core_properties
        now = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None, microsecond=0)
        cp.title = self.deck_title or os.path.splitext(os.path.basename(path))[0]
        cp.author = cp.last_modified_by = "Økonomistyrelsen"
        cp.created = cp.modified = now
        cp.revision = 1
        cp.comments = ""
        prs.save(path)
        if preview_dir:
            os.makedirs(preview_dir, exist_ok=True)
            for f in os.listdir(preview_dir):
                if f.startswith("slide-") and f.endswith(".png"):
                    os.remove(os.path.join(preview_dir, f))
            for i, s in enumerate(self.slides, 1):
                render(s).save(os.path.join(preview_dir, "slide-%02d.png" % i))
        for w in warnings:
            print("ADVARSEL:", w)
        return warnings


# ------------------------------------------------------------------ PPTX-BACKEND
def _rgb(hexs):
    return RGBColor.from_string(hexs)


def _strip_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _apply_theme(prs):
    """Sæt temaets fonte og farver, så tekst man selv tilføjer i PowerPoint følger systemet."""
    master = prs.slide_masters[0]
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    theme_part = master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    for tag, face in (("majorFont", FONTS["head"][0]), ("minorFont", FONTS["body"][0])):
        el = root.find(".//a:fontScheme/a:%s/a:latin" % tag, ns)
        if el is not None:
            el.set("typeface", face)
    colors = {"dk1": INK, "lt1": PAPER, "dk2": INK_SOFT, "lt2": PLACEHOLDER, "accent1": BRAND,
              "accent2": INK, "accent3": MUTED, "accent4": BRAND_LIGHT, "accent5": SIGNAL,
              "accent6": LINE, "hlink": BRAND, "folHlink": BRAND}
    cs = root.find(".//a:clrScheme", ns)
    cs.set("name", "Økonomistyrelsen")
    for name, val in colors.items():
        el = cs.find("a:%s" % name, ns)
        if el is None:
            continue
        for c in list(el):
            el.remove(c)
        etree.SubElement(el, "{%s}srgbClr" % ns["a"]).set("val", val)
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _set_line(ln, color, lw, dash=None):
    if color is None:
        ln.fill.background()
        return
    ln.color.rgb = _rgb(color)
    ln.width = Pt(lw * 0.75)
    if dash:
        ln.dash_style = dash


def _add_shadow(shape):
    spPr = shape._element.spPr
    eff = etree.SubElement(spPr, qn("a:effectLst"))
    sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad=str(px(18)), dist=str(px(8)), dir="5400000",
                          algn="t", rotWithShape="0")
    clr = etree.SubElement(sh, qn("a:srgbClr"), val="000000")
    etree.SubElement(clr, qn("a:alpha"), val="22000")


def _write_text(shape, el, warnings, idx):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "ctr": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[el.anchor]
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    for a in list(bodyPr):
        bodyPr.remove(a)
    etree.SubElement(bodyPr, qn("a:noAutofit"))
    lines, total = layout(el.paras, el.w)
    widest = max((sum(t[2] for t in ln) for _, ls, _, _ in lines for ln in ls), default=0)
    if widest > el.w + 1:
        warnings.append("slide %d: et ord er bredere end tekstboksen (%.0f > %.0f px): %r"
                        % (idx, widest, el.w, el.paras[0]["runs"][0]["text"][:40] if el.paras[0]["runs"] else ""))
    if total > el.h + 1:
        warnings.append("slide %d: tekst flyder ud af boksen (%.0f > %.0f px): %r"
                        % (idx, total, el.h, el.paras[0]["runs"][0]["text"][:40] if el.paras[0]["runs"] else ""))
    for i, p in enumerate(el.paras):
        P = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        P.alignment = {"l": PP_ALIGN.LEFT, "ctr": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[p["align"]]
        size = max(r["size"] for r in p["runs"]) if p["runs"] else 21
        P.line_spacing = Pt(size * p["lh"] * 0.75)
        if i > 0 and p["before"]:
            P.space_before = Pt(p["before"] * 0.75)
        else:
            P.space_before = Pt(0)
        P.space_after = Pt(0)
        for r in p["runs"]:
            if r["text"] == "\n":
                P.add_line_break()
                continue
            R = P.add_run()
            R.text = r["text"]
            fnt = R.font
            fnt.name = FONTS[r["f"]][0]
            fnt.size = Pt(r["size"] * 0.75)
            fnt.bold = bool(r["bold"])
            fnt.italic = bool(r["italic"])
            fnt.color.rgb = _rgb(r["color"])
            if r.get("hl"):                       # tekst-overstregning (highlight)
                hl = R._r.get_or_add_rPr().makeelement(qn("a:highlight"), {})
                etree.SubElement(hl, qn("a:srgbClr"), val=r["hl"])
                R._r.get_or_add_rPr().find(qn("a:latin")).addprevious(hl)
            rPr = R._r.get_or_add_rPr()
            rPr.set("lang", "da-DK")
            if r["caps"]:
                rPr.set("cap", "all")
            if r["sp"]:
                rPr.set("spc", str(int(round(r["sp"] * r["size"] * 0.75 * 100))))
            for tag in ("a:ea", "a:cs"):
                e = rPr.find(qn(tag))
                if e is None:
                    e = etree.SubElement(rPr, qn(tag))
                e.set("typeface", FONTS[r["f"]][0])


def _write_path(slide, el):
    xs = [c[k] for c in el.cmds for k in range(1, len(c), 2)]
    ys = [c[k] for c in el.cmds for k in range(2, len(c), 2)]
    x0, y0 = min(xs), min(ys)
    w, h = max(max(xs) - x0, 1), max(max(ys) - y0, 1)
    sid = slide.shapes._next_shape_id
    S = lambda v: str(int(round(v * EMU_PER_PX)))
    segs = []
    for c in el.cmds:
        if c[0] == "M":
            segs.append('<a:moveTo><a:pt x="%s" y="%s"/></a:moveTo>' % (S(c[1] - x0), S(c[2] - y0)))
        elif c[0] == "L":
            segs.append('<a:lnTo><a:pt x="%s" y="%s"/></a:lnTo>' % (S(c[1] - x0), S(c[2] - y0)))
        elif c[0] == "C":
            pts = "".join('<a:pt x="%s" y="%s"/>' % (S(c[k] - x0), S(c[k + 1] - y0)) for k in (1, 3, 5))
            segs.append("<a:cubicBezTo>%s</a:cubicBezTo>" % pts)
        elif c[0] == "Z":
            segs.append("<a:close/>")
    fill = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % el.fill) if el.fill else "<a:noFill/>"
    dash = '<a:prstDash val="%s"/>' % el.dash if el.dash else ""
    ln = ('<a:ln w="%d" cap="rnd"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s<a:round/></a:ln>'
          % (int(el.lw * 0.75 * 12700), el.color, dash)) if el.color else "<a:ln><a:noFill/></a:ln>"
    xml = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
           '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
           '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
           '<a:pathLst><a:path w="%s" h="%s">%s</a:path></a:pathLst></a:custGeom>%s%s</p:spPr>'
           '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="da-DK"/></a:p></p:txBody></p:sp>'
           % (sid, el.name or "Form %d" % sid, S(x0), S(y0), S(w), S(h), S(w), S(h), "".join(segs), fill, ln))
    slide.shapes._spTree.append(etree.fromstring(xml))


def _image_crop(path, w, h, fit, focus):
    im = PILImage.open(path)
    iw, ih = im.size
    if fit != "cover":
        return (0, 0, 0, 0), im
    box_r, img_r = w / h, iw / ih
    if img_r > box_r:      # for bred: beskær siderne
        keep = box_r / img_r
        cut = 1 - keep
        l = cut * focus[0]
        return (l, 0, cut - l, 0), im
    keep = img_r / box_r
    cut = 1 - keep
    t = cut * focus[1]
    return (0, t, 0, cut - t), im


def _write_slide(prs, s, idx, warnings):
    # Slides med overskrift bruger "Kun titel"-layoutet, så titlen ligger i
    # PowerPoints titel-pladsholder (navigation, disposition, skærmlæsere).
    has_title = any(isinstance(e, Text) and e.title for e in s.els)
    slide = prs.slides.add_slide(prs.slide_layouts[5 if has_title else 6])
    if s.bg != PAPER:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _rgb(s.bg)
    for el in s.els:
        if isinstance(el, (Rect, Ellipse)):
            shp = slide.shapes.add_shape(MSO_SHAPE.OVAL if isinstance(el, Ellipse) else MSO_SHAPE.RECTANGLE,
                                         px(el.x), px(el.y), px(el.w), px(el.h))
            _strip_style(shp)
            if el.fill:
                shp.fill.solid()
                shp.fill.fore_color.rgb = _rgb(el.fill)
            else:
                shp.fill.background()
            _set_line(shp.line, el.line, el.lw)
            if el.shadow:
                _add_shadow(shp)
            if el.name:
                shp.name = el.name
        elif isinstance(el, Line):
            c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, px(el.x1), px(el.y1), px(el.x2), px(el.y2))
            _strip_style(c)
            c.line.color.rgb = _rgb(el.color)
            c.line.width = Pt(el.lw * 0.75)
            ln = c.line._get_or_add_ln()
            if el.dash:
                etree.SubElement(ln, qn("a:prstDash")).set("val", el.dash)
            for tag, kind in (("a:headEnd", el.head), ("a:tailEnd", el.tail)):
                if kind:
                    etree.SubElement(ln, qn(tag), type=kind, w="med", len="med")
            if el.name:
                c.name = el.name
        elif isinstance(el, Path):
            _write_path(slide, el)
        elif isinstance(el, Text):
            if el.title and has_title and slide.shapes.title.text_frame.text == "":
                tb = slide.shapes.title
                tb.left, tb.top, tb.width, tb.height = px(el.x), px(el.y), px(el.w), px(el.h)
                # Pladsholderen oprettes før alle andre former og ligger derfor nederst.
                # Flyt den op i lagene, så den ligger hvor teksten hører til — ellers
                # dækker fx en sort titelblok over overskriften.
                sp = tb._element
                tree = sp.getparent()
                tree.remove(sp)
                tree.append(sp)
            else:
                tb = slide.shapes.add_textbox(px(el.x), px(el.y), px(el.w), px(el.h))
            _write_text(tb, el, warnings, idx)
            if el.name and not el.title:
                tb.name = el.name
        elif isinstance(el, Image):
            crop, im = _image_crop(el.path, el.w, el.h, el.fit, el.focus)
            x, y, w, h = el.x, el.y, el.w, el.h
            if el.fit == "contain":
                iw, ih = im.size
                sc = min(w / iw, h / ih)
                nw, nh = iw * sc, ih * sc
                x, y, w, h = x + (w - nw) / 2, y + (h - nh) / 2, nw, nh
            pic = slide.shapes.add_picture(el.path, px(x), px(y), px(w), px(h))
            pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = crop
            if el.alt:
                pic._element.nvPicPr.cNvPr.set("descr", el.alt)
            if el.name:
                pic.name = el.name
    if s.logo:
        lp = os.path.join(ASSETS, "oes-logo-%s.png" % s.logo)
        im = PILImage.open(lp)
        lh = LOGO_W * im.size[1] / im.size[0]
        pic = slide.shapes.add_picture(lp, px(W - LOGO_INSET - LOGO_W), px(H - LOGO_INSET - lh), px(LOGO_W), px(lh))
        pic.name = "ØS-logo"
        pic._element.nvPicPr.cNvPr.set("descr", "Økonomistyrelsen")
    if s.notes:
        slide.notes_slide.notes_text_frame.text = s.notes


def _panose(blob):
    import struct
    try:
        num = struct.unpack(">H", blob[4:6])[0]
        for i in range(num):
            off = 12 + i * 16
            if blob[off:off + 4] == b"OS/2":
                t_off = struct.unpack(">I", blob[off + 8:off + 12])[0]
                return blob[t_off + 32:t_off + 42].hex().upper()
    except Exception:
        pass
    return None


def _to_eot(ttf):
    """Pak en TTF ind i en Embedded OpenType-header (version 2.1, ukomprimeret).

    PowerPoint gemmer indlejrede fonte (.fntdata) som EOT, ikke som rå TTF.
    En rå TTF giver "Generel fejl" i dialogen "Installer integrerede skrifttyper"."""
    import io
    import struct
    from fontTools.ttLib import TTFont
    f = TTFont(io.BytesIO(ttf))
    os2, head, name = f["OS/2"], f["head"], f["name"]

    def nm(i):
        rec = name.getName(i, 3, 1, 0x409) or name.getName(i, 1, 0, 0)
        return str(rec).encode("utf-16-le") if rec else b""

    panose = bytes([os2.panose.bFamilyType, os2.panose.bSerifStyle, os2.panose.bWeight,
                    os2.panose.bProportion, os2.panose.bContrast, os2.panose.bStrokeVariation,
                    os2.panose.bArmStyle, os2.panose.bLetterForm, os2.panose.bMidline, os2.panose.bXHeight])
    body = panose + struct.pack("<BBLHH", 1, 1 if os2.fsSelection & 1 else 0, os2.usWeightClass,
                                os2.fsType, 0x504C)
    body += struct.pack("<LLLL", os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3,
                        os2.ulUnicodeRange4)
    body += struct.pack("<LL", getattr(os2, "ulCodePageRange1", 0), getattr(os2, "ulCodePageRange2", 0))
    body += struct.pack("<L", head.checkSumAdjustment) + b"\0" * 16
    for i in (1, 2, 5, 4):                      # familie, stil, version, fulde navn
        s = nm(i)
        body += struct.pack("<HH", 0, len(s)) + s
    body += struct.pack("<HH", 0, 0)            # Padding5 + tom RootString
    size = 16 + len(body) + len(ttf)
    return struct.pack("<LLLL", size, len(ttf), 0x00020001, 0) + body + ttf


def _embed_fonts(prs):
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    pres_part = prs.part
    root = pres_part._element
    for old in root.findall(qn("p:embeddedFontLst")):
        root.remove(old)
    lst = root.makeelement(qn("p:embeddedFontLst"), {})
    n = 0
    for typeface, pitch, slots in EMBED:
        entry = lst.makeelement(qn("p:embeddedFont"), {})
        fel = entry.makeelement(qn("p:font"), {"typeface": typeface, "pitchFamily": str(pitch), "charset": "0"})
        entry.append(fel)
        for slot, fn in slots:
            with open(os.path.join(FONTS_DIR, fn), "rb") as fh:
                blob = fh.read()
            if "panose" not in fel.attrib:
                pan = _panose(blob)
                if pan:
                    fel.set("panose", pan)
            n += 1
            part = Part(PackURI("/ppt/fonts/font%d.fntdata" % n), CT.X_FONTDATA, pres_part.package,
                        _to_eot(blob))
            rId = pres_part.relate_to(part, RT.FONT)
            sl = entry.makeelement(qn("p:" + slot), {})
            sl.set(qn("r:id"), rId)
            entry.append(sl)
        lst.append(entry)
    root.find(qn("p:notesSz")).addnext(lst)
    root.set("embedTrueTypeFonts", "1")
    return n


# ------------------------------------------------------------------ PREVIEW-BACKEND
SS = 4  # supersampling (høj nok til at hinting-afrunding ikke skubber ordene sammen)


def _c(hexs, a=255):
    return tuple(int(hexs[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


def _bez(p0, p1, p2, p3, n=24):
    out = []
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        out.append((mt ** 3 * p0[0] + 3 * mt * mt * t * p1[0] + 3 * mt * t * t * p2[0] + t ** 3 * p3[0],
                    mt ** 3 * p0[1] + 3 * mt * mt * t * p1[1] + 3 * mt * t * t * p2[1] + t ** 3 * p3[1]))
    return out


def _dashed(d, pts, color, width, dash):
    on, off = {"dash": (4, 3), "sysDash": (3, 1), "sysDot": (1, 1), "dot": (1, 2)}.get(dash, (4, 3))
    on, off = on * width, off * width
    for (xa, ya), (xb, yb) in zip(pts, pts[1:]):
        L = math.hypot(xb - xa, yb - ya)
        if L == 0:
            continue
        pos, draw_ = 0.0, True
        while pos < L:
            seg = on if draw_ else off
            end = min(pos + seg, L)
            if draw_:
                d.line([(xa + (xb - xa) * pos / L, ya + (yb - ya) * pos / L),
                        (xa + (xb - xa) * end / L, ya + (yb - ya) * end / L)], fill=color, width=max(1, int(width)))
            pos, draw_ = end, not draw_


def _arrow(d, x1, y1, x2, y2, kind, color, lw):
    ang = math.atan2(y2 - y1, x2 - x1)
    L = max(8 * SS, lw * 3)
    if kind in ("triangle", "arrow"):
        a1, a2 = ang + math.radians(150), ang - math.radians(150)
        d.polygon([(x2, y2), (x2 + L * math.cos(a1), y2 + L * math.sin(a1)),
                   (x2 + L * math.cos(a2), y2 + L * math.sin(a2))], fill=color)
    elif kind == "oval":
        r = L / 2.4
        d.ellipse([x2 - r, y2 - r, x2 + r, y2 + r], fill=color)


def render(s):
    im = PILImage.new("RGBA", (W * SS, H * SS), _c(s.bg))
    for el in s.els:
        if isinstance(el, (Rect, Ellipse)):
            box = [el.x * SS, el.y * SS, (el.x + el.w) * SS, (el.y + el.h) * SS]
            if el.shadow:
                sh = PILImage.new("RGBA", im.size, (0, 0, 0, 0))
                ImageDraw.Draw(sh).rectangle([box[0], box[1] + 8 * SS, box[2], box[3] + 8 * SS], fill=(0, 0, 0, 60))
                im.alpha_composite(sh.filter(ImageFilter.GaussianBlur(9 * SS)))
            d = ImageDraw.Draw(im)
            fn = d.ellipse if isinstance(el, Ellipse) else d.rectangle
            fn(box, fill=_c(el.fill) if el.fill else None,
               outline=_c(el.line) if el.line else None, width=int(el.lw * SS) if el.line else 0)
        elif isinstance(el, Line):
            d = ImageDraw.Draw(im)
            pts = [(el.x1 * SS, el.y1 * SS), (el.x2 * SS, el.y2 * SS)]
            if el.dash:
                _dashed(d, pts, _c(el.color), el.lw * SS, el.dash)
            else:
                d.line(pts, fill=_c(el.color), width=max(1, int(el.lw * SS)))
            if el.tail:
                _arrow(d, pts[0][0], pts[0][1], pts[1][0], pts[1][1], el.tail, _c(el.color), el.lw * SS)
            if el.head:
                _arrow(d, pts[1][0], pts[1][1], pts[0][0], pts[0][1], el.head, _c(el.color), el.lw * SS)
        elif isinstance(el, Path):
            d = ImageDraw.Draw(im)
            pts, cur, polys = [], None, []
            for c in el.cmds:
                if c[0] == "M":
                    if pts:
                        polys.append(pts)
                    cur = (c[1] * SS, c[2] * SS)
                    pts = [cur]
                elif c[0] == "L":
                    cur = (c[1] * SS, c[2] * SS)
                    pts.append(cur)
                elif c[0] == "C":
                    pts += _bez(cur, (c[1] * SS, c[2] * SS), (c[3] * SS, c[4] * SS), (c[5] * SS, c[6] * SS))
                    cur = pts[-1]
                elif c[0] == "Z":
                    pts.append(pts[0])
            if pts:
                polys.append(pts)
            for pl in polys:
                if el.fill:
                    d.polygon(pl, fill=_c(el.fill))
                if el.color:
                    if el.dash:
                        _dashed(d, pl, _c(el.color), el.lw * SS, el.dash)
                    else:
                        d.line(pl, fill=_c(el.color), width=max(1, int(el.lw * SS)), joint="curve")
        elif isinstance(el, Image):
            src = PILImage.open(el.path)
            src.seek(0)
            src = src.convert("RGBA")
            iw, ih = src.size
            if el.fit == "cover":
                (l, t, r, b), _ = _image_crop(el.path, el.w, el.h, "cover", el.focus)
                src = src.crop((int(l * iw), int(t * ih), int(iw - r * iw), int(ih - b * ih)))
                x, y, w, h = el.x, el.y, el.w, el.h
            else:
                sc = min(el.w / iw, el.h / ih)
                w, h = iw * sc, ih * sc
                x, y = el.x + (el.w - w) / 2, el.y + (el.h - h) / 2
            src = src.resize((max(1, int(w * SS)), max(1, int(h * SS))), PILImage.LANCZOS)
            im.alpha_composite(src, (int(x * SS), int(y * SS)))
        elif isinstance(el, Text):
            d = ImageDraw.Draw(im)
            lines, total = layout(el.paras, el.w)
            y = el.y
            if el.anchor == "ctr":
                y = el.y + (el.h - total) / 2
            elif el.anchor == "b":
                y = el.y + el.h - total
            for p, ls, lh, before in lines:
                y += before
                for ln in ls:
                    lw_ = sum(t[2] for t in ln)
                    x = el.x
                    if p["align"] == "ctr":
                        x = el.x + (el.w - lw_) / 2
                    elif p["align"] == "r":
                        x = el.x + el.w - lw_
                    base = y + lh * 0.80
                    for tok, r, tw in ln:
                        if r.get("hl"):
                            d.rectangle([x * SS, (base - r["size"] * 0.95) * SS, (x + tw) * SS,
                                         (base + r["size"] * 0.28) * SS], fill=_c(r["hl"]))
                        t = tok.upper() if r["caps"] else tok
                        fam = FONTS[r["f"]]
                        f = ImageFont.truetype(os.path.join(FONTS_DIR, fam[2] if r["bold"] else fam[1]),
                                               size=max(1, int(r["size"] * SS)))
                        if r["sp"]:
                            cx = x
                            for ch in t:
                                d.text((cx * SS, base * SS), ch, font=f, fill=_c(r["color"]), anchor="ls")
                                cx += f.getlength(ch) / SS + r["sp"] * r["size"]
                        else:
                            d.text((x * SS, base * SS), t, font=f, fill=_c(r["color"]), anchor="ls")
                        x += tw
                    y += lh
    if s.logo:
        lp = PILImage.open(os.path.join(ASSETS, "oes-logo-%s.png" % s.logo)).convert("RGBA")
        lh = LOGO_W * lp.size[1] / lp.size[0]
        lp = lp.resize((int(LOGO_W * SS), int(lh * SS)), PILImage.LANCZOS)
        im.alpha_composite(lp, (int((W - LOGO_INSET - LOGO_W) * SS), int((H - LOGO_INSET - lh) * SS)))
    return im.convert("RGB").resize((W, H), PILImage.LANCZOS)
