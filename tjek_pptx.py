"""
Tjekker en færdig .pptx for fejl, man ikke kan se i preview'et:
  1. Tekst der er dækket af en form, der ligger ovenpå (forkert lag-rækkefølge).
  2. Tekst med for lav kontrast mod den form/baggrund, den står på.

Læser selve filen (lag-rækkefølgen i spTree), ikke generatorens elementliste.

Brug:  python tjek_pptx.py "min præsentation.pptx"
Exit-kode 1 hvis der er fejl.
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn


def _hex(el):
    """Første faste farve (srgbClr) direkte under et fyld-element."""
    if el is None:
        return None
    c = el.find(qn("a:srgbClr"))
    return c.get("val") if c is not None else None


def _fill(sp):
    """Fyldfarve for en form, 'IMG' for billeder, None hvis gennemsigtig."""
    if sp.tag == qn("p:pic"):
        return "IMG"
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        return None
    return _hex(spPr.find(qn("a:solidFill")))


def _box(sp):
    off = sp.find(".//" + qn("a:off"))
    ext = sp.find(".//" + qn("a:ext"))
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")), int(off.get("y"))
    return x, y, x + int(ext.get("cx")), y + int(ext.get("cy"))


def _text(sp):
    return "".join(t.text or "" for t in sp.iter(qn("a:t"))).strip()


def _runs(sp):
    """(farve, størrelse i pt, fed) for hvert tekstløb med egen farve."""
    out = []
    for r in sp.iter(qn("a:rPr")):
        c = _hex(r.find(qn("a:solidFill")))
        if c:
            out.append((c, int(r.get("sz", "1800")) / 100, r.get("b") == "1"))
    return out


def _lum(hexs):
    r, g, b = (int(hexs[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast(a, b):
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def _overlap(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return max(w, 0) * max(h, 0)


def check(path):
    prs = Presentation(path)
    problems = []
    for n, slide in enumerate(prs.slides, 1):
        bg = "FFFFFF"
        bgfill = slide.background._cSld.find(".//" + qn("p:bg"))
        if bgfill is not None:
            bg = _hex(bgfill.find(".//" + qn("a:solidFill"))) or bg
        shapes = [sp for sp in slide.shapes._spTree if sp.tag in (qn("p:sp"), qn("p:pic"), qn("p:cxnSp"))]
        for i, sp in enumerate(shapes):
            txt = _text(sp)
            box = _box(sp)
            if not txt or box is None:
                continue
            # tekstens midte (lodret centreret tekst står i midten; ellers øverst)
            cx = (box[0] + box[2]) / 2
            cy = box[1] + min((box[3] - box[1]) / 2, 300000)
            area = max((box[2] - box[0]) * (box[3] - box[1]), 1)
            # 1. dækket af noget ovenpå?
            for top in shapes[i + 1:]:
                f, tb = _fill(top), _box(top)
                if f and tb and _overlap(box, tb) > 0.35 * area and not _text(top):
                    problems.append("slide %d: %r er dækket af en form ovenpå (%s)"
                                    % (n, txt[:40], top.find(".//" + qn("p:cNvPr")).get("name")))
                    break
            # 2. kontrast mod det, teksten står på
            under = bg
            for low in shapes[:i]:
                f, lb = _fill(low), _box(low)
                if f and f != "IMG" and lb and lb[0] <= cx <= lb[2] and lb[1] <= cy <= lb[3]:
                    under = f
            for col, pt, bold in _runs(sp):
                # WCAG AA: stor tekst (18 pt, eller 14 pt fed) skal have 3:1, ellers 4,5:1
                need = 3.0 if (pt >= 18 or (bold and pt >= 14)) else 4.5
                if under and _contrast(col, under) < need:
                    problems.append("slide %d: %r har for lav kontrast (%s på %s, %.1f:1 — kræver %.1f:1)"
                                    % (n, txt[:40], col, under, _contrast(col, under), need))
                    break
    return problems


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    probs = check(sys.argv[1])
    for p in probs:
        print("FEJL:", p)
    print("%d fejl fundet i %s" % (len(probs), sys.argv[1]))
    sys.exit(1 if probs else 0)
