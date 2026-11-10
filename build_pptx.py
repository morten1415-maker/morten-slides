# -*- coding: utf-8 -*-
"""
Økonomistyrelsens slide-system — stilen og slide-typerne (PowerPoint).

Motoren ligger i oes_pptx.py (elementer, tekst-ombrydning, .pptx-skrivning,
PNG-preview, font-indlejring). Denne fil er STILEN:

  - lys grå grund (#F4F4F2), sorte 3 px kanter, hårde skygger uden blur
  - overskriften står i en sort blok med grøn skygge ("skiltet"), øverst til venstre
  - kapitler: laks-label + stor sort titelblok
  - billeder får en grøn flade forskudt bagved (signaturen) — ingen kant
  - IBM Plex Sans = indhold, IBM Plex Mono = system (labels, numre, tal)
  - laks = signal (numre, labels, fald), lysegrøn = fremhævning

Alt måles i px på et 1280 x 720-canvas — de samme tal som i slides.css.

Brug:
    from build_pptx import Deck
    deck = Deck()
    deck.title("Årsrapport 2026", lead="…")
    deck.section("Årets resultater", num="01")
    deck.content("Tre fokusområder", ["…", "…"])
    deck.save("mit-deck.pptx", preview_dir="preview")   # preview_dir er valgfri

Mangler der en slide-type: brug deck.content("Titel") og tegn selv fra
CONTENT_TOP med byggeklodserne (card, block, chip, square_list, picture,
s.rect/s.line/s.text/s.media) — eller tilføj en ny metode her i samme stil.
"""
import math
import os

import oes_pptx as O
from oes_pptx import (BRAND, CONTENT_BOT, CONTENT_TOP, H, INK, LOGO_INSET, LOGO_W, MEDIA, MEDIA_Y, OFFSET,
                      PAD_TOP, PAD_X, PAPER, S_BODY, S_CAPTION, S_DISPLAY, S_LEAD, S_QUOTE, S_SMALL,
                      S_STATEMENT, S_TITLE, W, layout, para, run)

# ------------------------------------------------------------------ TOKENS
GROUND = "F4F4F2"       # baggrund
BLACK = "000000"        # kanter, skygger, titelblokke
LAKS = "ED5E66"         # signal på STORE flader og på sort (kapitel-label, citatstreg)
LAKS_TEXT = "C93A44"    # laks til tekst på lys grund (4,6:1 — ED5E66 har kun 3:1)
DIM = "6E6E6E"          # nedtonet tekst, stadig læsbar (4,6:1 på grunden)
HILITE = "B3E8CA"       # lysegrøn: overstregning, fremhævet række/kort
RULE = "D6D6D6"         # tynde indre delelinjer
EMPTY = "E2E2DE"        # billed-pladsholder
BORDER = 3              # kantbredde
SHADOW = 8              # hård skygge
GAP = 40                # luft mellem spalter

# Faste mål (px). verify.py holder dem op mod slides.css.
CONTENT_W = W - 2 * PAD_X        # 1104
TITLE_Y = PAD_TOP - 8            # titelblokkens top: 64
TITLE_SIZE = 31                  # tekst i titelblokken
TITLE_PAD = (19, 12)             # luft i titelblokken (vandret, lodret)
TITLE_SHADOW = 7                 # titelblokkens grønne skygge
DISPLAY_PAD = (36, 24)           # luft i de store blokke (forside, kapitel, tak)
DISPLAY_SHADOW = 12
SOURCE_BOTTOM = H - LOGO_INSET - 6   # kilde-linjens underkant: 682
SOURCE_SIZE = 12


# ------------------------------------------------------------------ TEKST-ROLLER
def mono(text, size=S_CAPTION, color=INK, caps=True):
    """Mono-label: altid versaler med sporing ('mono er system')."""
    return run(text, f="label", size=size, color=color, caps=caps, sp=0.06 if caps else 0)


def num(text, size=S_BODY, color=INK, bold=True):
    """Tal: mono, ikke versaler. bold=False til tabelceller."""
    return run(text, f="label" if bold else "mono", size=size, color=color)


def head(text, size=S_TITLE, color=PAPER):
    """Overskrift: fed IBM Plex Sans i normal case."""
    return run(text, f="head", size=size, color=color)


def body(text, size=S_BODY, color=INK, bold=False, hl=None):
    return run(text, size=size, color=color, bold=bold, hl=hl)


def rich(item, size=S_BODY, color=INK):
    """Gør 'tekst', ('Fed start: ', 'resten') eller en liste af run() til en liste af run()."""
    if isinstance(item, str):
        return [body(item, size, color)]
    if isinstance(item, tuple):
        return [body(item[0], size, color, bold=True)] + [body(t, size, color) for t in item[1:]]
    return list(item)


def highlight(text, part, size=S_BODY, color=INK):
    """Tekst hvor delstrengen `part` får lysegrøn overstregning."""
    if not part or part not in text:
        return [body(text, size, color)]
    a, b = text.split(part, 1)
    return [r for r in (body(a, size, color), body(part, size, color, hl=HILITE), body(b, size, color))
            if r["text"]]


def fit_size(text, maker, size, max_w, floor=20, step=2):
    """Skrump `size` til teksten kan stå på én linje i max_w."""
    while size > floor and O._measure(text, maker(text, size)) > max_w:
        size -= step
    return size


# ------------------------------------------------------------------ BYGGEKLODSER
def card(s, x, y, w, h, fill=PAPER, shadow=BLACK, border=BLACK, off=SHADOW, name="Kort"):
    """Boks med sort kant og hård skygge."""
    if shadow:
        s.rect(x + off, y + off, w, h, fill=shadow, name=name + ", skygge")
    s.rect(x, y, w, h, fill=fill, line=border, lw=BORDER if border else 1, name=name)


def block(s, x, y, paras, max_w, fill=BLACK, shadow=BRAND, off=SHADOW, pad=(22, 14), title=False,
          name="Titelblok", measure_only=False):
    """Tekst i en farvet blok, der slutter tæt om teksten. Returnerer (bredde, højde) inkl. skygge."""
    px_, py_ = pad
    lines, th = layout(paras, max_w - 2 * px_)
    widest = max((sum(t[2] for t in ln) for _, ls, _, _ in lines for ln in ls), default=0)
    w, h = widest + 2 * px_ + 2, th + 2 * py_
    if not measure_only:
        if shadow:
            s.rect(x + off, y + off, w, h, fill=shadow, name=name + ", skygge")
        s.rect(x, y, w, h, fill=fill, name=name)
        s.text(x + px_, y + py_, w - 2 * px_, th + 2, paras, title=title)
    return w + (off if shadow else 0), h + (off if shadow else 0)


def chip(s, x, y, text, fill=BLACK, color=PAPER, shadow=None, size=13):
    """Lille label med mono-tekst."""
    return block(s, x, y, [para(mono(text, size, color), lh=1.3)], 900, fill=fill, shadow=shadow, off=4,
                 pad=(10, 5), name="Label")


def chip_w(text, size=13):
    """Bredden af en chip (til centrering)."""
    return O._measure(text, mono(text, size, PAPER)) + 22


def square_list(s, x, y, w, items, gap=18, lh=1.4, size=S_BODY, accent_index=None):
    """Liste med små sorte firkanter (aldrig runde prikker). Returnerer y efter listen."""
    for i, it in enumerate(items):
        ps = [para(*rich(it, size), lh=lh)]
        _, th = layout(ps, w - 30)
        s.rect(x, y + size * lh / 2 - 6, 11, 11, fill=LAKS if i == accent_index else BLACK, name="Firkant")
        s.text(x + 30, y, w - 30, th + 2, ps)
        y += th + gap
    return y - gap if items else y


def list_height(w, items, gap=18, lh=1.4, size=S_BODY):
    return sum(layout([para(*rich(it, size), lh=lh)], w - 30)[1] for it in items) + gap * max(len(items) - 1, 0)


def picture(s, image, x, y, w, h, side=None, focus=(0.5, 0.5), alt="", label="Billede", fit="cover"):
    """Billede — eller en pladsholder hvis image er None. side='right'/'left' giver den
    grønne flade forskudt bagved (systemets signatur)."""
    if side:
        dx = OFFSET if side == "right" else -OFFSET
        s.rect(x + dx, y + OFFSET, w, h, fill=BRAND, name="Grøn flade")
    if image:
        return s.image(image, x, y, w, h, fit=fit, focus=focus, alt=alt)
    s.rect(x, y, w, h, fill=EMPTY, name="Billed-pladsholder")
    if label:
        s.text(x, y, w, h, [para(mono(label, 13, "555555"), align="ctr", lh=1.3)], anchor="ctr")


# ------------------------------------------------------------------ DECK
class Deck(O.Deck):
    """Alle slide-typer. Hver metode returnerer sliden, så du kan tegne videre på den."""

    def __init__(self, deck_title=None, embed_fonts=True, **_ignored):
        super().__init__()
        self.deck_title = deck_title
        self.embed_fonts = embed_fonts
        self.part = None          # "01", "02" … sættes af section() og vises i titelblokken
        self.warnings = []

    # ---------- fælles
    def blank(self, bg=GROUND, logo="green"):
        return super().blank(bg=bg, logo=logo)

    def _title(self, s, text):
        """Overskriften: sort skilt med grøn skygge; delnummeret i laks mono foran."""
        runs = []
        if self.part:
            runs += [mono(self.part, TITLE_SIZE * 0.8, LAKS), run("  ", size=TITLE_SIZE, color=PAPER)]
        runs.append(head(text, TITLE_SIZE))
        block(s, PAD_X, TITLE_Y, [para(*runs, lh=1.2)], CONTENT_W, title=True, off=TITLE_SHADOW,
              pad=TITLE_PAD)

    def _source(self, s, text):
        """Kilde-/note-linje nederst til venstre (mono, ikke versaler — den skal kunne læses)."""
        if text:
            s.text(PAD_X, SOURCE_BOTTOM - 18, 900, 18, [para(mono(text, SOURCE_SIZE, DIM, caps=False), lh=1.3)],
                   anchor="b", name="Kilde")

    def _display(self, s, text, y=None, label=None, below=None, max_w=1060):
        """Stor sort blok med grøn skygge, lodret centreret (forside, kapitel, tak)."""
        size = S_DISPLAY
        t_paras = [para(head(text, size), lh=1.05)]
        _, t_h = block(s, 0, 0, t_paras, max_w, off=DISPLAY_SHADOW, pad=DISPLAY_PAD, measure_only=True)
        l_h = 52 if label else 0
        b_h = (layout(below, 760)[1] + 32) if below else 0
        if y is None:
            y = (H - (l_h + t_h + b_h)) / 2 - 10
        if label:
            block(s, PAD_X, y, [para(head(label, 26), lh=1.2)], 600, fill=LAKS, shadow=BLACK, pad=(20, 8),
                  name="Kapitel-label")
            y += l_h
        block(s, PAD_X, y, t_paras, max_w, off=DISPLAY_SHADOW, pad=DISPLAY_PAD, title=True)
        if below:
            s.text(PAD_X, y + t_h + 32, 760, b_h, below)
        return s

    def _meta(self, s, text):
        if text:
            s.text(PAD_X, SOURCE_BOTTOM - 20, 800, 20, [para(mono(text, 13, DIM), lh=1.3)], anchor="b",
                   name="Meta")

    # ---------- åbning, kapitler, afslutning
    def title(self, title, lead=None, kicker=None, meta=None, image=None, eyebrow=None, focus=(0.5, 0.5),
              alt=""):
        """Forside. Med image: tekst til venstre, billede med grøn flade til højre.
        Uden: stor sort titelblok. meta = mono-linje nederst (dato / anledning / afsender)."""
        if image:
            return self.hero(title, image, eyebrow=eyebrow, lead=lead, meta=meta, focus=focus, alt=alt)
        s = self.blank()
        self._display(s, title, below=[para(body(lead, S_LEAD), lh=1.45)] if lead else None)
        self._meta(s, meta)
        return s

    def section(self, title, num=None, variant=None, label=None):
        """Kapitel: laks-label + stor sort titelblok. num='01' giver labelen 'Del 1' og sætter
        nummeret foran titlerne på de efterfølgende slides."""
        digits = "".join(ch for ch in str(num or label or "") if ch.isdigit())
        self.part = digits.zfill(2) if digits else None
        if label is None and digits:
            label = "Del %d" % int(digits)
        return self._display(self.blank(), title, label=label)

    def closing(self, title="Tak", cta=None, kicker=None, lines=None):
        """Afslutning: stor sort blok + valgfri linjer og en grøn mono-label (cta)."""
        s = self.blank()
        below = [para(*rich(ln, S_LEAD), lh=1.45) for ln in (lines or [])]
        size = S_DISPLAY
        t_paras = [para(head(title, size), lh=1.05)]
        _, t_h = block(s, 0, 0, t_paras, 1000, off=DISPLAY_SHADOW, pad=DISPLAY_PAD, measure_only=True)
        l_h = layout(below, 1000)[1] if below else 0
        c_h = 46 if cta else 0
        total = t_h + (32 + l_h if below else 0) + (32 + c_h if cta else 0)
        y = (H - total) / 2
        block(s, PAD_X, y, t_paras, 1000, off=DISPLAY_SHADOW, pad=DISPLAY_PAD, title=True)
        y += t_h
        if below:
            s.text(PAD_X, y + 32, 1000, l_h + 2, below)
            y += 32 + l_h
        if cta:
            block(s, PAD_X, y + 32, [para(mono(cta, 15, PAPER, caps=False), lh=1.3)], 900, fill=BRAND, shadow=BLACK, off=5,
                  pad=(18, 10), name="CTA")
        return s

    # ---------- overskrift + indhold
    def content(self, title=None, bullets=None, kicker=None, accent_index=None, sub=None):
        """Overskrift øverst + liste. Uden bullets: tegn selv videre fra CONTENT_TOP.
        sub = kort forklaring lige under titelblokken."""
        s = self.blank()
        if title:
            self._title(s, title)
        if sub:
            s.text(PAD_X, PAD_TOP + 82, 520, 70, [para(body(sub), lh=1.4)])
        if bullets:
            square_list(s, PAD_X, CONTENT_TOP + (70 if sub else 0), 1060, bullets, gap=24,
                        accent_index=accent_index)
        return s

    def agenda(self, items, title="Agenda", kicker=None, current=None):
        """Nummereret liste mellem sorte streger. current (0-baseret) fremhæver ét punkt."""
        s = self.content(title)
        n = len(items)
        row = min(99, (CONTENT_BOT - CONTENT_TOP) / max(n, 1))
        hs = 32 if row >= 90 else 26
        y = CONTENT_TOP
        s.rect(PAD_X, y, 900, 3, fill=BLACK, name="Streg")
        for i, t in enumerate(items):
            on = current is None or i == current
            s.text(PAD_X, y + 3, 90, row - 3, [para(num("%02d" % (i + 1), hs * 0.8, LAKS_TEXT if on else DIM),
                                                   lh=1.2)], anchor="ctr")
            r = head(t, hs, INK)
            if current is not None and i == current:
                r["hl"] = HILITE
            s.text(PAD_X + 90, y + 3, 800, row - 3, [para(r, lh=1.2)], anchor="ctr")
            y += row
            s.rect(PAD_X, y, 900, 1.5, fill=BLACK, name="Streg")
        return s

    def split(self, title, left, right, kicker=None, arrow=False):
        """To spalter under hver sin sorte streg. Hver side: {'label', 'head' (valgfri), 'body',
        'accent' (valgfri: grøn overskrift)}. arrow=True sætter en pil imellem (før → efter)."""
        s = self.content(title)
        col = (CONTENT_W - GAP * 2) / 2
        for i, d in enumerate((left, right)):
            x = PAD_X + i * (col + GAP * 2)
            s.rect(x, CONTENT_TOP, col, 3, fill=BLACK, name="Streg")
            ps = [para(mono(d.get("label", ""), 14, LAKS_TEXT), lh=1.3)]
            if d.get("head"):
                ps.append(para(head(d["head"], 34, BRAND if d.get("accent") else INK), lh=1.15, before=10))
            if d.get("body"):
                ps.append(para(*rich(d["body"]), lh=1.45, before=14))
            s.text(x, CONTENT_TOP + 20, col, CONTENT_BOT - CONTENT_TOP - 20, ps)
        if arrow:
            s.text(PAD_X + col + 8, CONTENT_TOP + 64, GAP * 2 - 16, 40, [para(head("→", 34, INK), align="ctr",
                                                                              lh=1.15)])
        return s

    def cards(self, items, title=None, kicker=None, accent_index=None):
        """2–4 spalter under hver sin sorte streg: mono-nummer, overskrift, tekst.
        accent_index: den spalte der skal bære ØS-grøn."""
        s = self.content(title)
        n = len(items)
        col = (CONTENT_W - GAP * (n - 1)) / n
        for i, it in enumerate(items):
            x = PAD_X + i * (col + GAP)
            on = i == accent_index
            s.rect(x, CONTENT_TOP, col, 3, fill=BRAND if on else BLACK, name="Streg")
            s.text(x, CONTENT_TOP + 22, col, CONTENT_BOT - CONTENT_TOP - 22, [
                para(num("%02d" % (i + 1), 15, LAKS_TEXT), lh=1.3),
                para(head(it["title"], 26, BRAND if on else INK), lh=1.2, before=10),
                para(*rich(it.get("body", "")), lh=1.4, before=10)])
        return s

    def comparison(self, left, right, title=None, kicker=None, win="right"):
        """To kort med +/– foran hvert punkt. Hver side: {'head', 'items': [(True/False, tekst)]}.
        Vinderen (win='left'/'right'/None) får skygge og grøn overskrift."""
        s = self.content(title)
        col = (CONTENT_W - GAP) / 2
        rows = max(len(left["items"]), len(right["items"]))
        h = 28 + 24 + 20 + rows * 46 + 18
        for i, (d, side) in enumerate(((left, "left"), (right, "right"))):
            x = PAD_X + i * (col + GAP)
            won = win == side
            card(s, x, CONTENT_TOP, col, h, shadow=BLACK if won else None)
            s.text(x + 28, CONTENT_TOP + 26, col - 56, 24, [para(mono(d["head"], 14, BRAND if won else INK),
                                                                 lh=1.3)])
            s.rect(x + 28, CONTENT_TOP + 60, col - 56, 2, fill=BLACK, name="Streg")
            y = CONTENT_TOP + 74
            for ok, text in d["items"]:
                s.text(x + 28, y, 30, 46, [para(num("+" if ok else "–", 24, BRAND if ok else LAKS_TEXT), lh=1.2)],
                       anchor="ctr")
                s.text(x + 62, y, col - 90, 46, [para(body(text, 20), lh=1.3)], anchor="ctr")
                y += 46
        return s

    def timeline(self, steps, title=None, kicker=None, current=None):
        """Vandret forløb. steps: [{'when', 'what', 'desc'}]. current (0-baseret) er grøn."""
        s = self.content(title)
        n = len(steps)
        col = CONTENT_W / n
        y = CONTENT_TOP + 70
        s.line(PAD_X + 14, y, PAD_X + (n - 1) * col + 14, y, color=BLACK, lw=3, name="Tidslinje")
        for i, st in enumerate(steps):
            x = PAD_X + i * col
            s.rect(x, y - 14, 28, 28, fill=BRAND if i == current else PAPER, line=BLACK, lw=3, name="Trin")
            ps = [para(mono(st.get("when", "%02d" % (i + 1)), 14, LAKS_TEXT), lh=1.3),
                  para(head(st["what"], S_BODY, INK), lh=1.25, before=8)]
            if st.get("desc"):
                ps.append(para(body(st["desc"], S_SMALL), lh=1.4, before=8))
            s.text(x, y + 36, col - 28, CONTENT_BOT - y - 36, ps)
        return s

    def process(self, steps, title=None, current=None, only=None):
        """Lodret proces med nummererede trin. current = navnet på det aktive trin (grøn,
        resten nedtones). only = vis kun dette trin (til at bygge listen op slide for slide)."""
        s = self.content(title)
        n = len(steps)
        gap = min(70, (CONTENT_BOT - CONTENT_TOP - 40) / max(n - 1, 1))
        x, y0 = PAD_X + 14, CONTENT_TOP + 20
        if only is None:
            s.line(x, y0, x, y0 + gap * (n - 1), color=BLACK, lw=3, name="Proceslinje")
        for i, st in enumerate(steps):
            if only is not None and st != only:
                continue
            y = y0 + gap * i
            on = current is None or st == current
            s.rect(x - 14, y - 14, 28, 28, fill=BRAND if st == current else PAPER,
                   line=BLACK if on else "B5B5B5", lw=3, name="Trin")
            s.text(x + 34, y - 16, 46, 32, [para(num("%02d" % (i + 1), 15, LAKS_TEXT if on else DIM), lh=1.3)],
                   anchor="ctr")
            s.text(x + 80, y - 17, 800, 34, [para(body(st, color=INK if on else DIM, bold=True), lh=1.3)],
                   anchor="ctr")
        return s

    # ---------- citater og udsagn
    def statement(self, quote, attribution=None, source=None, title=None, eyebrow=None, highlight_part=None,
                  width=1000):
        """Citat/udsagn med laks-streg i venstre side — ingen kasse.
        Med title: citatet står under titelblokken. highlight_part: delstreng med lysegrøn
        overstregning. quote kan også være en liste af run()."""
        s = self.content(title)
        size = S_QUOTE if title else S_STATEMENT
        runs = highlight(quote, highlight_part, size) if isinstance(quote, str) else [dict(r, size=size)
                                                                                     for r in quote]
        paras = []
        if eyebrow:
            paras.append(para(mono(eyebrow, 14, LAKS_TEXT), lh=1.3))
        paras.append(para(*runs, lh=1.25, before=16 if eyebrow else 0))
        if attribution:
            paras.append(para(mono("— " + attribution.lstrip("—– "), 14, INK), lh=1.3, before=24))
        indent = 34
        if title:
            width = min(width, 960)
            x = (W - width) / 2
            _, th = layout(paras, width - indent)
            y = (CONTENT_TOP - 20 + CONTENT_BOT) / 2 - th / 2
        else:
            x = PAD_X
            _, th = layout(paras, width - indent)
            y = (H - th) / 2
        s.rect(x, y, 6, th, fill=LAKS, name="Citatstreg")
        s.text(x + indent, y, width - indent, th + 2, paras)
        self._source(s, source)
        return s

    def testimonial(self, quote, name, role=None, kicker=None, image=None, focus=(0.5, 0.5), alt=""):
        """Citat med portræt: billede med grøn flade til venstre, citat med laks-streg til højre."""
        s = self.blank()
        size = 300
        y0 = (H - size - OFFSET) / 2
        picture(s, image, PAD_X + OFFSET, y0, size, size, side="left", focus=focus, alt=alt or name,
                label="Portræt")
        x = PAD_X + OFFSET + size + 72
        w = W - PAD_X - x
        paras = [para(body(quote, 34), lh=1.25),
                 para(body(name, S_BODY, bold=True), lh=1.3, before=26)]
        if role:
            paras.append(para(mono(role, 13, DIM), lh=1.3, before=6))
        _, th = layout(paras, w - 34)
        y = y0 + (size + OFFSET - th) / 2
        s.rect(x, y, 6, th, fill=LAKS, name="Citatstreg")
        s.text(x + 34, y, w - 34, th + 2, paras)
        return s

    # ---------- billeder
    def media_text(self, title, image=None, body_paras=None, bullets=None, eyebrow=None, side="right",
                   focus=(0.5, 0.5), alt="", size=S_TITLE + 8, free=False, accent_index=None):
        """Billede + tekst. Regel: højst én 'hård' ting pr. slide — billedet har den grønne
        flade, så overskriften står fri (ingen blok). free=True: frilagt motiv (PNG med
        gennemsigtighed) uden grøn flade."""
        s = self.blank()
        col_w = CONTENT_W - MEDIA - OFFSET - 96
        if side == "right":
            mx, x = W - PAD_X - OFFSET - MEDIA, PAD_X
        else:
            mx, x = PAD_X + OFFSET, PAD_X + OFFSET + MEDIA + 96
        if free and image:
            s.image(image, mx - 30, MEDIA_Y, MEDIA + 60 + OFFSET, MEDIA + OFFSET, fit="contain", alt=alt)
        else:
            picture(s, image, mx, MEDIA_Y, MEDIA, MEDIA, side=side, focus=focus, alt=alt)
        while size > 28 and max(O._measure(wd, head(wd, size)) for wd in title.split()) > col_w:
            size -= 2                      # det længste ord skal kunne stå på én linje
        paras = []
        if eyebrow:
            paras.append(para(mono(eyebrow, 14, LAKS_TEXT), lh=1.3))
        paras.append(para(head(title, size, INK), lh=1.08, before=14 if eyebrow else 0))
        for i, p in enumerate(body_paras or []):
            p = dict(p)
            if i == 0:
                p["before"] = max(p["before"], 24)
            paras.append(p)
        _, th = layout(paras, col_w)
        lh_ = (list_height(col_w, bullets) + 28) if bullets else 0
        y = MEDIA_Y + (MEDIA + OFFSET - th - lh_) / 2
        s.text(x, y, col_w, th + 2, paras, title=True)
        if bullets:
            square_list(s, x, y + th + 28, col_w, bullets, accent_index=accent_index)
        return s

    def hero(self, title, image=None, eyebrow=None, lead=None, side="right", focus=(0.5, 0.5), alt="",
             meta=None):
        """Som media_text, men med stor overskrift — forside og andre store øjeblikke."""
        lead_p = [para(body(lead, S_LEAD), lh=1.45)] if lead else None
        s = self.media_text(title, image, body_paras=lead_p, eyebrow=eyebrow, side=side, focus=focus, alt=alt,
                            size=S_DISPLAY - 4)
        self._meta(s, meta)
        return s

    def image_text(self, title, bullets=None, kicker=None, image=None, image_left=False, accent_index=None,
                   focus=(0.5, 0.5), alt=""):
        """Billede med grøn flade + overskrift og punkter. image_left=True spejler."""
        return self.media_text(title, image, bullets=bullets, side="left" if image_left else "right",
                               focus=focus, alt=alt, accent_index=accent_index)

    def cover(self, title, lead=None, kicker=None, image=None, focus=(0.5, 0.5), alt=""):
        """Fuldt billede med sort titelblok nederst til venstre."""
        s = self.blank(logo=None)
        picture(s, image, 0, 0, W, H, focus=focus, alt=alt, label="Baggrundsbillede 16:9")
        paras = [para(head(title, 52), lh=1.08)]
        if lead:
            paras.append(para(body(lead, S_BODY, PAPER), lh=1.4, before=14))
        _, bh = block(s, 0, 0, paras, 820, off=DISPLAY_SHADOW, pad=DISPLAY_PAD, measure_only=True)
        block(s, PAD_X, H - 64 - bh, paras, 820, off=DISPLAY_SHADOW, pad=DISPLAY_PAD, title=True)
        return s

    def image_full(self, caption=None, label=None, kicker=None, image=None, focus=(0.5, 0.5), alt=""):
        """Fuldt billede med billedtekst i en sort blok: mono-label i laks + tekst."""
        s = self.blank(logo=None)
        picture(s, image, 0, 0, W, H, focus=focus, alt=alt or (caption or ""), label="Fuldt billede")
        label = label or kicker
        runs = []
        if label:
            runs += [mono(label, 14, LAKS), run("   ", size=20, color=PAPER)]
        if caption:
            runs.append(body(caption, 20, PAPER, bold=True))
        if runs:
            paras = [para(*runs, lh=1.3)]
            _, bh = block(s, 0, 0, paras, 1000, off=6, pad=(20, 12), measure_only=True)
            block(s, PAD_X, H - 56 - bh, paras, 1000, off=6, pad=(20, 12), name="Billedtekst")
        return s

    def image_grid(self, title=None, n=6, cols=3, kicker=None, images=None):
        """Gitter af billeder under titlen. images: liste af stier (None = pladsholder)."""
        s = self.content(title)
        images = list(images or [None] * n)
        rows = math.ceil(len(images) / cols)
        g = 20
        top = CONTENT_TOP if title else PAD_TOP
        cw = (CONTENT_W - g * (cols - 1)) / cols
        ch = (CONTENT_BOT - top - g * (rows - 1)) / rows
        for i, im in enumerate(images):
            picture(s, im, PAD_X + (i % cols) * (cw + g), top + (i // cols) * (ch + g), cw, ch,
                    label="%02d" % (i + 1))
        return s

    # ---------- tal
    def bignum(self, value, sub=None, kicker=None, trend=None, source=None, title=None):
        """Ét stort tal. trend='up' giver ØS-grøn, 'down' laks."""
        s = self.content(title)
        col = {"up": BRAND, "down": LAKS_TEXT}.get(trend, INK)
        # "+18 %" -> tallet i mono, enheden mindre i sans (et mono-mellemrum er for bredt i 200 px)
        value, _, unit = value.strip().partition(" ")
        if not unit and value.endswith("%") and len(value) > 1:
            value, unit = value[:-1], "%"

        def runs(z):
            return [num(value, z, col)] + ([head(" " + unit, z * 0.45, col)] if unit else [])

        size = 200
        while size > 80 and sum(O._measure(r["text"], r) for r in runs(size)) > CONTENT_W:
            size -= 4
        sub_p = [para(body(sub, S_LEAD), lh=1.45)] if sub else []
        sub_h = layout(sub_p, 640)[1] if sub else 0
        total = size + (20 + sub_h if sub else 0)
        top = CONTENT_TOP if title else PAD_TOP
        y = (top + CONTENT_BOT - total) / 2
        s.text(PAD_X, y, CONTENT_W, size + 2, [para(*runs(size), lh=1.0)], name="Stort tal")
        if sub:
            s.text(PAD_X, y + size + 20, 640, sub_h + 2, sub_p)
        self._source(s, source)
        return s

    def stats(self, cards, title=None, kicker=None, primary_index=None, source=None):
        """Nøgletal i hvide kort. cards: [{'num', 'label', 'trend' (valgfri)}].
        primary_index: det kort der fremhæves (lysegrøn flade)."""
        s = self.content(title)
        n = len(cards)
        g = 32
        cw = (CONTENT_W - g * (n - 1)) / n
        ch = 230
        top = CONTENT_TOP if title else PAD_TOP
        y = (top + CONTENT_BOT - ch) / 2
        size = min(fit_size(c["num"], lambda t, z: num(t, z), 64, cw - 64, floor=28) for c in cards)
        for i, c in enumerate(cards):
            x = PAD_X + i * (cw + g)
            card(s, x, y, cw, ch, fill=HILITE if i == primary_index else PAPER, name="Nøgletal")
            col = {"up": BRAND, "down": LAKS_TEXT}.get(c.get("trend"), INK)
            s.text(x + 32, y + 24, cw - 64, ch - 48, [
                para(num(c["num"], size, col), lh=1.1),
                para(mono(c["label"], 14, INK if i == primary_index else DIM), lh=1.35, before=14)],
                anchor="ctr")
        self._source(s, source)
        return s

    def bars(self, rows, title=None, kicker=None, accent_index=None, source=None):
        """Vandrette søjler. rows: [{'label', 'pct' (0–100), 'value'}]. accent_index er grøn."""
        s = self.content(title)
        n = len(rows)
        pitch, bh = 64, 34
        top = CONTENT_TOP if title else PAD_TOP
        y0 = (top + CONTENT_BOT - (pitch * (n - 1) + bh)) / 2
        bx, bw = PAD_X + 240, CONTENT_W - 240 - 120
        s.line(bx, y0 - 16, bx, y0 + pitch * (n - 1) + bh + 16, color=BLACK, lw=3, name="Akse")
        for i, r in enumerate(rows):
            y = y0 + i * pitch
            on = i == accent_index
            w = max(bw * float(r["pct"]) / 100.0, 3)
            s.text(PAD_X, y, 224, bh, [para(body(r["label"], 20, bold=on), lh=1.3)], anchor="ctr")
            s.rect(bx, y, w, bh, fill=BRAND if on else BLACK, name="Søjle")
            s.text(bx + w + 14, y, 110, bh, [para(num(str(r.get("value", "")), 20, BRAND if on else INK),
                                                  lh=1.3)], anchor="ctr")
        self._source(s, source)
        return s

    def table(self, headers, rows, title=None, kicker=None, numeric_cols=(), highlight_row=None, source=None,
              col_widths=None):
        """Tabel: sort hovedrække med mono-versaler, hvide rækker, sort kant og hård skygge.
        numeric_cols højrestilles i mono. highlight_row får lysegrøn flade."""
        s = self.content(title)
        n = len(headers)
        wts = list(col_widths or [1] * n)
        cws = [CONTENT_W * w / float(sum(wts)) for w in wts]
        top = CONTENT_TOP if title else PAD_TOP
        hh = 46
        rh = min(54, (CONTENT_BOT - top - hh - SHADOW) / max(len(rows), 1))
        fs = 20 if rh >= 46 else 17
        th = hh + rh * len(rows)
        s.rect(PAD_X + SHADOW, top + SHADOW, CONTENT_W, th, fill=BLACK, name="Tabel, skygge")
        s.rect(PAD_X, top, CONTENT_W, th, fill=PAPER, name="Tabel")
        s.rect(PAD_X, top, CONTENT_W, hh, fill=BLACK, name="Hovedrække")

        def cell(ci, y, h, r):
            x = PAD_X + sum(cws[:ci])
            s.text(x + 18, y, cws[ci] - 36, h, [para(r, align="r" if ci in numeric_cols else "l", lh=1.25)],
                   anchor="ctr")

        for ci, t in enumerate(headers):
            cell(ci, top, hh, mono(t, 13, PAPER))
        for ri, row in enumerate(rows):
            y = top + hh + ri * rh
            hot = ri == highlight_row
            if hot:
                s.rect(PAD_X, y, CONTENT_W, rh, fill=HILITE, name="Fremhævet række")
            elif ri:
                s.rect(PAD_X, y, CONTENT_W, 1, fill=RULE, name="Delelinje")
            for ci, t in enumerate(row):
                cell(ci, y, rh, num(str(t), fs, INK, bold=hot) if ci in numeric_cols
                     else body(str(t), fs, bold=hot))
        s.rect(PAD_X, top, CONTENT_W, th, fill=None, line=BLACK, lw=BORDER, name="Tabel, kant")
        self._source(s, source)
        return s

    # ---------- gem
    def save(self, path, preview_dir=None):
        """Skriv .pptx (og PNG-preview hvis preview_dir er sat). Returnerer stien.
        Advarsler om tekst der ikke kan være i sin boks ligger i self.warnings."""
        self.warnings = super().save(path, preview_dir=preview_dir, embed_fonts=self.embed_fonts)
        return path
