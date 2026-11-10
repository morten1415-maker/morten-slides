# -*- coding: utf-8 -*-
"""
GALLERI — ét eksempel på hver slide-type i ØS-stilen.
Kør:  python deck.example.py   ->  example.pptx + preview/slide-NN.png

Dette er et komponentbibliotek, IKKE en deck du kopierer. Byg dit eget deck
ud fra indholdet: vælg de slide-typer fortællingen kræver, i din rækkefølge.
"""
import os
import sys

from build_pptx import CONTENT_TOP, GAP, W, Deck, body, card, para

HERE = os.path.dirname(os.path.abspath(__file__))
deck = Deck(deck_title="Økonomistyrelsen · Galleri")

# ---- Åbning ----
deck.title(
    "Årsrapport 2026",
    lead="Status på digitalisering, økonomi og de næste skridt.",
    meta="Økonomistyrelsen · Ledelsesoverblik · August 2026",
)

deck.hero(
    "Ét samlet system",
    eyebrow="Forside med billede",
    lead="deck.title(…, image=…) giver samme layout: tekst til venstre, billede med grøn flade til højre.",
)

deck.agenda(
    ["Årets resultater", "Digitalisering", "Økonomi og budget", "Næste skridt"],
    title="Agenda", current=1,
)

# ---- Del 1: tekst og struktur ----
deck.section("Årets resultater", num="01")

deck.content(
    "Fire fokusområder",
    [
        ("Digitalisering: ", "økonomiprocesser samles på tværs af staten."),
        ("Gennemsigtighed: ", "statsregnskabet bliver lettere at læse."),
        ("Datagrundlag: ", "bedre tal til politiske beslutninger."),
        ("Færre byrder: ", "mindre administration for institutionerne."),
    ],
    accent_index=2,
)

deck.split(
    "Før og efter",
    left={"label": "Før", "head": "Spredte systemer",
          "body": "Manuelle indberetninger, mange indgange og lange svartider."},
    right={"label": "Efter", "head": "Én samlet platform", "accent": True,
           "body": "Automatiske flows og overblik i realtid — ét sted at lede."},
    arrow=True,
)

deck.cards(
    [
        {"title": "Effektivitet", "body": "Færre manuelle trin gennem automatisering."},
        {"title": "Gennemsigtighed", "body": "Den fremhævede spalte bærer ØS-grøn — kun én pr. slide."},
        {"title": "Datagrundlag", "body": "Bedre tal til politiske beslutninger."},
    ],
    title="Tre indsatsområder", accent_index=1,
)

deck.comparison(
    left={"head": "I dag", "items": [
        (False, "Manuelle indberetninger"),
        (False, "Spredte systemer"),
        (True, "Kendt af brugerne"),
    ]},
    right={"head": "Fremover", "items": [
        (True, "Automatiske flows"),
        (True, "Samlet platform"),
        (True, "Overblik i realtid"),
    ]},
    title="Gammel og ny løsning", win="right",
)

deck.timeline(
    [
        {"when": "Q1", "what": "Analyse", "desc": "Behovene afdækkes."},
        {"when": "Q2", "what": "Design", "desc": "Løsningerne formes."},
        {"when": "Q3 · nu", "what": "Udrulning", "desc": "Første institutioner er i gang."},
        {"when": "Q4", "what": "Evaluering", "desc": "Læring og justering."},
    ],
    title="Året i fire kvartaler", current=2,
)

deck.process(
    ["Kortlægning af antagelser", "Testprotokol", "Pilottest", "Rekruttering", "Test", "Dataanalyse",
     "Formidling af indsigter"],
    title="Testprocessen", current="Pilottest",
)

# Egne layouts: deck.content("Titel") + byggeklodserne. sub = forklaring under titlen.
s = deck.content("Byg selv videre", sub="Kort, labels og streger er byggeklodser.")
for i, parts in enumerate((("Én ", "afgrænset", " antagelse"), ("Kan ", "testes", ""), ("Har et ", "skarpt", " mål"))):
    x = W / 2 - 236 * 1.5 - GAP + i * (236 + GAP)
    card(s, x, CONTENT_TOP + 90, 236, 236)
    s.text(x + 21, CONTENT_TOP + 111, 194, 194,
           [para(*[body(t, 24, bold=(j == 1)) for j, t in enumerate(parts) if t], align="ctr", lh=1.25)],
           anchor="ctr")

# ---- Del 2: citater og billeder ----
deck.section("Citater og billeder", num="02")

deck.statement(
    "“A person's perceptions and responses that result from the use or anticipated use of a product, "
    "system or service.”",
    attribution="ISO 9241-210", title="Hvad er UX?", highlight_part="use or anticipated use",
)

deck.statement("Gode beslutninger bygger på pålidelige tal.", attribution="Økonomistyrelsen")

deck.testimonial(
    "“Vi fik ét sted at lede — og tallene stemmer første gang.”",
    name="Navn Efternavn", role="Økonomichef, institution",
)

deck.image_text(
    "Billede og tekst",
    [
        "Billedet står i en fast kvadratisk ramme.",
        "Den grønne flade er forskudt bagved.",
        "image_left=True spejler layoutet.",
    ],
    image_left=True,
)

deck.media_text("Frilagt motiv", eyebrow="media_text(free=True)", body_paras=None,
                bullets=["Uden billede vises en pladsholder.", "Med free=True står et frilagt PNG uden flade."])

deck.image_grid(title="Billed-gitter", n=6, cols=3)

deck.cover(
    "Stort billede, stærk åbning",
    lead="Sort titelblok med grøn skygge oven på et billede i fuld flade.",
)

deck.image_full(caption="Billedtekst i sort blok oven på et billede i fuld flade.", label="Foto")

# ---- Del 3: tal ----
deck.section("Data og resultater", num="03")

deck.bignum("+18 %", sub="Effektivisering år over år. Ét tal, fuld opmærksomhed.",
            trend="up", source="Kilde: Statens Regnskab, 2026")

deck.stats(
    [
        {"num": "1,2 mio.", "label": "Behandlede bilag"},
        {"num": "+18 %", "label": "Effektivisering", "trend": "up"},
        {"num": "99,8 %", "label": "Oppetid"},
    ],
    title="Året i tal", source="Kilde: Statens Regnskab, 2026",
    primary_index=1,
)

deck.bars(
    [
        {"label": "Drift", "pct": 100, "value": "240"},
        {"label": "Udvikling", "pct": 49, "value": "118"},
        {"label": "Support", "pct": 27, "value": "64"},
    ],
    title="Budget pr. område", accent_index=1,
    source="Mio. kr. Kilde: Statens Regnskab, 2026",
)

deck.table(
    headers=["Område", "Ansvar", "Budget, mio. kr.", "Afvigelse"],
    rows=[
        ["Drift", "ØS", "240", "−2"],
        ["Udvikling", "ØS", "118", "+5"],
        ["Support", "ØS", "64", "0"],
    ],
    title="Budgetoversigt", source="Kilde: Statens Regnskab, 2026",
    numeric_cols=(2, 3),
    highlight_row=1,
)

deck.closing("Tak", lines=["Spørgsmål og kommentarer er velkomne."], cta="oes@oes.dk")

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out = deck.save(os.path.join(HERE, "example.pptx"), preview_dir=os.path.join(HERE, "preview"))
    print("Skrev %s — %d slides, %d advarsler" % (os.path.basename(out), len(deck.slides), len(deck.warnings)))
