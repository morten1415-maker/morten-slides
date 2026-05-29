"""
Eksempel: on-brand .pptx i Økonomistyrelsen-temaet.
Kør:  python3 deck.oes.example.py   ->  oes-example.pptx

Eneste forskel fra deck.example.py er theme="oes" — alt indhold-API er ens.
"""
from build_pptx import Deck

deck = Deck(deck_title="Økonomistyrelsen · Eksempel", theme="oes")

deck.title(
    "Årsrapport 2026.",
    lead="Et slidedeck i Økonomistyrelsens identitet — bygget på den brutalistiske skabelon med ØS-grøn som anker.",
    kicker="Økonomistyrelsen · 2026",
)

deck.section("Årets resultater", num="01")   # bruger automatisk brand-grøn flade

deck.content(
    "Tre fokusområder",
    [
        "Digitalisering af økonomiprocesser på tværs af staten.",
        "Øget gennemsigtighed i statsregnskabet.",
        "Stærkere datagrundlag for politiske beslutninger.",
        "Reduktion af administrative byrder.",
    ],
    kicker="Strategi",
    accent_index=2,
)

deck.split(
    "Før og efter",
    left={"label": "Før", "body": "Manuelle indberetninger, spredte systemer, lange svartider."},
    right={"label": "Efter", "body": "Samlet platform, automatiske flows, realtids-overblik."},
    kicker="Transformation",
)

deck.statement(
    "Gode beslutninger bygger på pålidelige tal.",
    source="— Økonomistyrelsen",
)

deck.stats(
    [
        {"num": "1,2 mio", "label": "Behandlede bilag"},
        {"num": "+18%", "label": "Effektivisering", "trend": "up"},
        {"num": "99,8", "label": "Oppetid %"},
    ],
    title="Året i tal",
    kicker="Resultater · 2026",
    primary_index=1,
)

deck.table(
    headers=["Område", "Ansvar", "Budget (mio)", "Afvigelse"],
    rows=[
        ["Drift", "ØS", "240", "−2"],
        ["Udvikling", "ØS", "118", "+5"],
        ["Support", "ØS", "64", "0"],
    ],
    title="Budgetoversigt",
    kicker="Økonomi",
    numeric_cols=(2, 3),
    highlight_row=1,
)

# ---- Præsentations-elementer (samme som primær stil, ØS-farvet) ----

deck.agenda(
    ["Årets resultater", "Digitalisering", "Økonomi & budget", "Næste skridt"],
    title="Agenda", kicker="Indhold", current=1,
)

deck.cover(
    "Stort billede, stærk åbning.",
    lead="Overlay-boks med hård shadow oven på fuldt baggrundsbillede.",
    kicker="Kapitel",
)

deck.cards(
    [
        {"title": "Effektivitet", "body": "Færre manuelle trin gennem automatisering."},
        {"title": "Gennemsigtighed", "body": "Det fremhævede kort — gul flade, kun ét per slide."},
        {"title": "Datagrundlag", "body": "Bedre tal til politiske beslutninger."},
    ],
    title="Tre indsatsområder", kicker="Indsatser", accent_index=1,
)

deck.timeline(
    [
        {"when": "Q1", "what": "Analyse", "desc": "Behov afdækkes."},
        {"when": "Q2", "what": "Design", "desc": "Løsninger formes."},
        {"when": "Q3", "what": "Udrulning", "desc": "Vi er her nu."},
        {"when": "Q4", "what": "Evaluering", "desc": "Læring og justering."},
    ],
    title="Året i fire kvartaler", kicker="Proces", current=2,
)

deck.bignum("+18%", sub="Effektivisering år over år — positivt tal bærer ØS-grøn.",
            kicker="Nøgletal", trend="up")

deck.bars(
    [
        {"label": "Drift", "pct": 100, "value": "240"},
        {"label": "Udvikling", "pct": 49, "value": "118"},
        {"label": "Support", "pct": 27, "value": "64"},
    ],
    title="Budget pr. område", kicker="Fordeling", accent_index=1,
)

deck.testimonial(
    "“Gode beslutninger bygger på pålidelige tal.”",
    name="Økonomistyrelsen", role="Statens økonomiforvaltning", kicker="Udtalelse",
)

deck.closing("Tak.", cta="oes@oes.dk →", kicker="Kontakt")

print("Skrev", deck.save("oes-example.pptx"))
