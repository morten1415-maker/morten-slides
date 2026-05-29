"""
Eksempel: byg et on-brand .pptx med Morten DS-generatoren.
Kør:  python3 deck.example.py   ->  skriver morten-ds-example.pptx

Dette er mønsteret jeg følger når brugeren beder om "et PowerPoint om X":
opret Deck, kald slide-typer i rækkefølge, gem.
"""
from build_pptx import Deck

deck = Deck(deck_title="Morten DS · Slide-template")

deck.title(
    "Bærbar brutalisme.",
    lead="Et slide-system der bærer din brand-karakter — sort kant, hård shadow, mono-label.",
    kicker="Morten Design System · 2026",
)

deck.section("Grundprincipper", num="01", variant="dark")

deck.content(
    "Reservér de stærke elementer",
    [
        "Maks. én rød-shadow per slide — det er din primary action.",
        "Maks. én gul flade per slide — det vigtigste KPI eller aktive sektion.",
        "Hvis alt råber, lytter ingen.",
        "Sort til alle yderkanter, lysegrå kun til indre delelinjer.",
        "Tal er altid mono + tabular-nums.",
    ],
    kicker="Princip",
    accent_index=2,
)

deck.split(
    "Mono er system. Sans er indhold.",
    left={"label": "Mono — IBM Plex Mono",
          "body": "Labels, metadata, status, tal, tider, tags, kbd. ALTID uppercase med letter-spacing."},
    right={"label": "Sans — IBM Plex Sans",
           "body": "Overskrifter, brødtekst, bullets, indhold. Normal case, naturlig læsning."},
    kicker="Typografi",
)

deck.statement(
    "Den røde shadow betyder “tryk her”. Brug den én gang.",
    source="— Designprincip nr. 3",
)

deck.stats(
    [
        {"num": "128", "label": "Aktive komponenter"},
        {"num": "+34%", "label": "Vækst i adoption", "trend": "up"},
        {"num": "99.9", "label": "Uptime %"},
    ],
    title="Tre tal der tæller",
    kicker="Resultater · Q1 2026",
    primary_index=1,
)

deck.table(
    headers=["Komponent", "Ejer", "Brug", "Issues"],
    rows=[
        ["Button", "Core", "1.240", "2"],
        ["DataTable", "Core", "512", "7"],
        ["Modal", "Core", "388", "1"],
        ["CommandPalette", "Core", "204", "0"],
    ],
    title="Komponent-status",
    kicker="Oversigt",
    numeric_cols=(2, 3),
    highlight_row=1,
)

# ---- Præsentations-elementer (billed-slides, agenda, kort, timeline osv.) ----

deck.agenda(
    ["Grundprincipper", "Slide-typer", "Billed-layouts", "Data & resultater"],
    title="Agenda", kicker="Indhold", current=2,
)

deck.cover(
    "Stort billede, stærk åbning.",
    lead="Overlay-boks med hård shadow oven på fuldt baggrundsbillede.",
    kicker="Kapitel",
    # image="sti/til/billede.jpg",   # udelad for placeholder
)

deck.image_full(
    caption="Caption i sort stribe — mono-kicker + titel.",
    kicker="Foto",
)

deck.image_text(
    "Billede + tekst side om side",
    [
        "Halv-til-halv layout med skarpt billed-felt.",
        "image_left=True giver spejlet variant.",
        "Billedet får 2px sort kant som alt andet.",
    ],
    kicker="Case", accent_index=2,
)

deck.image_grid(title="Billed-grid", n=6, cols=3, kicker="Galleri")

deck.cards(
    [
        {"title": "Hurtig", "body": "Offset-shadows gør hierarkiet øjeblikkeligt læsbart."},
        {"title": "Distinkt", "body": "Det fremhævede kort bruger gul flade — kun ét per slide."},
        {"title": "Konsistent", "body": "Samme tokens som UI-systemet, skaleret til projektor."},
    ],
    title="Tre kort med ikon-plads", kicker="Funktioner", accent_index=1,
)

deck.timeline(
    [
        {"when": "Q1", "what": "Research", "desc": "Indsigt og behov afdækkes."},
        {"when": "Q2", "what": "Design", "desc": "Koncept og prototyper."},
        {"when": "Q3", "what": "Build", "desc": "Vi er her nu."},
        {"when": "Q4", "what": "Launch", "desc": "Udrulning og læring."},
    ],
    title="Fire trin", kicker="Proces", current=2,
)

deck.comparison(
    left={"head": "Før", "items": [(False, "Spredte systemer"), (False, "Manuelle flows"), (True, "Kendt af alle")]},
    right={"head": "Efter", "items": [(True, "Samlet platform"), (True, "Automatiske flows"), (True, "Realtids-overblik")]},
    title="Før vs. efter", kicker="Valg", win="right",
)

deck.bignum("+34%", sub="Vækst i adoption år over år — ét tal, fuld fokus.",
            kicker="Nøgletal", trend="up")

deck.bars(
    [
        {"label": "Button", "pct": 100, "value": "1.240"},
        {"label": "DataTable", "pct": 41, "value": "512"},
        {"label": "Modal", "pct": 31, "value": "388"},
        {"label": "Palette", "pct": 16, "value": "204"},
    ],
    title="Brug pr. komponent", kicker="Fordeling", accent_index=1,
)

deck.testimonial(
    "“Systemet gør det nemt at lave slides der ser bevidste ud.”",
    name="Morten Mølgaard", role="Designer", kicker="Udtalelse",
)

deck.closing("Lad os bygge.", cta="morten@design.system →", kicker="Tak")

out = deck.save("morten-ds-example.pptx")
print("Skrev", out)
