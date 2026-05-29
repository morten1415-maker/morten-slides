# Mortens Præsentationer — Instruktioner til Claude Code

Dette projekt genererer on-brand slidedeck — HTML-slides og PowerPoint — der følger Mortens personlige design system ("bærbar brutalisme").

**Læs `PRESENTATIONS.md` før du genererer et deck.** Den indeholder de fulde brand-regler oversat til slides: type-skala, farvebrug, slide-typer, do/don'ts.

**To temaer:**
- **Primær præsentation stil** (standard, Mortens egen stil) — `PRESENTATIONS.md`, `slides.css`, `template.html`, `Deck(theme="default")` (alias `"primær"`).
- **Økonomistyrelsen** — `PRESENTATIONS-OES.md`, `slides-oes.css`, `template-oes.html`, `Deck(theme="oes")`. = Primær stil + ØS-grøn `#066b43` brugt sparsomt (section-dividers, logo, titel-topkant, positive tal, CTA) + Laks `#ed5e66` som signal + ØS-logo. Brug dette tema når brugeren beder om et ØS-deck.

## Workflow når brugeren beder om et deck

1. Læs `PRESENTATIONS.md` (brand-regler til slides).
2. Afklar kun det der mangler: emne, publikum, antal slides, HTML eller PPTX. Gæt fornuftigt.
3. Vælg tema: Primær præsentation stil (default) eller `oes` (se ovenfor). Spørg hvis det ikke er klart.
4. **HTML-deck:** opret en ny `.html`-fil baseret på `template.html` (primær) eller `template-oes.html` (ØS). Brug `slides.js`. Klar til browser og print-to-PDF.
5. **PowerPoint-deck:** skriv et ny `deck-<emne>.py` efter mønsteret i `deck.example.py` / `deck.oes.example.py`, kør det med `python3 deck-<emne>.py`. Genererer on-brand `.pptx`. Kræver `pip install python-pptx` (allerede installeret).
6. Brugeren skal **aldrig** gentage designreglerne — de er kodet ind i guiden og templatene.

## Filer

```
CLAUDE.md            — denne fil (entry point for Claude Code)
PRESENTATIONS.md     — fuld brand-guide: Primær præsentation stil (læs altid)
PRESENTATIONS-OES.md — Økonomistyrelsen-variant (kun forskellene)
slides.css           — al slide-styling, bygget på design-systemets tokens
slides-oes.css       — ØS-tema-overlay (importerer slides.css)
slides.js            — tastaturnavigation + fit-to-viewport (begge temaer)
template.html        — default demo-deck med alle slide-typer
template-oes.html    — ØS demo-deck med alle slide-typer
build_pptx.py        — Deck-helpers til .pptx (theme="default"|"oes")
deck.example.py      — primær eksempel-deck
deck.oes.example.py  — ØS eksempel-deck
assets/              — officielle ØS-logoer (grøn + hvid, transparente PNG)
```

## Relation til UI design systemet

Præsentationsguiden følger samme æstetik som `~/Desktop/Mortens Design system/morten-ds/` — samme farver, fonts (IBM Plex), shadows og signal-logik — men skaleret til 16:9-slides og projektor-afstand. De to systemer er bevidst holdt adskilt.
