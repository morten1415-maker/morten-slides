# Mortens Præsentationer — Instruktioner til Claude Code

Dette projekt genererer on-brand slidedeck — HTML-slides og PowerPoint — der følger Mortens personlige design system ("bærbar brutalisme").

**Læs `PRESENTATIONS.md` før du genererer et deck.** Den indeholder de fulde brand-regler oversat til slides: type-skala, farvebrug, slide-typer, do/don'ts.

## ⚠️ VIGTIGSTE REGEL — byg ud fra INDHOLDET, ikke ud fra templaten

`template.html`, `template-oes.html`, `deck.example.py` og `deck.oes.example.py` er **demo-gallerier / komponentbiblioteker** — de viser ÉT eksempel på *hver* tilgængelig slide-type. **De er IKKE en færdig deck der skal kopieres.**

Når brugeren beder om "et deck om X":

- ❌ **Lav IKKE** en kopi af galleriet med samme slides, samme antal og samme rækkefølge.
- ❌ **Brug ikke** alle slide-typer bare fordi de findes.
- ✅ **Start fra indholdet/formålet.** Find budskabet og en fortælling, og *vælg derudfra*:
  - **Hvilke** slide-typer der giver mening (drop resten).
  - **Hvor mange** slides emnet kræver (kan være 4, kan være 30 — galleriets ~21 er tilfældigt).
  - **Hvilken rækkefølge** der bedst bærer fortællingen.
- ✅ **Opfind nye slides/layouts** når indholdet kræver noget biblioteket ikke har — så længe du holder dig til DNA'et (sorte kanter, hårde offset-shadows, IBM Plex, mono = system/sans = indhold, sparsom rød/grøn, 0px på flader / 4px på interaktivt, tokens fra `slides.css`). Genbrug eksisterende klasser/helpers hvor du kan; udvid systemet konsistent hvor du ikke kan.

Kort sagt: **galleriet er paletten, ikke maleriet.** Stilen er fast — indholdet bestemmer hvilke og hvor mange slides.

**To temaer:**
- **Primær præsentation stil** (standard, Mortens egen stil) — `PRESENTATIONS.md`, `slides.css`, `template.html`, `Deck(theme="default")` (alias `"primær"`).
- **Økonomistyrelsen** — `PRESENTATIONS-OES.md`, `slides-oes.css`, `template-oes.html`, `Deck(theme="oes")`. = Primær stil + ØS-grøn `#066b43` brugt sparsomt (section-dividers, logo, titel-topkant, positive tal, CTA) + Laks `#ed5e66` som signal + ØS-logo. Brug dette tema når brugeren beder om et ØS-deck.

## Workflow når brugeren beder om et deck

1. Læs `PRESENTATIONS.md` (brand-regler til slides) + læs VIGTIGSTE REGEL ovenfor.
2. Afklar kun det der mangler: emne, publikum, HTML eller PPTX, tema. Gæt fornuftigt. (Spørg IKKE om antal slides — det udleder du af indholdet.)
3. **Lav en outline FØRST:** skriv fortællingen som en liste af slides — for hver: hvilken slide-type (eller ny idé) + det konkrete indhold. Vælg kun det der tjener budskabet. Det er her arbejdet ligger.
4. Vælg tema: Primær præsentation stil (default) eller `oes`.
5. **HTML-deck:** opret en ny `.html`-fil. Lås `<link>` til `slides.css` (primær) eller `slides-oes.css` (ØS) + `slides.js`. Byg DINE slides fra outlinen ved at bruge slide-type-klasserne (slå dem op i `template*.html` som reference) — ikke ved at kopiere hele filen. Tilføj nye `.slide--*`-klasser i en `<style>` eller egen CSS hvis indholdet kræver et nyt layout.
6. **PowerPoint-deck:** skriv et nyt `deck-<emne>.py` der bruger `Deck`-helperne i `build_pptx.py` og kalder netop de slide-metoder din outline kræver, i din rækkefølge. Mangler der en slide-type, så tilføj en ny metode i `build_pptx.py` i samme stil. Kør `python3 deck-<emne>.py`. Kræver `pip install python-pptx` (allerede installeret).
7. Brugeren skal **aldrig** gentage designreglerne — de er kodet ind i guiden, tokens og helpers.

## Filer

```
CLAUDE.md            — denne fil (entry point for Claude Code)
PRESENTATIONS.md     — fuld brand-guide: Primær præsentation stil (læs altid)
PRESENTATIONS-OES.md — Økonomistyrelsen-variant (kun forskellene)
slides.css           — al slide-styling, bygget på design-systemets tokens
slides-oes.css       — ØS-tema-overlay (importerer slides.css)
slides.js            — tastaturnavigation + fit-to-viewport (begge temaer)
template.html        — GALLERI/komponentbibliotek (primær): ét eksempel pr. slide-type — IKKE en deck at kopiere
template-oes.html    — GALLERI/komponentbibliotek (ØS): samme, men ØS-tema
build_pptx.py        — Deck-helpers til .pptx (theme="default"|"oes") — udvid med nye metoder ved behov
deck.example.py      — GALLERI/eksempel (primær): demonstrerer alle helpers — IKKE en deck at kopiere
deck.oes.example.py  — GALLERI/eksempel (ØS): samme
assets/              — officielle ØS-logoer (grøn + hvid, transparente PNG)
```

## Relation til UI design systemet

Præsentationsguiden følger samme æstetik som `~/Desktop/Mortens Design system/morten-ds/` — samme farver, fonts (IBM Plex), shadows og signal-logik — men skaleret til 16:9-slides og projektor-afstand. De to systemer er bevidst holdt adskilt.
