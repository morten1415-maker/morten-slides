# Økonomistyrelsen — Brand-/Designguide til slides

> Dette er **den eneste slide-stil** i systemet. **Læs altid denne fil før du genererer et deck.**
> Brugeren skal aldrig gentage designreglerne.
>
> Kode: `Deck()` i `build_pptx.py` · CSS: `slides.css` · gallerier: `template.html` og `deck.example.py`

Stilen er udmålt på præsentationen *UX og designprocessen* (SBI-kontormødet): lys grå
grund, titler i sorte blokke med grøn skygge, kapitel-labels i laks, billeder med en grøn
flade forskudt bagved, IBM Plex Sans til indhold og IBM Plex Mono til alt "system".

## Workflow når brugeren beder om et deck

> **Galleriet er paletten, ikke maleriet.** `template.html` og `deck.example.py` viser ÉT eksempel
> på hver slide-type — de er et bibliotek, IKKE en deck du kopierer. Stilen er fast; *indholdet*
> bestemmer hvilke og hvor mange slides.

1. Læs denne fil + `CLAUDE.md`.
2. Afklar kun det der mangler: **emne, publikum, HTML eller PPTX** (gæt fornuftigt). Spørg IKKE om antal slides eller tema.
3. **Skriv en outline FØRST:** fortællingen som en liste — for hver slide en type (eller en ny idé) + konkret indhold.
4. **PPTX:** skriv en ny `deck-<emne>.py` der kalder netop de `Deck`-metoder outlinen kræver. Gem med `preview_dir=`, og **se på PNG'erne** før du afleverer.
5. **HTML:** lav en NY `.html`, link `slides.css` + `slides.js`, og byg dine slides med klasserne fra `template.html`.
6. **Opfind nye layouts** når indholdet kræver det — med byggeklodserne og DNA'et nedenfor.
7. Kør `python verify.py mit-deck.pptx` (geometri, fonte, lag og kontrast).

---

## DNA — de regler der gør et slide til "et af vores"

| # | Regel | Hvordan |
|---|---|---|
| 1 | **Overskriften er et sort skilt med grøn skygge, øverst til venstre.** Én pr. slide | `Deck._title()` / `<h2 class="slide-title">` — top 64 px, venstre 88 px. Placér den aldrig selv |
| 2 | **Højst én "hård" ting pr. slide.** Har sliden et billede med grøn flade, står overskriften fri — ingen blok | `media_text` / `hero` / `image_text` gør det af sig selv (`.slide-h2` i HTML) |
| 3 | **Billeder har ingen kant — de har en grøn flade forskudt 24 px bagved** | `picture(…, side="right")` / `class="media"`. Frilagte motiver (`free=True` / `is-free`) har ingen flade |
| 4 | **Mono er system, sans er indhold.** Labels, numre, tal, kilder = IBM Plex Mono. Alt man læser = IBM Plex Sans | `mono()`, `num()`, `head()`, `body()` |
| 5 | **Mono-labels er versaler med sporing (0,06 em)** — undtagen tal, kilde-linjer og e-mails | `mono(text)` / `mono(text, caps=False)` |
| 6 | **Hårde skygger, aldrig blur. Aldrig gradienter. Aldrig runde hjørner** | Sort skygge 8 px på kort; grøn skygge på titelblokke |
| 7 | **Sorte 3 px kanter på kort; tynde grå linjer indeni** | `card()` / `.card` |
| 8 | **Lister har små sorte firkanter — aldrig runde prikker** | `square_list()` / `.bullets` |
| 9 | **Logoet sidder i højre hjørne** med 32 px luft til højre og bund | Kommer af sig selv. Fuldflade-billeder har intet logo |
| 10 | **Indhold starter på 176 px og slutter senest på 640 px** | `CONTENT_TOP` / `CONTENT_BOT` |

### Farver — hver farve har ét job

| Token | Hex | Job |
|---|---|---|
| `GROUND` | `#F4F4F2` | Slidens baggrund. Altid |
| `BLACK` | `#000000` | Titelblokke, kanter, skygger, søjler |
| `INK` | `#141414` | Tekst |
| `DIM` | `#6E6E6E` | Nedtonet tekst: kilde, meta, inaktive trin (4,6:1) |
| `BRAND` | `#066B43` | **ØS-grøn:** skyggen bag titelblokke, fladen bag billeder, det *ene* fremhævede element (aktivt trin, vinder, positivt tal) |
| `HILITE` | `#B3E8CA` | Lysegrøn overstregning i tekst, fremhævet tabelrække / nøgletalskort |
| `LAKS` | `#ED5E66` | Signal på **store flader og på sort**: kapitel-label, citatstreg, delnummer i titelblokken |
| `LAKS_TEXT` | `#C93A44` | Laks til **tekst på lys grund**: mono-labels, numre, negative tal (`LAKS` har kun 3:1 der) |
| `STICKY` | `#FFF9B1` | Kun post-its |

- **Grøn er det fremhævede, laks er systemets stemme.** Fremhæv højst ét element pr. slide.
- Gul bruges ikke længere som fremhævning (den gamle `#FEF3C7`) — brug `HILITE`.
- Der er ingen mørke eller grønne helflade-slides længere. Kapitler er en laks-label + sort titelblok.

### Typeskala (px på 1280 × 720 — gang med 0,75 for pt)

| Rolle | Størrelse | Brug |
|---|---|---|
| Display | 64 fed | Forside, kapitel, tak (i sort blok) |
| Hero | 60 fed | Fri overskrift ved siden af billede på forsiden |
| Fri overskrift | 44 fed | `media_text` — skrumper selv så længste ord kan stå på én linje |
| Udsagn / citat | 44 / 40 regular | Alene / under en titel |
| Titelblok | 31 fed | Slide-overskrift |
| Manchet | 24 | Lead under display |
| Brødtekst | 22 | Lister og tekst |
| Lille | 18 | Tætte spalter, diagram-tekst |
| Mono-label | 13–15 fed | Eyebrow, numre, akser |
| Kilde | 12 mono | Nederst til venstre |

Linjehøjde: 1,05–1,2 på overskrifter, 1,4–1,45 på brødtekst.

---

## Faste mål

Samme tal i PPTX og HTML — alt måles i px på et 1280 × 720-canvas (1 px = 0,75 pt = 9525 EMU).

```
                               px      build_pptx.py      slides.css
sidemargin                     88      PAD_X              --pad-x
titelblok, top                 64      TITLE_Y            --title-y
titelblok, tekst               31      TITLE_SIZE         .slide-title
titelblok, grøn skygge          7      TITLE_SHADOW       .slide-title
indhold starter               176      CONTENT_TOP        --content-top
indhold slutter               640      CONTENT_BOT        --pad-bottom (80)
billedramme (kvadrat)         470      MEDIA              --media
billedramme, top              125      MEDIA_Y            .slide--image-text
grøn flade, forskydning        24      OFFSET             --offset
kant / hård skygge          3 / 8      BORDER / SHADOW    --border / --shadow
spalteluft                     40      GAP                --gap
logo (bredde / hjørne-luft) 124 / 32   LOGO_W / LOGO_INSET  .slide::after
kilde-linje, underkant        682      SOURCE_BOTTOM      .slide__source
```

Ret dem centralt — aldrig pr. slide. `verify.py` holder de to filer op mod hinanden og mod
referencetallene i `REFERENCE_PX`.

---

## Slide-typer

`deck = Deck()` — hver metode returnerer sliden, så du kan tegne videre på den.

### Åbning, kapitler, afslutning

| Metode | HTML-klasse | Hvad |
|---|---|---|
| `title(title, lead, meta, image=None, eyebrow=None)` | `slide--title` / `slide--hero` | Forside. Uden billede: stor sort blok. Med `image`: tekst til venstre, billede med grøn flade til højre |
| `hero(title, image, eyebrow, lead, side, meta)` | `slide--hero` | Billede + stor fri overskrift — forside og store øjeblikke midt i decket |
| `agenda(items, title, current)` | `slide--agenda` | Nummereret liste mellem sorte streger; `current` får overstregning |
| `section(title, num="01")` | `slide--section` | Laks-label "Del 1" + sort titelblok. **Sætter "01" foran titlerne på de følgende slides** |
| `closing(title, lines, cta)` | `slide--closing` | Sort blok + linjer + grøn mono-label |

### Tekst og struktur

| Metode | HTML-klasse | Hvad |
|---|---|---|
| `content(title, bullets, accent_index, sub)` | `slide--content` | Titel + firkant-liste. Uden `bullets`: tom slide du selv tegner på. Et punkt kan være `("Fed start: ", "resten")` |
| `split(title, left, right, arrow)` | `slide--split` | To spalter under sorte streger: `{"label", "head", "body", "accent"}`. `arrow=True` = før → efter |
| `cards(items, title, accent_index)` | `slide--cards` | 2–4 nummererede spalter: `{"title", "body"}` |
| `comparison(left, right, title, win)` | `slide--compare` | To kort med +/– : `{"head", "items": [(True, "tekst")]}`. Vinderen får skygge |
| `timeline(steps, title, current)` | `slide--timeline` | Vandret forløb: `{"when", "what", "desc"}` |
| `process(steps, title, current, only)` | `slide--process` | Lodret nummereret proces; `current="Navn"` er grøn, resten nedtones |

### Citater

| Metode | HTML-klasse | Hvad |
|---|---|---|
| `statement(quote, attribution, title, highlight_part, eyebrow)` | `slide--statement` | Laks-streg, ingen kasse. `highlight_part="…"` overstreger en del (`<mark>` i HTML) |
| `testimonial(quote, name, role, image)` | `slide--testimonial` | Portræt med grøn flade + citat |

Vis samme citat flere gange med forskellig `highlight_part` for at gennemgå det led for led.

### Billeder

| Metode | HTML-klasse | Hvad |
|---|---|---|
| `image_text(title, bullets, image, image_left)` | `slide--image-text` | Kvadratisk billede med grøn flade + fri overskrift og punkter |
| `media_text(title, image, body_paras, bullets, eyebrow, side, free)` | `slide--image-text` | Den generelle udgave. `free=True`: frilagt PNG uden flade |
| `image_grid(title, images, cols)` | `slide--image-grid` | Gitter af billeder |
| `cover(title, lead, image)` | `slide--cover` | Fuldt billede + sort titelblok nederst |
| `image_full(caption, label, image)` | `slide--image-full` | Fuldt billede + billedtekst i sort blok |

Mangler billedet (`image=None`), tegnes en grå pladsholder. Sæt altid `alt=` på rigtige billeder.
`focus=(x, y)` (0–1) styrer beskæringen.

### Tal

| Metode | HTML-klasse | Hvad |
|---|---|---|
| `bignum(value, sub, trend, source)` | `slide--bignum` | Ét stort tal. `trend="up"` grøn, `"down"` laks |
| `stats(cards, title, primary_index, source)` | `slide--stats` | Nøgletal i hvide kort: `{"num", "label", "trend"}` |
| `bars(rows, title, accent_index, source)` | `slide--bars` | Vandrette søjler: `{"label", "pct", "value"}` |
| `table(headers, rows, title, numeric_cols, highlight_row, source, col_widths)` | `slide--table` | Sort hovedrække, hvide rækker, kant + skygge |

**Kilde på tal:** `source="Kilde: …"` giver en mono-linje nederst til venstre
(`<p class="slide__source">` i HTML — virker på enhver slide). Sæt den når tallene ikke er
brugerens egne. Har du kun tal uden kilde, så spørg efter kilden.

---

## Byg selv — byggeklodserne

Når ingen slide-type passer (diagrammer, modeller, post-it-øvelser):

```python
from build_pptx import (Deck, CONTENT_TOP, CONTENT_BOT, PAD_X, W, BLACK, BRAND, LAKS_TEXT, DIM, INK,
                        card, block, chip, chip_w, square_list, sticky, picture,
                        mono, num, head, body, para)

s = deck.content("Double Diamond")            # titelblok + tom flade fra CONTENT_TOP
s.path([("M", x1, y1), ("L", x2, y2), ("Z",)], color=BLACK, lw=3)     # frihåndsform
s.line(x1, y1, x2, y2, color=BLACK, lw=2, tail="triangle")            # pil
s.ellipse(x, y, w, h, fill=BRAND, line=BLACK, lw=3)
s.text(x, y, w, h, [para(mono("Label", 12, DIM), align="ctr")])
chip(s, x, y, "Problem")                      # lille sort label
card(s, x, y, w, h)                           # hvid boks, sort kant, hård skygge
sticky(s, x, y, 236, [body("Er "), body("nem", bold=True), body(" at bruge")])
square_list(s, x, y, w, ["Punkt 1", ("Fed: ", "resten")])
picture(s, "foto.jpg", x, y, w, h, side="right", alt="…")             # med grøn flade
```

Regler for egne layouts:
- Diagram-streger er sorte, 2–3 px. Kun ét element er grønt (det du taler om).
- Etiketter i diagrammer er mono, 12–13 px, `DIM` eller `INK`.
- Byg en model op **slide for slide** (samme layout, ét nyt element ad gangen) i stedet for animation.
- `deck.save()` advarer, hvis tekst ikke kan være i sin boks. Ret teksten eller boksen — ignorér den ikke.
- `s.notes = "…"` sætter talernoter.

I HTML: tilføj nye `.slide--*`-klasser i en `<style>` i decket, og brug tokens fra `:root`.

---

## Fonte — IBM Plex følger med filen

PowerPoint refererer fonte ved navn. Er IBM Plex ikke installeret på maskinen der åbner
filen, erstattes den lydløst, og mono/sans-forskellen forsvinder. Derfor **indlejres fonten
i hver genereret .pptx** (som EOT — en rå TTF giver fejl i PowerPoints font-dialog).

- Fontfilerne ligger i `assets/fonts/` (IBM Plex v6.4.0, OFL). `LICENSE.txt` **skal følge med**.
- `slides.css` bruger de samme filer via `@font-face` — HTML-slides virker offline.
- Slå indlejring fra med `Deck(embed_fonts=False)` (sparer ca. 0,5 MB).

---

## Don'ts

- ❌ Ingen grøn accent-streg under overskrifter (den gamle stil) — titlen *er* blokken.
- ❌ Ingen grønne eller sorte helflade-slides. Ingen gule flader (kun post-its).
- ❌ Ingen kant om billeder, ingen blur-skygger, ingen gradienter, ingen runde hjørner.
- ❌ Ingen eyebrow/kicker over en titelblok. (Over en *fri* overskrift ved et billede er den fin.)
- ❌ Ikke to titelblokke på én slide, og ikke titelblok + billede med grøn flade i `media_text`.
- ❌ Ingen runde bullets, ingen farvede ikoner som pynt.
- ❌ Ingen `LAKS` (`#ED5E66`) som tekstfarve på lys grund — brug `LAKS_TEXT`.
