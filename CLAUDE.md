# Økonomistyrelsens slide-system — instruktioner til Claude Code

Dette projekt genererer on-brand slidedeck — PowerPoint og HTML — i **Økonomistyrelsens**
stil: lys grå grund, titler i sorte blokke med grøn skygge, laks-labels, billeder med en
grøn flade forskudt bagved, IBM Plex Sans + Mono.

**Læs `PRESENTATIONS.md` før du genererer et deck.** Den indeholder de fulde regler:
DNA, farver, typeskala, faste mål, alle slide-typer og byggeklodserne.

## ⚠️ Der er ÉN stil

`slides.css`, `template.html` og `Deck()` ER stilen. Vælg ikke tema, spørg ikke om tema, og
lav ikke varianter. (`Deck(theme=...)` og `kicker=` tages stadig imod, men ignoreres — de
findes kun så gamle deck-scripts ikke går i stykker.)

## ⚠️ De fem fejl der bliver gentaget

De er kodet ind i `build_pptx.py` og `slides.css`. **Brug metoderne — placér dem ikke selv.**

1. **Overskriften er et sort skilt med grøn skygge, øverst til venstre** (top 64 px, venstre
   88 px). Ikke en fri tekst, ikke midt på sliden, ikke med en streg under.
2. **Højst én "hård" ting pr. slide.** Har sliden et billede med grøn flade (`media_text`,
   `hero`, `image_text`), står overskriften FRI ved siden af — ingen titelblok.
3. **Billeder har ingen kant — de har en grøn flade forskudt 24 px bagved.** Brug
   `picture(…, side=…)` / `class="media"`; tegn ikke fladen selv.
4. **Mono er system, sans er indhold.** Labels, numre, tal og kilder er IBM Plex Mono
   (`mono()`, `num()`); alt man læser er IBM Plex Sans. Laks-tekst på lys grund er
   `LAKS_TEXT`, aldrig `LAKS`.
5. **Kun ÉN overskrift pr. slide** og ingen eyebrow over en titelblok. Kapitlets nummer
   kommer af sig selv foran titlen, når du bruger `deck.section(…, num="01")`.

## ⚠️ Byg ud fra INDHOLDET, ikke ud fra galleriet

`template.html` og `deck.example.py` er **komponentbiblioteker** — ét eksempel pr. slide-type.
**De er IKKE en færdig deck der skal kopieres.**

- ❌ Lav ikke en kopi af galleriet med samme slides og rækkefølge. Brug ikke alle typer.
- ✅ Start fra indholdet: find budskabet og fortællingen, og vælg derudfra hvilke typer,
  hvor mange slides og hvilken rækkefølge.
- ✅ **Opfind nye layouts** når indholdet kræver det (diagrammer, modeller):
  `s = deck.content("Titel")` og tegn fra `CONTENT_TOP` med byggeklodserne (`card`, `block`,
  `chip`, `square_list`, `picture`, `s.line/rect/ellipse/path/text`). Se afsnittet
  "Byg selv" i `PRESENTATIONS.md`.
- ✅ Byg modeller op slide for slide (samme layout, ét nyt element) i stedet for animation.

## Workflow når brugeren beder om et deck

1. Læs `PRESENTATIONS.md`.
2. Afklar kun det der mangler: emne, publikum, HTML eller PPTX. Gæt fornuftigt.
3. **Lav en outline FØRST:** for hver slide en type (eller ny idé) + det konkrete indhold.
4. **PowerPoint:** skriv `deck-<emne>.py` med `from build_pptx import Deck`, og gem med
   `deck.save("navn.pptx", preview_dir="preview-navn")`.
5. **Se på PNG'erne i `preview_dir`** — det er sådan du ser sliden uden PowerPoint. Ret
   overlap, tomme flader og tekst der står for tæt. `save()` skriver `ADVARSEL:` hvis tekst
   ikke kan være i sin boks; ret det, ignorér det ikke.
6. Kør `python verify.py navn.pptx` (geometri, fonte, lag, kontrast).
7. **HTML:** ny `.html` der linker `slides.css` + `slides.js`; byg med klasserne fra `template.html`.

Kræver `pip install python-pptx pillow fonttools lxml`.

## Kilde på tal

`stats`, `table`, `bars`, `bignum` og `statement` tager `source="Kilde: …"`. I HTML:
`<p class="slide__source">`. **Sæt den når tallene ikke er brugerens egne.** Har du kun
tal uden kilde, så spørg efter kilden.

## Fonte

IBM Plex **indlejres** i hver genereret .pptx (`assets/fonts/`, pakket som EOT), og
`slides.css` bruger de samme filer via `@font-face`. Slet ikke `assets/fonts/LICENSE.txt` —
OFL kræver at licensen følger fonten.

## Kør verify.py efter enhver ændring i build_pptx.py, oes_pptx.py eller slides.css

```bash
python verify.py
```

Det bygger galleriet (+ `preview/`) og tjekker referencemålene, geometrien i filen,
font-indlejringen, HTML/PPTX-pariteten (mål og farver), galleri-dækningen, lag og kontrast,
og at ingen tekst flyder ud af sin boks. Exit-kode 1 ved fejl.

- Ny slide-type i `build_pptx.py` → kald den i `deck.example.py`, tilføj den i
  `template.html` + `slides.css`, og map den i `HTML_CLASS` i `verify.py`.
- Ændrer du bevidst et fast mål → ret det i `build_pptx.py`/`oes_pptx.py`, i `:root` i
  `slides.css` OG i `REFERENCE_PX` i `verify.py`.

## Filer

```
CLAUDE.md          — denne fil
PRESENTATIONS.md   — den fulde guide (læs altid)
build_pptx.py      — STILEN: tokens, byggeklodser og Deck med alle slide-typer
oes_pptx.py        — MOTOREN: elementer, tekst-ombrydning, .pptx-skrivning, PNG-preview, font-indlejring
deck.example.py    — galleri i PPTX-form -> example.pptx + preview/
slides.css         — al HTML-styling (samme mål og farver som build_pptx.py)
slides.js          — tastaturnavigation, fit-to-viewport, delnumre i titelblokke
template.html      — galleri i HTML-form
verify.py          — regressionstjek af hele systemet
tjek_pptx.py       — lag- og kontrasttjek af en vilkårlig .pptx (bruges af verify.py)
galleri.png        — oversigt over alle slide-typer
assets/            — ØS-logoer + IBM Plex (fonts/)
slides-oes.css     — udfaset shim: importerer blot slides.css
```
