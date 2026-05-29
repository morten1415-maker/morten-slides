# Primær præsentation stil — Brand-/Designguide til slides

> Dette er **"Primær præsentation stil"** — din standard slide-stil (HTML + PowerPoint),
> bygget på UI-designsystemets DNA. ØS-varianten ligger i `PRESENTATIONS-OES.md`.
> **Læs altid denne fil før du genererer et deck.** Når brugeren siger "lav et deck/slides om X",
> følg reglerne her og brug templates i denne mappe. Brugeren skal aldrig gentage designreglerne.
>
> Kode: `Deck(theme="default")` (alias `"primær"`/`"primary"`) · CSS: `slides.css` · template: `template.html`

## Workflow når brugeren beder om et deck

1. Læs denne fil + `../CLAUDE.md` (for det grundlæggende DNA).
2. Afklar kun det der mangler: **emne, publikum, antal slides, HTML eller PPTX** (gæt fornuftigt hvis ikke nævnt).
3. Skriv indholdet som struktureret outline (slide-type + indhold pr. slide).
4. **HTML:** kopiér `template.html`, udfyld med dine slides (samme `<section class="slide ...">`-mønstre).
5. **PPTX:** skriv en `deck.py` der kalder helpers i `build_pptx.py`, og kør den.
6. Hold dig til slide-typerne nedenfor. Opfind ikke nye layouts uden grund.

---

## Æstetisk DNA på slides — samme som UI, men forstørret

Systemet er **bærbar brutalisme**: sorte kanter, hårde offset-shadows (ingen blur),
IBM Plex (mono = system/metadata, sans = indhold), sparsom rød/gul.
På slides gælder præcis samme regler — men alt skaleres op til projektor-afstand.

De ufravigelige regler fra UI-guiden gælder også her:
- **Reservér de stærke elementer.** Maks. 1 rød-shadow og maks. 1 gul-flade per slide.
- **Mono er system, sans er indhold.** Mono er ALTID uppercase med 0.08–0.12em letter-spacing.
- **Sorte yderkanter, lysegrå indre delelinjer.** Border-radius: 0 på flader, 4px på interaktivt.
- **Offset-shadows kun.** Aldrig blurry shadows. Aldrig gradienter.
- **Tal er altid mono + tabular-nums.**

---

## Slide-format & canvas

- **Aspect ratio: 16:9.** HTML-canvas = `1280 × 720` px (skaleres til viewport). PPTX = `13.333 × 7.5 tommer`.
- Baggrund: `--bg-canvas` (#f4f4f2) som standard. Sektion-dividers må bruge sort eller gul fuldflade.
- **Sikker margin:** mindst 64px (≈0.66") luft til alle kanter. Indhold lever i et indre felt.
- Hver slide har en **footer-stribe** (mono, lille): venstre = deck-titel, højre = slide-nr `04 / 18`.

---

## Type-skala til slides (forstørret fra UI-skalaen)

Projektor kræver større tekst end en skærm-UI. Brug denne skala (px på 1280×720-canvas):

```
slide-display:  96px / 700 / -0.03em   → titel-slide hovedtekst
slide-title:    56px / 700 / -0.03em   → slide-overskrifter (h1 pr. slide)
slide-h2:       36px / 700 / -0.02em   → underoverskrift / sektion i slide
slide-lead:     28px / 400 / normal     → manchet/intro-sætning
slide-body:     22px / 400 / 1.4        → brødtekst, bullets
slide-stat:     88px / 700 / mono       → KPI-tal (tabular-nums)
slide-kicker:   16px / 600 / 0.12em uppercase MONO  → øjenbryns-label over titel
slide-footer:   13px / 600 / 0.1em uppercase MONO   → footer-metadata
```

Regel: **maks. ~6 bullets eller ~40 ord per slide.** Slides er ikke dokumenter.
Én pointe per slide. Hvis der er to pointer, lav to slides.

---

## Slide-typer (det faste bibliotek)

Hold dig til disse. De dækker stort set alt. HTML-klasse + PPTX-helper i parentes.

### 1. Title (`.slide--title` / `title_slide`)
Åbnings-slide. Mono-kicker øverst, stor `slide-display`-titel, evt. én manchetlinje,
footer med dato/forfatter i mono. Må have ét stort sort frame med offset-shadow om titlen.

### 2. Section divider (`.slide--section` / `section_slide`)
Markerer nyt kapitel. **Fuldflade-baggrund** (sort med hvid tekst, ELLER gul med sort tekst).
Stort sektion-nummer i mono (`01`, `02`) + sektion-titel. Brug sparsomt — det er et "åndedræt".

### 3. Content / bullets (`.slide--content` / `content_slide`)
Standard-sliden. Mono-kicker + `slide-title`, derefter en bullet-liste.
Bullets bruger en lille sort firkant-markør (ikke runde prikker). Maks. 6.

### 4. Two-column (`.slide--split` / `split_slide`)
To kolonner adskilt af 1px lysegrå linje. Fx tekst | liste, eller før | efter.
Hver kolonne har sin egen mono-label øverst.

### 5. Statement / quote (`.slide--statement` / `statement_slide`)
Én stor sætning eller citat, centreret, `slide-h2`/`slide-title`-størrelse.
Citater: ingen krøllede anførselstegn-grafik; brug en rød 4px venstre-kant i stedet.
Kilde i mono uppercase nederst.

### 6. Stat / KPI (`.slide--stats` / `stats_slide`)
1–3 store tal i bordede kort (2px sort border, 5px sort shadow).
Tallet i mono `slide-stat` + tabular-nums, label i mono under.
**Maks. ét kort må være fremhævet** (gul baggrund ELLER rød shadow) — det vigtigste KPI.

### 7. Table (`.slide--table` / `table_slide`)
Følger DataTable-reglerne: sort header, hvid mono-uppercase labels, 1px lysegrå rækkelinjer,
numeriske kolonner højrejusteret + mono + tabular-nums. Fremhævet række = gul baggrund.

### 8. Closing (`.slide--closing` / `closing()`)
Afslutning: "TAK" / kontakt / næste skridt. Spejler title-sliden. Må have den ene
tilladte rød-shadow på en primary "call to action"-boks.

### 9. Agenda / indhold (`.slide--agenda` / `agenda()`)
Nummereret indholdsfortegnelse. Mono-nummer + sans-titel pr. række, 1px delelinje.
Den aktuelle sektion må fremhæves (rødt nummer). Bruges også som kapitel-overblik.

### 10. Cover / hero med billede (`.slide--cover` / `cover()`)
Fuldt baggrundsbillede + hvid overlay-boks med stor shadow (kicker + titel + manchet).
Den stærkeste billed-åbning. Billede kan være placeholder.

### 11. Full-bleed billede (`.slide--image-full` / `image_full()`)
Billedet fylder hele sliden; sort caption-stribe nederst med mono-kicker + titel.
Til ét stærkt foto der skal tale for sig selv.

### 12. Billede + tekst (`.slide--image-text` / `image_text()`)
Halv-til-halv: billed-felt på den ene side, bullets/tekst på den anden.
`is-image-left` / `image_left=True` spejler. Billede får 2px sort kant.

### 13. Billed-grid / galleri (`.slide--image-grid` / `image_grid()`)
2×2 eller 3-op (`cols-3`) af billed-felter med ens størrelse og 2px kanter.
Til screenshots, produktbilleder, moodboard.

### 14. Feature-kort (`.slide--cards` / `cards()`)
2–4 bordede kort med ikon-plads (sort firkant), titel + kort tekst.
Maks. ét kort fremhævet (gul flade + rød ikon).

### 15. Timeline / proces (`.slide--timeline` / `timeline()`)
Vandrette trin med firkant-markører + forbindelseslinje. Mono-tid + titel + tekst.
Det aktuelle trin må fremhæves rødt.

### 16. Sammenligning (`.slide--compare` / `comparison()`)
To bordede kolonner med check (grøn firkant) / kryds (rød firkant) pr. punkt.
"Vinder"-kolonnen må have den ene tilladte rød-shadow.

### 17. Stort tal (`.slide--bignum` / `bignum()`)
Ét kæmpe mono-tal (tabular-nums) i fuld fokus + kort undertekst. Vækst grøn, fald rød.
Når ét nøgletal er hele pointen.

### 18. Søjlediagram (`.slide--bars` / `bars()`)
Vandrette søjler: label + sort fyld-bar + mono-værdi. Maks. én bar fremhævet (rød).
Brutalist-erstatning for et farvet diagram. Tal altid mono + tabular-nums.

### 19. Testimonial (`.slide--testimonial` / `testimonial()`)
Portræt-felt (firkant) + stort citat med rød venstre-kant + navn (sans) og rolle (mono).

> **Billeder:** alle billed-typer viser en grå placeholder med mono-label indtil et
> rigtigt billede sættes ind. HTML: byt `.img`-div ud med `<img>`. PPTX: send `image="sti.jpg"`.

---

## Farvebrug på slides (signal-logik, ikke dekoration)

Samme trafiklys-logik som UI:
- **Sort:** alle yderkanter, primær tekst, sektion-dividers, table-headers.
- **Gul (`--bg-accent`):** ÉN fremhævet flade per slide — aktivt KPI-kort, fremhævet tabelrække, en gul section-divider. Aldrig "bare for farvens skyld".
- **Rød (`--color-red`):** KUN til signal — primary CTA-shadow, "vigtigt/fald" i tal, citat-kant, danger. Maks. én rød-shadow per slide.
- **Grøn (`--color-success`):** positiv/færdig/vækst i tal og status-prikker.

Aldrig: farvede gradienter, mere end to accentfarver synlige på samme slide, farvet tekst som brødtekst.

---

## Hvad du IKKE skal gøre på slides

- ❌ Aldrig blurry/soft shadows — kun hårde offsets (sort eller rød).
- ❌ Aldrig vægtede tekstvægge. Maks. ~40 ord. Slides støtter tale, erstatter den ikke.
- ❌ Aldrig mono i normal case eller uden letter-spacing.
- ❌ Aldrig border-radius på flader/kort/dividers (0px). 4px kun på knap-agtige elementer.
- ❌ Aldrig runde bullet-prikker — brug små sorte firkanter.
- ❌ Aldrig mere end én rød-shadow eller én gul-flade per slide.
- ❌ Aldrig stock-foto-clutter eller emoji som primær ikon-strategi.
- ❌ Aldrig proportionale tal i tal-kontekst — altid tabular-nums + mono.

## Når du er i tvivl

Spørg: "Ville denne slide se rolig og bevidst ud fra bagerste række?
Og bærer den stadig systemets karakter (sort kant, hård shadow, mono-label)?"
Hvis ja til begge — kør på. Bryd kun reglerne hvis brugeren beder om det, og sig hvilken regel du bryder.

## Filer i denne mappe

```
PRESENTATIONS.md   — denne guide (læs altid først)
slides.css         — al slide-styling, bygget på ../styles/tokens.css
slides.js          — tastaturnavigation + fit-to-viewport-skalering
template.html      — demo-deck med ét eksempel af hver slide-type (kopiér denne)
build_pptx.py      — helpers til at generere on-brand .pptx
deck.example.py    — eksempel på hvordan man bygger et PPTX-deck med helperne
```
