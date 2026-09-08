# Økonomistyrelsen — Brand-/Designguide til slides

> Dette er **den eneste slide-stil** i systemet: Økonomistyrelsen. Der findes ikke
> længere en "primær" og en "ØS"-variant at vælge imellem — der er én stil, og den er ØS'.
> **Læs altid denne fil før du genererer et deck.** Brugeren skal aldrig gentage designreglerne.
>
> Kode: `Deck()` · CSS: `slides.css` · galleri: `template.html`

## Workflow når brugeren beder om et deck

> **Galleriet er paletten, ikke maleriet.** `template.html` og `deck.example.py` viser ÉT eksempel
> på hver slide-type — de er et bibliotek, IKKE en deck du kopierer. Stilen er fast; *indholdet*
> bestemmer hvilke og hvor mange slides.

1. Læs denne fil + `CLAUDE.md`.
2. Afklar kun det der mangler: **emne, publikum, HTML eller PPTX** (gæt fornuftigt). Spørg IKKE om antal slides — det følger af indholdet.
3. **Skriv en outline FØRST** ud fra formålet: fortællingen som en liste — for hver slide en type (eller en ny idé) + konkret indhold. Vælg kun det der tjener budskabet; drop resten.
4. **HTML:** lav en NY `.html`, link `slides.css` + `slides.js`, og byg DINE slides med slide-type-klasserne (slå dem op i `template.html` som reference). Kopiér ikke hele galleriet.
5. **PPTX:** skriv en ny `deck-<emne>.py` der kalder netop de `Deck`-helpers din outline kræver, og kør den.
6. **Opfind nye layouts** når indholdet kræver det — så længe DNA'et holdes (se nedenfor).

---

## ⚠️ De fejl der bliver gentaget — læs disse fem først

Det er de her fem der skal rettes i hånden bagefter hvis de bliver lavet forkert.
De er kodet ind i `slides.css` og `build_pptx.py`: **brug helperne, så rammer du dem automatisk.**

| # | Regel | Hvordan |
|---|---|---|
| 1 | **Overskriften sidder ØVERST** — ikke skubbet ned midt på sliden | PPTX: `TITLE_TOP` = 0,50" (1,27 cm). HTML: første element i sliden. Brug `Deck._title()` / `<h2 class="slide-title">` |
| 2 | **Den grønne linje hører til overskriften** — den skal sidde tæt under den, aldrig svæve mellem overskrift og indhold | PPTX: `ACCENT_Y` = 1,142" (2,90 cm), altid. HTML: `.slide-title::after` — kommer af sig selv. Placér den ALDRIG manuelt |
| 3 | **Den grønne linje flugter med overskriftens venstrekant** | PPTX: `ACCENT_X` = `MARGIN + 0.1"` (tekstboksens indryk). HTML: automatisk |
| 4 | **Logoet sidder i højre HJØRNE** med ens luft til højre og bund (0,28" / 27px) — ikke inde på tekstmargin | PPTX: `Deck._footer()`. HTML: `.slide::after` — kommer af sig selv. Tegn aldrig logoet selv |
| 5 | **Kun ÉN overskrift pr. slide** — den sorte titel. Ingen kicker-eyebrow over den | `kicker=`-argumenter findes stadig i API'et, men ignoreres. Skriv dem ikke |

**Én grøn linje pr. slide, og kun under en `slide-title`.** Sektions-dividers, title-slide, cover
og closing har ingen grøn linje — de bærer grøn på andre måder (fuldflade, shadow, CTA).

---

## Vertikal rytme (fast — pr. slide må den ikke justeres)

Samme rytme i HTML og PPTX, så de to formater ser ens ud:

```
                       PPTX (13.333 x 7.5")     HTML (1280 x 720 px)
overskrift, top            0.50"                    64px  (slide-padding)
grøn accent, top           1.142"                   tæt under titlen (10px)
indhold starter            1.70"                    +30px under accenten (fast)
indhold slutter            6.60"                    642px (padding-bottom 78px)
logo (h x b)               0.42" x 1.55"            41 x 150px
logo-hjørne (h/b-luft)     0.28"                    27px
sikker margin              0.66"                    64px
```

Konstanterne står i toppen af `build_pptx.py` (`TITLE_TOP`, `ACCENT_Y`, `CONTENT_TOP`,
`CONTENT_BOT`, `LOGO_INSET` …) og i `:root` i `slides.css`. **Ret dem der — ikke pr. slide.**

### Forsidens kasse (faste mål)

Forsidens hvide kasse må **ikke** hugge teksten — den skal være den samme brede,
rolige flade hver gang. Målene er identiske i HTML og PPTX:

```
                       PPTX (13.333 x 7.5")     HTML (1280 x 720 px)
kasse, venstre             1.68 cm / 0.66"          64px
kasse, top                 4.32 cm / 1.70"          163px
kasse, bredde              27.94 cm / 11.0"         1056px
kasse, højde (m. manchet)  10.16 cm / 4.0"          384px
kasse, højde (kun titel)    9.14 cm / 3.6"          346px
indre luft                 0.6" vandret / 0.55"     58px / 53px
overskrift                 58pt                     77px
manchet                    22pt                     29px
grøn offset-shadow         8pt ned-højre            8px 8px
meta-linje under kassen    0.38" under              36px under
```

Teksten er lodret centreret i kassen. Kassen har ingen grøn accent-linje —
den grønne offset-shadow er forsidens brand-signal.
`title(..., meta="Økonomistyrelsen · Ledelsesoverblik · August 2026")` sætter
mono-linjen under kassen (dato / anledning / afsender). I HTML: `<p class="meta">`
som søskende til `.frame`.

---

## Æstetisk DNA — bærbar brutalisme + ØS-grøn

Sorte kanter, hårde offset-shadows (ingen blur), IBM Plex (mono = system/metadata,
sans = indhold), sparsom gul/laks — og ØS-grøn som identitets-anker.
Alt skaleres op til projektor-afstand.

- **Reservér de stærke elementer.** Maks. 1 laks-shadow og maks. 1 gul flade per slide.
- **Mono er system, sans er indhold.** Mono er ALTID uppercase med 0.08–0.12em letter-spacing.
- **Sorte yderkanter, lysegrå indre delelinjer.** Border-radius: 0 på flader, 4px på interaktivt.
- **Offset-shadows kun.** Aldrig blurry shadows. Aldrig gradienter.
- **Tal er altid mono + tabular-nums.**

### Hvad der bevidst IKKE er taget fra ØS' brandguide

- ❌ ØS' fonte — vi beholder IBM Plex (stærkere, og ØS-fonten lå ikke i guiden).
- ❌ ØS' Templafy-layouts — vores brutalist-struktur er mere distinkt.
- ❌ Den fulde pastel-palette (himmelblå, latte, gråblå osv.) — ville udvande brutalismen.
  Kun grøn-ankeret + Laks er taget ind.

---

## Slide-format & canvas

- **Aspect ratio: 16:9.** HTML-canvas = `1280 × 720` px (skaleres til viewport). PPTX = `13.333 × 7.5"`.
- Baggrund: `--bg-canvas` (#f4f4f2). Section-dividers bruger ØS-grøn fuldflade.
- **Sikker margin:** mindst 64px (≈0.66") til alle kanter; 78px i bunden hvor logoet sidder.
- **Intet side-nummer og ingen footer-tekst.** Logoet nederst til højre er den eneste faste markør.

---

## Type-skala til slides

Projektor kræver større tekst end en skærm-UI (px på 1280×720-canvas):

```
slide-display:  96px / 700 / -0.03em   → titel-slide, sektions- og closing-overskrift
slide-title:    56px / 700 / -0.03em   → slide-overskrift (én pr. slide) + grøn accent
slide-h2:       36px / 700 / -0.02em   → underoverskrift inde i en slide
slide-lead:     28px / 400             → manchet/intro-sætning
slide-body:     22px / 400 / 1.4       → brødtekst, bullets
slide-stat:     88px / 700 / mono      → KPI-tal (tabular-nums)
slide-footer:   13px / 600 / 0.1em uppercase MONO   → sub-labels i kolonner, captions
```

Regel: **maks. ~6 bullets eller ~40 ord per slide.** Slides er ikke dokumenter.
Én pointe per slide. Hvis der er to pointer, lav to slides.

---

## Farvebrug (signal-logik, ikke dekoration)

| Rolle | Farve | HEX | Bruges til |
|---|---|---|---|
| Kanter, tekst, frames | Sort | `#000` | alle yderkanter, primær tekst, table-headers |
| Brand-anker | **ØS-grøn** | `#066b43` | titel-accent, section-dividers, logo, titel-shadow, closing-CTA |
| Positiv / vækst | ØS-grøn | `#066b43` | tal der stiger, check-markører |
| Signal | **Laks** | `#ed5e66` | danger, primary-shadow, citat-kant, "fald" i tal, accent-bullet |
| Highlight-flade | Gul | `#fef3c7` | ÉN fremhævet flade pr. slide — aktivt KPI-kort, fremhævet tabelrække |

**Grøn optræder præcis her — ikke andre steder:**
1. **Titel-accent** — kort grøn bjælke tæt under `slide-title`.
2. **Section-dividers** — grøn fuldflade, hvid tekst + hvidt logo. Det stærke brand-moment.
3. **Logo** — officielt ØS-logo i nederste højre hjørne.
4. **Titel-slidets ramme** — grøn offset-shadow (ikke en grøn topkant).
5. **Positive tal.**
6. **Closing-CTA** — grøn flade som brand-sign-off.

Aldrig: farvede gradienter, mere end to accentfarver synlige på samme slide, farvet brødtekst,
grøn og laks blandet i samme element.

---

## Logo

Det officielle ØS-logo (krone + "Økonomistyrelsen") ligger i `assets/` i to transparente versioner:

- `oes-logo-green.png` — på lyse slides.
- `oes-logo-white.png` — på den grønne section-flade (jf. brandguide s.3).

Det placeres automatisk i nederste højre hjørne. **Cover og full-bleed billed-slides får
intet logo** — de går helt til kanten, og der er ingen hjørne-luft.

Vil du bruge en højere-opløst/vektor-version: erstat filerne i `assets/` med samme navne,
eller send `Deck(logo_path="min-logo.png")`. Officielle filer ligger i den koncernfælles
billedbank (Skyfish) jf. brandguiden.

---

## Slide-typer (det faste bibliotek)

HTML-klasse + PPTX-helper i parentes. Alle typer med overskrift får automatisk
titel øverst + grøn accent + logo.

### 1. Title (`.slide--title` / `title()`)
Åbnings-slide: hvid kasse med faste mål (se "Forsidens kasse" ovenfor) og **grøn**
offset-shadow, `slide-display`-titel, evt. manchet, evt. `meta=`-linje under kassen.
Teksten er centreret i kassen. Ingen grøn accent-linje — shadowen bærer grøn.

### 2. Section divider (`.slide--section.is-brand` / `section()`)
Nyt kapitel: **ØS-grøn fuldflade**, hvid tekst, hvidt logo, stort mono-sektionsnummer.
Lodret centreret. Brug sparsomt — det er et "åndedræt".

### 3. Content / bullets (`.slide--content` / `content()`)
Standard-sliden: `slide-title` + grøn accent + bullet-liste med små sorte firkanter
(ikke runde prikker). Maks. 6. Én bullet må have laks-markør (`accent_index`).

### 4. Two-column (`.slide--split` / `split()`)
To kolonner adskilt af 1px lysegrå linje. Hver kolonne har sin egen mono-label øverst
(det er ikke en overskrift — den må godt stå der).

### 5. Statement / quote (`.slide--statement` / `statement()`)
Én stor sætning eller citat med **laks** 6px venstre-kant. Ingen krøllede anførselstegn-grafik.
Kilde i mono uppercase nederst.

### 6. Stat / KPI (`.slide--stats` / `stats()`)
1–3 store tal i bordede kort. Tallet i mono + tabular-nums, label i mono under.
**Maks. ét kort fremhævet** (gul flade + laks shadow). Vækst = grøn, fald = laks.

### 7. Table (`.slide--table` / `table()`)
Sort header med hvide mono-uppercase labels, 1px lysegrå rækkelinjer, numeriske kolonner
højrejusteret + mono + tabular-nums. Fremhævet række = gul.

### 8. Closing (`.slide--closing` / `closing()`)
"Tak" / kontakt / næste skridt. Spejler title-sliden. CTA-boksen er **grøn** med laks shadow.

### 9. Agenda (`.slide--agenda` / `agenda()`)
Nummereret indholdsfortegnelse: mono-nummer + sans-titel pr. række, 1px delelinje.
Den aktuelle sektion må fremhæves (laks nummer).

### 10. Cover / hero med billede (`.slide--cover` / `cover()`)
Fuldt baggrundsbillede + hvid overlay-boks med hård shadow. **Intet logo.**

### 11. Full-bleed billede (`.slide--image-full` / `image_full()`)
Billedet fylder hele sliden; sort caption-stribe nederst med hvid mono-label + titel. **Intet logo.**

### 12. Billede + tekst (`.slide--image-text` / `image_text()`)
Halv-til-halv: billed-felt på den ene side, bullets på den anden.
`is-image-left` / `image_left=True` spejler. Billede får 2px sort kant.

### 13. Billed-grid (`.slide--image-grid` / `image_grid()`)
2×2 eller 3-op (`cols-3`) billed-felter med ens størrelse og 2px kanter.

### 14. Feature-kort (`.slide--cards` / `cards()`)
2–4 bordede kort med ikon-plads (sort firkant), titel + kort tekst.
Maks. ét fremhævet (gul flade + laks ikon).

### 15. Timeline / proces (`.slide--timeline` / `timeline()`)
Vandrette trin med firkant-markører + forbindelseslinje. Det aktuelle trin må fremhæves laks.

### 16. Sammenligning (`.slide--compare` / `comparison()`)
To bordede kolonner med check (**grøn** firkant) / kryds (**laks** firkant) pr. punkt.
"Vinder"-kolonnen må have den ene tilladte laks-shadow.

### 17. Stort tal (`.slide--bignum` / `bignum()`)
Ét kæmpe mono-tal i fuld fokus + kort undertekst. Vækst grøn, fald laks.

### 18. Søjlediagram (`.slide--bars` / `bars()`)
Vandrette søjler: label + sort fyld-bar + mono-værdi. Maks. én bar fremhævet (laks).
Brutalist-erstatning for et farvet diagram.

### 19. Testimonial (`.slide--testimonial` / `testimonial()`)
Portræt-felt (firkant) + stort citat med laks venstre-kant + navn (sans) og rolle (mono).

> **Billeder:** alle billed-typer viser en grå placeholder med mono-label indtil et
> rigtigt billede sættes ind. HTML: byt `.img`-div ud med `<img>`. PPTX: send `image="sti.jpg"`.

---

## Hvad du IKKE skal gøre

- ❌ Aldrig placere overskrift, grøn accent eller logo manuelt — brug helperne (se de fem regler øverst).
- ❌ Aldrig en grøn linje andre steder end tæt under en `slide-title`.
- ❌ Aldrig en kicker-label over overskriften — kun ÉN overskrift pr. slide.
- ❌ Aldrig blurry/soft shadows — kun hårde offsets (sort, grøn eller laks).
- ❌ Aldrig tekstvægge. Maks. ~40 ord. Slides støtter tale, erstatter den ikke.
- ❌ Aldrig mono i normal case eller uden letter-spacing.
- ❌ Aldrig border-radius på flader/kort/dividers (0px). 4px kun på knap-agtige elementer.
- ❌ Aldrig runde bullet-prikker — brug små sorte firkanter.
- ❌ Aldrig mere end én laks-shadow eller én gul flade per slide.
- ❌ Aldrig side-numre eller footer-tekst — logoet er den eneste faste markør.
- ❌ Aldrig proportionale tal i tal-kontekst — altid tabular-nums + mono.

## Når du er i tvivl

Spørg: "Ville denne slide se rolig og bevidst ud fra bagerste række?
Og bærer den stadig systemets karakter (sort kant, hård shadow, mono-label, grøn kun hvor den hører hjemme)?"
Hvis ja til begge — kør på. Bryd kun reglerne hvis brugeren beder om det, og sig hvilken regel du bryder.

## Filer

```
PRESENTATIONS.md   — denne guide (læs altid først)
slides.css         — al slide-styling + ØS-identitet (ét stylesheet, intet tema-valg)
slides.js          — tastaturnavigation + fit-to-viewport-skalering
template.html      — GALLERI: ét eksempel pr. slide-type (reference, ikke en deck at kopiere)
build_pptx.py      — Deck-helpers til .pptx — udvid med nye metoder ved behov
deck.example.py    — GALLERI i PPTX-form -> example.pptx
assets/            — officielle ØS-logoer (grøn + hvid, transparente PNG)
slides-oes.css     — udfaset shim der blot importerer slides.css (gamle decks)
```
