# Økonomistyrelsen-tema — slide-variant

> **= Primær præsentation stil + ØS-grøn + logo.** Tilpasset **Økonomistyrelsen (ØS)**.
> Den holder sig bevidst TÆT på den primære stil — grøn bruges sparsomt som identitets-anker,
> ikke som gennemgående farve. **Læs `PRESENTATIONS.md` først**; denne fil beskriver kun forskellene.

## Princip: primær stil + ØS-grøn + officielt logo

ØS-temaet ER den primære stil. Vi ændrer kun farver på bevidste steder og lægger
ØS' officielle logo på. Alt andet (sorte kanter, hårde shadows, IBM Plex, gul highlight,
layouts) er **identisk med den primære stil**.

| Rolle | Primær stil | ØS-tema | HEX |
|---|---|---|---|
| Kicker-eyebrow over titlen | muted grå label | **fjernet** (kun den sorte titel) | — |
| Signal (danger/primary/vigtigt) | rød `#ff5757` | **Laks** (ØS' koral) | `#ed5e66` |
| Positiv / vækst | grøn `#16a34a` | ØS-grøn | `#066b43` |
| Highlight-flade | gul `#fef3c7` | **uændret** (gul) | `#fef3c7` |
| Kanter, tekst, frames | sort | **uændret** (sort) | `#000` |
| Brand-anker (ny) | — | ØS-grøn | `#066b43` |

**Grøn optræder her** (bevidst, ikke garish):
1. **Titel-accent** — kort grøn bjælke under content-titler, flush med titlens venstrekant.
2. **Section-dividers** — ØS-grøn flade, hvid tekst. Det stærke brand-moment mellem kapitler.
3. **Logo** — officielt ØS-logo nederst til højre på alle slides med plads (se nedenfor).
4. **Titel-frame** — **grøn offset-shadow** (i stedet for grøn topkant).
5. **Positive tal** (vækst/op) — ØS-grøn.
6. **Closing-CTA** — ØS-grøn flade som brand-sign-off.

**Hierarki-regel:** kun ÉN overskrift pr. slide = den sorte titel. Kicker-eyebrow'en over titlen
er **fjernet** i ØS-temaet (sub-labels i kolonner og billed-caption beholdes — de er ikke overskrifter).
Signal-shadow er Laks, ikke grøn, så de to ikke blandes.

## Logo (officielt, indlejret)

Det **officielle ØS-logo** (krone + "Økonomistyrelsen", udtrukket fra brandguiden) ligger i
`assets/` i to transparente versioner og placeres **nederst til højre**:
- `oes-logo-green.png` — grøn logo på lyse slides.
- `oes-logo-white.png` — hvid logo på den grønne section-flade (jf. brandguide s.3: hvid logo på ØS-grøn).
- Cover og full-bleed billed-slides har ingen footer → intet logo (ingen plads).

Laks bruges præcis hvor rød blev brugt: danger, primary-shadow, citat-kant, "fald" i tal, accent-bullet.

## Sådan bruger jeg det

- **HTML:** kopiér `template-oes.html` (linker `slides-oes.css`, som selv importerer `slides.css`). Logoet ligger i footeren via CSS — intet at gøre.
- **PowerPoint:** `Deck(theme="oes")`. Logoet indlejres automatisk nederst til højre. Alt indhold-API er identisk med primær — se `deck.oes.example.py`.

## Logo-asset

De officielle logoer ligger i `assets/` (udtrukket fra brandguiden, transparente PNG'er) og bruges automatisk.
Vil du bruge en højere-opløst/vektor-version: erstat filerne i `assets/` med samme navne,
eller send i PPTX `Deck(theme="oes", logo_path="min-logo.png")` for at override begge.
Officielle filer kan hentes i den koncernfælles billedbank (Skyfish) jf. brandguiden.

## Hvad der bevidst IKKE blev overtaget fra ØS' brandguide

- ❌ ØS' fonte — vi beholder IBM Plex (stærkere, og ØS-fonten lå ikke i guiden).
- ❌ ØS' Templafy-slide-layouts — vores brutalist-struktur er mere distinkt.
- ❌ Den fulde pastel-palette (himmelblå, latte, gråblå osv.) — ville udvande brutalismen. Kun grøn-ankeret + Laks er taget ind.
