# Økonomistyrelsens slide-system

Slide-system til **HTML-slides** og **PowerPoint** — bærbar brutalisme med ØS-grøn
som identitets-anker. **Én stil**, ét galleri af alle slide-typer.

## Galleri (GitHub Pages)

- **Alle slide-typer** → [`template.html`](template.html)

(Når Pages er slået til ligger det på `https://<bruger>.github.io/morten-slides/`.)

## Indhold

| Fil | Hvad |
|---|---|
| `template.html` | Galleri: ét eksempel pr. slide-type (reference — ikke en deck at kopiere) |
| `slides.css` | Al styling + ØS-identitet. Ét stylesheet, intet tema-valg |
| `slides.js` | Tastaturnavigation + fit-to-viewport |
| `build_pptx.py` | PowerPoint-generator (`Deck()`) |
| `deck.example.py` | Galleriet i PPTX-form → `example.pptx` |
| `PRESENTATIONS.md` | Brandguide (læses før et deck genereres) |
| `CLAUDE.md` | Entry point for Claude Code |
| `assets/` | Officielle ØS-logoer (grøn + hvid, transparente) |
| `slides-oes.css` | Udfaset shim der importerer `slides.css` — så ældre decks stadig virker |

## Før første brug: installér IBM Plex

PowerPoint refererer fonte ved navn — mangler IBM Plex på maskinen, erstattes den
lydløst med Calibri, og decket mister mono/sans-forskellen der bærer systemet.
`build_pptx.py` advarer hvis den mangler.

Hent IBM Plex Sans + Mono (gratis, OFL) fra
[github.com/IBM/plex/releases](https://github.com/IBM/plex/releases), marker
`.ttf`-filerne → højreklik → *Installer for mig* (kræver ikke admin).

## Brug

- **HTML:** åbn `template.html` i browseren. Pil/mellemrum/klik navigerer. Cmd/Ctrl+P → print til PDF.
- **PowerPoint:** `pip install python-pptx`, så `python deck.example.py`.

## De faste mål

Overskrift, den grønne accent-linje og logoet placeres automatisk og må ikke sættes pr. slide:

| | PPTX (13.333 × 7.5") | HTML (1280 × 720 px) |
|---|---|---|
| Overskrift, top | 0.50" | 64px |
| Grøn accent, top | 1.142" | tæt under titlen |
| Indhold | 1.70" → 6.60" | 64px → 642px |
| Logo, hjørne-luft | 0.28" | 27px |
| Forsidens kasse | 11.0" × 4.0" @ 1.70" | 1056 × 384px @ 163px |
| Kilde-linje | 6.85" | 657px |

Konstanterne står i toppen af `build_pptx.py` og i `:root` i `slides.css`. Ret dem der.
