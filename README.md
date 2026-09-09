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
| `assets/fonts/` | IBM Plex Sans + Mono (TTF) der indlejres i .pptx, med OFL-licens |
| `verify.py` | Regressionstjek af alle faste mål, font-indlejring og galleri-dækning |
| `slides-oes.css` | Udfaset shim der importerer `slides.css` — så ældre decks stadig virker |

## Fonte følger med filen

PowerPoint refererer fonte ved navn — mangler IBM Plex på maskinen der åbner filen,
erstattes den lydløst med Calibri, og decket mister mono/sans-forskellen der bærer
systemet. Derfor **indlejres IBM Plex i hver genereret .pptx**: typografien holder
på en låst VDI, hos modtagere og på mødelokale-PC'en.

Fontfilerne ligger i `assets/fonts/` (fra [github.com/IBM/plex](https://github.com/IBM/plex),
v6.4.0). IBM Plex er OFL-licenseret; `assets/fonts/LICENSE.txt` skal følge med.
Det koster ca. 0,5 MB pr. fil — slå fra med `Deck(embed_fonts=False)`.

## Brug

- **HTML:** åbn `template.html` i browseren. Pil/mellemrum/klik navigerer. Cmd/Ctrl+P → print til PDF.
- **PowerPoint:** `pip install python-pptx`, så `python deck.example.py`.

## Tjek at intet er skredet

```bash
python verify.py
```

Bygger galleriet og kontrollerer de invarianter der før blev rettet i hånden deck
efter deck: at konstanterne stadig er referencedeckenes mål, at den genererede
.pptx følger dem, at IBM Plex faktisk ligger indlejret i filen, at `slides.css`
og `build_pptx.py` taler om samme mål, at galleriet viser hver slide-type, og at
der ikke er efterladt kicker-labels over overskrifter. Exit-kode 1 ved fejl.

Kør det efter enhver ændring i `build_pptx.py` eller `slides.css` — der er ingen
PowerPoint- eller browser-rendering i pipelinen, så målene kan ellers skride
uden at nogen ser det.

## De faste mål

Overskrift, den grønne accent-linje og logoet placeres automatisk og må ikke sættes pr. slide:

| | PPTX (13.333 × 7.5") | HTML (1280 × 720 px) |
|---|---|---|
| Overskrift, top | 0.50" | 64px |
| Grøn accent, top | 3,26 cm | 5px under titlens linjeboks |
| Indhold | 1.70" → 6.60" | 64px → 642px |
| Logo, hjørne-luft | 0.28" | 27px |
| Forsidens kasse | 11.0" × 4.0" @ 1.70" | 1056 × 384px @ 163px |
| Kilde-linje | 6.85" | 657px |

Konstanterne står i toppen af `build_pptx.py` og i `:root` i `slides.css`. Ret dem der.
