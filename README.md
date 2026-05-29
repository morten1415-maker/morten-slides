# Mortens Præsentationer

Slide-system bygget på Mortens design system ("bærbar brutalisme") — til **HTML-slides** og **PowerPoint**.
To stilarter, hver med sit eget galleri af alle slide-typer.

## Sider (GitHub Pages)

- **Primær præsentation stil** → [`template.html`](template.html)
- **Økonomistyrelsen** → [`template-oes.html`](template-oes.html)

(Når Pages er slået til ligger de på `https://<bruger>.github.io/morten-slides/`.)

## Indhold

| Fil | Hvad |
|---|---|
| `index.html` | Forside med link til begge decks |
| `template.html` / `template-oes.html` | Demo-galleri pr. stil (alle slide-typer) |
| `slides.css` / `slides-oes.css` | Styling (ØS importerer den primære) |
| `slides.js` | Tastaturnavigation + fit-to-viewport |
| `build_pptx.py` | PowerPoint-generator (`Deck(theme="default"|"oes")`) |
| `deck.example.py` / `deck.oes.example.py` | Eksempel-decks → `.pptx` |
| `PRESENTATIONS.md` / `PRESENTATIONS-OES.md` | Brandguides (læses før et deck genereres) |
| `assets/` | Officielle ØS-logoer (grøn + hvid, transparente) |

## Brug

- **HTML:** åbn `template*.html` i browseren. Pil/mellemrum/klik navigerer. Cmd/Ctrl+P → print til PDF.
- **PowerPoint:** `pip install python-pptx`, så `python3 deck.example.py` (eller `deck.oes.example.py`).
