# Økonomistyrelsens slide-system — instruktioner til Claude Code

Dette projekt genererer on-brand slidedeck — HTML-slides og PowerPoint — i
**Økonomistyrelsens** stil: bærbar brutalisme med ØS-grøn som identitets-anker.

**Læs `PRESENTATIONS.md` før du genererer et deck.** Den indeholder de fulde brand-regler:
type-skala, farvebrug, slide-typer, den faste vertikale rytme og do/don'ts.

## ⚠️ Der er ÉN stil

Der er **ikke** længere et "primært" tema og et "ØS"-tema at vælge imellem.
`slides.css`, `template.html` og `Deck()` ER ØS-stilen. Vælg ikke tema, spørg ikke om tema,
og lav ikke en variant "uden ØS-farver".

(`Deck(theme=...)` tages stadig imod, men ignoreres — den findes kun så gamle
deck-scripts ikke går i stykker. `slides-oes.css` importerer blot `slides.css`.)

## ⚠️ De fem fejl der bliver gentaget

Disse fem er kodet ind i CSS'en og i `build_pptx.py`. **Brug helperne — placér dem ikke selv.**

1. **Overskriften sidder ØVERST** på sliden (PPTX: 0,50" / 1,27 cm), ikke skubbet ned midt på.
2. **Den grønne linje hører til overskriften** og sidder tæt under den (PPTX: 1,142" / 2,90 cm).
   Den må aldrig svæve mellem overskrift og indhold, og den må kun findes under en `slide-title`.
3. **Den grønne linje flugter med overskriftens venstrekant.**
4. **Logoet sidder i højre HJØRNE** med ens luft til højre og bund — ikke inde på tekstmargin.
5. **Kun ÉN overskrift pr. slide.** Ingen kicker-eyebrow over titlen.
   `kicker=`-argumenter i `Deck` ignoreres — undtagen i `image_full`, hvor det
   er mono-labelen i caption-striben (brug `label=` der). Sub-labels i kolonner
   og billed-captions er ikke overskrifter og bliver.

I HTML kommer 1–4 af sig selv når du bruger `<h2 class="slide-title">` og `.slide`-klasserne.
I PPTX kommer de af sig selv når du bruger `Deck`-metoderne. Tegn dem ikke manuelt.

## ⚠️ Byg ud fra INDHOLDET, ikke ud fra galleriet

`template.html` og `deck.example.py` er **demo-gallerier / komponentbiblioteker** — de viser
ÉT eksempel på *hver* tilgængelig slide-type. **De er IKKE en færdig deck der skal kopieres.**

Når brugeren beder om "et deck om X":

- ❌ **Lav IKKE** en kopi af galleriet med samme slides, samme antal og samme rækkefølge.
- ❌ **Brug ikke** alle slide-typer bare fordi de findes.
- ✅ **Start fra indholdet/formålet.** Find budskabet og en fortælling, og *vælg derudfra*:
  - **Hvilke** slide-typer der giver mening (drop resten).
  - **Hvor mange** slides emnet kræver (kan være 4, kan være 30 — galleriets 21 er tilfældigt).
  - **Hvilken rækkefølge** der bedst bærer fortællingen.
- ✅ **Opfind nye slides/layouts** når indholdet kræver noget biblioteket ikke har — så længe du
  holder dig til DNA'et (sorte kanter, hårde offset-shadows, IBM Plex, mono = system /
  sans = indhold, grøn kun hvor den hører hjemme, 0px på flader / 4px på interaktivt,
  tokens fra `slides.css`). Genbrug eksisterende klasser/helpers hvor du kan.

Kort sagt: **galleriet er paletten, ikke maleriet.** Stilen er fast — indholdet bestemmer
hvilke og hvor mange slides.

## Workflow når brugeren beder om et deck

1. Læs `PRESENTATIONS.md` + de to ⚠️-afsnit ovenfor.
2. Afklar kun det der mangler: emne, publikum, HTML eller PPTX. Gæt fornuftigt.
   (Spørg IKKE om antal slides — det udleder du af indholdet. Spørg IKKE om tema — der er kun én.)
3. **Lav en outline FØRST:** skriv fortællingen som en liste af slides — for hver: hvilken
   slide-type (eller ny idé) + det konkrete indhold. Det er her arbejdet ligger.
4. **HTML-deck:** opret en ny `.html`-fil. Lås `<link>` til `slides.css` + `slides.js`.
   Byg DINE slides fra outlinen med slide-type-klasserne (slå dem op i `template.html`
   som reference) — ikke ved at kopiere hele filen. Tilføj nye `.slide--*`-klasser i en
   `<style>` hvis indholdet kræver et nyt layout.
5. **PowerPoint-deck:** skriv et nyt `deck-<emne>.py` der bruger `Deck`-helperne i
   `build_pptx.py` og kalder netop de slide-metoder din outline kræver, i din rækkefølge.
   Mangler der en slide-type, tilføj en ny metode i `build_pptx.py` i samme stil —
   og brug `self._title()` / `self._footer()` så rytmen holder. Kør `python deck-<emne>.py`.
   Kræver `pip install python-pptx`.
6. Brugeren skal **aldrig** gentage designreglerne — de er kodet ind i guiden, tokens og helpers.

## Kilde på tal

`stats`, `table`, `bars` og `bignum` tager `source="Kilde: …"` — en mono-linje
nederst til venstre. I HTML: `<p class="slide__source">`. **Sæt den når tallene
ikke er brugerens egne.** Har du kun tal uden kilde, så spørg efter kilden i
stedet for at udgive slides uden.

## Fonte

IBM Plex **indlejres** i hver genereret .pptx (`assets/fonts/`), så typografien
holder også på maskiner uden fonten installeret. Rør ikke `embed_fonts=`, og slet
ikke `assets/fonts/LICENSE.txt` — OFL kræver at licensen følger fonten.

Dukker font-advarslen fra `save()` op, er indlejringen slået fra eller filerne
væk. Videregiv den til brugeren — decket ser stadig "rigtigt" ud, så den er let
at overse.

Bemærk to forskellige begreber: `statement(attribution=…)` er en citat-afsender
lige under citatet; `source=` på tal-slides er en datakilde nederst på sliden.

## Kør verify.py efter enhver ændring i build_pptx.py eller slides.css

```bash
python verify.py
```

Der er ingen PowerPoint- eller browser-rendering her, så de faste mål kan skride
uden at nogen ser det. `verify.py` bygger galleriet og tjekker referencemålene,
geometrien i den genererede fil, font-indlejringen, HTML/PPTX-pariteten,
galleri-dækningen og at der ikke er efterladte kicker-labels. Exit-kode 1 ved fejl.

Tilføjer du en ny slide-type i `build_pptx.py`, skal den også kaldes i
`deck.example.py` — ellers fejler galleri-dækningen. Ændrer du bevidst et fast
mål, skal det ændres BÅDE i `build_pptx.py` og i `REFERENCE_CM` i `verify.py`.

## Vertikal rytme — ret den centralt, ikke pr. slide

Målene ligger som konstanter i toppen af `build_pptx.py` (`TITLE_TOP`, `ACCENT_Y`, `ACCENT_X`,
`CONTENT_TOP`, `CONTENT_BOT`, `LOGO_W`, `LOGO_INSET`) og som CSS-variabler i `:root` i
`slides.css` (`--slide-pad`, `--slide-pad-bottom`). Skal noget flyttes, flyt det dér —
så flytter det sig ens i alle slide-typer og i begge formater.

## Filer

```
CLAUDE.md          — denne fil (entry point for Claude Code)
PRESENTATIONS.md   — fuld brand-guide (læs altid)
slides.css         — al slide-styling + ØS-identitet (ét stylesheet, intet tema-valg)
slides.js          — tastaturnavigation + fit-to-viewport (uændret)
template.html      — GALLERI/komponentbibliotek: ét eksempel pr. slide-type
build_pptx.py      — Deck-helpers til .pptx — udvid med nye metoder ved behov
deck.example.py    — GALLERI i PPTX-form -> example.pptx
example.pptx       — genereret galleri
assets/            — officielle ØS-logoer (grøn + hvid, transparente PNG)
slides-oes.css     — udfaset shim: importerer blot slides.css (så gamle decks stadig virker)
```

## Relation til UI design systemet

Præsentationsguiden følger samme æstetik som `~/Desktop/Mortens Design system/morten-ds/` —
samme fonts (IBM Plex), shadows og signal-logik — men skaleret til 16:9-slides og
projektor-afstand, og med ØS-grøn/Laks i stedet for den generiske rød/grøn.
De to systemer er bevidst holdt adskilt.
