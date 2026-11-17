# Schaalmodellen – lokale inventaris

Een lokale web-app om een collectie schaalmodellen te beheren: bekijken, zoeken, filteren, bewerken, afspraken met kopers plannen en zoekertjes klaarzetten. Alles draait op je eigen pc, zonder internet, externe scripts, lettertypen of diensten.

De bron is een Excel-bestand (`Inventaris.xlsx`, werkblad **Inventaris**). Dat bestand wordt alleen gelezen. Wat je via de pagina toevoegt of wijzigt, wordt apart bewaard.

---

## Inhoud

- [Snel starten](#snel-starten)
- [Waar staat wat](#waar-staat-wat)
- [Installatie](#installatie)
- [Functies](#functies)
- [Modellen toevoegen, bewerken en verwijderen](#modellen-toevoegen-bewerken-en-verwijderen)
- [Afspraken met kopers en kalender](#afspraken-met-kopers-en-kalender)
- [Foto's en zoekertjes](#fotos-en-zoekertjes)
- [Excel en pagina naast elkaar](#excel-en-pagina-naast-elkaar)
- [Parquet-export](#parquet-export)
- [Weergaveregels](#weergaveregels)
- [Projectstructuur](#projectstructuur)
- [Tests](#tests)

---

## Snel starten

| Wat wil je doen? | Hoe |
|---|---|
| **Alleen bekijken** | Open `data\index.html` in je browser |
| **Toevoegen of bewerken** | Dubbelklik op `Start inventaris.cmd` |
| **Excel opnieuw inlezen** | Sla Excel op en sluit het, dubbelklik op `Bijwerken.cmd`, vernieuw de pagina |

De app opent automatisch in je browser. Laat het appvenster open zolang je werkt en stop met <kbd>Ctrl</kbd>+<kbd>C</kbd>.

De app gebruikt standaard poort **18765**. Is die bezet, dan kiest ze zelf een vrije poort en opent ze het juiste adres.

```bash
python server.py [--port 18765] [--no-browser]
```

---

## Waar staat wat

**Code** staat in `C:\dev\schaalmodellen` (Git, GitHub `geert1981/schaalmodellen`).

**Gegevens** staan in OneDrive, zodat er automatisch een back-up van is:

```
C:\Users\ntm556\OneDrive - De Stroomlijn\Geert Kril\Schaalmodellen\data
├── index.html             # gegenereerde pagina
├── aanvullingen.json      # foto- en zoekertjeskoppelingen uit chats
├── lokale_modellen.json   # paginamodellen, aanpassingen, prullenbak, afspraken
├── inventaris.parquet     # export van het Excel-werkblad
└── fotos/
    ├── web/               # foto's die via de pagina zijn toegevoegd
    └── _mini/             # miniaturen voor het foto-overzicht
```

`C:\dev\schaalmodellen\data` is een **junction** naar die OneDrive-map. Het pad staat ook in `.data-root`.

Ontbreekt de koppeling, bijvoorbeeld op een andere laptop? Maak ze dan opnieuw aan in een opdrachtprompt:

```cmd
mklink /J "C:\dev\schaalmodellen\data" "C:\Users\ntm556\OneDrive - De Stroomlijn\Geert Kril\Schaalmodellen\data"
```

> [!IMPORTANT]
> Stel in OneDrive in dat deze map **altijd op dit apparaat** wordt bewaard.

**Bron:** `C:\Users\ntm556\OneDrive - De Stroomlijn\Geert Kril\Schaalmodellen\Inventaris.xlsx`

---

## Installatie

Je hebt Python 3 nodig, met:

- **pyarrow**: voor de Parquet-export
- **Pillow**: voor de miniaturen

Eenmalig installeren:

```bash
python -m pip install -r requirements.txt
```

De `.cmd`-bestanden zoeken zelf een geschikte Python. Ze nemen bij voorkeur de installatie waarin pyarrow al beschikbaar is: de Python-launcher (`py`), `python`, of de Python die met Codex is meegeleverd.

---

## Functies

### Weergaven
- **Foto-overzicht** (start): kaarten met foto, locatie, status en de vraagprijs zonder doos.
- **Tabel**: alle 15 kolommen. Sorteren kan via de kolomtitels.
- **Per locatie**: modellen gegroepeerd per locatienummer, met aantal en waarde te koop per locatie. Met **Afdrukken** krijg je een lijst per locatie.
- **Kalender**: maandoverzicht en een lijst met komende afspraken.

Zoeken, filters en sortering werken in alle weergaven, behalve in de kalender.

### Zoeken en filteren
- Zoeken gaat door alle 15 velden en de opmerkingen. Als je meerdere woorden typt, moeten ze allemaal voorkomen.
- Onder **Filters** staan extra filters, die je kunt combineren. Het aantal actieve filters staat bij de knop. Je kunt onder meer filteren op:
  - **prijsgrens** (vraagprijs zonder doos), bijvoorbeeld van 20 tot en met 50 euro. Modellen zonder prijs vallen dan weg.
  - modellen **met foto's en/of zoekertjes**.
- **Wis filters** zet zoeken, filters en sortering terug.
- Zoekterm, filters, sortering en weergave worden in de browser onthouden.

### Totalen
Bovenaan staat het aantal modellen dat online, gereserveerd of verkocht is. Houd de muis op een getal voor de waarde (vraagprijs zonder doos) en hoeveel modellen geen prijs hebben.

### Export
**Exporteer CSV** downloadt de modellen die je nu ziet (na zoeken, filters en sortering), met alle 15 kolommen plus Opmerking en ChatGPT URL. Het bestand opent meteen goed in Excel: puntkomma als scheidingsteken, komma als decimaalteken.

### Overig
- **Status zoekertje** kun je in *Bekijk model* meteen wijzigen via de keuzelijst. Dat wordt direct opgeslagen (alleen als de app via `Start inventaris.cmd` draait).
- **ChatGPT URL**: elk model heeft dit veld, aan te passen via *Model bewerken*. Het staat als klikbare link bij de details. Modellen met een gekoppelde bronchat krijgen die link als beginwaarde.
- **Licht/donker thema** wissel je met de knop rechtsboven. Zonder eigen keuze volgt de pagina de Windows-instelling.
- Een kolom die bij geen enkel model is ingevuld (nu **Doos**) wordt in tabel, filters en details verborgen. CSV en formulier bevatten ze wel.
- **Miniaturen** (`fotos/_mini`) worden automatisch gemaakt bij het starten, na opslaan en bij `Bijwerken.cmd`. Ontbreekt een miniatuur, dan toont de kaart de originele foto. Alles handmatig aanmaken kan met `python miniaturen.py`.

---

## Modellen toevoegen, bewerken en verwijderen

### Toevoegen
1. Start de app met `Start inventaris.cmd`.
2. Klik op **+ Model toevoegen**.
3. Vul minstens **automerk** en **model** in. De rest is optioneel. *Conditie* en *Status zoekertje* kies je uit een keuzelijst.
4. Kies foto's en plak de volledige tekst voor het zoekertje.
5. Zet de hoofdfoto vooraan met de pijltjes en klik **Model opslaan**.

Een nieuw model krijgt automatisch een nummer dat begint met `W-`.

### Bewerken
Open **Bekijk model**, dan **Model bewerken**. Je kunt gegevens, opmerkingen, foto's en de zoekertjestekst wijzigen, ook bij modellen uit Excel. Het modelnummer blijft hetzelfde.

### Verwijderen en herstellen
**Model verwijderen** verplaatst het model naar de **Prullenbak**. Daar kun je het herstellen, met foto's en tekst.

- Verwijderde modellen blijven ook na een Excel-update weg uit het overzicht.
- Foto's worden niet van de schijf gewist en Excel wordt niet gewijzigd.
- Haal je een foto weg in het formulier, dan verdwijnt alleen de koppeling. Het bestand zelf blijft bewaard.

---

## Afspraken met kopers en kalender

In **Bekijk model › Afspraak met koper** plan je een bezoek: datum en tijd, duur, koper (naam en telefoon) en een notitie. Per model is er één afspraak. Sla je opnieuw op, dan vervangt dat de vorige. Afspraken staan in `lokale_modellen.json`.

- Komende afspraken staan ook op de kaart van het model.
- **Exporteer .ics** maakt een agendabestand voor één afspraak (Outlook, Google, Apple), met een herinnering een uur vooraf.
- **Alle afspraken exporteren** maakt één `.ics`-bestand met alles.
- Importeer je een gewijzigde afspraak opnieuw, dan wordt de bestaande bijgewerkt (zelfde UID).
- **Toevoegen aan Google Agenda** opent Google Agenda met de afspraak al ingevuld. Hiervoor is internet nodig.

> [!NOTE]
> Een automatische koppeling met Google Agenda kan niet: de app draait alleen lokaal en Google kan er niet bij.

---

## Foto's en zoekertjes

Klik op een foto of op **Bekijk model**. Het detailvenster toont:

- de lokaal opgeslagen foto's. Een klik opent het originele bestand.
- de tekst voor het zoekertje, met een **kopieerknop**. Blokkeert de browser het kopiëren, dan wordt de tekst geselecteerd en druk je zelf op <kbd>Ctrl</kbd>+<kbd>C</kbd>.
- een link naar de bronchat (hiervoor is internet nodig).

Over de zoekertjesteksten:

- Ze komen uit bestaande chats. Opmaak en de vaste afsluiting zijn gelijkgetrokken, en fabrikantvoorvoegsels zijn uit de titel gehaald.
- Chats die niet eenduidig aan een model gekoppeld zijn, staan apart onder de tabel. Hun foto's worden pas bij een model gezet als de koppeling zeker is.
- Nieuwe chats worden niet automatisch opgehaald. Daarvoor is een nieuwe verwerking nodig.
- Verandert merk, model, bouwjaar, fabrikant of artikelnummer in Excel, dan wordt de oude koppeling uit voorzorg niet meer toegepast tot iemand ze opnieuw controleert.
- Bij sommige oudere chats waren de oorspronkelijke foto's niet toegankelijk.

> [!TIP]
> Controleer vóór je een zoekertje publiceert de vraagprijs, conditie en andere gegevens tegen de actuele inventaris.

---

## Excel en pagina naast elkaar

Je kunt modellen uit Excel zowel in Excel als via de pagina bijwerken.

- **Per veld gaan aanpassingen via de pagina voor.** De andere velden worden bij het opnieuw inlezen uit Excel bijgewerkt.
- De pagina schrijft **niets** terug naar Excel. Paginamodellen, aanpassingen en de prullenbak staan in `lokale_modellen.json`. Nieuwe foto's komen in `data/fotos/web`, in de gekozen volgorde.
- Bij het starten gebruikt de app de laatst ingelezen Excel-gegevens. Excel wordt pas opnieuw gelezen als je daarvoor kiest: via **Excel opnieuw inlezen** in de app of via `Bijwerken.cmd`. Beide manieren behouden de paginamodellen.
- Ook als Excel vergrendeld is, kun je paginamodellen opslaan.
- Gaat er iets mis bij het inlezen, dan blijft de vorige pagina behouden.

### Na wijzigingen in Excel
1. Sla `Inventaris.xlsx` op en sluit Excel.
2. Klik **Excel opnieuw inlezen** in de app, of dubbelklik op `Bijwerken.cmd`.
3. Vernieuw `data\index.html` in de browser.

> [!WARNING]
> Dubbele modellen tussen Excel en de pagina worden niet automatisch herkend. Voer elk model dus via één van beide in.
>
> Verandert de identiteit van een Excel-model, bijvoorbeeld omdat een nummer opnieuw is gebruikt? Dan worden de opgeslagen wijzigingen, foto's, teksten en verwijdering bewaard maar **niet toegepast**, en toont de pagina een melding. Zo wordt een nieuw model niet per ongeluk aangepast of verborgen.

---

## Parquet-export

`Bijwerken.cmd` maakt ook `data/inventaris.parquet`. Dit bestand:

- bevat alle oorspronkelijke kolommen van het werkblad **Inventaris**, ook Opmerking, en alle rijen met waarden, ook rijen die al vooraf ingevuld waren.
- laat lege cellen leeg (`null`). Ze worden niet vervangen door *Onbekend*.
- houdt numerieke kolommen numeriek. Kolommen met zowel getallen als tekst worden tekst, zodat artikelnummers en afwijkende waarden behouden blijven.
- bewaart datums en tijden als hun onderliggende getal of als ISO-tekst, zonder opmaak. Formules leveren hun laatst opgeslagen resultaat.
- bevat **geen** foto's, chatteksten of paginamodellen. Het blijft een weergave van de Excel-bron.
- heeft het bronpad, de SHA256 en het tijdstip in de metadata.

---

## Weergaveregels

- De 15 inventariskolommen staan altijd in deze volgorde:
  `nr` · `automerk` · `model` · `bouwjaar` · `merk modelauto` · `modelnr` · `conditie` · `schaal` · `oplage` · `locatie` · `doos` · `vraagprijs zonder doos` · `vraagprijs met doos` · `vraagprijs met doos als nieuw` · `2e hands`
- De extra Excel-kolom **Opmerking** staat onder Conditie en is doorzoekbaar.
- Lege velden worden getoond als *Onbekend*. Prijzen staan als bedragen in euro.
- Een schaal die Excel als tijd bewaart (`01:18`) wordt getoond als `1:18`.
- Rijen met alleen een nummer en de standaardschaal worden overgeslagen. Gedeeltelijke registraties, bijvoorbeeld met ook een locatie, blijven zichtbaar.
- Benamingen en eventuele tikfouten uit Excel blijven ongewijzigd.

---

## Projectstructuur

| Bestand | Doel |
|---|---|
| `Start inventaris.cmd` | Start de lokale app (`server.py`) |
| `Bijwerken.cmd` | Leest Excel opnieuw in en maakt `index.html` en `inventaris.parquet` (`bijwerken.py`) |
| `server.py` | Lokale webserver: opslaan, bewerken, prullenbak, afspraken |
| `bijwerken.py` | Leest Excel in, maakt de pagina en de Parquet-export |
| `miniaturen.py` | Maakt miniaturen in `fotos/_mini` |
| `inventaris.template.html` | Sjabloon voor de pagina (HTML, CSS en JavaScript, allemaal ingebed) |
| `test_inventaris.py` | Integratietests |
| `requirements.txt` | Python-afhankelijkheden |
| `data/` | Junction naar de gegevens in OneDrive (niet in Git) |

> [!NOTE]
> Houd `bijwerken.py`, `server.py`, `miniaturen.py`, `inventaris.template.html` en de `.cmd`-bestanden samen in één map.

---

## Tests

De integratietests draaien in tijdelijke mappen. Je echte collectie wordt dus niet gewijzigd.

```bash
python -m unittest test_inventaris -v
```
