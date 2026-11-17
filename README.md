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
- **Foto-overzicht** (start): kaarten met foto, locatie, status en een instelbare vraagprijs: zonder doos, met doos of met doos als nieuw. De keuze wordt onthouden.
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
Bovenaan staat het aantal modellen dat online, gereserveerd of verkocht is. Online en gereserveerd tonen in de toelichting de vraagwaarde zonder doos. Verkocht toont de werkelijke verkoopopbrengst en hoeveel verkoopprijzen nog onbekend zijn.

### Export
**Exporteer CSV** downloadt de modellen die je nu ziet (na zoeken, filters en sortering), met alle 15 kolommen plus Opmerking en ChatGPT URL. Het bestand opent meteen goed in Excel: puntkomma als scheidingsteken, komma als decimaalteken.

**Volledige export** werkt los van zoeken en filters en biedt twee formaten:

- **ZIP**: alle actieve modellen in `collectie.json` en `collectie.parquet`, de originele foto's, opgeslagen aanvullingen, afspraken, verkoopgegevens en prullenbak. Ook `index.html` zit erbij; die kun je uitgepakt offline bekijken. `export-info.json` vermeldt eventueel ontbrekende foto's.
- **Parquet**: de gecombineerde actieve collectie, inclusief lokale wijzigingen, fotopaden, teksten, afspraken en werkelijke verkoopprijs/datum. Foto's zelf worden via de ZIP meegenomen.

Deze volledige export vervangt `inventaris.parquet` niet: dat bestand blijft uitsluitend de Excel-bron weergeven.

### Back-ups en herstel

Vóór opslaan, verwijderen, herstellen uit de prullenbak, afspraken wijzigen en Excel opnieuw inlezen wordt automatisch een herstelpunt gemaakt. Via **Back-ups** kun je ook handmatig een herstelpunt maken en de hele collectie terugzetten. Vóór terugzetten wordt de huidige versie eveneens bewaard. De bron-Excel wordt nooit teruggezet of gewijzigd.

Herstelpunten staan in `data/backups/*.zip`. Foto's worden per inhoudsversie één keer opgeslagen in `data/backups/fotobestanden/`; alle herstelpunten verwijzen daarnaar. Bewaar dus de **hele backups-map** samen, niet alleen een losse ZIP. Bestaande herstelpunten worden niet automatisch verwijderd. De eerste keer kan het veiligstellen van de foto's langer duren. Ontbrekende foto's worden gemeld, niet verzonnen.

### Publicatiecontrole en dubbele modellen

Onder **Filters > Publicatiecontrole** kun je ontbrekende foto's, tekst of prijzen, afwijkende prijzen in zoekertjestekst en mogelijke dubbele modellen tonen. De details tonen de aandachtspunten. Voor prijzen vergelijkt de controle expliciete eurobedragen/vraagprijzen met de ingevulde vraagprijzen; dit is een hulpmiddel, geen inhoudelijke goedkeuring voor publicatie.

Bij toevoegen en bewerken waarschuwt de app voor hetzelfde automerk/model. Bekende verschillende bouwjaren, fabrikanten, artikelnummers of schalen worden onderscheiden. Onbekende gegevens kunnen tot een waarschuwing leiden. Een tweede exemplaar kan nog steeds worden opgeslagen na expliciete bevestiging; de app voegt niets automatisch samen.

### Verkoop registreren

De snelle statuskeuze werkt op de kaarten en in het detailvenster. Bij **Verkocht** vul je de werkelijke verkoopprijs (optioneel) en datum in. Ook via **Model bewerken** kun je die aanpassen. Een onbekende verkoopprijs telt niet als nul in de opbrengst; een expliciete prijs van nul blijft wel nul. Verkoopgegevens blijven na een Excel-update bewaard.

### Overig
- **Status zoekertje** kun je in *Bekijk model* meteen wijzigen via de keuzelijst. Dat wordt direct opgeslagen (alleen als de app via `Start inventaris.cmd` draait).
- **ChatGPT URL**: elk model heeft dit veld, aan te passen via *Model bewerken*. Het staat als klikbare link bij de details. Modellen met een gekoppelde bronchat krijgen die link als beginwaarde.
- **Licht/donker thema** wissel je met de knop rechtsboven. Zonder eigen keuze volgt de pagina de Windows-instelling.
- Een kolom die bij geen enkel model is ingevuld (nu **Doos**) wordt in tabel, filters en details verborgen. CSV en formulier bevatten ze wel.
- **Miniaturen** (`fotos/_mini`) worden automatisch gemaakt bij het starten, na opslaan en bij `Bijwerken.cmd`. Ontbreekt een miniatuur, dan toont de kaart de originele foto. Alles handmatig aanmaken kan met `python miniaturen.py`.

---


## Verkoopbundels

Via **Verkoopbundels > Nieuwe bundel** kies je minstens twee modellen, de koper, contactgegevens, een afgesproken totaalprijs en eventueel één afhaalafspraak. Ook vanuit een model kun je **Bundel maken** kiezen.

- Een model kan tegelijk bij één actieve bundel horen. Afzonderlijk verkochte modellen kunnen niet opnieuw in een bundel worden verkocht.
- **Gereserveerd** of **Verkocht** geldt voor alle geselecteerde modellen. Beheer die status, verkoopdatum en afspraak via de bundel; andere modelgegevens kun je gewoon blijven bewerken.
- De kalender en agenda-export tonen de gezamenlijke afspraak **één keer**, met alle modellen in de omschrijving.
- Een verkochte bundel telt bij de opbrengst **één keer** mee. Het totaal wordt niet automatisch over de modellen verdeeld. De modelregels hebben daarom geen fictieve individuele verkoopprijs.
- Je kunt leden toevoegen/verwijderen zolang minstens twee modellen overblijven. Verwijderde leden krijgen hun eerdere status terug.
- **Bundel opheffen** herstelt de eerdere statussen en laat eventuele oorspronkelijke individuele afspraken weer zien. De bundel blijft als geannuleerd bewaard en kan opnieuw worden gebruikt.
- Haal een model eerst uit de bundel voordat je het naar de prullenbak verplaatst. Wijzigingen van de Excel-identiteit worden gecontroleerd; een ander model met hetzelfde nummer wordt niet automatisch aan de bundel gekoppeld.

## Wijzigingsgeschiedenis per model

Open **Bekijk model > Wijzigingsgeschiedenis**. Vanaf het inschakelen van deze functie worden daadwerkelijke wijzigingen vastgelegd: datum/tijd, actie en waarden vóór/na. Dit omvat modelgegevens, prijs/status, foto's, tekst, verkoopgegevens, afspraken, bundelwijzigingen, prullenbakacties en wijzigingen door opnieuw ingelezen Excel.

**Deze wijziging terugdraaien** zet alleen de velden uit dat moment terug. Andere velden blijven behouden. De server controleert dat de betrokken velden nog dezelfde waarden hebben; een latere wijziging in hetzelfde veld wordt nooit stilzwijgend overschreven. Het terugdraaien krijgt zelf ook een geschiedenisregel.

Toevoegen/verwijderen beheer je via de prullenbak, samenhangende bundelwijzigingen via de bundel. Een Excel-nummer dat inmiddels bij een andere modelidentiteit hoort kan niet met een oude geschiedenisregel worden aangepast. Ontbreekt een oudere foto, herstel die eerst via Back-ups.

De geschiedenis staat samen met bundels in `data/lokale_modellen.json` en blijft behouden na een herstart of Excel-update. Er wordt geen geschiedenis van vóór deze functie verzonnen. Volledige back-upherstelpunten zetten ook de geschiedenis terug naar dat moment; de versie van vóór het terugzetten blijft in het vooraf gemaakte herstelpunt staan.

De volledige ZIP-export bevat ook `verkoopbundels.json` en `wijzigingsgeschiedenis.json`. In de Parquet-export staan per model de bundel-ID, bundelnaam en geschiedenis. Bundeltotalen staan afzonderlijk in de bundelgegevens; tel ze niet per model op.

## Modellen toevoegen, bewerken en verwijderen

### Toevoegen
1. Start de app met `Start inventaris.cmd`.
2. Klik op **+ Model toevoegen**.
3. Vul minstens **automerk** en **model** in. De rest is optioneel. *Conditie* en *Status zoekertje* kies je uit een keuzelijst.
4. Kies foto's, sleep ze in het formulier of plak een gekopieerde afbeelding met **Ctrl+V**. Plak de zoekertjestekst in het tekstveld.
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

### Foto's toevoegen en controleren

In **Model toevoegen** of **Model bewerken** kun je:

- foto's vanuit Verkenner in het formulier slepen, ook meerdere tegelijk;
- een gekopieerde afbeelding of schermafbeelding met **Ctrl+V** plakken; gewone tekst blijft normaal plakken in de tekstvelden;
- **Foto's kiezen…** gebruiken als je liever via de bestandskiezer werkt.

Nieuwe foto's verschijnen meteen in een grote voorvertoning. Klik op een miniatuur om die groter te bekijken, blader met de pijlen en kies **Als hoofdfoto** om de geselecteerde foto vooraan te zetten. De bestaande pijltjes en verwijderknoppen blijven beschikbaar. Toegevoegde foto's worden pas bewaard met **Model opslaan**.

JPG, PNG, WebP en GIF worden ondersteund. De limieten zijn 20 foto's per model, 15 MB per foto en 75 MB aan nog niet opgeslagen foto's. Ongeldige of te grote bestanden worden overgeslagen met een melding; de andere foto's blijven behouden. Plakken werkt wanneer de browser een echte afbeelding ontvangt, bijvoorbeeld na het kopiëren van een schermafbeelding, niet bij alleen een gekopieerde bestandsnaam of afbeeldingslink.

### Foto's en tekst bekijken

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
> Mogelijke dubbele modellen worden gemeld bij toevoegen en bewerken, ook tussen Excel en de pagina. Een tweede exemplaar vereist bevestiging; er wordt niets automatisch samengevoegd.
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
| `beheer.py` | Verkoopbundels, modelgeschiedenis en selectief terugdraaien |
| `collectie.py` | Herstelpunten, volledige export en publicatie-/dubbelcontroles |
| `miniaturen.py` | Maakt miniaturen in `fotos/_mini` |
| `inventaris.template.html` | Sjabloon voor de pagina (HTML, CSS en JavaScript, allemaal ingebed) |
| `test_inventaris.py` | Integratietests |
| `requirements.txt` | Python-afhankelijkheden |
| `data/` | Junction naar de gegevens in OneDrive (niet in Git) |

> [!NOTE]
> Houd `bijwerken.py`, `server.py`, `collectie.py`, `beheer.py`, `miniaturen.py`, `inventaris.template.html` en de `.cmd`-bestanden samen in één map.

---

## Tests

De integratietests draaien in tijdelijke mappen. Je echte collectie wordt dus niet gewijzigd.

```bash
python -m unittest test_inventaris -v
```
