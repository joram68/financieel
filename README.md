# HS4S – Financieel

Persoonlijke budgetplanner voor de HS4S-website.

## Bestanden
- `index.html`: gebruikersinterface (GitHub Pages).
- `Code.gs`: Google Apps Script-backend voor je eigen Google Sheet.

## Belangrijk
De repository bevat geen persoonlijke bedragen. Financiële gegevens worden in je eigen Google Sheet bewaard. Zet nooit je spreadsheet-ID, wachtwoorden of financiële export in GitHub.

## Installatie van de gegevensopslag
1. Upload `Persoonlijk budget 2026.xlsm` naar Google Drive en open het als Google Spreadsheet. Kies **Bestand → Opslaan als Google Spreadsheets** als Drive nog een Excelbestand toont.
2. Open de Google Spreadsheet en kies **Extensies → Apps Script**.
3. Vervang de voorbeeldcode in `Code.gs` door de inhoud van het bestand `Code.gs` uit deze repository.
4. Sla op en voer `setupFinancieel` één keer uit. Geef de gevraagde Google-toestemming.
5. Voer daarna `importeerExcel2026` uit. De bestaande werkbladen **Kosten 2026**, **Samenvatting 2026** en **SaldoOverzicht** worden gebruikt om je gegevens te importeren.
6. Implementeer als webapp: **Uitvoeren als: gebruiker die de webapp opent** en **Toegang: alleen jezelf**. Gebruik de URL die eindigt op `/exec`.
7. Open de Financieel-pagina en vul de webapp-URL in.

De toegang wordt gecontroleerd aan de hand van het Google-account. Deel de webapp niet met anderen.

## Jaarwerking
Een nieuw boekjaar neemt alle onkostenomschrijvingen over. Regels die als **Vast maandelijks** of **Vast jaarlijks** zijn gemarkeerd, nemen automatisch hun maandbedragen mee. Variabele en eenmalige bedragen worden leeggemaakt. Andere bedragen worden leeggemaakt om te voorkomen dat eenmalige uitgaven zomaar opnieuw worden gepland. Controleer de categorie en het soort van elke onkost.

## GitHub Pages
Ga naar **Settings → Pages**, kies **Deploy from a branch**, selecteer **main** en map **/(root)**. De pagina wordt daarna beschikbaar via de GitHub Pages-URL.
