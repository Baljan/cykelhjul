# TEKKANS GUIDE TILL ATT BEMÄSTRA CYKELSCRIPTET

Guide för dig som är **Cykelhjul** (eller mest troligt Nörd). Verktyget bygger ett personligt mail till varje par (schema för kvällen, adresser, portkoder, specialkost) och skickar ut dem automatiskt.

Nästan allt du ändrar finns i en enda fil: `variabler.py`. Läs den här guiden uppifrån och ned en gång, så är du igång!

## Filerna i mappen

| Fil | Vad den gör | Rör du den? |
|-----|-------------|-------------|
| **`variabler.py`** | ALLT du fyller i: sökväg till Excel, kolumnnamn, mailtext, datum, belopp, avsändare. | Ja - det är här du jobbar |
| `mailskapare.py` | Läser Excel och bygger alla mail → sparar dem i `mail_lista.xlsx`. | Nej |
| `mailskickare.py` | Skickar ut mailen från `mail_lista.xlsx`. | Nej |
| `mail_lista.xlsx` | Skapas automatiskt. Innehåller de färdiga mailen – **öppna och läs igenom före utskick!** | Läs |
| `misslyckade.txt` | Skapas bara om något mail inte gick fram – lista på adresser att följa upp. | Läs |

> Tanken: du öppnar bara `variabler.py`. De två andra `.py`-filerna ska lämnas ifred.

## Steg 0: Förberedelser

1. Installera Python, om du inte redan har det
2. Hämta detta repot
3. Kör:
```powershell
  pip install pandas openpyxl
```
Klart – det här behöver du inte göra om nästa gång.

## Steg 1: Excel-filen med svaren

Verktyget läser en Excel-fil (den ni samlar in parens svar i). Den måste ha **två flikar**:

1. **Info-fliken** (parens svar) – standardnamnet i koden är `Info`.
2. **Färdplans-fliken** – vem som äter vilken måltid var. Standardnamn: `Färdplan`.

**Info-fliken behöver dessa kolumner** (rubrikerna måste stämma med namnen i `variabler.py`):
`Par`, `Gemensamt parnamn:`, `Adress (där ni tänker servera er rätt):`, `Portkod:`,
`Telefonnummer vi når er på?`, `Mail person 1`, `Mail person 2`, och specialkost-kolumnen.

**Färdplan-fliken** behöver kolumnerna `Par`, `Förrätt`, `Huvudrätt`, `Efterrätt`
(varje par pekar på vilket par de äter respektive måltid hos).

> Heter en rubrik något annat i år? Då döper du **inte** om i Excel – du ändrar namnet i `variabler.py` istället (se avsnittet *"Hur man döper om saker"* längre ned).

## Steg 2: Fyll i `variabler.py`

Öppna `variabler.py`. Den är indelad i numrerade sektioner – gå igenom variablerna uppifrån:

**1. Filer**
- `SVARSFIL` – sökväg till Excel-filen (högerklicka filen → *"Kopiera som sökväg"*)
- `FLIK_DATA` / `FLIK_FARDPLAN` – namnen på flikarna i Excel
- `MAILLISTA` / `MISSLYCKADE_FIL` – filnamn som skapas automatiskt

**2. Kolumnnamn** (måste stämma EXAKT med rubrikerna i Excel)
- `KOL_PAR`, `KOL_PARNAMN`, `KOL_ADRESS`, `KOL_PORTKOD`, `KOL_TEL`
- `KOL_MEJL` / `KOL_MEJL2` – person 1:s och person 2:s mail
- `KOL_SPECIALKOST`
- `MALTID_FORRATT`, `MALTID_HUVUDRATT`, `MALTID_EFTERRATT`
- `CYKELHJUL` – namnet i signaturen

**3. Datum & belopp**
- `DATUM`
- `BUDGET_HUVUDRATT`, `BUDGET_OVRIGT`
- `UTHAMTNING_DAG`, `UTHAMTNING_TID`

**4. Specialkost**
- `NEJ_SVAR` – svar som betyder "ingen specialkost"

**5. SMTP / mailkonto**
- `SMTP_SERVER`, `SMTP_PORT`
- `SMTP_ANVANDARE` – avsändaradress
- `AVSANDARNAMN` – namnet mottagaren ser
- `SMTP_LOSENORD_ENV`

**6. Mailets text**
- `AMNE` – ämnesrad
- `MAILTEXT`, `SERVERA_TEXT`, `GAST_TEXT` – själva texterna

**Spara filen när du är klar.**

> I texterna finns platshållare i klammer, t.ex. `{parnamn}`. De fylls i automatiskt – **ändra texten runt dem, men låt klammrarna vara kvar** och skriv dem exakt (gemener).

## Hur man "döper om" saker

**Grundprincipen:** koden letar efter *exakt* de namn du skrivit i `variabler.py`.
Vill du ändra något byter du strängen där – inte i koden och oftast inte i Excel.

| Vill du ändra... | Ändra denna rad i `variabler.py` |
|------------------|----------------------------------|
| Namnet som står i signaturen | `CYKELHJUL = "Hedda Carlsson"` |
| Namnet mottagaren ser som avsändare | `AVSANDARNAMN = "Cykelhjul"` |
| Vilken Excel-fil som läses | `SVARSFIL = r"C:\..."` |
| Namnet på en flik i Excel | `FLIK_DATA` / `FLIK_FARDPLAN` |
| En kolumn heter annorlunda i år | motsvarande `KOL_...`-rad |
| Datum, budget, ämnesrad, själva texten | sektion 3 och 6 |

**Måste stämma EXAKT** – stor/liten bokstav, mellanslag och kolon räknas.
`Mail person 1` ≠ `mail person 1` ≠ `Mail Person 1:`. Stämmer det inte hittar koden inte kolumnen.

> `mailskapare.py` varnar i terminalen om ett kolumnnamn i `variabler.py` inte finns i Excel – så du ser direkt om du stavat fel.

## Steg 3: Mailkonto och lösenord

Mailen skickas via ett Gmail-konto. Av säkerhetsskäl ligger lösenordet **inte** i koden,
utan i en *miljövariabel* som heter `CYKELFEST_SMTP_PASSWORD`.

**1. Skapa ett App-lösenord i Google** (vanliga Gmail-lösenordet funkar inte):
- Slå på 2-stegsverifiering på kontot.
- Google-kontot → Säkerhet → **App-lösenord** → skapa ett → kopiera den 16-teckens koden.

**2. Lägg koden i miljövariabeln** – kör i terminalen (samma fönster du sedan skickar från):

```powershell
$env:CYKELFEST_SMTP_PASSWORD = "ditt-app-lösenord"
```

> Detta gäller bara det öppna terminalfönstret. Öppnar du ett nytt måste du köra raden igen.
> (Vill du slippa det kan du använda `setx CYKELFEST_SMTP_PASSWORD "..."` en gång – gäller då i nya fönster.)

## Steg 4: Kör igång

I terminalen, i mappen `cykelfest`, kör de två skripten i tur och ordning:

```powershell
python mailskapare.py
```
Skapar alla mail och sparar dem i `mail_lista.xlsx`. Skriver ut hur många par som når båda / bara en.

**Öppna nu `mail_lista.xlsx` och läs några mail** – stämmer datum, scheman och adresser? Rätta i `variabler.py` och kör om vid behov.

När allt ser bra ut – skicka:
```powershell
python mailskickare.py
```

> `mailskickare.py` skickar skarpt till alla på en gång, utan ångra-knapp. Dubbelkolla `mail_lista.xlsx` först!

## Testa säkert innan skarpt utskick

Innan du skickar till alla ~100 personer – testa att det landar rätt:

1. Gör en liten testvariant av Excel-filen med bara **ett par**, där du satt **din egen adress** i `Mail person 1`.
2. Peka `SVARSFIL` mot testfilen, kör `mailskapare.py` och sedan `mailskickare.py`.
3. Kolla mailet i din egen inkorg – ser avsändare, ämne, schema och signatur rätt ut?
4. Byt tillbaka `SVARSFIL` till den riktiga filen inför det skarpa utskicket.

## Om något går fel

- **Ett eller flera mail gick inte iväg** (t.ex. felskriven adress): skriptet hoppar över dem,
  fortsätter med resten, och skriver på slutet ut *antal skickade*, *antal misslyckade* och
  *listan på adresser* som strulade. Samma lista sparas i **`misslyckade.txt`**.
- **`KeyError` om en kolumn** → ett kolumnnamn i `variabler.py` stämmer inte med Excel. Kolla stavningen.
- **Inloggningsfel** → fel app-lösenord, eller så är miljövariabeln inte satt i det fönstret (se Steg 3).
- Tänk på att "skickat" betyder att Gmail accepterade mailet – en felstavad men giltig adress kan studsa
  tillbaka senare som ett *"Delivery failed"*-mail i din egen inkorg. Kika där efter utskicket.

## Snabb checklista

1. `pip install pandas openpyxl` (första gången)
2. Fyll i `variabler.py` – sökväg, kolumnnamn, datum, text, avsändare, `CYKELHJUL`
3. Sätt lösenordet: `$env:CYKELFEST_SMTP_PASSWORD = "..."`
4. `python mailskapare.py`
5. **Öppna och läs `mail_lista.xlsx`**
6. `python mailskickare.py`
7. Kolla sammanfattningen + ev. `misslyckade.txt`

Lycka till,
PoK Tess