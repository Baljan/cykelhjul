# -*- coding: utf-8 -*-
"""
==========================================================================
 VARIABLER – fyll i allt inför årets cykelfest HÄR.
==========================================================================
 Du ska INTE behöva öppna mailskapare.py eller mailskickare.py.
 Ändra värdena nedan, spara filen (Ctrl+S), och kör sedan skripten som
 vanligt. Platshållare som {parnamn} fylls i automatiskt – låt dem vara
 kvar, men ändra gärna texten runt dem.
==========================================================================
"""

# ========================================================================
# 1. FILER
# ========================================================================

# Sökväg till svars-excelen. Tips: högerklicka filen i Utforskaren ->
# "Kopiera som sökväg" och klistra in mellan citattecknen. Låt r:et stå
# kvar (det gör att \ funkar på Windows).
SVARSFIL = r"C:\Users\there\Cykelfest_Baljan_ht26_svar.xlsx"

# Namnen på flikarna inne i excel-filen
FLIK_DATA     = "Info"   # fliken med alla parens svar
FLIK_FARDPLAN = "Färdplan"    # fliken med vem som äter var

# Filen som de färdiga mailen sparas till / läses från
MAILLISTA = "mail_lista.xlsx"

# Fil där mail som INTE gick att skicka loggas (för uppföljning efteråt)
MISSLYCKADE_FIL = "misslyckade.txt"


# ========================================================================
# 2. KOLUMNNAMN  (måste stavas EXAKT som rubrikerna i excel-filen)
# ========================================================================

KOL_PAR         = "Par"
KOL_PARNAMN     = "Gemensamt parnamn:"
KOL_ADRESS      = "Adress (där ni tänker servera er rätt):"
KOL_PORTKOD     = "Portkod:"
KOL_TEL         = "Telefonnummer vi når er på?"
# Mejl till person 1 i paret
KOL_MEJL        = "Mail person 1"
# Mejl till person 2 i paret. Finns bara EN mejl-kolumn: lämna "" här.
# (Står båda adresserna i samma fält, t.ex. "a@x.se, b@y.se", funkar det
# också – de delas upp automatiskt.)
KOL_MEJL2       = "Mail person 2"
KOL_SPECIALKOST = "Specialkost (OBS: För båda i paret, dvs specificera om specialkosten gäller en eller båda i paret):"

# Namnen på de tre måltiderna (används både som kolumnrubrik i Färdplan
# OCH som text i mailet)
MALTID_FORRATT   = "Förrätt"
MALTID_HUVUDRATT = "Huvudrätt"
MALTID_EFTERRATT = "Efterrätt"

# Namn på Cykelhjul (används i mailet)
CYKELHJUL = "Hedda Carlsson"

# ========================================================================
# 3. DATUM & BELOPP
# ========================================================================

DATUM = "Lördagen den 17e oktober"

BUDGET_HUVUDRATT = "700 kr."   # bidrag till huvudrätts-paret
BUDGET_OVRIGT    = "400 kr."   # bidrag till förrätt/efterrätt

# När man kan hämta ut Flammanförköp + presentkort (används i MAILTEXT)
UTHAMTNING_DAG = "torsdag (15/10)"   # t.ex. "torsdag (15/10)"
UTHAMTNING_TID = "08:00"             # t.ex. "08:00"


# ========================================================================
# 4. SVAR SOM BETYDER "INGEN SPECIALKOST"
# ========================================================================
# Hela svaret jämförs (inte delar av det), så "Inget kött" tas aldrig bort
# av misstag. Lägg till fler varianter här om det dyker upp nya.
NEJ_SVAR = {
    "", "nej", "nope", "ingen", "inget", "inga", "nae", "naj", "nein", "nahhhh", "meh", "nox2",
    "allätare", "vi äter allt", "ingen specialkost", "nej ingen speckost", "nej jättebra",
}


# ========================================================================
# 5. SMTP / MAILKONTO
# ========================================================================

SMTP_SERVER    = "smtp.gmail.com"
SMTP_PORT      = 587                      # 587 för TLS, 465 för SSL
SMTP_ANVANDARE = "tess@baljan.org"   # avsändaradress (det konto som loggar in)
# Namnet mottagaren ser som avsändare (t.ex. "Cykelhjul"). Lämna "" för att
# bara visa mailadressen.
AVSANDARNAMN   = "Cykelhjul"

# Lösenordet läses ur en miljövariabel med DETTA namn (så att det inte
# ligger i klartext i koden). OBS: Gmail kräver numera ett "App-lösenord"
# (skapas under kontots säkerhetsinställningar, kräver att 2FA är på).
SMTP_LOSENORD_ENV = "CYKELFEST_SMTP_PASSWORD"


# ========================================================================
# 6. MAILETS TEXT
# ========================================================================

# Ämnesrad
AMNE = "VIKTIG INFO: BALJANS SUPERBALLA CYKELFEST"

# --- Själva mailet -------------------------------------------------------
# Platshållare som fylls i automatiskt:
#   {parnamn}    parets namn
#   {datum}      = DATUM ovan
#   {uthamtning_dag} {uthamtning_tid}   = UTHAMTNING_* ovan
#   {cykelhjul}  = CYKELHJUL ovan (namnet i signaturen)
#   {forratt} {huvudratt} {efterratt}   schemat (genereras av skriptet)
# \n = ny rad, \n\n = tom rad mellan stycken.
MAILTEXT = """Plinegeling {parnamn}!

{datum} är det DAGS för Baljans Balla Cykelfest a.k.a Baljans Balla BikeBröl a.k.a Tour de Baljan a.k.a Giro di Baljan a.k.a. Baljanrundan!

Fr.o.m. {uthamtning_dag} klockan {uthamtning_tid} kan man hämta ut sina: Hemköp-presentkort + anmälda Flammanförköp, under jourtider i Baljans kontor Nollberga.

Ert schema för kvällen ser ut såhär:

18:00 Förrätt - {forratt}

19:30 Huvudrätt - {huvudratt}

21:00 Efterrätt - {efterratt}

Efterpärla - Dörrarna står öppna på Flamman från 22.30!

VIKTIGT: Flammanförköpet är giltigt till 23.30, så för att komma in via förköps-kön behöver man stå i kön (med biljetten!!) innan dess.
ABSOLUT SENAST ÄR ALLTSÅ 23:29:59! Är man sen får man stå i vanliga kön och betala ordinarie pris!

Ses på lördag! TAGGA!
---------------------
{cykelhjul}
Cykelhjul,
Baljans Balla Styrelse"""

# --- Text till ett par som SJÄLVA ska servera en måltid ------------------
# Platshållare: {maltid} {budget} {specialkost}
SERVERA_TEXT = """Ni har fått äran att servera {maltid}.
BUDGET: Ni kommer få ett bidrag på {budget}
SPECIALKOST: {specialkost}
Observera att er egna specialkost står även med här."""

# --- Text till ett par som är GÄST hos ett annat par ---------------------
# Platshållare: {adress} {portkod} {tel} {vard}
GAST_TEXT = """Här är ni varmt välkomna till {adress}
Portkod: {portkod}
Nummer till värd: {tel}
Värd: {vard}"""
