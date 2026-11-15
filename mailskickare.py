import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
import os

import variabler as v   # <-- SMTP-inställningar och ämnesrad ligger i variabler.py

# Lösenordet läses ur miljövariabeln (se SMTP_LOSENORD_ENV i variabler.py).
# Gmail kräver ett App-lösenord – "less secure apps" finns inte längre.
smtp_password = os.environ[v.SMTP_LOSENORD_ENV]

# Avsändare som mottagaren ser: "Cykelhjul <tess@baljan.org>" om AVSANDARNAMN
# är ifyllt, annars bara adressen.
avsandare = formataddr((v.AVSANDARNAMN, v.SMTP_ANVANDARE)) if v.AVSANDARNAMN else v.SMTP_ANVANDARE

# Load data from an Excel file
df = pd.read_excel(v.MAILLISTA)

# Create an SMTP connection
server = smtplib.SMTP(v.SMTP_SERVER, v.SMTP_PORT)
server.starttls()
server.login(v.SMTP_ANVANDARE, smtp_password)

# Håll reda på vad som gick bra respektive dåligt
skickade = 0
misslyckade = []   # lista av mejladresser som inte gick iväg

# Loop through each row
for index, row in df.iterrows():
    # 'Mailadress' kan innehålla flera adresser (båda i paret), separerade
    # med komma. Dela upp till en lista så alla verkligen får mailet.
    mottagare = [a.strip() for a in str(row['Mailadress']).split(',') if a.strip()]
    email_text = row['Mail']

    if not mottagare:
        misslyckade.append(f"(rad {index}: ingen adress ifylld)")
        continue

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = avsandare
    msg['To'] = ", ".join(mottagare)
    msg['Subject'] = v.AMNE
    msg.attach(MIMEText(email_text, 'plain'))

    # Send the email (to_addrs måste vara en LISTA för att nå alla).
    # Går det inte -> spara adressen och fortsätt med resten.
    try:
        vägrade = server.sendmail(v.SMTP_ANVANDARE, mottagare, msg.as_string())
        if vägrade:
            misslyckade.extend(vägrade.keys())   # adresser servern nekade
        else:
            skickade += 1
    except smtplib.SMTPException:
        misslyckade.extend(mottagare)

# Close the SMTP connection
server.quit()

# Sammanfattning
print("\n================ SAMMANFATTNING ================")
print(f"Skickade:    {skickade}")
print(f"Misslyckade: {len(misslyckade)}")
if misslyckade:
    print("\nAdresser som inte gick iväg:")
    for adress in misslyckade:
        print(f"   {adress}")

    # Spara listan till fil så den finns kvar efter att terminalen stängts
    from datetime import datetime
    with open(v.MISSLYCKADE_FIL, "w", encoding="utf-8") as f:
        f.write(f"Misslyckade mail – {datetime.now():%Y-%m-%d %H:%M}\n")
        f.write(f"Skickade: {skickade}   Misslyckade: {len(misslyckade)}\n\n")
        for adress in misslyckade:
            f.write(adress + "\n")
    print(f"\nSparat till {v.MISSLYCKADE_FIL}")
