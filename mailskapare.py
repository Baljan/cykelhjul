import openpyxl as op
import pandas as pd
import numpy as np
import re

import variabler as v   # <-- all text, datum, kolumnnamn m.m. ligger i variabler.py


# Här laddas data från excel-filen
data = pd.read_excel(v.SVARSFIL, v.FLIK_DATA)
färdplan = pd.read_excel(v.SVARSFIL, v.FLIK_FARDPLAN)


# Kontroll: varnar om en kolumn i variabler.py inte hittas i excel-filen,
# så man inte får tyst tomma fält i mailen. KOL_MEJL2 kollas bara om ifylld.
def kontrollera_kolumner():
    finns = set(data.columns) | set(färdplan.columns)
    förväntade = [
        v.KOL_PAR, v.KOL_PARNAMN, v.KOL_ADRESS, v.KOL_PORTKOD, v.KOL_TEL,
        v.KOL_MEJL, v.KOL_SPECIALKOST,
        v.MALTID_FORRATT, v.MALTID_HUVUDRATT, v.MALTID_EFTERRATT,
    ]
    if v.KOL_MEJL2:
        förväntade.append(v.KOL_MEJL2)

    saknade = [k for k in förväntade if k not in finns]
    if saknade:
        print("⚠️  VARNING: dessa kolumner från variabler.py finns INTE i excel-filen:")
        for k in saknade:
            print(f"      {k!r}")
        print("   Kontrollera stavningen (exakt som rubriken i excel). Fälten blir annars tomma.\n")

kontrollera_kolumner()

data[v.KOL_TEL] = data[v.KOL_TEL].apply(str)


# Här mergas allt till en dataframe
cols = list(set(data.columns) & set(färdplan.columns))
dataset = pd.merge(data, färdplan, on=cols)


def är_nej(svar):
    # Svar som bara betyder "ingen specialkost" (se NEJ_SVAR i variabler.py).
    text = re.sub(r"\(.*?\)", "", str(svar))  # tar bort kommentarer inom parentes
    text = text.lower().strip(" .,!:)(\"'-\n\t")
    return text in v.NEJ_SVAR


# Här är funktionen som hämtar informationen om respektive par
def get_info(ID, dataset, måltid):

    for index, row in dataset.iterrows():
        if dataset[v.KOL_PAR][index] == ID:
            måltids_ID = dataset[måltid][index]  # hos vilket par ID ska äta måltiden
            if måltids_ID == ID:
                # Paret serverar själva denna måltid
                gäster = dataset[dataset[måltid] == måltids_ID]  # alla par som äter här, inkl. värdparet
                koster = [str(kost).strip() for kost in gäster[v.KOL_SPECIALKOST]
                          if pd.notna(kost) and not är_nej(kost)]
                allergier = ", ".join(koster) + "." if koster else "Ingen specialkost."
                budget = v.BUDGET_HUVUDRATT if måltid == v.MALTID_HUVUDRATT else v.BUDGET_OVRIGT
                return v.SERVERA_TEXT.format(maltid=måltid, budget=budget, specialkost=allergier)
            else:
                # Paret är gäst – hämta värdparets uppgifter
                for index, row in dataset.iterrows():
                    if dataset[v.KOL_PAR][index] == måltids_ID:
                        body = v.GAST_TEXT.format(
                            adress=dataset[v.KOL_ADRESS][index],
                            portkod=dataset[v.KOL_PORTKOD][index],
                            tel=dataset[v.KOL_TEL][index],
                            vard=dataset[v.KOL_PARNAMN][index],
                        )
                return body


def samla_mejl(row):
    # Slår ihop mejladresser så BÅDA i paret nås. Tar med KOL_MEJL och (om
    # ifylld) KOL_MEJL2, och delar upp fält som innehåller flera adresser
    # separerade med , ; eller mellanslag. Dubbletter tas bort.
    kolumner = [v.KOL_MEJL]
    if v.KOL_MEJL2:
        kolumner.append(v.KOL_MEJL2)

    adresser = []
    for kol in kolumner:
        värde = row.get(kol)
        if pd.isna(värde):
            continue
        for adr in re.split(r"[,;\s]+", str(värde).strip()):
            adr = adr.strip()
            if adr and adr not in adresser:
                adresser.append(adr)

    return ", ".join(adresser)


# Här byggs själva mailtexten per par
# Vill du ändra texten gör du det i MAILTEXT i variabler.py
def get_mail(df):
    mail_list = [None] * len(df)
    mejl_list = [None] * len(df)

    for index, row in df.iterrows():
        mail_list[index] = v.MAILTEXT.format(
            parnamn=df[v.KOL_PARNAMN][index],
            datum=v.DATUM,
            uthamtning_dag=v.UTHAMTNING_DAG,
            uthamtning_tid=v.UTHAMTNING_TID,
            cykelhjul=v.CYKELHJUL,
            forratt=get_info(df[v.KOL_PAR][index], dataset, v.MALTID_FORRATT),
            huvudratt=get_info(df[v.KOL_PAR][index], dataset, v.MALTID_HUVUDRATT),
            efterratt=get_info(df[v.KOL_PAR][index], dataset, v.MALTID_EFTERRATT),
        )
        mejl_list[index] = samla_mejl(row)

    lib = {
        'ID': df[v.KOL_PAR],
        'Mailadress': mejl_list,
        'Mail': mail_list,
    }

    return pd.DataFrame(lib)


mail_list = get_mail(dataset)
mail_list.to_excel(v.MAILLISTA)

# Sammanfattning: hur många mottagare varje par får
antal_mottagare = mail_list['Mailadress'].apply(
    lambda s: 0 if not str(s).strip() else len([a for a in str(s).split(',') if a.strip()])
)
utan = int((antal_mottagare == 0).sum())
en = int((antal_mottagare == 1).sum())
flera = int((antal_mottagare >= 2).sum())

print(f"Klart! {len(mail_list)} mail skrevs till {v.MAILLISTA}")
print(f"   Når båda i paret: {flera} par")
print(f"   Når bara en person:  {en} par")
if utan:
    print(f"   ⚠️  UTAN mejladress (skickas ej): {utan} par")
