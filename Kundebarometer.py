# -*- coding: utf-8 -*-
"""
Created on Tue Apr  1 18:33:08 2025

@author: oag

Program for å virtualisere kundetilfredshet.

"""

import pandas as pd
import matplotlib.pyplot as plt


# Utføring - Del a: Lese inn dataene fra uke 24
def Les_inn_uke_24():
    uke_logg = pd.read_excel("support_uke_24.xlsx", engine="openpyxl")
    uke_logg = uke_logg.fillna("")

    return (
        uke_logg["Ukedag"].tolist(),
        uke_logg["Klokkeslett"].tolist(),
        uke_logg["Varighet"].tolist(),
        uke_logg["Tilfredshet"].tolist()
    )


# Utføring - Del b: Finn og presenter antall henveldelser pr dag
def Daglige_henvendelser(u_dager):
    Antall_Ma = u_dager.count('Mandag')
    Antall_Ti = u_dager.count('Tirsdag')
    Antall_On = u_dager.count('Onsdag')
    Antall_To = u_dager.count('Torsdag')
    Antall_Fr = u_dager.count('Fredag')

    # Setter opp array for diagram
    Dager = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']
    Henvendelser = [Antall_Ma, Antall_Ti, Antall_On, Antall_To, Antall_Fr]

    # Plot diagrammet
    plt.bar(Dager, Henvendelser)
    plt.title('Antall henvendelser pr. dag for uke 24')
    plt.xlabel('Ukedag')
    plt.ylabel('Antall')
    plt.show()


# Utføring - Del c: Finn minste og lengste samtaletid for uke 24
def Finn_minste_lengste_samtaletid(varighet):
    Min_samtaletid = ""
    Max_samtaletid = ""
    for i in range(len(varighet)):
        if Min_samtaletid == "" or Min_samtaletid > varighet[i]:
            Min_samtaletid = varighet[i]
        if Max_samtaletid == "" or Max_samtaletid < varighet[i]:
            Max_samtaletid = varighet[i]
    print()
    print()
    print("Minimum samtaletid i uke 24 var: ", Min_samtaletid)
    print("Maksimal samtaletid i uke 24 var:", Max_samtaletid)


# Utføring - Del c: Beregn gjennomsnittlig samtaletid for uke 24
def Finn_gjennomsnitt_samtaletid(varighet):
    Total_samtaletid_sec = 0
    for i in range(len(varighet)):
        timer, minutter, sekunder = map(int, varighet[i].split(':'))
        Totalt_sekunder = timer * 3600 + minutter * 60 + sekunder
        Total_samtaletid_sec += Totalt_sekunder
    Gjennomsnitt_samtaletid = Total_samtaletid_sec/len(varighet)
    timer = Gjennomsnitt_samtaletid // 3600
    sekunder = Gjennomsnitt_samtaletid % 3600
    minutter = Gjennomsnitt_samtaletid // 60
    sekunder = Gjennomsnitt_samtaletid % 60
    print()
    print()
    print("Gjennomsnittlig samtaletid i uke 24 var:", f"{int(timer):02}:{int(minutter):02}:{int(sekunder):02}")


# Utføring - Del e: Finn antall henvendelser pr. vaktbolk for uke 24
def Finn_henvendelser_vakt_bolk(kl_slett):
    Antall_08_10 = 0
    Antall_10_12 = 0
    Antall_12_14 = 0
    Antall_14_16 = 0

    for i in range(len(kl_slett)):
        if kl_slett[i] < "10:00:00":
            Antall_08_10 += 1
        if kl_slett[i] >= "10:00:00" and kl_slett[i] < "12:00:00":
            Antall_10_12 += 1
        if kl_slett[i] >= "12:00:00" and kl_slett[i] < "14:00:00":
            Antall_12_14 += 1
        if kl_slett[i] >= "14:00:00":
            Antall_14_16 += 1

    # Setter opp array for diagram
    Vakt_bolker = ['08-10', '10-12', '12-14', '14-16']
    Henvendelser = [Antall_08_10, Antall_10_12, Antall_12_14, Antall_14_16]

    # Plot diagrammet
    plt.pie(Henvendelser, labels=Vakt_bolker, autopct=lambda pct: f'{round(pct / 100.*sum(Henvendelser))}')
    plt.title('Antall henvendelser pr. vaktbolk i uke 24')
    plt.show()


# Utføring - Del f: Finn NPS score for kunder som validerte tjenesten uke 24
def Finn_NPS_score(score):
    Antall_misfornøyde = 0
    Antall_nøytrale = 0
    Antall_fornøyde = 0

    for i in range(len(score)):
        if score[i] != "":
            int_score = int(score[i])
            if int_score <= 6:
                Antall_misfornøyde += 1
            if int_score >= 7 and int_score <= 8:
                Antall_nøytrale += 1
            if int_score >= 9:
                Antall_fornøyde += 1
    Antall_svar = Antall_misfornøyde + Antall_nøytrale + Antall_fornøyde

    NPS = (Antall_fornøyde/Antall_svar*100) - (Antall_misfornøyde/Antall_svar*100)

    # Setter opp array for diagram
    values = [Antall_misfornøyde, Antall_nøytrale, Antall_fornøyde]
    colors = ['red', 'yellow', 'green']

    # Lag kakediagram
    fig, ax = plt.subplots(figsize=(8, 8))
    wedges, texts, autotexts = ax.pie(
        values,
        autopct='%.1f%%',
        startangle=90,
        colors=colors,
        textprops=dict(color="black", fontsize=20),
        pctdistance=0.8
    )

    # Lag tom sirkel i midten av kakediagrammet
    centre_circle = plt.Circle((0, 0), 0.60, fc='white')
    fig.gca().add_artist(centre_circle)

    # Legg NPS verdien i sirkelen
    plt.text(0, 0, f'NPS\n{NPS:.0f}', ha='center', va='center', fontsize=60, fontweight='bold')

    # Plot diagrammet
    ax.axis('equal')
    plt.title('NPS for avdelingen MORSE - uke 24', fontsize=30, fontweight='bold')
    plt.tight_layout()
    plt.show()


# Del a: Lese inn dataene fra uke 24
u_dag, kl_slett, varighet, score = Les_inn_uke_24()

# Del b: Finn og presenter antall henveldelser pr dag
Daglige_henvendelser(u_dag)

# Del c: Finn minste og lengste samtaletid for uke 24
Finn_minste_lengste_samtaletid(varighet)

# Del d: Beregn gjennomsnittlig samtaletid for uke 24
Finn_gjennomsnitt_samtaletid(varighet)

# Del e: Finn antall henvendelser pr. vaktbolk for uke 24
Finn_henvendelser_vakt_bolk(kl_slett)

# Del f: Finn NPS score for kunder som validerte tjenesten uke 24
Finn_NPS_score(score)

# Kun for ekstra linjeskift før prompt
print()
