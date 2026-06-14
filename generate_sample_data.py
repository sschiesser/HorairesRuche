"""
Génère des fichiers Excel d'exemple dans le dossier data/ :
  - plages_horaires.xlsx : les créneaux/tâches à couvrir, jour par jour
  - personnes.xlsx       : l'équipe et ses capacités générales
  - blocages.xlsx        : les indisponibilités déclarées par chacun (priorité 1 ou 2)
  - shifts_fixes.xlsx    : les shifts déjà imposés pour la semaine à certaines personnes

Il n'y a pas de "périodes" fixes communes à toute la semaine : une même tâche
peut avoir plusieurs créneaux dans la journée (ex: Accueil 13h-15h puis
20h-24h), avec des horaires et des besoins en personnel différents, et tout
cela peut varier d'un jour à l'autre.

Ces fichiers servent de gabarit et de jeu de données de démonstration pour
generate_planning.py. Remplacez-les par les données réelles de l'événement
en conservant les mêmes noms de colonnes.
"""

import random
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent / "data"
DATA_DIR.mkdir(exist_ok=True)

random.seed(42)

JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# Pour chaque tâche, un ou plusieurs créneaux types dans la journée.
# Chaque créneau définit une plage d'heures de début, une plage de durées
# (en heures) et une plage de besoins en personnel : les valeurs exactes sont
# tirées au sort indépendamment pour chaque jour, donc le résultat varie d'un
# jour à l'autre tout en restant réaliste (ex: plus de monde le soir).
TACHES = {
    "Accueil": [
        {"heure_debut": (12, 14), "duree": (2, 3), "requis": (1, 2)},   # créneau midi
        {"heure_debut": (19, 20), "duree": (3, 4), "requis": (3, 4)},   # créneau soir
    ],
    "Billetterie": [
        {"heure_debut": (12, 14), "duree": (2, 3), "requis": (1, 1)},
        {"heure_debut": (19, 20), "duree": (3, 4), "requis": (1, 2)},
    ],
    "Bar": [
        {"heure_debut": (12, 14), "duree": (2, 3), "requis": (1, 2)},
        {"heure_debut": (19, 20), "duree": (3, 5), "requis": (3, 5)},
    ],
    "Vestiaire": [
        {"heure_debut": (18, 19), "duree": (4, 5), "requis": (1, 2)},
    ],
    "Sécurité": [
        {"heure_debut": (12, 14), "duree": (2, 3), "requis": (1, 2)},
        {"heure_debut": (19, 20), "duree": (4, 5), "requis": (2, 3)},
    ],
    "Technique": [
        {"heure_debut": (11, 13), "duree": (3, 4), "requis": (1, 2)},
        {"heure_debut": (18, 19), "duree": (5, 6), "requis": (2, 3)},
    ],
}

# --- 1. Plages horaires : un ou plusieurs créneaux par tâche et par jour ----
plages_rows = []
for jour in JOURS:
    for tache, slots in TACHES.items():
        for slot in slots:
            debut_h = random.randint(*slot["heure_debut"])
            fin_h = min(debut_h + random.randint(*slot["duree"]), 24)
            plages_rows.append({
                "Jour": jour,
                "Tâche": tache,
                "Début": f"{debut_h:02d}:00",
                "Fin": f"{fin_h:02d}:00",
                "Nb_personnes": random.randint(*slot["requis"]),
            })

pd.DataFrame(plages_rows).to_excel(DATA_DIR / "plages_horaires.xlsx", index=False)


# --- 2. Personnes -----------------------------------------------------------
TACHES_LISTE = list(TACHES.keys())
personnes_rows = []
for i in range(1, 71):
    nb_caps = random.randint(2, 4)
    caps = random.sample(TACHES_LISTE, nb_caps)
    # ~10% de personnes "polyvalentes", capables de toutes les tâches
    capacites = "Tous" if random.random() < 0.1 else ", ".join(caps)
    personnes_rows.append({"Nom": f"Personne {i:02d}", "Capacités": capacites})

pd.DataFrame(personnes_rows).to_excel(DATA_DIR / "personnes.xlsx", index=False)


# --- 3. Blocages -------------------------------------------------------------
HEURES_BLOCAGE_POSSIBLES = list(range(10, 22))
DUREES_BLOCAGE_POSSIBLES = [2, 3, 4]  # heures

blocages_rows = []
for i in range(1, 71):
    nom = f"Personne {i:02d}"
    if random.random() < 0.4:
        for _ in range(random.randint(1, 2)):
            jour = random.choice(JOURS)
            debut_h = random.choice(HEURES_BLOCAGE_POSSIBLES)
            fin_h = min(debut_h + random.choice(DUREES_BLOCAGE_POSSIBLES), 24)
            blocages_rows.append({
                "Nom": nom,
                "Jour": jour,
                "Début": f"{debut_h:02d}:00",
                "Fin": f"{fin_h:02d}:00",
                "Priorité": random.choice([1, 2]),
            })

pd.DataFrame(blocages_rows).to_excel(DATA_DIR / "blocages.xlsx", index=False)


# --- 4. Shifts fixes (1 ou 2 par personne concernée, pour toute la semaine) -
def trouver_plage(jour, tache, occurrence=0):
    matches = [row for row in plages_rows if row["Jour"] == jour and row["Tâche"] == tache]
    if occurrence >= len(matches):
        raise ValueError(f"Pas assez de créneaux pour {jour} / {tache}")
    return matches[occurrence]

p1 = trouver_plage("Lundi", "Technique", occurrence=1)    # créneau du soir
p2a = trouver_plage("Mardi", "Sécurité", occurrence=0)     # créneau de midi
p2b = trouver_plage("Vendredi", "Sécurité", occurrence=1)  # créneau du soir

shifts_fixes_rows = [
    {"Nom": "Personne 01", "Jour": p1["Jour"], "Tâche": p1["Tâche"], "Début": p1["Début"], "Fin": p1["Fin"]},
    {"Nom": "Personne 02", "Jour": p2a["Jour"], "Tâche": p2a["Tâche"], "Début": p2a["Début"], "Fin": p2a["Fin"]},
    {"Nom": "Personne 02", "Jour": p2b["Jour"], "Tâche": p2b["Tâche"], "Début": p2b["Début"], "Fin": p2b["Fin"]},
]
pd.DataFrame(shifts_fixes_rows).to_excel(DATA_DIR / "shifts_fixes.xlsx", index=False)

print(f"Fichiers d'exemple générés dans {DATA_DIR}")
