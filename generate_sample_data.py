"""
Génère des fichiers Excel d'exemple dans le dossier data/ :
  - plages_horaires.xlsx : les créneaux/tâches à couvrir, jour par jour
  - personnes.xlsx       : l'équipe et ses capacités générales
  - blocages.xlsx        : les indisponibilités déclarées par chacun (priorité 1 ou 2)

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

# Pour chaque tâche : nombre de personnes requises par période de la journée
TACHES = {
    "Accueil":     {"Après-midi": 2, "Soirée": 3},
    "Billetterie": {"Après-midi": 2, "Soirée": 2},
    "Bar":         {"Après-midi": 2, "Soirée": 4},
    "Vestiaire":   {"Après-midi": 1, "Soirée": 2},
    "Sécurité":    {"Après-midi": 2, "Soirée": 3},
    "Technique":   {"Après-midi": 2, "Soirée": 2},
}

PERIODES = {
    "Après-midi": ("13:00", "18:00"),
    "Soirée": ("18:00", "23:30"),
}

# --- 1. Plages horaires ---------------------------------------------------
rows = []
for jour in JOURS:
    for periode, (debut, fin) in PERIODES.items():
        for tache, requis in TACHES.items():
            rows.append({
                "Jour": jour,
                "Tâche": tache,
                "Début": debut,
                "Fin": fin,
                "Nb_personnes": requis[periode],
            })

pd.DataFrame(rows).to_excel(DATA_DIR / "plages_horaires.xlsx", index=False)

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
blocages_rows = []
for i in range(1, 71):
    nom = f"Personne {i:02d}"
    if random.random() < 0.4:
        for _ in range(random.randint(1, 2)):
            jour = random.choice(JOURS)
            periode = random.choice(list(PERIODES.keys()))
            debut, fin = PERIODES[periode]
            priorite = random.choice([1, 2])
            blocages_rows.append({
                "Nom": nom,
                "Jour": jour,
                "Début": debut,
                "Fin": fin,
                "Priorité": priorite,
            })

pd.DataFrame(blocages_rows).to_excel(DATA_DIR / "blocages.xlsx", index=False)

print(f"Fichiers d'exemple générés dans {DATA_DIR}")
