
"""
Génère le planning d'une équipe (60-80 personnes, 2 shifts/jour) pour un jour
donné, à partir de fichiers Excel saisis par le planificateur, et produit un
classeur Excel indiquant qui est affecté à quoi ce jour-là.

Fichiers d'entrée attendus dans data/ :
  plages_horaires.xlsx
      Un onglet par jour. Chaque onglet contient :
      Tâche | Début | Fin | Nb_personnes
      -> une ligne par créneau à couvrir ce jour-là (une tâche, un horaire,
         et le nombre de personnes nécessaires sur ce créneau).

  personnes.xlsx
      Nom | Prénom | Fonction | Capacités
      -> une ligne par membre de l'équipe. "Nom" et "Prénom" identifient la
         personne (plusieurs personnes peuvent partager le même nom de
         famille). "Fonction" détermine le nombre maximum de shifts par
         jour pour cette personne : "Junior.e" -> 1, "Staff"/"Adjoint.e"
         -> 2, "COF" -> 0 (jamais affecté, sauf shift fixe). "Capacités"
         liste les tâches que la personne peut effectuer. Trois syntaxes :
           - vide / "Tous" / "*"           → toutes les tâches
           - "Bar loges, Bar public"        → uniquement ces tâches
           - "* - Bar loges, Bar public"    → toutes les tâches SAUF ces-là
         Une personne avec une seule capacité (ex: "Bourdon.ne") s'auto-gère
         sur cette tâche et n'est jamais insérée dans le planning par le
         solveur (sauf shift fixe) ; elle ne compte jamais dans le
         Nb_personnes requis sur cette tâche, même via un shift fixe : un
         créneau demandant par exemple 3 personnes a besoin de 3 personnes
         du pool normal, en plus de la ou des personnes auto-gérées.

  blocages.xlsx
      Un onglet par jour (comme plages_horaires.xlsx). Chaque onglet
      contient une ligne par personne, avec ses colonnes fixes (par
      position, les en-têtes "Début"/"Fin" étant dupliqués) :
      Nom | Prénom | Raison P1 | Début P1 | Fin P1 | Raison P2 | Début P2 | Fin P2 | Raison P3 | Début P3 | Fin P3
      -> "Raison P1"/"Raison P2"/"Raison P3" (ex : nom d'un concert) sont
         informatifs. Il y a 3 niveaux de priorité, décroissants : le
         solveur essaie de respecter la priorité 1 (colonnes C-E) avant la
         priorité 2 (colonnes F-H), elle-même avant la priorité 3 (colonnes
         I-K), mais peut enfreindre n'importe laquelle si c'est nécessaire
         pour couvrir tous les besoins. Chaque case peut être vide (aucune
         contrainte sur cette priorité ce jour-là). Un jour sans aucun
         blocage peut ne pas avoir d'onglet.

  shifts_fixes.xlsx (optionnel)
      Un onglet par jour (comme plages_horaires.xlsx). Chaque onglet
      contient :
      Nom | Prénom | Tâche | Début | Fin
      -> une ligne par shift déjà imposé à certaines personnes ce jour-là.
         Nom + Prénom doivent correspondre exactement à une ligne de
         personnes.xlsx. Ce créneau vient s'ajouter à ceux de
         plages_horaires.xlsx : il n'a pas besoin de correspondre à un
         créneau existant (Tâche/Début/Fin libres). Ce shift est garanti
         dans le planning, même s'il dépasse les capacités ou un blocage de
         la personne. Un jour sans aucun shift fixe peut ne pas avoir
         d'onglet.

Règles appliquées :
  - une personne n'est jamais affectée à une tâche pour laquelle elle n'a
    pas la capacité requise (sauf shift fixe) ;
  - une personne fait au maximum 2 shifts dans la journée (1 pour un
    Junior.e, 0 pour un COF, selon la Fonction), et ne peut pas être
    affectée à deux créneaux qui se chevauchent ou qui laissent moins de
    2h de pause entre eux ;
  - un Junior.e n'est jamais affecté à un créneau qui se termine après
    23:00 (sauf shift fixe) ;
  - une personne avec une seule capacité s'auto-gère et n'est jamais
    affectée par le solveur (sauf shift fixe), et ne compte jamais dans le
    Nb_personnes requis d'une tâche (même via un shift fixe) ;
  - les shifts fixes sont toujours attribués et comptent dans le quota de
    shifts par jour ;
  - sur la semaine, chaque personne (sauf Junior.e, COF et auto-gérée)
    doit passer au moins une fois par un créneau "Déambule" ou par un
    créneau tardif (qui commence à 22:00 ou après) ; le solveur regarde ce
    qui a déjà été généré pour les autres jours (output/planning_<Jour>.xlsx)
    pour répartir ce passage sur la semaine plutôt que de le forcer chaque
    jour ;
  - si le solveur ne peut pas tout résoudre, l'ordre de priorité est :
    1) couvrir tous les besoins en personnel (Nb_personnes) ; 2) respecter
    au maximum les blocages de priorité 1, puis 2, puis 3 (en violant le
    moins de créneaux bloqués possible, plutôt que d'ignorer tout le
    blocage) ; 3) faire passer chacun par un créneau Déambule/tardif au
    moins une fois sur la semaine ; 4) répartir la charge équitablement
    entre les personnes ce jour-là.

Utilisation :
    pip install -r requirements.txt
    python generate_planning.py
    -> le script demande quel jour planifier, puis écrit le résultat dans
       output/planning_<jour>.xlsx
"""

import datetime as dt
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from ortools.sat.python import cp_model
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import sys as _sys
if getattr(_sys, "frozen", False):
    # Exécutable PyInstaller : les données sont à côté du .exe
    BASE_DIR = Path(_sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

# Couverture, puis blocages priorité 1/2/3, puis passage Déambule/tardif,
# puis équité : chacun est résolu dans sa propre phase (cascade lexico-
# graphique), du plus important au moins important. Voir generer_planning().

TARGET_SHIFTS_PER_DAY = 2
SOLVER_TIME_LIMIT_SECONDS = 30
PAUSE_MINIMALE_MINUTES = 120  # pause minimale entre deux shifts d'une même personne
HEURE_LIMITE_JUNIOR_MINUTES = 23 * 60  # un Junior.e ne termine jamais après 23:00
HEURE_DEBUT_TARDIF_MINUTES = 22 * 60  # un créneau est "tardif" s'il commence à 22:00 ou après

# Nombre maximum de shifts par jour selon la "Fonction" (personnes.xlsx).
# Une Fonction non reconnue utilise TARGET_SHIFTS_PER_DAY par défaut.
MAX_SHIFTS_PAR_FONCTION = {
    "junior.e": 1,
    "staff": 2,
    "adjoint.e": 2,
    "cof": 0,
}


# ---------------------------------------------------------------------------
# Utilitaires
# ---------------------------------------------------------------------------

def to_minutes(value) -> int:
    """Convertit une heure (str "HH:MM", datetime.time, fraction Excel...) en minutes."""
    if isinstance(value, dt.datetime):
        return value.hour * 60 + value.minute
    if isinstance(value, dt.time):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)):
        # Excel stocke parfois une heure comme fraction de journée (0.5 = midi)
        return round(value * 24 * 60) % (24 * 60)
    if isinstance(value, str):
        value = value.strip()
        parts = value.split(":")
        return int(parts[0]) * 60 + int(parts[1])
    raise ValueError(f"Format d'heure non reconnu : {value!r}")


def minutes_to_str(minutes: int) -> str:
    if minutes != 0 and minutes % (24 * 60) == 0:
        return "24:00"
    minutes = minutes % (24 * 60)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def overlaps(start1: int, end1: int, start2: int, end2: int) -> bool:
    return start1 < end2 and start2 < end1


def gap_insuffisant(s1, s2, pause_minimale: int) -> bool:
    """True si les deux créneaux se chevauchent, ou s'ils laissent moins de
    pause_minimale minutes de battement entre la fin de l'un et le début de l'autre."""
    if overlaps(s1["debut"], s1["fin"], s2["debut"], s2["fin"]):
        return True
    gap = s2["debut"] - s1["fin"] if s1["fin"] <= s2["debut"] else s1["debut"] - s2["fin"]
    return gap < pause_minimale


def parse_capacites(value):
    """Renvoie un set de tâches (en minuscules), None si la personne peut tout
    faire, ou un tuple ("exclude", set) pour la syntaxe "* - tâche1, tâche2"
    (toutes les tâches sauf celles listées). La résolution du tuple en set
    final se fait dans load_data, une fois la liste complète des tâches connue."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    value = str(value).strip()
    if value == "" or value.lower() in ("tous", "toutes", "*", "all"):
        return None
    if value.startswith("*"):
        # "* - tâche1, tâche2" → toutes les tâches sauf celles-ci
        reste = value[1:].strip()
        if reste.startswith("-"):
            exclusions = {p.strip().lower() for p in reste[1:].split(",") if p.strip()}
            return ("exclude", exclusions)
        return None
    return {part.strip().lower() for part in value.split(",") if part.strip()}


def person_key(nom, prenom) -> tuple:
    """Clé d'identification d'une personne : plusieurs personnes peuvent
    partager le même Nom de famille, on identifie donc par Nom + Prénom."""
    return (str(nom).strip().lower(), str(prenom).strip().lower())


# ---------------------------------------------------------------------------
# Sélection du jour
# ---------------------------------------------------------------------------

JOURS_VALIDES = {"lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"}


def lister_jours_disponibles():
    onglets = pd.read_excel(DATA_DIR / "plages_horaires.xlsx", sheet_name=None).keys()
    return [o for o in onglets if o.strip().lower() in JOURS_VALIDES]


def demander_jour(jours_disponibles):
    print("Jours disponibles dans plages_horaires.xlsx :")
    for i, jour in enumerate(jours_disponibles, start=1):
        print(f"  {i}. {jour}")
    while True:
        choix = input("Quel jour faut-il planifier ? ").strip()
        for jour in jours_disponibles:
            if choix.lower() == jour.lower():
                return jour
        if choix.isdigit() and 1 <= int(choix) <= len(jours_disponibles):
            return jours_disponibles[int(choix) - 1]
        print("Jour non reconnu, merci de réessayer.")


# ---------------------------------------------------------------------------
# Rotation Déambule / créneaux tardifs sur la semaine
# ---------------------------------------------------------------------------

def est_deambule_ou_tardif(tache_norm: str, debut_minutes: int) -> bool:
    return "déambule" in tache_norm or debut_minutes >= HEURE_DEBUT_TARDIF_MINUTES


def construire_lookup_affichage(people):
    """Associe chaque nom_affichage (tel qu'écrit dans les plannings déjà
    générés) à la clé de la personne correspondante. Un nom_affichage
    partagé par deux personnes (cas non prévu) est écarté : on ne peut pas
    deviner sans ambiguïté de qui il s'agit."""
    lookup = {}
    ambigus = set()
    for p in people:
        cle_affichage = p["nom_affichage"].strip().lower()
        if cle_affichage in lookup and lookup[cle_affichage] != p["key"]:
            ambigus.add(cle_affichage)
        else:
            lookup[cle_affichage] = p["key"]
    for cle in ambigus:
        lookup.pop(cle, None)
    return lookup


def charger_deja_rotes(jour_actuel, lookup_affichage):
    """Personnes ayant déjà eu un créneau Déambule ou tardif un autre jour,
    d'après les fichiers output/planning_<Jour>.xlsx déjà générés."""
    deja = set()
    for autre_jour in JOURS_VALIDES:
        if autre_jour == jour_actuel.lower():
            continue
        nom_jour = autre_jour.capitalize()
        chemin = OUTPUT_DIR / f"planning_{nom_jour}.xlsx"
        if not chemin.exists():
            continue
        wb = load_workbook(chemin, data_only=True)
        if nom_jour not in wb.sheetnames:
            continue
        for row in wb[nom_jour].iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None or str(row[0]).startswith("("):
                continue
            tache, debut, _fin, _requis, _affectes, _manquant, personnes = row[:7]
            if not personnes:
                continue
            if not est_deambule_ou_tardif(str(tache).strip().lower(), to_minutes(debut)):
                continue
            for nom in str(personnes).split(","):
                nom_normalise = nom.strip()
                for marqueur in (" (*)", " (**)", " (***)"):
                    nom_normalise = nom_normalise.replace(marqueur, "")
                cle = lookup_affichage.get(nom_normalise.strip().lower())
                if cle:
                    deja.add(cle)
    return deja


# ---------------------------------------------------------------------------
# Chargement des données
# ---------------------------------------------------------------------------

def load_data(jour):
    plages_df = pd.read_excel(DATA_DIR / "plages_horaires.xlsx", sheet_name=jour).dropna(
        subset=["Tâche", "Début", "Fin", "Nb_personnes"]
    )
    personnes_df = pd.read_excel(DATA_DIR / "personnes.xlsx").dropna(subset=["Nom", "Prénom"])

    # blocages.xlsx : une ligne par personne, colonnes fixes par position
    # (A=Nom, B=Prénom, C=Raison P1, D=Début P1, E=Fin P1, F=Raison P2,
    # G=Début P2, H=Fin P2). Les en-têtes "Début"/"Fin" sont dupliqués donc
    # on lit par position de colonne plutôt que par nom.
    blocages_path = DATA_DIR / "blocages.xlsx"
    blocages_sheet = next(
        (s for s in pd.ExcelFile(blocages_path).sheet_names if s.lower() == jour.lower()), None
    )
    if blocages_sheet is not None:
        blocages_df = pd.read_excel(blocages_path, sheet_name=blocages_sheet, header=0)
    else:
        blocages_df = pd.DataFrame()

    shifts = []
    for _, row in plages_df.iterrows():
        debut = to_minutes(row["Début"])
        fin = to_minutes(row["Fin"])
        if fin <= debut:
            fin += 24 * 60
        shifts.append({
            "tache": str(row["Tâche"]).strip(),
            "tache_norm": str(row["Tâche"]).strip().lower(),
            "origine": "plages",
            "debut": debut,
            "fin": fin,
            "requis": int(row["Nb_personnes"]),
        })

    people = []
    for _, row in personnes_df.iterrows():
        nom = str(row["Nom"]).strip()
        prenom = str(row["Prénom"]).strip()
        fonction = str(row.get("Fonction", "")).strip()
        nb_shifts_col = row.get("Nombre de shifts")
        if nb_shifts_col is not None and not (isinstance(nb_shifts_col, float) and pd.isna(nb_shifts_col)):
            max_shifts_p = int(nb_shifts_col)
        else:
            max_shifts_p = MAX_SHIFTS_PAR_FONCTION.get(fonction.lower(), TARGET_SHIFTS_PER_DAY)
        raw_caps = parse_capacites(row.get("Capacités"))
        # capacites_explicites : uniquement les tâches nommées directement (pas
        # via wildcard "* - X"). Seules ces personnes peuvent être assignées
        # aux tâches spécifiques (colonne B de l'onglet Tâches).
        if raw_caps is None or isinstance(raw_caps, tuple):
            capacites_explicites = None
        else:
            capacites_explicites = raw_caps
        # Disponibilité pour ce jour : col F-K (Mardi-Dimanche). 0 = absent.
        dispo_val = row.get(jour)
        disponible = not (
            dispo_val is not None
            and not (isinstance(dispo_val, float) and pd.isna(dispo_val))
            and str(dispo_val).strip() == "0"
        )
        people.append({
            "nom": nom,
            "prenom": prenom,
            "key": person_key(nom, prenom),
            "fonction": fonction.lower(),
            "max_shifts": max_shifts_p,
            "capacites": raw_caps,           # résolu plus bas si tuple "exclude"
            "capacites_explicites": capacites_explicites,
            "auto_gere": raw_caps is not None and not isinstance(raw_caps, tuple) and len(raw_caps) == 1,
            "disponible": disponible,
        })

    # Nom d'affichage : le prénom seul pour gagner de la place ; complété de
    # la 1re lettre du nom de famille seulement si le prénom n'est pas unique.
    prenom_counts = Counter(p["prenom"].lower() for p in people)
    for p in people:
        if prenom_counts[p["prenom"].lower()] > 1 and p["nom"]:
            p["nom_affichage"] = f"{p['prenom']} {p['nom'][0]}."
        else:
            p["nom_affichage"] = p["prenom"]

    key_counts = Counter(p["key"] for p in people)
    doublons = sorted({
        f"{p['prenom']} {p['nom']}".strip() for p in people if key_counts[p["key"]] > 1
    })
    if doublons:
        raise ValueError(
            f"personnes.xlsx : personne(s) en double (même Nom + Prénom) : {', '.join(doublons)}"
        )
    key_to_idx = {p["key"]: idx for idx, p in enumerate(people)}

    blocages_map = defaultdict(list)
    for _, row in blocages_df.iterrows():
        nom = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        prenom = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if not nom and not prenom:
            continue
        key = person_key(nom, prenom)
        if key not in key_to_idx:
            raise ValueError(
                f"blocages.xlsx : personne inconnue '{prenom} {nom}' (absente de personnes.xlsx)"
            )
        # (priorité, colonne raison, colonne début, colonne fin)
        for priorite, col_raison, col_debut, col_fin in ((1, 2, 3, 4), (2, 5, 6, 7), (3, 8, 9, 10)):
            if pd.isna(row.iloc[col_debut]) or pd.isna(row.iloc[col_fin]):
                continue
            debut = to_minutes(row.iloc[col_debut])
            fin = to_minutes(row.iloc[col_fin])
            if fin <= debut:
                fin += 24 * 60
            raison = str(row.iloc[col_raison]).strip() if pd.notna(row.iloc[col_raison]) else ""
            blocages_map[key].append({
                "debut": debut,
                "fin": fin,
                "priorite": priorite,
                "raison": raison,
            })

    # Shifts fixes (optionnel) : créneaux déjà imposés ce jour-là, ajoutés en
    # plus de ceux de plages_horaires.xlsx (pas besoin d'y correspondre).
    fixed_assignments = []
    fixed_path = DATA_DIR / "shifts_fixes.xlsx"
    if fixed_path.exists():
        fixed_sheet = next(
            (s for s in pd.ExcelFile(fixed_path).sheet_names if s.lower() == jour.lower()), None
        )
        if fixed_sheet is not None:
            fixed_df = pd.read_excel(fixed_path, sheet_name=fixed_sheet).dropna(
                subset=["Nom", "Prénom", "Tâche", "Début", "Fin"]
            )
        else:
            fixed_df = pd.DataFrame(columns=["Nom", "Prénom", "Tâche", "Début", "Fin"])
        for _, row in fixed_df.iterrows():
            nom = str(row["Nom"]).strip()
            prenom = str(row["Prénom"]).strip()
            key = person_key(nom, prenom)
            if key not in key_to_idx:
                raise ValueError(
                    f"shifts_fixes.xlsx : personne inconnue '{prenom} {nom}' (absente de personnes.xlsx)"
                )
            tache = str(row["Tâche"]).strip()
            debut = to_minutes(row["Début"])
            fin = to_minutes(row["Fin"])
            if fin <= debut:
                fin += 24 * 60
            s_idx = len(shifts)
            shifts.append({
                "tache": tache,
                "tache_norm": tache.lower(),
                "origine": "fixe",
                "debut": debut,
                "fin": fin,
                "requis": 1,
            })
            fixed_assignments.append((key, s_idx))

    # Résolution des capacités "* - exclusions" : maintenant que toutes les
    # tâches (plages + fixes) sont connues, on peut construire le set final.
    all_taches_norm = {s["tache_norm"] for s in shifts}
    for p in people:
        if isinstance(p["capacites"], tuple):
            _, exclusions = p["capacites"]
            resolved = all_taches_norm - exclusions
            p["capacites"] = resolved if resolved else None
            p["auto_gere"] = p["capacites"] is not None and len(p["capacites"]) == 1

    # Tâches spécifiques (colonne B de l'onglet Tâches dans plages_horaires.xlsx) :
    # seules les personnes qui ont explicitement déclaré cette tâche dans leurs
    # capacités peuvent s'y voir assignées (pas le pool général).
    taches_specifiques = set()
    xl_sheets = pd.ExcelFile(DATA_DIR / "plages_horaires.xlsx").sheet_names
    if "Tâches" in xl_sheets:
        taches_df = pd.read_excel(DATA_DIR / "plages_horaires.xlsx", sheet_name="Tâches", header=0)
        if taches_df.shape[1] > 1:
            col_b = taches_df.iloc[:, 1]
            taches_specifiques = {str(v).strip().lower() for v in col_b if pd.notna(v) and str(v).strip()}

    deja_rotes = charger_deja_rotes(jour, construire_lookup_affichage(people))

    return shifts, people, blocages_map, fixed_assignments, deja_rotes, taches_specifiques


# ---------------------------------------------------------------------------
# Construction et résolution du modèle
# ---------------------------------------------------------------------------

def build_model(shifts, people, blocages_map, fixed_assignments, deja_rotes, taches_specifiques=None):
    model = cp_model.CpModel()
    taches_specifiques = taches_specifiques or set()

    key_to_idx = {p["key"]: idx for idx, p in enumerate(people)}
    fixed_pairs = {(key_to_idx[key], s_idx) for key, s_idx in fixed_assignments}

    x = {}          # (p_idx, s_idx) -> BoolVar, uniquement pour les paires possibles
    p1_pairs = set()  # paires en conflit avec un blocage de priorité 1
    p2_pairs = set()  # paires en conflit avec une préférence de priorité 2
    p3_pairs = set()  # paires en conflit avec une préférence de priorité 3

    for s_idx, s in enumerate(shifts):
        for p_idx, p in enumerate(people):
            is_fixed = (p_idx, s_idx) in fixed_pairs

            if not is_fixed:
                if s["tache_norm"] in taches_specifiques:
                    # Tâche spécifique : uniquement les personnes qui l'ont
                    # déclarée explicitement dans leurs capacités (pas le pool
                    # général "Tous" ni les wildcards "* - X").
                    caps_expl = p["capacites_explicites"]
                    if caps_expl is None or s["tache_norm"] not in caps_expl:
                        continue
                else:
                    # Tâche commune : vérification habituelle par capacités.
                    caps = p["capacites"]
                    if caps is not None and s["tache_norm"] not in caps:
                        continue

            # Personne absente ce jour-là (colonne F-K dans personnes.xlsx = 0).
            if not is_fixed and not p["disponible"]:
                continue

            # Une seule capacité déclarée : la personne s'auto-gère sur cette
            # tâche et n'est jamais insérée dans le planning par le solveur.
            if not is_fixed and p["auto_gere"]:
                continue

            if not is_fixed and p["fonction"] == "junior.e" and s["fin"] > HEURE_LIMITE_JUNIOR_MINUTES:
                continue

            x[(p_idx, s_idx)] = model.new_bool_var(f"x_{p_idx}_{s_idx}")

            blocks = blocages_map.get(p["key"], [])
            if any(
                b["priorite"] == 1 and overlaps(s["debut"], s["fin"], b["debut"], b["fin"])
                for b in blocks
            ):
                p1_pairs.add((p_idx, s_idx))

            if any(
                b["priorite"] == 2 and overlaps(s["debut"], s["fin"], b["debut"], b["fin"])
                for b in blocks
            ):
                p2_pairs.add((p_idx, s_idx))

            if any(
                b["priorite"] == 3 and overlaps(s["debut"], s["fin"], b["debut"], b["fin"])
                for b in blocks
            ):
                p3_pairs.add((p_idx, s_idx))

    # Les shifts fixes sont garantis
    for pair in fixed_pairs:
        model.add(x[pair] == 1)

    # Couverture des besoins : assignés + manquants == requis. Sur les
    # créneaux de plages_horaires.xlsx, les personnes auto-gérées (une seule
    # capacité) ne comptent jamais dans ce besoin, même affectées via un
    # shift fixe : le besoin doit être couvert par le pool normal, en plus
    # d'elles. Sur un créneau créé par shifts_fixes.xlsx (requis=1, réservé à
    # la personne fixée), on compte tout le monde normalement.
    shortfall = {}
    for s_idx, s in enumerate(shifts):
        shortfall[s_idx] = model.new_int_var(0, s["requis"], f"shortfall_{s_idx}")
        assigned = [
            x[(p_idx, s_idx)] for p_idx, p in enumerate(people)
            if (p_idx, s_idx) in x and not (s["origine"] == "plages" and p["auto_gere"])
        ]
        model.add(sum(assigned) + shortfall[s_idx] == s["requis"])

    # Chevauchements/pause insuffisante et plafond de shifts (selon la Fonction)
    for i in range(len(shifts)):
        for j in range(i + 1, len(shifts)):
            if gap_insuffisant(shifts[i], shifts[j], PAUSE_MINIMALE_MINUTES):
                for p_idx in range(len(people)):
                    if (p_idx, i) in x and (p_idx, j) in x:
                        model.add(x[(p_idx, i)] + x[(p_idx, j)] <= 1)

    # Une personne ne fait pas deux fois la même tâche dans la journée.
    shifts_par_tache = defaultdict(list)
    for s_idx, s in enumerate(shifts):
        shifts_par_tache[s["tache_norm"]].append(s_idx)
    for p_idx in range(len(people)):
        for s_idxs in shifts_par_tache.values():
            vars_tache = [x[(p_idx, s_idx)] for s_idx in s_idxs if (p_idx, s_idx) in x]
            if len(vars_tache) > 1:
                model.add(sum(vars_tache) <= 1)

    for p_idx, p in enumerate(people):
        today = [x[(p_idx, s_idx)] for s_idx in range(len(shifts)) if (p_idx, s_idx) in x]
        if today:
            model.add(sum(today) <= p["max_shifts"])

    # Équité : minimiser l'écart entre la personne la plus et la moins sollicitée
    totals = []
    for p_idx in range(len(people)):
        assigned = [x[(p_idx, s_idx)] for s_idx in range(len(shifts)) if (p_idx, s_idx) in x]
        total_p = model.new_int_var(0, len(shifts), f"total_{p_idx}")
        model.add(total_p == sum(assigned))
        totals.append(total_p)

    max_total = model.new_int_var(0, len(shifts), "max_total")
    min_total = model.new_int_var(0, len(shifts), "min_total")
    model.add_max_equality(max_total, totals)
    model.add_min_equality(min_total, totals)

    # Déambule à éviter pour les personnes limitées à 1 shift : elles ont peu
    # de créneaux disponibles et un Déambule est moins utile qu'une tâche fixe.
    deambule_un_shift_pairs = {
        (p_idx, s_idx) for (p_idx, s_idx) in x
        if people[p_idx]["max_shifts"] == 1 and "déambule" in shifts[s_idx]["tache_norm"]
    }

    # Rotation Déambule/tardif : pour chaque personne concernée qui n'est
    # pas déjà passée par un tel créneau un autre jour, un indicateur vaut 1
    # si elle n'en a aucun aujourd'hui non plus (à minimiser), 0 sinon. Les
    # Junior.e, COF et personnes auto-gérées ne sont jamais concernés.
    deambule_tardif = {
        s_idx for s_idx, s in enumerate(shifts)
        if est_deambule_ou_tardif(s["tache_norm"], s["debut"])
    }
    manque_rotation = []  # liste de (p_idx, indicateur BoolVar)
    for p_idx, p in enumerate(people):
        if p["key"] in deja_rotes or p["fonction"] in ("junior.e", "cof") or p["auto_gere"]:
            continue
        creneaux_rotation = [x[(p_idx, s_idx)] for s_idx in deambule_tardif if (p_idx, s_idx) in x]
        if not creneaux_rotation:
            continue
        indicateur = model.new_bool_var(f"manque_rotation_{p_idx}")
        model.add(sum(creneaux_rotation) + indicateur >= 1)
        manque_rotation.append((p_idx, indicateur))

    return model, x, shortfall, p1_pairs, p2_pairs, p3_pairs, max_total, min_total, manque_rotation, deambule_un_shift_pairs


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SHORTAGE_FILL = PatternFill("solid", fgColor="F4B084")
TITLE_FONT = Font(bold=True, size=12)


def style_header_row(ws, row=1):
    for cell in ws[row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")


def trouver_raison(blocages_map, key, s, priorite):
    for b in blocages_map.get(key, []):
        if b["priorite"] == priorite and overlaps(s["debut"], s["fin"], b["debut"], b["fin"]):
            return b.get("raison", "")
    return ""


def export_results(jour, shifts, people, x, shortfall, p1_pairs, p2_pairs, p3_pairs, blocages_map, manque_rotation, solver, output_path):
    wb = Workbook()
    wb.remove(wb.active)

    # --- Feuille du jour : un tableau par tâche/créneau ----------------------
    headers = ["Tâche", "Début", "Fin", "Requis", "Affectés", "Manquant", "Personnes affectées"]
    ws = wb.create_sheet(title=str(jour)[:31])
    ws.append(headers)
    style_header_row(ws)

    for s_idx, s in enumerate(shifts):
        assigned_names = []
        for p_idx, p in enumerate(people):
            if (p_idx, s_idx) in x and solver.value(x[(p_idx, s_idx)]) == 1:
                name = p["nom_affichage"]
                if (p_idx, s_idx) in p1_pairs:
                    name += " (*)"
                if (p_idx, s_idx) in p2_pairs:
                    name += " (**)"
                if (p_idx, s_idx) in p3_pairs:
                    name += " (***)"
                assigned_names.append(name)

        manquant = solver.value(shortfall[s_idx])
        row = [
            s["tache"],
            minutes_to_str(s["debut"]),
            minutes_to_str(s["fin"]),
            s["requis"],
            s["requis"] - manquant,
            manquant,
            ", ".join(assigned_names),
        ]
        ws.append(row)
        if manquant > 0:
            for col in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=col).fill = SHORTAGE_FILL

    ws.append([])
    note1 = ws.cell(row=ws.max_row + 1, column=1, value="(*) Personne affectée malgré un blocage de priorité 1")
    note1.font = Font(italic=True, size=9)
    note2 = ws.cell(row=ws.max_row + 1, column=1, value="(**) Personne affectée malgré un blocage de priorité 2")
    note2.font = Font(italic=True, size=9)
    note3 = ws.cell(row=ws.max_row + 1, column=1, value="(***) Personne affectée malgré un blocage de priorité 3")
    note3.font = Font(italic=True, size=9)

    for col, width in enumerate([20, 8, 8, 8, 10, 10, 70], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"

    # --- Récapitulatif par personne -----------------------------------------
    ws = wb.create_sheet("Récapitulatif", 0)
    ws.append(["Nom", "Shifts", "Tâches et horaires"])
    style_header_row(ws)

    ordre_affichage = sorted(range(len(people)), key=lambda p_idx: (people[p_idx]["prenom"].lower(), people[p_idx]["nom"].lower()))
    for p_idx in ordre_affichage:
        p = people[p_idx]
        assigned = sorted(
            (s_idx for s_idx in range(len(shifts)) if (p_idx, s_idx) in x and solver.value(x[(p_idx, s_idx)]) == 1),
            key=lambda s_idx: shifts[s_idx]["debut"],
        )
        detail_parts = []
        for s_idx in assigned:
            s = shifts[s_idx]
            part = f"{s['tache']} {minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}"
            if (p_idx, s_idx) in p1_pairs:
                part += " (*)"
            if (p_idx, s_idx) in p2_pairs:
                part += " (**)"
            if (p_idx, s_idx) in p3_pairs:
                part += " (***)"
            detail_parts.append(part)
        ws.append([p["nom_affichage"], len(assigned), "; ".join(detail_parts)])

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 60
    ws.freeze_panes = "A2"

    # --- Alertes : sous-effectifs et préférences non respectées -------------
    ws = wb.create_sheet("Alertes")
    ws.append(["Type", "Tâche", "Horaire", "Détail"])
    style_header_row(ws)

    for s_idx, s in enumerate(shifts):
        manquant = solver.value(shortfall[s_idx])
        if manquant > 0:
            ws.append([
                "Sous-effectif",
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                f"{manquant} personne(s) manquante(s) sur {s['requis']} requise(s)",
            ])

    for (p_idx, s_idx) in p1_pairs:
        if solver.value(x[(p_idx, s_idx)]) == 1:
            s = shifts[s_idx]
            p = people[p_idx]
            raison = trouver_raison(blocages_map, p["key"], s, priorite=1)
            detail = f"{p['nom_affichage']} avait un blocage (priorité 1) sur ce créneau"
            if raison:
                detail += f" ({raison})"
            ws.append([
                "Blocage priorité 1 non respecté",
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                detail,
            ])

    for (p_idx, s_idx) in p2_pairs:
        if solver.value(x[(p_idx, s_idx)]) == 1:
            s = shifts[s_idx]
            p = people[p_idx]
            raison = trouver_raison(blocages_map, p["key"], s, priorite=2)
            detail = f"{p['nom_affichage']} avait indiqué une préférence (priorité 2) sur ce créneau"
            if raison:
                detail += f" ({raison})"
            ws.append([
                "Préférence non respectée (priorité 2)",
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                detail,
            ])

    for (p_idx, s_idx) in p3_pairs:
        if solver.value(x[(p_idx, s_idx)]) == 1:
            s = shifts[s_idx]
            p = people[p_idx]
            raison = trouver_raison(blocages_map, p["key"], s, priorite=3)
            detail = f"{p['nom_affichage']} avait indiqué une préférence (priorité 3) sur ce créneau"
            if raison:
                detail += f" ({raison})"
            ws.append([
                "Préférence non respectée (priorité 3)",
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                detail,
            ])

    for (p_idx, indicateur) in manque_rotation:
        if solver.value(indicateur) == 1:
            p = people[p_idx]
            ws.append([
                "Rotation Déambule/tardif manquante",
                "-",
                "-",
                f"{p['nom_affichage']} n'a encore eu aucun créneau Déambule ou tardif cette semaine",
            ])

    for p_idx, p in enumerate(people):
        if p["auto_gere"] or p["max_shifts"] < 2:
            continue
        nb_shifts = sum(
            1 for s_idx in range(len(shifts))
            if (p_idx, s_idx) in x and solver.value(x[(p_idx, s_idx)]) == 1
        )
        if 0 < nb_shifts < p["max_shifts"]:
            ws.append([
                "Shifts incomplets",
                "-",
                "-",
                f"{p['nom_affichage']} n'a fait que {nb_shifts} shift(s) sur {p['max_shifts']} attendus",
            ])

    if ws.max_row == 1:
        ws.append(["Aucune", "-", "-", "Tous les besoins sont couverts et tous les blocages respectés"])

    for col, width in enumerate([24, 16, 14, 70], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"

    OUTPUT_DIR.mkdir(exist_ok=True)
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Orchestration (utilisée par la CLI et par l'application web)
# ---------------------------------------------------------------------------

def generer_planning(jour):
    """Charge les données, résout le modèle et exporte le planning pour un
    jour donné. Renvoie un dict (success, messages, output_path, totaux).
    Les erreurs de données (ValueError levée par load_data) ne sont pas
    interceptées : à l'appelant de les afficher comme il convient."""
    messages = []

    shifts, people, blocages_map, fixed_assignments, deja_rotes, taches_specifiques = load_data(jour)
    messages.append(
        f"{jour} : {len(shifts)} créneaux à couvrir, {len(people)} personnes dans l'équipe, "
        f"{len(fixed_assignments)} shift(s) fixe(s)."
    )

    model, x, shortfall, p1_pairs, p2_pairs, p3_pairs, max_total, min_total, manque_rotation, deambule_un_shift_pairs = build_model(
        shifts, people, blocages_map, fixed_assignments, deja_rotes, taches_specifiques
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT_SECONDS
    solver.parameters.num_search_workers = 8

    # Phase 1 : couvrir les besoins en personnel est la priorité absolue.
    total_shortfall_expr = sum(shortfall.values())
    model.minimize(total_shortfall_expr)
    status = solver.solve(model)

    messages.append(f"Statut du solveur : {solver.status_name(status)}")
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        messages.append(
            "Aucune solution trouvée. Vérifiez notamment les shifts fixes : "
            "deux shifts fixes qui se chevauchent, ou plus de shifts fixes que le "
            "quota autorisé par sa Fonction pour une même personne ce jour-là, "
            "rendent le planning impossible."
        )
        return {"success": False, "messages": messages, "output_path": None}

    shortfall_min = sum(solver.value(v) for v in shortfall.values())
    model.add(total_shortfall_expr <= shortfall_min)

    # Phase 2 : à couverture égale, respecter au maximum les blocages de priorité 1.
    p1_penalty = sum(x[pair] for pair in p1_pairs) if p1_pairs else 0
    model.minimize(p1_penalty)
    solver.solve(model)
    p1_violations_min = sum(solver.value(x[pair]) for pair in p1_pairs) if p1_pairs else 0
    model.add(p1_penalty <= p1_violations_min)

    # Phase 3 : à priorité 1 respectée au mieux, idem pour la priorité 2.
    p2_penalty = sum(x[pair] for pair in p2_pairs) if p2_pairs else 0
    model.minimize(p2_penalty)
    solver.solve(model)
    p2_violations_min = sum(solver.value(x[pair]) for pair in p2_pairs) if p2_pairs else 0
    model.add(p2_penalty <= p2_violations_min)

    # Phase 4 : à priorité 2 respectée au mieux, idem pour la priorité 3.
    p3_penalty = sum(x[pair] for pair in p3_pairs) if p3_pairs else 0
    model.minimize(p3_penalty)
    solver.solve(model)
    p3_violations_min = sum(solver.value(x[pair]) for pair in p3_pairs) if p3_pairs else 0
    model.add(p3_penalty <= p3_violations_min)

    # Phase 5 : éviter d'assigner un Déambule aux personnes limitées à 1 shift.
    d1_penalty = sum(x[pair] for pair in deambule_un_shift_pairs) if deambule_un_shift_pairs else 0
    model.minimize(d1_penalty)
    solver.solve(model)
    d1_min = sum(solver.value(x[pair]) for pair in deambule_un_shift_pairs) if deambule_un_shift_pairs else 0
    model.add(d1_penalty <= d1_min)

    # Phase 6 : à blocages respectés au mieux, faire passer chacun par un
    # créneau Déambule/tardif au moins une fois sur la semaine.
    rotation_penalty = sum(ind for _, ind in manque_rotation) if manque_rotation else 0
    model.minimize(rotation_penalty)
    solver.solve(model)
    rotation_min = sum(solver.value(ind) for _, ind in manque_rotation) if manque_rotation else 0
    model.add(rotation_penalty <= rotation_min)

    # Phase 7 : enfin, répartir la charge équitablement entre les personnes.
    model.minimize(max_total - min_total)
    solver.solve(model)

    total_shortfall = sum(solver.value(v) for v in shortfall.values())
    total_p1 = sum(solver.value(x[pair]) for pair in p1_pairs) if p1_pairs else 0
    total_p2 = sum(solver.value(x[pair]) for pair in p2_pairs) if p2_pairs else 0
    total_p3 = sum(solver.value(x[pair]) for pair in p3_pairs) if p3_pairs else 0
    total_manque_rotation = sum(solver.value(ind) for _, ind in manque_rotation) if manque_rotation else 0
    messages.append(f"Créneaux non couverts (somme) : {total_shortfall}")
    messages.append(f"Blocages priorité 1 non respectés : {total_p1}")
    messages.append(f"Préférences (priorité 2) non respectées : {total_p2}")
    messages.append(f"Préférences (priorité 3) non respectées : {total_p3}")
    messages.append(f"Personnes sans créneau Déambule/tardif cette semaine : {total_manque_rotation}")

    output_path = OUTPUT_DIR / f"planning_{jour}.xlsx"
    export_results(
        jour, shifts, people, x, shortfall, p1_pairs, p2_pairs, p3_pairs,
        blocages_map, manque_rotation, solver, output_path,
    )
    messages.append(f"Planning généré : {output_path}")

    return {
        "success": True,
        "messages": messages,
        "output_path": output_path,
        "total_shortfall": total_shortfall,
        "total_p1": total_p1,
        "total_p2": total_p2,
        "total_p3": total_p3,
    }


# ---------------------------------------------------------------------------
# Programme principal (CLI)
# ---------------------------------------------------------------------------

def main():
    jour = demander_jour(lister_jours_disponibles())
    resultat = generer_planning(jour)
    for message in resultat["messages"]:
        print(message)


if __name__ == "__main__":
    main()
