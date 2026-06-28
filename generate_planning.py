
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
         liste les tâches que la personne peut effectuer, séparées par des
         virgules (ex: "Bar, Accueil"). Laisser vide ou écrire "Tous" si la
         personne peut effectuer n'importe quelle tâche. Une personne avec
         une seule capacité (ex: "Bourdon.ne") s'auto-gère sur cette tâche
         et n'est jamais insérée dans le planning par le solveur (sauf
         shift fixe).

  blocages.xlsx
      Un onglet par jour (comme plages_horaires.xlsx). Chaque onglet
      contient une ligne par personne, avec ses colonnes fixes (par
      position, les en-têtes "Début"/"Fin" étant dupliqués) :
      Nom | Prénom | Raison P1 | Début P1 | Fin P1 | Raison P2 | Début P2 | Fin P2
      -> "Raison P1"/"Raison P2" (ex : nom d'un concert) sont informatifs.
         Priorité 1 (colonnes C-E) : la personne souhaite être libre sur ce
         créneau ; le solveur essaie de le respecter en priorité, mais peut
         l'enfreindre si c'est nécessaire pour couvrir tous les besoins.
         Priorité 2 (colonnes F-H) : préférence secondaire, encore moins
         contraignante. Chaque case peut être vide (aucune contrainte sur
         cette priorité ce jour-là). Un jour sans aucun blocage peut ne pas
         avoir d'onglet.

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
    affectée par le solveur (sauf shift fixe) ;
  - les shifts fixes sont toujours attribués et comptent dans le quota de
    shifts par jour ;
  - si le solveur ne peut pas tout résoudre, l'ordre de priorité est :
    1) couvrir tous les besoins en personnel (Nb_personnes) ; 2) respecter
    au maximum les blocages de priorité 1 (en violant le moins de
    créneaux bloqués possible, plutôt que d'ignorer tout le blocage) ;
    3) respecter au maximum les préférences de priorité 2 et répartir la
    charge équitablement entre les personnes ce jour-là.

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
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

# Couvrir les besoins en personnel (shortfall) est résolu en priorité absolue
# (1re phase), avant d'optimiser les objectifs secondaires suivants (2e phase),
# du plus important au moins important :
W_PRIORITY2 = 50     # éviter d'affecter quelqu'un sur une préférence (priorité 2)
W_FAIRNESS = 1       # répartir la charge de travail équitablement

TARGET_SHIFTS_PER_DAY = 2
SOLVER_TIME_LIMIT_SECONDS = 30
PAUSE_MINIMALE_MINUTES = 120  # pause minimale entre deux shifts d'une même personne
HEURE_LIMITE_JUNIOR_MINUTES = 23 * 60  # un Junior.e ne termine jamais après 23:00

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
    """Renvoie un set de tâches (en minuscules) ou None si la personne peut tout faire."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    value = str(value).strip()
    if value == "" or value.lower() in ("tous", "toutes", "*", "all"):
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
            "debut": debut,
            "fin": fin,
            "requis": int(row["Nb_personnes"]),
        })

    people = []
    for _, row in personnes_df.iterrows():
        nom = str(row["Nom"]).strip()
        prenom = str(row["Prénom"]).strip()
        fonction = str(row.get("Fonction", "")).strip()
        people.append({
            "nom": nom,
            "prenom": prenom,
            "nom_complet": f"{prenom} {nom}".strip(),
            "key": person_key(nom, prenom),
            "fonction": fonction.lower(),
            "max_shifts": MAX_SHIFTS_PAR_FONCTION.get(fonction.lower(), TARGET_SHIFTS_PER_DAY),
            "capacites": parse_capacites(row.get("Capacités")),
        })

    key_counts = Counter(p["key"] for p in people)
    doublons = sorted({p["nom_complet"] for p in people if key_counts[p["key"]] > 1})
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
        for priorite, col_raison, col_debut, col_fin in ((1, 2, 3, 4), (2, 5, 6, 7)):
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
                "debut": debut,
                "fin": fin,
                "requis": 1,
            })
            fixed_assignments.append((key, s_idx))

    return shifts, people, blocages_map, fixed_assignments


# ---------------------------------------------------------------------------
# Construction et résolution du modèle
# ---------------------------------------------------------------------------

def build_model(shifts, people, blocages_map, fixed_assignments):
    model = cp_model.CpModel()

    key_to_idx = {p["key"]: idx for idx, p in enumerate(people)}
    fixed_pairs = {(key_to_idx[key], s_idx) for key, s_idx in fixed_assignments}

    x = {}          # (p_idx, s_idx) -> BoolVar, uniquement pour les paires possibles
    p1_pairs = set()  # paires en conflit avec un blocage de priorité 1
    p2_pairs = set()  # paires en conflit avec une préférence (priorité 2)

    for s_idx, s in enumerate(shifts):
        for p_idx, p in enumerate(people):
            is_fixed = (p_idx, s_idx) in fixed_pairs

            caps = p["capacites"]
            if not is_fixed and caps is not None and s["tache_norm"] not in caps:
                continue

            # Une seule capacité déclarée : la personne s'auto-gère sur cette
            # tâche et n'est jamais insérée dans le planning par le solveur.
            if not is_fixed and caps is not None and len(caps) == 1:
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

    # Les shifts fixes sont garantis
    for pair in fixed_pairs:
        model.add(x[pair] == 1)

    # Couverture des besoins : assignés + manquants == requis
    shortfall = {}
    for s_idx, s in enumerate(shifts):
        shortfall[s_idx] = model.new_int_var(0, s["requis"], f"shortfall_{s_idx}")
        assigned = [x[(p_idx, s_idx)] for p_idx in range(len(people)) if (p_idx, s_idx) in x]
        model.add(sum(assigned) + shortfall[s_idx] == s["requis"])

    # Chevauchements/pause insuffisante et plafond de shifts (selon la Fonction)
    for i in range(len(shifts)):
        for j in range(i + 1, len(shifts)):
            if gap_insuffisant(shifts[i], shifts[j], PAUSE_MINIMALE_MINUTES):
                for p_idx in range(len(people)):
                    if (p_idx, i) in x and (p_idx, j) in x:
                        model.add(x[(p_idx, i)] + x[(p_idx, j)] <= 1)

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

    return model, x, shortfall, p1_pairs, p2_pairs, max_total, min_total


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


def export_results(jour, shifts, people, x, shortfall, p1_pairs, p2_pairs, blocages_map, solver, output_path):
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
                name = p["nom_complet"]
                if (p_idx, s_idx) in p1_pairs:
                    name += " (**)"
                if (p_idx, s_idx) in p2_pairs:
                    name += " (*)"
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
    note1 = ws.cell(row=ws.max_row + 1, column=1, value="(**) Personne affectée malgré un blocage de priorité 1")
    note1.font = Font(italic=True, size=9)
    note = ws.cell(row=ws.max_row + 1, column=1, value="(*) Personne affectée malgré une préférence de blocage (priorité 2)")
    note.font = Font(italic=True, size=9)

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
                part += " (**)"
            if (p_idx, s_idx) in p2_pairs:
                part += " (*)"
            detail_parts.append(part)
        ws.append([p["nom_complet"], len(assigned), "; ".join(detail_parts)])

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
            detail = f"{p['nom_complet']} avait un blocage (priorité 1) sur ce créneau"
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
            detail = f"{p['nom_complet']} avait indiqué une préférence (priorité 2) sur ce créneau"
            if raison:
                detail += f" ({raison})"
            ws.append([
                "Préférence non respectée",
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                detail,
            ])

    if ws.max_row == 1:
        ws.append(["Aucune", "-", "-", "Tous les besoins sont couverts et tous les blocages respectés"])

    for col, width in enumerate([24, 16, 14, 70], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"

    OUTPUT_DIR.mkdir(exist_ok=True)
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------

def main():
    jour = demander_jour(lister_jours_disponibles())

    shifts, people, blocages_map, fixed_assignments = load_data(jour)
    print(f"{jour} : {len(shifts)} créneaux à couvrir, {len(people)} personnes dans l'équipe, "
          f"{len(fixed_assignments)} shift(s) fixe(s).")

    model, x, shortfall, p1_pairs, p2_pairs, max_total, min_total = build_model(
        shifts, people, blocages_map, fixed_assignments
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT_SECONDS
    solver.parameters.num_search_workers = 8

    # Phase 1 : couvrir les besoins en personnel est la priorité absolue.
    total_shortfall_expr = sum(shortfall.values())
    model.minimize(total_shortfall_expr)
    status = solver.solve(model)

    print(f"Statut du solveur : {solver.status_name(status)}")
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("Aucune solution trouvée. Vérifiez notamment les shifts fixes : "
              "deux shifts fixes qui se chevauchent, ou plus de shifts fixes que le "
              "quota autorisé par sa Fonction pour une même personne ce jour-là, "
              "rendent le planning impossible.")
        return

    shortfall_min = sum(solver.value(v) for v in shortfall.values())

    # Phase 2 : à couverture égale, respecter au maximum les blocages priorité 1.
    p1_penalty = sum(x[pair] for pair in p1_pairs) if p1_pairs else 0
    model.add(total_shortfall_expr <= shortfall_min)
    model.minimize(p1_penalty)
    solver.solve(model)
    p1_violations_min = sum(solver.value(x[pair]) for pair in p1_pairs) if p1_pairs else 0

    # Phase 3 : à couverture et respect priorité 1 égaux, optimiser préférences et équité.
    p2_penalty = sum(x[pair] for pair in p2_pairs) if p2_pairs else 0
    model.add(p1_penalty <= p1_violations_min)
    model.minimize(W_PRIORITY2 * p2_penalty + W_FAIRNESS * (max_total - min_total))
    solver.solve(model)

    total_shortfall = sum(solver.value(v) for v in shortfall.values())
    total_p1 = sum(solver.value(x[pair]) for pair in p1_pairs) if p1_pairs else 0
    total_p2 = sum(solver.value(x[pair]) for pair in p2_pairs)
    print(f"Créneaux non couverts (somme) : {total_shortfall}")
    print(f"Blocages priorité 1 non respectés : {total_p1}")
    print(f"Préférences (priorité 2) non respectées : {total_p2}")

    output_path = OUTPUT_DIR / f"planning_{jour}.xlsx"
    export_results(jour, shifts, people, x, shortfall, p1_pairs, p2_pairs, blocages_map, solver, output_path)
    print(f"Planning généré : {output_path}")


if __name__ == "__main__":
    main()
