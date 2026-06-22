
"""
Génère le planning hebdomadaire d'une équipe (60-80 personnes, 2 shifts/jour)
à partir de 3 fichiers Excel saisis par le planificateur, et produit un
classeur Excel avec un tableau par jour (et par tâche) indiquant qui est
affecté à quoi.

Fichiers d'entrée attendus dans data/ :

  taches.xlsx
      Tâche
      -> liste de référence des tâches valides, dans l'ordre d'affichage
         souhaité pour le planning. Tout nom de tâche utilisé dans
         plages_horaires.xlsx ou dans les Capacités de personnes.xlsx doit
         y figurer exactement (à la casse près) ; une tâche inconnue
         provoque une erreur au chargement, et une tâche jamais planifiée
         déclenche un simple avertissement.

  plages_horaires.xlsx
      Jour | Tâche | Début | Fin | Nb_personnes
      -> une ligne par créneau à couvrir (une tâche, un horaire, un jour,
         et le nombre de personnes nécessaires sur ce créneau).

  personnes.xlsx
      Nom | Prénom | Capacités
      -> une ligne par membre de l'équipe. "Nom" et "Prénom" identifient la
         personne (plusieurs personnes peuvent partager le même nom de
         famille). "Capacités" liste les tâches que la personne peut
         effectuer, séparées par des virgules (ex: "Bar, Accueil"). Laisser
         vide ou écrire "Tous" si la personne peut effectuer n'importe
         quelle tâche.

  blocages.xlsx
      Nom | Prénom | Jour | Début | Fin | Priorité
      -> une ligne par plage horaire que la personne ne souhaite pas
         travailler. Nom + Prénom doivent correspondre exactement à une
         ligne de personnes.xlsx. Priorité 1 = blocage absolu (jamais
         affecté sur ce créneau). Priorité 2 = préférence (évité si
         possible, mais peut être utilisé pour compléter le planning).
         Chaque personne a droit à au maximum 2 blocages par jour.

  shifts_fixes.xlsx (optionnel)
      Nom | Prénom | Jour | Tâche | Début | Fin
      -> une ligne par shift déjà imposé pour toute la semaine à certaines
         personnes (typiquement 1 ou 2 par personne concernée). Nom +
         Prénom doivent correspondre exactement à une ligne de
         personnes.xlsx, et Jour/Tâche/Début/Fin doivent correspondre
         exactement à un créneau de plages_horaires.xlsx. Ce shift est
         garanti dans le planning, même s'il dépasse les capacités ou un
         blocage de la personne.

Règles appliquées :
  - une personne n'est jamais affectée à une tâche pour laquelle elle n'a
    pas la capacité requise (sauf shift fixe) ;
  - une personne n'est jamais affectée sur un créneau couvert par un
    blocage de priorité 1 (sauf shift fixe) ;
  - une personne fait au maximum 2 shifts par jour, et ne peut pas être
    affectée à deux créneaux qui se chevauchent le même jour ;
  - les shifts fixes sont toujours attribués et comptent dans le quota de
    2 shifts/jour ;
  - le solveur essaie de couvrir tous les besoins (Nb_personnes), d'éviter
    les créneaux en priorité 2, de tendre vers 2 shifts/jour par personne
    et de répartir la charge équitablement sur la semaine.

Utilisation :
    pip install -r requirements.txt
    python generate_sample_data.py   # (optionnel) génère des données d'exemple
    python generate_planning.py
    -> résultat dans output/planning_resultat.xlsx
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

# Poids de l'objectif (du plus important au moins important)
W_SHORTFALL = 1000   # couvrir les besoins en personnel
W_PRIORITY2 = 50     # éviter d'affecter quelqu'un sur une préférence (priorité 2)
W_FAIRNESS = 1       # répartir la charge de travail équitablement

TARGET_SHIFTS_PER_DAY = 2
SOLVER_TIME_LIMIT_SECONDS = 30


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
# Chargement des données
# ---------------------------------------------------------------------------

def load_data():
    taches_df = pd.read_excel(DATA_DIR / "taches.xlsx").dropna(subset=["Tâche"])
    taches_order = [str(t).strip() for t in taches_df["Tâche"]]
    task_order_index = {t.lower(): i for i, t in enumerate(taches_order)}

    plages_df = pd.read_excel(DATA_DIR / "plages_horaires.xlsx").dropna(
        subset=["Jour", "Tâche", "Début", "Fin", "Nb_personnes"]
    )
    personnes_df = pd.read_excel(DATA_DIR / "personnes.xlsx").dropna(subset=["Nom", "Prénom"])
    blocages_df = pd.read_excel(DATA_DIR / "blocages.xlsx").dropna(
        subset=["Nom", "Prénom", "Jour", "Début", "Fin", "Priorité"]
    )

    taches_inconnues = sorted({
        str(t).strip() for t in plages_df["Tâche"] if str(t).strip().lower() not in task_order_index
    })
    if taches_inconnues:
        raise ValueError(
            f"plages_horaires.xlsx : tâche(s) absente(s) de taches.xlsx : {', '.join(taches_inconnues)}"
        )

    shifts = []
    for _, row in plages_df.iterrows():
        debut = to_minutes(row["Début"])
        fin = to_minutes(row["Fin"])
        if fin <= debut:
            fin += 24 * 60
        shifts.append({
            "jour": str(row["Jour"]).strip(),
            "tache": str(row["Tâche"]).strip(),
            "tache_norm": str(row["Tâche"]).strip().lower(),
            "debut": debut,
            "fin": fin,
            "requis": int(row["Nb_personnes"]),
        })

    taches_non_planifiees = sorted(
        t for t in taches_order if t.lower() not in {s["tache_norm"] for s in shifts}
    )
    if taches_non_planifiees:
        print(f"Attention : tâche(s) de taches.xlsx jamais planifiée(s) dans plages_horaires.xlsx : "
              f"{', '.join(taches_non_planifiees)}")

    people = []
    for _, row in personnes_df.iterrows():
        nom = str(row["Nom"]).strip()
        prenom = str(row["Prénom"]).strip()
        people.append({
            "nom": nom,
            "prenom": prenom,
            "nom_complet": f"{prenom} {nom}".strip(),
            "key": person_key(nom, prenom),
            "capacites": parse_capacites(row.get("Capacités")),
        })

    key_counts = Counter(p["key"] for p in people)
    doublons = sorted({p["nom_complet"] for p in people if key_counts[p["key"]] > 1})
    if doublons:
        raise ValueError(
            f"personnes.xlsx : personne(s) en double (même Nom + Prénom) : {', '.join(doublons)}"
        )
    key_to_idx = {p["key"]: idx for idx, p in enumerate(people)}

    capacites_inconnues = sorted({
        cap for p in people if p["capacites"] is not None for cap in p["capacites"]
        if cap not in task_order_index
    })
    if capacites_inconnues:
        raise ValueError(
            f"personnes.xlsx : capacité(s) absente(s) de taches.xlsx : {', '.join(capacites_inconnues)}"
        )

    blocages_map = defaultdict(list)
    for _, row in blocages_df.iterrows():
        nom = str(row["Nom"]).strip()
        prenom = str(row["Prénom"]).strip()
        key = person_key(nom, prenom)
        if key not in key_to_idx:
            raise ValueError(
                f"blocages.xlsx : personne inconnue '{prenom} {nom}' (absente de personnes.xlsx)"
            )
        jour = str(row["Jour"]).strip()
        debut = to_minutes(row["Début"])
        fin = to_minutes(row["Fin"])
        if fin <= debut:
            fin += 24 * 60
        blocages_map[(key, jour)].append({
            "debut": debut,
            "fin": fin,
            "priorite": int(row["Priorité"]),
        })

    for (key, jour), blocks in blocages_map.items():
        if len(blocks) > 2:
            nom_complet = people[key_to_idx[key]]["nom_complet"]
            print(f"Attention : {nom_complet} a {len(blocks)} blocages déclarés le {jour} (maximum recommandé : 2).")

    # Shifts fixes (optionnel) : créneaux déjà imposés pour toute la semaine
    shift_lookup = {
        (s["jour"], s["tache_norm"], s["debut"], s["fin"]): s_idx
        for s_idx, s in enumerate(shifts)
    }

    fixed_assignments = []
    fixed_path = DATA_DIR / "shifts_fixes.xlsx"
    if fixed_path.exists():
        fixed_df = pd.read_excel(fixed_path).dropna(
            subset=["Nom", "Prénom", "Jour", "Tâche", "Début", "Fin"]
        )
        for _, row in fixed_df.iterrows():
            nom = str(row["Nom"]).strip()
            prenom = str(row["Prénom"]).strip()
            key = person_key(nom, prenom)
            if key not in key_to_idx:
                raise ValueError(
                    f"shifts_fixes.xlsx : personne inconnue '{prenom} {nom}' (absente de personnes.xlsx)"
                )
            jour = str(row["Jour"]).strip()
            tache_norm = str(row["Tâche"]).strip().lower()
            debut = to_minutes(row["Début"])
            fin = to_minutes(row["Fin"])
            if fin <= debut:
                fin += 24 * 60
            shift_key = (jour, tache_norm, debut, fin)
            if shift_key not in shift_lookup:
                raise ValueError(
                    f"shifts_fixes.xlsx : créneau introuvable dans plages_horaires.xlsx pour "
                    f"{prenom} {nom} ({jour}, {row['Tâche']}, {minutes_to_str(debut)}-{minutes_to_str(fin)})"
                )
            fixed_assignments.append((key, shift_lookup[shift_key]))

    return shifts, people, blocages_map, fixed_assignments, task_order_index


# ---------------------------------------------------------------------------
# Construction et résolution du modèle
# ---------------------------------------------------------------------------

def build_model(shifts, people, blocages_map, fixed_assignments):
    model = cp_model.CpModel()

    key_to_idx = {p["key"]: idx for idx, p in enumerate(people)}
    fixed_pairs = {(key_to_idx[key], s_idx) for key, s_idx in fixed_assignments}

    x = {}          # (p_idx, s_idx) -> BoolVar, uniquement pour les paires possibles
    p2_pairs = set()  # paires en conflit avec une préférence (priorité 2)

    for s_idx, s in enumerate(shifts):
        for p_idx, p in enumerate(people):
            is_fixed = (p_idx, s_idx) in fixed_pairs

            caps = p["capacites"]
            if not is_fixed and caps is not None and s["tache_norm"] not in caps:
                continue

            blocks = blocages_map.get((p["key"], s["jour"]), [])
            hard_blocked = any(
                b["priorite"] == 1 and overlaps(s["debut"], s["fin"], b["debut"], b["fin"])
                for b in blocks
            )
            if hard_blocked and not is_fixed:
                continue

            x[(p_idx, s_idx)] = model.new_bool_var(f"x_{p_idx}_{s_idx}")

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

    # Chevauchements et plafond de 2 shifts/jour
    shifts_by_day = defaultdict(list)
    for s_idx, s in enumerate(shifts):
        shifts_by_day[s["jour"]].append(s_idx)

    for jour, s_idxs in shifts_by_day.items():
        for i in range(len(s_idxs)):
            for j in range(i + 1, len(s_idxs)):
                s1, s2 = s_idxs[i], s_idxs[j]
                if overlaps(shifts[s1]["debut"], shifts[s1]["fin"], shifts[s2]["debut"], shifts[s2]["fin"]):
                    for p_idx in range(len(people)):
                        if (p_idx, s1) in x and (p_idx, s2) in x:
                            model.add(x[(p_idx, s1)] + x[(p_idx, s2)] <= 1)

        for p_idx in range(len(people)):
            today = [x[(p_idx, s_idx)] for s_idx in s_idxs if (p_idx, s_idx) in x]
            if today:
                model.add(sum(today) <= TARGET_SHIFTS_PER_DAY)

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

    p2_penalty = sum(x[pair] for pair in p2_pairs) if p2_pairs else 0
    model.minimize(
        W_SHORTFALL * sum(shortfall.values())
        + W_PRIORITY2 * p2_penalty
        + W_FAIRNESS * (max_total - min_total)
    )

    return model, x, shortfall, p2_pairs


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


def export_results(shifts, people, x, shortfall, p2_pairs, solver, output_path, task_order_index):
    wb = Workbook()
    wb.remove(wb.active)

    days = list(dict.fromkeys(s["jour"] for s in shifts))

    # --- Une feuille par jour, un tableau par tâche/créneau -----------------
    headers = ["Tâche", "Début", "Fin", "Requis", "Affectés", "Manquant", "Personnes affectées"]
    for jour in days:
        ws = wb.create_sheet(title=str(jour)[:31])
        ws.append(headers)
        style_header_row(ws)

        day_s_idxs = sorted(
            (s_idx for s_idx, s in enumerate(shifts) if s["jour"] == jour),
            key=lambda s_idx: (task_order_index[shifts[s_idx]["tache_norm"]], shifts[s_idx]["debut"]),
        )
        for s_idx in day_s_idxs:
            s = shifts[s_idx]

            assigned_names = []
            for p_idx, p in enumerate(people):
                if (p_idx, s_idx) in x and solver.value(x[(p_idx, s_idx)]) == 1:
                    name = p["nom_complet"]
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
        note = ws.cell(row=ws.max_row + 1, column=1, value="(*) Personne affectée malgré une préférence de blocage (priorité 2)")
        note.font = Font(italic=True, size=9)

        for col, width in enumerate([20, 8, 8, 8, 10, 10, 70], start=1):
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes = "A2"

    # --- Récapitulatif par personne -----------------------------------------
    ws = wb.create_sheet("Récapitulatif", 0)
    ws.append(["Nom"] + days + ["Total semaine"])
    style_header_row(ws)

    for p_idx, p in enumerate(people):
        counts = []
        total = 0
        for jour in days:
            count = sum(
                1
                for s_idx, s in enumerate(shifts)
                if s["jour"] == jour and (p_idx, s_idx) in x and solver.value(x[(p_idx, s_idx)]) == 1
            )
            counts.append(count)
            total += count
        ws.append([p["nom_complet"]] + counts + [total])

    ws.column_dimensions["A"].width = 22
    for col in range(2, len(days) + 3):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.freeze_panes = "A2"

    # --- Alertes : sous-effectifs et préférences non respectées -------------
    ws = wb.create_sheet("Alertes")
    ws.append(["Type", "Jour", "Tâche", "Horaire", "Détail"])
    style_header_row(ws)

    for s_idx, s in enumerate(shifts):
        manquant = solver.value(shortfall[s_idx])
        if manquant > 0:
            ws.append([
                "Sous-effectif",
                s["jour"],
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                f"{manquant} personne(s) manquante(s) sur {s['requis']} requise(s)",
            ])

    for (p_idx, s_idx) in p2_pairs:
        if solver.value(x[(p_idx, s_idx)]) == 1:
            s = shifts[s_idx]
            p = people[p_idx]
            ws.append([
                "Préférence non respectée",
                s["jour"],
                s["tache"],
                f"{minutes_to_str(s['debut'])}-{minutes_to_str(s['fin'])}",
                f"{p['nom_complet']} avait indiqué une préférence (priorité 2) sur ce créneau",
            ])

    if ws.max_row == 1:
        ws.append(["Aucune", "-", "-", "-", "Tous les besoins sont couverts et toutes les préférences respectées"])

    for col, width in enumerate([24, 12, 16, 14, 70], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"

    OUTPUT_DIR.mkdir(exist_ok=True)
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------

def main():
    shifts, people, blocages_map, fixed_assignments, task_order_index = load_data()
    print(f"{len(shifts)} créneaux à couvrir, {len(people)} personnes dans l'équipe, "
          f"{len(fixed_assignments)} shift(s) fixe(s).")

    model, x, shortfall, p2_pairs = build_model(shifts, people, blocages_map, fixed_assignments)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT_SECONDS
    solver.parameters.num_search_workers = 8
    status = solver.solve(model)

    print(f"Statut du solveur : {solver.status_name(status)}")
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("Aucune solution trouvée. Vérifiez notamment les shifts fixes : "
              "deux shifts fixes qui se chevauchent, ou plus de 2 shifts fixes "
              "le même jour pour une même personne, rendent le planning impossible.")
        return

    total_shortfall = sum(solver.value(v) for v in shortfall.values())
    total_p2 = sum(solver.value(x[pair]) for pair in p2_pairs)
    print(f"Créneaux non couverts (somme) : {total_shortfall}")
    print(f"Préférences (priorité 2) non respectées : {total_p2}")

    output_path = OUTPUT_DIR / "planning_resultat.xlsx"
    export_results(shifts, people, x, shortfall, p2_pairs, solver, output_path, task_order_index)
    print(f"Planning généré : {output_path}")


if __name__ == "__main__":
    main()
