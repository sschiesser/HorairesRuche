"""
Génère une grille horaire visuelle (un onglet par jour) à partir :
  - de data/horaires_paleo.xlsx     : les horaires des concerts, par scène
  - de output/planning_<Jour>.xlsx  : les plannings d'équipe déjà générés
    par generate_planning.py (un par jour)

Pour chaque jour disponible (présent à la fois dans horaires_paleo.xlsx et
sous forme de output/planning_<Jour>.xlsx déjà généré), produit un onglet
avec :
  - une ligne de bandeau horaire (de 15h à 03h, par pas de 5 minutes) ;
  - une ligne par scène, avec le nom de chaque concert positionné et
    dimensionné selon son horaire réel ;
  - une ligne par tâche du planning, avec les personnes affectées
    positionnées et dimensionnées selon l'horaire de chaque créneau.

Les jours dont le planning n'a pas encore été généré (output/planning_<Jour>.xlsx
manquant) sont ignorés, avec un avertissement.

Utilisation :
    python generate_grille.py
    -> écrit output/Grille_horaire.xlsx
"""

import datetime as dt
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from generate_planning import DATA_DIR, JOURS_VALIDES, OUTPUT_DIR, to_minutes

JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MOIS_FR = [
    "janvier", "février", "mars", "avril", "mai", "juin",
    "juillet", "août", "septembre", "octobre", "novembre", "décembre",
]

# La grille couvre de 15h à 03h le lendemain, par pas de 5 minutes. Toute
# heure antérieure à midi est considérée comme appartenant au petit matin
# du jour suivant (continuité de la grille après minuit).
PIVOT_NOCTURNE_MINUTES = 12 * 60
GRID_START_MINUTES = 15 * 60
GRID_END_MINUTES = (24 + 3) * 60
MINUTES_PAR_COLONNE = 5
COL_DEBUT_GRILLE = 3
COL_FIN_GRILLE = COL_DEBUT_GRILLE + (GRID_END_MINUTES - GRID_START_MINUTES) // MINUTES_PAR_COLONNE - 1

FOND_SOMBRE = PatternFill("solid", fgColor="0D1B2A")
FOND_ENTETE = PatternFill("solid", fgColor="1F2A38")
FOND_TIMELINE = PatternFill("solid", fgColor="2F3640")
FOND_CONCERT = PatternFill("solid", fgColor="FFD966")
FOND_PLANNING = PatternFill("solid", fgColor="FFE699")
TEXTE_BLANC = Font(color="FFFFFF", bold=True)
TEXTE_GRIS = Font(color="AAAAAA", size=8)
TEXTE_CONCERT = Font(color="7F6000", bold=True, size=8)
TEXTE_PLANNING = Font(color="7F5800", size=7)


def minutes_ajustees(valeur_brute) -> int:
    """Heure brute -> minutes depuis minuit, repoussée de 24h si elle
    appartient en réalité au petit matin du jour suivant."""
    minutes = to_minutes(valeur_brute)
    return minutes + 24 * 60 if minutes < PIVOT_NOCTURNE_MINUTES else minutes


def plage_minutes(debut_brut, fin_brut) -> tuple:
    debut = minutes_ajustees(debut_brut)
    fin = minutes_ajustees(fin_brut)
    if fin <= debut:
        fin += 24 * 60
    return debut, fin


def colonne_pour_minutes(minutes: int) -> int:
    return COL_DEBUT_GRILLE + (minutes - GRID_START_MINUTES) // MINUTES_PAR_COLONNE


def plage_colonnes(debut_minutes: int, fin_minutes: int):
    """Renvoie (col_debut, col_fin) bornées à la grille, ou None si la
    plage tombe entièrement hors de la grille."""
    col_debut = max(COL_DEBUT_GRILLE, colonne_pour_minutes(debut_minutes))
    col_fin = min(COL_FIN_GRILLE, colonne_pour_minutes(fin_minutes) - 1)
    if col_fin < col_debut:
        return None
    return col_debut, col_fin


# ---------------------------------------------------------------------------
# Chargement des horaires de concerts (data/horaires_paleo.xlsx)
# ---------------------------------------------------------------------------

def charger_horaires_concerts(path):
    """Lit le tableau empilé (une scène, puis une date, puis ses concerts,
    répété pour chaque scène) et renvoie (scenes, concerts) où :
      - scenes : liste ordonnée des noms de scène ;
      - concerts : dict (scène, date) -> liste de (artiste, début_min, fin_min).
    """
    df = pd.read_excel(path, sheet_name="Sheet1", header=None)

    scenes = []
    concerts = defaultdict(list)
    scene_courante = None
    date_courante = None

    for _, row in df.iterrows():
        col0, col1, col2 = row[0], row[1], row[2]

        if pd.isna(col0):
            scene_courante = None
            date_courante = None
            continue

        if isinstance(col0, (dt.datetime, pd.Timestamp)):
            date_courante = col0.date()
            continue

        if pd.isna(col1) and pd.isna(col2):
            if str(col0).strip().lower() != "band":
                scene_courante = str(col0).strip()
                if scene_courante not in scenes:
                    scenes.append(scene_courante)
            continue

        if scene_courante is None or date_courante is None:
            continue
        if pd.isna(col1) or pd.isna(col2):
            continue

        debut, fin = plage_minutes(col1, col2)
        concerts[(scene_courante, date_courante)].append((str(col0).strip(), debut, fin))

    return scenes, concerts


def construire_dates_par_jour(concerts):
    """Déduit, à partir des dates trouvées dans horaires_paleo.xlsx, la date
    calendaire de chaque jour de la semaine (Mardi -> 2026-07-21, etc.)."""
    dates = sorted({date for (_, date) in concerts.keys()})
    return {JOURS_FR[date.weekday()]: date for date in dates}


# ---------------------------------------------------------------------------
# Chargement des tâches déjà planifiées (output/planning_<Jour>.xlsx)
# ---------------------------------------------------------------------------

def charger_taches_planning(planning_path, jour):
    """Renvoie un dict tache -> liste de (personnes, début_min, fin_min),
    dans l'ordre d'apparition des tâches dans le planning."""
    wb = load_workbook(planning_path, data_only=True)
    ws = wb[jour]

    taches = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None or str(row[0]).startswith("("):
            continue
        tache, debut, fin, _requis, _affectes, _manquant, personnes = row[:7]
        taches.setdefault(str(tache).strip(), [])
        if not personnes:
            continue
        debut_min, fin_min = plage_minutes(debut, fin)
        taches[str(tache).strip()].append((str(personnes).strip(), debut_min, fin_min))

    return taches


# ---------------------------------------------------------------------------
# Construction de la grille
# ---------------------------------------------------------------------------

def ecrire_entete_heures(ws, ligne):
    ws.cell(row=ligne, column=1).fill = FOND_ENTETE
    minutes = GRID_START_MINUTES
    while minutes < GRID_END_MINUTES:
        col = colonne_pour_minutes(minutes)
        cell = ws.cell(row=ligne, column=col, value=f"{(minutes % (24 * 60)) // 60:02d}h")
        cell.font = TEXTE_GRIS
        minutes += 60
    for col in range(COL_DEBUT_GRILLE, COL_FIN_GRILLE + 1):
        ws.cell(row=ligne, column=col).fill = FOND_ENTETE
    ws.row_dimensions[ligne].height = 14


def ecrire_ligne(ws, ligne, label, evenements, fond_evenement, texte_evenement, hauteur=20):
    cell_label = ws.cell(row=ligne, column=1, value=label)
    cell_label.fill = FOND_SOMBRE
    cell_label.font = TEXTE_BLANC
    cell_label.alignment = Alignment(horizontal="right", vertical="center")

    for col in range(COL_DEBUT_GRILLE, COL_FIN_GRILLE + 1):
        ws.cell(row=ligne, column=col).fill = FOND_TIMELINE

    for texte, debut_min, fin_min in evenements:
        plage = plage_colonnes(debut_min, fin_min)
        if plage is None:
            continue
        col_debut, col_fin = plage
        if col_fin > col_debut:
            ws.merge_cells(start_row=ligne, start_column=col_debut, end_row=ligne, end_column=col_fin)
        cell = ws.cell(row=ligne, column=col_debut, value=texte)
        cell.fill = fond_evenement
        cell.font = texte_evenement
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[ligne].height = hauteur


def ajuster_largeurs_colonnes(ws):
    ws.column_dimensions["A"].width = 22
    for col in range(2, COL_FIN_GRILLE + 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 1.5


def construire_grille(output_path):
    scenes, concerts = charger_horaires_concerts(DATA_DIR / "horaires_paleo.xlsx")
    dates_par_jour = construire_dates_par_jour(concerts)

    wb = Workbook()
    wb.remove(wb.active)

    for jour in JOURS_FR:
        if jour.lower() not in JOURS_VALIDES or jour not in dates_par_jour:
            continue

        planning_path = OUTPUT_DIR / f"planning_{jour}.xlsx"
        if not planning_path.exists():
            print(f"Attention : {planning_path.name} introuvable, jour {jour} ignoré.")
            continue

        date = dates_par_jour[jour]
        titre = f"{jour} {date.day} {MOIS_FR[date.month - 1]} {date.year}"
        ws = wb.create_sheet(title=titre[:31])

        titre_cell = ws.cell(row=1, column=1, value=titre)
        titre_cell.font = Font(bold=True, size=14, color="FFFFFF")
        titre_cell.fill = FOND_SOMBRE
        ws.row_dimensions[1].height = 20

        ecrire_entete_heures(ws, 2)

        ligne = 3
        for scene in scenes:
            ecrire_ligne(ws, ligne, scene, concerts.get((scene, date), []), FOND_CONCERT, TEXTE_CONCERT, hauteur=26)
            ligne += 1

        ligne += 1
        ws.cell(row=ligne, column=1, value="Plannings").font = Font(bold=True, color="FFFFFF", size=8)
        ws.cell(row=ligne, column=1).fill = FOND_ENTETE
        ecrire_entete_heures(ws, ligne)
        ligne += 1

        taches = charger_taches_planning(planning_path, jour)
        for tache, evenements in taches.items():
            ecrire_ligne(ws, ligne, tache, evenements, FOND_PLANNING, TEXTE_PLANNING, hauteur=20)
            ligne += 1

        ajuster_largeurs_colonnes(ws)
        ws.freeze_panes = "C3"

    if not wb.sheetnames:
        raise RuntimeError(
            "Aucun jour n'a pu être construit : vérifiez que des fichiers "
            "output/planning_<Jour>.xlsx existent pour au moins un jour présent "
            "dans data/horaires_paleo.xlsx."
        )

    OUTPUT_DIR.mkdir(exist_ok=True)
    wb.save(output_path)
    return output_path


def main():
    output_path = OUTPUT_DIR / "Grille_horaire.xlsx"
    construire_grille(output_path)
    print(f"Grille horaire générée : {output_path}")


if __name__ == "__main__":
    main()
