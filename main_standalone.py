"""
VoeuxConcerts — application standalone.

Génère en une seule action le planning et la grille horaire pour un jour donné.
Les fichiers Excel dans data/ restent editables à côté de l'exécutable.
Les résultats sont écrits dans output/.

Utilisation :
    double-cliquer sur VoeuxConcerts.exe
    — ou —
    python main_standalone.py
"""

import sys

from generate_planning import (
    DATA_DIR, OUTPUT_DIR,
    demander_jour, generer_planning, lister_jours_disponibles,
)
from generate_grille import (
    charger_horaires_concerts, charger_taches_planning,
    construire_dates_par_jour, construire_grille_jour,
)


def main():
    print("=" * 52)
    print("  VoeuxConcerts — Générateur de planning")
    print("=" * 52)
    print()

    try:
        jours = lister_jours_disponibles()
    except FileNotFoundError as e:
        print(f"Erreur : {e}")
        print("Vérifiez que le dossier data/ se trouve à côté de l'application.")
        input("\nAppuyez sur Entrée pour quitter.")
        sys.exit(1)

    if not jours:
        print("Aucun jour disponible dans plages_horaires.xlsx.")
        input("\nAppuyez sur Entrée pour quitter.")
        sys.exit(1)

    jour = demander_jour(jours)

    # ---- Planning -------------------------------------------------------
    print(f"\n--- Planning {jour} ---")
    try:
        resultat = generer_planning(jour)
    except (ValueError, FileNotFoundError) as e:
        print(f"Erreur : {e}")
        input("\nAppuyez sur Entrée pour quitter.")
        sys.exit(1)

    for msg in resultat["messages"]:
        print(msg)

    if not resultat["success"]:
        input("\nAppuyez sur Entrée pour quitter.")
        sys.exit(1)

    # ---- Grille ---------------------------------------------------------
    print(f"\n--- Grille {jour} ---")
    try:
        scenes, concerts = charger_horaires_concerts(DATA_DIR / "horaires_paleo.xlsx")
        dates_par_jour = construire_dates_par_jour(concerts)
        if jour not in dates_par_jour:
            print(f"Attention : {jour} introuvable dans horaires_paleo.xlsx — grille ignorée.")
        else:
            taches = charger_taches_planning(OUTPUT_DIR / f"planning_{jour}.xlsx", jour)
            wb = construire_grille_jour(jour, dates_par_jour[jour], scenes, concerts, taches)
            OUTPUT_DIR.mkdir(exist_ok=True)
            output_path = OUTPUT_DIR / f"Grille_{jour}.xlsx"
            wb.save(output_path)
            print(f"Grille générée : {output_path}")
    except Exception as e:
        print(f"Erreur lors de la génération de la grille : {e}")

    print()
    input("Appuyez sur Entrée pour quitter.")


if __name__ == "__main__":
    main()
