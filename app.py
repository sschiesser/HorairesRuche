"""
Application web pour générer le planning sans utiliser de terminal.

Les fichiers Excel (data/personnes.xlsx, plages_horaires.xlsx, etc.) restent
édités directement sur cette machine, exactement comme avec
generate_planning.py en ligne de commande. Cette application web sert
uniquement à choisir le jour et à générer/télécharger le planning depuis un
navigateur.

Utilisation :
    pip install -r requirements.txt
    python app.py
    -> ouvre http://localhost:5000 sur cette machine.
       Les autres ordinateurs du même réseau local (même WiFi) peuvent y
       accéder via http://<adresse-IP-de-cette-machine>:5000
"""

from flask import Flask, abort, render_template, request, send_file

from generate_planning import OUTPUT_DIR, generer_planning, lister_jours_disponibles

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html", jours=lister_jours_disponibles())


@app.route("/generer", methods=["POST"])
def generer():
    jour = request.form.get("jour", "").strip()
    jours = lister_jours_disponibles()
    if jour not in jours:
        return render_template("index.html", jours=jours, erreur=f"Jour invalide : {jour}"), 400

    try:
        resultat = generer_planning(jour)
    except (ValueError, FileNotFoundError) as e:
        return render_template("index.html", jours=jours, erreur=str(e)), 400

    return render_template("resultat.html", jour=jour, resultat=resultat)


@app.route("/telecharger/<jour>")
def telecharger(jour):
    chemin = OUTPUT_DIR / f"planning_{jour}.xlsx"
    if not chemin.exists():
        abort(404)
    return send_file(chemin, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
