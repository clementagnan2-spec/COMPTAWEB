"""
Application web SYSCOHADA — Fondations (étape 1)
===================================================
Réutilise core.py tel quel (moteur métier inchangé, déjà testé).
Remplace Tkinter par des pages HTML servies par Flask.

Lancement local :
    pip install -r requirements.txt
    python app.py
    -> http://127.0.0.1:5000

Variables d'environnement (pour la prod, voir DEPLOIEMENT.md) :
    DB_PATH        chemin du fichier SQLite (défaut : ./data/comptabilite.db)
    SECRET_KEY     clé secrète des sessions Flask (à définir en prod !)
"""
import os
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, session, g, flash

import core

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "data", "comptabilite.db"))
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-cle-a-changer-en-production")


# ---------------------------------------------------------------------------
# Connexion base de données — une connexion par requête (pattern Flask standard)
# ---------------------------------------------------------------------------
def get_db():
    if "db" not in g:
        g.db = core.get_connection(DB_PATH)
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


# ---------------------------------------------------------------------------
# Authentification — réutilise core.verify_password (déjà en place pour le
# serveur réseau existant, hash + sel, rien à réinventer)
# ---------------------------------------------------------------------------
def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def menu_requis(menu_key):
    """Décorateur : vérifie que l'utilisateur connecté a accès à ce sous-menu
    (réutilise core.get_menus_autorises, même règle que le bureau/client)."""
    def decorator(view):
        @wraps(view)
        @login_required
        def wrapped(*args, **kwargs):
            autorises = core.get_menus_autorises(get_db(), session["user"]["niveau_acces"])
            if menu_key not in autorises:
                flash("Vous n'avez pas accès à cette section.", "error")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapped
    return decorator


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        nom = request.form.get("nom_utilisateur", "")
        mdp = request.form.get("mot_de_passe", "")
        user = core.verify_password(get_db(), nom, mdp)
        if user:
            session["user"] = user
            dest = request.args.get("next") or url_for("dashboard")
            return redirect(dest)
        flash("Identifiant ou mot de passe incorrect.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Tableau de bord — menu filtré par niveau d'accès (core.MENU_STRUCTURE /
# core.get_menus_autorises, exactement la même règle que sur le bureau)
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    db = get_db()
    autorises = core.get_menus_autorises(db, session["user"]["niveau_acces"])
    menu = [
        (titre, [(label, key) for label, key in items if key in autorises])
        for titre, items in core.MENU_STRUCTURE
        if any(key in autorises for _label, key in items)
    ]
    return render_template("dashboard.html", menu=menu)


@app.route("/module/<menu_key>")
@login_required
def module_placeholder(menu_key):
    """Provisoire : les modules seront reconstruits un par un. Empêche un
    lien mort tant que l'écran correspondant n'existe pas encore."""
    db = get_db()
    autorises = core.get_menus_autorises(db, session["user"]["niveau_acces"])
    if menu_key not in autorises:
        flash("Vous n'avez pas accès à cette section.", "error")
        return redirect(url_for("dashboard"))
    titre = next((label for _t, items in core.MENU_STRUCTURE for label, key in items if key == menu_key),
                 menu_key)
    return render_template("placeholder.html", titre=titre)


# ---------------------------------------------------------------------------
# Amorçage : premier utilisateur administrateur si la base est neuve
# ---------------------------------------------------------------------------
@app.route("/premiere-configuration", methods=["GET", "POST"])
def premiere_configuration():
    db = get_db()
    if core.list_utilisateurs(db):
        return redirect(url_for("login"))
    if request.method == "POST":
        nom = request.form.get("nom_utilisateur", "").strip()
        mdp = request.form.get("mot_de_passe", "")
        if not nom or not mdp:
            flash("Identifiant et mot de passe obligatoires.", "error")
        else:
            core.ajouter_niveaux_acces_suggeres(db)
            core.ajouter_niveaux_acces_suggeres_menus(db)
            core.add_utilisateur(db, nom, mdp, nom_complet="Administrateur",
                                  niveau_acces="Administrateur", actif=True)
            flash("Compte administrateur créé. Connectez-vous.", "success")
            return redirect(url_for("login"))
    return render_template("premiere_configuration.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
