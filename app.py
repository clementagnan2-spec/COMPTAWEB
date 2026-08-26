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
import json as _json
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, session, g, flash, jsonify

import core

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "data", "comptabilite.db"))
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-cle-a-changer-en-production")


@app.template_filter("cfa")
def fmt_cfa(v):
    """Formate un montant en F CFA (espace comme séparateur de milliers,
    aucune décimale — convention SYSCOHADA). Vide si non numérique."""
    if v is None or v == "":
        return ""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    sign = "-" if v < 0 else ""
    v = abs(v)
    entier = f"{v:,.0f}".replace(",", " ")
    return f"{sign}{entier}"


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
    if not core.list_utilisateurs(get_db()):
        return redirect(url_for("premiere_configuration"))
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
# Table de correspondance menu_key -> endpoint Flask réellement construit.
# Les clés absentes de ce dict retombent sur l'écran "en construction"
# (module_placeholder). Complétée au fur et à mesure des modules livrés.
# ---------------------------------------------------------------------------
MENU_ROUTES = {
    "saisie": "saisie",
    "ouverture": "ouverture",
    "grand_livre": "grand_livre",
    "balance": "balance",
    "bilan_syscohada": "bilan_syscohada",
    "compte_resultat_sig": "compte_resultat_sig",
    "tft": "tft",
    "situation_financiere": "situation_financiere",
    "fournisseurs": "fournisseurs",
    "clients": "clients",
    "facturation": "facturation",
    "stocks": "stocks",
    "immobilisations": "immobilisations",
    "amortissements": "amortissements",
}


@app.template_global()
def menu_url(menu_key):
    if menu_key in MENU_ROUTES:
        return url_for(MENU_ROUTES[menu_key])
    return url_for("module_placeholder", menu_key=menu_key)


@app.context_processor
def inject_menu():
    """Rend le menu (filtré par niveau d'accès) disponible dans TOUS les
    templates automatiquement, sans que chaque route ait à le recalculer et
    le passer explicitement — évite qu'une nouvelle page oublie de le faire
    et se retrouve avec une barre latérale vide."""
    if "user" not in session:
        return {"menu": []}
    try:
        db = get_db()
        autorises = core.get_menus_autorises(db, session["user"]["niveau_acces"])
        menu = [
            (titre, [(label, key) for label, key in items if key in autorises])
            for titre, items in core.MENU_STRUCTURE
            if any(key in autorises for _label, key in items)
        ]
        return {"menu": menu}
    except Exception:
        return {"menu": []}


# ---------------------------------------------------------------------------
# Tableau de bord — menu filtré par niveau d'accès (core.MENU_STRUCTURE /
# core.get_menus_autorises, exactement la même règle que sur le bureau)
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    return render_template("dashboard.html")


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


# ---------------------------------------------------------------------------
# Utilitaire commun : exercice comptable actif (choisi en haut de chaque
# page ; mémorisé en session, comme le sélecteur "EXERCICE COMPTABLE" du
# bureau)
# ---------------------------------------------------------------------------
def exercice_actif(db):
    ex = request.args.get("exercice") or session.get("exercice") or core.get_current_exercice(db)
    session["exercice"] = ex
    return ex


@app.route("/api/comptes")
@login_required
def api_comptes():
    """Recherche de comptes pour l'auto-complétion (Saisie, Grand livre...)."""
    q = request.args.get("q", "")
    rows = core.search_accounts(get_db(), q, limit=30)
    return jsonify([dict(r) for r in rows])


# ---------------------------------------------------------------------------
# SAISIE > Saisie des écritures
# ---------------------------------------------------------------------------
@app.route("/module/saisie", methods=["GET", "POST"])
@menu_requis("saisie")
def saisie():
    db = get_db()
    exercice = exercice_actif(db)

    if request.method == "POST":
        try:
            date_str = request.form["date"]
            piece = request.form.get("piece", "").strip()
            journal = request.form.get("journal", "OD")
            tiers = request.form.get("tiers", "").strip()
            comptes = request.form.getlist("ligne_compte")
            libelles = request.form.getlist("ligne_libelle")
            debits = request.form.getlist("ligne_debit")
            credits = request.form.getlist("ligne_credit")
            lignes = []
            for i in range(len(comptes)):
                if not comptes[i].strip():
                    continue
                lignes.append({
                    "compte": comptes[i].strip(),
                    "libelle": libelles[i].strip() if i < len(libelles) else "",
                    "debit": float(debits[i]) if i < len(debits) and debits[i] else 0,
                    "credit": float(credits[i]) if i < len(credits) and credits[i] else 0,
                })
            core.add_ecriture_multi_lignes(db, date_str, piece, journal, lignes, tiers=tiers)
            flash("Écriture enregistrée.", "success")
            return redirect(url_for("saisie", exercice=exercice))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")

    entries = core.list_entries(db, order_by="date", exercice=exercice)
    entries.sort(key=lambda e: (e["date"], e["id"]), reverse=True)
    return render_template("saisie.html", entries=entries, exercice=exercice,
                            exercices=core.list_exercices(db),
                            today=core.date.today().isoformat())


@app.route("/module/saisie/supprimer/<int:entry_id>", methods=["POST"])
@menu_requis("saisie")
def saisie_supprimer(entry_id):
    db = get_db()
    try:
        core.delete_entry(db, entry_id)
        flash("Écriture supprimée.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("saisie", exercice=session.get("exercice")))


# ---------------------------------------------------------------------------
# SAISIE > Soldes d'ouverture
# ---------------------------------------------------------------------------
@app.route("/module/ouverture", methods=["GET", "POST"])
@menu_requis("ouverture")
def ouverture():
    db = get_db()
    exercice = exercice_actif(db)

    if request.method == "POST":
        try:
            code = request.form["compte"].strip()
            valeur = float(request.form.get("solde") or 0)
            if not core.account_exists(db, code):
                raise ValueError(f"Le compte « {code} » n'existe pas dans le plan comptable.")
            core.set_opening_balance(db, code, valeur, exercice=exercice)
            flash("Solde d'ouverture enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("ouverture", exercice=exercice))

    soldes = core.list_opening_balances(db, exercice=exercice)
    return render_template("ouverture.html", soldes=soldes, exercice=exercice,
                            exercices=core.list_exercices(db))


# ---------------------------------------------------------------------------
# RAPPORTS FINANCIERS
# ---------------------------------------------------------------------------
@app.route("/module/grand_livre")
@menu_requis("grand_livre")
def grand_livre():
    db = get_db()
    exercice = exercice_actif(db)
    compte = request.args.get("compte", "").strip()
    lignes = core.compute_grand_livre(db, compte, exercice=exercice) if compte else []
    return render_template("grand_livre.html", lignes=lignes, compte=compte,
                            exercice=exercice, exercices=core.list_exercices(db))


@app.route("/module/balance")
@menu_requis("balance")
def balance():
    db = get_db()
    exercice = exercice_actif(db)
    lignes = core.compute_balance(db, only_with_movement=True, exercice=exercice)
    lignes.sort(key=lambda b: b["code"])
    totaux = {
        "debit": sum(l["debit"] for l in lignes), "credit": sum(l["credit"] for l in lignes),
        "solde_ouverture": sum(l["solde_ouverture"] for l in lignes),
        "solde_cloture": sum(l["solde_cloture"] for l in lignes),
    }
    return render_template("balance.html", lignes=lignes, totaux=totaux, exercice=exercice,
                            exercices=core.list_exercices(db))


@app.route("/module/bilan_syscohada")
@menu_requis("bilan_syscohada")
def bilan_syscohada():
    db = get_db()
    exercice = exercice_actif(db)
    try:
        d = core.compute_bilan_plat(db, exercice=exercice)
        erreur = None
    except Exception as exc:
        d, erreur = None, str(exc)
    return render_template("bilan.html", d=d, erreur=erreur, exercice=exercice,
                            exercices=core.list_exercices(db))


def _rapport_formule_generique(menu_key, template_name, compute_fn):
    """Fabrique une route pour un écran RAPPORTS FINANCIERS basé sur
    compute_etat_formule_generique (CR, TFT, Situation financière) — les
    trois ont exactement la même structure de résultat
    ({colonnes, lignes, errors})."""
    def view():
        db = get_db()
        exercice = exercice_actif(db)
        try:
            d = compute_fn(db, exercice=exercice)
            erreur = None
        except Exception as exc:
            d, erreur = None, str(exc)
        return render_template(template_name, d=d, erreur=erreur, exercice=exercice,
                                exercices=core.list_exercices(db))
    view.__name__ = menu_key
    return view


app.add_url_rule("/module/compte_resultat_sig", "compte_resultat_sig",
                  menu_requis("compte_resultat_sig")(
                      _rapport_formule_generique("compte_resultat_sig", "etat_formule.html", core.compute_cr)))
app.add_url_rule("/module/tft", "tft",
                  menu_requis("tft")(
                      _rapport_formule_generique("tft", "etat_formule.html", core.compute_tft_gabarit)))
app.add_url_rule("/module/situation_financiere", "situation_financiere",
                  menu_requis("situation_financiere")(
                      _rapport_formule_generique("situation_financiere", "etat_formule.html",
                                                  core.compute_situation_fin)))


# ---------------------------------------------------------------------------
# ENGAGEMENTS-PROJETS > Fournisseurs & COMMERCIAL > Clients
# ---------------------------------------------------------------------------
@app.route("/module/fournisseurs", methods=["GET", "POST"])
@menu_requis("fournisseurs")
def fournisseurs():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_fournisseur(
                db, request.form["code"], request.form["raison_sociale"],
                contact=request.form.get("contact", ""), telephone=request.form.get("telephone", ""),
                adresse=request.form.get("adresse", ""),
                delai_paiement_jours=request.form.get("delai_paiement_jours") or 30,
                delai_livraison_jours=request.form.get("delai_livraison_jours") or 15,
            )
            flash("Fournisseur enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("fournisseurs"))
    q = request.args.get("q", "")
    return render_template("fournisseurs.html", fournisseurs=core.list_fournisseurs(db, query=q or None), q=q)


@app.route("/module/fournisseurs/supprimer/<code>", methods=["POST"])
@menu_requis("fournisseurs")
def fournisseurs_supprimer(code):
    try:
        core.delete_fournisseur(get_db(), code)
        flash("Fournisseur supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("fournisseurs"))


@app.route("/module/clients", methods=["GET", "POST"])
@menu_requis("clients")
def clients():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_client(
                db, request.form["code"], request.form["raison_sociale"],
                contact=request.form.get("contact", ""), telephone=request.form.get("telephone", ""),
                adresse=request.form.get("adresse", ""),
                delai_paiement_jours=request.form.get("delai_paiement_jours") or 30,
            )
            flash("Client enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("clients"))
    q = request.args.get("q", "")
    return render_template("clients.html", clients=core.list_clients(db, query=q or None), q=q)


@app.route("/module/clients/supprimer/<code>", methods=["POST"])
@menu_requis("clients")
def clients_supprimer(code):
    try:
        core.delete_client(get_db(), code)
        flash("Client supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("clients"))


# ---------------------------------------------------------------------------
# COMMERCIAL > Facturation
# ---------------------------------------------------------------------------
@app.route("/module/facturation", methods=["GET", "POST"])
@menu_requis("facturation")
def facturation():
    db = get_db()
    if request.method == "POST":
        try:
            fid = core.create_facture_vente(
                db, request.form["numero"], request.form["date_facture"], request.form["client_code"],
            )
            flash("Facture créée en brouillon — ajoutez des lignes.", "success")
            return redirect(url_for("facturation_detail", facture_id=fid))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("facturation"))
    return render_template("facturation.html", factures=core.list_factures_vente(db),
                            clients=core.list_clients(db), today=core.date.today().isoformat())


@app.route("/module/facturation/<int:facture_id>", methods=["GET", "POST"])
@menu_requis("facturation")
def facturation_detail(facture_id):
    db = get_db()
    facture = core.get_facture_vente(db, facture_id)
    if not facture:
        flash("Facture introuvable.", "error")
        return redirect(url_for("facturation"))

    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "ajouter_ligne":
                core.add_ligne_facture_vente(
                    db, facture_id, request.form["compte_vente"], request.form["libelle"],
                    float(request.form["quantite"]), float(request.form["prix_unitaire"]),
                )
                flash("Ligne ajoutée.", "success")
            elif action == "valider":
                warnings = core.valider_facture_vente(db, facture_id)
                flash("Facture validée — écritures envoyées en Saisie.", "success")
                for w in warnings:
                    flash(w, "error")
            elif action == "devalider":
                core.devalider_facture_vente(db, facture_id)
                flash("Facture repassée en brouillon.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("facturation_detail", facture_id=facture_id))

    facture = core.get_facture_vente(db, facture_id)
    lignes = core.list_lignes_facture_vente(db, facture_id)
    totals = core.compute_facture_totals(db, facture_id)
    return render_template("facturation_detail.html", facture=facture, lignes=lignes, totals=totals)


@app.route("/module/facturation/<int:facture_id>/supprimer", methods=["POST"])
@menu_requis("facturation")
def facturation_supprimer(facture_id):
    try:
        core.delete_facture_vente(get_db(), facture_id)
        flash("Facture supprimée.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("facturation"))


@app.route("/module/facturation/ligne/<int:ligne_id>/supprimer", methods=["POST"])
@menu_requis("facturation")
def facturation_ligne_supprimer(ligne_id):
    facture_id = request.form.get("facture_id", type=int)
    core.delete_ligne_facture_vente(get_db(), ligne_id)
    flash("Ligne supprimée.", "success")
    return redirect(url_for("facturation_detail", facture_id=facture_id))


@app.route("/module/facturation/<int:facture_id>/apercu")
@menu_requis("facturation")
def facturation_apercu(facture_id):
    try:
        html = core.render_facture_vente_html(get_db(), facture_id)
        return html
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("facturation_detail", facture_id=facture_id))


# ---------------------------------------------------------------------------
# COMMERCIAL/PRODUCTION > Stocks (Matières premières, Produits finis —
# les 3 sous-menus du bureau pointent tous vers cette même clé "stocks")
# ---------------------------------------------------------------------------
@app.route("/module/stocks", methods=["GET", "POST"])
@menu_requis("stocks")
def stocks():
    db = get_db()
    exercice = exercice_actif(db)

    if request.method == "POST":
        try:
            code = request.form["compte"].strip()
            valeur = float(request.form.get("valeur") or 0)
            quantite = float(request.form.get("quantite") or 0)
            core.set_stock_initial(db, code, valeur, exercice=exercice)
            core.set_stock_qte_initiale(db, code, quantite, exercice=exercice)
            flash("Stock initial enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("stocks", exercice=exercice))

    synthese = core.compute_stocks(db, exercice=exercice)
    detail = core.compute_stocks_detail(db, exercice=exercice)
    return render_template("stocks.html", synthese=synthese, detail=detail, exercice=exercice,
                            exercices=core.list_exercices(db))


# ---------------------------------------------------------------------------
# IMMOBILISATIONS > Immobilisations & Amortissements
# ---------------------------------------------------------------------------
@app.route("/module/immobilisations", methods=["GET", "POST"])
@menu_requis("immobilisations")
def immobilisations():
    db = get_db()
    exercice = exercice_actif(db)

    if request.method == "POST":
        try:
            compte = request.form["compte"].strip()
            if not core.account_exists(db, compte):
                raise ValueError(f"Le compte « {compte} » n'existe pas dans le plan comptable.")
            base_qte = request.form.get("base_repartition_quantite")
            amort_manuel = request.form.get("amortissement_annuel_manuel")
            core.set_immobilisation_fiche(
                db, compte,
                fournisseur_code=request.form.get("fournisseur_code") or None,
                prix_achat=float(request.form["prix_achat"]) if request.form.get("prix_achat") else None,
                date_acquisition=request.form.get("date_acquisition") or None,
                base_repartition_quantite=float(base_qte) if base_qte else None,
                base_repartition_unite=request.form.get("base_repartition_unite") or None,
                amortissement_annuel_manuel=float(amort_manuel) if amort_manuel else None,
            )
            flash("Fiche immobilisation enregistrée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("immobilisations", exercice=exercice))

    liste = core.compute_immobilisations_liste(db, exercice=exercice)
    return render_template("immobilisations.html", liste=liste, exercice=exercice,
                            exercices=core.list_exercices(db), fournisseurs=core.list_fournisseurs(db))


@app.route("/module/amortissements", methods=["GET", "POST"])
@menu_requis("amortissements")
def amortissements():
    db = get_db()
    categories = [c for c, _b, _a in core.IMMO_CATEGORIES]
    if request.method == "POST":
        for i, categorie in enumerate(categories):
            key = f"taux_{i}"
            if key in request.form and request.form[key] != "":
                core.set_taux_amortissement(db, categorie, float(request.form[key]))
        flash("Taux d'amortissement enregistrés.", "success")
        return redirect(url_for("amortissements"))
    taux = core.list_taux_amortissement(db)
    for i, t in enumerate(taux):
        t["field_key"] = f"taux_{i}"
    return render_template("amortissements.html", taux=taux)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
