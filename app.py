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
import tempfile
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, session, g, flash, jsonify, send_file

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


DEFAULT_ADMIN_USER = "admin"
DEFAULT_ADMIN_PASSWORD = "1234"


def ensure_default_admin(db):
    """Tant qu'aucun disque persistant n'est configuré (voir DEPLOIEMENT.md),
    la base repart à zéro à chaque redéploiement. Pour éviter de repasser par
    l'écran de création à chaque fois, on recrée automatiquement un compte
    admin/1234 s'il n'existe aucun utilisateur.
    SÉCURITÉ : ce mot de passe par défaut est volontairement faible et connu
    de quiconque lit ce code — à changer dès que possible via ADMIN >
    Utilisateurs, en particulier une fois des données réelles saisies."""
    if core.list_utilisateurs(db):
        return False
    core.ajouter_niveaux_acces_suggeres(db)
    core.ajouter_niveaux_acces_suggeres_menus(db)
    core.add_utilisateur(db, DEFAULT_ADMIN_USER, DEFAULT_ADMIN_PASSWORD,
                          nom_complet="Administrateur", niveau_acces="Administrateur", actif=True)
    return True


@app.route("/login", methods=["GET", "POST"])
def login():
    db = get_db()
    if ensure_default_admin(db):
        flash(f"Compte administrateur par défaut créé : identifiant « {DEFAULT_ADMIN_USER} », "
              f"mot de passe « {DEFAULT_ADMIN_PASSWORD} ». Changez ce mot de passe dès que possible "
              "dans ADMIN > Utilisateurs.", "error")
    if request.method == "POST":
        nom = request.form.get("nom_utilisateur", "")
        mdp = request.form.get("mot_de_passe", "")
        user = core.verify_password(db, nom, mdp)
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
    "grh_personnel": "grh_personnel",
    "grh_paie": "grh_paie",
    "exercices": "exercices",
    "plan_comptable": "plan_comptable",
    "plan_analytique": "plan_analytique",
    "plan_budgetaire": "plan_budgetaire",
    "plan_bailleur": "plan_bailleur",
    "taux_tva": "taux_tva",
    "taux_retenue": "taux_retenue",
    "niveaux_acces": "niveaux_acces",
    "utilisateurs": "utilisateurs",
    "production": "production",
    "tresorerie": "tresorerie",
    "contrats": "contrats",
    "expression_besoin": "expression_besoin",
    "ep_bon_commande": "ep_bon_commande",
    "bordereau_livraison": "bordereau_livraison",
    "reglements": "reglements",
    "recouvrement": "recouvrement",
    "marges": "marges",
    "admin_factures": "admin_factures",
    "admin_modele_bon_commande": "admin_modele_bon_commande",
    "reinitialisation": "reinitialisation",
    "grh_time_sheet": "grh_time_sheet",
    "grh_kpi": "grh_kpi",
    "grh_tableau_bord": "grh_tableau_bord",
    "grh_hs": "grh_hs",
    "transport": "transport",
    "missions": "missions",
    "pieces_rechange": "pieces_rechange",
    "reparations": "reparations",
    "energie": "energie",
    "maintenance": "maintenance",
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


# ---------------------------------------------------------------------------
# Import / Export Excel — mécanisme générique réutilisé par tous les écrans
# concernés (chaque fonction core.export_xxx_xlsx(conn, path, ...) écrit sur
# disque ; chaque core.import_xxx_xlsx(conn, path) lit un fichier uploadé).
# ---------------------------------------------------------------------------
def export_xlsx_response(export_fn, filename, *args, **kwargs):
    """Appelle export_fn(*args, tmp_path, **kwargs) (signature core.py :
    (conn, path, ...) ou parfois juste (path,)) et renvoie le fichier."""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        export_fn(*args, tmp_path, **kwargs)
        return send_file(tmp_path, as_attachment=True, download_name=filename,
                          mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        # send_file lit le fichier en streaming ; on nettoie via un callback
        # n'est pas nécessaire ici car l'OS supprimera /tmp — mais on essaie
        # quand même de bonne foi, en silence si le fichier est encore ouvert.
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def handle_import_upload(import_fn, *args, redirect_endpoint, file_field="fichier", **kwargs):
    """Gère l'upload d'un .xlsx depuis un <input type=file>, appelle
    import_fn(*args, tmp_path, **kwargs) puis affiche un résumé et redirige.
    Les fonctions core.import_xxx_xlsx renvoient des formats hétérogènes
    (dict {crees,mis_a_jour,erreurs} / tuple (nb, avertissements) / simple
    entier) — ce bloc les interprète tous sans supposer un format précis."""
    fichier = request.files.get(file_field)
    if not fichier or not fichier.filename:
        flash("Aucun fichier sélectionné.", "error")
        return redirect(url_for(redirect_endpoint))
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        fichier.save(tmp_path)
        resultat = import_fn(*args, tmp_path, **kwargs)
        erreurs = []
        if isinstance(resultat, dict):
            crees = resultat.get("crees")
            maj = resultat.get("mis_a_jour")
            erreurs = resultat.get("erreurs") or []
            parts = []
            if crees is not None:
                parts.append(f"{crees} créé(s)")
            if maj is not None:
                parts.append(f"{maj} mis à jour")
            flash("Import terminé — " + (", ".join(parts) if parts else f"{len(erreurs)} avertissement(s)") + ".",
                  "success")
        elif isinstance(resultat, tuple) and len(resultat) == 2:
            nb, erreurs = resultat
            erreurs = erreurs or []
            flash(f"Import terminé — {nb} ligne(s) importée(s).", "success")
        elif isinstance(resultat, int):
            flash(f"Import terminé — {resultat} ligne(s) importée(s).", "success")
        else:
            flash("Import terminé.", "success")
        for e in erreurs[:10]:
            flash(str(e), "error")
    except (ValueError, KeyError) as exc:
        flash(f"Erreur d'import : {exc}", "error")
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
    return redirect(url_for(redirect_endpoint))


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
            analytics = request.form.getlist("ligne_analytic")
            lignes = []
            for i in range(len(comptes)):
                if not comptes[i].strip():
                    continue
                lignes.append({
                    "compte": comptes[i].strip(),
                    "libelle": libelles[i].strip() if i < len(libelles) else "",
                    "debit": float(debits[i]) if i < len(debits) and debits[i] else 0,
                    "credit": float(credits[i]) if i < len(credits) and credits[i] else 0,
                    "analytic_code": analytics[i].strip() if i < len(analytics) and analytics[i].strip() else None,
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


# ---------------------------------------------------------------------------
# GRH > Liste du personnel
# ---------------------------------------------------------------------------
@app.route("/module/grh_personnel", methods=["GET", "POST"])
@menu_requis("grh_personnel")
def grh_personnel():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_personnel(
                db, request.form["nom"], matricule=request.form.get("matricule", ""),
                prenom=request.form.get("prenom", ""), poste=request.form.get("poste", ""),
                service=request.form.get("service", ""), date_embauche=request.form.get("date_embauche", ""),
                telephone=request.form.get("telephone", ""), email=request.form.get("email", ""),
                salaire_base=float(request.form.get("salaire_base") or 0),
                statut=request.form.get("statut", "actif"),
            )
            flash("Employé ajouté.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_personnel"))
    return render_template("grh_personnel.html", personnel=core.list_personnel(db))


@app.route("/module/grh_personnel/supprimer/<int:personnel_id>", methods=["POST"])
@menu_requis("grh_personnel")
def grh_personnel_supprimer(personnel_id):
    try:
        core.delete_personnel(get_db(), personnel_id)
        flash("Employé supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("grh_personnel"))


# ---------------------------------------------------------------------------
# GRH > Paie (Bulletins, État de paie, Paramètres)
# ---------------------------------------------------------------------------
def _periode_actuelle():
    return request.args.get("periode") or session.get("paie_periode") or core.date.today().strftime("%Y-%m")


@app.route("/module/grh_paie")
@menu_requis("grh_paie")
def grh_paie():
    periode = _periode_actuelle()
    session["paie_periode"] = periode
    return redirect(url_for("grh_paie_bulletins", periode=periode))


@app.route("/module/grh_paie/bulletins", methods=["GET", "POST"])
@menu_requis("grh_paie")
def grh_paie_bulletins():
    db = get_db()
    periode = _periode_actuelle()
    session["paie_periode"] = periode

    if request.method == "POST":
        try:
            personnel_id = int(request.form["personnel_id"])
            core.set_bulletin_paie(
                db, personnel_id, periode,
                classification=request.form.get("classification", "AUTRE"),
                salaire_base=float(request.form.get("salaire_base") or 0),
                prime_anciennete=float(request.form.get("prime_anciennete") or 0),
                heures_sup=float(request.form.get("heures_sup") or 0),
                sursalaire=float(request.form.get("sursalaire") or 0),
                gratification=float(request.form.get("gratification") or 0),
                indemnite_caisse=float(request.form.get("indemnite_caisse") or 0),
                indemnite_logement=float(request.form.get("indemnite_logement") or 0),
                indemnite_fonction=float(request.form.get("indemnite_fonction") or 0),
                indemnite_transport=float(request.form.get("indemnite_transport") or 0),
                personnes_a_charge=int(request.form.get("personnes_a_charge") or 0),
                retenue_pret=float(request.form.get("retenue_pret") or 0),
            )
            flash("Bulletin enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_paie_bulletins", periode=periode))

    bulletins = core.list_bulletins_paie(db, periode=periode)
    verrouille = core.est_periode_paie_validee(db, periode)
    return render_template("grh_paie_bulletins.html", personnel=core.list_personnel(db, actifs_only=True),
                            bulletins=bulletins, periode=periode, verrouille=verrouille)


@app.route("/module/grh_paie/bulletins/supprimer/<int:bulletin_id>", methods=["POST"])
@menu_requis("grh_paie")
def grh_paie_bulletin_supprimer(bulletin_id):
    try:
        core.delete_bulletin_paie(get_db(), bulletin_id)
        flash("Bulletin supprimé.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("grh_paie_bulletins", periode=session.get("paie_periode")))


@app.route("/module/grh_paie/etat", methods=["GET", "POST"])
@menu_requis("grh_paie")
def grh_paie_etat():
    db = get_db()
    periode = _periode_actuelle()
    session["paie_periode"] = periode

    if request.method == "POST" and request.form.get("action") == "valider":
        try:
            etat, piece = core.valider_paie_periode(db, periode)
            flash(f"Paie {periode} comptabilisée (pièce {piece}).", "success")
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_paie_etat", periode=periode))

    etat = core.compute_paie_periode(db, periode)
    verrouille = core.est_periode_paie_validee(db, periode)
    return render_template("grh_paie_etat.html", etat=etat, periode=periode, verrouille=verrouille)


@app.route("/module/grh_paie/parametres", methods=["GET", "POST"])
@menu_requis("grh_paie")
def grh_paie_parametres():
    db = get_db()
    if session["user"]["niveau_acces"] != "Administrateur":
        flash("Réservé à l'administrateur.", "error")
        return redirect(url_for("grh_paie_bulletins"))
    if request.method == "POST":
        try:
            params = core.get_paie_parametres(db)
            params["taux_cnss_salarie"] = float(request.form["taux_cnss_salarie"]) / 100
            params["plafond_cnss"] = float(request.form["plafond_cnss"])
            params["cnss_salariale_plafonnee"] = float(request.form["cnss_salariale_plafonnee"])
            params["taux_cnss_patronale"] = float(request.form["taux_cnss_patronale"]) / 100
            params["taux_tpa"] = float(request.form["taux_tpa"]) / 100
            params["taux_retenue_obligatoire"] = float(request.form["taux_retenue_obligatoire"]) / 100
            params["abattement_cadre"] = float(request.form["abattement_cadre"]) / 100
            params["abattement_autre"] = float(request.form["abattement_autre"]) / 100
            core.set_paie_parametres(db, params)
            flash("Paramètres de paie enregistrés.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_paie_parametres"))
    params = core.get_paie_parametres(db)
    return render_template("grh_paie_parametres.html", p=params)


@app.route("/module/grh_paie/bulletin/<int:bulletin_id>/apercu")
@menu_requis("grh_paie")
def grh_paie_apercu(bulletin_id):
    try:
        return core.render_bulletin_paie_html(get_db(), bulletin_id)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("grh_paie_bulletins"))


# ---------------------------------------------------------------------------
# PARAMÈTRES > Exercices comptables (clôture)
# ---------------------------------------------------------------------------
@app.route("/module/exercices", methods=["GET", "POST"])
@menu_requis("exercices")
def exercices():
    db = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "nouveau":
                core.set_current_exercice(db, request.form["exercice"].strip())
                flash("Exercice créé/activé.", "success")
            elif action == "activer":
                core.set_current_exercice(db, request.form["exercice"])
                session["exercice"] = request.form["exercice"]
                flash("Exercice actif changé.", "success")
            elif action == "cloturer":
                ex = request.form["exercice"]
                next_ex = core.close_exercice(db, ex)
                core.set_current_exercice(db, next_ex)
                session["exercice"] = next_ex
                flash(f"Exercice {ex} clôturé. Soldes d'ouverture de {next_ex} calculés.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("exercices"))

    liste = core.list_exercices(db)
    bilans = {}
    for e in liste:
        try:
            bilans[e["exercice"]] = core.compute_bilan(db, exercice=e["exercice"])["ecart"]
        except Exception:
            bilans[e["exercice"]] = None
    return render_template("exercices.html", liste=liste, bilans=bilans,
                            exercice_actif=core.get_current_exercice(db))


# ---------------------------------------------------------------------------
# PARAMÈTRES > Plans auxiliaires (comptable / analytique / budgétaire / bailleurs)
# ---------------------------------------------------------------------------
@app.route("/module/plan_comptable", methods=["GET", "POST"])
@menu_requis("plan_comptable")
def plan_comptable():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_account(db, request.form["code"].strip(), request.form["label"].strip())
            flash("Compte enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("plan_comptable"))
    q = request.args.get("q", "")
    comptes = core.search_accounts(db, q, limit=500) if q else core.search_accounts(db, "", limit=500)
    return render_template("plan_comptable.html", comptes=comptes, q=q)


@app.route("/module/plan_comptable/supprimer/<code>", methods=["POST"])
@menu_requis("plan_comptable")
def plan_comptable_supprimer(code):
    try:
        core.delete_account(get_db(), code)
        flash("Compte supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("plan_comptable"))


@app.route("/module/plan_analytique", methods=["GET", "POST"])
@menu_requis("plan_analytique")
def plan_analytique():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_analytic_code(db, request.form["code"].strip(), request.form["label"].strip(),
                                    unite=request.form.get("unite") or None)
            flash("Code analytique enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("plan_analytique"))
    return render_template("plan_analytique.html", codes=core.list_analytic_codes(db))


@app.route("/module/plan_analytique/supprimer/<code>", methods=["POST"])
@menu_requis("plan_analytique")
def plan_analytique_supprimer(code):
    try:
        core.delete_analytic_code(get_db(), code)
        flash("Code analytique supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("plan_analytique"))


@app.route("/module/plan_budgetaire", methods=["GET", "POST"])
@menu_requis("plan_budgetaire")
def plan_budgetaire():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_budget_code(db, request.form["code"].strip(), request.form["label"].strip(),
                                  montant=float(request.form.get("montant") or 0))
            flash("Ligne budgétaire enregistrée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("plan_budgetaire"))
    return render_template("plan_budgetaire.html", codes=core.list_budget_codes(db))


@app.route("/module/plan_budgetaire/supprimer/<code>", methods=["POST"])
@menu_requis("plan_budgetaire")
def plan_budgetaire_supprimer(code):
    try:
        core.delete_budget_code(get_db(), code)
        flash("Ligne budgétaire supprimée.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("plan_budgetaire"))


@app.route("/module/plan_bailleur", methods=["GET", "POST"])
@menu_requis("plan_bailleur")
def plan_bailleur():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_donor_code(db, request.form["code"].strip(), request.form["label"].strip())
            flash("Bailleur enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("plan_bailleur"))
    return render_template("plan_bailleur.html", codes=core.list_donor_codes(db))


@app.route("/module/plan_bailleur/supprimer/<code>", methods=["POST"])
@menu_requis("plan_bailleur")
def plan_bailleur_supprimer(code):
    try:
        core.delete_donor_code(get_db(), code)
        flash("Bailleur supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("plan_bailleur"))


# ---------------------------------------------------------------------------
# ADMIN > Taux de TVA / Taux de retenue à la source
# ---------------------------------------------------------------------------
@app.route("/module/taux_tva", methods=["GET", "POST"])
@menu_requis("taux_tva")
def taux_tva():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_taux_tva(db, request.form["code"].strip(), request.form["label"].strip(),
                               montant=float(request.form.get("montant") or 0),
                               compte=request.form.get("compte") or None)
            flash("Taux de TVA enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("taux_tva"))
    return render_template("taux_tva.html", taux=core.list_taux_tva(db))


@app.route("/module/taux_tva/supprimer/<code>", methods=["POST"])
@menu_requis("taux_tva")
def taux_tva_supprimer(code):
    try:
        core.delete_taux_tva(get_db(), code)
        flash("Taux de TVA supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("taux_tva"))


@app.route("/module/taux_retenue", methods=["GET", "POST"])
@menu_requis("taux_retenue")
def taux_retenue():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_taux_retenue(db, request.form["code"].strip(), request.form["label"].strip(),
                                   montant=float(request.form.get("montant") or 0),
                                   compte=request.form.get("compte") or None)
            flash("Taux de retenue enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("taux_retenue"))
    return render_template("taux_retenue.html", taux=core.list_taux_retenue(db))


@app.route("/module/taux_retenue/supprimer/<code>", methods=["POST"])
@menu_requis("taux_retenue")
def taux_retenue_supprimer(code):
    try:
        core.delete_taux_retenue(get_db(), code)
        flash("Taux de retenue supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("taux_retenue"))


# ---------------------------------------------------------------------------
# ADMIN > Niveaux d'accès / Utilisateurs
# ---------------------------------------------------------------------------
@app.route("/module/niveaux_acces", methods=["GET", "POST"])
@menu_requis("niveaux_acces")
def niveaux_acces():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_niveau_acces(db, request.form["nom"].strip(), description=request.form.get("description", ""))
            flash("Niveau d'accès créé.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("niveaux_acces"))
    niveaux = core.list_niveaux_acces(db)
    menus_par_niveau = {n["nom"]: core.get_menus_autorises(db, n["nom"]) for n in niveaux}
    tous_les_menus = [(label, key) for _titre, items in core.MENU_STRUCTURE for label, key in items]
    return render_template("niveaux_acces.html", niveaux=niveaux, menus_par_niveau=menus_par_niveau,
                            tous_les_menus=tous_les_menus)


@app.route("/module/niveaux_acces/supprimer/<nom>", methods=["POST"])
@menu_requis("niveaux_acces")
def niveaux_acces_supprimer(nom):
    try:
        core.delete_niveau_acces(get_db(), nom)
        flash("Niveau d'accès supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("niveaux_acces"))


@app.route("/module/niveaux_acces/menus/<nom>", methods=["POST"])
@menu_requis("niveaux_acces")
def niveaux_acces_menus(nom):
    keys = request.form.getlist("menu_key")
    core.set_menus_autorises(get_db(), nom, keys)
    flash(f"Menus autorisés mis à jour pour « {nom} ».", "success")
    return redirect(url_for("niveaux_acces"))


@app.route("/module/utilisateurs", methods=["GET", "POST"])
@menu_requis("utilisateurs")
def utilisateurs():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_utilisateur(db, request.form["nom_utilisateur"].strip(), request.form["mot_de_passe"],
                                  nom_complet=request.form.get("nom_complet", ""),
                                  niveau_acces=request.form.get("niveau_acces", "Lecture seule"),
                                  actif=bool(request.form.get("actif")))
            flash("Utilisateur créé.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("utilisateurs"))
    return render_template("utilisateurs.html", utilisateurs=core.list_utilisateurs(db),
                            niveaux=core.list_niveaux_acces(db))


@app.route("/module/utilisateurs/supprimer/<int:user_id>", methods=["POST"])
@menu_requis("utilisateurs")
def utilisateurs_supprimer(user_id):
    if user_id == session["user"]["id"]:
        flash("Vous ne pouvez pas supprimer votre propre compte.", "error")
        return redirect(url_for("utilisateurs"))
    try:
        core.delete_utilisateur(get_db(), user_id)
        flash("Utilisateur supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("utilisateurs"))


@app.route("/module/utilisateurs/reinitialiser/<int:user_id>", methods=["POST"])
@menu_requis("utilisateurs")
def utilisateurs_reinitialiser(user_id):
    nouveau_mdp = request.form.get("nouveau_mot_de_passe", "")
    if not nouveau_mdp:
        flash("Nouveau mot de passe requis.", "error")
    else:
        core.update_utilisateur(get_db(), user_id, nouveau_mot_de_passe=nouveau_mdp)
        flash("Mot de passe réinitialisé.", "success")
    return redirect(url_for("utilisateurs"))


# ---------------------------------------------------------------------------
# PRODUCTION > Fabrication (recettes / nomenclature + coût de production)
# ---------------------------------------------------------------------------
@app.route("/module/production", methods=["GET", "POST"])
@menu_requis("production")
def production():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_produit_fini(
                db, request.form["code"].strip(), request.form["nom"].strip(),
                description=request.form.get("description", ""),
                quantite_produite=float(request.form.get("quantite_produite") or 1),
                marge_pourcentage=float(request.form.get("marge_pourcentage") or 30),
                compte_stock=request.form.get("compte_stock") or "360000",
            )
            flash("Produit fini enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("production"))
    return render_template("production.html", produits=core.list_produits_finis(db))


@app.route("/module/production/<code>/supprimer", methods=["POST"])
@menu_requis("production")
def production_supprimer(code):
    try:
        core.delete_produit_fini(get_db(), code)
        flash("Produit fini supprimé.", "success")
    except Exception as exc:
        flash(f"Suppression impossible : {exc}", "error")
    return redirect(url_for("production"))


@app.route("/module/production/<code>", methods=["GET", "POST"])
@menu_requis("production")
def production_detail(code):
    db = get_db()
    exercice = exercice_actif(db)
    produit = core.get_produit_fini(db, code)
    if not produit:
        flash("Produit introuvable.", "error")
        return redirect(url_for("production"))

    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "ajouter_ligne":
                core.add_recette_ligne(
                    db, code, request.form["type_ligne"], request.form["libelle"],
                    float(request.form["quantite"]), compte=request.form.get("compte") or None,
                    cout_unitaire=float(request.form["cout_unitaire"]) if request.form.get("cout_unitaire") else None,
                    analytic_code=request.form.get("analytic_code") or None,
                )
                flash("Ligne de recette ajoutée.", "success")
            elif action == "produire":
                resultat, warnings = core.valider_fabrication(db, code, exercice=exercice)
                flash(f"Fabrication comptabilisée — {resultat['quantite_produite']} unité(s) produite(s).", "success")
                for w in warnings:
                    flash(w, "error")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("production_detail", code=code, exercice=exercice))

    cout = core.compute_cout_production(db, code, exercice=exercice)
    return render_template("production_detail.html", produit=produit, cout=cout,
                            ligne_types=core.LIGNE_TYPES, exercice=exercice,
                            exercices=core.list_exercices(db))


@app.route("/module/production/<code>/ligne/<int:ligne_id>/supprimer", methods=["POST"])
@menu_requis("production")
def production_ligne_supprimer(ligne_id, code):
    core.delete_recette_ligne(get_db(), ligne_id)
    flash("Ligne supprimée.", "success")
    return redirect(url_for("production_detail", code=code))


# ---------------------------------------------------------------------------
# TRESORERIE
# ---------------------------------------------------------------------------
@app.route("/module/tresorerie")
@menu_requis("tresorerie")
def tresorerie():
    db = get_db()
    exercice = exercice_actif(db)
    date_from = request.args.get("date_from") or f"{exercice}-01-01"
    date_to = request.args.get("date_to") or f"{exercice}-12-31"
    lignes, totaux = core.compute_tresorerie_banques_horizontal(db, date_from=date_from, date_to=date_to,
                                                                  exercice=exercice)
    engagements = core.compute_engagements_a_payer(db)
    return render_template("tresorerie.html", lignes=lignes, totaux=totaux, engagements=engagements,
                            exercice=exercice, exercices=core.list_exercices(db),
                            date_from=date_from, date_to=date_to)


# ---------------------------------------------------------------------------
# ENGAGEMENTS-PROJETS > Contrats (suivi des délais fournisseurs)
# ---------------------------------------------------------------------------
@app.route("/module/contrats", methods=["GET", "POST"])
@menu_requis("contrats")
def contrats():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_commande(
                db, request.form["fournisseur_code"], request.form.get("piece", ""),
                request.form.get("libelle", ""), float(request.form.get("montant") or 0),
                request.form["date_commande"],
            )
            flash("Commande/contrat enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("contrats"))
    return render_template("contrats.html", commandes=core.list_commandes(db), fournisseurs=core.list_fournisseurs(db),
                            today=core.date.today().isoformat())


@app.route("/module/contrats/<int:commande_id>/livraison", methods=["POST"])
@menu_requis("contrats")
def contrats_livraison(commande_id):
    core.update_commande(get_db(), commande_id, date_livraison_reelle=request.form.get("date_livraison_reelle"))
    flash("Date de livraison réelle enregistrée.", "success")
    return redirect(url_for("contrats"))


@app.route("/module/contrats/<int:commande_id>/paiement", methods=["POST"])
@menu_requis("contrats")
def contrats_paiement(commande_id):
    core.update_commande(get_db(), commande_id, date_paiement_reel=request.form.get("date_paiement_reel"))
    flash("Date de paiement réelle enregistrée.", "success")
    return redirect(url_for("contrats"))


@app.route("/module/contrats/<int:commande_id>/supprimer", methods=["POST"])
@menu_requis("contrats")
def contrats_supprimer(commande_id):
    core.delete_commande(get_db(), commande_id)
    flash("Commande supprimée.", "success")
    return redirect(url_for("contrats"))


# ---------------------------------------------------------------------------
# ENGAGEMENTS-PROJETS > Expression de besoin
# ---------------------------------------------------------------------------
@app.route("/module/expression_besoin", methods=["GET", "POST"])
@menu_requis("expression_besoin")
def expression_besoin():
    db = get_db()
    if request.method == "POST":
        try:
            eid = core.create_expression_besoin(
                db, request.form["numero"], request.form["date_demande"],
                demandeur=request.form.get("demandeur", ""), service=request.form.get("service", ""),
            )
            flash("Expression de besoin créée.", "success")
            return redirect(url_for("expression_besoin_detail", expression_id=eid))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("expression_besoin"))
    return render_template("expression_besoin.html", liste=core.list_expressions_besoin(db),
                            today=core.date.today().isoformat())


@app.route("/module/expression_besoin/<int:expression_id>", methods=["GET", "POST"])
@menu_requis("expression_besoin")
def expression_besoin_detail(expression_id):
    db = get_db()
    exp = core.get_expression_besoin(db, expression_id)
    if not exp:
        flash("Expression de besoin introuvable.", "error")
        return redirect(url_for("expression_besoin"))
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "ajouter_ligne":
                core.add_ligne_expression_besoin(db, expression_id, request.form["libelle"],
                                                  float(request.form["quantite"]), unite=request.form.get("unite"))
                flash("Ligne ajoutée.", "success")
            elif action == "valider":
                bon_id = core.valider_expression_besoin(db, expression_id)
                flash("Expression validée — basculée en Bon de commande.", "success")
                return redirect(url_for("ep_bon_commande_detail", bon_id=bon_id))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("expression_besoin_detail", expression_id=expression_id))
    return render_template("expression_besoin_detail.html", exp=exp,
                            lignes=core.list_lignes_expression_besoin(db, expression_id))


@app.route("/module/expression_besoin/<int:expression_id>/ligne/<int:ligne_id>/supprimer", methods=["POST"])
@menu_requis("expression_besoin")
def expression_besoin_ligne_supprimer(expression_id, ligne_id):
    core.delete_ligne_expression_besoin(get_db(), ligne_id)
    flash("Ligne supprimée.", "success")
    return redirect(url_for("expression_besoin_detail", expression_id=expression_id))


# ---------------------------------------------------------------------------
# ENGAGEMENTS-PROJETS > Bon de commande (comptabilise l'achat à la validation)
# ---------------------------------------------------------------------------
@app.route("/module/ep_bon_commande", methods=["GET", "POST"])
@menu_requis("ep_bon_commande")
def ep_bon_commande():
    db = get_db()
    if request.method == "POST":
        try:
            bid = core.create_ep_bon_commande(
                db, request.form["numero"], request.form["date_commande"],
                fournisseur_code=request.form.get("fournisseur_code", ""),
            )
            flash("Bon de commande créé en brouillon.", "success")
            return redirect(url_for("ep_bon_commande_detail", bon_id=bid))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("ep_bon_commande"))
    return render_template("ep_bon_commande.html", liste=core.list_ep_bons_commande(db),
                            fournisseurs=core.list_fournisseurs(db), today=core.date.today().isoformat())


@app.route("/module/ep_bon_commande/<int:bon_id>", methods=["GET", "POST"])
@menu_requis("ep_bon_commande")
def ep_bon_commande_detail(bon_id):
    db = get_db()
    bon = core.get_ep_bon_commande(db, bon_id)
    if not bon:
        flash("Bon de commande introuvable.", "error")
        return redirect(url_for("ep_bon_commande"))
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "ajouter_ligne":
                core.add_ligne_ep_bon_commande(
                    db, bon_id, request.form["libelle"], float(request.form["quantite"]),
                    prix_unitaire=float(request.form.get("prix_unitaire") or 0),
                    unite=request.form.get("unite"), compte_charge=request.form.get("compte_charge") or None,
                    analytic_code=request.form.get("analytic_code") or None,
                )
                flash("Ligne ajoutée.", "success")
            elif action == "valider":
                bordereau_id, reglement_id = core.valider_ep_bon_commande(db, bon_id)
                flash("Bon de commande validé — achat comptabilisé, Bordereau et Règlement générés.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("ep_bon_commande_detail", bon_id=bon_id))
    lignes = core.list_lignes_ep_bon_commande(db, bon_id)
    totals = core.compute_ep_bon_commande_totals(db, bon_id)
    return render_template("ep_bon_commande_detail.html", bon=bon, lignes=lignes, totals=totals)


@app.route("/module/ep_bon_commande/<int:bon_id>/ligne/<int:ligne_id>/supprimer", methods=["POST"])
@menu_requis("ep_bon_commande")
def ep_bon_commande_ligne_supprimer(bon_id, ligne_id):
    core.delete_ligne_ep_bon_commande(get_db(), ligne_id)
    flash("Ligne supprimée.", "success")
    return redirect(url_for("ep_bon_commande_detail", bon_id=bon_id))


# ---------------------------------------------------------------------------
# ENGAGEMENTS-PROJETS > Bordereau de livraison
# ---------------------------------------------------------------------------
@app.route("/module/bordereau_livraison")
@menu_requis("bordereau_livraison")
def bordereau_livraison():
    return render_template("bordereau_livraison.html", liste=core.list_bordereaux_livraison(get_db()))


@app.route("/module/bordereau_livraison/<int:bordereau_id>", methods=["GET", "POST"])
@menu_requis("bordereau_livraison")
def bordereau_livraison_detail(bordereau_id):
    db = get_db()
    bordereau = core.get_bordereau_livraison(db, bordereau_id)
    if not bordereau:
        flash("Bordereau introuvable.", "error")
        return redirect(url_for("bordereau_livraison"))
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "maj_quantite":
                core.update_ligne_bordereau_livraison(db, int(request.form["ligne_id"]),
                                                        float(request.form["quantite_livree"]))
                flash("Quantité livrée mise à jour.", "success")
            elif action == "valider":
                core.valider_bordereau_livraison(db, bordereau_id)
                flash("Bordereau validé (réception confirmée).", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("bordereau_livraison_detail", bordereau_id=bordereau_id))
    lignes = core.list_lignes_bordereau_livraison(db, bordereau_id)
    return render_template("bordereau_livraison_detail.html", bordereau=bordereau, lignes=lignes)


# ---------------------------------------------------------------------------
# ENGAGEMENTS-PROJETS > Règlements (comptabilise si créé hors Bon de commande)
# ---------------------------------------------------------------------------
@app.route("/module/reglements", methods=["GET", "POST"])
@menu_requis("reglements")
def reglements():
    db = get_db()
    if request.method == "POST":
        try:
            rid = core.create_reglement(
                db, request.form["numero"], request.form["date_reglement"],
                fournisseur_code=request.form.get("fournisseur_code", ""),
            )
            flash("Règlement créé en brouillon.", "success")
            return redirect(url_for("reglements_detail", reglement_id=rid))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("reglements"))
    return render_template("reglements.html", liste=core.list_reglements(db), fournisseurs=core.list_fournisseurs(db),
                            today=core.date.today().isoformat())


@app.route("/module/reglements/<int:reglement_id>", methods=["GET", "POST"])
@menu_requis("reglements")
def reglements_detail(reglement_id):
    db = get_db()
    reglement = core.get_reglement(db, reglement_id)
    if not reglement:
        flash("Règlement introuvable.", "error")
        return redirect(url_for("reglements"))
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "ajouter_ligne":
                core.add_ligne_reglement(
                    db, reglement_id, request.form.get("compte_charge") or None, request.form["libelle"],
                    float(request.form["quantite"]), prix_unitaire=float(request.form.get("prix_unitaire") or 0),
                    analytic_code=request.form.get("analytic_code") or None,
                )
                flash("Ligne ajoutée.", "success")
            elif action == "valider":
                core.valider_reglement(db, reglement_id)
                flash("Règlement validé — charge comptabilisée.", "success")
            elif action == "payer":
                montant = core.enregistrer_paiement_reglement(
                    db, reglement_id, request.form["date_paiement"], request.form["compte_paiement"])
                flash(f"Paiement de {montant:,.0f} F CFA comptabilisé.".replace(",", " "), "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("reglements_detail", reglement_id=reglement_id))
    lignes = core.list_lignes_reglement(db, reglement_id)
    totals = core.compute_reglement_totals(db, reglement_id)
    return render_template("reglements_detail.html", reglement=reglement, lignes=lignes, totals=totals)


@app.route("/module/reglements/<int:reglement_id>/ligne/<int:ligne_id>/supprimer", methods=["POST"])
@menu_requis("reglements")
def reglements_ligne_supprimer(reglement_id, ligne_id):
    core.delete_ligne_reglement(get_db(), ligne_id)
    flash("Ligne supprimée.", "success")
    return redirect(url_for("reglements_detail", reglement_id=reglement_id))


# ---------------------------------------------------------------------------
# COMMERCIAL > Recouvrement (factures clients simples, hors Facturation TVA)
# ---------------------------------------------------------------------------
@app.route("/module/recouvrement", methods=["GET", "POST"])
@menu_requis("recouvrement")
def recouvrement():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_facture(
                db, request.form["client_code"], request.form.get("piece", ""), request.form.get("libelle", ""),
                float(request.form["montant"]), request.form["date_facture"],
            )
            flash("Facture enregistrée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("recouvrement"))
    return render_template("recouvrement.html", factures=core.list_factures(db), clients=core.list_clients(db),
                            today=core.date.today().isoformat())


@app.route("/module/recouvrement/<int:facture_id>/payer", methods=["POST"])
@menu_requis("recouvrement")
def recouvrement_payer(facture_id):
    try:
        core.enregistrer_paiement_facture(get_db(), facture_id, request.form["date_paiement_reel"],
                                           request.form["compte_reglement"])
        flash("Paiement comptabilisé.", "success")
    except (ValueError, KeyError) as exc:
        flash(str(exc), "error")
    return redirect(url_for("recouvrement"))


@app.route("/module/recouvrement/<int:facture_id>/supprimer", methods=["POST"])
@menu_requis("recouvrement")
def recouvrement_supprimer(facture_id):
    core.delete_facture(get_db(), facture_id)
    flash("Facture supprimée.", "success")
    return redirect(url_for("recouvrement"))


# ---------------------------------------------------------------------------
# COMMERCIAL > Marges bénéficiaires
# ---------------------------------------------------------------------------
@app.route("/module/marges")
@menu_requis("marges")
def marges():
    db = get_db()
    exercice = exercice_actif(db)
    cr = core.compute_liasse_resultat(db, exercice=exercice)
    return render_template("marges.html", cr=cr, exercice=exercice, exercices=core.list_exercices(db))


# ---------------------------------------------------------------------------
# ADMIN > Modification des factures (dévalidation consolidée vente + achat)
# ---------------------------------------------------------------------------
@app.route("/module/admin_factures")
@menu_requis("admin_factures")
def admin_factures():
    db = get_db()
    factures = (
        [{"type": "vente", "id": f["id"], "numero": f["numero"], "date": f["date_facture"],
          "tiers": f["raison_sociale"]} for f in core.list_factures_vente(db) if f["statut"] == "validee"]
        + [{"type": "achat", "id": f["id"], "numero": f["numero"], "date": f["date_facture"],
            "tiers": f["raison_sociale"]} for f in core.list_factures_achat(db) if f["statut"] == "validee"]
    )
    return render_template("admin_factures.html", factures=factures)


@app.route("/module/admin_factures/devalider", methods=["POST"])
@menu_requis("admin_factures")
def admin_factures_devalider():
    db = get_db()
    type_facture = request.form.get("type")
    facture_id = int(request.form.get("id"))
    try:
        if type_facture == "vente":
            core.devalider_facture_vente(db, facture_id)
        else:
            core.devalider_facture_achat(db, facture_id)
        flash("Facture dévalidée — repassée en brouillon, écritures retirées.", "success")
    except (ValueError, KeyError) as exc:
        flash(str(exc), "error")
    return redirect(url_for("admin_factures"))


# ---------------------------------------------------------------------------
# ADMIN > Modèle de bon de commande (en-tête/pied de page par défaut)
# ---------------------------------------------------------------------------
@app.route("/module/admin_modele_bon_commande", methods=["GET", "POST"])
@menu_requis("admin_modele_bon_commande")
def admin_modele_bon_commande():
    db = get_db()
    if request.method == "POST":
        core.set_text_setting(db, "bon_commande_entete_defaut", request.form.get("entete", "").strip())
        core.set_text_setting(db, "bon_commande_pied_defaut", request.form.get("pied", "").strip())
        flash("Modèle de bon de commande enregistré.", "success")
        return redirect(url_for("admin_modele_bon_commande"))
    entete = core.get_text_setting(db, "bon_commande_entete_defaut", "")
    pied = core.get_text_setting(db, "bon_commande_pied_defaut", "")
    return render_template("admin_modele_bon_commande.html", entete=entete, pied=pied)


# ---------------------------------------------------------------------------
# ADMIN > Réinitialisation des données (réservé à l'administrateur)
# ---------------------------------------------------------------------------
@app.route("/module/reinitialisation", methods=["GET", "POST"])
@menu_requis("reinitialisation")
def reinitialisation():
    db = get_db()
    if session["user"]["niveau_acces"] != "Administrateur":
        flash("Réservé à l'administrateur.", "error")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        categories = request.form.getlist("categorie")
        confirmation = request.form.get("confirmation", "")
        if confirmation != "SUPPRIMER":
            flash("Tapez SUPPRIMER en majuscules pour confirmer.", "error")
        elif not categories:
            flash("Sélectionnez au moins une catégorie.", "error")
        else:
            exercice = request.form.get("exercice") or None
            rapport = core.reinitialiser_donnees(db, categories, exercice=exercice)
            details = " · ".join(f"{core.REINIT_CATEGORIES.get(k, k)} : {v} ligne(s)" for k, v in rapport.items())
            flash(f"Données réinitialisées — {details}", "success")
        return redirect(url_for("reinitialisation"))
    return render_template("reinitialisation.html", categories=core.REINIT_CATEGORIES,
                            exercices=core.list_exercices(db))


# ---------------------------------------------------------------------------
# GRH > Time sheet
# ---------------------------------------------------------------------------
@app.route("/module/grh_time_sheet", methods=["GET", "POST"])
@menu_requis("grh_time_sheet")
def grh_time_sheet():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_time_sheet(db, int(request.form["personnel_id"]), request.form["date_pointage"],
                                 float(request.form["heures"]), activite=request.form.get("activite", ""),
                                 notes=request.form.get("notes", ""))
            flash("Pointage enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_time_sheet"))
    return render_template("grh_time_sheet.html", pointages=core.list_time_sheet(db),
                            personnel=core.list_personnel(db, actifs_only=True),
                            today=core.date.today().isoformat())


@app.route("/module/grh_time_sheet/<int:ts_id>/supprimer", methods=["POST"])
@menu_requis("grh_time_sheet")
def grh_time_sheet_supprimer(ts_id):
    core.delete_time_sheet(get_db(), ts_id)
    flash("Pointage supprimé.", "success")
    return redirect(url_for("grh_time_sheet"))


# ---------------------------------------------------------------------------
# GRH > KPI
# ---------------------------------------------------------------------------
@app.route("/module/grh_kpi", methods=["GET", "POST"])
@menu_requis("grh_kpi")
def grh_kpi():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_kpi(
                db, request.form["indicateur"], description=request.form.get("description", ""),
                personnel_id=int(request.form["personnel_id"]) if request.form.get("personnel_id") else None,
                service=request.form.get("service", ""), periode=request.form.get("periode", ""),
                valeur_cible=float(request.form.get("valeur_cible") or 0),
                valeur_realisee=float(request.form.get("valeur_realisee") or 0),
                unite=request.form.get("unite", ""), statut=request.form.get("statut", "en_cours"),
            )
            flash("KPI enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_kpi"))
    return render_template("grh_kpi.html", kpis=core.list_kpi(db), personnel=core.list_personnel(db))


@app.route("/module/grh_kpi/<int:kpi_id>/statut", methods=["POST"])
@menu_requis("grh_kpi")
def grh_kpi_statut(kpi_id):
    core.update_kpi(get_db(), kpi_id, statut=request.form["statut"],
                     valeur_realisee=float(request.form.get("valeur_realisee") or 0))
    flash("KPI mis à jour.", "success")
    return redirect(url_for("grh_kpi"))


@app.route("/module/grh_kpi/<int:kpi_id>/supprimer", methods=["POST"])
@menu_requis("grh_kpi")
def grh_kpi_supprimer(kpi_id):
    core.delete_kpi(get_db(), kpi_id)
    flash("KPI supprimé.", "success")
    return redirect(url_for("grh_kpi"))


# ---------------------------------------------------------------------------
# GRH > HS (hygiène santé)
# ---------------------------------------------------------------------------
@app.route("/module/grh_hs", methods=["GET", "POST"])
@menu_requis("grh_hs")
def grh_hs():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_hs(
                db, request.form["date_evenement"], type_evenement=request.form.get("type_evenement", "incident"),
                personnel_id=int(request.form["personnel_id"]) if request.form.get("personnel_id") else None,
                description=request.form.get("description", ""), gravite=request.form.get("gravite", ""),
                statut=request.form.get("statut", "ouvert"),
            )
            flash("Événement HS enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("grh_hs"))
    return render_template("grh_hs.html", liste=core.list_hs(db), personnel=core.list_personnel(db),
                            today=core.date.today().isoformat())


@app.route("/module/grh_hs/<int:hs_id>/statut", methods=["POST"])
@menu_requis("grh_hs")
def grh_hs_statut(hs_id):
    core.update_hs(get_db(), hs_id, statut=request.form["statut"])
    flash("Statut mis à jour.", "success")
    return redirect(url_for("grh_hs"))


@app.route("/module/grh_hs/<int:hs_id>/supprimer", methods=["POST"])
@menu_requis("grh_hs")
def grh_hs_supprimer(hs_id):
    core.delete_hs(get_db(), hs_id)
    flash("Événement supprimé.", "success")
    return redirect(url_for("grh_hs"))


# ---------------------------------------------------------------------------
# GRH > Tableau de bord GRH (synthèse lecture seule)
# ---------------------------------------------------------------------------
@app.route("/module/grh_tableau_bord")
@menu_requis("grh_tableau_bord")
def grh_tableau_bord():
    return render_template("grh_tableau_bord.html", d=core.compute_tableau_bord_grh(get_db()))


# ---------------------------------------------------------------------------
# TRANSPORT > Parc auto
# ---------------------------------------------------------------------------
@app.route("/module/transport", methods=["GET", "POST"])
@menu_requis("transport")
def transport():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_vehicule(
                db, request.form["immatriculation"], marque=request.form.get("marque", ""),
                modele=request.form.get("modele", ""), type_vehicule=request.form.get("type_vehicule", ""),
                date_acquisition=request.form.get("date_acquisition", ""),
                chauffeur_affecte=request.form.get("chauffeur_affecte", ""),
                statut=request.form.get("statut", "actif"),
            )
            flash("Véhicule enregistré.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("transport"))
    return render_template("transport.html", vehicules=core.list_vehicules(db))


@app.route("/module/transport/<int:vehicule_id>/supprimer", methods=["POST"])
@menu_requis("transport")
def transport_supprimer(vehicule_id):
    core.delete_vehicule(get_db(), vehicule_id)
    flash("Véhicule supprimé.", "success")
    return redirect(url_for("transport"))


# ---------------------------------------------------------------------------
# TRANSPORT > Missions
# ---------------------------------------------------------------------------
@app.route("/module/missions", methods=["GET", "POST"])
@menu_requis("missions")
def missions():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_mission(
                db, request.form["destination"],
                vehicule_id=int(request.form["vehicule_id"]) if request.form.get("vehicule_id") else None,
                chauffeur=request.form.get("chauffeur", ""), motif=request.form.get("motif", ""),
                date_depart=request.form.get("date_depart", ""), date_retour=request.form.get("date_retour", ""),
                km_depart=float(request.form["km_depart"]) if request.form.get("km_depart") else None,
                km_retour=float(request.form["km_retour"]) if request.form.get("km_retour") else None,
            )
            flash("Mission enregistrée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("missions"))
    return render_template("missions.html", missions=core.list_missions(db), vehicules=core.list_vehicules(db))


@app.route("/module/missions/<int:mission_id>/statut", methods=["POST"])
@menu_requis("missions")
def missions_statut(mission_id):
    core.update_mission(get_db(), mission_id, statut=request.form["statut"])
    flash("Statut mis à jour.", "success")
    return redirect(url_for("missions"))


@app.route("/module/missions/<int:mission_id>/supprimer", methods=["POST"])
@menu_requis("missions")
def missions_supprimer(mission_id):
    core.delete_mission(get_db(), mission_id)
    flash("Mission supprimée.", "success")
    return redirect(url_for("missions"))


# ---------------------------------------------------------------------------
# TRANSPORT/MAINTENANCE-QUALITÉ > Pièces de rechange (stock partagé)
# ---------------------------------------------------------------------------
@app.route("/module/pieces_rechange", methods=["GET", "POST"])
@menu_requis("pieces_rechange")
def pieces_rechange():
    db = get_db()
    if request.method == "POST":
        try:
            core.add_piece_rechange(
                db, request.form["designation"], code=request.form.get("code", ""),
                quantite_stock=float(request.form.get("quantite_stock") or 0), unite=request.form.get("unite", ""),
                cout_unitaire=float(request.form.get("cout_unitaire") or 0),
                fournisseur_code=request.form.get("fournisseur_code", ""),
            )
            flash("Pièce de rechange enregistrée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("pieces_rechange"))
    return render_template("pieces_rechange.html", pieces=core.list_pieces_rechange(db),
                            fournisseurs=core.list_fournisseurs(db))


@app.route("/module/pieces_rechange/<int:piece_id>/supprimer", methods=["POST"])
@menu_requis("pieces_rechange")
def pieces_rechange_supprimer(piece_id):
    core.delete_piece_rechange(get_db(), piece_id)
    flash("Pièce supprimée.", "success")
    return redirect(url_for("pieces_rechange"))


# ---------------------------------------------------------------------------
# TRANSPORT > Réparations (décrémente automatiquement le stock de pièces)
# ---------------------------------------------------------------------------
@app.route("/module/reparations", methods=["GET", "POST"])
@menu_requis("reparations")
def reparations():
    db = get_db()
    if request.method == "POST":
        try:
            core.create_reparation(
                db, request.form["description"],
                vehicule_id=int(request.form["vehicule_id"]) if request.form.get("vehicule_id") else None,
                date_reparation=request.form.get("date_reparation") or None, garage=request.form.get("garage", ""),
                cout_main_oeuvre=float(request.form.get("cout_main_oeuvre") or 0),
            )
            flash("Réparation créée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("reparations"))
    return render_template("reparations.html", liste=core.list_reparations(db), vehicules=core.list_vehicules(db),
                            today=core.date.today().isoformat())


@app.route("/module/reparations/<int:reparation_id>", methods=["GET", "POST"])
@menu_requis("reparations")
def reparations_detail(reparation_id):
    db = get_db()
    reparation = core.get_reparation(db, reparation_id)
    if not reparation:
        flash("Réparation introuvable.", "error")
        return redirect(url_for("reparations"))
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "ajouter_piece":
                core.add_ligne_reparation(db, reparation_id, int(request.form["piece_id"]),
                                           quantite=float(request.form["quantite"]))
                flash("Pièce ajoutée — stock décrémenté.", "success")
            elif action == "terminer":
                core.update_reparation(db, reparation_id, statut="terminee")
                flash("Réparation marquée terminée.", "success")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "error")
        return redirect(url_for("reparations_detail", reparation_id=reparation_id))
    lignes = core.list_lignes_reparation(db, reparation_id)
    cout_total = core.compute_cout_total_reparation(db, reparation_id)
    return render_template("reparations_detail.html", reparation=reparation, lignes=lignes, cout_total=cout_total,
                            pieces=core.list_pieces_rechange(db))


@app.route("/module/reparations/<int:reparation_id>/ligne/<int:ligne_id>/supprimer", methods=["POST"])
@menu_requis("reparations")
def reparations_ligne_supprimer(reparation_id, ligne_id):
    core.delete_ligne_reparation(get_db(), ligne_id)
    flash("Ligne supprimée — stock restitué.", "success")
    return redirect(url_for("reparations_detail", reparation_id=reparation_id))


# ---------------------------------------------------------------------------
# MAINTENANCE-QUALITÉ > Énergie / Maintenance (même écran générique que le
# bureau : coûts par code analytique, alimentés automatiquement par Saisie
# et Fabrication)
# ---------------------------------------------------------------------------
def _analytique_periode_view(prefix, suggestions, titre, description):
    db = get_db()
    date_from = request.args.get("date_from") or None
    date_to = request.args.get("date_to") or None
    if request.method == "POST" and request.form.get("action") == "ajouter_suggestions":
        n = core.ajouter_codes_analytiques_suggeres(db, suggestions)
        flash(f"{n} code(s) analytique(s) ajouté(s)." if n else "Tous les codes courants existent déjà.", "success")
        return redirect(request.path)
    codes = core.compute_couts_analytiques_categorie(db, prefix, date_from=date_from, date_to=date_to)
    totaux = {
        "solde_debut_periode": sum(c["solde_debut_periode"] for c in codes),
        "debit_periode": sum(c["debit_periode"] for c in codes),
        "credit_periode": sum(c["credit_periode"] for c in codes),
        "solde_fin_periode": sum(c["solde_fin_periode"] for c in codes),
    }
    return render_template("analytique_periode.html", codes=codes, totaux=totaux, titre=titre,
                            description=description, date_from=date_from, date_to=date_to,
                            exercice=core.get_current_exercice(db))


@app.route("/module/energie", methods=["GET", "POST"])
@menu_requis("energie")
def energie():
    return _analytique_periode_view(
        core.PREFIX_ENERGIE, core.SUGGESTIONS_ENERGIE, "Énergie",
        "Coûts d'énergie (eau, électricité, essence, gasoil, gaz...) par code analytique, sur une "
        "période choisie — alimentés par les écritures de Saisie taguées avec un code « ENERGIE- » "
        "et par les lignes « Énergie » des recettes de Fabrication.")


@app.route("/module/maintenance", methods=["GET", "POST"])
@menu_requis("maintenance")
def maintenance():
    return _analytique_periode_view(
        core.PREFIX_MAINTENANCE, core.SUGGESTIONS_MAINTENANCE, "Maintenance",
        "Coûts de maintenance (véhicules, bâtiments, machines, informatique...) par code analytique, "
        "sur une période choisie — alimentés par les écritures de Saisie taguées avec un code "
        "« MAINT- » et par les lignes « Autre charge » des recettes de Fabrication qui leur sont associées.")


# ---------------------------------------------------------------------------
# IMPORT / EXPORT EXCEL — un export + un import par écran concerné
# ---------------------------------------------------------------------------

# ---- Plan comptable ----
@app.route("/module/plan_comptable/export")
@menu_requis("plan_comptable")
def plan_comptable_export():
    return export_xlsx_response(core.export_plan_comptable_xlsx, "plan_comptable.xlsx", get_db())


@app.route("/module/plan_comptable/import", methods=["POST"])
@menu_requis("plan_comptable")
def plan_comptable_import():
    return handle_import_upload(core.import_plan_comptable_xlsx, get_db(), redirect_endpoint="plan_comptable")


# ---- Soldes d'ouverture ----
@app.route("/module/ouverture/export")
@menu_requis("ouverture")
def ouverture_export():
    db = get_db()
    exercice = exercice_actif(db)
    return export_xlsx_response(core.export_opening_balances_xlsx, f"soldes_ouverture_{exercice}.xlsx",
                                 db, exercice=exercice)


@app.route("/module/ouverture/import", methods=["POST"])
@menu_requis("ouverture")
def ouverture_import():
    db = get_db()
    exercice = exercice_actif(db)
    return handle_import_upload(core.import_opening_balances_xlsx, db, redirect_endpoint="ouverture",
                                 exercice=exercice)


# ---- Plan analytique ----
@app.route("/module/plan_analytique/export")
@menu_requis("plan_analytique")
def plan_analytique_export():
    return export_xlsx_response(core.export_analytic_codes_xlsx, "plan_analytique.xlsx", get_db())


@app.route("/module/plan_analytique/import", methods=["POST"])
@menu_requis("plan_analytique")
def plan_analytique_import():
    return handle_import_upload(core.import_analytic_codes_xlsx, get_db(), redirect_endpoint="plan_analytique")


# ---- Plan budgétaire ----
@app.route("/module/plan_budgetaire/export")
@menu_requis("plan_budgetaire")
def plan_budgetaire_export():
    return export_xlsx_response(core.export_budget_codes_xlsx, "plan_budgetaire.xlsx", get_db())


@app.route("/module/plan_budgetaire/import", methods=["POST"])
@menu_requis("plan_budgetaire")
def plan_budgetaire_import():
    return handle_import_upload(core.import_budget_codes_xlsx, get_db(), redirect_endpoint="plan_budgetaire")


# ---- Plan bailleurs ----
@app.route("/module/plan_bailleur/export")
@menu_requis("plan_bailleur")
def plan_bailleur_export():
    return export_xlsx_response(core.export_donor_codes_xlsx, "plan_bailleurs.xlsx", get_db())


@app.route("/module/plan_bailleur/import", methods=["POST"])
@menu_requis("plan_bailleur")
def plan_bailleur_import():
    return handle_import_upload(core.import_donor_codes_xlsx, get_db(), redirect_endpoint="plan_bailleur")


# ---- Taux de TVA ----
@app.route("/module/taux_tva/export")
@menu_requis("taux_tva")
def taux_tva_export():
    return export_xlsx_response(core.export_taux_tva_xlsx, "taux_tva.xlsx", get_db())


@app.route("/module/taux_tva/import", methods=["POST"])
@menu_requis("taux_tva")
def taux_tva_import():
    return handle_import_upload(core.import_taux_tva_xlsx, get_db(), redirect_endpoint="taux_tva")


# ---- Taux de retenue ----
@app.route("/module/taux_retenue/export")
@menu_requis("taux_retenue")
def taux_retenue_export():
    return export_xlsx_response(core.export_taux_retenue_xlsx, "taux_retenue.xlsx", get_db())


@app.route("/module/taux_retenue/import", methods=["POST"])
@menu_requis("taux_retenue")
def taux_retenue_import():
    return handle_import_upload(core.import_taux_retenue_xlsx, get_db(), redirect_endpoint="taux_retenue")


# ---- Fournisseurs (modèle + import — pas d'export des données réelles) ----
@app.route("/module/fournisseurs/modele")
@menu_requis("fournisseurs")
def fournisseurs_modele():
    return export_xlsx_response(core.export_fournisseurs_template, "modele_fournisseurs.xlsx")


@app.route("/module/fournisseurs/import", methods=["POST"])
@menu_requis("fournisseurs")
def fournisseurs_import():
    return handle_import_upload(core.import_fournisseurs_from_xlsx, get_db(), redirect_endpoint="fournisseurs")


# ---- Clients (modèle + import) ----
@app.route("/module/clients/modele")
@menu_requis("clients")
def clients_modele():
    return export_xlsx_response(core.export_clients_template, "modele_clients.xlsx")


@app.route("/module/clients/import", methods=["POST"])
@menu_requis("clients")
def clients_import():
    return handle_import_upload(core.import_clients_from_xlsx, get_db(), redirect_endpoint="clients")


# ---- Saisie : import d'écritures ----
@app.route("/module/saisie/import", methods=["POST"])
@menu_requis("saisie")
def saisie_import():
    return handle_import_upload(core.import_entries_from_xlsx, get_db(), redirect_endpoint="saisie")


# ---- Balance (export seul) ----
@app.route("/module/balance/export")
@menu_requis("balance")
def balance_export():
    db = get_db()
    exercice = exercice_actif(db)
    return export_xlsx_response(core.export_balance_xlsx, f"balance_{exercice}.xlsx", db, exercice=exercice)


# ---- Bilan SYSCOHADA (export seul, 2 formats) ----
@app.route("/module/bilan_syscohada/export")
@menu_requis("bilan_syscohada")
def bilan_syscohada_export():
    db = get_db()
    exercice = exercice_actif(db)
    detaille = request.args.get("detaille") == "1"
    fn = core.export_bilan_detaille_xlsx if detaille else core.export_bilan_gabarit_xlsx
    suffix = "detaille" if detaille else "gabarit"
    return export_xlsx_response(fn, f"bilan_{suffix}_{exercice}.xlsx", db, exercice=exercice)


# ---- Immobilisations (import seul) ----
@app.route("/module/immobilisations/import", methods=["POST"])
@menu_requis("immobilisations")
def immobilisations_import():
    return handle_import_upload(core.import_immobilisations_from_xlsx, get_db(), redirect_endpoint="immobilisations")


# ---- Niveaux d'accès ----
@app.route("/module/niveaux_acces/export")
@menu_requis("niveaux_acces")
def niveaux_acces_export():
    return export_xlsx_response(core.export_niveaux_acces_xlsx, "niveaux_acces.xlsx", get_db())


@app.route("/module/niveaux_acces/import", methods=["POST"])
@menu_requis("niveaux_acces")
def niveaux_acces_import():
    return handle_import_upload(core.import_niveaux_acces_xlsx, get_db(), redirect_endpoint="niveaux_acces")


# ---- GRH Personnel : modèle + import ----
@app.route("/module/grh_personnel/modele")
@menu_requis("grh_personnel")
def grh_personnel_modele():
    return export_xlsx_response(lambda path: core.export_personnel_template_xlsx(path), "modele_personnel.xlsx")


@app.route("/module/grh_personnel/import", methods=["POST"])
@menu_requis("grh_personnel")
def grh_personnel_import():
    return handle_import_upload(core.import_personnel_xlsx, get_db(), redirect_endpoint="grh_personnel")


# ---- GRH Time sheet : modèle + import ----
@app.route("/module/grh_time_sheet/modele")
@menu_requis("grh_time_sheet")
def grh_time_sheet_modele():
    return export_xlsx_response(lambda path: core.export_time_sheet_template_xlsx(path), "modele_time_sheet.xlsx")


@app.route("/module/grh_time_sheet/import", methods=["POST"])
@menu_requis("grh_time_sheet")
def grh_time_sheet_import():
    return handle_import_upload(core.import_time_sheet_xlsx, get_db(), redirect_endpoint="grh_time_sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
