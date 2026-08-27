# -*- coding: utf-8 -*-
"""
SERVEUR — expose le moteur comptable (core.py) sur le réseau local ou
Internet, pour que plusieurs postes « client » (voir client_main.py)
puissent travailler EN MÊME TEMPS sur la même base de données.

Architecture :
- HTTP + JSON, avec la bibliothèque standard Python uniquement (aucune
  dépendance supplémentaire — cohérent avec le reste du projet).
- Une seule base SQLite partagée, en mode WAL (Write-Ahead Logging) pour
  de bonnes performances en lecture/écriture concurrente, protégée par un
  verrou global pour sérialiser les écritures (une opération métier à la
  fois — le moteur comptable n'a pas été conçu pour une écriture
  simultanée sur les mêmes lignes, donc cette sérialisation garantit la
  cohérence des données, au prix d'un débit plus faible qu'un vrai SGBD
  multi-utilisateur — largement suffisant pour une équipe).
- Authentification par utilisateur/mot de passe (table `utilisateurs`,
  déjà existante — voir core.verify_password()), avec un jeton de session
  à durée de vie limitée.
- Liste blanche explicite des fonctions core.py accessibles à distance
  (voir RPC_WHITELIST ci-dessous) — pour ne jamais exposer l'exécution de
  code arbitraire sur le serveur.

Lancement :
    python server.py [--port 8765] [--db chemin/vers/comptabilite.db]

IMPORTANT — sécurité réseau :
- Sur un réseau LOCAL (même bureau/même box internet), aucune configuration
  supplémentaire n'est nécessaire : les postes clients se connectent à
  l'adresse IP locale du serveur (ex. 192.168.1.10) sur le port choisi.
- Pour un accès depuis INTERNET (hors réseau local), il faut soit :
  (a) configurer une redirection de port (« port forwarding ») sur le
      routeur/box vers ce serveur, avec un mot de passe fort pour chaque
      utilisateur — le trafic reste alors en clair (HTTP), donc réservé à
      un usage de confiance (VPN recommandé en plus si possible) ; soit
  (b) placer le serveur derrière un VPN d'entreprise, solution la plus
      sûre pour un accès distant.
Ce serveur n'implémente PAS le chiffrement TLS/HTTPS par défaut — à
prévoir séparément (ex. reverse proxy) pour un usage sur Internet ouvert.
"""
import argparse
import json
import secrets
import sqlite3
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import core

# Numéro de version — change à CHAQUE fois que ce fichier est modifié.
# Permet de vérifier en un coup d'œil (affiché au démarrage ET dans
# /ping) que le serveur en cours d'exécution est bien la dernière
# version, sans avoir à deviner.
SERVER_VERSION = "2026-08-27-v1"

# ---------------------------------------------------------------------------
# SÉCURITÉ — modèle en LISTE NOIRE (pas liste blanche) : toute fonction
# publique de core.py prenant `conn` en premier argument est autorisée à
# distance PAR DÉFAUT, SAUF celles listées explicitement ci-dessous.
#
# Pourquoi ce choix : avec une liste BLANCHE, chaque nouvel écran construit
# côté client nécessite de mettre à jour ET reconstruire le serveur, avec un
# risque réel de désynchronisation entre les deux (déjà vécu concrètement :
# "Fonction non autorisée à distance" alors que le client était à jour mais
# pas le serveur). Avec une liste NOIRE, un nouvel écran fonctionne
# immédiatement dès que le client sait l'appeler, sans jamais retoucher au
# serveur — seules les quelques opérations réellement sensibles listées
# ci-dessous restent explicitement protégées.
# ---------------------------------------------------------------------------
RPC_BLOCKLIST = {
    # Authentification — déjà gérée en interne par /login, ne doit jamais
    # être appelable directement (permettrait de tester des mots de passe
    # sans passer par le mécanisme de session).
    "verify_password",
    # Gestion des utilisateurs et des niveaux d'accès — réservée à
    # l'application de bureau : un utilisateur distant ne doit jamais
    # pouvoir se créer un compte Administrateur ou modifier les
    # autorisations à distance (risque d'élévation de privilèges).
    "add_utilisateur", "update_utilisateur", "delete_utilisateur", "list_utilisateurs",
    "add_niveau_acces", "delete_niveau_acces", "set_menus_autorises",
    "ajouter_niveaux_acces_suggeres", "ajouter_niveaux_acces_suggeres_menus",
    # Opérations destructrices ou d'infrastructure — réservées au poste
    # serveur/bureau local.
    "reinitialiser_donnees", "init_db", "load_plan_comptable", "synchroniser_base",
    "get_app_icon_path",
}


def _construire_rpc_whitelist():
    """Construit dynamiquement l'ensemble des fonctions autorisées à
    distance : toutes les fonctions publiques de core.py prenant `conn` en
    premier argument, MOINS RPC_BLOCKLIST et les fonctions travaillant sur
    des fichiers locaux au serveur (export_*/import_*/generate_etat_xlsx —
    leurs chemins désignent le disque du serveur, sans signification
    pertinente pour un client distant tel quel)."""
    import inspect
    fonctions = set()
    for nom, obj in vars(core).items():
        if nom.startswith("_") or not inspect.isfunction(obj):
            continue
        if nom.startswith(("export_", "import_")) or nom == "generate_etat_xlsx":
            continue
        try:
            params = list(inspect.signature(obj).parameters)
        except (ValueError, TypeError):
            continue
        if params and params[0] == "conn":
            fonctions.add(nom)
    return fonctions - RPC_BLOCKLIST


RPC_WHITELIST = _construire_rpc_whitelist()

SESSION_DURATION_SECONDS = 8 * 3600  # 8h de travail avant reconnexion


class SessionStore:
    """Jetons de session en mémoire — {token: {"utilisateur":..., "niveau_acces":..., "expire":...}}."""

    def __init__(self):
        self._sessions = {}
        self._lock = threading.Lock()

    def create(self, nom_utilisateur, niveau_acces):
        token = secrets.token_hex(32)
        with self._lock:
            self._sessions[token] = {
                "utilisateur": nom_utilisateur,
                "niveau_acces": niveau_acces,
                "expire": time.time() + SESSION_DURATION_SECONDS,
            }
        return token

    def get(self, token):
        with self._lock:
            session = self._sessions.get(token)
            if not session:
                return None
            if session["expire"] < time.time():
                del self._sessions[token]
                return None
            return session

    def revoke(self, token):
        with self._lock:
            self._sessions.pop(token, None)


class AccountingServer:
    """Encapsule la connexion SQLite partagée (mode WAL) et le verrou
    global qui sérialise les écritures — un seul point d'accès à la base
    pour toutes les requêtes réseau, quel que soit le thread qui les traite."""

    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA foreign_keys=ON")
        core.init_db(self.conn)
        self.write_lock = threading.RLock()
        self.sessions = SessionStore()

    def call(self, function_name, args, kwargs):
        if function_name not in RPC_WHITELIST:
            raise PermissionError(f"Fonction « {function_name} » non autorisée à distance.")
        fn = getattr(core, function_name, None)
        if fn is None:
            raise AttributeError(f"Fonction « {function_name} » introuvable.")
        with self.write_lock:
            # Clôt toute transaction implicite en attente AVANT d'exécuter la
            # fonction — sans ça, une connexion SQLite longue durée en mode
            # WAL peut rester figée sur un instantané ancien de la base
            # (notamment sous Windows) tant qu'aucune transaction n'est
            # explicitement terminée, même si un AUTRE processus (l'application
            # de bureau) a bien validé des changements depuis. Sans effet si
            # rien n'était en attente (commit() est alors un no-op).
            self.conn.commit()
            return fn(self.conn, *args, **kwargs)


def _json_default(obj):
    """Sérialisation JSON pour les types non natifs (ex. sqlite3.Row déjà
    converti en dict par les fonctions core.py, mais par prudence)."""
    if isinstance(obj, (bytes, bytearray)):
        return obj.decode("utf-8", errors="replace")
    return str(obj)


def make_handler(server_state: AccountingServer):
    class Handler(BaseHTTPRequestHandler):
        server_version = "SaisieComptableServer/1.0"

        def log_message(self, format, *args):
            pass  # silencieux — évite de polluer la console ; activer si besoin de diagnostic

        def _send_json(self, status, payload):
            body = json.dumps(payload, default=_json_default, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _read_json_body(self):
            length = int(self.headers.get("Content-Length", 0) or 0)
            if length == 0:
                return {}
            raw = self.rfile.read(length)
            return json.loads(raw.decode("utf-8"))

        def do_POST(self):
            try:
                if self.path == "/login":
                    return self._handle_login()
                if self.path == "/rpc":
                    return self._handle_rpc()
                if self.path == "/logout":
                    return self._handle_logout()
                self._send_json(404, {"ok": False, "error": "Route inconnue."})
            except Exception as exc:
                self._send_json(500, {"ok": False, "error": f"Erreur serveur : {exc}"})

        def do_GET(self):
            if self.path == "/ping":
                return self._send_json(200, {
                    "ok": True, "message": "Serveur SaisieComptable actif.",
                    "version": SERVER_VERSION, "nb_fonctions_autorisees": len(RPC_WHITELIST),
                })
            self._send_json(404, {"ok": False, "error": "Route inconnue."})

        def _handle_login(self):
            data = self._read_json_body()
            nom_utilisateur = (data.get("nom_utilisateur") or "").strip()
            mot_de_passe = data.get("mot_de_passe") or ""
            if not nom_utilisateur or not mot_de_passe:
                return self._send_json(400, {"ok": False, "error": "Identifiant et mot de passe requis."})
            with server_state.write_lock:
                server_state.conn.commit()  # voir AccountingServer.call() -- force une lecture a jour (WAL)
                utilisateur = core.verify_password(server_state.conn, nom_utilisateur, mot_de_passe)
                if utilisateur:
                    menus_autorises = sorted(core.get_menus_autorises(server_state.conn, utilisateur.get("niveau_acces")))
            if not utilisateur:
                return self._send_json(401, {"ok": False, "error": "Identifiant ou mot de passe incorrect."})
            token = server_state.sessions.create(nom_utilisateur, utilisateur.get("niveau_acces"))
            self._send_json(200, {
                "ok": True,
                "session": token,
                "utilisateur": nom_utilisateur,
                "menus_autorises": menus_autorises,
                "niveau_acces": utilisateur.get("niveau_acces"),
                "version": SERVER_VERSION,
            })

        def _handle_logout(self):
            data = self._read_json_body()
            token = data.get("session")
            if token:
                server_state.sessions.revoke(token)
            self._send_json(200, {"ok": True})

        def _handle_rpc(self):
            data = self._read_json_body()
            token = data.get("session")
            session = server_state.sessions.get(token) if token else None
            if not session:
                return self._send_json(401, {"ok": False, "error": "Session expirée ou invalide — reconnectez-vous."})

            function_name = data.get("function")
            args = data.get("args") or []
            kwargs = data.get("kwargs") or {}
            try:
                result = server_state.call(function_name, args, kwargs)
            except PermissionError as exc:
                return self._send_json(403, {"ok": False, "error": str(exc)})
            except Exception as exc:
                return self._send_json(400, {"ok": False, "error": f"{type(exc).__name__} : {exc}"})
            self._send_json(200, {"ok": True, "result": result})

    return Handler


def run_server(db_path, host="0.0.0.0", port=8765):
    state = AccountingServer(db_path)
    handler_cls = make_handler(state)
    httpd = ThreadingHTTPServer((host, port), handler_cls)
    print("=" * 60)
    print(f"VERSION DU SERVEUR : {SERVER_VERSION}")
    print(f"Fonctions autorisées à distance : {len(RPC_WHITELIST)}")
    print("=" * 60)
    print(f"Serveur SaisieComptable démarré sur {host}:{port}")
    print(f"Base de données : {db_path}")
    print("Adresses locales possibles pour les postes clients : voir 'ipconfig' (Windows) sur cette machine.")
    print("Ctrl+C pour arrêter.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nArrêt du serveur.")
        httpd.shutdown()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Serveur SaisieComptable — accès réseau multi-utilisateur.")
    parser.add_argument("--port", type=int, default=8765, help="Port d'écoute (défaut : 8765)")
    parser.add_argument("--host", default="0.0.0.0", help="Adresse d'écoute (défaut : 0.0.0.0, toutes les interfaces)")
    parser.add_argument("--db", default=None, help="Chemin de la base de données (défaut : emplacement standard)")
    args = parser.parse_args()
    db_path = args.db or core.default_db_path()
    run_server(db_path, host=args.host, port=args.port)
