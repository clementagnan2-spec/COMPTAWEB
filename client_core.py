# -*- coding: utf-8 -*-
"""
CLIENT — module miroir de core.py côté client : au lieu d'exécuter les
fonctions comptables directement sur une base SQLite locale, chaque appel
est transformé en requête réseau (JSON/HTTP) vers le serveur (voir
server.py) qui exécute la VRAIE fonction core.py sur la base partagée.

Usage typique dans l'interface client (voir client_main.py) :

    remote = RemoteConnection("192.168.1.10", 8765)
    remote.login("alice", "motdepasse")
    # 'remote' se comporte alors comme la connexion `conn` habituelle :
    entries = client_core.list_entries(remote, exercice="2026")
    client_core.add_balanced_entry(remote, date, piece, journal, ...)

Aucune dépendance externe (bibliothèque standard uniquement — urllib).
"""
import json
import urllib.error
import urllib.request


class RemoteConnectionError(Exception):
    """Erreur de connexion réseau au serveur (serveur injoignable, etc.)."""
    pass


class RemoteAuthError(Exception):
    """Identifiants refusés, ou session expirée."""
    pass


class RemoteCallError(Exception):
    """La fonction a été exécutée côté serveur mais a levé une erreur
    métier (ex. ValueError d'une règle de gestion) — le message
    d'origine est conservé pour affichage à l'utilisateur."""
    pass


class RemoteConnection:
    """Remplace la connexion sqlite3 locale (`conn`) par une connexion
    réseau au serveur — objet passé en premier argument aux fonctions de
    ce module, exactement comme `conn` l'est aux fonctions de core.py."""

    def __init__(self, host, port, timeout=15):
        self.host = host
        self.port = port
        self.base_url = f"http://{host}:{port}"
        self.timeout = timeout
        self.session = None
        self.nom_utilisateur = None
        self.niveau_acces = None
        self.menus_autorises = set()
        self.server_version = None

    def ping(self):
        """Vérifie que le serveur répond, sans authentification — pour un
        test de connexion rapide avant de demander les identifiants.
        Renvoie le dict complet (avec "version") si joignable, sinon None."""
        try:
            with urllib.request.urlopen(f"{self.base_url}/ping", timeout=self.timeout) as resp:
                data = json.loads(resp.read())
            return data if data.get("ok") else None
        except (urllib.error.URLError, OSError, TimeoutError):
            return None

    def login(self, nom_utilisateur, mot_de_passe):
        status, data = self._post("/login", {
            "nom_utilisateur": nom_utilisateur, "mot_de_passe": mot_de_passe,
        })
        if status != 200 or not data.get("ok"):
            raise RemoteAuthError(data.get("error", "Échec de connexion."))
        self.session = data["session"]
        self.nom_utilisateur = data["utilisateur"]
        self.niveau_acces = data.get("niveau_acces")
        self.menus_autorises = set(data.get("menus_autorises") or [])
        self.server_version = data.get("version", "?")
        return data

    def logout(self):
        if self.session:
            try:
                self._post("/logout", {"session": self.session})
            except Exception:
                pass  # sans conséquence — la session expirera d'elle-même
        self.session = None

    def call(self, function_name, *args, **kwargs):
        if not self.session:
            raise RemoteAuthError("Non connecté — veuillez vous reconnecter.")
        status, data = self._post("/rpc", {
            "session": self.session, "function": function_name,
            "args": list(args), "kwargs": kwargs,
        })
        if status == 401:
            self.session = None
            raise RemoteAuthError(data.get("error", "Session expirée — reconnectez-vous."))
        if not data.get("ok"):
            raise RemoteCallError(data.get("error", "Erreur serveur inconnue."))
        return data.get("result")

    def _post(self, path, payload):
        body = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{self.base_url}{path}", data=body,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                return resp.status, json.loads(resp.read())
        except urllib.error.HTTPError as exc:
            try:
                return exc.code, json.loads(exc.read())
            except Exception:
                return exc.code, {"ok": False, "error": str(exc)}
        except (urllib.error.URLError, OSError, TimeoutError) as exc:
            raise RemoteConnectionError(
                f"Impossible de joindre le serveur {self.host}:{self.port} — {exc}"
            )


def __getattr__(name):
    """Miroir dynamique de core.py : `client_core.add_entry(remote, ...)`
    transforme automatiquement l'appel en `remote.call("add_entry", ...)`
    — même signature que la fonction core.py correspondante (le premier
    argument `conn` devient la connexion réseau `remote`), pour permettre
    de réutiliser presque tel quel le code des écrans existants."""
    if name.startswith("_"):
        raise AttributeError(name)

    def _proxy(conn, *args, **kwargs):
        if not isinstance(conn, RemoteConnection):
            raise TypeError(
                "client_core attend une RemoteConnection en premier argument (pas une connexion SQLite locale)."
            )
        return conn.call(name, *args, **kwargs)

    _proxy.__name__ = name
    return _proxy
