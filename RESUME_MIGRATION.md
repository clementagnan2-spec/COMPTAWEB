# RÉSUMÉ POUR NOUVELLE CONVERSATION — Application SYSCOHADA + Paie

## À FAIRE EN PREMIER dans la nouvelle conversation
Donner ce fichier ZIP complet à Claude et dire : « Voici mon projet, lis
RESUME_MIGRATION.md pour te mettre à jour, puis continue le travail. »

---

## CE QU'EST LE PROJET

Application de gestion intégrée SYSCOHADA (Burkina Faso) — comptabilité,
GRH/Paie, stocks, immobilisations, achats/ventes, transport, trésorerie —
en Python/Tkinter, avec 3 programmes :
1. **`main.py`** (~9800 lignes) → compile en `SaisieComptable.exe` —
   application de bureau autonome (ouvre directement une base SQLite
   locale).
2. **`server.py`** → compile en `SaisieComptableServeur.exe` — serveur
   réseau (HTTP/JSON, bibliothèque standard uniquement) qui expose la
   même base de données pour un accès multi-utilisateur simultané.
3. **`client_main.py`** (~5600 lignes) → compile en
   `SaisieComptableClient.exe` — application de bureau séparée qui se
   connecte au serveur par réseau (LAN), sans jamais toucher de fichier
   local (sauf lecture/écriture de fichiers .xlsx locaux pour les
   imports/exports/impressions, voir plus bas).

`core.py` (~9600 lignes) contient tout le moteur métier (comptabilité,
GRH, paie, achats, ventes, stocks, immobilisations, trésorerie...),
partagé par les 3 programmes. `client_core.py` est un module miroir qui
transforme les appels `core.py` en requêtes réseau HTTP/JSON vers le
serveur.

Build automatique via GitHub Actions : `.github/workflows/main.yml`
(⚠ le fichier s'appelle **main.yml**, pas build.yml). Un seul push
déclenche les 3 compilations PyInstaller et produit **un seul artifact**
nommé `SaisieComptable-windows` contenant les 3 `.exe`.

## L'UTILISATEUR — CONTEXTE IMPORTANT

- **Non technique**, a beaucoup de mal avec les manipulations Windows
  (invite de commandes, GitHub, gestion de fichiers). Donner des
  instructions très explicites, étape par étape, sans jargon.
- **Ne veut PAS utiliser Python directement** — reste sur le flux .exe +
  GitHub Actions.
- Port **8765 par défaut bloqué par Windows** sur sa machine (WinError
  10013/10061) — utilise **le port 8080** à la place, systématiquement.
- Utilise deux scripts `.vbs` faits sur mesure pour démarrer le serveur
  en arrière-plan sans fenêtre visible : `Lancer_Serveur_Arriere_Plan.vbs`
  détecte automatiquement les cartes réseau de la machine et demande le
  port + l'adresse d'écoute (mémorise le dernier port choisi). **Cette
  détection réseau existe UNIQUEMENT dans le .vbs — l'utilisateur a
  explicitement demandé de NE PAS la remettre dans le logiciel
  lui-même** (une tentative d'ajouter un bouton « Rechercher les
  serveurs sur le réseau » côté client + réponse UDP côté serveur a été
  faite puis intégralement retirée sur sa demande).
- Utilise couramment un dossier `Desktop\USINE PRO` comme emplacement de
  travail réel (visible dans plusieurs captures).

## PROBLÈME RÉCURRENT LE PLUS FRÉQUENT — À VÉRIFIER EN PREMIER

**Symptôme** : "Fonction non autorisée à distance", `TypeError: ... got
an unexpected keyword argument`, écran vide, ou tout comportement qui ne
correspond pas à ce qui vient d'être codé → dans l'immense majorité des
cas, **l'utilisateur teste encore un ancien .exe / ancien serveur**, pas
le dernier build.

**Solution en place** : `server.py` a une constante
`SERVER_VERSION` (actuellement **`"2026-08-26-v7"`** — à incrémenter à
CHAQUE modification de `core.py` ou `server.py`, format
`AAAA-MM-JJ-vN`), affichée à 3 endroits : console du serveur au
démarrage, bouton "Tester la connexion" du client, barre du haut du
client une fois connecté. **Toujours demander/vérifier ce numéro avant
de chercher un bug ailleurs.** Rituel de redéploiement systématique
demandé à l'utilisateur après chaque livraison : arrêter l'ancien
serveur (`.vbs` ou Gestionnaire des tâches → Fin de tâche sur
`SaisieComptableServeur.exe`) → remplacer les 3 `.exe` → relancer →
vérifier le numéro de version avant de retester.

## ARCHITECTURE DE SÉCURITÉ DU SERVEUR (important, ne pas régresser)

`server.py` utilise un modèle **liste noire** (pas liste blanche) :
`RPC_WHITELIST` est calculée dynamiquement au démarrage (toutes les
fonctions publiques de `core.py` prenant `conn` en premier argument),
moins `RPC_BLOCKLIST` et les fonctions dont le nom commence par
`export_`/`import_` (chemins de fichiers locaux, sans usage réseau
direct — voir section Imports Excel plus bas pour le contournement
utilisé). **Ne JAMAIS revenir à une liste blanche manuelle.**

Toute nouvelle fonction `core.py` prenant `conn` en premier paramètre
est donc **automatiquement exposée** au client réseau, sauf ajout
explicite à `RPC_BLOCKLIST`. Vérifier ce point avant d'ajouter une
fonction sensible.

## VALIDATION DES DONNÉES CÔTÉ SERVEUR (important, ne pas régresser)

Suite à des saisies « n'importe quoi » constatées (comptes inventés,
fournisseurs vides, montants texte...), une règle de fond a été posée :
**toute donnée métier doit être validée côté SERVEUR**, jamais
seulement côté UI, pour protéger TOUS les clients (bureau, réseau,
futurs) de façon uniforme :
- `add_ecriture_multi_lignes`, `update_entry` : vérifient l'existence
  réelle du compte (`account_exists`), du code analytique
  (`analytic_code_exists`), du fournisseur/client, et imposent le
  tiers auxiliaire obligatoire dès qu'un compte racine 40x
  (Fournisseurs) ou 41x (Clients) est utilisé — y compris lors d'une
  **modification** d'écriture (pas seulement à la création).
- `create_facture_vente`, `add_ligne_facture_vente` : client/compte de
  vente obligatoires et vérifiés.
- `set_immobilisation_fiche` : compte, fournisseur vérifiés.
- Toute nouvelle saisie doit suivre ce même principe.

En complément, le client réseau (`RemoteSaisieTab`) valide aussi **au
moment de la saisie** (pas seulement à l'enregistrement final) : le
champ Compte est vérifié dès « Ajouter la ligne », et
Fournisseur/Client/Code budgétaire/Code bailleur sont vérifiés dès
qu'on quitte le champ (`<FocusOut>`), avec effacement + message clair
si invalide.

## ÉTAT D'AVANCEMENT — CLIENT RÉSEAU

Les 48 sous-menus de l'application sont couverts dans `client_main.py`.
Un très gros travail de mise à niveau du client a été fait dans cette
conversation (le client était initialement beaucoup plus pauvre que le
bureau sur plusieurs écrans) :

**Écrans reconstruits pour la parité complète avec le bureau** (avant :
lecture seule ou incomplets ; maintenant : CRUD complet, comme le
bureau) : Fabrication (recettes + coût de production + validation),
Stocks (synthèse éditable + mouvements comptables), Recouvrement
(factures + paiement + bug de balance âgée corrigé), Contrats
fournisseurs (livraison/paiement/suppression), Réparations (création +
pièces utilisées), Immobilisations (fiche éditable + import Excel),
Saisie (édition/suppression des écritures, tous les champs
Fournisseur/Client/Code analytique/budgétaire/bailleur/Quantité comme
le bureau), Facturation (client/compte validés, suppression, aperçu
avant impression), GRH > Paie (nouveau module entier, voir plus bas).

**5 écrans ADMIN sensibles** (Utilisateurs, Niveaux d'accès,
Réinitialisation, etc.) restent volontairement réservés au bureau.

**Filtrage des menus par profil** (niveau d'accès) fonctionne sur les
deux applications — voir `core.MENU_STRUCTURE`,
`core.get_menus_autorises()`. **Attention** : ajouter une clé à
`MENU_STRUCTURE` ne donne PAS automatiquement accès à ce menu aux
niveaux d'accès déjà configurés en base (ex. ajout de `grh_paie` :
il a fallu que l'utilisateur aille cocher la case dans ADMIN > Niveaux
d'accès pour le niveau concerné). `ajouter_niveaux_acces_suggeres_menus`
ne s'applique qu'aux niveaux pas encore configurés.

## MODULE PAIE (GRH > Paie) — AJOUTÉ DANS CETTE CONVERSATION

Porté depuis un projet séparé fourni par l'utilisateur (« PaieBurkina »,
un calculateur de paie autonome avec son propre payroll_engine.py). Le
moteur de calcul a été **repris à l'identique** (mêmes formules CNSS,
IUTS à 9 tranches, abattement CADRE/AUTRE, exonérations
Logement/Fonction/Transport, réduction pour charges de famille) — un
test croisé confirme des résultats **identiques au franc près** avec le
moteur d'origine. Intégration native (pas un .exe séparé lancé à côté) :
utilise les employés déjà saisis dans GRH > Personnel plutôt qu'une
liste séparée, et les comptes/niveaux d'accès déjà en place.

Fonctions clé dans `core.py` : `compute_bulletin_paie`,
`get_paie_parametres`/`set_paie_parametres` (JSON dans settings),
`set_bulletin_paie`/`get_bulletin_paie`/`list_bulletins_paie`,
`compute_paie_periode`, `valider_paie_periode` (comptabilise + verrouille
la période), `render_bulletin_paie_html` (aperçu avant impression),
`export_paie_bulletins_template`/`parse_paie_bulletins_xlsx`/
`apply_paie_bulletins_rows` (import Excel en masse par matricule).

Écran (3 sous-onglets, bureau ET client) : **Bulletins** (saisie par
employé/mois + import Excel), **État de paie** (calcul + export Excel +
bouton **Valider la paie (comptabiliser)**), **Paramètres de paie**
(taux CNSS/TPA/abattements, réservé admin).

**Comptabilisation automatique** (bouton Valider) génère, par employé :
Débit 661100 (salaires), Crédit 431000 (CNSS), 447210 (IUTS), 447220
(retenue obligatoire 1%), 421000 (remboursement prêt si applicable),
422000 (net à payer) ; puis charges patronales : Débit 664100/664200,
Crédit 431000/442810 (TPA). Tous ces comptes existent réellement dans
le plan comptable bundlé — vérifié. Écritures testées équilibrées.
**Une période validée est verrouillée** (bulletins non modifiables/
supprimables ensuite) — table `paie_periodes_validees`.

## AMORTISSEMENT D'ÉQUIPEMENT DANS LES RECETTES DE FABRICATION

Nouveau type de composant dans PRODUCTION > Fabrication > Recettes :
« Amortissement d'équipement (depuis une immobilisation) ». Permet
d'incorporer le coût d'usage d'un équipement (pelle mécanique, camion,
concasseur...) dans le coût de production, en F CFA par tonne/heure/etc.

Sur la fiche de l'équipement (écran IMMOBILISATIONS), deux nouveaux
champs : **Base de répartition** (quantité annuelle de référence, ex.
5000 tonnes/an) + **Unité**, et **Amortissement annuel (si pas
comptabilisé)** — un montant déclaré manuellement, utilisé en attendant
que de vraies dotations aux amortissements soient comptabilisées (dès
qu'une vraie dotation existe pour ce compte, elle prend automatiquement
le relais du montant manuel — priorité testée et vérifiée).
`compute_cout_amortissement_unitaire()` = amortissement (réel ou
manuel) ÷ base de répartition.

Import/export Excel des fiches d'immobilisations également ajouté
(mêmes boutons que pour les bulletins de paie, même principe).

## IMPRESSION / APERÇU AVANT IMPRESSION

Système HTML unifié (`_html_facture_pro` dans `core.py`) pour les
factures de vente, factures d'achat, bons de commande ET bulletins de
paie : cadre entreprise (repris de ADMIN > Liasse fiscale — inclut
maintenant aussi Téléphone et RCCM, ajoutés dans cette conversation),
bloc « Doit : », tableau des lignes, totaux, montant en toutes lettres
(`nombre_en_lettres_fr()`, testé et corrigé — bug initial sur
« quatre-vingt-dix »). Un bouton « Aperçu avant impression » ouvre le
document dans le navigateur (rien n'est imprimé à ce stade) ; un
bouton/texte à l'intérieur de la page déclenche l'impression réelle
(Ctrl+P ou bouton dédié) — vocabulaire choisi précisément pour éviter
toute confusion avec une impression accidentelle.

Fonctions `render_*_html(conn, id)` renvoient le HTML en `str` (pas de
fichier) pour usage à distance par le client réseau, qui écrit le
fichier localement puis l'ouvre avec `webbrowser`.

## IMPORTS EXCEL — PATTERN À RÉUTILISER

Pour toute nouvelle fonctionnalité d'import Excel en masse, séparer
systématiquement en 2 fonctions (comme fait pour Immobilisations et
Paie) :
1. `parse_XXX_xlsx(path)` — lecture pure, AUCUN accès `conn`,
   utilisable localement par le bureau ET par le client réseau (lit le
   fichier sur son propre poste).
2. `apply_XXX_rows(conn, rows)` — écriture en base à partir des lignes
   déjà lues, appelée directement par le bureau ou via RPC par le
   client. **Important** : ce nom ne doit PAS commencer par `import_`
   ni `export_`, sinon le filtre RPC l'exclut automatiquement de la
   liste blanche dynamique (voir section sécurité serveur).
3. `import_XXX_from_xlsx(conn, path)` = wrapper `parse` + `apply`, pour
   l'usage direct et pratique du bureau uniquement.

`openpyxl` a dû être ajouté explicitement au build PyInstaller du
client réseau (`--hidden-import openpyxl` et
`--hidden-import openpyxl.cell._writer` dans `.github/workflows/main.yml`,
sur les 3 builds) — absent par défaut car le client n'en avait jamais eu
besoin avant.

## RECHERCHE FOURNISSEUR/CLIENT SUR INTERNET

Écrans Fournisseurs et Clients (bureau + client) : bloc « Trouver de
nouveaux fournisseurs/clients sur Internet » — ouvre une recherche
Google préremplie (produit + ville) dans le navigateur par défaut.
Choix délibéré de rester simple/gratuit (pas d'API de recherche
payante) suite à une discussion sur les coûts avec l'utilisateur.

## CORRECTIONS DE MISE EN PAGE (39 écrans)

Deux défauts de mise en page systémiques corrigés dans toute
l'application (bureau + client) :
1. **Labels de texte dynamique sans `wraplength`** poussant la fenêtre
   hors de l'écran (ex. messages d'aperçu de coût en Fabrication) —
   corrigé partout où repéré.
2. **Widgets `pack(fill="both", expand=True)` suivis d'autres widgets**
   (boutons, totaux) qui se retrouvaient poussés hors de la zone
   visible sur les petites fenêtres — `expand=True` retiré partout où
   un widget de taille naturelle bornée (Treeview avec ou sans
   `height=`, `tk.Text`) était suivi d'autre chose (31 classes dans
   `main.py`, 8 dans `client_main.py`). Les tableaux gardent une
   hauteur naturelle raisonnable (~10 lignes par défaut) avec
   défilement automatique si besoin.

## PROBLÈME EN COURS, NON RÉSOLU — TFT (Tableau des flux de trésorerie)

L'utilisateur signale « 2 formule(s) du gabarit n'ont pas pu être
évaluées » sur RAPPORTS FINANCIERS > TFT. **Vérifié : aucune des 31
formules du gabarit TFT (`_tft_template_path` /
`etats_financiers_data.TFT_TEMPLATE_B64`) ne contient de division** —
le message d'avertissement générique (« souvent une division par
zéro ») est donc trompeur pour cet écran précis, et la cause exacte n'a
pas encore été identifiée (non reproduite avec des données de test
simples). **Amélioration apportée** : le message affiche maintenant le
détail précis (cellule + texte d'erreur exact) au lieu d'un simple
compteur — sur le bureau ET le client (qui ignorait ces erreurs
silencieusement avant). **Prochaine étape** : demander à l'utilisateur
une capture du nouveau message détaillé pour identifier les 2 cellules
exactes et corriger la formule ou la donnée en cause dans
`evaluate_formula`/`evaluate_sheet_formulas` (`core.py`, section
gabarits à formules `CtaCptSolde...`).

## BUGS MAJEURS DÉJÀ TROUVÉS ET CORRIGÉS (ne pas réintroduire)

1. `compute_balance()` excluait les comptes absents du Plan comptable
   bundlé — corrigé : parcourt l'union Plan comptable + soldes
   d'ouverture + écritures.
2. Formules `...Nm1` (N-1) utilisaient un exercice séparé au lieu du
   solde d'ouverture de l'exercice courant — corrigé (TFT, CR, Situation
   financière, Bilan gabarit).
3. PyInstaller + import dynamique (`importlib.import_module`) ne bundle
   pas le module — toujours utiliser des `import X` littéraux à
   l'intérieur des fonctions, jamais dynamiques.
4. Le serveur restait figé sur un instantané ancien de la base (WAL) —
   corrigé avec un `conn.commit()` avant chaque requête réseau.
5. Workflow GitHub nommé **`main.yml`**, pas `build.yml`.
6. `RemoteRecetteFabricationTab` traitait le résultat de
   `compute_production()` (un résumé/dict) comme une liste de lignes —
   plantage. Corrigé, testé.
7. `RemoteRecouvrementTab` traitait `tranches` (une LISTE indexée
   renvoyée par `compute_balance_agee`) comme un dict par libellé de
   tranche — plantage dès qu'il y avait des impayés. Corrigé.
8. `compute_cout_amortissement_unitaire` renvoyait un montant NÉGATIF
   (l'amortissement est stocké en négatif par convention interne,
   `valeur_nette = brut + amortissement`) — corrigé avec `abs()`.
9. Fonction `nombre_en_lettres_fr()` : bug sur les nombres finissant par
   ...70/...90 (« quatre-vingt » au lieu de « quatre-vingt-dix ») — testé
   et corrigé.
10. Bouton `ttk.Button(row=6...)` dans ADMIN > Liasse fiscale codé en
    dur, chevauchant les nouveaux champs Téléphone/RCCM — corrigé en
    position dynamique `(len(core.COMPANY_FIELDS) + 1) // 2 * 2`.

## STYLE DE TRAVAIL ATTENDU (déjà établi, à poursuivre)

- Toujours compiler (`python3 -m py_compile`) les 5 fichiers Python
  avant de livrer.
- Pour toute modification serveur/client réseau : tester avec un VRAI
  serveur + VRAI client (lancer `server.py --port <libre> --db
  <db_temp>` en arrière-plan dans le même appel bash que le test client,
  car chaque appel `bash_tool` est une session séparée — le serveur lancé
  en arrière-plan dans un appel précédent ne survit PAS à l'appel
  suivant).
- Toujours vérifier la non-régression du moteur comptable (écritures
  équilibrées : total débit = total crédit) après une modification
  touchant la comptabilisation.
- Incrémenter `SERVER_VERSION` dans `server.py` à chaque modification de
  `core.py` ou `server.py`.
- Toujours repackager le zip complet (`zip -r ... -x "*.pyc" -x
  "__pycache__/*"`) et le fournir via `present_files` après chaque
  modification.
- Donner des instructions Windows très explicites et simples.
- Ne pas deviner à l'aveugle un comportement UI ambigu — poser une
  question de clarification courte (l'utilisateur a explicitement
  redemandé un retour en arrière une fois après une supposition
  incorrecte sur un bouton "Modifier").

## FICHIERS DU PROJET

```
main.py                    Application de bureau (~9800 lignes)
core.py                    Moteur métier partagé (~9600 lignes)
server.py                  Serveur réseau (SERVER_VERSION à incrémenter)
client_core.py             Module miroir réseau (proxy RPC générique)
client_main.py             Application client réseau (~5600 lignes)
bilan_template_data.py     Gabarit Bilan encodé en base64
etats_financiers_data.py   Gabarits CR/TFT/Situation financière (base64)
factory_icon_data.py       Icône de l'application (base64)
factory_icon.ico
plan_comptable.json        Plan comptable SYSCOHADA de référence
templates/                 Gabarits Excel bruts (sources des .py base64)
.github/workflows/main.yml Build GitHub Actions (3 exécutables, openpyxl inclus)
requirements.txt
README.md                  Journal détaillé historique (très long)
RESUME_MIGRATION.md        Ce fichier
```

Scripts fournis séparément à l'utilisateur (pas dans le zip du code,
donnés directement en pièce jointe dans la conversation) :
`Lancer_Serveur_Arriere_Plan.vbs` (détection réseau + port, mémorise le
dernier choix) — à replacer dans le dossier des 3 `.exe` si perdu.
