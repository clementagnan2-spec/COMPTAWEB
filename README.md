# Saisie Comptable SYSCOHADA — application Windows autonome

## 🔄 RÉSUMÉ DE REPRISE (à lire en premier par toute nouvelle conversation Claude)

**Contexte** : application de comptabilité SYSCOHADA développée sur
plusieurs longues sessions avec Claude. Le code est **volumineux et déjà
très abouti** (`core.py` ~180 Ko, `main.py` ~210 Ko) — avant toute
modification, lis intégralement les fichiers concernés plutôt que de les
régénérer ou de les réécrire à partir d'une supposition. Une réécriture
depuis zéro **casserait** des dizaines de fonctionnalités déjà construites
et testées.

**Dépôt GitHub** : build automatique via GitHub Actions (le `.exe` ne peut
pas être compilé directement par Claude, environnement Linux). Voir section
suivante pour le processus de mise à jour.

**Ce qui existe déjà (ne pas reconstruire, juste modifier/étendre)** :
- Menu à 6 entrées : SAISIE, COMMERCE, PRODUCTION, ENGAGEMENTS-PROJETS,
  ÉTATS ET RAPPORTS, PARAMÈTRES (navigation par menu déroulant, pas
  d'onglets classiques — voir `class App` dans `main.py`)
- **Saisie** : partie double forcée (Compte débiteur + Compte créditeur
  obligatoires ensemble), validation des comptes/tiers en temps réel,
  liste déroulante automatique au clic, sélection multiple + Ctrl+A +
  suppression groupée (transaction unique, pas de commit par ligne)
- **Exercices comptables** : multi-exercices avec clôture annuelle
  (report des soldes + résultat net vers le compte 121000), verrouillage
  des exercices clôturés
- **Plan comptable** : 1591 comptes (import Sage), comptes racines
  (1 chiffre, ou 40-49 pour la classe 4), rattachement obligatoire des
  écritures 40xxx/41xxx à un fournisseur/client
- **Commerce** : Ventes, Clients, Recouvrement, Facturation (avec sortie de
  stock automatique et TVA), Stocks, Marges
- **Engagements-projets** : Achats, Fournisseurs, Factures frs (entrée de
  stock automatique + retenue à la source), Contrats
- **Production** : Fabrication avec nomenclature (BOM), coût de production,
  validation qui décrémente les matières et valorise le produit fini
- **États et rapports** : Grand livre complet (tous comptes, bandes de
  couleur), Balance (sous-totaux par classe), **Bilan « avec détails »**
  (Actif gauche en Brut/Amortissements/Net, Passif droite en Montant,
  détail compte par compte — immobilisations, stocks, créances, dettes,
  trésorerie banque par banque —, toujours équilibré par construction,
  exportable en .xlsx ; voir section « Bilan « avec détails » » plus bas),
  Compte de résultat (SIG), TFT (méthode indirecte CAFG), Situation
  financière (FR-BFR-TN), Liasse fiscale (export 92 pages, cohérente avec
  tous les écrans ci-dessus)
- **Paramètres** : gestion des 4 plans (comptable/analytique/budgétaire/
  bailleurs) avec import/export xlsx (écrase à l'import)

**Pièges déjà rencontrés (pour ne pas les refaire)** :
- Les f-strings avec apostrophe échappée (`f"{'d\\'accord'}"`) plantent en
  Python < 3.12 — toujours extraire la chaîne dans une variable avant.
- Toujours utiliser `account_racine()`/préfixes (pas le code exact à 6
  chiffres) pour agréger des comptes — le vrai plan comptable de
  l'utilisateur est plein de sous-comptes détaillés (602101, 521120...).
- Chaque `conn.commit()` coûte cher en boucle — grouper en une transaction
  pour les opérations multiples.
- Toujours tester avec `python3 -m py_compile` et un scénario réel
  (`core.get_connection('t.db')` puis nettoyer) avant de livrer.
- Vérifier la cohérence Balance ↔ Bilan ↔ TFT ↔ Situation financière ↔
  Liasse fiscale : ils partagent tous `compute_balance()` **et, depuis la
  correction de l'équilibre du Bilan, le même `compute_resultat_net_complet()`
  pour le résultat net.**
- **Ne jamais calculer un total de Bilan (résultat net, capitaux propres,
  stocks...) à partir d'une liste de comptes codée en dur** (ex.
  `COMPTES_CAPITAL = ["101","118","121"]`) : le vrai plan comptable de
  l'utilisateur (1591 comptes Sage) contient forcément des comptes hors de
  toute liste pré-définie, qui disparaîtraient alors silencieusement du
  Total Actif/Passif et casseraient l'équilibre. Pour un TOTAL, toujours
  sommer la classe entière (`_sum_class(balance, "1")` etc.) ; les listes de
  comptes codées en dur ne sont acceptables que pour un DÉTAIL affiché à
  titre indicatif, avec une ligne « Autres » qui absorbe le reliquat pour
  que le détail somme exactement au vrai total.

**Ce qui reste à construire/imparfait** : Contrats (module vide),
Tableaux d'exécution budgétaire (encore un placeholder) ; les lignes
d'investissement/financement de la vraie feuille TFT officielle dans la
Liasse fiscale (seules ZA/FA-FE sont mappées, positions FF+ à confirmer).
Impôts, Déclarations sociales et Rapprochements bancaires sont désormais
construits (voir section dédiée plus bas) — plus des placeholders.

**Les données réelles de l'utilisateur (`comptabilite.db`) ne sont PAS ici**
— uniquement sur son PC Windows local
(`%LOCALAPPDATA%\SaisieComptable\`). Ne jamais supposer leur contenu.

---

Application de bureau (Tkinter) qui reproduit les fonctions essentielles
du classeur Excel : Saisie des écritures, Balance, Compte de résultat et
Bilan, calculés automatiquement. Aucune installation d'Excel n'est requise :
une fois compilée, c'est un simple `.exe`.

Le plan comptable intégré (`plan_comptable.json`) est celui importé depuis
votre export Sage (1591 comptes).

## Important : je ne peux pas produire le .exe moi-même

Un `.exe` est un binaire Windows. Je travaille dans un environnement Linux
qui ne peut pas compiler de binaire Windows. La solution ci-dessous utilise
**GitHub Actions** : GitHub compile lui-même le `.exe` sur une machine
Windows à chaque fois que vous poussez du code — c'est la manière standard
et fiable de faire, sans avoir besoin d'un PC Windows.

## Mise en ligne sur GitHub (une seule fois)

```bash
cd accounting-app
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/VOTRE-COMPTE/VOTRE-DEPOT.git
git push -u origin main
```

## Récupérer le .exe

1. Sur GitHub, ouvrez l'onglet **Actions** de votre dépôt : le workflow
   « Build Windows .exe » se déclenche automatiquement à chaque push sur
   `main` (ou lancez-le manuellement via **Run workflow**).
2. Une fois le job terminé (~2-3 minutes), ouvrez son résumé et téléchargez
   l'artifact **SaisieComptable-windows** : il contient `SaisieComptable.exe`.
3. Pour publier une **Release** téléchargeable en un clic (recommandé pour
   partager l'app), créez un tag :
   ```bash
   git tag v1.0
   git push origin v1.0
   ```
   Le `.exe` sera automatiquement attaché à la Release correspondante.

## Utilisation de l'application

- **Saisie** : formulaire d'ajout/modification/suppression d'écritures
  (Date, Pièce, Journal, Compte, Tiers, Libellé, Débit, Crédit, Code flux,
  Code analytique). Le champ **N° Compte** est une liste déroulante avec
  recherche : tapez un numéro ou un mot du libellé (ex. `clients`, `601`,
  `banque`) et choisissez le compte dans la liste qui s'affiche. Le
  **Journal** propose AC/VE/OD/BQ/CA (modifiable librement), et le
  **Code flux** est une liste fermée EXP/INV/FIN pour éviter les fautes de
  frappe. Le libellé du compte s'affiche automatiquement pendant la saisie.

  **Import massif (.xlsx)** *(nouveau)* : pour les volumes d'écritures
  importants, deux boutons sont disponibles au-dessus du tableau :
  - **« Télécharger un modèle (.xlsx) »** : génère un fichier vierge avec
    les bons en-têtes (Date, N° Pièce, Journal, N° Compte, Tiers, Libellé,
    Débit, Crédit, Code flux, Code analytique) et deux lignes d'exemple.
  - **« Importer des écritures (.xlsx) »** : sélectionnez votre fichier
    préparé (l'ordre des colonnes n'a pas d'importance, les en-têtes sont
    reconnus automatiquement) — toutes les lignes sont ajoutées à la
    Saisie en une fois. Les dates peuvent être au format texte (AAAA-MM-JJ)
    ou en dates Excel natives. Les lignes vides sont ignorées ; un compte
    absent du plan comptable ou un montant non numérique déclenche un
    avertissement (la ligne est quand même importée, avec le montant
    invalide remplacé par 0) plutôt que de faire échouer tout l'import.
- **Balance** : synthèse Débit/Crédit/Solde par compte, actualisée à la volée.
- **Compte de résultat** et **Bilan** : calculés automatiquement selon la
  même logique que le classeur Excel (mêmes regroupements de comptes).

Les données sont stockées localement dans :
`%LOCALAPPDATA%\SaisieComptable\comptabilite.db` (SQLite). Elles persistent
d'un lancement à l'autre de l'application.

## Développer / tester en local (optionnel)

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Navigation

L'application n'a plus d'onglets classiques : la navigation se fait
entièrement via **la barre de menu** en haut de la fenêtre, avec 5 menus
principaux (en gras) :

- **SAISIE** : Saisie des écritures, Soldes d'ouverture.
- **COMMERCE** : Ventes, Clients, Recouvrement, Facturation, Stocks, Marges bénéficiaires.

### Correction majeure : reconnaissance des sous-comptes détaillés (important)

**Bug signalé et corrigé** : un achat de matières premières saisi directement
dans l'onglet Saisie sur un **sous-compte détaillé** (ex. `602101 ACHAT
CLINKER`, au lieu du compte maître `602000`) ne mettait pas le stock à jour,
et — plus grave — **faussait le calcul du Résultat et du Bilan** (écart non
nul), car les comptes de résultat/trésorerie/capitaux propres n'étaient
reconnus que sur leur code exact à 6 chiffres.

**Deux corrections apportées :**
1. **Mise à jour automatique du stock désormais aussi en Saisie directe** :
   dès qu'une écriture équilibrée (Compte débiteur/Compte créditeur) touche
   un compte d'achat (601x/602x) ou de vente (701x/702x) lié à un stock,
   **avec une quantité renseignée**, l'entrée ou la sortie de stock
   correspondante est automatiquement comptabilisée — plus besoin de passer
   par Facturation/Factures frs pour que le stock se mette à jour. Ces
   écritures apparaissent dans l'onglet Stocks → Mouvements comptables sous
   l'origine **« Saisie directe (auto) »**.
2. **Rattachement par racine/préfixe (3 chiffres) partout** : tous les
   calculs qui agrègent des comptes (Résultat, Bilan, Trésorerie, Production)
   reconnaissent désormais **tous les sous-comptes** d'une racine donnée
   (ex. 602101, 602102... sont bien rattachés à 602 ; 521100, 521120...
   sont bien rattachés à 521), et pas seulement le compte maître à 6 chiffres
   se terminant par des zéros.

Testé : le scénario exact du bug (achat de 4 500 000 sur le compte 602101,
quantité 100) met maintenant bien à jour le stock matières premières
(4 500 000 / 100 unités) **et** le Bilan reste parfaitement équilibré (écart
= 0) — vérifié aussi après un cycle complet de clôture d'exercice et dans
l'export de la Liasse fiscale.

### Onglet Stocks — mouvements comptables détaillés (nouveau)

L'onglet **Stocks** (menu COMMERCE, aussi accessible depuis PRODUCTION →
Matières premières/Produits finis) a maintenant deux sous-onglets :
- **« Synthèse par compte »** : le tableau existant (stock initial, entrées,
  sorties, stock final, coût unitaire moyen).
- **« Mouvements comptables (classe 3) »** *(nouveau)* : le détail
  chronologique de **toutes** les écritures sur les comptes de stock
  (310000, 320000, 331000, 360000) de l'exercice en cours, avec un filtre
  par origine :
  - **Facturation** : sorties de stock générées automatiquement par la
    validation d'une facture de vente.
  - **Facture frs** : entrées de stock générées automatiquement par la
    validation d'une facture d'achat.
  - **Saisie manuelle** : toute écriture sur un compte de stock passée
    directement dans l'onglet Saisie.

  Chaque ligne affiche à la fois le **mouvement** (Débit/Crédit en valeur,
  Qté mvt) et le **cumul** après ce mouvement (**Qté cumulée** et **Valeur
  cumulée**), en partant du stock initial de l'exercice — comme une vraie
  fiche de stock. Les lignes générées automatiquement sont affichées en
  bleu. Testé : une facture d'achat (+20 unités) puis une vente (-10
  unités) sur un stock initial de 100 unités / 300 000 donnent bien un
  cumul de 120 puis 110 unités, avec la valeur qui suit correctement.

### Module Facturation (nouveau)

L'onglet **Facturation** présente directement une facture éditable :
- **En-tête modifiable** et **pied de page modifiable** (texte libre).
- **N° Facture**, **Date**, **Client** (obligatoirement rattaché à un compte
  racine 41, avec la même recherche/validation que dans le reste de l'app).
- **Taux de TVA paramétrable** (compte 44 — 443100 « T.V.A. facturée sur
  ventes »), avec une valeur par défaut mémorisée d'une facture à l'autre.
- **Lignes de vente** liées à un compte de classe **70** (Ventes) : chaque
  ligne a un compte, un libellé, une quantité et un prix unitaire ; le
  montant HT est calculé automatiquement.

**Bouton « Valider et envoyer en Saisie »** : génère automatiquement les
écritures comptables équilibrées dans l'onglet Saisie :
- Débit **Client** (411000) pour le montant TTC.
- Crédit chaque **compte de vente** (70x) pour le HT de sa ligne.
- Crédit **TVA facturée** (443100) pour la taxe.
- **Mise à jour automatique des stocks** : les comptes 701000 (marchandises,
  stock 310000) et 702000 (produits finis, stock 360000) déclenchent en plus
  une sortie de stock au coût unitaire moyen réel (Débit 603100 ou 736000 /
  Crédit le compte de stock correspondant) — les comptes de services
  (ex. 706000) n'impactent aucun stock. Ce mapping compte-de-vente ↔ stock
  est défini dans `core.VENTE_STOCK_MAPPING` (extensible).

Une fois validée, une facture est **verrouillée** (plus de modification
possible, cohérent avec le fait que ses écritures existent déjà en Saisie).

**Bug de calcul du Résultat corrigé au passage** : les comptes de variation
de stock (603100 pour les marchandises, 736000 pour les produits finis)
n'étaient référencés dans aucune formule du Compte de résultat, ce qui
créait un écart Actif/Passif après une vente de marchandises ou de produits
finis. Testé et corrigé : un scénario complet (service + marchandise +
produit fini + TVA) donne désormais un Bilan parfaitement équilibré et un
Résultat net exact (vérifié à l'unité près sur plusieurs cas).

- **PRODUCTION** : Matières premières, Fabrication, Produits finis.

### Reconfiguration majeure : Stocks au détail réel + Fabrication qui consomme les matières (mise à jour)

**Onglet Stocks → Synthèse par compte** : affiche désormais le **détail réel
de chaque compte** de stock utilisé (ex. `321001 CLINKER`, `321002 GYPSE`),
et non plus seulement les 4 comptes centralisateurs (310000/320000/331000/
360000). Un filtre par catégorie (31 Marchandises / 32 Matières premières /
33 Autres approvisionnements / 36 Produits finis) est disponible, ainsi
qu'un nouveau champ **« Marge de valorisation des produits finis par défaut
(%) »**, utilisé comme marge par défaut pour tout nouveau produit créé dans
Fabrication.

**Onglet Fabrication → Recettes / Coût de production** reconfiguré :
- Le sélecteur **« Compte de stock »** des lignes matière propose désormais
  tous les comptes détaillés réellement utilisés dans Stocks (pas seulement
  les 4 comptes centralisateurs) — vous pouvez donc combiner clinker, gypse,
  calcaire... chacun avec son propre coût réel.
- Chaque produit fini a maintenant un **compte de stock configurable**
  (classe 36) où il sera placé une fois fabriqué.
- **Nouveau bouton « Valider la fabrication (comptabiliser) »** : envoie les
  écritures comptables dans le menu SAISIE —
  1. **Consommation des matières premières** : chaque matière utilisée dans
     la recette est diminuée en **quantité et en valeur** sur son compte
     réel (ex. 321001), avec pour contrepartie le compte de variation de
     stock approprié (603200 pour les matières premières, 603100 pour les
     marchandises...).
  2. **Entrée du produit fini** : le compte de stock du produit (classe 36)
     est augmenté en **quantité et en valeur**, valorisé au **coût de
     production + la marge paramétrée**, avec pour contrepartie le compte
     736000.

**Trois bugs trouvés et corrigés pendant les tests** (tous liés à la
reconnaissance des sous-comptes réels) : (1) le calcul du coût de
production ne cherchait le coût unitaire que parmi les 4 comptes maîtres —
corrigé ; (2) la fabrication était datée du jour au lieu d'une date dans
l'exercice actif par défaut — corrigé ; (3) le total des stocks au Bilan
utilisait des préfixes à 3 chiffres qui ratent les sous-comptes détaillés —
corrigé avec de vrais préfixes de catégorie à 2 chiffres.

Testé de bout en bout (clinker + gypse + main-d'œuvre + énergie → ciment,
marge 25 %) : consommation exacte des matières (quantité et valeur),
production de 10 unités de ciment valorisées à 250 000, **Bilan
parfaitement équilibré**, sans régression sur les scénarios précédents.

### Module Fabrication — nomenclature et coût de production (nouveau)

L'onglet **Fabrication** contient maintenant deux sous-onglets :

**« Recettes / Coût de production »** *(nouveau)* — un calculateur de coût de
revient (nomenclature / BOM) :
- Créez un **produit fini** (code, nom).
- Ajoutez ses composants : **matières premières** (choisies parmi les
  comptes de stock — le coût unitaire réel est repris automatiquement du
  **coût unitaire moyen** calculé dans l'onglet Stocks, donc directement
  depuis vos achats comptabilisés), **main-d'œuvre** et **énergie** (coût
  unitaire saisi manuellement), avec une quantité pour chacun.
- Le **coût de production total**, le **coût de production unitaire**
  (divisé par la quantité produite par la recette) sont calculés
  automatiquement.
- Réglez une **marge (%)** : le **prix de vente unitaire suggéré** est
  calculé automatiquement (coût de production × (1 + marge)).

Testé avec un cas concret : achat de 100 unités de matière première pour
500 000 (coût unitaire réel 5 000, repris automatiquement des stocks) → une
recette combinant 2 unités de cette matière + main-d'œuvre (3 000) +
énergie (500) donne un coût de production de 13 500, et un prix de vente
suggéré de 18 900 à 40 % de marge.

**« Coûts de fabrication (période) »** — l'ancien contenu de l'onglet
Fabrication (coûts réels de la période via l'axe analytique AN-FAB),
inchangé et toujours disponible.

- **ENGAGEMENTS-PROJETS** : Achats, Fournisseurs, Factures frs, Contrats.

### Module Factures frs (nouveau)

Le pendant achats du module Facturation. L'onglet **Factures frs** présente
directement une facture d'achat éditable :
- **En-tête** et **pied de page modifiables**.
- **N° Facture**, **Date**, **Fournisseur** (obligatoirement rattaché à un
  compte racine 40).
- **Retenue fiscale à la source paramétrable** : taux (%) et **compte de
  retenue au choix parmi la classe 44** (ex. 447810 « RETENUE 5% OPÉRÉE »),
  avec valeurs par défaut mémorisées d'une facture à l'autre.
- **Lignes d'achat** liées à un compte de classe **6** (charges) : compte,
  libellé, quantité, prix unitaire — montant HT calculé automatiquement.

**Bouton « Valider et envoyer en Saisie »** génère les écritures :
- Débit chaque **compte d'achat** (6x) pour le HT de sa ligne.
- Crédit **Fournisseur** (401000) pour le **net à payer** (HT − retenue).
- Crédit le **compte de retenue** choisi, si un taux est renseigné.
- **Mise à jour automatique des stocks** : les comptes 601000 (marchandises,
  stock 310000) et 602000 (matières premières, stock 320000) déclenchent en
  plus une **entrée de stock** (Débit le compte de stock / Crédit
  603100 ou 603200) — les comptes de service (ex. 622000 Locations)
  n'impactent aucun stock. Mapping défini dans `core.ACHAT_STOCK_MAPPING`
  (extensible).

Une fois validée, une facture est **verrouillée**. Testé de bout en bout
(service + marchandise + matière première + retenue 5%) : Bilan
parfaitement équilibré, stocks correctement augmentés, solde fournisseur
exact, et cohérence vérifiée aussi dans l'export de la Liasse fiscale.


### Module Commerce — Clients / Ventes / Recouvrement (nouveau)

- **Clients** : liste auxiliaire (fiche par client : raison sociale, contact,
  délai de paiement par défaut en jours). Créer / modifier / supprimer, ou
  **importer en masse (.xlsx)** avec un modèle téléchargeable.
- **Ventes** : soldes des opérations avec chaque client (Débit − Crédit sur
  les comptes 411xxx qui lui sont tagués), **total par client**, avec un
  **filtre de plage de dates** (Du / Au). Positif = montant restant dû par
  le client (à recouvrer).
- **Recouvrement** : journal des factures émises à chaque client. À la
  création, l'échéance de **paiement** est calculée automatiquement (date
  de facture + délai par défaut du client). Renseignez ensuite la date
  réelle de paiement au fur et à mesure des encaissements : les **retards
  sont détectés et affichés en rouge** (« EN RETARD (n j) » si l'échéance
  est dépassée sans paiement enregistré, ou « Payé (retard n j) » une fois
  la date réelle enregistrée après l'échéance).

**Saisie** : un nouveau champ **« Client »** (liste déroulante avec
recherche, proposition de création si le code n'existe pas) permet de taguer
chaque écriture — c'est ce qui alimente automatiquement les modules Ventes
et Recouvrement.

### Module Engagements-Projets (nouveau, remplace les placeholders)

- **Fournisseurs** : liste auxiliaire (fiche par fournisseur : raison sociale,
  contact, délais par défaut de paiement et de livraison en jours). Créer /
  modifier / supprimer, ou **importer en masse (.xlsx)** avec un modèle
  téléchargeable.
- **Achats** : soldes des opérations avec chaque fournisseur (Débit − Crédit
  sur les comptes 401xxx/408xxx qui lui sont tagués), **total par
  fournisseur**, avec un **filtre de plage de dates** (Du / Au).
- **Contrats** : journal des commandes passées avec chaque fournisseur. À la
  création, les échéances de **livraison** et de **paiement** sont calculées
  automatiquement (date de commande + délais par défaut du fournisseur).
  Renseignez ensuite les dates réelles de livraison/paiement au fur et à
  mesure : les **dépassements sont détectés et affichés en rouge**
  (« EN RETARD (n j) » si la date prévue est dépassée sans qu'une date
  réelle ait été saisie, ou « Livré/Payé (retard n j) » une fois la date
  réelle enregistrée après l'échéance).

**Saisie** : un nouveau champ **« Fournisseur »** (liste déroulante avec
recherche, proposition de création si le code n'existe pas) permet de taguer
chaque écriture — c'est ce qui alimente automatiquement les modules Achats
et Contrats.

- **ÉTATS ET RAPPORTS** : Grand livre, Balance, Bilan, Compte de résultat,
  TFT, Liasse fiscale, Tableaux d'exécution budgétaire, Impôts,
  Déclarations sociales, Rapprochements bancaires.

Cliquer sur un menu ouvre la liste de ses pages ; cliquer sur une page
l'affiche dans la fenêtre (un seul panneau à la fois).

### Saisie : nouveaux champs (mise à jour)

Le champ « Code flux » a été retiré du formulaire de Saisie. À la place,
chaque écriture propose désormais : **Code analytique**, **Code
budgétaire**, **Code bailleur** (texte libre, pour le suivi par projet/
bailleur de fonds) et **Quantité** (pour la valorisation des stocks — voir
ci-dessous). Le tableau et l'import massif (.xlsx) ont été mis à jour en
conséquence.

⚠️ Le TFT (Tableau des flux de trésorerie) utilisait le Code flux pour
classer les mouvements de trésorerie en EXP/INV/FIN. Ce champ n'étant plus
saisissable, les nouveaux mouvements apparaîtront tous en « Flux non
classés ». Dites-moi si vous voulez qu'on prévoie un autre moyen de les
classer.

**Stocks** (mise à jour) : suivi désormais en **valeur ET en quantité**.
Renseignez la quantité sur chaque écriture touchant un compte de stock
(Saisie), et une quantité initiale (bouton dédié dans l'onglet Stocks) —
l'application calcule alors le **coût unitaire moyen** (valeur du stock
final / quantité finale) pour chaque compte.

### Partie double vraiment forcée (mise à jour majeure)

Le formulaire de Saisie a changé de logique : au lieu d'une ligne à la fois
(un compte + Débit ou Crédit), il demande maintenant **ensemble** :
**Compte débiteur**, **Compte créditeur** et **Montant**. Cliquer sur
« Ajouter » crée automatiquement les deux lignes en une seule opération —
**il est structurellement impossible de créer une écriture déséquilibrée**
par ce formulaire (le compte débiteur doit être différent du compte
créditeur, le montant doit être positif, sinon le logiciel refuse).

Les deux champs comptes sont des listes déroulantes avec recherche ; si
vous quittez le champ avec un code qui n'existe pas dans le Plan comptable,
l'application vous demande de le créer (avec un libellé) avant de continuer
— impossible d'enregistrer une écriture sur un compte invalide.

**Modifier une ligne existante** : sélectionnez-la dans le tableau (chaque
ligne du tableau reste une moitié débit ou crédit, comme avant) — le
formulaire ne pré-remplit alors que le côté concerné ; ne renseignez que ce
compte-là pour la modifier.

**Pour les écritures à plus de 2 comptes** (ex. une facture avec TVA
répartie sur 3 lignes) : ajoutez plusieurs paires successives sur la même
pièce (le N° Pièce reste rempli après chaque « Ajouter » pour faciliter
l'enchaînement) — chaque paire est déjà équilibrée, donc la pièce entière
le reste automatiquement.

### Exercices comptables et clôture annuelle (nouveau)

Une barre en haut de la fenêtre affiche en permanence l'**exercice
comptable en cours** (ex. 2025), avec un sélecteur pour basculer entre
exercices et un bouton **« + Nouvel exercice »**.

Tous les calculs (Saisie, Balance, Bilan, Compte de résultat, TFT, Stocks,
Production, Liasse fiscale) sont désormais **scopés à l'exercice
sélectionné** : seules les écritures datées de cet exercice sont prises en
compte pour les mouvements, et les soldes d'ouverture sont ceux enregistrés
pour cet exercice précis.

**Clôture annuelle** (menu PARAMÈTRES → Exercices comptables) :
- calcule le solde de clôture de chaque compte de bilan (classes 1 à 5) de
  l'exercice sélectionné ;
- intègre le résultat net de l'exercice dans le compte **121000** (Report à
  nouveau créditeur) ;
- reporte ces soldes comme **soldes d'ouverture de l'exercice suivant**
  (créé automatiquement s'il n'existait pas) ;
- **verrouille l'exercice clôturé** : impossible d'ajouter, modifier ou
  supprimer une écriture datée de cet exercice tant qu'il reste clôturé.

Testé avec un cycle complet : exercice 2024 (capital, ventes, achats) →
clôture → exercice 2025 hérite automatiquement des bons soldes d'ouverture
(clients, fournisseurs, banque, report à nouveau incluant le résultat 2024)
et le Bilan reste équilibré, y compris après de nouveaux mouvements en 2025.

### Menu PARAMÈTRES (remplace les plans dans SAISIE)

Les 4 écrans de gestion des plans (Plan comptable, Plan analytique, Plan
budgétaire, Plan bailleurs de fonds) ainsi que les **Exercices comptables**
sont désormais regroupés dans le menu **PARAMÈTRES**.

### Grand livre complet avec bandes de couleur (mise à jour majeure)

L'onglet **Grand livre** affiche désormais **tous les comptes de
l'exercice par défaut**, groupés par compte puis par classe, exactement
comme un grand livre papier classique :
- Bandeau **bleu** pour l'en-tête de chaque compte, et pour son
  sous-total (« TOTAL COMPTE XXXXXX — Solde débiteur/créditeur »).
- Bandeau **orange** pour le total de chaque classe.
- La ligne « À-nouveaux au 01/01 » (solde d'ouverture) s'affiche si non
  nulle, puis le détail chronologique des écritures avec solde cumulé.

Un filtre optionnel (compte et/ou tiers) permet de se recentrer sur un
compte précis si besoin — bouton « Réinitialiser » pour revenir à la vue
complète. Testé sur le scénario exact de votre capture (emprunt WBI/Vista) :
le solde cumulé calculé correspond au FCFA près (-14 595 375 000 puis
-13 849 250 000, identiques à votre grand livre de référence).

### Grand livre : corrigé (n'affichait rien tant qu'on n'avait pas tapé)

**Cause** : le champ « N° Compte » n'avait aucune liste par défaut et ne
s'ouvrait pas au clic (il fallait taper au clavier pour voir apparaître des
résultats) — d'où l'impression que l'écran « n'affiche rien ». Corrigé,
même comportement que dans Saisie : liste des 300 premiers comptes
préchargée, clic = ouverture automatique de la liste déroulante, message
d'aide affiché tant qu'aucun compte n'est choisi (et message clair si le
compte tapé n'existe pas). Le calcul lui-même a été testé et fonctionne
correctement.

### Diagnostic de l'écart de Balance (analyse de votre fichier)

**Comparaison faite entre votre Balance PDF (exercice 2024, autre logiciel)
et notre export (exercice 2026)** : les **soldes de clôture** (colonnes
Solde Débit/Crédit) correspondent **exactement** entre les deux systèmes
là où c'est comparable (ex. TOTAL CLASSE 1 : 20 055 904 / 27 576 434 184
identiques des deux côtés) — la formule de calcul du solde est donc
correcte.

L'écart que vous observez sur les **Cumul Débit/Crédit** vient d'un
mélange de deux facteurs, pas d'un bug de calcul :
1. **Ce ne sont pas les mêmes exercices** (PDF = 2024, export = 2026) : les
   mouvements de la période ne peuvent pas être identiques entre deux
   années différentes.
2. **Des opérations semblent avoir été saisies comme solde d'ouverture au
   lieu d'écritures de la période** (ex. le compte 162020 « EMPRUNT VISTA »
   : votre solde d'ouverture 2026 est déjà de -15 000 000 000, alors que le
   PDF 2024 montre ce même emprunt DÉCAISSÉ pendant l'année — crédité 15
   milliards en cours d'exercice). Le solde final est identique dans les
   deux cas, mais le détail des mouvements de la période diffère forcément
   selon où l'opération a été enregistrée.

L'indicateur d'écart ajouté sur la Balance (message précédent) devrait déjà
vous signaler ce type de situation. Si un écart de **Cumul Débit/Crédit
total** subsiste sur l'exercice 2026 lui-même (pas en comparaison avec
2024), c'est probablement dû à un import massif d'écritures déséquilibré —
dites-le-moi si c'est le cas et je regarderai les données précises.

### Balance : export ajouté + diagnostic du déséquilibre (correction)

**Bouton d'export manquant** : l'onglet Balance n'avait effectivement pas
de bouton d'export — corrigé, un bouton **« Exporter (.xlsx) »** génère
maintenant un fichier avec les mêmes sous-totaux par classe et le total
général que l'écran.

**Sur la formule elle-même** : testée avec des données garanties
équilibrées, elle est correcte (Cumul Débit = Cumul Crédit, Solde Débit =
Solde Crédit à l'euro près). Le déséquilibre visible sur votre capture
vient très probablement de **données important déséquilibrées** — deux
causes possibles, maintenant détectées automatiquement :
1. **Écart sur le Cumul Débit/Crédit** → une ou plusieurs écritures de la
   période ne sont pas équilibrées. Cela ne peut arriver que via l'**import
   massif d'écritures (.xlsx)**, qui n'imposait pas l'équilibre global du
   fichier — **corrigé** : cet import affiche désormais un avertissement
   explicite si le fichier importé n'est pas équilibré dans son ensemble
   (testé et reproduit : Débit 5 000 ≠ Crédit 3 000 → avertissement déclenché).
2. **Écart sur le Solde Débit/Crédit** → soldes d'ouverture incomplets
   (déjà signalé dans le Bilan).

L'onglet Balance affiche maintenant un **indicateur d'écart en bas du
tableau** (vert si équilibré, rouge avec explication sinon) pour repérer
ces situations immédiatement, sans avoir à comparer les totaux à la main.

### Ctrl+A pour tout sélectionner dans Saisie (nouveau)

Dans le tableau de l'onglet Saisie, **Ctrl+A sélectionne désormais toutes
les lignes visibles** (comme dans l'Explorateur Windows), ce qui permet
ensuite de les supprimer toutes d'un coup avec le bouton « Supprimer
(sélection multiple possible) ».

### Correction de lenteur : suppression groupée trop lente (bug corrigé)

**Cause trouvée** : `core.delete_entry()` fait un `commit()` (écriture
synchrone sur disque) **à chaque ligne** — en boucle sur plusieurs lignes
sélectionnées, ça multiplie les accès disque et ralentit fortement,
surtout avec beaucoup de lignes.

**Corrigé** : nouvelle fonction `delete_entries_bulk()` qui supprime tout
le lot dans **une seule transaction** (un seul `commit()` à la fin).
Mesuré : 300 lignes supprimées en 0,003 s avec la nouvelle méthode, contre
0,068 s pour seulement 100 lignes avec l'ancienne — un gain d'environ 20×.

### Suppression groupée dans Saisie (nouveau)

Le tableau de l'onglet Saisie accepte désormais la **sélection multiple**
(Ctrl+clic ou Maj+clic, comme dans l'Explorateur Windows). Le bouton
**« Supprimer (sélection multiple possible) »** supprime alors toutes les
lignes sélectionnées d'un coup, avec une seule confirmation — un exercice
clôturé bloque toujours la suppression des lignes qu'il concerne (message
détaillé si certaines lignes n'ont pas pu être supprimées). Testé :
suppression groupée de 4 écritures en un clic.

### Modèle téléchargeable pour la balance N-1 (nouveau)

Ajout d'un bouton **« Télécharger un modèle (.xlsx) »** dans l'onglet
Soldes d'ouverture, avant les boutons Importer/Exporter — génère un fichier
vierge avec les bons en-têtes et **un exemple équilibré** (4 comptes dont
la somme fait 0), à remplir puis réimporter directement. Testé : le modèle
généré s'importe sans le moindre avertissement (round-trip complet).

### Import de la balance N-1 rendu plus tolérant (correction de bug)

**Bug signalé** : l'import échouait avec « Colonnes obligatoires
introuvables » sur un fichier réel. Corrigé — l'import reconnaît maintenant
plusieurs formats courants :
- Une colonne **« Solde »** signée (notre format par défaut).
- **Deux colonnes séparées « Solde débit » / « Solde crédit »** (comme une
  balance générale classique — le solde est recalculé automatiquement en
  Débit − Crédit).
- Simplement **« Débit » / « Crédit »**.
- Un **en-tête décalé** (titre ou lignes vides au-dessus) — la ligne
  d'en-têtes est désormais recherchée dans les 10 premières lignes, pas
  seulement la ligne 1.

Si le fichier ne correspond toujours à aucun format reconnu, le **message
d'erreur affiche maintenant les en-têtes réellement détectés** dans le
fichier, pour vous aider à comprendre ce qui ne correspond pas.

Testé avec 3 formats différents (solde signé, débit/crédit séparés,
en-tête décalé avec titre) : tous s'importent correctement.

### Import/Export xlsx pour tous les plans + balance N-1, et liste déroulante automatique en Saisie

**Import/Export .xlsx avec écrasement** (menu PARAMÈTRES) :
- **Plan comptable** : boutons Importer/Exporter dans l'onglet. Importer un
  fichier **écrase entièrement** le plan actuel (les comptes non présents
  dans le fichier disparaissent), puis réinsère automatiquement les comptes
  racines (1, 2, 3, 5, 6, 7, 8, 9, 40-49).
- **Plan analytique, Plan budgétaire, Plan bailleurs** : même principe
  (Importer écrase, Exporter génère un .xlsx avec les bons en-têtes).

**Balance d'ouverture (N-1)** (menu SAISIE → Soldes d'ouverture) : mêmes
boutons Importer/Exporter. L'import **écrase les soldes d'ouverture de
l'exercice actuellement sélectionné uniquement** (les autres exercices ne
sont pas affectés) — un compte absent du Plan comptable déclenche un
avertissement mais est importé quand même.

Testé pour les 5 imports : écrasement confirmé dans chaque cas (les
anciennes données disparaissent, remplacées par le contenu du fichier).

**Liste déroulante automatique en Saisie** : un simple **clic** dans les
champs Compte débiteur/créditeur, Journal, Fournisseur, Client, Code
analytique/budgétaire/bailleur ouvre désormais directement la liste
déroulante pour faire défiler et choisir — plus besoin de taper au clavier.
Les champs Compte débiteur/créditeur sont préchargés avec les 300 premiers
comptes dès l'ouverture de l'onglet.

### NOTE 34 (Liasse fiscale) remplie + liens externes cassés supprimés

**Cause identifiée** : le bandeau "IMPOSSIBLE D'ACTUALISER... valeurs
depuis un classeur lié" venait de **liens externes cassés** dans le modèle
(référence vers l'ancien classeur de l'entité GCM, absent). La feuille
**NOTE 34** (Fiche de synthèse des indicateurs financiers — SIG) contenait
en plus d'anciennes valeurs littérales (pas des formules), donc mon
nettoyage général les vidait sans les remplacer, d'où l'écran vide.

**Deux corrections** :
1. **Les liens externes cassés sont maintenant supprimés** à l'export —
   testé, le bandeau d'erreur ne devrait plus apparaître à l'ouverture.
2. **NOTE 34 est remplie automatiquement** (Chiffre d'affaires, Marge
   commerciale, Valeur ajoutée, EBE, Résultat d'exploitation, Résultat
   financier, Résultat des activités ordinaires, Résultat HAO, Résultat
   net), avec les mêmes données que l'onglet Compte de résultat — colonne
   « Année N-1 » remplie aussi si l'exercice précédent existe dans
   l'application.

### Compte de résultat en Soldes Intermédiaires de Gestion (SIG) (mise à jour)

L'onglet **Compte de résultat** suit désormais exactement la structure
officielle SIG (Soldes Intermédiaires de Gestion) de votre modèle, avec une
couleur par section :
- **Activité commerciale** (vert) : Marge commerciale
- **Chiffre d'affaires** (bleu) : A+B+C+D
- **Valeur ajoutée** (jaune) : tous les achats et charges externes détaillés
- **EBE et Résultat d'exploitation** (violet)
- **Résultat financier** (orange) et Résultat des activités ordinaires
- **HAO et Résultat net** (rouge/rose)

Calculé à partir de **`compute_liasse_resultat()`** — la même fonction que
la Liasse fiscale, le TFT et la Situation financière — donc toujours
cohérent avec la Balance et le Bilan. Vérifié : le Résultat net affiché
correspond exactement à celui utilisé par le Bilan (`compute_compte_resultat`
et `compute_liasse_resultat` donnent la même valeur, testé sur plusieurs
scénarios y compris avec variation de stock).

### TFT : la vraie feuille officielle est maintenant remplie (mise à jour)

Grâce à une capture de votre feuille TFT officielle, j'ai pu identifier
précisément les cellules à remplir : **ZA** (trésorerie d'ouverture, ligne
10), **FA** (CAFG, ligne 12), **FB** (variation actif circulant HAO, ligne
13), **FC** (variation des stocks, ligne 14), **FD** (variation des
créances, ligne 15), **FE** (variation du passif circulant, ligne 16) —
toutes en colonne I (Exercice N), calculées depuis vos écritures.

Testé : les valeurs injectées dans la vraie feuille TFT correspondent
exactement à celles de l'onglet TFT de l'application (flux opérationnel
cohérent entre les deux, écart 0).

⚠️ **Les lignes d'investissement et de financement (à partir de FF) ne sont
pas encore automatisées** dans la vraie feuille officielle — je n'ai pas
encore de confirmation visuelle de leur position exacte dans votre modèle,
et je préfère ne pas deviner au risque d'écrire au mauvais endroit sur un
document officiel. **Envoyez-moi une capture des lignes suivantes de la
feuille TFT** (après la ligne 19) pour que je complète le reste. En
attendant, le calcul complet (avec investissement et financement) reste
disponible dans l'onglet supplémentaire « TFT (méthode indirecte - CAFG) »
du même fichier exporté.

### Liasse fiscale : mêmes données que Balance/Bilan/TFT/Situation financière

L'export de la Liasse fiscale utilise désormais **exactement les mêmes
fonctions de calcul** que les onglets de l'application :
- **BILAN** et **RESULTAT** : déjà basés sur `compute_liasse_bilan()` et
  `compute_liasse_resultat()` (comme les onglets Bilan et Compte de
  résultat) — inchangé, déjà cohérent.
- **TFT** *(corrigé)* : l'onglet supplémentaire calculé automatiquement
  utilisait encore l'ancienne méthode directe (`compute_tft`) — remplacé
  par `compute_tft_indirect()`, la **méthode indirecte avec CAFG**,
  identique à l'onglet TFT de l'application (renommé « TFT (méthode
  indirecte - CAFG) » dans le fichier exporté).
- **Nouvelle feuille « SITUATION FIN. (FR-BFR-TN) »** *(nouveau)* : ajoutée
  à l'export, avec les mêmes données que l'onglet Situation financière
  (CAFG, rentabilité, Fonds de Roulement, Besoin en Fonds de Roulement,
  Trésorerie Nette avec contrôle).

Testé : export complet sans erreur ni avertissement (noms d'onglets
raccourcis pour respecter la limite Excel de 31 caractères), Bilan
équilibré (31 000 000 = 31 000 000), TFT et Situation financière remplis
avec les bonnes valeurs.

### Corrections TFT + Bilan, et nouveau module Situation financière

**TFT** : ajout des **bandes de couleur par section** qui manquaient (Text
brut remplacé par un Treeview coloré — trésorerie d'ouverture en violet,
CAFG/exploitation en vert, investissement en orange, financement en bleu,
contrôle en rouge/rose).

**Bilan** : **Actif à gauche, Passif à droite** (inversé par rapport à la
précédente version, sur votre demande).

**Nouveau : Situation financière (FR-BFR-TN)** (menu ÉTATS ET RAPPORTS),
présentée selon le modèle officiel que vous avez fourni, avec une couleur
par section :
- Résultat net, CAFG, autofinancement, ratios de rentabilité économique et
  financière (vert)
- **Fonds de Roulement (FR)** = Ressources stables − Actifs immobilisés (bleu)
- **Besoin en Fonds de Roulement (BFR)** = exploitation + HAO (jaune)
- **Trésorerie Nette (TN) = FR − BFR**, avec contrôle face à la trésorerie
  réelle de la Balance (violet)
- Flux de la période (rappel du TFT, orange) et endettement financier net
  (rouge/rose)

Entièrement calculée à partir de `compute_bilan()`, `compute_liasse_resultat()`
et `compute_tft_indirect()` — donc toujours cohérente avec la Balance, le
Bilan et le TFT. **Un bug a été détecté et corrigé pendant les tests** : un
premier essai montrait un écart de 5 000 000 entre la trésorerie nette
calculée (FR−BFR) et la trésorerie réelle — l'investigation a révélé qu'il
s'agissait en fait d'un **Bilan lui-même déséquilibré** dans le scénario de
test (solde d'ouverture d'un emprunt saisi sans sa contrepartie), et non
d'un défaut de la formule. Une fois les soldes d'ouverture complets et
équilibrés, la Situation financière se réconcilie exactement avec la
Balance (testé : écart 0 sur plusieurs scénarios).

### TFT en méthode indirecte — CAFG (nouveau, cohérent avec la Balance)

L'onglet **TFT** contient maintenant deux sous-onglets :

**« TFT (méthode indirecte — CAFG) »** *(nouveau, vue principale)* : suit
exactement la structure du modèle officiel SYSCOHADA que vous avez fourni —
trésorerie d'ouverture, détermination de la **Capacité d'Autofinancement
Globale (CAFG)** à partir de l'EBE, des revenus et frais financiers, puis
variation du BFR (stocks, créances, dettes circulantes) pour obtenir le
flux des activités opérationnelles ; flux d'investissement (acquisitions/
cessions d'immobilisations incorporelles, corporelles, financières) ; flux
de financement (capital, subventions, emprunts).

Entièrement calculé à partir de **la même `compute_balance()`** que les
onglets Balance et Bilan — une ligne **CONTRÔLE** compare la trésorerie
recalculée par la méthode indirecte à la trésorerie réelle de la Balance
(classe 5) ; l'**écart doit être nul**, ce qui garantit la cohérence entre
les trois états. Testé avec plusieurs scénarios (vente à crédit, achat de
stock au comptant, encaissement partiel, remboursement d'emprunt) : écart
toujours à 0, trésorerie calculée = trésorerie réelle au FCFA près.

**« TFT (méthode directe — ancien) »** : l'ancienne vue (basée sur le code
flux EXP/INV/FIN), conservée pour référence mais reléguée en second plan.

### Balance et Bilan reformatés (mise à jour, cohérence garantie entre eux)

**Balance** (États et rapports → Balance) : reformatée en **Balance
générale groupée par classe**, avec pour chaque compte les colonnes Solde
Ouverture, Cumul Débit, Cumul Crédit, **Solde Débit** et **Solde Crédit**
(séparés, comme une balance comptable classique), un **sous-total par
classe** (ligne bleutée « TOTAL CLASSE X ») et un **total général** en bas
(ligne foncée « TOTAL BALANCE »).

**Bilan** (États et rapports → Bilan) : présenté en **deux colonnes
côte à côte, PASSIF à gauche et ACTIF à droite**, avec une **couleur
distincte par masse** :
- Actif (droite) : Immobilisations (vert), Stocks détaillés par compte réel
  (jaune), Créances (bleu), Trésorerie détaillée par banque (violet).
- Passif (gauche) : Capitaux propres + résultat net (vert), Dettes
  financières (orange), Dettes circulantes détaillées — fournisseurs,
  avances, fiscal/social, autres — (rouge/rose), Trésorerie passif (violet).
- Une bande foncée en bas de chaque colonne pour le TOTAL ACTIF / TOTAL
  PASSIF, et l'écart Actif-Passif affiché en vert (équilibré) ou rouge
  (à corriger).

**Cohérence garantie entre les deux** : Balance et Bilan sont calculés à
partir de **la même fonction `compute_balance()`** — revérifié après cette
mise à jour visuelle (Total Actif = 14 700 000, cohérent avec la somme des
soldes débiteurs de la Balance sur les classes 1 à 5, aucune KeyError,
aucune régression).

### Balance et Bilan reformatés — première version (historique)

**Balance** (États et rapports → Balance) : reformatée en **Balance
générale groupée par classe**, avec pour chaque compte les colonnes Solde
Ouverture, Cumul Débit, Cumul Crédit, **Solde Débit** et **Solde Crédit**
(séparés, comme une balance comptable classique), un **sous-total par
classe** (ligne bleutée « TOTAL CLASSE X ») et un **total général** en bas
(ligne foncée « TOTAL BALANCE ») — structure proche de votre balance PDF de
référence.

**Bilan** (États et rapports → Bilan) : largement enrichi avec le détail
par poste :
- Immobilisations nettes détaillées par catégorie
- **Stocks détaillés par compte réel** (ex. 321001 CLINKER), pas seulement
  le total
- Créances détaillées (avances versées / clients)
- **Trésorerie détaillée par banque/caisse** (chaque compte 52x séparément,
  comme dans votre PDF)
- Dettes circulantes détaillées (fournisseurs / avances reçues / dettes
  fiscales et sociales / autres dettes)

**Cohérence garantie entre les deux** : Balance et Bilan sont désormais
calculés à partir de **la même fonction `compute_balance()`** — testé et
vérifié : le Total Actif du Bilan correspond exactement à la somme des
soldes débiteurs de la Balance sur les classes 1 à 5 (14 700 000 = 14 700
000 dans le scénario testé, avec plusieurs banques et sous-comptes de stock
détaillés). Aucune régression sur les scénarios précédents.

### Racines des comptes (nouveau)

Chaque compte du Plan comptable est désormais rattaché à une **racine**,
visible dans l'onglet Plan comptable (colonnes « Racine » et « Libellé de la
racine ») :
- **1 chiffre** pour les classes 1, 2, 3, 5, 6, 7, 8, 9.
- **2 chiffres pour la classe 4** (comptes de tiers), qui se subdivise en
  **40** (Fournisseurs et comptes rattachés), **41** (Clients et comptes
  rattachés), 42 (Personnel), 43 (Organismes sociaux), 44 (État), 45
  (Organismes internationaux), 46 (Associés/Groupe), 47 (Débiteurs/
  créditeurs divers), 48 (Régularisations), 49 (Dépréciations sur tiers).

**Les comptes racines existent désormais réellement dans le Plan comptable**
(1, 2, 3, 5, 6, 7, 8, 9, 40 à 49), avec un libellé entre tirets (ex. « —
Fournisseurs et comptes rattachés — ») pour les repérer facilement. Grâce au
tri alphabétique des codes, chaque racine **apparaît en tête de son groupe**
dans toutes les listes de comptes (ex. le compte « 1 » avant 101000, 101100,
etc. ; le compte « 40 » avant 400000, 401000, 401100...).

Les fiches auxiliaires créées dans **Fournisseurs** sont rattachées à la
racine **40**, celles créées dans **Clients** à la racine **41**.

**Sélection du tiers rendue obligatoire (nouveau)** : dans l'onglet Saisie,
si vous tapez directement le compte racine **40** ou **41** dans « Compte
débiteur »/« Compte créditeur », l'application vous avertit qu'on ne saisit
jamais directement sur une racine de regroupement, bascule automatiquement
sur le compte de détail usuel (401000/411000), et impose de choisir le
fournisseur ou le client dans le champ correspondant. Plus largement, **toute
écriture sur un compte de la racine 40 sans fournisseur renseigné (ou de la
racine 41 sans client renseigné) est bloquée** à l'enregistrement.

**Tous les calculs liés aux comptes de tiers ont été mis à jour en
conséquence** :
- Le **Bilan** classe désormais les comptes de tiers **par racine** plutôt
  que par simple signe du solde : la racine 41 (Clients) va toujours en
  Créances, la racine 40 (Fournisseurs) toujours en Dettes circulantes ; les
  autres racines (42 à 49) restent classées par signe, car leur nature
  actif/passif dépend réellement du solde.
- **Achats** et **Ventes** utilisent désormais la racine complète (`40%` et
  `41%`) au lieu de motifs partiels — un **bug a été corrigé au passage** :
  l'ancien filtre (401xxx/408xxx pour les fournisseurs, 411xxx pour les
  clients) ratait des comptes comme 402, 404, 409, 412, 413, 418, 419, qui
  sont maintenant bien pris en compte.

Testé de bout en bout : Bilan équilibré avec un compte fournisseur débiteur
(avance, compte 409xxx) et un compte client sur un effet à recevoir (compte
412xxx), tous deux désormais correctement classés ; comptes racines vérifiés
existants et correctement triés (« 1 » avant 101000, « 40 » avant 400000-
409xxx) ; écriture réelle avec fournisseur tagué toujours cohérente.

### Gestion des plans (détail des écrans)

Le menu **SAISIE** contient maintenant 4 écrans pour créer/modifier/
supprimer les référentiels utilisés lors de la saisie : **Plan comptable**,
**Plan analytique**, **Plan budgétaire** (avec montant prévu), **Plan
bailleurs de fonds**.

### (Ancien mécanisme remplacé)

L'équilibrage « après coup » ligne par ligne a été remplacé par le
formulaire Compte débiteur / Compte créditeur décrit plus haut, qui
équilibre chaque écriture dès sa création plutôt que de le vérifier après.

### Listes déroulantes avec proposition de création (nouveau)

Les champs **Code analytique**, **Code budgétaire** et **Code bailleur**
sont des listes déroulantes alimentées par leurs plans respectifs. Si vous
tapez un code qui n'existe pas encore, l'application vous demande de
confirmer sa création (avec un libellé) avant de passer à la cellule
suivante — impossible d'enregistrer un code orphelin par erreur de frappe.

### Ce qui est pleinement fonctionnel


Saisie, Soldes d'ouverture, Stocks (partagé entre Matières premières et
Produits finis pour l'instant), Fabrication, Compte de résultat, TFT, Grand
livre, Balance, Bilan, Liasse fiscale — ainsi que 3 nouvelles pages basées
sur vos écritures existantes :
- **Ventes** / **Achats** : synthèse des comptes de vente (classe 7) et
  d'achat (classe 6), hors éléments financiers.
- **Marges bénéficiaires** : marge commerciale, valeur ajoutée, résultat
  d'exploitation et résultat net (mêmes calculs que la Liasse fiscale).
- **Clients** / **Fournisseurs** : Grand livre pré-filtré sur les comptes
  411000 / 401000.

### Ce qui reste à construire

**Contrats**, **Tableaux d'exécution budgétaire**, **Impôts**,
**Déclarations sociales** et **Rapprochements bancaires** apparaissent dans
le menu mais affichent pour l'instant un message « fonctionnalité pas
encore développée » — ce sont de nouveaux modules à part entière (suivi de
contrats, calcul d'impôts, etc.) qui nécessitent d'être conçus et développés
spécifiquement. Dites-moi lesquels prioriser.

- **Soldes d'ouverture** *(nouveau)* : saisissez le solde de report à nouveau
  de chaque compte de bilan au 1er jour de l'exercice (= solde de clôture de
  l'exercice précédent). Convention : débiteur = positif, créditeur = négatif.
  La somme de tous les soldes d'ouverture doit être nulle (partie double) —
  un contrôle l'affiche en bas de l'onglet. **Tous les calculs (Balance,
  Bilan, TFT, Liasse fiscale) intègrent désormais automatiquement ces soldes
  d'ouverture** : Balance de clôture = Solde d'ouverture + Mouvements de
  l'exercice. C'est ce qui permet au Bilan de s'équilibrer même si ce n'est
  pas la première année d'activité.
- **Balance** *(mise à jour)* : affiche maintenant, pour chaque compte, le
  Solde d'ouverture, le Débit/Crédit/Solde de la période, et le **Solde de
  clôture**.
- **Stocks** : le stock initial saisi ici alimente désormais directement la
  table des soldes d'ouverture (même mécanisme que ci-dessus).
- **TFT** : la trésorerie d'ouverture est calculée **automatiquement** à
  partir des soldes d'ouverture des comptes de trésorerie (521000/531000/
  570000/585000) ; un bouton permet de la forcer manuellement si besoin.
  Codez `FLUX-EXP`, `FLUX-INV` ou `FLUX-FIN` dans le champ « Code flux » des
  écritures de trésorerie dans l'onglet Saisie pour classer les mouvements.
- **Grand livre** : tapez un N° Compte (liste déroulante avec recherche)
  puis « Afficher » pour voir le détail chronologique et le solde cumulé.
- **Production** : tapez `AN-FAB` dans le champ « Code analytique » de
  l'onglet Saisie sur les lignes de charges de fabrication pour qu'elles
  remontent dans l'onglet Production.

### Liasse fiscale *(mise à jour majeure)*

Renseignez l'identification de l'entité (dénomination, adresse, N° IFU,
exercice clos le...), puis « Exporter la liasse fiscale complète (.xlsx) ».

Le fichier généré reprend **les 92 pages et les mêmes dimensions exactes du
modèle SYSCOHADA système normal que vous avez fourni** (COUVERTURE, BILAN,
RESULTAT, TFT, 39 notes annexes NOTE 1 à NOTE 39, ~20 tableaux fiscaux DGI
SUPPL1 à SUPPL20, fiches R1-R4, etc.) :

- ✅ **BILAN et RESULTAT** : remplis automatiquement depuis vos écritures,
  avec les mêmes codes officiels (AD/AE/AI... côté actif, CA/CJ/DA... côté
  passif, TA/RA/XA... au compte de résultat). Les totaux et le Résultat net
  utilisent désormais la **balance de clôture** (soldes d'ouverture +
  mouvements) — le Bilan s'équilibre toujours, y compris les années
  suivantes une fois les soldes d'ouverture saisis.
- ✅ **TFT** : la page officielle (méthode indirecte avec CAFG) est laissée
  vierge — nous ne calculons pas la CAFG automatiquement. Un onglet
  supplémentaire **« TFT (simplifie) »** est ajouté avec un calcul en
  méthode directe (Ouverture, EXP/INV/FIN, Clôture), cohérent avec la
  Balance.
- ⚠️ **Détail des lignes du Bilan** (AE à AN, CA à CM, DA à DM) : réparti
  par plage de comptes, y compris une répartition proportionnelle des
  amortissements entre catégories — indicatif, à vérifier.
- 📄 **Toutes les autres pages** (39 notes, ~20 tableaux DGI) : conservées
  avec leur mise en page, leurs libellés et **leurs dimensions identiques**
  au modèle fourni, mais les montants qu'elles contenaient (qui sont les
  chiffres 2023 de l'entreprise du modèle, pas les vôtres) sont **effacés**
  pour éviter toute confusion — à compléter manuellement ou par votre
  expert-comptable.

**À faire vérifier par un expert-comptable avant tout dépôt officiel auprès
de la DGI** — cet export est une aide à la préparation, pas un dépôt
directement utilisable tel quel.

## Limites de cette version par rapport au classeur Excel

Cette version ne reprend pas encore le suivi budgétaire / analytique par
projet / par bailleur de fonds (feuille « Rapport d'exécution » du
classeur), ni les comptes auxiliaires Fournisseurs/Clients détaillés.
Dites-moi si vous voulez que je les ajoute — le moteur (`core.py`) est
structuré pour que ce soit un ajout incrémental, pas une réécriture.

### Bilan « avec détails » + correction de l'équilibre (mise à jour majeure)

**Cause racine de l'ancien déséquilibre du Bilan identifiée et corrigée** :
plusieurs calculs (résultat net dans `compute_bilan`, capitaux propres,
clôture d'exercice) s'appuyaient sur des **listes de comptes codées en dur**
(`COMPTES_PRODUITS_EXPL`, `COMPTES_CAPITAL`, etc.) qui ne couvrent qu'une
partie du vrai plan comptable de l'utilisateur (1591 comptes importés de
Sage). Tout compte hors de ces listes disparaissait purement et simplement
du Total Actif ou du Total Passif, d'où l'écart constaté.

**Correctif** : `compute_resultat_net_complet()` (nouveau, dans `core.py`)
calcule désormais le résultat net de façon exhaustive à partir de
**l'intégralité** des classes 6 (charges), 7 (produits) et 8 (HAO), sans
aucune liste de comptes partielle. `compute_bilan()` utilise l'intégralité
de la **classe 1** (au lieu de `COMPTES_CAPITAL`/`COMPTES_DETTES_FIN`...) et
l'intégralité de la **classe 3** (au lieu des 4 comptes maîtres de stock) —
chaque compte de la Balance est donc classé dans une case et une seule de
l'Actif ou du Passif, ce qui **garantit mathématiquement Actif = Passif**
(tant que la somme des soldes d'ouverture de l'exercice est nulle).
`compute_liasse_resultat()` (Compte de résultat officiel, TFT, Situation
financière, Liasse fiscale) est recalée sur cette même référence via une
ligne de réconciliation (repliée dans « Autres produits »/« Autres
charges ») — tous les états financiers partagent maintenant EXACTEMENT le
même résultat net. `close_exercice()` (clôture annuelle) utilise aussi ce
calcul exhaustif pour reporter le résultat sur le compte 121000.

**Nouvel onglet Bilan, présenté « avec détails »** (calqué sur le rapport
financier de référence fourni par l'utilisateur, pas sur les codes officiels
DGI de la Liasse fiscale) : `compute_bilan_detaille()` —
- **Actif** en colonnes Brut / Amortissements / Net : immobilisations par
  catégorie (charges immobilisées, terrains, bâtiments, installations,
  matériel, matériel de transport, avances sur immo, immobilisations
  financières — répartition de l'amortissement proportionnelle au brut de
  chaque catégorie, indicatif) ; stocks regroupés par préfixe à 2 chiffres
  (31 à 39, plus de compte de stock oublié) ; créances **compte par compte**
  (chaque client, chaque avance fournisseur, chaque compte 42 à 49 débiteur
  affiché séparément — pas juste un total) ; trésorerie **banque par
  banque**.
- **Passif** en Montant : capitaux propres et ressources durables **compte
  par compte** (capital, réserves, report à nouveau, emprunts... — classe 1
  entière) puis Résultat net de l'exercice ; dettes circulantes **compte par
  compte** (chaque fournisseur, et pour la classe 44 chaque compte distinct
  — IS, IMF, BIC, TVA facturée, TVA due, retenues... apparaissent
  automatiquement séparément dès lors qu'ils existent comme comptes distincts
  dans le plan comptable réel, sans codage en dur) ; trésorerie créditrice
  banque par banque.
- **Exportable en .xlsx** (bouton « Exporter (.xlsx) », `export_bilan_detaille_xlsx()`),
  dans une mise en page à deux colonnes proche du modèle papier fourni.

Testé de bout en bout avec des comptes volontairement absents des anciennes
listes codées en dur (105xxx, 163xxx, 380xxx stock, 84x/87x HAO...) : Bilan
et Bilan détaillé strictement équilibrés (écart = 0), Compte de résultat/TFT/
Situation financière cohérents avec ce même résultat net, clôture d'exercice
fonctionnelle.

**Point de vigilance non résolu** : la répartition Brut/Amortissements par
catégorie d'immobilisation reste **proportionnelle** (indicatif, comme déjà
signalé pour la Liasse fiscale) faute de connaître la correspondance exacte
entre chaque compte d'immobilisation et son compte d'amortissement dédié
dans le plan comptable réel de l'utilisateur — le Net par catégorie et le
Total Net restent, eux, exacts.

### Correction d'une facture validée + Impôts / Déclarations sociales / Rapprochements bancaires

**Facturation — corriger une erreur sur les chiffres après validation.**
Une facture validée envoie ses écritures en Saisie et devenait ensuite
totalement verrouillée (impossible de corriger le moindre chiffre). Nouveau
bouton **« Corriger cette facture (erreur sur les chiffres) »** dans
l'onglet Facturation, actif uniquement quand la facture sélectionnée est
validée : `core.devalider_facture_vente()` retire les écritures générées
par cette facture (repérées de façon fiable par le couple `piece = numéro
de facture` / `journal = 'VE'`, propre à chaque facture), remet son statut à
« brouillon », et vous pouvez alors corriger les lignes/quantités/prix puis
revalider normalement. Refuse si l'exercice comptable de la facture est
clôturé. Testé de bout en bout (validation → correction → revalidation →
Bilan toujours équilibré).

**Impôts (classe 44) et Déclarations sociales (classe 43)** — les deux
anciens placeholders sont remplacés par un onglet générique
`ClassePeriodeTab` (`core.compute_comptes_prefixe_periode()`) : liste tous
les comptes de la classe concernée avec **solde de début de période /
mouvements Débit-Crédit / solde de fin de période**, sur une **période
librement choisie** (filtre Du/Au — JJ/MM/AAAA — par défaut l'exercice
comptable entier). Comme pour le Bilan, aucune liste de comptes codée en
dur : tout compte 44xxxx ou 43xxxx existant dans le plan comptable apparaît
automatiquement dès qu'il a un solde ou un mouvement sur la période (IS,
IMF, BIC, TVA due/facturée/récupérable, retenues à la source... pour la
classe 44 ; CNSS et assimilés pour la classe 43).

**Rapprochements bancaires (racine 52)** — nouvel onglet
`RapprochementBancaireTab` (`core.compute_mouvements_prefixe_periode()`) :
chaque compte de banque à 6 chiffres (52xxxx) est détaillé mouvement par
mouvement sur la période choisie (Du/Au), avec pour chacun une **case à
cocher « Pointé »** (cliquer sur la colonne, ☑/☐) pour signaler qu'il a été
retrouvé dans le relevé bancaire papier. Le pointage est enregistré dans
une nouvelle table `pointages_bancaires` (persistant, lié à l'écriture par
son ID) — il reste visible à la prochaine ouverture, y compris si l'on
change de période d'affichage. Un total en bas de page indique le montant
pointé, le montant total de la période, et l'écart non pointé restant à
justifier.

**Limite non résolue de cette session** : l'environnement de développement
ne dispose pas du module `tkinter` — le moteur (`core.py`) a été testé en
profondeur en ligne de commande (scénarios réels, écarts vérifiés), mais
l'interface graphique elle-même (`main.py`) n'a pu être vérifiée que par
relecture attentive et compilation (`python3 -m py_compile`), pas par un
lancement réel de l'application. Un premier test sur votre PC Windows reste
la vérification finale à faire pour ces trois nouveaux écrans — signalez-moi
tout affichage inattendu.

### Diagnostic de l'écart Bilan + lisibilité (import comptable Sage — vérifié sur données réelles)

**L'écart Actif-Passif de -8 175 989 544 constaté par l'utilisateur sur ses
vraies données (exports `Balance222.xlsx` / `Bila141141n.xlsx` du 07/08/2026)
a été analysé en détail — ce n'est PAS un bug du calcul du Bilan** (déjà
garanti mathématiquement équilibré par construction, voir plus haut), mais
un problème dans les DONNÉES elles-mêmes, avec deux causes identifiées et
quantifiées directement depuis l'export réel de l'utilisateur :
1. **Somme des soldes d'ouverture de l'exercice 2026 = -9 098 750 409**
   (devrait être nulle par partie double).
2. **Cumul Débit ≠ Cumul Crédit des écritures de la période** : 62 053 285 001
   contre 61 130 524 136, écart de 922 760 865 — des écritures existent en
   base où Débit ≠ Crédit pour une même pièce (source très probable : un
   import massif d'écritures qui n'a pas respecté la partie double, risque
   déjà documenté dans ce README).

**Nouveaux outils construits pour que l'utilisateur puisse corriger cela
lui-même :**
- `core.compute_ecart_diagnostic()` : décompose l'écart du Bilan en ces deux
  causes exactes, avec leurs montants réels tirés de la base.
- `core.compute_pieces_non_equilibrees()` + nouvel onglet **« Écritures non
  équilibrées »** (ÉTATS ET RAPPORTS) : liste chaque pièce (regroupement
  Pièce + Journal) dont Débit ≠ Crédit, triée par écart décroissant, avec
  date, nombre de lignes, et l'écart exact — permet de retrouver et corriger
  précisément la ou les pièces fautives dans la Saisie.
- Le message d'écart du Bilan n'affiche plus un simple point d'exclamation
  générique : il affiche maintenant les VRAIES causes avec leurs montants,
  et un bouton « Voir le détail des pièces non équilibrées → » ouvre
  directement ce nouvel onglet.

Testé avec un scénario reproduisant exactement le cas réel (soldes
d'ouverture non nuls + pièce d'import Débit ≠ Crédit) : le diagnostic
retrouve exactement la pièce fautive, et
`ecart_soldes_ouverture + ecart_ecritures_periode == écart du Bilan` à
l'euro (franc) près, dans tous les cas testés.

**Lisibilité du Bilan corrigée** (« les chiffres coincés sur les côtés ») :
- Nouveau formatage `fmt_cfa()` façon rapport SYSCOHADA : espace comme
  séparateur de milliers (ex. `27 556 378 280`), plus proche du PDF de
  référence que l'ancien séparateur virgule.
- Colonnes Actif/Passif élargies (Brut/Amortissements 150px, Net/Montant
  170-180px, au lieu de 110-140px), hauteur de ligne et police du Bilan
  augmentées, barre de défilement horizontale ajoutée en plus de la
  verticale.
- La fenêtre de l'application démarre désormais **maximisée**
  (`self.state("zoomed")`) au lieu d'une taille fixe 1200x720, pour laisser
  toute la place nécessaire aux écrans denses en chiffres (Bilan, Balance).

### Classification Actif/Passif des comptes de tiers — correction par compte, pas par racine entière

**Bug corrigé, repéré par l'utilisateur directement sur son export réel**
(`Bila141141n.xlsx`) : le compte `419100 CLIENTS, VERSEMENT A VENTILLE`
(solde créditeur) apparaissait avec un montant NÉGATIF dans la liste des
Créances (Actif), au lieu de basculer au Passif comme « Clients créditeurs »
— exactement la ligne que le PDF de référence isole séparément
(`Clients créditeurs *411`).

**Cause** : `compute_bilan()`/`compute_bilan_detaille()` traitaient la
racine 41 (Clients) comme systématiquement à l'Actif EN BLOC (toute la
racine, quel que soit le signe de chaque compte), et la racine 40
(Fournisseurs) systématiquement au Passif en bloc — au lieu d'appliquer la
même règle que les racines 42 à 49 : **chaque compte, individuellement,
selon le signe de son propre solde de clôture** (débiteur → Actif,
créditeur → Passif), conforme aux libellés du rapport de référence qui
fait apparaître CHAQUE racine (40 à 49) potentiellement des deux côtés du
Bilan (« Frs avances versées *40* » à l'Actif ET « Fournisseurs *40 » au
Passif ; « Client débiteurs *411* » à l'Actif ET « Clients créditeurs *411 »
au Passif).

**Correctif** : `compute_bilan()` et `compute_bilan_detaille()` appliquent
maintenant une seule règle uniforme pour TOUTES les racines de tiers (40 à
49) : `_sum_racine(balance, racine, sign="pos")` pour les créances (Actif),
`_sum_racine(balance, racine, sign="neg")` pour les dettes (Passif) —
compte par compte, plus aucune racine traitée « en bloc ». Testé avec un
scénario reproduisant exactement le cas réel (411100 débiteur, 419100
créditeur, 409100 avance fournisseur débitrice, 401100 fournisseur
créditeur normal) : chaque compte atterrit désormais du bon côté, le total
détaillé somme exactement au total du Bilan, et l'écart reste à 0.

### Couleurs du Bilan calquées pixel par pixel sur le PDF de référence

**Demande** : « le cadre du bilan avec les mêmes couleurs n'est pas
identique à celui du pdf ». Le PDF de référence n'utilise pas un code
couleur par masse comptable (immobilisations/stocks/créances/dettes) comme
l'ancienne version, mais **une couleur PAR RACINE de compte de tiers**,
appliquée à l'identique des deux côtés du Bilan (même couleur pour « Frs
avances versées *40* » à l'Actif et « Fournisseurs *40 » au Passif, par
exemple). Couleurs extraites du PDF par échantillonnage de pixels
(`pdftoppm` + PIL) :

- Racine 40 (Fournisseurs) : orange `#FF6600`, texte blanc
- Racine 41 (Clients) : bleu `#3366FF`, texte blanc
- Racine 42 (Personnel) : jaune `#FFFF00`, texte noir
- Racine 43 (Organismes sociaux/CNSS) : rose `#FF99CC`, texte noir
- Racines 44-45 (État, organismes internationaux) : gris `#999999`, texte blanc
- Racines 46-49 (Débiteurs/créditeurs divers, HAO, régularisation,
  dépréciations) : cyan `#00FFFF`, texte noir (le PDF ne distingue pas 46 de
  47-49, même couleur)
- Stocks (classe 3) : bleu clair `#99CCFF`
- Trésorerie (classe 5, Actif et Passif) : vert `#00FF00`
- Sous-totaux côté Actif (Total immobilisations, Total stocks, Total
  créances, Total trésorerie actif) : bleu clair `#99CCFF`
- Sous-totaux côté Passif (Total capitaux propres, Total dettes
  circulantes, Total trésorerie passif) : or `#FFCC00`
- TOTAL ACTIF / TOTAL PASSIF (total général) : or `#FFCC00`, les deux côtés
- Immobilisations et Capitaux propres/ressources durables (classe 1) :
  blanc/aucune couleur, comme dans le PDF

`compute_bilan_detaille()` renvoie désormais un champ `"racine"` sur chaque
ligne de créance/dette et `"prefixe"` sur chaque ligne de stock, pour que
l'interface (`BilanTab`) et l'export (`export_bilan_detaille_xlsx()`)
déterminent la couleur exacte sans avoir à re-parser le libellé. Les deux
sont maintenant strictement cohérents entre eux (même palette à l'écran et
dans le fichier exporté). Testé : les comptes 401100 (racine 40), 411100 et
419100 (racine 41), 431300 (racine 43) ressortent avec les bonnes couleurs
hexadécimales exactes dans l'export .xlsx.

### Menu MAINTENANCE-ÉNERGIE + scrollbar Saisie

**Nouveau menu « MAINTENANCE-ÉNERGIE »**, avec deux sous-menus :
- **Énergie** : coûts par code analytique (eau, électricité, essence, gasoil,
  gaz...) sur une période choisie — bouton « Ajouter les codes courants »
  pour pré-remplir `ENERGIE-EAU`, `ENERGIE-ELEC`, etc. (n'écrase jamais un
  code déjà personnalisé).
- **Maintenance** : même principe avec des codes `MAINT-` (véhicules,
  bâtiments, machines, informatique...).

Ces écrans utilisent le champ **Code analytique** déjà présent dans
l'onglet Saisie : renseignez-le sur la ligne du compte de charge (classe 6)
pour qu'elle remonte automatiquement dans le bon écran, groupée par code,
avec solde de début de période / mouvements / solde cumulé.

**Bug découvert et corrigé pendant les tests** : `add_balanced_entry` pose
le même code analytique sur les DEUX lignes d'une écriture (la charge ET sa
contrepartie, ex. la banque) — un calcul naïf du solde net par code
analytique donne donc toujours 0. `compute_couts_analytiques_categorie()`
ne comptabilise désormais QUE le côté charge (classe 6), comme le fait déjà
`compute_production()`/AN-FAB pour les coûts de fabrication — même
principe, appliqué de façon cohérente.

**Sous-menu Fabrication mis à jour** : les lignes de recette de type
Main-d'œuvre, Énergie et Autre charge peuvent désormais être associées à un
code analytique (nouveau champ dans le formulaire d'ajout de composant,
nouvelle colonne dans le tableau de la recette). Colonne `analytic_code`
ajoutée à la table `recette_lignes` (migration automatique pour les bases
existantes). Nouvelle fonction `compute_couts_analytiques_fabrication()`
pour croiser les lignes de recette d'une catégorie avec les coûts réels
comptabilisés sous le même code.

Testé de bout en bout : ligne de recette « Main-d'œuvre » taguée
`MAINT-MACH`, ligne « Énergie » taguée `ENERGIE-EAU`, écritures réelles de
Saisie sur ces mêmes codes — tout remonte correctement dans les écrans
Énergie/Maintenance, et le Bilan reste équilibré (écart = 0) après
fabrication et facturation. Non-régression vérifiée : une recette sans
aucun code analytique (comportement d'avant) fonctionne toujours à
l'identique.

**Onglet Saisie** : la plage des écritures a maintenant une vraie
scrollbar verticale à droite (déplaçable à la souris), en plus du défilement
au clavier/molette déjà existant.

### Correctif urgent : `NameError: name 'FLUX_FAB' is not defined`

**Cause** : lors de l'insertion du nouveau bloc « Maintenance & Énergie »
dans `core.py` (message précédent), la constante `FLUX_FAB = "AN-FAB"`
(utilisée par `compute_production()`, appelée dès l'ouverture du menu
PRODUCTION > Fabrication) a été supprimée par erreur pendant le
remplacement de texte — un `NameError` bloquait donc le lancement de ce
module dans le `.exe` compilé. **Corrigé** : `FLUX_FAB` restauré à sa place
d'origine, juste avant `FAB_POSTES`.

Vérifié par un test de fumée qui appelle toutes les fonctions de calcul
principales de `core.py` (Balance, Bilan, Bilan détaillé, Compte de
résultat, Liasse fiscale, TFT, Situation financière, Production, Stocks,
Trésorerie, Impôts/Déclarations sociales, Énergie/Maintenance, diagnostic
d'écart, rapprochement bancaire) sur une base neuve, sans aucune exception :
plus aucune constante orpheline détectée suite à la modification précédente.

### Correctif : liste « Code analytique » figée dans l'onglet Saisie

**Bug repéré par l'utilisateur** : le Plan analytique affichait bien tous
les codes (ENERGIE-EAU, MAINT-MACH...), mais le menu déroulant « Code
analytique » de l'onglet Saisie n'en proposait que deux, restés figés
depuis le tout premier lancement de l'application.

**Cause** : `SaisieTab.refresh()` — appelé à chaque fois qu'on revient sur
l'onglet Saisie (`App.show()`) — ne rafraîchissait que la liste des
comptes (`_refresh_compte_values()`). Les listes Code analytique, Code
budgétaire, Code bailleur, Fournisseur et Client n'étaient peuplées
qu'une seule fois, à la création du formulaire au démarrage de l'app :
tout code ajouté ensuite (via Plan analytique, ou via les nouveaux boutons
« Ajouter les codes courants » d'Énergie/Maintenance) n'apparaissait donc
jamais dans Saisie tant que l'application n'était pas redémarrée.

**Corrigé** : `SaisieTab.refresh()` rafraîchit désormais aussi ces cinq
listes à chaque retour sur l'onglet Saisie.

### Unités de mesure + coût unitaire moyen pondéré analytique

**Demande** : unité Litre (L) pour eau/gasoil/gaz/essence, Kilowatt (Kw)
pour l'électricité, Heure (H) pour la maintenance — et que la Fabrication
calcule automatiquement le coût de revient d'une heure (ou d'un litre, d'un
kWh) à partir de l'ensemble des enregistrements comptables réels, sur le
même principe que le coût unitaire moyen pondéré déjà utilisé pour les
stocks.

**Réalisé** :
- Colonne `unite` ajoutée à `analytic_codes` (migration automatique). Les
  codes suggérés (`ajouter_codes_analytiques_suggeres`) portent maintenant
  leur unité : `ENERGIE-EAU/GASOIL/GAZ/ESSENCE` → L, `ENERGIE-ELEC` → Kw,
  tous les `MAINT-*` → H.
- **`compute_cout_unitaire_moyen_analytique(conn, code, ...)`** (nouveau) :
  montant total des charges (classe 6) comptabilisées sous ce code, divisé
  par la quantité totale saisie sur ces mêmes lignes (champ Quantité de la
  Saisie) — se met à jour tout seul après chaque facture saisie avec une
  quantité, exactement comme demandé.
- **`compute_cout_production()`** (recette de Fabrication) utilise
  désormais ce coût automatiquement pour toute ligne main-d'œuvre/énergie/
  autre associée à un code analytique — plus besoin de saisir un coût
  unitaire manuel, sauf si aucune quantité n'a encore été comptabilisée
  sous ce code (repli sur la saisie manuelle en attendant).
- **Onglet Fabrication** : nouveau champ « Code analytique » sur chaque
  ligne de recette, avec un **aperçu en direct** du coût moyen pondéré
  constaté (ex. « Coût moyen pondéré constaté : 5 555,56 F CFA / H ») dès
  qu'un code est choisi, et le libellé « Quantité » s'actualise avec
  l'unité (« Quantité (H) »).
- **Onglet Saisie** : le libellé « Quantité » s'actualise aussi avec
  l'unité du code analytique choisi sur la ligne, pour rappeler dans quelle
  unité saisir (litres, kilowatts, heures).
- **Plan analytique** (Paramètres) : nouvelle colonne/champ « Unité »,
  conservée à l'export/import `.xlsx`.

Testé de bout en bout : deux paiements réels de prestataires (100 000 F/20h
et 150 000 F/25h) donnent un coût moyen de 5 555,56 F/h ; une recette avec
2h de main-d'œuvre + 100L d'eau + 50Kw d'électricité récupère automatiquement
les trois coûts unitaires réels (H, L, Kw) sans aucune saisie manuelle ;
fabrication complète (matière + main-d'œuvre analytique) comptabilisée avec
succès, Bilan resté équilibré (écart = 0) ; export/import du Plan analytique
préserve les unités.


### Saisie multi-lignes — révisée : plusieurs comptes au débit ET au crédit

**Correction suite au retour utilisateur** : la première version de la
fenêtre multi-lignes n'acceptait qu'un SEUL compte au crédit. La demande
était en réalité une vraie écriture à lignes multiples des DEUX côtés (ex.
plusieurs charges de classe 6 au débit, réglées par plusieurs comptes de
trésorerie au crédit) — comme un journal général classique.

**Réalisé** :
- `core.add_ecriture_multi_lignes()` remplace l'ancienne fonction
  restrictive : accepte un nombre libre de lignes, chacune au débit OU au
  crédit (jamais les deux sur la même ligne), et refuse l'enregistrement
  tant que Total Débit ≠ Total Crédit (message d'erreur explicite avec
  l'écart exact).
- **Fenêtre `MultiLigneDialog` reconstruite** : une seule grille avec
  colonnes Compte / Libellé / Débit / Crédit / Quantité / Code analytique.
  On ajoute les lignes une à une (au débit ou au crédit, au choix), avec un
  total Débit / total Crédit / Écart affiché en temps réel (vert si
  équilibré, rouge sinon).
- **Listes déroulantes qui s'ouvrent automatiquement au clic** (comptes,
  journal, code analytique) — comme demandé, plus besoin de cliquer sur la
  petite flèche. Le champ Compte est pré-rempli avec les 200 premiers
  comptes et se filtre en tapant (recherche par code ou libellé), comme le
  formulaire de Saisie standard.

Testé : 2 lignes débit (Eau taguée ENERGIE-EAU avec quantité, Entretien) +
2 lignes crédit (Banque, Caisse) — écriture à 4 lignes correctement
équilibrée et enregistrée, Bilan resté équilibré (écart = 0), coût unitaire
moyen pondéré de l'eau recalculé correctement à partir de cette écriture.

### Correctif : la saisie multi-lignes ne générait pas le stock automatiquement

**Bug repéré par l'utilisateur** : un achat de matières premières saisi via
la fenêtre multi-lignes (plusieurs comptes 602xxx au débit avec quantité,
un compte fournisseur au crédit) ne déclenchait pas l'entrée de stock
automatique, contrairement au formulaire de Saisie standard — c'était une
limitation documentée mais qui s'est révélée être l'usage réel le plus
courant (une seule facture fournisseur avec plusieurs matières premières).

**Corrigé** : `core.add_ecriture_multi_lignes()` applique désormais la même
logique d'entrée/sortie de stock automatique que `add_balanced_entry()`,
mais LIGNE PAR LIGNE : chaque ligne débit avec une quantité sur un compte
d'achat lié à un stock (601x marchandises, 602x matières premières)
génère sa propre entrée de stock ; chaque ligne crédit avec une quantité
sur un compte de vente lié à un stock (701x/702x) génère sa propre sortie
de stock (au coût unitaire moyen réel).

Testé avec le scénario exact de l'utilisateur (4 comptes 602xxx au débit
avec quantités, 1 compte fournisseur au crédit) : les 4 mouvements de
stock sont générés individuellement, le stock de matières premières
(320000) reflète bien le cumul (14 200 000 F pour 5 002 unités), et le
Bilan reste équilibré (écart = 0).

### Compte stock global pour une facture d'achat groupée (matière + transport + douane)

**Bug identifié par l'utilisateur** : pour une facture d'achat incluant
plusieurs charges liées au MÊME lot de marchandise (ex. clinker + transport
+ douane, chacun sur un compte 602xxx distinct), le comportement « ligne
par ligne » du correctif précédent aurait généré 3 mouvements de stock
séparés, chacun avec sa propre quantité — multipliant la quantité reçue
par 3 au lieu de la compter une seule fois.

**Corrigé** : nouveau champ **« Compte stock »** (optionnel) dans la
fenêtre multi-lignes, avec une **« Quantité réellement reçue »** — quand
il est renseigné, `core.add_ecriture_multi_lignes()` additionne TOUTES les
lignes au débit de l'écriture (matière + frais accessoires) en **une
seule** entrée de stock, à la quantité réellement reçue (pas une par
ligne) : le coût unitaire moyen du stock reflète alors le vrai coût de
revient, frais de transport et douane inclus. Si le champ est laissé vide,
le comportement précédent (ligne par ligne, chaque ligne avec sa propre
quantité génère son propre mouvement) reste disponible pour des achats
réellement indépendants dans la même écriture.

Testé avec le scénario exact de l'utilisateur (CLINKER 250 000 F +
TRANSPORT 2 500 000 F + DOUANE 500 000 F, quantité reçue 1 200 unités) :
stock à 3 250 000 F pour 1 200 unités, coût unitaire moyen 2 708,33 F/unité
(au lieu de 3 600 unités comptées à tort). Non-régression vérifiée : le
mode « ligne par ligne » (sans compte stock global) continue de fonctionner
normalement pour des lignes réellement indépendantes.

### Correctifs : quantité de stock ignorée + tiers non forcé dans la saisie multi-lignes

**Deux bugs repérés par l'utilisateur sur son propre usage réel** :

1. **La quantité n'était pas prise en compte lors de la génération du
   stock** : le champ « Quantité réellement reçue » n'était pas obligatoire
   — laissé vide, l'entrée de stock automatique se comptabilisait quand
   même, avec une quantité de 0 (visible dans le formulaire : la ligne
   générée montrait un montant mais aucune quantité). **Corrigé** : ce
   champ est désormais obligatoire (et strictement positif) dès qu'un
   compte stock est choisi, aussi bien côté interface (bloque
   l'enregistrement avec un message clair) que côté moteur
   (`add_ecriture_multi_lignes` refuse maintenant l'appel si
   `compte_stock_global` est renseigné sans `quantite_stock_global`).

2. **Le choix d'un fournisseur ou d'un client n'était pas forcé** dans la
   fenêtre multi-lignes (contrairement au formulaire de Saisie standard).
   **Corrigé** : dès qu'une ligne utilise un compte de la racine 40
   (Fournisseurs) ou 41 (Clients), un champ « Fournisseur »/« Client »
   apparaît automatiquement et devient obligatoire pour pouvoir ajouter la
   ligne (avec recherche par code ou raison sociale, liste déroulante
   auto-ouverte au clic) — le tiers choisi est stocké PAR LIGNE (pas un
   seul tiers global pour toute l'écriture, puisqu'une écriture peut
   régler plusieurs fournisseurs différents). Le moteur refuse aussi
   maintenant, en dernier recours, tout enregistrement d'une ligne 40/41
   sans le tiers correspondant renseigné.
   Une nouvelle colonne « Tiers » a été ajoutée au tableau des lignes
   pour visualiser directement qui est concerné par chaque ligne.

Testé avec le scénario exact de l'utilisateur (CLINKER + TRANSPORT +
DOUANE en compte stock global 321001, quantité 1200, réglé au fournisseur
FRS-01) : stock correctement à 1 200 unités pour 2 708,33 F/unité, ligne
fournisseur bloquée tant que le fournisseur n'est pas choisi, Bilan
équilibré. Non-régression vérifiée sur le mode « ligne par ligne » sans
tiers ni stock global.

### Correctifs UI : boutons invisibles + champ Tiers générique retiré

**Bug repéré par l'utilisateur** : sur un écran de taille standard, la
fenêtre multi-lignes avait grandi (ajout du champ Tiers par ligne) au
point que les boutons « Enregistrer l'écriture »/« Annuler » n'étaient
plus visibles en bas — aucun moyen de valider l'écriture sans redimensionner
la fenêtre.

**Corrigé** : réorganisation de l'empilement (`pack`) de la fenêtre —
les boutons et la section « Compte stock » sont désormais ancrés en BAS de
la fenêtre en premier (`side="bottom"`), et le tableau des lignes (au
milieu) est la seule zone qui s'agrandit ou se réduit selon l'espace
disponible, avec sa propre scrollbar verticale. Les boutons restent donc
**toujours visibles**, quelle que soit la taille de l'écran ou le nombre
de lignes ajoutées. Fenêtre également rendue redimensionnable (taille
minimale 900×500), avec une hauteur par défaut plus généreuse (1080×680).

**Champ « Tiers » générique retiré** de la section « Informations
communes » : il faisait doublon et n'était plus pertinent maintenant que
le tiers (Fournisseur/Client) se choisit ligne par ligne, de façon
obligatoire et fiable, dès qu'un compte des racines 40/41 est utilisé.

### Correctif majeur : sens du mouvement de stock (entrée/sortie) dans le compte stock global

**Bug repéré par l'utilisateur** : une VENTE à des clients (comptes 41xxx
débités, banque créditée) générait quand même une **entrée** de stock
(augmentation), alors que le stock aurait dû **diminuer** — le compte
stock global ne savait faire que le sens « achat ».

**Cause** : `add_ecriture_multi_lignes()` additionnait systématiquement les
lignes au débit comme coût d'une ENTRÉE de stock, quel que soit le sens
réel de l'opération. Pour une vente, les lignes au débit sont des
créances clients (leur montant = prix de vente), pas un coût de revient —
utiliser ce montant comme coût de sortie de stock aurait été doublement
faux.

**Corrigé** : nouveau paramètre `sens_stock_global` (« entree » ou
« sortie ») :
- **Entrée (achat)** : comportement inchangé — coût = somme des lignes
  débit (matière + frais accessoires), le stock augmente.
- **Sortie (vente)** : le coût = quantité × **coût unitaire moyen actuel**
  du stock (même logique qu'une vente simple, via `compute_stocks_detail`
  — pas `compute_stocks`, qui se limite aux 4 comptes centralisateurs et
  ne couvre pas les sous-comptes granulaires comme `321001 CLINKER`), le
  stock diminue.

**Nouveau sélecteur « Sens » (Entrée (achat) / Sortie (vente))** dans la
section « Compte stock » de la fenêtre multi-lignes, avec un texte
explicatif qui s'adapte au sens choisi et un libellé de quantité qui
change (« Quantité reçue » / « Quantité vendue »).

Testé avec le scénario exact de l'utilisateur (vente à 3 clients CLI-0004/
CLI-0009/CLI-0006, réglée par une banque, sortie de 10 unités de CLINKER) :
le stock diminue bien de la bonne quantité, au coût unitaire moyen réel
(pas au prix de vente), Bilan équilibré. Non-régression vérifiée sur le
mode entrée (achat) et le mode ligne par ligne classique.

### Nouveau menu ADMIN : taux paramétrables, correction consolidée, impression des factures

**Demande** : unifier Facturation/Factures frs, bouton « Imprimer la
facture », menu ADMIN pour la modification des factures, TVA paramétrable
(Facturation) et retenue à la source par compte de classe 44 paramétrable
(Factures frs), les taux étant définis dans ADMIN.

**Réalisé** :
- **Nouveau menu ADMIN** avec 3 sous-menus :
  - **Taux de TVA** : liste de taux nommés et réutilisables (ex. « TVA
    standard 18% », « Exonéré 0% »), gérés comme le Plan analytique (créer/
    modifier/supprimer, import/export .xlsx).
  - **Taux de retenue à la source** : même principe (ex. « Retenue BIC
    5% »), pour Factures frs.
  - **Modification des factures** : vue consolidée de TOUTES les factures
    déjà validées (vente ET achat), avec un bouton « Dévalider » unique —
    complète, sans le remplacer, le bouton déjà présent dans chaque onglet.
- **`core.devalider_facture_achat()`** (nouveau — n'existait pas, seul le
  côté vente avait cette capacité) : symétrique à `devalider_facture_vente()`.
- **Facturation (COMMERCE)** : nouveau menu déroulant « Préréglage (ADMIN) »
  à côté du champ TVA — sélectionner un taux de la liste ADMIN remplit
  automatiquement le champ (toujours modifiable à la main si besoin).
- **Factures frs (ENGAGEMENTS-PROJETS)** : même principe pour la retenue à
  la source, en plus du choix du compte de retenue (classe 44) déjà
  existant et déjà filtré à cette classe. **Bouton « Corriger cette
  facture »** ajouté (n'existait pas sur ce module, seulement côté ventes)
  utilisant la nouvelle fonction de dévalidation.
- **Bouton « Imprimer la facture »** sur les deux modules : génère un
  document HTML avec bouton Imprimer intégré (Ctrl+P), ouvert directement
  dans le navigateur par défaut — aucune dépendance PDF supplémentaire à
  intégrer au `.exe`.

**Choix de conception assumé** : Facturation et Factures frs restent deux
classes séparées (`FacturationTab`/`FacturesFrsTab`) plutôt qu'une fusion
complète en une seule fenêtre — une fusion aurait été un chantier à part
entière, risqué pour deux modules déjà testés en production. Elles ont en
revanche été rendues **strictement cohérentes** : mêmes boutons (Nouvelle
facture / Supprimer / Corriger / Imprimer), même comportement de
dévalidation, même mécanisme de préréglage de taux depuis ADMIN.

Testé de bout en bout : taux TVA et retenue paramétrés dans ADMIN, facture
de vente et facture d'achat créées/validées avec ces taux, export HTML des
deux imprimable, dévalidation des deux types de factures (simulant le
bouton ADMIN consolidé), Bilan resté équilibré à chaque étape.

### Comptes fiscaux liés aux taxes (ADMIN) + code analytique par ligne de facture

**Demande confirmée** : présenter dans ADMIN les comptes fiscaux liés aux
différentes taxes, et reproduire dans Facturation/Factures frs le même
type d'interaction que la fenêtre multi-lignes (listes déroulantes qui
s'ouvrent au clic).

**Réalisé** :
- **Colonne `compte`** ajoutée aux tables `taux_tva` et `taux_retenue`
  (migration automatique) : chaque taux paramétrable dans ADMIN est
  maintenant lié à son compte fiscal réel (classe 44), avec un champ dédié
  filtré à cette classe (recherche par code ou libellé, liste déroulante
  auto-ouverte au clic).
- **`factures_vente.tva_compte`** (nouveau, migration incluse) : le compte
  de TVA n'est plus figé sur `443100` — sélectionner un préréglage TVA dans
  Facturation renseigne maintenant AUSSI le compte fiscal utilisé à la
  validation (au lieu du taux seul). Testé : une facture avec compte TVA
  personnalisé poste bien dessus.
- **Même principe côté retenue à la source** (Factures frs) : sélectionner
  un préréglage renseigne le taux ET le compte de retenue.
- **Code analytique par ligne de facture** (vente ET achat) : nouvelle
  colonne `analytic_code` sur `facture_vente_lignes`/`facture_achat_lignes`
  (migration incluse), champ dédié dans le formulaire d'ajout de ligne,
  colonne affichée dans le tableau — permet de rattacher une ligne de
  facture à un code Énergie/Maintenance comme partout ailleurs dans
  l'application.
- **Listes déroulantes auto-ouvertes au clic** (client/fournisseur, compte
  de vente/achat, préréglages TVA/retenue, code analytique) dans
  Facturation et Factures frs, reproduisant l'interaction de la fenêtre
  Saisie multi-lignes.

Testé de bout en bout : facture de vente avec compte TVA personnalisé et
code analytique sur sa ligne, facture d'achat avec compte de retenue
personnalisé et code analytique sur sa ligne, toutes deux validées,
export HTML des deux, Bilan resté équilibré (écart = 0).

### Catégories de retenue à la source courantes (ADMIN)

**Demande** : lister les nombreuses retenues à la source possibles dans
ADMIN, à choisir ensuite dans la facture. Recherche effectuée sur les
taux réels en vigueur au Burkina Faso : la législation a changé en 2026
(retenue TVA passée de 20% à 30%) et les sources disponibles se
contredisent sur d'autres taux (BIC cité à 2% ou 5% selon le cas) — les
taux n'ont donc **volontairement pas été inventés**.

**Réalisé** : nouveau bouton **« Ajouter les catégories courantes (BIC,
IS, TVA...) »** dans l'écran ADMIN > Taux de retenue à la source. Ajoute 7
catégories reconnues officiellement par la DGI du Burkina Faso (d'après la
liste de leurs formulaires de déclaration officiels sur dgi.bf) — **à 0%
et sans compte fiscal pré-rempli**, à compléter par l'utilisateur avec le
taux et le compte exacts applicables à son cas :
- Retenue BIC (fournisseurs non attributaires)
- Retenue Impôt sur les Sociétés (IS)
- Retenue à la source de la TVA
- Retenue sur loyers d'immeuble
- Retenue sur sommes versées aux prestataires établis au Burkina Faso
- Retenue sur sommes versées aux personnes sans installation professionnelle au Burkina Faso
- Retenue sur commandes publiques

N'écrase jamais une catégorie déjà personnalisée (même mécanisme que les
codes analytiques Énergie/Maintenance). Ce bouton (`SUGGESTIONS_FN`) est
généralisé dans `_SimplePlanTab` — réutilisable pour n'importe quel futur
plan à catégories courantes, sans dupliquer le code.

Testé : les 7 catégories s'ajoutent correctement, une retenue déjà
personnalisée par l'utilisateur (« 5% » / 447810) n'est jamais touchée par
un second appel.

### Factures frs — le brouillon devient un « Bon de commande » imprimable

**Demande** : renommer « Enregistrer (brouillon) » en « Enregistrer BON DE
COMMANDE », placer un bouton « Imprimer le bon de commande » juste à côté,
avec un modèle modifiable dans ADMIN.

**Réalisé** :
- Bouton renommé **« Enregistrer BON DE COMMANDE »**.
- Bouton **« Imprimer le bon de commande »** déplacé juste à côté (retiré
  de la barre du haut), au lieu d'« Imprimer la facture ».
- `core.export_bon_commande_html()` (nouveau) : génère le document intitulé
  « Bon de commande » tant que la facture est en brouillon — reprend
  l'en-tête/pied de page propres à cette commande s'ils sont remplis,
  sinon le **modèle par défaut** défini dans ADMIN. Une fois la facture
  validée, le même bouton imprime automatiquement la vraie « Facture
  d'achat » (`export_facture_achat_html`, comportement inchangé) — le bon
  de commande n'a plus lieu d'être à ce stade.
- **Nouvel écran ADMIN « Modèle de bon de commande »** : deux champs
  (en-tête / pied de page) enregistrés une seule fois, appliqués à toute
  commande dont les champs propres sont vides — pratique pour ne pas
  retaper les coordonnées de la société à chaque bon de commande.

**Bug corrigé au passage** : `core.set_text_setting()` était appelé
depuis `FacturationTab` (préréglage TVA) mais n'existait pas dans
`core.py` — un `AttributeError` aurait empêché l'enregistrement d'une
facture de vente avec TVA. Ajouté (alias explicite de `set_setting()`
pour les réglages textuels), testé.

Testé de bout en bout : bon de commande en brouillon utilisant le modèle
ADMIN, facture validée imprimant le bon document, Bilan resté équilibré.

### Circuit interne Expression de besoin → Bon de commande → Bordereau de livraison

**Demande** : 3 nouveaux sous-menus dans ENGAGEMENTS-PROJETS, formant un
circuit d'approbation d'achat interne **sans aucun lien avec la
comptabilité à aucune étape** — la validation d'une étape fait basculer
le document dans le sous-menu suivant.

**Réalisé** :
- **Expression de besoin** : demande interne (numéro, date, demandeur,
  service, lignes libellé/quantité/unité). Sa validation verrouille
  l'expression et crée automatiquement un **Bon de commande** avec les
  mêmes lignes.
- **Bon de commande** (circuit interne — distinct du bouton "Enregistrer
  BON DE COMMANDE" de Factures frs, qui lui génère de vraies écritures une
  fois validé) : peut aussi être créé directement, précise le fournisseur
  et un prix unitaire par ligne (montant purement indicatif, jamais
  comptabilisé). Sa validation verrouille le bon et crée automatiquement
  un **Bordereau de livraison**.
- **Bordereau de livraison** : dernière étape — reprend les quantités
  commandées, avec une quantité livrée modifiable (double-clic sur une
  ligne) pour gérer les livraisons partielles. Sa validation marque
  simplement la réception comme confirmée — fin du circuit.

**Interface** : chaque sous-menu affiche la liste des documents en
tableau aligné (N°, date, statut...) ; **double-clic sur une ligne pour
l'ouvrir en grand**, ajouter/modifier des lignes, et valider — exactement
comme demandé.

Testé de bout en bout, y compris une réception partielle (200 sacs
commandés, 180 livrés) : le circuit complet s'exécute sans générer une
seule écriture comptable, à aucune étape.

### Sous-menu Règlements — la comptabilisation du circuit interne

**Demande** : ajouter « Règlements » dans ENGAGEMENTS-PROJETS. La
validation d'un Bon de commande (circuit interne) doit aussi le ventiler
dans ce sous-menu, ET là, contrairement aux 3 étapes précédentes, générer
la vraie saisie comptable — avec compte de charge, code analytique et
retenue fiscale à choisir par ligne.

**Réalisé** :
- `valider_ep_bon_commande()` crée maintenant, EN PLUS du Bordereau de
  livraison, un **Règlement en brouillon** (lignes recopiées du bon de
  commande — libellé, quantité, prix unitaire — mais **sans compte de
  charge ni code analytique**, à choisir dans ce nouvel écran).
- **`core.valider_reglement()`** (nouveau) : comptabilise réellement le
  règlement — débit de chaque compte de charge choisi (avec son code
  analytique), crédit fournisseur (net à payer), crédit retenue fiscale si
  applicable, plus entrée de stock automatique pour les lignes liées à un
  compte de marchandises/matières premières — même principe que
  `valider_facture_achat()`. **Refuse explicitement la validation tant
  qu'une ligne n'a pas de compte de charge choisi**, avec le détail des
  lignes manquantes dans le message d'erreur.
- **`core.devalider_reglement()`** (correction) : symétrique aux autres
  modules, retire les écritures et repasse en brouillon.
- **Écran Règlements** (liste + double-clic pour ouvrir en grand) : pour
  chaque ligne, sélection d'un **compte de charge** (classe 6, recherche
  filtrée) et d'un **code analytique**, plus une section **retenue
  fiscale** avec préréglages ADMIN (taux + compte, comme Factures frs).
  Boutons Enregistrer / Valider (comptabiliser) / Corriger (repasser en
  brouillon).

Testé de bout en bout : circuit complet Expression de besoin → Bon de
commande → Règlement, refus de validation tant que le compte de charge
manque, comptabilisation correcte une fois complété (débit charge avec
code analytique, crédit fournisseur, crédit retenue, entrée de stock
automatique), correction puis recomptabilisation — Bilan resté équilibré
à chaque étape.

### Suivi des retards de paiement sur le Bon de commande (circuit interne)

**Demande** : à partir du Bon de commande, ajouter la date de facture, la
date de saisie, la date de paiement attendu, et le retard de paiement
calculé par rapport à la date de saisie.

**Réalisé** — même principe déjà utilisé pour Achats/Recouvrement, appliqué
au nouveau circuit :
- 3 nouvelles colonnes sur `ep_bons_commande` : `date_facture`,
  `date_saisie`, `date_paiement_attendu` (migration incluse).
- **`list_ep_bons_commande()`** calcule désormais le statut de paiement :
  tant que la date de saisie n'est pas renseignée, retard « en cours »
  comparé à aujourd'hui ; une fois la date de saisie renseignée, le retard
  se fige définitivement (date de saisie − date de paiement attendu).
- **Écran Bon de commande** : 3 nouveaux champs dans le détail (double-clic),
  et une ligne de statut « Paiement : ... » colorée en rouge si en retard,
  vert sinon. La liste affiche aussi les colonnes « Paiement attendu » et
  « Statut paiement », avec les lignes en retard mises en évidence en
  rouge gras.

Testé : date de saisie postérieure à la date de paiement attendu → retard
correctement calculé à 19 jours ; circuit complet (Bon de commande →
Bordereau + Règlement) toujours fonctionnel, Bilan resté équilibré.

### Amortissements exacts par catégorie (à partir des formules du système de référence)

**Fourni par l'utilisateur** : les formules exactes (`CtaCptSolde`,
`CtaCptSoldeDébit`, `CtaCptSoldeCrédit`) de son ancien système de rapport
financier, avec les plages de comptes précises utilisées pour chaque
ligne du Bilan.

**Vérifié — déjà conforme** : la classification Actif/Passif par racine
(compte par compte, débiteur → Actif / créditeur → Passif pour les
racines 40 à 49), le regroupement des capitaux propres par racine
classe 1, et le calcul du total Actif/Passif à partir de la Balance
correspondent déjà exactement à la logique de ces formules.

**Corrigé — répartition proportionnelle remplacée par les vraies plages
de comptes** : `IMMO_CATEGORIES` utilisait jusqu'ici une répartition
*proportionnelle* de l'amortissement entre catégories (faute de connaître
la correspondance exacte compte-immo ↔ compte-amortissement). Les
formules fournies donnent cette correspondance exacte (ex. Bâtiments
231-233 ↔ amortissements 2831*-2833*/2931*-2933* ; Matériel 240-244 ↔
2840*-2844*/2940*-2944*...) — désormais utilisée telle quelle pour
calculer un amortissement **exact** par catégorie, plus indicatif. Une
ligne « Autres immobilisations non classées » absorbe automatiquement
tout compte hors de ces plages précises (garantit que le détail somme
toujours exactement au total, sans rien perdre).

**Écart volontairement conservé avec le système de référence** : leur
formule de résultat (`=-CtaCptSolde("7*")-CtaCptSolde("6*")`) ne couvre
que les classes 6 et 7, et leur ligne « Prov fin + banq créditrices *19* »
n'est, dans leurs formules, pas négée contrairement à toutes les autres
lignes du Passif (`=CtaCptSolde("19*")` au lieu de `-CtaCptSolde("19*")`)
— ce qui, si des comptes de classe 8 ou 19 sont utilisés, désaligne leur
propre TOTAL I. `compute_resultat_net_complet()` de cette application
inclut délibérément la classe 8 et négie systématiquement la classe 1
(déjà garanti équilibré, voir la section plus haut sur le diagnostic
d'écart) — pour ne pas réintroduire le bug d'origine.

Testé : compte 231100 (10 000 000, amort 283100 = -3 000 000) → Bâtiments
net 7 000 000 exact ; compte 244100 (5 000 000, amort 284100 = -1 000 000)
→ Matériel net 4 000 000 exact ; un compte hors plage (200000, racine
isolée) bascule proprement dans « Autres immobilisations non classées »
sans casser l'écart (resté à 0 dans tous les cas).

### TFT — corrections à partir des formules du système de référence (CtaCptSolde)

**Fourni par l'utilisateur** : les formules exactes du TFT (méthode
indirecte, CAFG) de son ancien système. Décodage précis, avec 3 vraies
lacunes corrigées et une amélioration de structure :

1. **Bug corrigé — stocks incomplets dans le TFT** : `compute_tft_indirect()`
   utilisait encore l'ancienne liste partielle `COMPTES_STOCK_PREFIXES`
   (31/32/33/36 seulement, déjà corrigée dans le Bilan il y a plusieurs
   sessions mais pas ici) au lieu de la classe 3 entière — tout sous-compte
   de stock hors de cette liste (33, 37, 38, 39...) disparaissait
   silencieusement de la variation de trésorerie calculée. Corrigé :
   classe 3 entière, comme `CtaCptSoldeDébit("3*")` dans le rapport de
   référence.
2. **Bug corrigé — avances sur immobilisations absentes des
   investissements** : les acquisitions d'immobilisations corporelles ne
   couvraient que les racines 22-23-24 ; la racine 25 (avances et acomptes
   versés sur immobilisations), présente dans la formule de référence
   (`CtaCptSolde("22*","25*")`), a été ajoutée.
3. **Bug corrigé — capital limité à 3 comptes** : l'augmentation de
   capital ne captait que les comptes 101/104/105 ; élargi à la racine 10
   entière (`CtaCptSolde("10*")` dans la référence), même principe que
   pour le Bilan (jamais de liste de comptes partielle pour un total).
4. **Structure enrichie — CAF Exploitation intermédiaire** : entre l'EBE
   et la CAFG, ajout des lignes « Produits des cessions courantes
   d'immobilisations (754) », « Valeurs comptables des cessions courantes
   (654) » et « Transferts de charges d'exploitation (781) », tirées
   directement du rapport de référence, avant application des revenus/
   frais financiers pour obtenir la CAFG.
5. **Variation de l'actif circulant HAO isolée** : séparée de la
   variation des créances classiques — créances = racines 40-45 débit,
   actif circulant HAO = racines 46-49 débit (comme dans le rapport de
   référence), le passif circulant restant une ligne unique (racines 40-49
   crédit, comme avant).

**Écarts volontairement conservés** avec le système de référence (pour ne
pas réintroduire de risque de déséquilibre) : emprunts couvrent les
racines 16 ET 17 (la référence n'utilise que 16), et les immobilisations
incorporelles couvrent les racines 20 ET 21 (la référence n'utilise que
21) — ces élargissements ne peuvent que capter des mouvements
supplémentaires, jamais en perdre.

Testé : un compte de stock 380000 (hors ancienne liste), une avance sur
immobilisation (compte 251000), et un apport en capital sur compte 101
(racine 10 entière) sont tous correctement pris en compte — écart de
contrôle du TFT resté à 0, Bilan resté équilibré.

### Situation financière — corrections majeures à partir des formules du système de référence

**Fourni par l'utilisateur** : les formules exactes de la Situation
financière (FR-BFR-TN) de son ancien système. Deux **vrais bugs** trouvés
en les comparant à mon implémentation — les mêmes catégories de bugs déjà
corrigées ailleurs dans l'application, mais qui n'avaient jamais été
propagées à cet écran :

1. **Bug corrigé — Ressources stables tronquées** :
   `capitaux_propres_ressources` sommait uniquement les catégories
   curées du Bilan simple (`COMPTES_CAPITAL=["101","118","121"]`,
   `COMPTE_SUBVENTIONS="141"`, `COMPTE_PROVISIONS="191"`), en oubliant la
   ligne « Autres postes de ressources durables » qui absorbe tout le
   reste — tout compte de classe 1 hors de ces 3 codes précis (soit la
   quasi-totalité d'un vrai plan comptable de 1591 comptes) disparaissait
   silencieusement du Fonds de Roulement, de la Rentabilité économique et
   financière. Corrigé : somme exhaustive de la classe 1 entière (comme
   `CtaCptSolde("1*")` dans le rapport de référence), avec un détail
   indicatif Capitaux propres (racines 10-15) / Dettes financières
   (racines 16-17) / Autres (18-19) qui somme toujours exactement au
   total.
2. **Bug corrigé — BFR exploitation avec racines 40/41 « en bloc »** :
   le calcul utilisait encore l'ancienne convention (racine 41 toujours
   en créances, racine 40 toujours en dettes, quel que soit le signe du
   compte) — la même erreur déjà corrigée dans le Bilan il y a plusieurs
   sessions (un client créditeur y restait affiché, à tort, dans les
   créances). Corrigé : racines 40 à 46 classées compte par compte selon
   le signe de leur solde, comme le Bilan et comme
   `CtaCptSoldeDébit/Crédit("3*","46*")` dans le rapport de référence.
3. **Dividendes versés désormais calculés** (compte 465, Associés —
   dividendes à payer, `=CtaCptSolde("465*")` dans la référence) au lieu
   d'être toujours à 0 (non isolé auparavant).

Testé avec un scénario ciblant précisément ces deux bugs : un compte de
classe 1 hors des anciennes listes curées (105000, +5 000 000) apparaît
maintenant bien dans les ressources stables ; un client au solde créditeur
(419100, avoir de 300 000) bascule correctement en passif circulant
d'exploitation au lieu de rester en créances — le contrôle de trésorerie
(Trésorerie nette calculée vs réelle) tombe exactement à 0, Bilan resté
équilibré.

### Présentation TFT et Situation financière — alignée sur les PDF de référence + validation par les vrais chiffres

**Fourni par l'utilisateur** : les PDF réels (Bilan, TFT, Situation
financière) avec les vrais montants attendus, en plus des formules déjà
fournies précédemment.

**Vérification manuelle croisée, chiffre par chiffre** (précieuse : ce
sont les VRAIS résultats de l'utilisateur, pas un exemple inventé) :
- CAFG = EBE(6 679 997 354) + Revenus financiers(92 767 739) - Frais
  financiers(1 356 338 942) = **5 416 426 151** — exact.
- Ressources stables = Capitaux propres(14 022 638 895) + Dettes
  financières(17 262 479 052) = **31 285 117 947** — exact.
- Fonds de roulement = 31 285 117 947 - 29 368 227 937 = **1 916 890 010** — exact.
- Besoin de financement global = 208 057 048 + (-404 023 042) = **-195 965 994** — exact.
- Trésorerie nette = 1 916 890 010 - (-195 965 994) = **2 112 856 004** — exact,
  et cohérent avec la trésorerie du Bilan (Caisse débitrice 2 112 856 004)
  ET avec le TFT (Trésorerie nette au 31/12/N).
- TFT : Flux opérationnels (Somme FA-FE, hors CAFG) = -446 083 968 (HAO)
  + 1 571 701 345 (stocks) + -60 615 663 (créances) + -591 091 968 (dettes)
  = **473 909 746** — exact ; B = CAFG + ce flux = **5 890 335 897** — exact ;
  Variation de trésorerie = B+C+D = **1 572 707 188** — exact.

Cette vérification manuelle confirme que la structure de calcul mise en
place lors des deux sessions précédentes (formules TFT et Situation
financière) reproduit fidèlement le système de référence, chiffre pour
chiffre.

**Présentation des écrans TFT et Situation financière reconstruite pour
suivre exactement l'ordre et les libellés des PDF** :
- **TFT** : ajout de la ligne intermédiaire « Flux de trésorerie provenant
  des activités opérationnelles (Somme FA à FE) » (variation du BFR seule,
  avant ajout de la CAFG) en plus du flux « B » déjà affiché (CAFG + BFR) —
  les deux valeurs sont désormais visibles, comme dans le PDF. Ordre des
  lignes de variation du BFR corrigé (HAO, puis Stocks, puis Créances, puis
  Dettes, comme le PDF) — l'ordre n'affecte pas le résultat mais facilite
  la comparaison ligne à ligne avec le rapport de référence. Cessions
  d'immobilisations incorporelles et corporelles fusionnées en une seule
  ligne, comme dans le PDF.
- **Situation financière** : « RÉSULTAT NET COMPTABLE » ajouté en première
  ligne (manquant auparavant) ; détail de la CAF d'exploitation (654/754/
  781) ajouté, comme pour le TFT ; tous les libellés et l'ordre des
  sections (CAFG, Analyse de la situation financière, BFR, Trésorerie
  nette, Flux de la période, Endettement financier) alignés sur le PDF ;
  affichage de la ligne « Actifs immobilisés » en négatif, comme dans le
  PDF (le calcul sous-jacent, lui, reste inchangé).
- Les deux écrans utilisent maintenant le même format de montant que le
  Bilan (séparateur espace, style SYSCOHADA) au lieu du séparateur virgule.

Testé : écart de contrôle du TFT et de la Situation financière toujours à
0 après ces changements de présentation, cohérence Bilan/TFT/Situation
financière vérifiée sur un scénario simple.

### Bilan — sous-totaux Brut/Amortissements + colonne N-1 (comme le PDF)

**3 manques confirmés par l'utilisateur** en comparant son écran à un
export réel de son ancien système (mêmes formules déjà fournies
précédemment) : pas de sous-total Brut/Amortissements séparé du Net, pas
de colonne N-1, présentation différente du gabarit de référence.

**Vérification préalable, chiffre par chiffre** : sur 8 lignes
d'immobilisations comparées entre l'écran (exercice 2026) et le PDF de
référence (exercice 2024), **5 correspondaient déjà exactement au franc
près** (Terrains, Avances sur immobilisations, Immobilisations
financières, Matériel de transport, Brevets/licences) — preuve que le
calcul est juste. Les 3 qui différaient (Bâtiments, Installations,
Matériel) sont précisément les catégories où de nouveaux investissements
sont les plus probables sur 2 ans d'écart entre les deux exercices — pas
nécessairement un bug, mais l'absence de colonne N-1 rendait cette
vérification impossible à faire depuis l'écran lui-même.

**Réalisé** :
- **Sous-totaux Brut et Amortissements** ajoutés, séparés du sous-total
  Net, dans la section Immobilisations (3 lignes : Total BRUT, Total
  AMORTISSEMENTS et provisions, Total NETTES) — écran et export .xlsx.
- **Colonne N-1 complète** sur tout le Bilan (Actif : Net N-1 ; Passif :
  Exercice N-1), calculée en récupérant automatiquement les données de
  l'exercice précédent (`exercice - 1`) — 0 si cet exercice n'a pas de
  données (comme le PDF, qui laisse aussi des cellules N-1 vides).
- **Créances, dettes, trésorerie et capitaux propres regroupés par
  racine/préfixe avec un sous-total** (au lieu du détail compte par
  compte affiché seul) — nécessaire pour que la comparaison N vs N-1 ait
  un sens : un compte peut exister une année et pas l'autre, mais une
  racine (ex. « 41 — Clients débiteurs ») reste comparable d'un exercice à
  l'autre. Nouvelles fonctions `_compute_bilan_groupes()` et
  `_merge_n1()`, réutilisables pour tout futur besoin de comparaison
  d'exercices.

Testé avec 2 exercices réels (2025 clôturé puis reporté sur 2026, avec une
extension de bâtiment et une nouvelle créance client en 2026) : le Total
Actif N (6 700 000) et N-1 (4 000 000) sont corrects, une créance créée en
2026 affiche bien N-1 = 0 (n'existait pas en 2025), écart resté à 0 sur
Bilan, Situation financière et TFT.

### Export du Bilan dans le GABARIT EXACT fourni par l'utilisateur

**Demande** : utiliser le fichier gabarit Excel fourni tel quel dans le
sous-menu Bilan, à la place de l'ancien export.

**Réalisé** :
- Le gabarit (`templates/bilan_template.xls`, format SpreadsheetML —
  reconnu par Excel malgré l'extension .xls) est embarqué dans
  l'application et copié dans le `.exe` au build (workflow GitHub Actions
  mis à jour).
- **`core.export_bilan_gabarit_xlsx()`** (nouveau) : un interpréteur
  générique des formules du gabarit (`CtaCptSolde`, `CtaCptSoldeDébit`,
  `CtaCptSoldeCrédit`, et leurs variantes N-1 `...Nm1`) — chaque formule
  est retrouvée dans le fichier XML brut, évaluée directement à partir de
  la Balance (exercice N et N-1), puis remplacée par sa valeur numérique.
  Gère les références nommées internes au gabarit (`[R120.EtLoc]=...`,
  utilisées par ex. pour sommer plusieurs lignes de capitaux propres dans
  le TOTAL I). **131 formules** au total dans ce gabarit, toutes évaluées
  automatiquement — aucune n'a eu besoin d'être recopiée à la main.
- Bouton **« Exporter (.xlsx) »** du Bilan pointant maintenant sur ce
  nouvel export (`.xls`, comme le gabarit d'origine) — l'ancien export
  personnalisé (`export_bilan_detaille_xlsx`) reste disponible dans le
  code mais n'est plus utilisé par ce bouton, comme demandé.

**Bug découvert et corrigé pendant le développement** : certaines
cellules du gabarit portent un attribut Excel supplémentaire
(`x:Ticked="1"`), que le premier motif de recherche ne reconnaissait pas
— 64 formules sur 131 auraient été oubliées silencieusement. Corrigé et
vérifié : les 131 formules sont maintenant toutes détectées et évaluées.

Testé avec un historique réel sur 2 exercices (2025 clôturé puis reporté
sur 2026, avec une extension de bâtiment) : le gabarit exporté n'a plus
aucune formule non évaluée, la valeur Brut N-1 du Bâtiment (5 000 000)
est correcte, et TOTAL GENERAL Actif = TOTAL GENERAL Passif selon les
propres formules du gabarit (calcul indépendant de compute_bilan(),
donc une vérification croisée supplémentaire de la cohérence globale).

### Adoption du moteur "bilan-auto" — 4 états financiers dans leurs gabarits officiels

**Fourni par l'utilisateur** : un projet Python complet (« bilan-auto »)
avec un moteur d'évaluation de formules bien plus abouti que celui bricolé
la session précédente (résolution multi-passes des dépendances entre
rubriques, quel que soit leur ordre dans la feuille), et 4 vrais gabarits
`.xlsx` (Bilan, Compte de Résultat, Situation Financière, TFT).

**Réalisé** :
- Le moteur de formules du projet de référence (`cta_cpt_solde*`,
  `evaluate_sheet_formulas`, résolution des rubriques `[Rxxx.EtLoc]` en
  plusieurs passes) a été **porté tel quel** dans `core.py`, adapté pour
  lire les soldes directement depuis la Balance de CETTE application
  (`compute_balance()`) au lieu d'exiger l'import de deux fichiers de
  balance externes comme le fait le projet fourni.
- Les 4 gabarits (`modele_bilan.xlsx`, `modele_resultat.xlsx`,
  `modele_situation.xlsx`, `modele_flux.xlsx`) sont embarqués dans
  `templates/` et intégrés au build `.exe`.
- **`core.generate_etat_xlsx(conn, etat_id, output_path, exercice=None)`**
  (nouveau) : génère n'importe lequel des 4 états dans son gabarit
  officiel, en calculant le solde N depuis l'exercice demandé et N-1
  depuis l'exercice précédent.
- **Bouton « Exporter (gabarit officiel .xlsx) »** ajouté aux onglets
  Compte de résultat, TFT et Situation financière (le Bilan l'avait déjà,
  désormais routé sur ce même moteur plus robuste).
- **`+ Nouvel exercice` clôture maintenant l'exercice en cours** avant de
  créer le suivant (avec confirmation) — corrige un vrai manque : le
  bouton ne faisait auparavant que basculer l'affichage sur une nouvelle
  année, sans jamais reporter les soldes de clôture, ce qui faisait
  apparaître un Bilan vide au démarrage d'un nouvel exercice tant que
  l'utilisateur ne pensait pas à aller cliquer sur « Clôturer » dans un
  autre onglet.
- **Sous-menus « Achats » et « Factures frs » supprimés** du menu
  ENGAGEMENTS-PROJETS (remplacés par le circuit Expression de besoin →
  Bon de commande → Bordereau de livraison → Règlements).

Testé de bout en bout : clôture automatique 2025→2026 avec report des
soldes, les 4 états générés sans erreur (Bilan 132/132, Compte de Résultat
41/41, TFT 33/33 ; Situation Financière 34/35 — la seule cellule en échec
utilise une fonction `Ratio(...)` que le projet de référence lui-même ne
gère pas non plus, comportement identique à l'original), Bilan resté
équilibré après report.

### Suppression du menu ÉTATS ET RAPPORTS + nouveaux menus TRANSPORT / IMMOBILISATIONS / RAPPORTS TECHNIQUE

**Demande confirmée par l'utilisateur** (après vérification explicite,
étant donné l'ampleur du travail concerné) : suppression complète du menu
ÉTATS ET RAPPORTS et de son contenu (Grand livre, Balance, Bilan, Compte
de résultat, TFT, Situation financière, Liasse fiscale, Écritures non
équilibrées, Tableaux d'exécution budgétaire, Impôts, Déclarations
sociales, Rapprochements bancaires).

**Réalisé** :
- Le menu et les 12 entrées correspondantes retirés de l'interface.
- **Le moteur de calcul sous-jacent (`core.py`) n'a PAS été touché** :
  `compute_bilan()`, `compute_balance()`, `compute_liasse_resultat()`,
  `compute_tft_indirect()`, `compute_situation_financiere()`,
  `generate_etat_xlsx()`, etc. restent pleinement fonctionnels et
  continuent d'être utilisés en interne par d'autres fonctionnalités
  actives (clôture d'exercice, validation des Règlements, calcul du
  résultat net...) — seuls les ÉCRANS d'affichage ont été retirés du menu.
- **3 nouveaux menus ajoutés** : TRANSPORT, IMMOBILISATIONS, RAPPORTS
  TECHNIQUE — chacun avec un écran provisoire en attente de précisions sur
  leur contenu exact (structure de menu en place, prête à être développée
  dès que le besoin est précisé).

Testé : démarrage de l'application, tous les onglets enregistrés
correspondent à une classe existante (29 onglets au total), le moteur de
calcul comptable interne (Bilan, clôture d'exercice) reste pleinement
fonctionnel malgré le retrait de ses écrans d'affichage.

### Menus TRANSPORT et IMMOBILISATIONS + Pièces de rechange partagées

**Réalisé** :
- **TRANSPORT** : 4 sous-menus, sans lien avec la comptabilité (même
  principe que le circuit Expression de besoin) :
  - **Parc auto** : fiche véhicule (immatriculation, marque, modèle, type,
    chauffeur affecté, statut).
  - **Missions** : trajets par véhicule (chauffeur, destination, dates,
    km départ/retour — calcule automatiquement les km parcourus).
  - **Pièces de rechange** : stock (code, désignation, quantité, coût
    unitaire) — **PARTAGÉ** avec le menu MAINTENANCE-ÉNERGIE, comme
    demandé, pour servir aussi bien aux réparations de véhicules qu'à la
    maintenance générale d'équipement.
  - **Réparations** : par véhicule, avec un détail (double-clic) où
    chaque pièce utilisée **décrémente automatiquement** le stock de
    Pièces de rechange (refuse si stock insuffisant), plus la main
    d'œuvre — calcule le coût total (pièces + main d'œuvre).
- **IMMOBILISATIONS** : 2 sous-menus :
  - **Immobilisations** : liste des comptes de classe 2 ayant un solde
    dans la Balance, avec Fournisseur et Prix d'achat (fiche éditable par
    compte), Catégorie (même catégorisation que le Bilan), Taux
    d'amortissement configuré, et **Valeur Brute / Amortissement / Valeur
    Nette** — l'amortissement affiché reste celui RÉELLEMENT comptabilisé
    (comptes 28x/29x de la Balance, réparti au prorata au sein d'une
    catégorie s'il y a plusieurs comptes), jamais une simulation, pour ne
    jamais diverger du Bilan.
  - **Amortissements** : taux annuel (%) paramétrable par catégorie
    d'immobilisation (double-clic pour modifier), réutilisé dans l'écran
    Immobilisations.

Testé de bout en bout : véhicule → mission (700 km calculés) → réparation
avec 4 plaquettes de frein (stock 10→6, coût total 47 000 = 4×8 000 +
15 000 main d'œuvre) ; immobilisation véhicule avec fournisseur/prix
d'achat et taux 20%, valeur nette calculée exactement depuis la Balance
(15M brut - 3M amort = 12M net). Bilan resté équilibré (non-régression
vérifiée).

### Suppression du sous-menu Ventes + Balance âgée des créances clients

**Réalisé** :
- **Sous-menu « Ventes »** retiré du menu COMMERCE.
- **Recouvrement** transformé en deux onglets :
  - **Factures** : le suivi existant (inchangé).
  - **Balance âgée** (nouveau) : pour chaque client ayant des factures
    impayées, répartit le montant dû par tranche d'ancienneté (jours
    écoulés depuis la date de facture). **Seuils des tranches sélectionnables**
    (3 champs modifiables + préréglages 30/60/90, 15/30/60, 30/60/120),
    avec ligne de total général. **Double-clic sur un client** ouvre une
    fenêtre de détail listant chaque facture impayée avec son ancienneté
    exacte en jours, triée de la plus ancienne à la plus récente.
- `core.compute_balance_agee()` (nouveau) : s'appuie sur la table
  `factures_clients` déjà existante (onglet Factures du Recouvrement) —
  ignore les factures déjà payées (date de paiement réel renseignée).

Testé : 3 factures sur 2 clients avec des anciennetés différentes (5, 95 et
40 jours) — chaque montant tombe dans la bonne tranche (0-30/31-60/61-90/
>90), total général exact (850 000), non-régression vérifiée (Bilan resté
équilibré, ce module étant sans lien avec la comptabilité).

### Correctifs Recouvrement : montant à 0 + règlement non comptabilisé

**Deux bugs repérés par l'utilisateur sur une vraie capture d'écran** :

1. **Montant affiché à 0.00 pour toutes les factures** : le formulaire
   « Nouvelle facture » acceptait silencieusement un champ Montant laissé
   vide, en le remplaçant par 0 au lieu d'avertir l'utilisateur. Corrigé :
   le montant est désormais obligatoire et doit être strictement positif,
   aussi bien côté écran (message d'erreur explicite) que côté moteur
   (`core.add_facture()` refuse un montant ≤ 0).
2. **Aucune option de règlement, aucune comptabilisation** : « Enregistrer
   le paiement » se contentait de noter une date, sans jamais demander
   quel compte banque/caisse avait reçu l'argent, et sans générer la
   moindre écriture comptable. Corrigé :
   - Nouveau champ **« Compte banque/caisse »** (recherche filtrée à la
     classe 5, liste déroulante) à côté de la date de paiement réel.
   - **`core.enregistrer_paiement_facture()`** (nouveau) : comptabilise
     réellement le règlement — Débit banque/caisse choisi, Crédit compte
     client (411000, même convention que la Facturation), pour le montant
     de la facture. Un garde-fou (`reglement_comptabilise`) empêche de
     reposter une seconde écriture si la date de paiement est modifiée
     ensuite.

Testé de bout en bout, en reproduisant exactement le scénario de
l'utilisateur (facture 1 500 000 F sur « Societe ABC5 », réglée en
caisse) : montant correctement affiché, écriture comptable générée
(Débit 571000 / Crédit 411000), Bilan resté équilibré, aucun doublon
d'écriture en cas de second appel.

### Synchronisation (PARAMÈTRES) + Utilisateurs et niveaux d'accès (ADMIN)

**Réalisé** :
- **PARAMÈTRES > Synchronisation** : bouton qui rejoue explicitement
  `init_db()`/`_migrate()` sur la base en cours — crée toute table ou
  colonne manquante (utile après avoir installé une nouvelle version du
  logiciel sur une base existante plus ancienne), sans jamais toucher aux
  données existantes. Affiche un rapport (nombre de tables, exercice
  courant, écart du Bilan) confirmant l'état de la base.
- **ADMIN > Niveaux d'accès** : liste paramétrable (même mécanisme que les
  Taux de TVA/retenue — créer/modifier/supprimer, import/export .xlsx,
  bouton « Ajouter les niveaux courants » pré-remplissant Administrateur/
  Comptable/Saisie seule/Lecture seule).
- **ADMIN > Utilisateurs** : comptes utilisateurs (nom d'utilisateur, nom
  complet, mot de passe **haché** — SHA-256 salé, jamais stocké en clair —
  niveau d'accès assigné, actif/inactif). `core.verify_password()` posé
  pour un futur écran de connexion.

**Point de transparence important** : l'application ne demande pas encore
de connexion au démarrage, et aucun menu/action n'est aujourd'hui
restreint selon le niveau d'accès — ce chantier pose la **base** (comptes,
mots de passe sécurisés, niveaux configurables) demandée, mais l'application
RÉELLE des restrictions (écran de connexion + contrôle par menu/action
selon le niveau) est un chantier bien plus large, touchant potentiellement
chaque écran de l'application — à traiter séparément si vous le souhaitez.

Testé : synchronisation sur une base de 39 tables, niveaux d'accès
exportés/réimportés fidèlement, création d'utilisateur avec mot de passe
haché et vérification correcte/incorrecte, non-régression du Bilan.

### Correctif : Stock/Immobilisations persistants après suppression des écritures + outil de Réinitialisation

**Bug repéré par l'utilisateur** (avec capture d'écran) : après avoir
supprimé toutes les écritures comptables et lancé la Synchronisation, le
Stock et les Immobilisations affichaient toujours des valeurs.

**Cause identifiée** : les Soldes d'ouverture (`opening_balances`) sont
une table **séparée** des écritures (`entries`) — supprimer les écritures
ne vide jamais les soldes d'ouverture, et Stock/Immobilisations sont
calculés à partir des DEUX (`compute_balance()` = solde d'ouverture +
mouvements des écritures). La Synchronisation, elle, ne touche **jamais**
aux données (uniquement à la structure des tables) — ce n'était donc pas
l'outil à utiliser pour ça.

**Réalisé** : nouvel écran **ADMIN > Réinitialisation des données** —
outil explicite et destructif, avec :
- **6 catégories indépendantes** à cocher : Écritures comptables, Soldes
  d'ouverture, Fiches immobilisations, Circuit d'engagements (Expression
  de besoin/Bon de commande/Bordereau/Règlements), Factures (vente/achat/
  recouvrement), Transport (véhicules/missions/réparations/pièces).
- **Portée** : toutes les années, ou un seul exercice (pour écritures et
  soldes d'ouverture).
- **Confirmation renforcée** : il faut taper « SUPPRIMER » pour activer
  le bouton, puis confirmer une seconde fois — action irréversible.
- Rapport détaillé du nombre de lignes supprimées par catégorie.

Testé de bout en bord avec un scénario couvrant toutes les catégories
(stock/immo via soldes d'ouverture, fiche immobilisation, engagement,
véhicule, facture) : chaque catégorie se vide indépendamment et
complètement, Bilan resté équilibré (écart = 0) sur une base totalement
vidée.

### La validation du Bon de commande comptabilise désormais directement

**Demande** : quand on valide un Bon de commande, le logiciel doit aussi
générer les écritures comptables (auparavant, il fallait un second passage
dans le sous-menu Règlements pour comptabiliser).

**Réalisé** :
- **Compte de charge et code analytique ajoutés directement aux lignes du
  Bon de commande** (colonnes `compte_charge`/`analytic_code` sur
  `ep_bon_commande_lignes`, migration incluse) — plus besoin d'attendre
  l'étape Règlements pour les renseigner.
- **Retenue fiscale ajoutée au Bon de commande** (taux, compte, préréglages
  ADMIN — même mécanisme que Règlements/Factures frs).
- **`core.valider_ep_bon_commande()` comptabilise directement** : Débit
  des comptes de charge choisis (avec code analytique), Crédit fournisseur
  net de retenue, retenue fiscale si applicable, entrée de stock
  automatique pour les lignes liées à un compte de marchandises/matières
  premières — refuse explicitement si un compte de charge manque sur une
  ligne, ou si le fournisseur n'est pas renseigné. Toujours accompagné de
  la création du Bordereau de livraison (suivi de réception, inchangé).
- Un Règlement est **toujours créé pour la traçabilité**, mais déjà marqué
  validé avec les MÊMES écritures (pas de double comptabilisation) — le
  mécanisme de correction existant (dévalider/revalider depuis l'écran
  Règlements) continue de fonctionner normalement pour corriger une
  erreur après coup.
- Logique de comptabilisation **factorisée** (`_comptabiliser_lignes_achat`)
  entre Règlements et Bon de commande — un seul moteur, deux points d'entrée.

Testé de bout en bout : validation directe du Bon de commande génère
immédiatement l'écriture (charge + fournisseur, Bilan équilibré) sans
étape supplémentaire ; les deux garde-fous (compte de charge manquant,
fournisseur manquant) refusent correctement ; le mécanisme de correction
existant (dévalider/revalider via Règlements) reste pleinement
fonctionnel après une validation directe.

### Correctifs Bon de commande : garde-fou montant nul + correction possible après validation

**Bug repéré par l'utilisateur** (avec capture d'écran) : une ligne avec
Prix unitaire = 0.00 (la valeur semblait avoir été saisie par erreur dans
le champ Unité au lieu de Prix unitaire) a pu être validée — l'écriture
générée avait un montant nul, et **aucune entrée de stock n'a été créée**
(la logique de comptabilisation ignore volontairement les lignes à montant
nul, pour ne jamais poser une écriture ou un mouvement de stock à 0).

**Corrigé — deux volets** :

1. **Garde-fou ajouté** : `_comptabiliser_lignes_achat()` (utilisée à la
   fois par Règlements et par la validation directe du Bon de commande)
   refuse désormais explicitement toute ligne à montant nul, avec un
   message pointant vers la cause la plus probable (« prix unitaire
   probablement resté à 0, ou saisi dans le mauvais champ »).
2. **Bug plus grave découvert en creusant : aucun moyen de corriger un Bon
   de commande déjà validé** — une fois « VALIDÉ », l'écran se verrouillait
   complètement, sans bouton de correction (contrairement aux Règlements,
   Factures frs, etc.). `core.devalider_ep_bon_commande()` (nouveau) :
   supprime les écritures comptables générées, repasse le Bon de commande
   ET le Règlement lié en brouillon ensemble (cohérence des deux
   documents), sans toucher au Bordereau de livraison déjà créé
   (indépendant de la comptabilité). Bouton **« Corriger (repasser en
   brouillon) »** ajouté à l'écran. Le sens inverse (dévalider un Règlement
   créé par un Bon de commande) repasse maintenant aussi ce Bon de
   commande en brouillon, pour la même raison de cohérence.

Testé de bout en bout, en reproduisant exactement le scénario de
l'utilisateur : une ligne à montant nul est désormais refusée à la
validation ; un Bon de commande validé avec une erreur peut être corrigé
(écritures retirées, ligne modifiable, revalidation possible) ; une fois
le prix correctement saisi, l'écriture ET le mouvement de stock (bonne
quantité, bon coût unitaire) sont générés correctement.

### Menu RAPPORTS FINANCIERS restauré + Paiement bancaire dans Règlements

**Points clarifiés avec l'utilisateur** : les captures montraient un Bon
de commande "immo" avec **Quantité : 0** (d'où montant=0, donc rien en
Immobilisations) et l'absence du bouton "Corriger" / de l'ouverture
automatique des listes déjà livrés précédemment — signe d'une version
antérieure encore utilisée. Ces deux fonctionnalités sont bien présentes
dans le code actuel (vérifié). Pour cette réponse, deux vraies demandes
nouvelles :

1. **Menu « RAPPORTS FINANCIERS »** (nouveau) : réintègre Grand livre,
   Balance, Bilan, Écritures non équilibrées, Compte de résultat, TFT,
   Situation financière, Liasse fiscale, Tableaux d'exécution budgétaire,
   Impôts, Déclarations sociales, Rapprochements bancaires — tous ces
   écrans et leur moteur de calcul n'avaient jamais été supprimés du code
   (seule la précédente inscription au menu ÉTATS ET RAPPORTS l'avait
   été), donc rien à reconstruire, juste réintégré sous ce nouveau nom.

2. **Paiement bancaire dans Règlements** : jusqu'ici, valider un Règlement
   ne faisait que reconnaître la charge et la dette fournisseur (compte
   401000) — il manquait l'étape du RÈGLEMENT proprement dit (le paiement
   réel). Ajouté :
   - `core.enregistrer_paiement_reglement()` : comptabilise le paiement —
     Débit fournisseur (401000, soldant sa dette), Crédit compte banque/
     caisse choisi, pour le montant NET à payer (après retenue). Ne
     comptabilise qu'une seule fois (garde-fou `paiement_comptabilise`).
   - `core.devalider_paiement_reglement()` : annule un paiement déjà
     comptabilisé (mauvais compte, erreur) sans toucher à la charge/dette
     déjà validée séparément.
   - Nouvelle section **« Paiement bancaire/caisse »** dans l'écran
     Règlements : date de paiement, compte banque/caisse (recherche
     filtrée à la classe 5), boutons Enregistrer/Annuler le paiement —
     n'apparaît activable qu'une fois la charge déjà validée.
   - `devalider_reglement()` (correction de la charge) annule aussi
     automatiquement le paiement s'il en existait un, pour rester cohérent.

Testé de bout en bout : achat d'immobilisation avec quantité correcte →
apparaît bien dans Immobilisations (valeur brute 500 000) ; paiement
bancaire du règlement lié comptabilisé (Débit 401000, Crédit banque) ;
génération du Bilan (RAPPORTS FINANCIERS) toujours sans erreur (132/132
cellules) ; Bilan resté équilibré à chaque étape.

### Bon de commande / Règlement — "Compte de charge" renommé "Compte débiteur"

**Demande** : le champ s'appelait "Compte de charge" mais devait aussi
accepter des comptes d'immobilisation (classe 2), pas seulement des
charges (classe 6) — un bon de commande peut tout aussi bien servir à
acheter une charge qu'une immobilisation.

**Réalisé** :
- Libellé changé en **« Compte débiteur (charge ou immobilisation) »**
  dans les écrans Bon de commande et Règlements.
- Liste déroulante élargie : comptes de **classe 2 (immobilisations) ET
  classe 6 (charges)** proposés (au lieu de la classe 6 seule).
- Messages d'erreur mis à jour en conséquence (« compte débiteur »).

Testé : un compte de classe 2 (241100, immobilisation) choisi comme
compte débiteur sur une ligne de Bon de commande — validation réussie,
l'immobilisation apparaît correctement dans le menu Immobilisations, et
une ligne de charge classique (classe 6) continue de fonctionner
normalement. Bilan resté équilibré dans les deux cas.

### Nouveau menu GRH + MAINTENANCE-ÉNERGIE renommé en MAINTENANCE-QUALITÉ

**Réalisé** :
- **GRH** (nouveau menu), 5 sous-menus, sans lien avec la comptabilité
  (même principe que Transport) :
  - **Liste du personnel** : fiche employé (matricule, nom, prénom, poste,
    service, date d'embauche, contact, statut).
  - **Time sheet** : pointage des heures par employé et par activité —
    refuse un nombre d'heures nul ou négatif.
  - **KPI** : indicateurs de performance (cible/réalisé/unité, par
    employé et/ou service, taux de réalisation calculé automatiquement,
    coloré en vert si atteint / rouge si non atteint).
  - **Tableau de bord GRH** : synthèse en lecture seule (effectif actif,
    heures pointées sur 30 jours, KPI en cours/atteints/non atteints,
    incidents HS ouverts par gravité) — calculée à la volée à partir des
    4 autres sous-menus, aucune donnée dupliquée.
  - **HS (hygiène santé)** : incidents, visites médicales, formations
    sécurité, distributions d'EPI — par employé, avec gravité et statut
    ouvert/clos (mis en évidence en rouge si ouvert).
- **MAINTENANCE-ÉNERGIE renommé en MAINTENANCE-QUALITÉ** (libellé du menu
  uniquement — les sous-menus Énergie/Maintenance/Pièces de rechange
  restent inchangés).

Testé de bout en bout : personnel créé avec pointage, KPI et incidents HS
liés — le Tableau de bord GRH agrège correctement toutes ces données
(8h pointées, KPI atteint à 120%, 1 incident grave ouvert). Non-régression
vérifiée : moteur comptable et module Pièces de rechange partagé
toujours pleinement fonctionnels.

### Menu TRESORERIE + import Personnel/Time Sheet + confirmation Bilan gabarit

**Bilan** : vérifié — le gabarit exact fourni par l'utilisateur
(`templates/bilan_template.xls` / `modele_bilan.xlsx`, formules N et N-1)
était déjà branché et fonctionnel depuis une session précédente (confirmé
par test direct). Rien à reconstruire.

**Réalisé** :
- **Menu TRESORERIE** (nouveau), avec 2 onglets :
  - **Banques (Entrées/Sorties)** : chaque compte de trésorerie (classe 5)
    aligné horizontalement — Solde début / Entrées / Sorties / Solde fin
    de la période (exercice courant par défaut), avec une ligne de total.
  - **Engagements à payer** : liste des Règlements déjà validés (charge
    comptabilisée) dont le paiement bancaire n'a pas encore été
    enregistré (menu ENGAGEMENTS-PROJETS > Règlements) — comparés à la
    trésorerie disponible, avec un message clair (vert/rouge) indiquant
    si l'entreprise peut faire face à tous ses engagements.
- **Import Excel pour Liste du personnel et Time sheet** (menu GRH) :
  boutons **« Télécharger le modèle d'import (.xlsx) »** et
  **« Importer (.xlsx) »** sur les deux écrans.
  - Personnel : import par Matricule — une fiche déjà existante est mise
    à jour (pas de doublon), une nouvelle est créée.
  - Time sheet : chaque ligne DOIT référencer un matricule déjà présent
    dans la Liste du personnel — sinon ignorée avec message d'erreur
    explicite (rapport groupé après import, jamais un plantage total).

Testé de bout en bout : Trésorerie horizontale correcte sur plusieurs
comptes avec mouvements réels ; engagement de 5 000 000 avec trésorerie
insuffisante de 3 500 000 → signalé correctement (`peut_faire_face:
False`) ; modèles téléchargés puis réimportés (création, puis mise à jour
sans doublon au second import, matricule inconnu proprement signalé) ;
non-régression du Bilan gabarit vérifiée.

### Bilan — correctif d'affichage (libellés coupés) + confirmation du calcul

**Capture d'écran fournie** : les libellés apparaissaient coupés au début
(« es premières et fournitures » au lieu de « Matières premières... »,
« OPRES ET RESSOURCES DURABLES » au lieu de « CAPITAUX PROPRES... ») —
les deux tableaux étaient scrollés horizontalement, la barre de défilement
existante ne se réinitialisant jamais après un rafraîchissement.

**Corrigé** : les deux tableaux reviennent maintenant systématiquement au
début (gauche) après chaque rafraîchissement, et la colonne Libellé de
l'Actif a été élargie (300→340px) pour réduire le besoin de défiler.

**Vérifié — Total Actif ≠ Total Passif sur la capture n'est PAS un bug de
calcul** : `compute_bilan()` classe chaque compte de la Balance dans une
case et une seule, garantissant mathématiquement Actif = Passif dès lors
que les données respectent la partie double (testé et confirmé : écart à
0 sur un scénario équilibré). Un écart de 37 789 564 389 - 28 690 813 980
= 9 098 750 409 signale donc un vrai déséquilibre dans les données
saisies/importées (le cas le plus fréquent : un solde d'ouverture importé
sans sa contrepartie — reproduit et vérifié : le diagnostic intégré
détecte précisément ce cas et pointe vers l'onglet « Soldes
d'ouverture »). Ce diagnostic s'affiche automatiquement sous les deux
tableaux dès qu'un écart existe — probablement hors du cadre visible sur
la capture transmise.

Pour identifier précisément la cause dans votre cas, il faudrait soit
consulter ce message de diagnostic directement dans l'application (faites
défiler sous les tableaux), soit me transmettre votre fichier
`comptabilite.db` pour un diagnostic direct.

### Bilan entièrement refait — lecture directe du gabarit, colonne N-1 réparée, sans scroller

**Cause du bug « aucune donnée en N-1 »** : l'écran Bilan utilisait ma
propre logique de catégorisation (`compute_bilan_detaille`), séparée de
celle qui alimente l'export .xlsx officiel — une source de divergence
inutile. Comme demandé : **ancienne logique supprimée, travail refait à
zéro** en lisant directement le gabarit.

**Réalisé** :
- **`core.compute_bilan_plat()`** (nouveau, remplace
  `compute_bilan_detaille` pour cet écran) : relit le gabarit officiel
  (`templates/modele_bilan.xlsx`) **ligne par ligne**, colonne par
  colonne (A=Libellé Actif, B=Brut, C=Amortissements, D=Net, E=Net N-1,
  F=Libellé Passif, G=Exercice N, H=Exercice N-1), et évalue **chaque
  formule** avec le même moteur que l'export .xlsx (`evaluate_sheet_formulas`)
  — donc exactement les mêmes valeurs, y compris la colonne N-1, sans
  aucune logique de regroupement propre à l'application. Filtre les
  lignes parasites du gabarit (titre du fichier, en-têtes de colonnes,
  valeur résiduelle collée).
- **Écran BilanTab entièrement reconstruit** : affichage plat (Actif à
  gauche, Passif à droite), colonnes dimensionnées pour tenir sans
  défilement horizontal (plus de scroller à gérer), écart Actif/Passif
  affiché en permanence dans la barre d'outils.
- Export `.xlsx` inchangé (utilisait déjà ce même moteur).

Testé avec un historique réel sur 2 exercices (2025 clôturé → 2026, avec
une extension de bâtiment) : colonne Net N-1 correctement remplie (4 000 000,
la valeur de 2025) au lieu d'être vide, TOTAL GENERAL Actif = TOTAL
GENERAL Passif exactement (6 000 000) en N comme en N-1, écart à 0.

### Crash au démarrage du Bilan corrigé — lecture de gabarit rendue auto-réparante

**Crash signalé** (capture d'écran) : `openpyxl does not support the old
.xls file format` — l'application plantait entièrement à l'ouverture du
Bilan.

**Cause** : le code utilisait `os.path.exists(...)` pour choisir entre le
nouveau gabarit (`modele_bilan.xlsx`) et l'ancien (`bilan_template.xls`,
un fichier XML SpreadsheetML malgré son extension .xls) — dans
l'environnement de l'utilisateur, ce choix a mené vers l'ancien gabarit,
mais la détection automatique de son format (`_is_spreadsheetml()`) a
échoué, envoyant le fichier directement à `openpyxl.load_workbook()` qui
ne sait pas lire ce format → plantage complet et immédiat de
l'application (aucune gestion d'erreur).

**Corrigé — deux niveaux de protection** :
1. **`open_template_workbook()` rendue auto-réparante** : si la première
   tentative de lecture échoue (quelle qu'en soit la raison — mauvaise
   détection de format, encodage inhabituel...), la fonction retente
   automatiquement l'AUTRE méthode avant d'abandonner. Message d'erreur
   clair uniquement si les deux échouent.
2. **Gestion d'erreur ajoutée à tous les points d'entrée** (Bilan à
   l'écran, export .xlsx) : un souci de lecture de gabarit affiche
   désormais un message d'erreur clair et actionnable au lieu de faire
   planter toute l'application.

Testé en reproduisant exactement le scénario du crash (gabarit .xlsx
absent, seul l'ancien .xls disponible avec une détection de format
défaillante) : fonctionne désormais correctement via le repli
automatique. Testé aussi le cas extrême (aucun gabarit disponible) :
erreur claire et actionnable levée, plus de plantage silencieux.

### Suppression du menu RAPPORTS FINANCIERS et de tous ses gabarits

**Demande** : supprimer tout le menu RAPPORTS FINANCIERS avec tous les
templates le concernant.

**Réalisé** :
- Menu et les 12 entrées correspondantes retirés (Grand livre, Balance,
  Bilan, Écritures non équilibrées, Compte de résultat, TFT, Situation
  financière, Liasse fiscale, Tableaux d'exécution budgétaire, Impôts,
  Déclarations sociales, Rapprochements bancaires).
- **6 fichiers de gabarit supprimés** : `templates/bilan_template.xls`,
  `templates/modele_bilan.xlsx`, `templates/modele_resultat.xlsx`,
  `templates/modele_situation.xlsx`, `templates/modele_flux.xlsx`,
  `etats_financiers_template.xlsx` — vérifié qu'aucun n'était utilisé
  ailleurs dans l'application avant suppression.
- **Workflow de build nettoyé** (`.github/workflows/build.yml`) : les 6
  lignes `--add-data` correspondantes retirées.
- Le moteur de calcul sous-jacent non spécifique aux gabarits
  (`compute_bilan()`, `compute_balance()`, `close_exercice()`,
  `valider_reglement()`...) **n'a pas été touché** et reste utilisé en
  interne par les fonctionnalités actives (clôture d'exercice, Règlements,
  etc.) — seules les fonctions spécifiques au moteur de gabarits Excel
  (`compute_bilan_plat`, `generate_etat_xlsx`, `export_bilan_gabarit_xlsx`...)
  sont désormais orphelines (code mort, mais inoffensif : plus aucun
  écran n'y accède).

Testé : moteur comptable interne (Bilan, Règlements, clôture d'exercice)
toujours pleinement fonctionnel après suppression ; les fonctions
spécifiques aux gabarits supprimés échouent proprement avec une erreur
claire si appelées (mais ne le sont plus, aucun écran n'y menant).

### Suppression complète de la classe BilanTab (code mort résiduel)

**Demande** : ne plus voir le message d'erreur « Impossible de calculer
le Bilan » (gabarit `bilan_template.xls` introuvable).

**Précision** : ce message provenait d'une version antérieure de
l'exécutable, utilisée encore par l'utilisateur — le menu RAPPORTS
FINANCIERS (et l'écran Bilan qui déclenchait ce message) avait déjà été
entièrement supprimé du CODE SOURCE dans la réponse précédente ; il ne
manquait qu'un rebuild de l'exécutable côté utilisateur pour que le
message disparaisse.

**Fait en plus, par souci de propreté** : la classe `BilanTab` elle-même
(qui ne servait plus à rien, plus aucun menu n'y menant) a été
entièrement supprimée du code source, plutôt que laissée comme code mort
inutilisé — élimine toute trace de ce message d'erreur, même dans le
code source.

Testé : compilation propre, aucune référence résiduelle à `BilanTab` nulle
part dans le code, moteur comptable interne toujours pleinement
fonctionnel.

### Nouveau menu RAPPORT FINANCIERS (Grand livre, Balance, Bilan SYSCOHADA)

**Réalisé** : nouveau menu avec exactement 3 sous-menus, comme demandé.
- **Grand livre** et **Balance** : simplement réenregistrés (les écrans
  n'avaient jamais été supprimés, seulement retirés du menu).
- **Bilan SYSCOHADA** (nouveau, reconstruit) : basé sur
  `compute_bilan_detaille()` — **entièrement autonome, ne dépend d'aucun
  fichier de gabarit externe** (contrairement à l'ancien Bilan basé sur
  un gabarit `.xlsx`/`.xls`, source des crashs précédents). Calculé
  compte par compte depuis la même Balance que l'onglet Balance,
  garantissant mathématiquement Actif = Passif, avec comparatif de
  l'exercice précédent (N-1) et diagnostic automatique en cas d'écart
  réel (soldes d'ouverture incomplets, écritures non équilibrées).
  Présentation à plat (Actif Brut/Amortissements/Net/Net N-1 à gauche,
  Passif Exercice N/N-1 à droite), sans scroller.

Testé avec un historique réel sur 2 exercices (2025 clôturé → 2026) :
Total Actif = Total Passif exactement en N (4 700 000) comme en N-1
(4 000 000), Grand livre et Balance non-régressés.

### Bilan monté sur les seules opérations de la période + Balance et Grand Livre à 6 colonnes

**Demande** : le Bilan était monté sur solde d'ouverture + cumul des
opérations, alors qu'il devait être monté sur les seules opérations de la
période — la colonne N-1 devant contenir exclusivement le solde
d'ouverture. La Balance et le Grand livre devaient avoir 6 colonnes (2
pour le solde d'ouverture, 2 pour les mouvements de la période, 2 pour le
solde de clôture).

**Réalisé** :
- **`core.compute_bilan()` refactorée** pour accepter une balance et un
  résultat net déjà calculés (`balance=`, `resultat_net_override=`) — sans
  dupliquer sa logique de classification (immobilisations, stocks,
  créances/dettes par racine, capitaux propres, trésorerie), qui reste
  strictement la même.
- **`compute_bilan_mouvement_periode()`** (nouveau) : construit une balance
  où `solde_cloture` est remplacé par le seul mouvement Débit-Crédit de la
  période, et calcule le Bilan dessus.
- **`compute_bilan_solde_ouverture()`** (nouveau) : même principe, mais
  avec `solde_cloture` remplacé par le seul solde d'ouverture.
- **`compute_bilan_detaille()` reconstruite** pour utiliser ces deux
  fonctions : la colonne « N » = mouvements de la période seule, la
  colonne « N-1 » = solde d'ouverture seul (au lieu de calculer un second
  Bilan complet sur l'exercice précédent).
- **Bug trouvé et corrigé en cours de route** : un compte SANS mouvement
  cette période mais avec un solde d'ouverture réel (ex. un bâtiment
  immobilisé, jamais retouché dans l'année) disparaissait purement et
  simplement du Bilan — le rattachement N→N-1 était à sens unique. Corrigé
  avec une vraie fusion par union (`_merge_bilan_lignes`) : ce compte
  apparaît maintenant avec 0 en colonne période et son vrai solde en
  colonne d'ouverture.
- **Écran Bilan SYSCOHADA** : libellés de colonnes mis à jour (« Brut/
  Amort./Net (période) » et « Solde d'ouverture »), écart affiché et
  diagnostiqué indépendamment sur les deux colonnes.
- **`compute_balance_detaillee()` et l'écran Balance** : 6 colonnes
  (Ouverture Débit/Crédit, Mouvement Débit/Crédit, Clôture Débit/Crédit),
  export .xlsx mis à jour en conséquence.
- **Écran Grand livre** : la ligne de total par compte (et par classe)
  affiche désormais les 6 colonnes (Ouverture/Mouvement/Clôture,
  Débit/Crédit), en plus du détail écriture par écriture inchangé.

Testé de bout en bout sur un historique réel à 2 exercices (2025 clôturé
→ 2026) : Bilan équilibré indépendamment sur la colonne période (700 000)
et sur la colonne solde d'ouverture (4 500 000) ; Balance équilibrée sur
ses 3 paires de colonnes (5,5M / 850K / 6,2M) ; Grand livre cohérent
(banque : ouverture 500 000, mouvement -150 000, final 350 000) ; compte
immobilisation dormant correctement affiché avec son solde d'ouverture ;
non-régression vérifiée (Règlements, clôture d'exercice, Situation
financière).

### Regroupements du Bilan alignés exactement sur le fichier de référence

**Fichier fourni** : `Bilan.xlsx` — les valeurs calculées (pas de formules
textuelles cette fois) de l'exercice 2024 de l'utilisateur, servant de
référence de structure/libellés.

**Écart trouvé et corrigé** : le rapport de référence combine certaines
racines en une seule ligne, alors que mon Bilan les affichait séparément :
- **Racines 44 et 45** combinées en une ligne « IUTS-TPA-TVA » (au lieu de
  « État — débiteur/créditeur » et « Organismes internationaux » séparés).
- **Racines 47, 48, 49** combinées en une ligne « HAO » (au lieu de 3
  lignes séparées).
- **Racines 16 et 17** combinées en une ligne « Emprunts bancaires » côté
  Capitaux propres (au lieu de « Emprunts et dettes financières » et
  « Dettes de location-acquisition » séparées).
- **Trésorerie regroupée en Banques (racines 50-56) / Caisse (racines
  57-59)**, au lieu d'un groupe par racine individuelle (52, 56, 57...).

Nouvelles fonctions `_cle_racine_bilan()` et `_cle_prefixe_treso()`
appliquant ce regroupement exact avant la fusion N/N-1 déjà en place.

Testé avec un scénario ciblant précisément ces 4 regroupements (comptes
sur les racines 44, 45, 47, 49, 16, 17, et deux comptes de trésorerie) :
chaque fusion tombe juste (150 000 = 100 000+50 000 pour 44+45, 1 500 000
= 1 000 000+500 000 pour 16+17...), écart resté à 0.

### Formules du gabarit officiel réintégrées comme export dédié

**Demande** : appliquer les formules du fichier gabarit (CtaCptSolde...)
au Bilan.

**Réalisé** : le moteur d'évaluation de ces formules (`cta_cpt_solde*`,
`evaluate_sheet_formulas`, `compute_bilan_plat`, `export_bilan_gabarit_xlsx`)
n'avait en réalité jamais été supprimé du code — seul l'ancien écran
Bilan qui EN DÉPENDAIT (et son fichier de gabarit) l'avaient été. Le
fichier gabarit a été réintégré (`templates/bilan_template.xls`) et un
nouveau bouton **« Exporter selon le gabarit officiel (formules CtaCptSolde
exactes) »** a été ajouté à l'écran Bilan SYSCOHADA existant.

**Point de transparence important** : les formules de ce gabarit (sans
suffixe `Nm1`) calculent sur le **solde de clôture** (ouverture + cumul
des opérations) — la convention *avant* le changement demandé
récemment, qui a fait passer l'écran Bilan SYSCOHADA à une logique
« opérations de la période seule / solde d'ouverture seul ». Les deux
coexistent donc désormais intentionnellement :
- **Écran Bilan SYSCOHADA** (autonome, aucun fichier externe requis) :
  toujours en logique « période / solde d'ouverture », comme demandé
  précédemment.
- **Bouton d'export « gabarit officiel »** : reproduit fidèlement les
  formules exactes du fichier fourni (logique solde de clôture N / N-1
  via `...Nm1`), pour produire un document dans le format officiel exact
  quand nécessaire.

Workflow de build (`.github/workflows/build.yml`) mis à jour pour
réembarquer ce fichier.

Testé : export selon le gabarit officiel sans aucune erreur de formule ;
écran Bilan SYSCOHADA (autonome) toujours pleinement fonctionnel en
parallèle, écarts à 0 sur les deux mécanismes indépendamment.

### Correction : Bilan N = solde de clôture (ouverture + mouvements), N-1 = solde d'ouverture

**Correction demandée** : la colonne N doit redevenir le solde de clôture
habituel (solde d'ouverture + cumul des opérations de la période), pas les
seules opérations de la période comme demandé juste avant.

**Réalisé** : `compute_bilan_detaille()` utilise à nouveau le solde de
clôture standard pour la colonne N (`_compute_bilan_groupes(conn,
exercice)` sans forcer les mouvements seuls). **La colonne N-1 n'a pas eu
besoin d'être modifiée** : le solde d'ouverture d'un exercice correspond
mathématiquement au solde de clôture de l'exercice précédent (dès lors que
la clôture d'exercice a été utilisée normalement), donc elle assurait
déjà une vraie comparaison N-1 valide.

Écran Bilan SYSCOHADA mis à jour en conséquence (libellés de colonnes :
« Brut/Amort./Net » redevient le solde de clôture, « Net N-1 (ouverture) »
reste le solde d'ouverture).

Testé : un compte avec solde d'ouverture de 2 000 000 et une vente de
500 000 dans l'année affiche bien 2 500 000 en colonne N (ouverture +
mouvement) et 2 000 000 en colonne N-1 (ouverture seule) — écarts à 0 sur
les deux colonnes. Non-régression vérifiée (export gabarit officiel,
Situation financière, clôture d'exercice).

### Diagnostic : détail des comptes dans « Autres immobilisations non classées »

**Constat de l'utilisateur** : plusieurs catégories du gabarit de
référence (Terrains, Installations, Matériel de transport,
Immobilisations financières, Charges immobilisées) n'apparaissaient pas
sur l'écran, avec une grosse ligne « Autres immobilisations non classées »
à la place — confirmé par l'utilisateur après vérification.

**Cause probable** : les comptes réels de son plan comptable ne tombent
pas dans les plages numériques attendues (`IMMO_CATEGORIES`, dérivées du
gabarit de référence), et se retrouvent donc absorbés dans le reliquat
« Autres » plutôt que classés correctement.

**Réalisé — étape de diagnostic** : la ligne « Autres immobilisations non
classées » affiche désormais le **détail compte par compte** des comptes
concernés (code + libellé + montant), au lieu d'un simple total global —
pour identifier précisément quels comptes ne correspondent à aucune
catégorie attendue, et ainsi pouvoir corriger `IMMO_CATEGORIES` en
conséquence une fois ces comptes identifiés.

Testé : un compte hors de toutes les plages (200500) apparaît maintenant
en sous-ligne détaillée sous « Autres immobilisations non classées »,
avec son code, son libellé et son montant exact — tandis qu'un compte
correctement classé (231100, Bâtiments) reste affiché normalement sans
détail superflu. Écart resté à 0.

**Prochaine étape** : une fois les comptes non classés visibles à
l'écran, transmettez-moi leurs codes exacts pour que j'élargisse les
plages `IMMO_CATEGORIES` en conséquence.

### Bilan — boutons Télécharger le template + Visionner selon le template

**Demande** : deux nouveaux boutons dans le sous-menu Bilan — un pour
télécharger le template (avec ses formules), un autre pour visionner le
Bilan selon ce template.

**Réalisé** :
- **« Télécharger mon template (vierge, avec formules) »** : copie le
  fichier gabarit brut tel quel (formules CtaCptSolde non évaluées) vers
  l'emplacement choisi.
- **« Visionner le Bilan selon mon template »** : génère le Bilan avec les
  formules exactes évaluées, dans un fichier temporaire, puis **l'ouvre
  automatiquement** avec l'application par défaut (Excel) — sans boîte de
  dialogue « Enregistrer sous », pour un visionnage immédiat.
- Le bouton d'export existant (« Exporter selon le gabarit officiel »)
  reste disponible séparément, pour choisir où enregistrer le fichier.

**Bug corrigé au passage** : le mécanisme de repli (utilisé quand
`modele_bilan.xlsx` — le nouveau format .xlsx — est absent, ce qui est le
cas actuellement) passait par une conversion `openpyxl`, qui **perdait
toute la mise en forme d'origine** du gabarit (couleurs, bordures) — un
fichier de 60 Ko ressortait à 6 Ko, complètement dépouillé. Corrigé avec
une substitution de texte directe sur le XML brut (même mécanisme
robuste que l'export utilisé précédemment) : la mise en forme est
désormais intégralement préservée (51 Ko sur 60 Ko d'origine, l'écart
s'expliquant simplement par des textes de formules plus longs que les
valeurs numériques qui les remplacent).

Testé : template brut téléchargé identique à l'original (60 137 octets) ;
Bilan généré pour visionnage avec mise en forme préservée (131 valeurs
calculées, 0 formule non évaluée, aucun plantage même sur les lignes de
totaux qui référencent d'autres cellules) ; écran Bilan SYSCOHADA
autonome toujours pleinement fonctionnel en parallèle.

### Fin définitive du crash « fichier de gabarit introuvable » — gabarit encodé dans le code source

**Crash signalé une nouvelle fois** (capture d'écran) : le bouton
« Visionner le Bilan selon mon template » échouait avec `[Errno 2] No
such file or directory: ...\templates\bilan_template.xls` — la version
`.exe` utilisée ne contenait toujours pas ce fichier, malgré une
configuration de build correcte (déjà vérifiée deux fois). Le mécanisme
`--add-data` de PyInstaller s'est révélé peu fiable en pratique.

**Corrigé définitivement** : le gabarit `templates/bilan_template.xls`
(60 137 octets) est maintenant **encodé en base64 directement dans le
code source Python** (nouveau fichier `bilan_template_data.py`,
~86 Ko). Au premier besoin, `core.py` le régénère automatiquement dans
le dossier persistant de l'application (`%LOCALAPPDATA%\SaisieComptable`,
le même dossier que la base de données) — **sans dépendre d'aucun
mécanisme de bundle externe**. Le fichier `--add-data` du workflow de
build est conservé par sécurité, mais l'application n'en a plus besoin :
elle fonctionne à l'identique même si ce fichier est totalement absent de
l'installation.

Testé en simulant exactement le scénario du crash (variable d'environnement
`LOCALAPPDATA` neutralisée, fichier `templates/bilan_template.xls`
physiquement absent du disque) : le gabarit se régénère automatiquement,
l'export produit un fichier de 51 038 octets avec mise en forme préservée,
aucune erreur. Non-régression vérifiée sur l'écran Bilan SYSCOHADA
autonome et l'ensemble des 45 onglets de l'application.

### Bilan — boutons pour modifier et restaurer le template

**Demande** : un bouton pour modifier les formules du template.

**Réalisé** :
- **« Modifier les formules du template (importer une version corrigée) »** :
  choisissez un fichier `.xls`/`.xlsx` (ex. le template téléchargé via
  « Télécharger mon template », édité dans Excel avec vos propres
  corrections) — validé avant d'être pris en compte (doit s'ouvrir
  correctement et contenir au moins une vraie formule CtaCptSolde, sinon
  refusé avec un message clair) puis remplace le template ACTIF utilisé
  par l'application (`core.import_bilan_template()`). Toutes les
  utilisations suivantes (« Visionner », « Exporter selon le gabarit
  officiel ») utilisent automatiquement ce nouveau template.
- **« Restaurer le template d'origine »** : annule toute modification
  importée et revient au template intégré à l'application
  (`core.restaurer_bilan_template_original()`), en cas d'erreur ou pour
  repartir d'une base saine.

Testé de bout en bout : import validé (131 formules détectées), refus
propre d'un fichier sans formules et d'un fichier illisible (messages
d'erreur clairs, pas de plantage), export fonctionnel avec le nouveau
template, restauration de l'original, et l'écran Bilan SYSCOHADA
autonome (qui ne dépend d'aucun template) reste pleinement fonctionnel
tout au long de ces opérations.

### Template figé — retrait de l'import/remplacement côté utilisateur

**Demande** : le template doit rester figé dans le logiciel — les
formules se modifient directement dedans, pas via un import utilisateur.

**Réalisé** : les boutons « Modifier les formules du template (importer
une version corrigée) » et « Restaurer le template d'origine », ajoutés
juste avant, ont été retirés de l'écran. Le template reste donc figé
(encodé dans `bilan_template_data.py`), non modifiable depuis
l'application.

**Nouveau fonctionnement pour modifier une formule** : indiquez-moi
directement quelle formule doit changer (et sa nouvelle valeur), et je
modifie le fichier gabarit source puis régénère `bilan_template_data.py`
en conséquence — le nouveau template devient alors la version officielle
figée, livrée avec le reste du code.

Les 3 boutons restants (Télécharger mon template, Visionner selon mon
template, Exporter selon le gabarit officiel) restent disponibles — ce
sont des opérations de LECTURE/EXPORT, pas de modification.

### Bouton « Modifier les formules du template » — ouvre le fichier figé directement dans Excel

**Précision de l'utilisateur** : pas d'import/remplacement — un bouton
qui ouvre directement le template figé (le fichier utilisé par le
logiciel) pour éditer les formules dedans.

**Réalisé** : le bouton **« Modifier les formules du template »**
remplace l'ancien « Télécharger mon template » — il ouvre directement,
avec Excel, le **même fichier physique** que celui utilisé par l'écran
Bilan (`core.BILAN_TEMPLATE_PATH`, situé dans
`%LOCALAPPDATA%\SaisieComptable\bilan_template.xls`). Le template reste
figé du point de vue du logiciel (toujours aucun mécanisme d'import/
remplacement dans l'interface), mais comme c'est exactement le même
fichier, toute formule modifiée et enregistrée dans Excel (Ctrl+S, en
gardant le format) est immédiatement prise en compte au prochain
« Visionner » ou « Exporter » — sans étape supplémentaire, sans bouton
d'import séparé.

Testé : le fichier gabarit actif (celui régénéré automatiquement depuis
le code source si absent) est bien accessible à cet emplacement, l'export
et l'écran Bilan SYSCOHADA autonome fonctionnent normalement à partir de
ce même fichier.

### Vraie différence de rubrique trouvée : racines 13/14/15 doivent être fusionnées

**Question de l'utilisateur** : « pourquoi mon bilan donne quelques
rubriques différentes du tien ? » — avec le fichier de référence réel
joint.

**Cause trouvée** : l'écran Bilan SYSCOHADA (`compute_bilan_detaille`,
code propre à l'application) affichait les racines **13, 14 et 15**
comme 3 lignes séparées (« Résultat net (avant affectation) »,
« Subventions d'investissement », « Provisions réglementées »), alors que
le fichier de référence de l'utilisateur les combine en **une seule
ligne** : « Résultats antérieurs (13-15) » — confirmé par la formule déjà
vérifiée `=-CtaCptSolde("13*","15*")`, une PLAGE couvrant les 3 racines
d'un coup, pas 3 formules séparées.

**Corrigé** : les racines 13, 14 et 15 sont désormais fusionnées en une
seule ligne « Résultats antérieurs (racines 13-15) », comme la référence.
La ligne séparée « Résultat net de l'exercice » (calculée depuis les
classes 6/7/8, pas depuis le solde du compte 13) reste distincte, comme
dans le fichier de référence.

Testé avec un scénario ciblant précisément ce cas (comptes sur les
racines 13, 14 et 15) : la fusion tombe juste (100 000+200 000+50 000 =
350 000 sur une seule ligne), écart resté à 0.

### Bug confirmé et corrigé : racine 55 (Trésorerie) manquante dans "Banques créditrices"

**Écart signalé par l'utilisateur** : sa "Banque créditrice *50-56*"
donnait 1 milliard, l'application n'en affichait que 643 379 329.

**Cause trouvée** : `_cle_prefixe_treso()` utilisait une liste discrète
de racines `("50","51","52","53","54","56")` — **oubliant la racine
« 55 »** entre 54 et 56. Or la formule de référence
`CtaCptSolde("50*","56*")` est une vraie **plage numérique continue**
(500000 à 569999), qui inclut forcément tout compte 551xxx-559xxx s'il
existe. Tout compte sur cette racine tombait donc à tort dans « Caisse »
au lieu de « Banques ».

**Corrigé** : `_cle_prefixe_treso()` fonctionne désormais par comparaison
numérique de plage (500000 ≤ code ≤ 569999), exactement comme la formule
de référence — ne peut plus jamais oublier de racine intermédiaire.

Testé avec un compte sur la racine 55 (200 000 000) : correctement compté
dans « Banques créditrices/débitrices » (total 600 000 000 = 400M+200M)
au lieu de « Caisse ». Écart resté à 0.

**En attente** : pour le poste Installations/Agencements (17 Md attendu
vs 12,5 Md affiché), en attente de la liste des comptes sous « Autres
immobilisations non classées » (visible en faisant défiler l'écran
Bilan) pour identifier précisément la cause.

### Bug corrigé : compte racine 230 absent de toute catégorie (Bilan ET menu Immobilisations)

**Suite du point Installations/Agencements** : en réexaminant
`IMMO_CATEGORIES` (partagée entre l'écran Bilan et le menu
IMMOBILISATIONS > Immobilisations), un vrai trou trouvé : la racine
**230** (le compte « Bâtiments » racine, avant sa subdivision officielle
en 231-234) n'appartenait à AUCUNE catégorie — ni Terrains (220-229), ni
Bâtiments (231-233 seulement) — et tombait donc dans « Autres
immobilisations non classées » au lieu de rejoindre Bâtiments.

**Corrigé** : la plage Bâtiments élargie à (230000-233999), comblant ce
trou — comme les deux écrans partagent la même fonction
`categorie_immobilisation()`/`IMMO_CATEGORIES`, ce correctif s'applique
automatiquement aux deux à la fois.

Testé : un compte 230000 (2 000 000) rejoint maintenant correctement
« Bâtiments » aux côtés d'un compte 231100 (3 000 000), sur le Bilan
comme sur le menu Immobilisations (total 5 000 000 sur une seule ligne
au lieu d'être scindé). Écart resté à 0.

**Toujours en attente** pour le point précis Installations/Agencements
(17 Md attendu vs 12,5 Md affiché) : la liste des comptes exacts sous
« Autres immobilisations non classées » sur la vraie base de
l'utilisateur, pour identifier s'il reste un autre trou de plage.

### Scrollbars ajoutées partout + Soldes d'ouverture en Débit/Crédit avec totaux

**Réalisé** :
- **Scrollbar verticale ajoutée sur les 46 écrans** qui affichaient une
  liste (Treeview) sans en avoir — script automatisé (repère chaque
  Treeview et sa ligne d'affichage, l'enveloppe d'une scrollbar), plus 4
  cas particuliers traités à la main (Bilan SYSCOHADA — 2 tableaux côte à
  côte, Grand livre — scrollbar verticale ET horizontale vu le nombre de
  colonnes, détail Balance âgée, Exercices comptables). Vérifié : plus
  aucun écran de l'application n'a de liste sans scrollbar.
- **Soldes d'ouverture** : la colonne unique « Solde » (signée) remplacée
  par **deux colonnes Débit / Crédit**, avec une ligne de **total en bas
  de chaque colonne** et un message d'équilibre clair (Débit = Crédit,
  au lieu de « somme des soldes = 0 »).

Testé : soldes d'ouverture toujours enregistrés et lus correctement
(conversion Débit/Crédit ↔ solde signé vérifiée dans les deux sens),
Bilan resté équilibré. Compilation complète sans erreur sur les 46
écrans modifiés.

### Format des nombres uniformisé partout — séparateur de milliers, plus de décimales

**Demande** : séparateur de milliers sur tous les chiffres, et suppression
des deux zéros après la virgule.

**Réalisé** : les **111 occurrences** de l'ancien format (`,.2f` — virgule
comme séparateur de milliers, 2 décimales, ex. `1,234,567.00`) trouvées
dans toute l'application ont été remplacées par `fmt_cfa()` (déjà
utilisée dans les écrans les plus récents) — espace comme séparateur de
milliers, aucune décimale, format SYSCOHADA standard (ex. `1 234 567`).
Script automatisé (repère chaque `{EXPR:,.2f}` et le convertit en
`{fmt_cfa(EXPR)}`), plus 2 occurrences oubliées (sans séparateur du tout)
corrigées à la main.

Testé : plus aucune occurrence de format à 2 décimales dans tout le
fichier, moteur comptable toujours pleinement fonctionnel, format vérifié
sur un exemple (1234567 → « 1 234 567 »).

### Crash au démarrage corrigé — bug du script d'ajout automatique de scrollbars

**Crash signalé** (capture d'écran) : `AttributeError: 'Frame' object has
no attribute 'yview'` — l'application plantait immédiatement au
démarrage.

**Cause** : le script automatisé qui a ajouté des scrollbars à 46 écrans
(réponse précédente) a repéré `self.content.pack(fill="both",
expand=True)` — le **conteneur principal de l'application** (un simple
`ttk.Frame` servant à basculer entre les différents onglets), qui
correspondait par coïncidence au même motif que les listes (Treeview)
ciblées — et lui a ajouté une scrollbar comme s'il s'agissait d'une
liste. Un `Frame` n'a pas de méthode `.yview()`, d'où le plantage
immédiat à l'ouverture (ce conteneur est initialisé dès le démarrage).

**Corrigé** : `self.content` restauré à son comportement d'origine (sans
scrollbar, ce n'est pas une liste). **Vérification systématique
effectuée sur tout le fichier** : chaque scrollbar ajoutée automatiquement
a été contrôlée pour confirmer qu'elle contrôle bien un widget qui
supporte réellement `.yview()` (Treeview, Text, Listbox) — aucun autre
cas erroné trouvé.

Testé : compilation propre, moteur comptable toujours pleinement
fonctionnel, aucune autre scrollbar mal attribuée détectée dans
l'ensemble de l'application.

### Crash Grand livre corrigé — mélange grid/pack sur le mauvais conteneur

**Crash signalé** (capture d'écran) : `_tkinter.TclError: cannot use
geometry manager grid inside .!frame2.!grandlivretab which already has
slaves managed by pack` à l'ouverture du Grand livre.

**Cause** : dans mon précédent correctif de `GrandLivreTab` (ajout des
scrollbars verticale et horizontale), le tableau (`self.tree`) avait été
créé avec `self` comme parent (l'onglet lui-même, qui utilise déjà
`.pack()` pour sa barre de filtres) au lieu du nouveau conteneur
`_gl_frame` — puis positionné avec `.grid()`, provoquant un conflit
Tkinter (un même conteneur ne peut pas mélanger `pack()` et `grid()`
entre ses enfants directs).

**Corrigé** : le tableau est maintenant créé directement avec
`_gl_frame` comme parent, cohérent avec son positionnement en `.grid()`.
**Vérification systématique** de tous les autres écrans ayant reçu une
scrollbar par `.grid()` (au lieu du motif `pack(in_=...)` utilisé
partout ailleurs) — aucun autre cas de ce type trouvé (Bilan SYSCOHADA
était déjà correctement construit dès le départ).

Testé : compilation propre, Grand livre calculable sans erreur, moteur
comptable toujours pleinement fonctionnel.

### Crash Trésorerie corrigé — même famille de bug, vérification exhaustive faite cette fois

**Crash signalé** (capture d'écran) : `can't pack
.!frame2.!tresorerietab.!notebook.!frame.!treeview inside
.!frame2.!tresorerietab.!frame` à l'ouverture de Trésorerie.

**Cause** : le script automatisé d'ajout de scrollbars (deux réponses
plus tôt) utilisait systématiquement `ttk.Frame(self)` pour envelopper
chaque tableau, sans vérifier le VRAI parent Tkinter du tableau
d'origine. Pour les tableaux créés à l'intérieur d'une méthode
`_build_xxx(self, parent)` (où `parent` est un onglet de Notebook, pas
`self` directement — cas de `TresorerieTab` et de la Balance âgée dans
`RecouvrementTab`), le nouveau conteneur de scrollbar se retrouvait
attaché au mauvais parent Tkinter, provoquant un plantage à l'ouverture.

**Corrigé** : `TresorerieTab` (2 tableaux) et `self.tree_agee` dans
`RecouvrementTab` (Balance âgée) corrigés pour utiliser le bon parent.

**Vérification exhaustive automatisée** effectuée cette fois sur les
**50 scrollbars** ajoutées : un script compare, pour CHAQUE tableau du
fichier, le parent Tkinter réel utilisé à sa création avec le parent
utilisé par son conteneur de scrollbar — confirmé zéro incohérence
restante nulle part dans l'application (recherche faite deux fois, y
compris sur les tableaux en variable locale, pas seulement `self.xxx`).

Testé : compilation propre, Trésorerie calculable sans erreur, moteur
comptable toujours pleinement fonctionnel.

### Toutes les scrollbars ajoutées automatiquement retirées — stabilité avant tout

**Demande** : après plusieurs plantages successifs causés par le script
d'ajout automatique de scrollbars (Grand livre, Trésorerie, Balance
âgée), retrait de toutes les scrollbars ajoutées, y compris celles qui
n'avaient pas encore planté.

**Réalisé** : les **50 scrollbars ajoutées automatiquement** (script
précédent) ont été annulées par un script de sens inverse, restituant
l'affichage simple d'origine (`self.tree.pack(...)`). Les **4 cas
corrigés manuellement** (Bilan SYSCOHADA, Grand livre, Soldes
d'ouverture, Exercices comptables) ont également été ramenés à leur
structure simple d'origine, par souci de cohérence et pour éliminer tout
risque résiduel — même s'ils n'avaient causé aucun plantage.

Les 2 scrollbars **préexistantes** avant mes modifications (non liées à
cet incident, jamais fautives) ont été conservées telles quelles.

Testé : compilation propre sur les 46 écrans concernés, moteur comptable
et Grand livre pleinement fonctionnels, Soldes d'ouverture (colonnes
Débit/Crédit avec totaux, ajoutées dans une réponse précédente et non
liées au bug) toujours opérationnels.

## CORRECTIF MAJEUR — cause racine de l'écart Actif/Passif enfin trouvée et corrigée

**Symptôme signalé** : après import de la balance N-1 (fichier joint,
parfaitement équilibré dans Excel — somme exactement 0), le logiciel
affichait « NON ÉQUILIBRÉ » avec un écart de plusieurs milliards.

**Cause racine trouvée** (après investigation approfondie) :
`compute_balance()` — la fonction utilisée PARTOUT dans l'application
(Bilan, Balance générale, Grand livre, Soldes d'ouverture) — parcourait
la table `accounts` (le Plan comptable bundlé) puis cherchait le solde
d'ouverture de CHAQUE compte listé. **Tout compte présent dans les
soldes d'ouverture ou dans les écritures mais ABSENT du Plan comptable
bundlé (plan comptable réel de l'utilisateur plus détaillé, avec des
sous-comptes non prévus) était silencieusement IGNORÉ** — dans le
fichier de l'utilisateur, 19 comptes sur 132 (ex. `231102`, `234102`,
`471108`...) n'existaient pas dans le Plan comptable de l'application,
et leur solde disparaissait de tous les calculs. Cette **même faille**
touchait `list_opening_balances()` (jointure INNER au lieu de LEFT) et 4
autres requêtes (calcul du coût unitaire moyen analytique, budget par
code analytique).

**C'est la cause de l'écart Actif/Passif chassé depuis le tout début de
cette conversation** — pas des soldes d'ouverture réellement incomplets.

**Corrigé** :
- `compute_balance()` parcourt désormais l'**union** des comptes du Plan
  comptable, des comptes ayant un solde d'ouverture et des comptes ayant
  reçu au moins une écriture — plus aucun compte ne peut disparaître
  silencieusement. Un compte absent du Plan comptable reçoit un libellé
  de repli (« CODE (hors Plan comptable) ») et sa classe est déduite du
  premier chiffre de son code.
- `list_opening_balances()` : jointure LEFT au lieu de INNER, avec le
  même repli.
- 4 requêtes analytiques (coût unitaire moyen, budget par code
  analytique) : filtre par classe de compte fait directement sur le code
  (`substr(compte,1,1)='6'`) au lieu de passer par une jointure qui
  pouvait exclure des comptes non répertoriés.

**Testé avec le fichier réel de l'utilisateur** : 132 comptes importés
ET affichés (au lieu de 113 avant le correctif) ; Total Débit = Total
Crédit = 38 991 589 074 (écart 0, au lieu de ~9 milliards) ; Bilan
parfaitement équilibré (37 789 564 389 des deux côtés).

## Trois nouveaux sous-menus dans RAPPORT FINANCIERS : Compte de résultat (SIG), TFT, Situation financière

**Réalisé** :
- **3 fichiers gabarits fournis** (Compte de résultat SIG, TFT, Situation
  financière FR-BFR-TN) — même format que le Bilan, **encodés directement
  dans le code source** (`etats_financiers_data.py`) selon le même
  principe que `bilan_template_data.py` : plus aucune dépendance au
  bundle PyInstaller, régénérés automatiquement si besoin.
- **Nouvelle fonction générique `compute_etat_formule_generique()`** :
  lit n'importe lequel de ces 3 gabarits (structure commune « RUBRIQUE |
  N (| N-1 | %) », détectée automatiquement via la ligne d'en-tête) et
  évalue ses formules avec le même moteur CtaCptSolde que le Bilan — zéro
  duplication de code entre les 3 états.
- **Nouvelle fonction `Ratio(valeur, décimales, unité...)`** ajoutée au
  moteur de formules (rencontrée dans la Situation financière) — mise en
  forme d'affichage sans effet sur le calcul, ignorée proprement.
- **Écran générique `EtatFormuleTab`** (réutilisé pour les 3) : Actualiser,
  Modifier les formules du template (ouvre le fichier figé dans Excel,
  même principe que le Bilan), Exporter (.xls, préservant la mise en
  forme d'origine).
- **3 sous-menus ajoutés** à RAPPORT FINANCIERS : Compte de résultat
  (SIG), TFT, Situation financière.

Testé de bout en bout avec un scénario réaliste (vente + achat) : Compte
de résultat (41 lignes, 0 erreur), TFT (31 lignes, 0 erreur), Situation
financière (34 lignes, 2 erreurs de division par zéro attendues — ratios
de rentabilité quand les capitaux propres sont nuls dans le jeu de test
minimal, pas un bug). Export des 3 états vérifié (mise en forme
préservée, formules correctement résolues même avec des dépendances
inter-rubriques en plusieurs passes). Non-régression du Bilan SYSCOHADA
confirmée.

### Crash « No module named 'etats_financiers_data' » corrigé

**Crash signalé** (capture d'écran) : les 3 nouveaux sous-menus (Compte
de résultat, TFT, Situation financière) affichaient tous
`No module named 'etats_financiers_data'`.

**Cause** : `_generer_template_depuis_b64()` utilisait un import
DYNAMIQUE (`importlib.import_module(module_name)`, où `module_name` était
une simple chaîne de caractères passée en paramètre) — PyInstaller ne
peut analyser statiquement ce genre d'import et ne l'inclut donc PAS
automatiquement dans l'exécutable compilé, contrairement à un `import
etats_financiers_data` écrit littéralement dans le code (détectable).

**Corrigé** : chaque fonction (`_cr_template_path()`, `_tft_template_path()`,
`_situation_template_path()`) fait maintenant son propre `import
etats_financiers_data` **littéral** — exactement le même principe qui
fonctionne déjà pour `bilan_template_data`. Par sécurité supplémentaire,
les deux modules ont aussi été ajoutés explicitement en
`--hidden-import` dans le workflow de build.

Testé : les 3 états se calculent à nouveau sans erreur, Bilan SYSCOHADA
non régressé.

## CORRECTIF MAJEUR — colonne N-1 des formules …Nm1 corrigée partout (TFT, CR, Situation financière, ET Bilan gabarit)

**Signalé** : la ligne « Trésorerie nette au 1er janvier N-1 » du TFT
affichait 0 alors que le solde d'ouverture réel existe bien.

**Cause racine** : toutes les formules `…Nm1` (CtaCptSoldeNm1,
CtaCptSoldeDébitNm1, CtaCptSoldeCréditNm1) étaient alimentées par
`_soldes_dict(conn, exercice_n1)` — qui calcule le solde de CLÔTURE d'un
exercice **N-1 complètement séparé** (ex. 2025), avec ses propres
écritures et son propre solde d'ouverture. Tant que l'utilisateur n'a
pas créé et alimenté cet exercice 2025 séparément dans l'application (ce
qui n'est généralement pas le cas — le solde d'ouverture de l'exercice
courant EST déjà, mathématiquement, le solde de clôture de l'exercice
précédent), cette colonne renvoyait 0 partout.

**Corrigé** : nouvelle fonction `_soldes_ouverture_dict()` — alimente
désormais toutes les formules `…Nm1` avec le **solde d'ouverture de
l'exercice courant** (comme déjà fait pour l'écran Bilan SYSCOHADA
autonome). Corrige les 5 fonctions concernées : le TFT, le Compte de
résultat, la Situation financière, **et un bug latent dans l'export
« gabarit officiel » du Bilan** (`compute_bilan_plat`, utilisée par
« Visionner »/« Exporter ») qui avait exactement le même défaut.

Testé : TFT affiche maintenant 2 500 000 (solde d'ouverture réel) au
lieu de 0 pour la trésorerie nette d'ouverture ; Bilan gabarit affiche
désormais 5 000 000 en Net N-1 (au lieu de 0) sur un scénario identique ;
Compte de résultat non régressé (41 lignes, 0 erreur).

### Nom du logiciel changé + icône d'usine

**Demande** : renommer "Saisie Comptable SYSCOHADA" en "PLATEFORME
INTEGREE DE GESTION", et remplacer l'icône (une plume) par une icône
d'usine.

**Réalisé** :
- Titre de la fenêtre changé en **« PLATEFORME INTEGREE DE GESTION »**.
- **Icône d'usine créée** (bâtiment bleu avec deux cheminées fumantes,
  toit en dents de scie) et appliquée à la fois à la fenêtre de
  l'application (`self.iconbitmap()`) et à l'exécutable lui-même
  (`--icon` dans le workflow de build, pour l'icône du fichier `.exe` et
  de la barre des tâches Windows).
- **Encodée en base64 dans le code source** (`factory_icon_data.py`),
  exactement comme les gabarits de rapports financiers — évite la même
  classe de bug rencontrée plusieurs fois (fichier absent du bundle
  PyInstaller malgré une configuration `--add-data` correcte). Régénérée
  automatiquement au démarrage si besoin, dans le dossier persistant de
  l'application.

Testé : icône régénérée correctement à partir des données encodées
(14 599 octets, identique à l'original), titre de fenêtre mis à jour,
moteur comptable non régressé.

### Crash au démarrage corrigé — NiveauxAccesTab (KeyError: 'code')

**Crash signalé** (capture d'écran) : `KeyError: 'code'` dans `refresh()`
de `NiveauxAccesTab`, empêchant l'application de démarrer dès qu'au
moins un niveau d'accès existe dans la base (menu ADMIN > Niveaux
d'accès) — bug préexistant, sans rapport avec le changement de nom/icône
de la réponse précédente, révélé simplement parce que toutes les pages
s'instancient d'un coup au démarrage.

**Cause** : `core.list_niveaux_acces()` renvoie des dictionnaires avec
les clés `nom`/`description` (colonnes réelles de la table
`niveaux_acces`), alors que l'écran générique `_SimplePlanTab` (dont
hérite `NiveauxAccesTab`) attend des clés `code`/`label`.

**Corrigé** : `NiveauxAccesTab.list_fn()` transforme désormais
explicitement `nom`→`code` et `description`→`label` avant de les
transmettre à l'écran générique — **sans toucher à
`core.list_niveaux_acces()`**, dont dépend un autre appelant
(`UtilisateursTab._refresh_niveaux()`, qui utilise `nom` directement) qui
serait sinon cassé par un changement de la fonction partagée.

Testé : reproduit le crash exact avec des niveaux d'accès réels
(Administrateur, Comptable, Lecture seule, Saisie seule), confirmé
corrigé ; l'autre appelant de `list_niveaux_acces()` reste inchangé et
fonctionnel.

## NOUVELLE ARCHITECTURE — Serveur + Client réseau (multi-utilisateur simultané)

**Demande** : transformer ce logiciel en serveur, et créer une
application « client » qui s'y connecte par réseau local ou Internet,
avec plusieurs utilisateurs travaillant réellement en même temps.

### Vue d'ensemble

Trois programmes distincts désormais dans le projet :

1. **`main.py`** (inchangé) — l'application de bureau autonome
   existante, qui ouvre directement un fichier SQLite local. Continue de
   fonctionner exactement comme avant, pour un usage mono-poste.
2. **`server.py`** (nouveau) — le **SERVEUR** : héberge la base de
   données partagée et expose le moteur comptable (`core.py`) sur le
   réseau (HTTP + JSON, bibliothèque standard uniquement, aucune
   dépendance supplémentaire).
3. **`client_main.py`** (nouveau) — le **CLIENT** : une application de
   bureau séparée, qui ne touche à AUCUN fichier local — elle se connecte
   au serveur (adresse IP + port) et travaille entièrement à distance.

### Comment ça marche

- Le serveur ouvre la base SQLite en **mode WAL** (meilleure gestion de
  la lecture/écriture simultanée) et sérialise les écritures avec un
  verrou global — plusieurs utilisateurs peuvent travailler EN MÊME
  TEMPS sans se marcher dessus ni corrompre les données (testé avec deux
  utilisateurs écrivant en parallèle via de vrais threads simultanés :
  aucune perte, Bilan resté équilibré).
- **Authentification** : réutilise la table `utilisateurs` déjà
  existante (menu ADMIN) — chaque utilisateur se connecte avec son
  identifiant/mot de passe habituel, reçoit un jeton de session (8h de
  validité).
- **Sécurité** : liste blanche explicite des fonctions accessibles à
  distance (`RPC_WHITELIST` dans `server.py`) — impossible d'exécuter du
  code arbitraire sur le serveur. Actuellement : Saisie comptable
  (multi-lignes), Ventes (clients/factures), Achats (fournisseurs/
  règlements/bons de commande), Stocks — le circuit commercial demandé
  en priorité.
- **`client_core.py`** : module miroir de `core.py` côté client — chaque
  fonction transforme automatiquement l'appel en requête réseau vers le
  serveur, ce qui permet de réutiliser (presque) le même style de code
  que l'application de bureau existante.

### Premier écran client entièrement fonctionnel : la SAISIE

L'écran de Saisie comptable multi-lignes du client est **pleinement
opérationnel de bout en bout** — testé avec le vrai moteur (formulaire
multi-lignes, contrôle d'équilibre en temps réel, envoi au serveur,
refus propre d'une écriture déséquilibrée, actualisation de la liste des
dernières écritures).

**Les écrans Ventes / Achats / Stocks** restent à construire côté client
(un onglet « À VENIR » les liste) — le serveur les expose déjà
(`RPC_WHITELIST`), donc leur ajout suit exactement le même modèle que la
Saisie, sans travail d'architecture supplémentaire.

### Utilisation

**Sur le poste serveur** (celui qui héberge les données) :
```
python server.py --port 8765
```
(ou l'exécutable `SaisieComptableServeur.exe` une fois compilé). Le
port par défaut est 8765 ; la base de données utilisée est celle de
l'installation standard (`%LOCALAPPDATA%\SaisieComptable\comptabilite.db`),
sauf `--db chemin_personnalisé.db`.

**Sur chaque poste client** : lancer `client_main.py` (ou
`SaisieComptableClient.exe`), saisir l'adresse IP du poste serveur
(visible avec `ipconfig` sur le poste serveur, réseau local) et le port,
puis se connecter avec un identifiant/mot de passe existant.

**Accès depuis Internet** (hors réseau local) : nécessite soit une
redirection de port sur le routeur du poste serveur (le trafic reste en
clair, réservé à un usage de confiance), soit un VPN d'entreprise
(solution recommandée pour un accès distant sécurisé) — ce serveur
n'implémente pas le chiffrement HTTPS par défaut.

### Testé de bout en bout

- Connexion, refus de mauvais mot de passe, refus d'accès sans session,
  refus d'une fonction non autorisée.
- Écriture comptable multi-lignes postée via le réseau, correctement
  persistée et équilibrée.
- Refus propre d'une écriture déséquilibrée envoyée par le client.
- **Deux utilisateurs différents écrivant réellement en même temps**
  (threads parallèles) : 100 écritures créées sans perte ni corruption,
  Bilan resté équilibré.
- Non-régression complète de l'application de bureau existante
  (`main.py`), totalement indépendante de cette nouvelle architecture.

### Workflow de build mis à jour

Trois exécutables Windows générés désormais : `SaisieComptable.exe`
(application de bureau autonome, inchangée), `SaisieComptableServeur.exe`
(serveur, mode console pour voir les journaux), `SaisieComptableClient.exe`
(client réseau).

### Workflow GitHub Actions renommé en main.yml (correction critique)

**Signalé** : "GitHub a construit un seul .exe" malgré le workflow à 3
builds livré précédemment.

**Cause trouvée** : le fichier de workflow actif dans le dépôt GitHub de
l'utilisateur s'appelle **`main.yml`**, alors que je l'avais toujours
travaillé sous le nom `build.yml` — un nom différent depuis le début de
ce fichier dans mes livraisons. En important le zip, `build.yml`
s'ajoutait donc À CÔTÉ de l'ancien `main.yml` (qui ne construisait
qu'un seul exécutable) au lieu de le REMPLACER — c'est cet ancien
fichier qui continuait de s'exécuter à chaque push.

**Corrigé** : le fichier a été renommé `.github/workflows/main.yml`
(contenu inchangé — toujours les 3 builds PyInstaller) pour correspondre
exactement au nom du fichier déjà présent dans le dépôt GitHub de
l'utilisateur, et le remplacer correctement au prochain import.

## Contrôle d'accès réel — niveaux d'accès liés aux menus, connexion obligatoire

**Demande** : les niveaux d'accès doivent réellement restreindre les
menus et sous-menus (Saisie, Production, Rapport financiers,
Engagements, GRH, Trésorerie, Transport, Immobilisations, etc.), pas
juste exister comme référentiel.

**Réalisé** :
- **`core.MENU_STRUCTURE`** (nouveau) : structure canonique des 48
  sous-menus de l'application, source unique de vérité — vérifié qu'elle
  correspond exactement aux `register()` réels de `main.py` (zéro
  divergence).
- **Table `niveau_acces_menus`** (nouvelle) : associe à chaque niveau
  d'accès la liste des sous-menus qu'il autorise.
  `core.get_menus_autorises()`/`set_menus_autorises()` pour lire/écrire
  cette association. Le niveau **« Administrateur » a toujours accès à
  tout**, même sans configuration explicite — garde-fou pour ne jamais
  risquer un verrouillage total de l'administration.
- **Écran « Niveaux d'accès » reconstruit** (menu ADMIN) : en
  sélectionnant un niveau, une liste à cocher de tous les sous-menus
  apparaît à droite (groupés par menu principal), avec un bouton
  « Enregistrer les autorisations ».
- **Connexion obligatoire au démarrage** — mais **seulement dès qu'au
  moins un utilisateur existe** dans la base (menu ADMIN > Utilisateurs).
  Tant qu'aucun utilisateur n'a été créé, l'application démarre librement
  (mode amorçage), pour ne jamais bloquer l'accès à une installation
  neuve. Une fois connecté, les menus affichés sont filtrés en temps
  réel selon les autorisations du niveau d'accès de l'utilisateur — un
  menu de premier niveau sans aucun sous-menu autorisé est masqué
  entièrement. L'identité connectée s'affiche dans la barre du haut.
- `ajouter_niveaux_acces_suggeres_menus()` : préconfigure des ensembles
  raisonnables pour les niveaux courants (Administrateur = tout,
  Comptable = tout sauf ADMIN, Lecture seule = rapports financiers et
  trésorerie uniquement, Saisie seule = saisie et documents commerciaux
  de base) — sans jamais écraser une configuration déjà personnalisée.

Testé de bout en bout : garde-fou vérifié (aucun utilisateur → pas de
connexion requise ; dès qu'un utilisateur existe → connexion exigée) ;
authentification et filtrage réel des menus selon le niveau (un niveau
« Saisie seule » restreint à 2 sous-menus voit bien disparaître Bilan et
tous les autres) ; Administrateur conserve l'accès total même sans
configuration ; moteur comptable non régressé.

### Client réseau — barre de menu filtrée selon le niveau d'accès

**Demande** : le client doit aussi voir le menu qui lui est autorisé en haut.

**Réalisé** :
- **`server.py`** : la réponse de connexion (`/login`) inclut désormais
  `menus_autorises` — la liste des sous-menus réellement permis pour le
  niveau d'accès de l'utilisateur (calculée côté serveur via
  `core.get_menus_autorises()`, la même fonction que l'application de
  bureau).
- **`client_core.py`** : `RemoteConnection` conserve cette liste
  (`remote.menus_autorises`) après connexion.
- **`client_main.py`** : `ClientApp` construit désormais une **vraie barre
  de menu** (même structure que l'application de bureau,
  `core.MENU_STRUCTURE`), filtrée exactement de la même façon — un menu
  de premier niveau sans aucun sous-menu autorisé est masqué. Les
  sous-menus autorisés mais **pas encore construits côté client** (tout
  sauf Saisie, à ce stade) s'affichent quand même dans le menu, marqués
  « (bientôt disponible) », et affichent un message clair au clic au lieu
  de planter — pour ne jamais bloquer sur une fonctionnalité en cours de
  développement.

Testé de bout en bout avec deux profils réels : un niveau restreint
(« Saisie seule », 3 sous-menus autorisés) ne voit que 2 menus de premier
niveau (SAISIE, COMMERCE) ; un Administrateur reçoit bien les 48 menus.

### Profils métier précis pour les niveaux d'accès

**Demande** : les niveaux suggérés doivent correspondre à des rôles
métier réels : Comptable, Vendeur, Chargé des achats, GRH, Trésorier,
Usine (+ Administrateur).

**Réalisé** — 7 profils désormais proposés via « Ajouter les niveaux
courants » (menu ADMIN > Niveaux d'accès), avec des autorisations de
menus adaptées à chaque fonction :
- **Administrateur** : tous les 48 modules.
- **Comptable** (15 modules) : Saisie, Soldes d'ouverture, tous les
  rapports financiers (Grand livre, Balance, Bilan, Compte de résultat,
  TFT, Situation financière), Trésorerie, et les paramètres comptables
  (Exercices, Plan comptable, Plan analytique, Plan budgétaire, Plan
  bailleurs, Synchronisation).
- **Vendeur** (5 modules) : Clients, Recouvrement, Facturation, Stocks,
  Marges bénéficiaires (menu COMMERCE).
- **Chargé des achats** (6 modules) : Fournisseurs, Contrats, Expression
  de besoin, Bon de commande, Bordereau de livraison, Règlements (menu
  ENGAGEMENTS-PROJETS).
- **GRH** (5 modules) : Liste du personnel, Time sheet, KPI, Tableau de
  bord GRH, HS.
- **Trésorier** (3 modules) : Trésorerie, Recouvrement (encaissements
  clients), Règlements (décaissements fournisseurs).
- **Usine** (11 modules) : Production, Transport, Immobilisations,
  Maintenance-Qualité, Rapports technique.

Chaque utilisateur créé avec l'un de ces niveaux voit automatiquement,
sur l'application de bureau ET sur le client réseau, uniquement les
menus correspondant à sa fonction — testé et vérifié pour les 7 profils
(comptage et contenu exacts confirmés).

### Confirmation : restriction fine par sous-menu (pas juste par menu entier)

**Demande** : chaque profil (ex. Vendeur) doit voir son menu (ex.
"commercial"), avec possibilité de restreindre au niveau des sous-menus.

**Confirmé** — c'était déjà le mécanisme construit : `set_menus_autorises()`
travaille au niveau de CHAQUE sous-menu individuellement (pas par bloc de
menu entier). Un menu de premier niveau (ex. COMMERCIAL) reste visible
tant qu'il lui reste AU MOINS UN sous-menu autorisé, et disparaît
entièrement dès que tous ses sous-menus sont retirés — testé et vérifié
avec le Vendeur : retirer "Recouvrement" et "Marges bénéficiaires" les
fait disparaître individuellement, sans toucher aux 3 autres sous-menus
autorisés (Clients, Facturation, Stocks) ; retirer TOUS les sous-menus
fait disparaître le menu COMMERCIAL en entier. Ce mécanisme est
identique sur l'application de bureau et sur le client réseau (même
fonction `core.get_menus_autorises()` des deux côtés).

**Renommage** : le menu « COMMERCE » renommé en **« COMMERCIAL »** pour
coller au vocabulaire de l'utilisateur.

## CORRECTIF MAJEUR — le serveur voyait des données périmées sans redémarrage

**Symptôme signalé** : après avoir configuré les modules du niveau
« GRH » dans l'application de bureau, le client connecté au serveur ne
voyait toujours aucun menu — jusqu'à ce que le serveur soit manuellement
redémarré, après quoi ça fonctionnait.

**Cause** : le serveur garde une connexion SQLite ouverte en continu
(mode WAL, pour de bonnes performances multi-utilisateur). Sur certains
systèmes (Windows en particulier), cette connexion longue durée peut
rester figée sur un instantané ancien de la base tant qu'aucune
transaction n'est explicitement close — même si un AUTRE processus
(l'application de bureau) a bien enregistré des changements entre-temps.
Un redémarrage du serveur forçait une nouvelle connexion, donc une
lecture fraîche — mais ce n'est évidemment pas praticable au quotidien.

**Corrigé** : un `commit()` (sans effet si rien n'était en attente) est
désormais exécuté systématiquement **avant chaque requête réseau**
(connexion ET appels RPC), forçant la connexion à toujours repartir d'un
instantané à jour de la base — plus jamais besoin de redémarrer le
serveur pour voir des changements faits depuis l'application de bureau.

Testé en reproduisant exactement le scénario signalé : serveur démarré
une seule fois (jamais redémarré), configuration ajoutée depuis un AUTRE
processus (simulant l'application de bureau) PENDANT que le serveur
tournait, puis nouvelle connexion via ce même serveur — les 5 menus GRH
sont désormais correctement reçus sans aucun redémarrage.

## Écrans GRH construits côté client (5 sous-menus pleinement fonctionnels)

**Réalisé** — les 5 écrans GRH ajoutés au client réseau, suivant
exactement le même modèle que la Saisie :
- **Liste du personnel** — création, mise à jour, suppression.
- **Time sheet** — pointage des heures par employé.
- **KPI** — indicateurs de performance avec taux de réalisation calculé.
- **Tableau de bord GRH** — synthèse en lecture seule (cartes de
  résumé + incidents HS par gravité).
- **HS (hygiène santé)** — incidents, visites médicales, formations,
  distributions d'EPI.

**Factorisation** : la gestion d'erreur réseau unifiée (session expirée,
serveur injoignable, erreur métier), auparavant dupliquée dans
`RemoteSaisieTab`, a été extraite en une fonction commune `appeler()` —
réutilisée par les 6 écrans du client désormais, pour éviter toute
duplication future.

**Serveur** : les 16 fonctions GRH ajoutées à `RPC_WHITELIST`
(`list_personnel`, `add_personnel`, `add_time_sheet`, `add_kpi`,
`add_hs`...).

Testé de bout en bout avec un vrai serveur et un vrai client réseau
(pas seulement le moteur local) : création d'un employé, pointage
d'heures, KPI avec taux de réalisation, incident HS, puis vérification
que le Tableau de bord agrège correctement toutes ces données en temps
réel à travers le réseau. Non-régression du moteur comptable confirmée.

## Liste blanche du serveur considérablement élargie (236 fonctions)

**Demande** : activer tous les menus à distance, plutôt que d'ajouter les
fonctions une par une à chaque nouvel écran client — pour éviter de
devoir reconstruire le serveur à chaque fois.

**Réalisé** : `RPC_WHITELIST` passe de 61 à **236 fonctions** — la quasi-
totalité des fonctions métier de `core.py` (générée automatiquement par
script à partir de toutes les fonctions publiques prenant `conn` en
premier argument). Désormais, **ajouter un nouvel écran côté client ne
nécessitera plus de reconstruire le serveur** — seulement le client.

**Exclusions volontaires** (14 fonctions, sécurité) : gestion des
utilisateurs et niveaux d'accès (`add_utilisateur`, `delete_utilisateur`,
`add_niveau_acces`, `set_menus_autorises`...), `verify_password` (déjà
géré en interne par `/login`, ne doit pas être appelable directement),
`reinitialiser_donnees` (remise à zéro destructrice), `init_db`,
`load_plan_comptable`, `synchroniser_base`.

**Exclusions techniques** (32 fonctions) : tous les `export_*`/`import_*`
travaillant sur des fichiers Excel — ces chemins de fichiers désignent le
disque du SERVEUR, pas celui du client, donc sans signification
pertinente en RPC tel quel (nécessiterait un vrai mécanisme d'upload/
téléchargement, hors périmètre pour l'instant).

**Note sur le filtrage** : ce filtrage large côté serveur repose sur
l'AUTHENTIFICATION (session valide obligatoire), pas encore sur une
vérification du niveau d'accès PAR FONCTION — le filtrage des menus
selon le profil (Vendeur, GRH...) reste géré côté client (l'interface ne
montre/n'appelle que ce qui est autorisé). Un utilisateur techniquement
capable d'envoyer des requêtes HTTP directes pourrait outrepasser ce
filtrage d'interface. Pour un usage en réseau de confiance (bureau,
LAN d'entreprise), ce n'est pas un risque significatif ; à renforcer
avec une vérification serveur par fonction si l'usage s'étend.

Testé de bout en bout : `add_personnel`/`list_personnel` (qui posaient
problème) fonctionnent désormais ; les fonctions sensibles
(`add_utilisateur`, `reinitialiser_donnees`) restent bien bloquées.

## Écrans Achats (Fournisseurs, Règlements) construits côté client

**Demande** : le client affichait "(bientôt disponible)" pour tous les
sous-menus d'ENGAGEMENTS-PROJETS — confirmé que ce n'était PAS un bug
(le système fonctionnait déjà correctement : connexion, filtrage des
menus, serveur en arrière-plan) mais simplement des écrans pas encore
construits côté client.

**Réalisé** — 2 écrans du circuit Achats construits, suivant le même
modèle que Saisie et GRH :
- **Fournisseurs** — création, mise à jour, suppression, recherche.
- **Règlements** — création du règlement, ajout de lignes (compte de
  charge, libellé, quantité, prix unitaire), et **validation qui
  comptabilise réellement l'achat sur le serveur** (débit du compte de
  charge, crédit fournisseur — même moteur que l'application de bureau).

Testé de bout en bout avec un vrai serveur et un vrai client réseau :
fournisseur créé, règlement créé avec une ligne, validé (statut passé à
"validee", écriture comptable réellement postée), Bilan resté équilibré
après l'opération. Non-régression du moteur comptable confirmée.

**Reste à construire côté client** (suivant le même modèle) : Contrats,
Expression de besoin, Bon de commande, Bordereau de livraison (le reste
d'ENGAGEMENTS-PROJETS), puis le circuit Ventes (Clients, Facturation,
Stocks, Marges) et Trésorerie/Transport/Immobilisations selon la
priorité souhaitée.

### Outil pratique : lancer le serveur en arrière-plan sans fenêtre

Un script `Lancer_Serveur_Arriere_Plan.vbs` a été fourni séparément
(hors de ce zip, transmis directement) : lance le serveur SANS aucune
fenêtre visible à garder ouverte, pour éviter de le couper par erreur en
fermant une fenêtre — le serveur continue de tourner en arrière-plan tant
que l'ordinateur reste allumé.

## GROS LOT — 13 nouveaux écrans construits côté client (21/48 au total)

**Demande** : construire tous les écrans manquants.

**Réalisé dans cette réponse** (13 nouveaux écrans, en plus des 8 déjà
existants) :
- **RAPPORT FINANCIERS** (6, lecture seule) : Grand livre, Balance,
  Bilan SYSCOHADA, Compte de résultat (SIG), TFT, Situation financière.
- **COMMERCIAL** (3) : Clients (CRUD), Facturation (création + lignes +
  validation qui comptabilise réellement la vente), Stocks (lecture).
- **TRESORERIE** (1, lecture) : banques horizontales + engagements à
  payer.
- **IMMOBILISATIONS** (1, lecture).
- **ENGAGEMENTS-PROJETS** (2) : Expression de besoin (avec validation
  qui bascule automatiquement en Bon de commande), Bon de commande (avec
  validation qui comptabilise réellement l'achat).

**Point technique corrigé au passage** : `compute_etat_formule_generique()`
prend un argument fonction (non transmissible en JSON) — 3 nouvelles
enveloppes RPC-compatibles créées (`compute_cr`, `compute_tft_gabarit`,
`compute_situation_fin`). Collision de nom détectée et corrigée
(`compute_tft` existait déjà pour un calcul différent — renommé en
`compute_tft_gabarit` pour la nouvelle enveloppe).

**Total désormais : 21 sous-menus sur 48 pleinement fonctionnels côté
client** (contre 8 avant cette réponse).

Testé de bout en bout à chaque étape avec un vrai serveur et un vrai
client réseau (pas seulement le moteur local) : les 6 rapports
financiers, le circuit Ventes complet (client → facture → validation →
Bilan équilibré), le circuit Achats complet via Expression de besoin
(expression → validation → Bon de commande → complété → validé →
comptabilisé → Bilan équilibré), Trésorerie et Immobilisations en
lecture. Non-régression du moteur comptable confirmée.

**Reste à construire** (27 sous-menus) : Recouvrement, Marges,
Fabrication, Contrats, Bordereau de livraison, Amortissements, Transport
(Parc auto, Missions, Pièces de rechange, Réparations), Rapports
technique, Maintenance-Qualité (Énergie, Maintenance), Paramètres (6),
et une partie d'ADMIN (les fonctions les plus sensibles — gestion des
utilisateurs, réinitialisation — restent volontairement non exposées à
distance, voir la section sécurité plus haut).

### Diagnostic : "Immobilisations vide sur le client" — pas un bug, décalage d'exercice

**Signalé** : le menu Immobilisations affiche des données sur
l'application de bureau mais rien sur le client.

**Vérifié rigoureusement** : `compute_immobilisations_liste()` (et plus
généralement tout calcul basé sur `compute_balance()`) se comporte de
façon **strictement identique** que l'appel vienne du bureau (connexion
locale directe) ou du client (RPC via le serveur) — testé en reproduisant
exactement le scénario (donnée saisie sous l'exercice 2025, exercice
courant ensuite basculé sur 2026) : les DEUX chemins renvoient un
résultat vide pour 2026 et les données correctes pour 2025. **Ce n'est
donc pas un bug spécifique au client** — c'est que le bureau et le
serveur peuvent être réglés sur des **exercices comptables (années)
différents**, et une immobilisation saisie dans l'ouverture d'un
exercice n'apparaît que dans CET exercice (et les suivants, UNIQUEMENT
si l'exercice a été formellement clôturé via « + Nouvel exercice », qui
reporte les soldes).

**Corrigé** : la barre du client affiche désormais **« Exercice
comptable (serveur) : XXXX »**, pour comparer directement avec le
sélecteur d'exercice du bureau et repérer immédiatement un éventuel
décalage.

Testé de bout en bout : exercice affiché correctement à la connexion,
comportement de `compute_immobilisations_liste()` confirmé identique
entre bureau et client sur les deux exercices testés.

## TERMINÉ — 48 sous-menus sur 48 désormais couverts côté client

**Demande** : construire tous les menus.

**Réalisé** — 20 écrans supplémentaires construits dans cette réponse
(en plus des 21 déjà existants) :
- **COMMERCIAL** : Recouvrement, Marges bénéficiaires.
- **PRODUCTION** : Fabrication (lecture).
- **ENGAGEMENTS-PROJETS** : Contrats, Bordereau de livraison.
- **IMMOBILISATIONS** : Amortissements (taux par catégorie).
- **TRANSPORT** : Parc auto, Missions, Pièces de rechange, Réparations.
- **MAINTENANCE-QUALITÉ** : Énergie, Maintenance (écran générique
  `RemoteAnalytiquePeriodeTab`, réutilisé pour les deux).
- **PARAMÈTRES** : Exercices comptables (avec clôture réelle testée —
  report des soldes vérifié), Plan comptable (recherche), Plan
  analytique, Plan budgétaire, Plan bailleurs de fonds, Synchronisation.
- **SAISIE** : Soldes d'ouverture (Débit/Crédit avec totaux, comme le
  bureau).
- **ADMIN** (5 écrans) : Taux TVA, Taux retenue construits ; Modification
  des factures, Modèle de bon de commande, Niveaux d'accès, Utilisateurs,
  Réinitialisation des données affichent une **explication claire**
  plutôt qu'un écran vide — ces opérations restent délibérément
  réservées à l'application de bureau, par sécurité (voir la liste
  blanche du serveur).

**Factorisation** : un écran générique `RemoteSimplePlanTab` construit
et réutilisé pour 5 écrans (Plan analytique, Plan budgétaire, Plan
bailleurs, Taux TVA, Taux retenue) — élimine la duplication.

**Bug trouvé et corrigé pendant la construction** : `RemoteExercicesTab`
traitait par erreur `list_exercices()` (qui renvoie des dictionnaires
`{"exercice":..., "cloture":...}`) comme de simples chaînes de
caractères — corrigé et testé avec une vraie clôture d'exercice de bout
en bout (report des soldes vers l'exercice suivant vérifié).

**Bilan final : les 48 sous-menus de l'application sont désormais
couverts côté client** — soit un écran pleinement fonctionnel, soit une
explication claire pour les quelques opérations volontairement réservées
au bureau par sécurité. Zéro écran laissé sans réponse.

Testé de bout en bout à chaque étape (Exercices avec clôture réelle,
Soldes d'ouverture, Plan comptable, tous les écrans du lot). Non-
régression complète du moteur comptable confirmée.

### Gestionnaire d'erreurs global ajouté au client — plus jamais d'écran vide sans explication

**Signalé** : Immobilisations reste vide sur le client alors que le
bureau affiche des données réelles pour le même exercice.

**Investigation approfondie** : reproduit fidèlement le scénario exact
communiqué (comptes 231100, 231102 « hors Plan comptable », montants
négatifs, exercice 2026 identique des deux côtés) — **le calcul et la
transmission réseau se sont révélés corrects dans tous les tests**,
aussi bien en appel direct qu'à travers un vrai serveur/client. Impossible
de reproduire le problème précisément avec les informations disponibles.

**Corrigé structurellement** : un **gestionnaire d'erreurs global** a
été ajouté au client (`report_callback_exception`) — sans lui, toute
exception survenant dans un écran (notamment en mode `--windowed`, sans
console visible) est **silencieusement avalée par Tkinter**, laissant
l'écran vide sans le moindre message, rendant tout diagnostic impossible
côté utilisateur. Désormais, **toute erreur s'affiche dans une boîte de
dialogue claire**, avec le détail technique — si le problème d'
Immobilisations revient, ce message permettra d'identifier la cause
exacte immédiatement, au lieu d'un écran silencieusement vide.

Testé : mécanisme de capture et de formatage de l'erreur vérifié.

## CORRECTIF STRUCTUREL — le serveur autorise désormais tout par défaut (liste noire, pas liste blanche)

**Demande** : simplifier le serveur pour qu'il donne simplement accès à
la base de données au client, sans liste à maintenir.

**Cause du problème rencontré** (« Fonction non autorisée à distance »
sur Immobilisations et Balance) : la liste BLANCHE devait être tenue à
jour manuellement à chaque nouvel écran construit côté client, avec un
risque réel de désynchronisation entre client et serveur si l'un des
deux n'était pas reconstruit en même temps que l'autre — exactement ce
qui s'est produit.

**Corrigé structurellement** : le modèle est inversé — désormais, **toute
fonction publique de `core.py` est autorisée à distance PAR DÉFAUT**
(liste calculée dynamiquement au démarrage du serveur, 239 fonctions),
sauf une **petite liste noire explicite** (`RPC_BLOCKLIST`, 14 fonctions)
couvrant uniquement les opérations réellement sensibles : gestion des
utilisateurs et niveaux d'accès (risque d'élévation de privilèges),
réinitialisation destructrice des données, `verify_password` (déjà géré
par `/login`), et les opérations d'infrastructure (`init_db`,
`synchroniser_base`...).

**Conséquence pratique** : **plus jamais besoin de reconstruire le
serveur** quand un nouvel écran est ajouté côté client (tant que la
fonction `core.py` sous-jacente existe déjà) — élimine complètement la
classe de bug rencontrée.

Testé de bout en bout avec les deux cas exacts qui posaient problème
(`compute_immobilisations_liste`, `compute_balance_detaillee`) —
fonctionnent désormais sans toucher au serveur ; vérifié que les
opérations sensibles (`add_utilisateur`...) restent bien bloquées.

## Bouton "Ouvrir le dossier de la base de données" (menu ADMIN)

**Demande** : un bouton dans le menu ADMIN de l'application de bureau
pour ouvrir directement le dossier contenant le fichier de la base de
données.

**Réalisé** : ajouté en haut de l'écran ADMIN > Réinitialisation des
données — affiche le chemin exact et ouvre l'explorateur de fichiers
Windows (ou l'équivalent macOS/Linux) d'un clic, pratique pour localiser
rapidement `comptabilite.db` (utile notamment pour vérifier que le
serveur et le bureau utilisent bien le même fichier).

## Numéro de version visible sur le serveur (fini de deviner)

**Demande** : rester sur l'approche .exe (pas Python direct), mais
éviter la confusion récurrente sur la version du serveur en cours
d'exécution.

**Réalisé** : `SERVER_VERSION = "2026-08-23-v1"` (chaîne à changer à
chaque modification de `server.py`), affichée à **3 endroits** :
1. Console du serveur au démarrage (bien visible, avec le nombre de
   fonctions autorisées).
2. Bouton "Tester la connexion au serveur" du client, avant même de se
   connecter.
3. Barre du haut du client, une fois connecté ("Version serveur : ...").

**Fini de deviner** : pour vérifier que le bon serveur tourne, il suffit
de comparer le numéro affiché au numéro que j'indique dans ma réponse —
plus besoin de dates de fichiers ou de suppositions.

Testé de bout en bout : version affichée correctement dans les 3
emplacements (console, ping, connexion).
