# Guide de déploiement — étape 1 (Fondations)

Ce guide explique comment mettre cette première version en ligne. On
utilise **Render.com**, qui fonctionne comme GitHub Actions : vous
poussez du code sur GitHub, et Render construit + déploie automatiquement.

## Ce que contient cette étape

- Page de connexion + création du premier administrateur
- Tableau de bord avec le menu complet, **filtré automatiquement selon
  le niveau d'accès** (même règle que le bureau/client réseau)
- Les 48 sous-menus sont cliquables mais affichent « en construction »
  (le contenu de chaque module sera ajouté un par un dans les prochaines
  étapes — la prochaine sera **Saisie comptable + Rapports financiers**)

## Étape A — Mettre le code sur GitHub

1. Créez un nouveau dépôt GitHub (par exemple `grand-livre-web`), ou
   utilisez un dépôt existant dédié à la version web (à part de celui
   du bureau/exe, pour ne pas mélanger les deux projets).
2. Mettez-y tous les fichiers de ce zip.

## Étape B — Créer le compte Render

1. Allez sur **https://render.com** → « Get Started » → connectez-vous
   avec votre compte GitHub (le même que celui utilisé pour les
   GitHub Actions actuelles).
2. Cliquez **New +** → **Web Service**.
3. Choisissez le dépôt GitHub créé à l'étape A.

## Étape C — Configuration du service

Render détecte Python automatiquement. Renseignez :

| Champ | Valeur |
|---|---|
| Name | `grand-livre` (ou ce que vous voulez) |
| Region | Frankfurt (le plus proche du Burkina Faso) |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `gunicorn app:app` |
| Instance Type | Starter (~7 $/mois) — nécessaire pour le disque persistant |

### Disque persistant (important — sans ça, la base de données serait effacée à chaque redéploiement)

Dans l'onglet **Disks** du service :
- Add Disk → Name: `data` → Mount Path: `/data` → Size: 1 Go (largement
  suffisant au départ, agrandissable plus tard)

### Variables d'environnement

Onglet **Environment** → ajoutez :

| Clé | Valeur |
|---|---|
| `DB_PATH` | `/data/comptabilite.db` |
| `SECRET_KEY` | une longue chaîne aléatoire (Render peut la générer — bouton « Generate ») |

Cliquez **Create Web Service**. Le premier déploiement prend 2-3 minutes.

## Étape D — Premier accès

1. Render vous donne une adresse du type
   `https://grand-livre.onrender.com`. Ouvrez-la.
2. Vous arrivez sur « Première configuration » → créez votre compte
   administrateur.
3. Connectez-vous. Le tableau de bord doit afficher tous les menus.

## Étape E — Nom de domaine (optionnel, pour une adresse personnalisée)

Au lieu de `grand-livre.onrender.com`, vous pouvez utiliser par exemple
`gestion.votre-entreprise.com` :

1. Achetez un nom de domaine chez **OVH** (https://www.ovh.com, simple
   et courant en Afrique francophone) — comptez ~10-15 €/an.
2. Dans Render, onglet **Settings** → **Custom Domains** → ajoutez votre
   domaine. Render vous donne 1-2 valeurs à recopier dans OVH.
3. Chez OVH, zone DNS de votre domaine → ajoutez ces valeurs (Render
   affiche des instructions précises pour OVH directement).
4. Attendez 10 minutes à quelques heures (propagation DNS) — Render
   active alors automatiquement le HTTPS (cadenas) sur votre domaine.

Je peux vous accompagner pas à pas, écran par écran, quand vous serez
rendu à cette étape.

## Redéploiement (à chaque nouvelle version que je vous fournirai)

Contrairement au flux .exe, **pas besoin de télécharger/remplacer de
fichiers** : vous poussez le nouveau code sur GitHub (ou je vous le
fournis prêt à pousser), et Render reconstruit + redéploie tout seul en
1-2 minutes. Aucune interruption pour vos utilisateurs pendant ce temps
n'est garantie sur le plan Starter — on pourra passer à un plan avec
« zero downtime » plus tard si besoin.

## Ce qui reste à construire (dans l'ordre convenu)

1. ~~Fondations (connexion, tableau de bord, menu filtré)~~ ✅ fait
2. ~~Saisie comptable + Rapports financiers~~ ✅ fait
3. ~~Fournisseurs / Clients / Facturation~~ ✅ fait (voir ci-dessous)
4. ~~Stocks~~ ✅ fait (voir ci-dessous)
5. ~~Immobilisations~~ ✅ fait (voir ci-dessous)
6. ~~GRH / Paie (Personnel + moteur de paie)~~ ✅ fait (voir ci-dessous)
7. ~~Paramètres + Admin (fondations transverses)~~ ✅ fait (voir ci-dessous)
8. ~~Fabrication / Production~~ ✅ fait (voir ci-dessous)
9. ~~Trésorerie~~ ✅ fait (voir ci-dessous)
10. ~~ENGAGEMENTS-PROJETS (Contrats, Expression de besoin, Bon de
    commande, Bordereau de livraison, Règlements)~~ ✅ fait (voir
    ci-dessous)
11. ~~COMMERCIAL (Recouvrement, Marges bénéficiaires)~~ ✅ fait (voir
    ci-dessous)
12. ~~ADMIN restants (Modification des factures, Modèle de bon de
    commande, Réinitialisation des données)~~ ✅ fait (voir ci-dessous)
13. ~~GRH restants (Time sheet, KPI, Tableau de bord, HS)~~ ✅ fait (voir
    ci-dessous)
14. ~~TRANSPORT (Parc auto, Missions, Pièces de rechange, Réparations)~~
    ✅ fait (voir ci-dessous)
15. ~~MAINTENANCE-QUALITÉ (Énergie, Maintenance)~~ ✅ fait (voir ci-dessous)
16. RAPPORTS TECHNIQUES — reste volontairement "en construction" : c'était
    déjà un écran non défini dans l'application de bureau elle-même
    ("dites-moi quels rapports vous voulez ici"), donc rien n'a été
    oublié — dites-moi ce que vous voulez y voir et je le construis.
17. Synchronisation — n'a pas de sens en version web (spécifique au
    partage de fichier `.db` en réseau local du bureau/client) ; à
    remplacer si besoin par un export/sauvegarde de la base, dites-moi
    si vous en voulez un.

## 🎉 Toutes les fonctionnalités comptables et opérationnelles sont
## construites et testées (46 sous-menus sur 48 — les 2 restants sont
## des cas particuliers expliqués ci-dessus, pas des oublis).

## Détail de l'étape 7 — Paramètres + Admin (fondations transverses)

- **Exercices comptables** : création, activation, **clôture** (report
  des soldes vers l'exercice suivant, verrouillage en lecture seule,
  alerte si le Bilan n'est pas équilibré) — testé (clôture 2026 → 2027
  créé et activé automatiquement).
- **Plan comptable, Plan analytique, Plan budgétaire, Plan bailleurs** :
  consultation/recherche, ajout, suppression.
- **Taux de TVA** et **Taux de retenue à la source** : listes
  paramétrables (utilisées par Facturation et les futurs écrans Achats).
- **Niveaux d'accès** : création de niveaux personnalisés + sélection
  des sous-menus autorisés par niveau (cases à cocher) — testé avec un
  niveau sur-mesure (2 menus seulement) et vérifié que l'utilisateur
  rattaché ne voit bien que ces 2 menus.
- **Utilisateurs** : création, réinitialisation de mot de passe,
  suppression (protection contre l'auto-suppression) — testé de bout
  en bout, y compris connexion avec le mot de passe réinitialisé.

Ces écrans sont les fondations dont dépendent plusieurs autres modules
(Facturation utilise déjà Taux de TVA, tout utilisateur dépend de
Niveaux d'accès) — les avoir maintenant sécurise la suite.

Sous-menus GRH restants (Time sheet, KPI, Tableau de bord GRH, Heures
sup.) : pas encore construits, moins prioritaires — seront traités dans
un lot ultérieur.

## Détail de l'étape 6 — GRH / Paie

- **Liste du personnel** : ajout, liste, suppression des employés.
- **Paie** (3 onglets) :
  - **Bulletins** : saisie des éléments de gain par employé et par
    période (salaire de base, primes, indemnités, retenue prêt...).
  - **État de paie** : calcul automatique complet — CNSS salariale et
    patronale, TPA patronale, IUTS (barème réel à 9
    tranches avec réduction selon charges de famille), net perçu, coût
    total employeur — puis **Valider la paie** génère toutes les
    écritures comptables (charge salaires, retenues CNSS/IUTS,
    charges patronales) automatiquement, équilibrées, envoyées en
    Saisie sous la pièce `PAIE-AAAA-MM`. La période se verrouille
    ensuite (bulletins non modifiables), comme sur le bureau.
  - **Paramètres** (administrateur uniquement) : taux CNSS, TPA,
    abattements — modifiables sans toucher au code.
- Testé de bout en bout avec un vrai bulletin (200 000 F CFA de salaire
  de base) : calculs vérifiés (CNSS 11 000, IUTS 14 729, net perçu
  172 528) et écritures comptables parfaitement équilibrées (238 000 =
  238 000) en base après validation.

## Détail de l'étape 5 — Immobilisations

- **Immobilisations** : liste de tous les comptes de classe 2 ayant un
  solde (valeur brute, amortissement réellement comptabilisé, valeur
  nette) + fiche éditable par compte (fournisseur, prix d'achat, date
  d'acquisition, base de répartition pour le coût de production en
  Fabrication, amortissement manuel en attendant les vraies dotations).
- **Amortissements** : taux indicatif par catégorie (le montant réellement
  affiché ailleurs reste toujours celui comptabilisé dans la Balance).

## Détail de l'étape 4 — Stocks

- **Synthèse** des 4 comptes centralisateurs (Marchandises, Matières
  premières, Autres approvisionnements, Produits finis) : stock initial,
  entrées, sorties, stock final — en valeur ET en quantité, avec coût
  unitaire moyen pondéré calculé automatiquement.
- **Détail par compte** (tous les sous-comptes de la classe 3 réellement
  utilisés).
- Modification du stock initial (valeur + quantité) par compte et par
  exercice.
- Les entrées/sorties sont calculées **automatiquement** à partir des
  écritures de Saisie (colonne Quantité) et des validations de facture —
  aucune double saisie nécessaire, testé de bout en bout (stock initial
  + écriture d'achat avec quantité → stock final recalculé correctement).

## Détail de l'étape 3 — Fournisseurs / Clients / Facturation

- **Fournisseurs** et **Clients** : liste, recherche, création, suppression.
- **Facturation** : création de facture en brouillon → ajout de lignes
  (compte de vente + quantité + prix, avec recherche de compte en direct,
  calcul HT/TVA/TTC automatique) → **Valider (comptabiliser)** génère les
  écritures comptables équilibrées (Débit Client, Crédit Ventes, Crédit
  TVA) et les envoie directement dans Saisie comptable, exactement comme
  le bureau → **Corriger (dévalider)** repasse la facture en brouillon et
  supprime les écritures générées en cas d'erreur. Aperçu avant impression
  disponible (ouvre le document dans un nouvel onglet).
- Testé de bout en bout : création facture → ligne → validation →
  écritures équilibrées vérifiées en base → apparition correcte dans
  Saisie → dévalidation → écritures bien supprimées.

## Détail de l'étape 8 — Fabrication / Production

- **Fabrication** : création de produits finis (quantité par lot, marge,
  compte de stock) + **recette de fabrication** (nomenclature) par
  produit — chaque ligne (matière première, main-d'œuvre, énergie,
  amortissement d'équipement, autre charge) reprend **automatiquement**
  son coût réel : coût unitaire moyen du stock pour les matières, coût
  d'amortissement réellement comptabilisé pour les équipements, coût
  moyen pondéré du code analytique pour la main-d'œuvre/énergie — sinon
  coût saisi manuellement.
- Calcul automatique du **coût de production total**, coût unitaire, et
  **prix de vente suggéré** selon la marge paramétrée.
- **Produire (comptabiliser)** : diminue le stock de matières (quantité
  + valeur) et augmente le stock de produits finis, écritures générées
  automatiquement.
- Testé de bout en bout avec un cas réel (matière première reprise du
  stock à 1000 F/unité + main-d'œuvre manuelle 2000 F → coût unitaire
  produit 700 F, prix de vente 910 F) : écritures de fabrication
  vérifiées parfaitement équilibrées (14 100 = 14 100) et impact correct
  sur les stocks.

## Détail de l'étape 9 — Trésorerie

- Vue par compte de trésorerie (banques, caisse — classe 5) : solde en
  début de période choisie, entrées, sorties, solde en fin de période.
- **Capacité à faire face aux engagements** : compare la trésorerie
  disponible aux règlements déjà validés mais pas encore décaissés
  (alerte visuelle si le solde après engagements serait négatif).
- Testé : une écriture bancaire (dépôt de capital 500 000 F sur
  521000) apparaît correctement dans le tableau.

## Détail de l'étape 10 — ENGAGEMENTS-PROJETS (circuit d'achat complet)

- **Contrats** : journal des commandes fournisseurs avec échéances de
  livraison et de paiement calculées automatiquement (délais du
  fournisseur), alerte visuelle en cas de retard.
- **Expression de besoin** : création + lignes → validation bascule
  automatiquement en Bon de commande (lignes recopiées), sans écriture
  comptable à ce stade.
- **Bon de commande** : ajout de lignes (compte de charge + code
  analytique par ligne) → **validation** comptabilise directement
  l'achat (Débit charge, Crédit fournisseur, retenue fiscale
  optionnelle) **et génère automatiquement** le Bordereau de livraison
  et le Règlement correspondants (déjà marqués validés, pas de double
  écriture).
- **Bordereau de livraison** : confirmation des quantités réellement
  reçues (aucune écriture — la comptabilisation a déjà eu lieu au Bon
  de commande).
- **Règlements** : peut aussi être créé directement (dépense hors
  circuit d'achat complet) → validation comptabilise la charge →
  **Enregistrer le paiement** comptabilise l'encaissement bancaire/caisse
  et solde la dette fournisseur.
- Testé de bout en bout : Bon de commande (10 000 F) → validation
  (écriture 604000/401000 équilibrée) → Bordereau et Règlement
  auto-créés → paiement du règlement → dette fournisseur (401000)
  ramenée exactement à 0. Circuit Expression de besoin → Bon de
  commande également vérifié (ligne recopiée correctement).

## Détail de l'étape 11 — COMMERCIAL (Recouvrement + Marges)

- **Recouvrement** : factures clients simples (montant global, hors
  détail TVA — différent de Facturation), échéance calculée selon le
  délai de paiement du client, encaissement comptabilisé
  automatiquement (Débit banque/caisse, Crédit client 411000) — testé
  avec écriture équilibrée vérifiée.
- **Marges bénéficiaires** : marge commerciale, valeur ajoutée et
  résultat d'exploitation, calculés directement depuis la liasse
  fiscale (même moteur que le Compte de résultat).

## Détail de l'étape 12 — ADMIN restants

- **Modification des factures** : vue consolidée de toutes les factures
  déjà validées (vente ET achat), avec dévalidation en un clic (retire
  les écritures comptables générées, repasse en brouillon modifiable) —
  testé avec une vraie facture.
- **Modèle de bon de commande** : en-tête/pied de page par défaut,
  appliqué automatiquement aux bons de commande qui n'ont pas leur
  propre en-tête renseigné.
- **Réinitialisation des données** : suppression par catégorie (écritures,
  soldes d'ouverture, fiches immobilisations, circuit d'engagements,
  factures, transport), avec double garde-fou — confirmation JavaScript
  + obligation de taper "SUPPRIMER" en majuscules — testé (2 écritures
  supprimées après confirmation correcte, refusé si mot de confirmation
  incorrect).

## Détail de l'étape 13 — GRH restants

- **Time sheet** : pointage des heures travaillées par employé et par jour.
- **KPI** : indicateurs de performance (cible/réalisé/taux), par employé
  ou par service, avec statut (en cours/atteint/non atteint).
- **HS (hygiène santé)** : journal des incidents/accidents/observations,
  avec gravité et statut (ouvert/fermé).
- **Tableau de bord GRH** : synthèse en cartes agrégeant automatiquement
  les 3 écrans ci-dessus + Personnel — testé (1 employé actif, 8h
  pointées, KPI à 95%, incident créé, tous les chiffres agrégés
  correctement).

## Détail de l'étape 14 — TRANSPORT

- **Parc auto** : liste des véhicules (immatriculation, marque, modèle,
  chauffeur affecté, statut).
- **Missions** : déplacements (destination, chauffeur, dates, kilométrage),
  statut en cours/terminée.
- **Pièces de rechange** : stock partagé entre Transport et Maintenance
  (désignation, quantité, coût unitaire).
- **Réparations** : par véhicule (ou équipement, sans véhicule) — chaque
  pièce utilisée **décrémente automatiquement le stock**, refuse si le
  stock est insuffisant (protection testée : demande de 100 pièces sur un
  stock de 10 → refusée, stock resté intact), coût total = pièces +
  main-d'œuvre calculé automatiquement (testé : 3 filtres × 5000 + 15000
  main-d'œuvre = 30 000 F CFA).

## Détail de l'étape 15 — MAINTENANCE-QUALITÉ (Énergie, Maintenance)

- **Énergie** et **Maintenance** : coûts par code analytique (eau,
  électricité, essence... / véhicules, bâtiments, machines...), sur une
  période librement choisie — alimentés automatiquement par toute
  écriture de Saisie taguée avec le bon code analytique, et par les
  lignes de recette de Fabrication qui leur sont associées. Bouton
  "Ajouter les codes courants" pour préremplir les codes suggérés.
- **Correctif important au passage** : le champ "Code analytique" par
  ligne manquait dans le formulaire web de Saisie des écritures (présent
  côté moteur `core.py` depuis le début, mais pas encore exposé dans
  l'interface HTML) — ajouté et testé de bout en bout (écriture taguée
  ENERGIE-EAU → apparaît correctement dans l'écran Énergie).

## Correctif — Import / Export Excel (oublié lors de la construction initiale)

Ajouté et testé de bout en bout sur les 16 écrans concernés :

| Écran | Export | Modèle | Import |
|---|---|---|---|
| Plan comptable | ✅ | | ✅ |
| Soldes d'ouverture | ✅ | | ✅ |
| Plan analytique | ✅ | | ✅ |
| Plan budgétaire | ✅ | | ✅ |
| Plan bailleurs | ✅ | | ✅ |
| Taux de TVA | ✅ | | ✅ |
| Taux de retenue | ✅ | | ✅ |
| Fournisseurs | | ✅ | ✅ |
| Clients | | ✅ | ✅ |
| Saisie des écritures | | | ✅ (import en masse) |
| Balance | ✅ | | |
| Bilan SYSCOHADA | ✅ (2 formats) | | |
| Immobilisations | | | ✅ |
| Niveaux d'accès | ✅ | | ✅ |
| Liste du personnel | | ✅ | ✅ |
| Time sheet | | ✅ | ✅ |

Testé de bout en bout : export d'un code analytique → suppression en base
→ réimport du même fichier → code bien restauré ; téléchargement du
modèle Fournisseurs → remplissage avec un vrai fournisseur → import →
fournisseur bien créé.

## Correctif — Menu latéral compact + titre de page

- Le menu latéral n'affiche plus que les 13 grands titres (SAISIE,
  COMMERCIAL, PRODUCTION...) — les sous-menus apparaissent maintenant au
  survol de la souris, dans un petit panneau qui se déploie sur la
  droite (comme un menu déroulant), et se replie automatiquement vers le
  haut s'il est trop près du bas de l'écran.
- Ajout du titre **"PLATEFORME INTÉGRÉE DE GESTION"** en haut de chaque
  page, comme sur le bureau.
- Bug trouvé et corrigé pendant les tests : un minuteur partagé entre
  tous les groupes de menu faisait qu'un survol sur un nouveau groupe
  annulait par erreur la fermeture du précédent, laissant deux
  sous-menus ouverts en même temps. Corrigé, testé (un seul sous-menu
  visible à la fois, quel que soit l'ordre de survol) et vérifié que le
  clic sur un lien de sous-menu navigue toujours correctement.
