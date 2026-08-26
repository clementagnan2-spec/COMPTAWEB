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
6. **GRH / Paie** ← prochaine étape
7. Fabrication / Production
8. Trésorerie
9. Écrans ADMIN

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
