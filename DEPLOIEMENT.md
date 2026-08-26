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
2. ~~Saisie comptable + Rapports financiers~~ ✅ fait (voir ci-dessous)
3. **Fournisseurs / Clients / Facturation** ← prochaine étape
4. Stocks
5. Immobilisations
6. GRH / Paie
7. Fabrication / Production
8. Trésorerie
9. Écrans ADMIN

## Détail de l'étape 2 — Saisie comptable + Rapports financiers

Écrans construits et testés (identiques dans leur logique au bureau,
`core.py` non modifié) :
- **Saisie des écritures** : saisie multi-lignes avec recherche de compte
  en direct, vérification d'équilibre en temps réel, liste des écritures
  de l'exercice, suppression.
- **Soldes d'ouverture** : saisie + liste par exercice.
- **Grand livre** : détail chronologique d'un compte, solde cumulé.
- **Balance** : balance générale de l'exercice.
- **Bilan SYSCOHADA**, **Compte de résultat (SIG)**, **TFT**,
  **Situation financière** : calculés depuis les mêmes gabarits Excel
  officiels que le bureau — si des formules du gabarit TFT posent
  toujours problème, le détail des cellules en erreur s'affiche
  maintenant directement à l'écran.

Un sélecteur d'exercice comptable est apparu en haut de chaque page
concernée (mémorisé pendant la session).
