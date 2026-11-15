# Clôture de Paris · compte-rendu du soir

Site statique qui publie chaque soir, après la fermeture d'Euronext Paris (17h30),
un compte-rendu de chaque action suivie : cours, variation, commentaire rédigé
automatiquement, volumes, tendances, plus haut/bas annuel, carte de la séance et
synthèse par secteur. Chaque séance est archivée.

## Fichiers

| Fichier | Rôle |
|---|---|
| `generer.ps1` | Récupère les cours (Yahoo Finance, sans clé) et génère le site dans `site/` |
| `actions.json` | Actions suivies, favoris (`favoris`) et cryptomonnaies (`cryptos`) — à modifier librement |
| `serveur.ps1` | Serveur local pour consulter le site sur http://localhost:8080 |
| `.github/workflows/compte-rendu.yml` | Mise en ligne automatique chaque soir (GitHub Pages) |
| `planifier-windows.ps1` | Tâches planifiées Windows : 18h15 (site) et 22h30 (site + e-mail) |
| `configurer-mail.ps1` | À lancer une fois pour recevoir le compte-rendu par e-mail (Gmail) |

## Utilisation locale

```powershell
powershell -ExecutionPolicy Bypass -File generer.ps1
powershell -ExecutionPolicy Bypass -File serveur.ps1
```

## Mise en ligne automatique (gratuit)

1. Créer un dépôt GitHub et y pousser ce dossier.
2. Dans *Settings → Pages*, choisir **Source : GitHub Actions**.
3. Le workflow tourne du lundi au vendredi à 18h15 (heure d'hiver) ; on peut aussi le lancer
   à la main depuis l'onglet *Actions*.

Le site est alors publié sur `https://<utilisateur>.github.io/<dépôt>/`.

## Ajouter ou retirer une action

Éditer `actions.json`. Le symbole est celui de Yahoo Finance (ex. `MC.PA` pour LVMH,
`AAPL` pour Apple). La composition du CAC 40 évolue : vérifiez la liste de temps en temps.

*Les commentaires sont générés automatiquement à partir des cours et ne constituent pas un conseil en investissement.*
