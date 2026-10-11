# Bibliographie de Christo Datso

Liste des textes publiés de **Christo Datso** (nouvelles, novellas, recueils, chroniques, essais, poésie),
servie sous <https://www.cdatso.be/bibliographie/>.

> **Ce dépôt est un DÉRIVÉ GÉNÉRÉ, jamais édité à la main.** `index.html` est produit par
> `outils/genere-bibliographie.py` à partir d'un classeur de référence qui vit sur le poste de l'auteur
> (hors git, non publié). Toute correction se fait dans le classeur, puis on régénère.

## Ce que la page contient, et ne contient pas

- Contient : titres, types, années, volumes ou revues, éditeurs, liens vers les pages d'éditeurs.
- Ne contient aucun texte intégral ni extrait.
- Hors page : les inédits (usage privé, manuscrits) et les textes d'un autre auteur, tant qu'aucune
  autorisation écrite n'est au dossier.

## Régénérer

```
python outils/genere-bibliographie.py --date "AAAA-MM-JJ HH:MM"
python outils/genere-bibliographie.py --verifier
python outils/genere-bibliographie.py --simuler
```

Python 3.12 et `openpyxl`. Le générateur lit le classeur en lecture seule (option `--classeur` pour un autre
chemin) et n'écrit qu'un fichier, `index.html`. Il refuse (code 2) toute valeur hors des listes fermées de
type de texte et de type de publication, tout ID en double, toute URL qui ne commence pas par `http`.
`--verifier` régénère en mémoire et compare à `index.html` (code 0 si égal). La date du pied de page est
passée en argument, jamais estimée.

## Publication

La publication est un push sur `main` décidé par l'auteur, puis GitHub Pages (`main`, `/`).
Aucune licence n'est attachée à ce dépôt à ce jour.

## Fichiers

| Fichier | Rôle |
|---|---|
| `index.html` | La page, générée. |
| `outils/genere-bibliographie.py` | Le générateur. |
| `404.html` | Page d'erreur minimale. |
| `.nojekyll` | Sert les fichiers tels quels. |
| `.gitattributes` | Fins de ligne LF. |
