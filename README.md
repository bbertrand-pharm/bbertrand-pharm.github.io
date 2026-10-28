# Site de Benjamin Bertrand

Site statique (une page `index.html`) dont tout le contenu est dans deux fichiers :

- `data/profil.json` : présentation, contact, expertise, engagements, enseignement
- `data/travaux.json` : articles, communications orales et posters

## Mettre le site en ligne (gratuit, GitHub Pages)

1. Créer un compte sur github.com puis un dépôt public, par exemple `site`.
2. Y déposer le contenu de ce dossier (`index.html`, `data/`, `.pages.yml`).
3. Dans le dépôt : Settings › Pages › Source « Deploy from a branch », branche `main`, dossier `/ (root)`.
4. Le site est en ligne à `https://<votre-identifiant>.github.io/site/` en une à deux minutes.
5. Option : brancher un nom de domaine (ex. `benjaminbertrand.fr`) dans Settings › Pages › Custom domain.

## Mettre à jour le contenu sans coder (Pages CMS)

1. Aller sur https://app.pagescms.org et se connecter avec son compte GitHub.
2. Choisir le dépôt du site.
3. Deux menus apparaissent : « Profil et contact » et « Travaux ». On ajoute, modifie ou supprime une entrée dans un formulaire, puis on clique sur « Save ».
4. Le site se met à jour tout seul en une à deux minutes.

Pour un nouveau poster : Travaux › « Add an entry » › type Poster, année, titre, auteurs, congrès, lien PDF (Google Drive par exemple).
L'ordre de saisie n'a pas d'importance : le site trie par année et propose filtres et recherche.
