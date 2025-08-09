## Mail-assist — Complément Outlook (Task Pane + Commandes)

Projet de complément Office pour Outlook, basé sur Webpack 5, Babel et Office.js. Le code applicatif se trouve dans le dossier `My Office Add-in/`.

## Aperçu

- Complément Outlook avec volet des tâches (task pane) et commandes UI-less
- Pile technique: JavaScript, Webpack 5, Babel (`@babel/preset-env`), polyfills (`core-js`, `regenerator-runtime`), Office.js
- Hébergement dev: HTTPS local via serveur webpack, sideload d’Outlook
- Manifeste: XML pointant vers `https://localhost:3001`

## Fonctionnalités

- Volet Outlook (Task Pane):
  - Accueil “Hello world”, affichage du sujet du message sélectionné.
  - Démo d’appel API OpenAI (clé saisie côté client, stockage localStorage à des fins de développement seulement).
- Commande UI-less: action `action(event)` affichant une notification dans l’élément courant.

## Prérequis

- Windows et Outlook Desktop (Microsoft 365)
- Node.js LTS et npm
- Accès administrateur pour approuver les certificats de développement si nécessaire

## Installation

Depuis la racine du dépôt:

```bash
cd "My Office Add-in"
npm install
```

## Démarrage (développement)

Lancement du serveur local + sideload de l’add-in dans Outlook:

```bash
npm run start
```

Notes:

- Le serveur se lance en HTTPS sur le port défini dans `package.json > config.dev_server_port` (3001). Le manifeste référence `https://localhost:3001`.
- Pour arrêter et décharger l’add-in d’Outlook:

```bash
npm run stop
```

## Scripts npm utiles

- `dev-server`: lance uniquement le serveur webpack en mode dev (sans sideload)
- `build:dev`: build en mode développement (source maps)
- `build`: build en mode production
- `watch`: build incrémental en mode dev
- `validate`: valide le manifeste XML
- `lint` / `lint:fix` / `prettier`: qualité et formatage du code
- `signin` / `signout`: connexion/déconnexion M365 pour les outils de debug

## Structure du projet

```
Firstaddin/
  └─ My Office Add-in/
      ├─ src/
      │  ├─ taskpane/            # UI du volet (HTML/CSS/JS)
      │  └─ commands/            # Commandes UI-less (FunctionFile)
      ├─ assets/                 # Icônes et images
      ├─ manifest-xml-out/
      │  └─ manifest.xml         # Manifeste consommé par les scripts de debug
      ├─ webpack.config.js       # Bundling + dev-server HTTPS
      ├─ babel.config.json       # Transpilation cible
      └─ package.json            # Scripts et dépendances
```

Entrées Webpack principales:

- `taskpane`: `src/taskpane/taskpane.js` + `src/taskpane/taskpane.html`
- `commands`: `src/commands/commands.js` (servi via `commands.html`)

## Développement

- Logique Outlook (lecture/écriture de l’item courant): `src/taskpane/taskpane.js`
- UI et styles du volet: `src/taskpane/taskpane.html` et `src/taskpane/taskpane.css`
- Commandes UI-less (exécution sans UI): `src/commands/commands.js` (associées via `Office.actions.associate`)
- Icônes/visuels: `assets/` (référencés par le manifeste)

Astuce: la page de task pane inclut Office.js depuis le CDN; le code s’initialise sur `Office.onReady`.

## Manifeste (manifest.xml)

Fichier: `My Office Add-in/manifest-xml-out/manifest.xml`

- Hôte: Outlook (`Mailbox`)
- URL du task pane: `https://localhost:3001/taskpane.html`
- URL du FunctionFile (commandes): `https://localhost:3001/commands.html`
- Icônes: `https://localhost:3001/assets/...`
- Permissions: `ReadWriteItem`
- Règles d’activation: Message (Read/Compose)

Valider le manifeste:

```bash
cd "My Office Add-in"
npm run validate
```

## Build

Développement:

```bash
npm run build:dev
```

Production:

```bash
npm run build
```

La sortie webpack (par défaut `dist/`) est servable statiquement en HTTPS.

## Déploiement

1. Héberger les fichiers générés (`dist/`) sur une origine HTTPS publique.
2. Mettre à jour le manifeste pour que `Taskpane.Url`, `Commands.Url`, `IconUrl` et `HighResolutionIconUrl` pointent vers votre domaine de production.
3. Distribuer le manifeste (catalogue d’organisation, Centre d’administration, ou Store).

Astuce: `webpack.config.js` expose une constante `urlProd` à ajuster si vous souhaitez automatiser certaines substitutions, mais le manifeste XML doit être mis à jour explicitement.

## Qualité du code

```bash
npm run lint
npm run lint:fix
npm run prettier
```

## Sécurité

- La démo `taskpane` inclut un test d’appel à l’API OpenAI avec stockage de la clé en `localStorage` (usage DEV uniquement). Ne pas utiliser ce mécanisme en production.
- Toujours servir en HTTPS et limiter les domaines autorisés dans le manifeste.

## Dépannage

- L’add-in ne se charge pas: vérifier que le serveur écoute sur `3001` et que le certificat est approuvé.
- Icons ou HTML non trouvés: aligner les chemins dans le manifeste avec le domaine/port effectif.
- Changement de port: mettre à jour `package.json > config.dev_server_port`, `webpack.config.js` (port déjà synchronisé via npm config), et le manifeste si nécessaire.
- Nettoyer le sideload: `npm run stop` puis redémarrer Outlook.

## Liens utiles

- Documentation Office Add-ins: [learn.microsoft.com/office/dev/add-ins](https://learn.microsoft.com/office/dev/add-ins/)
- Déboguer des compléments Office: [learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins)

## Licence

MIT (voir `license` dans `package.json`)

