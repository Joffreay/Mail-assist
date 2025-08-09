## Mail-assist — Complément Outlook (Task Pane + Commandes)

Projet de complément Office pour Outlook, basé sur Webpack 5, Babel et Office.js. Développement actif à la racine du dépôt (`webpack.config.js`, `src/`, `manifest-xml-out/`). Le dossier `My Office Add-in/` est présent comme gabarit distinct (dépôt imbriqué) et n’est pas utilisé par défaut.

## Aperçu

- Complément Outlook avec volet des tâches (task pane) et commandes UI-less
- Pile technique: JavaScript, Webpack 5, Babel (`@babel/preset-env`), polyfills (`core-js`, `regenerator-runtime`), Office.js
- Hébergement dev: HTTPS local via webpack-dev-server (port 3001)
- Manifeste: XML pointant vers `https://localhost:3001`

## Fonctionnalités

- Volet Outlook (Task Pane):
  - Accueil “Hello world”, affichage du sujet du message courant
  - Démo d’appel API OpenAI (clé saisie côté client, stockage localStorage en DEV uniquement)
  - Lister les prochains rendez-vous (7 jours) via Microsoft Graph après auth MSAL (Dialog)
  - Récupérer les catégories utilisateur Outlook (master categories)
- Commande UI-less: `action(event)` affichant une notification dans l’élément courant

## Prérequis

- Windows et Outlook (Desktop ou OWA selon tests)
- Node.js LTS et npm
- Droits admin recommandés pour approuver les certificats de dev et l’exemption loopback WebView2

## Installation

Depuis la racine du dépôt:

```bash
npm install
```

## Démarrage (développement)

Lancer le serveur local + sideload de l’add-in dans Outlook:

```bash
npm run start
```

Notes:

- Dev-server en HTTPS sur `https://localhost:3001/`
- Arrêter/décharger l’add-in:

```bash
npm run stop
```

Serveur seul (sans sideload):

```bash
npm run dev-server
```

## Scripts npm utiles

- `dev-server`: lance uniquement le serveur webpack (HTTPS)
- `build:dev`: build dev (source maps)
- `build`: build prod
- `watch`: build incrémental
- `validate`: valide le manifeste XML
- `lint` / `lint:fix` / `prettier`: qualité et formatage
- `signin` / `signout`: connexion/déconnexion M365 pour les outils Office Add-in

## Structure du projet (racine)

```
Firstaddin/
  ├─ src/
  │  ├─ taskpane/
  │  │  ├─ taskpane.html / .css / .js
  │  │  ├─ auth.html           # Auth MSAL (Dialog)
  │  │  └─ auth2.html          # Variante d’auth pour contourner le cache
  │  └─ commands/
  ├─ assets/
  ├─ manifest-xml-out/
  │  └─ manifest.xml           # Manifeste utilisé par les scripts de debug
  ├─ dist/                     # Sortie webpack
  ├─ webpack.config.js         # Bundling + dev-server HTTPS
  └─ package.json
```

Entrées Webpack principales:

- `taskpane`: `src/taskpane/taskpane.js` + `src/taskpane/taskpane.html`
- `commands`: `src/commands/commands.js` (servi via `commands.html`)
- Pages statiques servies: `auth.html`, `auth2.html`

## Authentification (MSAL) et Microsoft Graph

- Dialog: `Office.context.ui.displayDialogAsync(https://localhost:3001/auth2.html?clientId=...)`
- MSAL Browser (2.x):
  - `redirectUri`: `https://localhost:3001/auth2.html` (ou `auth.html`)
  - `navigateToLoginRequestUrl: false`
  - `system.allowRedirectInIframe: true`, `cache.storeAuthStateInCookie: true`
  - Fallback popup: `loginPopup`/`acquireTokenPopup` si `Silent` échoue (interaction_required/consent_required)
- Scopes Graph requis: `User.Read`, `Calendars.Read`
- App registration (Entra ID): ajouter des URIs de redirection exactes selon le port/page utilisés:
  - `https://localhost:3001/auth.html`
  - `https://localhost:3001/auth2.html`

## Catégories Outlook (master categories)

- Bouton “Récupérer les catégories”: appelle `Office.context.mailbox.masterCategories.getAsync`
- Exigences:
  - Requirements Mailbox 1.8
  - Permissions du manifeste: `ReadWriteMailbox` (déjà configuré)

## Manifeste (manifest.xml)

Fichier: `manifest-xml-out/manifest.xml`

- Hôte: Outlook (`Mailbox`)
- Task Pane: `https://localhost:3001/taskpane.html`
- FunctionFile: `https://localhost:3001/commands.html`
- Icônes: `https://localhost:3001/assets/...`
- Requirements: Mailbox ≥ 1.8
- Permissions: `ReadWriteMailbox`
- Règles d’activation: Message (Read/Compose)

Valider le manifeste:

```bash
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

La sortie webpack (`dist/`) est servable en HTTPS.

## Dépannage

- Boucle auth bloquée: vérifier que la page servie est bien `auth2.html` récente (voir le code source: présence de `id="status"`, mentions `loginPopup/acquireTokenPopup`).
- Exemption loopback WebView2: autoriser si demandé. En cas d’échec, relancer la commande en admin.
- Port 3001 occupé: `npx kill-port 3001` puis relancer `npm run dev-server`.
- Sous-module accidentel: le dossier `My Office Add-in/` est un dépôt imbriqué; si non désiré dans l’index Git: `git rm --cached "My Office Add-in"`.

## Qualité du code

```bash
npm run lint
npm run lint:fix
npm run prettier
```

## Liens utiles

- Office Add-ins: https://learn.microsoft.com/office/dev/add-ins/
- Debug Office Add-ins: https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins

## Licence

MIT

