# Office Add-in Task Pane - Assistant IA Intégré

## 📋 Description

Cet Office Add-in pour Outlook intègre des fonctionnalités d'intelligence artificielle via OpenAI et Microsoft Graph pour améliorer la productivité des utilisateurs. Il permet la génération automatique de réponses d'email, la gestion des calendriers, et l'analyse intelligente du contenu des messages.

## ✨ Fonctionnalités Principales

### 🤖 Intégration OpenAI (ChatGPT)
- **Génération de réponses automatiques** : Création de réponses professionnelles basées sur le contenu des emails
- **Assistant IA conversationnel** : Interface de chat pour poser des questions et obtenir des réponses
- **Analyse de contenu** : Extraction et analyse intelligente du contenu des messages

### 🔐 Authentification Microsoft Graph
- **Connexion sécurisée** : Authentification OAuth2 avec Microsoft Identity Platform
- **Gestion des tokens** : Gestion automatique des tokens d'accès et de rafraîchissement
- **Sécurité renforcée** : Protection CSRF et validation des états de sécurité

### 📅 Gestion des Calendriers
- **Synchronisation des événements** : Récupération des rendez-vous via Microsoft Graph
- **Vue calendrier** : Affichage des événements à venir dans les 7 prochains jours
- **Intégration Outlook** : Utilisation native des fonctionnalités Outlook

### 📧 Fonctionnalités Outlook
- **Extraction de contenu** : Récupération automatique du sujet et du corps des emails
- **Gestion des catégories** : Intégration avec le système de catégories Outlook
- **Interface utilisateur native** : Design cohérent avec l'interface Office

## 🚀 Installation et Configuration

### Prérequis
- **Node.js** : Version 16 ou supérieure
- **Office 365** : Compte Microsoft 365 avec Outlook
- **Azure AD** : Application enregistrée pour Microsoft Graph (optionnel)
- **OpenAI** : Clé API OpenAI pour les fonctionnalités IA

### Installation

1. **Cloner le repository**
   ```bash
   git clone <repository-url>
   cd "My Office Add-in"
   ```

2. **Installer les dépendances**
   ```bash
   npm install
   ```

3. **Configuration des certificats de développement**
   ```bash
   npx office-addin-dev-certs install
   ```

4. **Configuration des variables d'environnement**
   ```bash
   cp config.env.example config.env
   # Éditer config.env avec vos clés API
   ```

### Configuration Azure AD (Optionnel)

1. **Créer une application dans Azure AD**
   - Aller sur [Azure Portal](https://portal.azure.com)
   - Créer une nouvelle application d'enregistrement
   - Configurer les redirections URI

2. **Configurer les permissions**
   - `Mail.ReadWrite`
   - `Calendars.Read`
   - `User.Read`

3. **Obtenir le Client ID**
   - Copier l'ID d'application (Client ID)
   - L'utiliser dans l'interface de l'add-in

## 🛠️ Développement

### Scripts NPM Disponibles

```bash
# Développement
npm run dev-server          # Démarrer le serveur de développement
npm run build:dev          # Build en mode développement
npm run watch              # Build en mode watch

# Production
npm run build              # Build de production
npm run start              # Démarrer l'add-in

# Utilitaires
npm run lint               # Vérification du code
npm run lint:fix           # Correction automatique du linting
npm run validate           # Validation du manifest
npm run signin             # Connexion M365
npm run signout            # Déconnexion M365

# Services
npm run callback-server    # Démarrer le serveur de callback
npm run webhook-server     # Démarrer le serveur webhook
npm run start-all          # Démarrer tous les services
```

### Structure du Projet

```
My Office Add-in/
├── src/                          # Code source
│   ├── taskpane/                # Interface principale
│   │   ├── taskpane.js         # Logique principale
│   │   ├── taskpane.html       # Interface utilisateur
│   │   ├── taskpane.css        # Styles
│   │   ├── graph-auth-service.js # Service d'authentification
│   │   └── config.js           # Configuration
│   ├── commands/                # Commandes du ruban
│   │   ├── commands.js         # Logique des commandes
│   │   └── commands.html       # Interface des commandes
│   └── assets/                  # Ressources statiques
├── manifest-xml-out/            # Manifest XML généré
├── dist/                        # Fichiers compilés
├── webpack.config.js            # Configuration Webpack
├── server.js                    # Serveur webhook
└── package.json                 # Dépendances et scripts
```

### Architecture du Code

#### TaskPaneManager
Classe principale qui gère l'initialisation et le cycle de vie de l'add-in.

#### MicrosoftGraphAuthService
Service dédié à l'authentification et à la gestion des tokens Microsoft Graph.

#### Configuration Centralisée
Fichier `config.js` qui centralise tous les paramètres et constantes de l'application.

## 🔧 Configuration

### Variables d'Environnement

```env
# OpenAI
OPENAI_API_KEY=sk-your-api-key-here
OPENAI_MODEL=gpt-4

# Microsoft Graph
GRAPH_CLIENT_ID=your-client-id
GRAPH_TENANT_ID=your-tenant-id

# Serveur
PORT=8080
NODE_ENV=development
DEBUG=true
```

### Configuration Webpack

Le fichier `webpack.config.js` gère :
- Compilation Babel pour la compatibilité ES6+
- Génération des pages HTML
- Gestion des assets et certificats HTTPS
- Configuration du serveur de développement

## 📱 Utilisation

### Interface Principale

1. **Section OpenAI** : Saisir votre clé API et poser des questions
2. **Actions Outlook** : Boutons pour afficher le sujet et générer des réponses
3. **Calendrier** : Saisir votre Client ID Azure AD pour lister les événements

### Fonctionnalités Clés

#### Génération de Réponse Automatique
1. Saisir votre clé API OpenAI
2. Cliquer sur "Réponse automatique"
3. L'IA analyse le contenu et génère une réponse appropriée
4. La réponse s'ouvre dans un formulaire de réponse Outlook

#### Consultation du Calendrier
1. Saisir votre Client ID Azure AD
2. Cliquer sur "Lister mes prochains RDV"
3. S'authentifier via Microsoft Graph
4. Consulter vos événements des 7 prochains jours

## 🔒 Sécurité

### Authentification
- **OAuth2** : Utilisation du protocole standard Microsoft
- **Tokens sécurisés** : Gestion automatique des tokens d'accès
- **Validation d'état** : Protection contre les attaques CSRF

### Données
- **Stockage local** : Tokens stockés localement (développement uniquement)
- **Chiffrement** : Communication HTTPS obligatoire
- **Permissions minimales** : Accès limité aux données nécessaires

### API Keys
- **OpenAI** : Stockage local temporaire (à sécuriser en production)
- **Azure AD** : Gestion sécurisée via Microsoft Identity Platform

## 🧪 Tests

### Tests Unitaires
```bash
npm run test:unit
```

### Tests d'Intégration
```bash
npm run test:integration
```

### Tests End-to-End
```bash
npm run test:e2e
```

## 📦 Déploiement

### Développement
```bash
npm run build:dev
npm run start
```

### Production
```bash
npm run build
# Déployer le contenu du dossier dist/
```

### Manifest XML
Le manifest est généré au format XML pour une meilleure compatibilité avec les versions récentes d'Office.

## 🐛 Dépannage

### Problèmes Courants

#### Erreur de Certificat HTTPS
```bash
npx office-addin-dev-certs install
npx office-addin-dev-certs verify
```

#### Erreur d'Authentification Graph
- Vérifier que l'application Azure AD est correctement configurée
- Contrôler les permissions et scopes
- Vérifier les URIs de redirection

#### Erreur OpenAI API
- Vérifier la validité de la clé API
- Contrôler les quotas et limites
- Vérifier la connectivité réseau

### Logs et Débogage

```bash
# Activer le mode debug
DEBUG=true npm run dev-server

# Consulter les logs du serveur
npm run webhook-server
```

## 🤝 Contribution

### Guidelines de Code
- **ES6+** : Utilisation des fonctionnalités modernes JavaScript
- **Commentaires** : Documentation JSDoc complète
- **Linting** : Respect des règles ESLint
- **Tests** : Couverture de tests pour les nouvelles fonctionnalités

### Processus de Contribution
1. Fork du repository
2. Création d'une branche feature
3. Développement et tests
4. Pull request avec description détaillée

## 📄 Licence

Ce projet est sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus de détails.

## 🙏 Remerciements

- **Microsoft** : Pour l'infrastructure Office Add-ins
- **OpenAI** : Pour l'API d'intelligence artificielle
- **Communauté** : Pour les contributions et retours

## 📞 Support

### Documentation
- [Documentation Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/)
- [Microsoft Graph API](https://docs.microsoft.com/graph/)
- [OpenAI API](https://platform.openai.com/docs/)

### Issues
Pour signaler un bug ou demander une fonctionnalité, utilisez les [Issues GitHub](https://github.com/your-repo/issues).

### Contact
- **Email** : support@your-domain.com
- **Discord** : [Serveur communautaire](https://discord.gg/your-server)

---

**Version** : 1.0.0  
**Dernière mise à jour** : Décembre 2024  
**Compatibilité** : Office 365, Outlook Desktop/Web

