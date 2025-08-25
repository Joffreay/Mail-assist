/**
 * Configuration pour le consentement administrateur Microsoft Graph
 * Contient les paramètres et exemples d'URLs pour le consentement administrateur
 */

const AdminConsentConfig = {
    // Configuration par défaut
    default: {
        // Scopes par défaut pour Microsoft Graph
        scopes: ['https://graph.microsoft.com/.default'],
        
        // URI de redirection par défaut
        redirectUri: 'https://localhost:3001/consent/callback',
        
        // Paramètres de l'interface utilisateur
        ui: {
            dialogHeight: 80,
            dialogWidth: 60,
            displayInIframe: false
        }
    },

    // Configuration des endpoints
    endpoints: {
        // Endpoint de consentement administrateur
        adminConsent: 'https://login.microsoftonline.com/{TENANT_ID}/v2.0/adminconsent',
        
        // Endpoint de consentement utilisateur
        userConsent: 'https://login.microsoftonline.com/{TENANT_ID}/v2.0/authorize',
        
        // Endpoint de token
        token: 'https://login.microsoftonline.com/{TENANT_ID}/v2.0/token'
    },

    // Configuration des permissions
    permissions: {
        // Permissions pour les emails
        mail: [
            'Mail.ReadWrite',
            'Mail.ReadWrite.Shared',
            'MailboxSettings.ReadWrite'
        ],
        
        // Permissions pour les utilisateurs
        user: [
            'User.Read',
            'User.Read.All',
            'User.ReadWrite.All'
        ],
        
        // Permissions pour les calendriers
        calendar: [
            'Calendars.ReadWrite',
            'Calendars.ReadWrite.Shared'
        ],
        
        // Permissions pour les contacts
        contacts: [
            'Contacts.ReadWrite',
            'Contacts.ReadWrite.Shared'
        ],
        
        // Permissions pour les fichiers
        files: [
            'Files.ReadWrite',
            'Files.ReadWrite.All',
            'Files.ReadWrite.AppFolder'
        ]
    },

    // Configuration des scopes
    scopes: {
        // Scope par défaut pour toutes les permissions
        default: 'https://graph.microsoft.com/.default',
        
        // Scopes spécifiques par catégorie
        mail: 'https://graph.microsoft.com/Mail.ReadWrite',
        user: 'https://graph.microsoft.com/User.Read.All',
        calendar: 'https://graph.microsoft.com/Calendars.ReadWrite',
        contacts: 'https://graph.microsoft.com/Contacts.ReadWrite',
        files: 'https://graph.microsoft.com/Files.ReadWrite.All'
    },

    // Exemples d'URLs de consentement administrateur
    examples: {
        // Pour un tenant spécifique
        specificTenant: {
            description: 'Consentement administrateur pour un tenant spécifique',
            url: 'https://login.microsoftonline.com/{TENANT_ID}/v2.0/adminconsent?client_id={APP_CLIENT_ID}&scope=https://graph.microsoft.com/.default&redirect_uri=https://localhost:3001/consent/callback',
            parameters: {
                TENANT_ID: '12345678-1234-1234-1234-123456789012',
                APP_CLIENT_ID: '87654321-4321-4321-4321-210987654321'
            }
        },
        
        // Pour tous les tenants (common)
        allTenants: {
            description: 'Consentement administrateur pour tous les tenants',
            url: 'https://login.microsoftonline.com/common/v2.0/adminconsent?client_id={APP_CLIENT_ID}&scope=https://graph.microsoft.com/.default&redirect_uri=https://localhost:3001/consent/callback',
            parameters: {
                APP_CLIENT_ID: '87654321-4321-4321-4321-210987654321'
            }
        },
        
        // Avec des scopes spécifiques
        specificScopes: {
            description: 'Consentement administrateur avec des scopes spécifiques',
            url: 'https://login.microsoftonline.com/{TENANT_ID}/v2.0/adminconsent?client_id={APP_CLIENT_ID}&scope=Mail.ReadWrite%20User.Read.All%20Calendars.ReadWrite&redirect_uri=https://localhost:3001/consent/callback',
            parameters: {
                TENANT_ID: '12345678-1234-1234-1234-123456789012',
                APP_CLIENT_ID: '87654321-4321-4321-4321-210987654321'
            }
        }
    },

    // Configuration des messages d'erreur
    errorMessages: {
        // Erreurs de configuration
        config: {
            CLIENT_ID_MISSING: 'Client ID manquant. Veuillez configurer l\'ID client de l\'application.',
            TENANT_ID_MISSING: 'Tenant ID manquant. Veuillez spécifier l\'ID du tenant.',
            REDIRECT_URI_MISSING: 'URI de redirection manquant. Veuillez configurer l\'URI de redirection.',
            SCOPES_MISSING: 'Aucun scope défini. Veuillez configurer au moins un scope.'
        },
        
        // Erreurs de consentement
        consent: {
            CONSENT_REFUSED: 'Le consentement administrateur a été refusé.',
            CONSENT_TIMEOUT: 'Le consentement administrateur a expiré.',
            CONSENT_ERROR: 'Une erreur s\'est produite lors du consentement administrateur.',
            INVALID_TENANT: 'Tenant invalide ou non autorisé.',
            INSUFFICIENT_PERMISSIONS: 'Permissions insuffisantes pour accorder le consentement.'
        },
        
        // Erreurs de réseau
        network: {
            NETWORK_ERROR: 'Erreur de connexion réseau.',
            TIMEOUT_ERROR: 'Délai d\'attente dépassé.',
            SERVER_ERROR: 'Erreur du serveur Azure AD.'
        }
    },

    // Configuration des logs
    logging: {
        // Niveaux de log
        levels: {
            DEBUG: 'debug',
            INFO: 'info',
            WARN: 'warn',
            ERROR: 'error'
        },
        
        // Activation des logs
        enabled: true,
        
        // Niveau de log par défaut
        defaultLevel: 'info',
        
        // Préfixe des logs
        prefix: '[AdminConsent]'
    },

    // Configuration de sécurité
    security: {
        // Timeout du consentement (en millisecondes)
        consentTimeout: 300000, // 5 minutes
        
        // Validation des paramètres
        validateParameters: true,
        
        // Vérification de l'état
        validateState: true,
        
        // Chiffrement des données sensibles
        encryptSensitiveData: false
    }
};

// Fonctions utilitaires pour la configuration
const AdminConsentUtils = {
    /**
     * Génère une URL de consentement administrateur
     * @param {Object} config - Configuration pour l'URL
     * @returns {string} URL de consentement administrateur
     */
    generateAdminConsentUrl(config = {}) {
        const {
            tenantId = 'common',
            clientId,
            scopes = AdminConsentConfig.default.scopes,
            redirectUri = AdminConsentConfig.default.redirectUri,
            prompt = 'consent'
        } = config;

        if (!clientId) {
            throw new Error(AdminConsentConfig.errorMessages.config.CLIENT_ID_MISSING);
        }

        // Construire l'URL
        const url = new URL(AdminConsentConfig.endpoints.adminConsent.replace('{TENANT_ID}', tenantId));
        
        // Ajouter les paramètres
        url.searchParams.set('client_id', clientId);
        url.searchParams.set('scope', scopes.join(' '));
        url.searchParams.set('redirect_uri', redirectUri);
        
        if (prompt) {
            url.searchParams.set('prompt', prompt);
        }

        return url.toString();
    },

    /**
     * Valide la configuration du consentement administrateur
     * @param {Object} config - Configuration à valider
     * @returns {Object} Résultat de la validation
     */
    validateConfig(config) {
        const errors = [];
        const warnings = [];

        // Vérifier le Client ID
        if (!config.clientId) {
            errors.push(AdminConsentConfig.errorMessages.config.CLIENT_ID_MISSING);
        }

        // Vérifier le Tenant ID
        if (!config.tenantId) {
            warnings.push(AdminConsentConfig.errorMessages.config.TENANT_ID_MISSING);
        }

        // Vérifier l'URI de redirection
        if (!config.redirectUri) {
            errors.push(AdminConsentConfig.errorMessages.config.REDIRECT_URI_MISSING);
        }

        // Vérifier les scopes
        if (!config.scopes || config.scopes.length === 0) {
            errors.push(AdminConsentConfig.errorMessages.config.SCOPES_MISSING);
        }

        return {
            isValid: errors.length === 0,
            errors: errors,
            warnings: warnings,
            hasWarnings: warnings.length > 0
        };
    },

    /**
     * Obtient les permissions par catégorie
     * @param {string} category - Catégorie de permissions
     * @returns {Array} Liste des permissions
     */
    getPermissionsByCategory(category) {
        return AdminConsentConfig.permissions[category] || [];
    },

    /**
     * Obtient tous les scopes disponibles
     * @returns {Object} Tous les scopes par catégorie
     */
    getAllScopes() {
        return AdminConsentConfig.scopes;
    },

    /**
     * Génère un exemple d'URL avec des paramètres réels
     * @param {string} exampleKey - Clé de l'exemple
     * @param {Object} parameters - Paramètres à remplacer
     * @returns {string} URL d'exemple avec paramètres
     */
    generateExampleUrl(exampleKey, parameters = {}) {
        const example = AdminConsentConfig.examples[exampleKey];
        if (!example) {
            throw new Error(`Exemple non trouvé: ${exampleKey}`);
        }

        let url = example.url;
        
        // Remplacer les paramètres
        Object.keys(parameters).forEach(key => {
            const placeholder = `{${key}}`;
            url = url.replace(new RegExp(placeholder, 'g'), parameters[key]);
        });

        return url;
    },

    /**
     * Formate un message d'erreur
     * @param {string} errorKey - Clé de l'erreur
     * @param {Object} context - Contexte de l'erreur
     * @returns {string} Message d'erreur formaté
     */
    formatErrorMessage(errorKey, context = {}) {
        const errorMessage = AdminConsentConfig.errorMessages[errorKey];
        if (!errorMessage) {
            return `Erreur inconnue: ${errorKey}`;
        }

        // Remplacer les variables dans le message
        let message = errorMessage;
        Object.keys(context).forEach(key => {
            const placeholder = `{${key}}`;
            message = message.replace(new RegExp(placeholder, 'g'), context[key]);
        });

        return message;
    }
};

// Export pour utilisation dans d'autres modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { AdminConsentConfig, AdminConsentUtils };
} else if (typeof window !== 'undefined') {
    window.AdminConsentConfig = AdminConsentConfig;
    window.AdminConsentUtils = AdminConsentUtils;
}
