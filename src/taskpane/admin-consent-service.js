/**
 * Service de gestion du consentement administrateur Microsoft Graph
 * Gère l'URL de consentement administrateur et le traitement des callbacks
 */
class AdminConsentService {
    constructor() {
        this.clientId = null;
        this.tenantId = null;
        this.redirectUri = null;
        this.scopes = ['https://graph.microsoft.com/.default'];
        
        // Événements personnalisés
        this.events = {
            onConsentSuccess: null,
            onConsentError: null,
            onConsentRefused: null,
            onConsentPending: null
        };
    }

    /**
     * Initialise le service de consentement administrateur
     * @param {string} clientId - ID client de l'application
     * @param {string} tenantId - ID du tenant (optionnel, utilise 'common' si non spécifié)
     * @param {string} redirectUri - URI de redirection
     */
    initialize(clientId, tenantId = null, redirectUri = null) {
        this.clientId = clientId;
        this.tenantId = tenantId || 'common';
        this.redirectUri = redirectUri || `${window.location.origin}/consent/callback`;
        
        console.log('Service de consentement administrateur initialisé');
        console.log('Client ID:', this.clientId);
        console.log('Tenant ID:', this.tenantId);
        console.log('Redirect URI:', this.redirectUri);
    }

    /**
     * Génère l'URL de consentement administrateur
     * @param {Object} options - Options de configuration
     * @returns {string} URL de consentement administrateur
     */
    generateAdminConsentUrl(options = {}) {
        if (!this.clientId) {
            throw new Error('Service non initialisé. Appelez initialize() d\'abord.');
        }

        const {
            tenantId = this.tenantId,
            scopes = this.scopes,
            redirectUri = this.redirectUri,
            prompt = 'consent'
        } = options;

        // Construire l'URL de consentement administrateur
        const adminConsentUrl = new URL(`https://login.microsoftonline.com/${tenantId}/v2.0/adminconsent`);
        
        // Paramètres obligatoires
        adminConsentUrl.searchParams.set('client_id', this.clientId);
        adminConsentUrl.searchParams.set('scope', scopes.join(' '));
        adminConsentUrl.searchParams.set('redirect_uri', redirectUri);
        
        // Paramètres optionnels
        if (prompt) {
            adminConsentUrl.searchParams.set('prompt', prompt);
        }

        return adminConsentUrl.toString();
    }

    /**
     * Génère l'URL de consentement administrateur avec des paramètres personnalisés
     * @param {string} customTenantId - ID du tenant personnalisé
     * @param {string} customRedirectUri - URI de redirection personnalisée
     * @returns {string} URL de consentement administrateur personnalisée
     */
    generateCustomAdminConsentUrl(customTenantId = null, customRedirectUri = null) {
        return this.generateAdminConsentUrl({
            tenantId: customTenantId || this.tenantId,
            redirectUri: customRedirectUri || this.redirectUri
        });
    }

    /**
     * Traite la réponse du consentement administrateur
     * @param {Object} urlParams - Paramètres de l'URL de callback
     * @returns {Object} Résultat du traitement
     */
    processAdminConsentResponse(urlParams) {
        try {
            const adminConsent = urlParams.get('admin_consent');
            const tenant = urlParams.get('tenant');
            const error = urlParams.get('error');
            const errorDescription = urlParams.get('error_description');

            // Vérifier s'il y a une erreur
            if (error) {
                const errorResult = {
                    success: false,
                    error: error,
                    errorDescription: errorDescription,
                    tenant: tenant,
                    type: 'error'
                };

                if (this.events.onConsentError) {
                    this.events.onConsentError(errorResult);
                }

                return errorResult;
            }

            // Traiter le résultat du consentement
            if (adminConsent === 'True' && tenant) {
                const successResult = {
                    success: true,
                    adminConsent: true,
                    tenant: tenant,
                    type: 'success',
                    message: 'Consentement administrateur accordé avec succès'
                };

                if (this.events.onConsentSuccess) {
                    this.events.onConsentSuccess(successResult);
                }

                return successResult;

            } else if (adminConsent === 'False') {
                const refusedResult = {
                    success: false,
                    adminConsent: false,
                    tenant: tenant,
                    type: 'refused',
                    message: 'Consentement administrateur refusé'
                };

                if (this.events.onConsentRefused) {
                    this.events.onConsentRefused(refusedResult);
                }

                return refusedResult;

            } else {
                const pendingResult = {
                    success: false,
                    adminConsent: null,
                    tenant: tenant,
                    type: 'pending',
                    message: 'En attente du consentement administrateur'
                };

                if (this.events.onConsentPending) {
                    this.events.onConsentPending(pendingResult);
                }

                return pendingResult;
            }

        } catch (error) {
            console.error('Erreur lors du traitement de la réponse de consentement:', error);
            
            const errorResult = {
                success: false,
                error: 'PROCESSING_ERROR',
                errorDescription: error.message,
                type: 'error'
            };

            if (this.events.onConsentError) {
                this.events.onConsentError(errorResult);
            }

            return errorResult;
        }
    }

    /**
     * Ouvre la page de consentement administrateur
     * @param {Object} options - Options de configuration
     * @returns {Promise<boolean>} True si l'ouverture a réussi
     */
    async openAdminConsentPage(options = {}) {
        try {
            const adminConsentUrl = this.generateAdminConsentUrl(options);
            
            // Ouvrir dans une nouvelle fenêtre
            const consentWindow = window.open(
                adminConsentUrl,
                'admin_consent',
                'width=800,height=600,scrollbars=yes,resizable=yes'
            );

            if (!consentWindow) {
                throw new Error('Impossible d\'ouvrir la fenêtre de consentement administrateur. Vérifiez que les popups ne sont pas bloqués.');
            }

            return true;

        } catch (error) {
            console.error('Erreur lors de l\'ouverture de la page de consentement:', error);
            throw error;
        }
    }

    /**
     * Ouvre la page de consentement administrateur dans un dialogue Office
     * @param {Object} options - Options de configuration
     * @returns {Promise<Object>} Résultat du consentement
     */
    async openAdminConsentDialog(options = {}) {
        return new Promise((resolve, reject) => {
            try {
                const adminConsentUrl = this.generateAdminConsentUrl(options);
                
                // Ouvrir le dialogue Office
                Office.context.ui.displayDialogAsync(
                    adminConsentUrl,
                    { 
                        height: 80, 
                        width: 60, 
                        displayInIframe: false 
                    },
                    (res) => {
                        if (res.status !== Office.AsyncResultStatus.Succeeded) {
                            reject(new Error("Impossible d'ouvrir la fenêtre de consentement administrateur."));
                            return;
                        }

                        const dialog = res.value;
                        
                        // Écouter les messages de la page de callback
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                            try {
                                const msg = JSON.parse(arg.message || "{}");
                                
                                if (msg.type === "consent_result") {
                                    dialog.close();
                                    resolve(msg.result);
                                } else if (msg.type === "close_consent_window") {
                                    dialog.close();
                                    resolve({ type: 'cancelled', message: 'Consentement annulé par l\'utilisateur' });
                                }
                            } catch (e) {
                                console.error('Erreur lors du traitement du message:', e);
                                dialog.close();
                                reject(e);
                            }
                        });

                        // Gérer la fermeture de la fenêtre
                        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                            if (arg.errorCode) {
                                reject(new Error(`Erreur de dialogue: ${arg.errorCode}`));
                            }
                        });

                        // Timeout de sécurité
                        setTimeout(() => {
                            dialog.close();
                            reject(new Error('Timeout du dialogue de consentement'));
                        }, 300000); // 5 minutes
                    }
                );

            } catch (error) {
                reject(error);
            }
        });
    }

    /**
     * Vérifie le statut du consentement administrateur
     * @param {string} tenantId - ID du tenant à vérifier
     * @returns {Promise<Object>} Statut du consentement
     */
    async checkAdminConsentStatus(tenantId = null) {
        // Note: Cette fonction nécessiterait une API backend pour vérifier le statut
        // Pour l'instant, elle retourne un statut simulé
        
        const targetTenantId = tenantId || this.tenantId;
        
        return {
            tenantId: targetTenantId,
            hasAdminConsent: false, // À implémenter avec une vraie vérification
            lastChecked: new Date().toISOString(),
            message: 'Vérification du statut non implémentée'
        };
    }

    /**
     * Génère un rapport de consentement administrateur
     * @param {Object} consentResult - Résultat du consentement
     * @returns {Object} Rapport formaté
     */
    generateConsentReport(consentResult) {
        const report = {
            timestamp: new Date().toISOString(),
            clientId: this.clientId,
            tenantId: this.tenantId,
            redirectUri: this.redirectUri,
            result: consentResult,
            summary: this.generateConsentSummary(consentResult)
        };

        return report;
    }

    /**
     * Génère un résumé du consentement
     * @param {Object} consentResult - Résultat du consentement
     * @returns {string} Résumé formaté
     */
    generateConsentSummary(consentResult) {
        switch (consentResult.type) {
            case 'success':
                return `✅ Consentement administrateur accordé avec succès pour le tenant ${consentResult.tenant}`;
            case 'refused':
                return `❌ Consentement administrateur refusé pour le tenant ${consentResult.tenant}`;
            case 'pending':
                return `⏳ En attente du consentement administrateur`;
            case 'error':
                return `🚨 Erreur lors du consentement: ${consentResult.error} - ${consentResult.errorDescription}`;
            default:
                return `❓ Statut de consentement inconnu`;
        }
    }

    /**
     * Valide la configuration du service
     * @returns {Object} Résultat de la validation
     */
    validateConfiguration() {
        const errors = [];
        const warnings = [];

        if (!this.clientId) {
            errors.push('Client ID manquant');
        }

        if (!this.redirectUri) {
            errors.push('URI de redirection manquant');
        }

        if (this.tenantId === 'common') {
            warnings.push('Utilisation du tenant "common" - le consentement sera demandé pour chaque tenant');
        }

        if (!this.scopes || this.scopes.length === 0) {
            errors.push('Aucun scope défini');
        }

        return {
            isValid: errors.length === 0,
            errors: errors,
            warnings: warnings,
            hasWarnings: warnings.length > 0
        };
    }

    /**
     * Définit un gestionnaire d'événement
     * @param {string} eventName - Nom de l'événement
     * @param {Function} handler - Gestionnaire d'événement
     */
    on(eventName, handler) {
        if (this.events.hasOwnProperty(eventName)) {
            this.events[eventName] = handler;
        } else {
            console.warn(`Événement inconnu: ${eventName}`);
        }
    }

    /**
     * Supprime un gestionnaire d'événement
     * @param {string} eventName - Nom de l'événement
     */
    off(eventName) {
        if (this.events.hasOwnProperty(eventName)) {
            this.events[eventName] = null;
        }
    }

    /**
     * Obtient la configuration actuelle du service
     * @returns {Object} Configuration actuelle
     */
    getConfiguration() {
        return {
            clientId: this.clientId,
            tenantId: this.tenantId,
            redirectUri: this.redirectUri,
            scopes: this.scopes
        };
    }

    /**
     * Met à jour la configuration du service
     * @param {Object} config - Nouvelle configuration
     */
    updateConfiguration(config) {
        if (config.clientId !== undefined) this.clientId = config.clientId;
        if (config.tenantId !== undefined) this.tenantId = config.tenantId;
        if (config.redirectUri !== undefined) this.redirectUri = config.redirectUri;
        if (config.scopes !== undefined) this.scopes = config.scopes;

        console.log('Configuration mise à jour:', this.getConfiguration());
    }
}

// Export pour utilisation dans d'autres modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = AdminConsentService;
} else if (typeof window !== 'undefined') {
    window.AdminConsentService = AdminConsentService;
}
