/**
 * Service d'authentification Microsoft Graph
 * 
 * Ce service gère l'authentification OAuth2 avec Microsoft Graph, incluant :
 * - L'initialisation du flux d'authentification
 * - L'échange de codes d'autorisation contre des tokens
 * - La gestion des tokens d'accès et de rafraîchissement
 * - Le rafraîchissement automatique des tokens expirés
 * - La persistance locale des tokens (développement uniquement)
 * 
 * @author Microsoft Corporation (modifié)
 * @license MIT
 * @version 1.0.0
 */

/**
 * Configuration par défaut du service d'authentification
 */
const DEFAULT_CONFIG = {
  // Scopes d'autorisation Microsoft Graph
  SCOPES: [
    'openid',
    'profile', 
    'email',
    'Mail.ReadWrite',
    'Mail.ReadWrite.Shared',
    'MailboxSettings.ReadWrite',
    'User.Read',
    'Calendars.Read'
  ],
  
  // Endpoints Microsoft Identity
  ENDPOINTS: {
    AUTHORIZE: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
    TOKEN: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    GRAPH_ME: 'https://graph.microsoft.com/v1.0/me'
  },
  
  // Paramètres de sécurité
  SECURITY: {
    STATE_LENGTH: 30,
    TOKEN_REFRESH_THRESHOLD: 5 * 60 * 1000 // 5 minutes
  },
  
  // Clés de stockage local
  STORAGE_KEYS: {
    TOKENS: 'msal_tokens',
    STATE: 'msal_state'
  }
};

/**
 * Service d'authentification Microsoft Graph
 */
class MicrosoftGraphAuthService {
  /**
   * Constructeur du service
   */
  constructor() {
    // Configuration de base
    this.clientId = null;
    this.redirectUri = null;
    this.scopes = [...DEFAULT_CONFIG.SCOPES];
    
    // État de l'authentification
    this.accessToken = null;
    this.refreshToken = null;
    this.tokenExpiry = null;
    this.userInfo = null;
    
    // Gestionnaire d'événements
    this.eventHandlers = new Map();
    
    // Initialiser les gestionnaires d'événements par défaut
    this.initializeDefaultEventHandlers();
    
    console.log('🔐 Service d\'authentification Microsoft Graph créé');
  }

  /**
   * Initialise les gestionnaires d'événements par défaut
   */
  initializeDefaultEventHandlers() {
    // Événements disponibles
    const defaultEvents = [
      'onAuthSuccess',
      'onAuthError', 
      'onTokenRefresh',
      'onUserInfoLoaded',
      'onLogout'
    ];
    
    defaultEvents.forEach(eventName => {
      this.eventHandlers.set(eventName, null);
    });
  }

  /**
   * Initialise le service d'authentification
   * 
   * @param {string} clientId - ID client Azure AD
   * @param {string} redirectUri - URI de redirection après authentification
   * @param {string[]} [customScopes] - Scopes personnalisés (optionnel)
   */
  initialize(clientId, redirectUri, customScopes = null) {
    if (!clientId || !redirectUri) {
      throw new Error('clientId et redirectUri sont requis pour l\'initialisation');
    }

    this.clientId = clientId;
    this.redirectUri = redirectUri;
    
    // Utiliser des scopes personnalisés si fournis
    if (customScopes && Array.isArray(customScopes)) {
      this.scopes = [...customScopes];
    }
    
    // Charger les tokens sauvegardés
    this.loadSavedTokens();
    
    console.log('✅ Service d\'authentification initialisé avec:', {
      clientId: this.clientId,
      redirectUri: this.redirectUri,
      scopes: this.scopes
    });
  }

  /**
   * Démarre le processus d'authentification
   * 
   * @returns {string} URL d'authentification Microsoft
   * @throws {Error} Si le service n'est pas initialisé
   */
  startAuth() {
    if (!this.isInitialized()) {
      throw new Error('Service non initialisé. Appelez initialize() d\'abord.');
    }

    // Générer un état unique pour la sécurité
    const state = this.generateSecureState();
    
    // Construire l'URL d'authentification
    const authUrl = new URL(DEFAULT_CONFIG.ENDPOINTS.AUTHORIZE);
    const params = {
      client_id: this.clientId,
      response_type: 'code',
      redirect_uri: this.redirectUri,
      scope: this.scopes.join(' '),
      response_mode: 'query',
      state: state,
      prompt: 'consent'
    };
    
    // Ajouter les paramètres à l'URL
    Object.entries(params).forEach(([key, value]) => {
      authUrl.searchParams.set(key, value);
    });
    
    // Sauvegarder l'état pour validation
    this.saveState(state);
    
    console.log('🚀 Démarrage de l\'authentification avec l\'état:', state);
    return authUrl.toString();
  }

  /**
   * Traite la réponse d'authentification
   * 
   * @param {string} code - Code d'autorisation reçu
   * @param {string} state - État de sécurité reçu
   * @returns {Promise<Object>} Informations de l'utilisateur
   * @throws {Error} Si l'authentification échoue
   */
  async processAuthResponse(code, state) {
    try {
      console.log('📥 Traitement de la réponse d\'authentification');
      
      // Valider l'état de sécurité
      if (!this.validateState(state)) {
        throw new Error('État de sécurité invalide - possible attaque CSRF');
      }
      
      // Échanger le code contre un token
      const tokenResponse = await this.exchangeCodeForToken(code);
      
      // Sauvegarder les tokens
      this.saveTokens(tokenResponse);
      
      // Charger les informations utilisateur
      await this.loadUserInfo();
      
      // Déclencher l'événement de succès
      this.triggerEvent('onAuthSuccess', this.userInfo);
      
      console.log('✅ Authentification réussie pour:', this.userInfo.displayName);
      return this.userInfo;
      
    } catch (error) {
      console.error('❌ Erreur lors du traitement de la réponse d\'authentification:', error);
      
      // Déclencher l'événement d'erreur
      this.triggerEvent('onAuthError', error);
      
      throw error;
    }
  }

  /**
   * Échange le code d'autorisation contre un token
   * 
   * @param {string} code - Code d'autorisation
   * @returns {Promise<Object>} Réponse avec les tokens
   * @throws {Error} Si l'échange échoue
   */
  async exchangeCodeForToken(code) {
    console.log('🔄 Échange du code d\'autorisation contre un token');
    
    // Note: En production, cet échange devrait se faire côté serveur
    // pour des raisons de sécurité (secret client)
    
    const tokenUrl = DEFAULT_CONFIG.ENDPOINTS.TOKEN;
    const body = new URLSearchParams({
      client_id: this.clientId,
      scope: this.scopes.join(' '),
      code: code,
      redirect_uri: this.redirectUri,
      grant_type: 'authorization_code'
    });

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: body
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Erreur lors de l'échange de token: ${response.status} - ${errorText}`);
    }

    const tokenData = await response.json();
    console.log('✅ Token reçu avec succès');
    
    return tokenData;
  }

  /**
   * Charge les informations de l'utilisateur connecté
   * 
   * @returns {Promise<Object>} Informations utilisateur
   * @throws {Error} Si le chargement échoue
   */
  async loadUserInfo() {
    if (!this.accessToken) {
      throw new Error('Aucun token d\'accès disponible');
    }

    try {
      console.log('👤 Chargement des informations utilisateur');
      
      const response = await fetch(DEFAULT_CONFIG.ENDPOINTS.GRAPH_ME, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`Erreur lors du chargement des informations utilisateur: ${response.status}`);
      }

      this.userInfo = await response.json();
      
      // Déclencher l'événement
      this.triggerEvent('onUserInfoLoaded', this.userInfo);
      
      console.log('✅ Informations utilisateur chargées:', this.userInfo.displayName);
      return this.userInfo;
      
    } catch (error) {
      console.error('❌ Erreur lors du chargement des informations utilisateur:', error);
      throw error;
    }
  }

  /**
   * Vérifie si l'utilisateur est authentifié
   * 
   * @returns {boolean} True si l'utilisateur est authentifié et le token valide
   */
  isAuthenticated() {
    return !!(this.accessToken && this.tokenExpiry && Date.now() < this.tokenExpiry);
  }

  /**
   * Vérifie si le token doit être rafraîchi
   * 
   * @returns {boolean} True si un rafraîchissement est nécessaire
   */
  needsTokenRefresh() {
    if (!this.tokenExpiry) return true;
    
    return Date.now() > (this.tokenExpiry - DEFAULT_CONFIG.SECURITY.TOKEN_REFRESH_THRESHOLD);
  }

  /**
   * Rafraîchit le token d'accès
   * 
   * @returns {Promise<string>} Nouveau token d'accès
   * @throws {Error} Si le rafraîchissement échoue
   */
  async refreshAccessToken() {
    if (!this.refreshToken) {
      throw new Error('Aucun token de rafraîchissement disponible');
    }

    try {
      console.log('🔄 Rafraîchissement du token d\'accès');
      
      const tokenUrl = DEFAULT_CONFIG.ENDPOINTS.TOKEN;
      const body = new URLSearchParams({
        client_id: this.clientId,
        scope: this.scopes.join(' '),
        refresh_token: this.refreshToken,
        grant_type: 'refresh_token'
      });

      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: body
      });

      if (!response.ok) {
        throw new Error(`Erreur lors du rafraîchissement du token: ${response.status}`);
      }

      const tokenResponse = await response.json();
      
      // Mettre à jour les tokens
      this.updateTokens(tokenResponse);
      
      // Déclencher l'événement
      this.triggerEvent('onTokenRefresh', this.accessToken);
      
      console.log('✅ Token rafraîchi avec succès');
      return this.accessToken;
      
    } catch (error) {
      console.error('❌ Erreur lors du rafraîchissement du token:', error);
      
      // Si le rafraîchissement échoue, déconnecter l'utilisateur
      this.logout();
      throw error;
    }
  }

  /**
   * Obtient un token d'accès valide
   * 
   * @returns {Promise<string>} Token d'accès valide
   * @throws {Error} Si l'utilisateur n'est pas authentifié
   */
  async getValidAccessToken() {
    if (!this.isAuthenticated()) {
      throw new Error('Utilisateur non authentifié');
    }

    if (this.needsTokenRefresh()) {
      await this.refreshAccessToken();
    }

    return this.accessToken;
  }

  /**
   * Déconnecte l'utilisateur
   */
  logout() {
    console.log('🚪 Déconnexion de l\'utilisateur');
    
    // Réinitialiser l'état
    this.accessToken = null;
    this.refreshToken = null;
    this.tokenExpiry = null;
    this.userInfo = null;
    
    // Nettoyer le stockage local
    this.clearSavedTokens();
    this.clearState();
    
    // Déclencher l'événement de déconnexion
    this.triggerEvent('onLogout');
    
    console.log('✅ Utilisateur déconnecté');
  }

  /**
   * Sauvegarde les tokens localement
   */
  saveTokens(tokenResponse) {
    this.accessToken = tokenResponse.access_token;
    this.refreshToken = tokenResponse.refresh_token;
    this.tokenExpiry = Date.now() + (tokenResponse.expires_in * 1000);
    
    this.persistTokens();
  }

  /**
   * Met à jour les tokens existants
   */
  updateTokens(tokenResponse) {
    this.accessToken = tokenResponse.access_token;
    this.refreshToken = tokenResponse.refresh_token || this.refreshToken;
    this.tokenExpiry = Date.now() + (tokenResponse.expires_in * 1000);
    
    this.persistTokens();
  }

  /**
   * Persiste les tokens dans le stockage local
   */
  persistTokens() {
    try {
      const tokens = {
        accessToken: this.accessToken,
        refreshToken: this.refreshToken,
        tokenExpiry: this.tokenExpiry,
        userInfo: this.userInfo
      };
      
      localStorage.setItem(DEFAULT_CONFIG.STORAGE_KEYS.TOKENS, JSON.stringify(tokens));
      console.log('💾 Tokens sauvegardés localement');
      
    } catch (error) {
      console.warn('⚠️ Impossible de sauvegarder les tokens:', error);
    }
  }

  /**
   * Charge les tokens sauvegardés
   */
  loadSavedTokens() {
    try {
      const saved = localStorage.getItem(DEFAULT_CONFIG.STORAGE_KEYS.TOKENS);
      if (saved) {
        const tokens = JSON.parse(saved);
        
        // Vérifier si les tokens sont encore valides
        if (tokens.tokenExpiry && Date.now() < tokens.tokenExpiry) {
          this.accessToken = tokens.accessToken;
          this.refreshToken = tokens.refreshToken;
          this.tokenExpiry = tokens.tokenExpiry;
          this.userInfo = tokens.userInfo;
          
          console.log('📥 Tokens restaurés depuis le stockage local');
        } else {
          // Tokens expirés, les supprimer
          console.log('⏰ Tokens expirés, suppression du stockage local');
          this.clearSavedTokens();
        }
      }
    } catch (error) {
      console.warn('⚠️ Erreur lors du chargement des tokens sauvegardés:', error);
      this.clearSavedTokens();
    }
  }

  /**
   * Supprime les tokens sauvegardés
   */
  clearSavedTokens() {
    try {
      localStorage.removeItem(DEFAULT_CONFIG.STORAGE_KEYS.TOKENS);
    } catch (error) {
      console.warn('⚠️ Erreur lors de la suppression des tokens:', error);
    }
  }

  /**
   * Sauvegarde l'état de sécurité
   */
  saveState(state) {
    try {
      sessionStorage.setItem(DEFAULT_CONFIG.STORAGE_KEYS.STATE, state);
    } catch (error) {
      console.warn('⚠️ Impossible de sauvegarder l\'état:', error);
    }
  }

  /**
   * Valide l'état de sécurité
   */
  validateState(state) {
    try {
      const savedState = sessionStorage.getItem(DEFAULT_CONFIG.STORAGE_KEYS.STATE);
      return state === savedState;
    } catch (error) {
      console.warn('⚠️ Erreur lors de la validation de l\'état:', error);
      return false;
    }
  }

  /**
   * Efface l'état de sécurité
   */
  clearState() {
    try {
      sessionStorage.removeItem(DEFAULT_CONFIG.STORAGE_KEYS.STATE);
    } catch (error) {
      console.warn('⚠️ Erreur lors de la suppression de l\'état:', error);
    }
  }

  /**
   * Génère un état sécurisé unique
   * 
   * @returns {string} État unique sécurisé
   */
  generateSecureState() {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    
    for (let i = 0; i < DEFAULT_CONFIG.SECURITY.STATE_LENGTH; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    
    return result;
  }

  /**
   * Vérifie si le service est initialisé
   * 
   * @returns {boolean} True si le service est initialisé
   */
  isInitialized() {
    return !!(this.clientId && this.redirectUri);
  }

  /**
   * Définit un gestionnaire d'événement
   * 
   * @param {string} eventName - Nom de l'événement
   * @param {Function} handler - Gestionnaire d'événement
   */
  on(eventName, handler) {
    if (this.eventHandlers.has(eventName)) {
      this.eventHandlers.set(eventName, handler);
      console.log(`📝 Gestionnaire d'événement défini pour: ${eventName}`);
    } else {
      console.warn(`⚠️ Événement inconnu: ${eventName}`);
    }
  }

  /**
   * Supprime un gestionnaire d'événement
   * 
   * @param {string} eventName - Nom de l'événement
   */
  off(eventName) {
    if (this.eventHandlers.has(eventName)) {
      this.eventHandlers.set(eventName, null);
      console.log(`🗑️ Gestionnaire d'événement supprimé pour: ${eventName}`);
    }
  }

  /**
   * Déclenche un événement
   * 
   * @param {string} eventName - Nom de l'événement
   * @param {*} data - Données de l'événement
   */
  triggerEvent(eventName, data) {
    const handler = this.eventHandlers.get(eventName);
    if (handler && typeof handler === 'function') {
      try {
        handler(data);
      } catch (error) {
        console.error(`❌ Erreur dans le gestionnaire d'événement ${eventName}:`, error);
      }
    }
  }

  /**
   * Obtient les informations de l'utilisateur actuel
   * 
   * @returns {Object|null} Informations utilisateur ou null si non connecté
   */
  getCurrentUser() {
    return this.userInfo;
  }

  /**
   * Obtient le statut de l'authentification
   * 
   * @returns {Object} Statut détaillé de l'authentification
   */
  getAuthStatus() {
    return {
      isAuthenticated: this.isAuthenticated(),
      needsRefresh: this.needsTokenRefresh(),
      tokenExpiry: this.tokenExpiry,
      hasUserInfo: !!this.userInfo,
      isInitialized: this.isInitialized()
    };
  }
}

// Export pour utilisation dans d'autres modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = MicrosoftGraphAuthService;
} else if (typeof window !== 'undefined') {
  window.MicrosoftGraphAuthService = MicrosoftGraphAuthService;
}
