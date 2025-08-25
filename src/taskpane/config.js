/**
 * Configuration centralisée de l'Office Add-in
 * 
 * Ce fichier contient toutes les constantes de configuration, les URLs des APIs,
 * et les paramètres globaux utilisés dans l'application.
 * 
 * @author Microsoft Corporation (modifié)
 * @license MIT
 * @version 1.0.0
 */

/**
 * Configuration de l'environnement
 */
const ENV = {
  // Déterminer l'environnement actuel
  IS_DEVELOPMENT: process.env.NODE_ENV === 'development' || !process.env.NODE_ENV,
  IS_PRODUCTION: process.env.NODE_ENV === 'production',
  
  // Variables d'environnement spécifiques
  DEBUG_MODE: process.env.DEBUG === 'true' || false,
  LOG_LEVEL: process.env.LOG_LEVEL || 'info'
};

/**
 * Configuration des APIs externes
 */
const API_CONFIG = {
  // OpenAI API
  OPENAI: {
    BASE_URL: 'https://api.openai.com/v1',
    ENDPOINTS: {
      CHAT_COMPLETIONS: '/chat/completions',
      MODELS: '/models'
    },
    DEFAULT_MODEL: 'gpt-4',
    MAX_TOKENS: 4000,
    TIMEOUT: 30000 // 30 secondes
  },
  
  // Microsoft Graph API
  GRAPH: {
    BASE_URL: 'https://graph.microsoft.com/v1.0',
    ENDPOINTS: {
      ME: '/me',
      CALENDAR_VIEW: '/me/calendarView',
      MAIL: '/me/messages',
      USERS: '/users'
    },
    TIMEOUT: 30000
  },
  
  // Microsoft Identity Platform
  IDENTITY: {
    AUTHORIZE_URL: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
    TOKEN_URL: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    SCOPES: [
      'openid',
      'profile',
      'email',
      'Mail.ReadWrite',
      'Mail.ReadWrite.Shared',
      'MailboxSettings.ReadWrite',
      'User.Read',
      'Calendars.Read',
      'Calendars.ReadWrite'
    ]
  }
};

/**
 * Configuration de l'interface utilisateur
 */
const UI_CONFIG = {
  // Dimensions des fenêtres de dialogue
  DIALOGS: {
    AUTH: {
      WIDTH: 50,
      HEIGHT: 70,
      DISPLAY_IN_IFRAME: false
    },
    CONSENT: {
      WIDTH: 60,
      HEIGHT: 80,
      DISPLAY_IN_IFRAME: false
    }
  },
  
  // Messages utilisateur
  MESSAGES: {
    LOADING: 'Chargement en cours...',
    SUCCESS: 'Opération réussie',
    ERROR: 'Une erreur s\'est produite',
    VALIDATION: {
      REQUIRED_FIELD: 'Ce champ est requis',
      INVALID_FORMAT: 'Format invalide',
      API_KEY_REQUIRED: 'Clé API requise'
    }
  },
  
  // Classes CSS
  CSS_CLASSES: {
    LOADING: 'loading',
    SUCCESS: 'success',
    ERROR: 'error',
    HIDDEN: 'hidden',
    VISIBLE: 'visible'
  }
};

/**
 * Configuration du stockage local
 */
const STORAGE_CONFIG = {
  // Clés de stockage
  KEYS: {
    OPENAI_API_KEY: 'openai_api_key',
    MSAL_TOKENS: 'msal_tokens',
    MSAL_STATE: 'msal_state',
    USER_PREFERENCES: 'user_preferences',
    LAST_SYNC: 'last_sync_timestamp'
  },
  
  // Configuration des tokens
  TOKENS: {
    REFRESH_THRESHOLD: 5 * 60 * 1000, // 5 minutes
    MAX_AGE: 24 * 60 * 60 * 1000 // 24 heures
  }
};

/**
 * Configuration des fonctionnalités
 */
const FEATURES = {
  // OpenAI
  OPENAI_ENABLED: true,
  AUTO_REPLY_ENABLED: true,
  
  // Microsoft Graph
  GRAPH_ENABLED: true,
  CALENDAR_SYNC_ENABLED: true,
  MAIL_SYNC_ENABLED: true,
  
  // Outlook
  OUTLOOK_FEATURES: {
    CATEGORIES_ENABLED: true,
    MAIL_CONTENT_EXTRACTION: true,
    AUTO_REPLY_GENERATION: true
  }
};

/**
 * Configuration des logs et débogage
 */
const LOGGING_CONFIG = {
  // Niveaux de log
  LEVELS: {
    DEBUG: 0,
    INFO: 1,
    WARN: 2,
    ERROR: 3
  },
  
  // Configuration des logs
  CONFIG: {
    ENABLE_CONSOLE: true,
    ENABLE_FILE: false,
    LOG_FILE: 'addin.log',
    MAX_LOG_SIZE: 10 * 1024 * 1024, // 10 MB
    MAX_LOG_FILES: 5
  },
  
  // Messages d'erreur personnalisés
  ERROR_MESSAGES: {
    NETWORK_ERROR: 'Erreur de réseau. Vérifiez votre connexion.',
    AUTH_ERROR: 'Erreur d\'authentification. Veuillez vous reconnecter.',
    API_ERROR: 'Erreur de l\'API. Veuillez réessayer plus tard.',
    PERMISSION_ERROR: 'Permissions insuffisantes pour cette opération.',
    TIMEOUT_ERROR: 'Délai d\'attente dépassé. Veuillez réessayer.'
  }
};

/**
 * Configuration de sécurité
 */
const SECURITY_CONFIG = {
  // Validation des entrées
  INPUT_VALIDATION: {
    MAX_LENGTH: {
      PROMPT: 10000,
      SUBJECT: 255,
      BODY: 100000
    },
    ALLOWED_CHARS: /^[a-zA-Z0-9\s\-_.,!?@#$%^&*()+=<>{}[\]|\\/:;"'`~]+$/,
    SANITIZE_HTML: true
  },
  
  // Protection CSRF
  CSRF: {
    ENABLED: true,
    TOKEN_LENGTH: 32,
    EXPIRY_TIME: 30 * 60 * 1000 // 30 minutes
  },
  
  // Rate limiting
  RATE_LIMITING: {
    ENABLED: true,
    MAX_REQUESTS: 100,
    WINDOW_MS: 15 * 60 * 1000 // 15 minutes
  }
};

/**
 * Configuration des performances
 */
const PERFORMANCE_CONFIG = {
  // Cache
  CACHE: {
    ENABLED: true,
    TTL: 5 * 60 * 1000, // 5 minutes
    MAX_SIZE: 100 // Nombre maximum d'éléments
  },
  
  // Debouncing
  DEBOUNCE: {
    SEARCH: 300, // 300ms
    API_CALLS: 1000, // 1 seconde
    UI_UPDATES: 100 // 100ms
  },
  
  // Lazy loading
  LAZY_LOADING: {
    ENABLED: true,
    THRESHOLD: 0.1 // 10% de la vue
  }
};

/**
 * Configuration des tests
 */
const TEST_CONFIG = {
  // Mode test
  ENABLED: process.env.NODE_ENV === 'test',
  
  // Mocks
  MOCKS: {
    OPENAI_API: false,
    GRAPH_API: false,
    OFFICE_API: false
  },
  
  // Timeouts
  TIMEOUTS: {
    UNIT_TEST: 5000,
    INTEGRATION_TEST: 30000,
    E2E_TEST: 120000
  }
};

/**
 * Configuration des localisations
 */
const LOCALIZATION_CONFIG = {
  // Langue par défaut
  DEFAULT_LANGUAGE: 'fr-FR',
  
  // Langues supportées
  SUPPORTED_LANGUAGES: ['fr-FR', 'en-US'],
  
  // Fallback
  FALLBACK_LANGUAGE: 'en-US',
  
  // Format des dates
  DATE_FORMAT: {
    SHORT: 'DD/MM/YYYY',
    LONG: 'DD MMMM YYYY',
    TIME: 'HH:mm',
    DATETIME: 'DD/MM/YYYY HH:mm'
  }
};

/**
 * Configuration des webhooks
 */
const WEBHOOK_CONFIG = {
  // Endpoints
  ENDPOINTS: {
    VALIDATION: '/graph/notifications',
    NOTIFICATIONS: '/graph/notifications'
  },
  
  // Configuration des notifications
  NOTIFICATIONS: {
    ENABLED: true,
    TYPES: ['created', 'updated', 'deleted'],
    RESOURCES: ['messages', 'events', 'contacts']
  },
  
  // Sécurité
  SECURITY: {
    VALIDATION_TOKEN_REQUIRED: true,
    MAX_PAYLOAD_SIZE: '10mb'
  }
};

/**
 * Configuration complète exportée
 */
const CONFIG = {
  ENV,
  API_CONFIG,
  UI_CONFIG,
  STORAGE_CONFIG,
  FEATURES,
  LOGGING_CONFIG,
  SECURITY_CONFIG,
  PERFORMANCE_CONFIG,
  TEST_CONFIG,
  LOCALIZATION_CONFIG,
  WEBHOOK_CONFIG,
  
  // Méthodes utilitaires
  isDevelopment: () => ENV.IS_DEVELOPMENT,
  isProduction: () => ENV.IS_PRODUCTION,
  isTest: () => TEST_CONFIG.ENABLED,
  
  // Validation de configuration
  validate: () => {
    const errors = [];
    
    // Vérifications de base
    if (!API_CONFIG.OPENAI.BASE_URL) {
      errors.push('URL de base OpenAI manquante');
    }
    
    if (!API_CONFIG.GRAPH.BASE_URL) {
      errors.push('URL de base Graph manquante');
    }
    
    if (errors.length > 0) {
      console.error('❌ Erreurs de configuration:', errors);
      return false;
    }
    
    console.log('✅ Configuration validée avec succès');
    return true;
  },
  
  // Obtenir la configuration pour un environnement spécifique
  getForEnvironment: (env) => {
    const envConfig = { ...CONFIG };
    
    if (env === 'production') {
      // Masquer les informations sensibles en production
      envConfig.ENV.DEBUG_MODE = false;
      envConfig.LOGGING_CONFIG.CONFIG.ENABLE_CONSOLE = false;
    }
    
    return envConfig;
  }
};

// Validation automatique au chargement
if (typeof window !== 'undefined') {
  CONFIG.validate();
}

// Export de la configuration
if (typeof module !== 'undefined' && module.exports) {
  module.exports = CONFIG;
} else if (typeof window !== 'undefined') {
  window.ADDIN_CONFIG = CONFIG;
}

export default CONFIG;
