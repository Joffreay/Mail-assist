/**
 * Serveur Webhook Microsoft Graph
 * 
 * Ce serveur gère les notifications et validations des webhooks Microsoft Graph.
 * Il implémente le protocole de validation et traite les notifications de changement.
 * 
 * @author Microsoft Corporation (modifié)
 * @license MIT
 * @version 1.0.0
 */

import express from 'express';
import bodyParser from 'body-parser';

/**
 * Configuration du serveur
 */
const SERVER_CONFIG = {
  PORT: process.env.PORT || 8080,
  ENDPOINTS: {
    WEBHOOK: '/graph/notifications',
    HEALTH: '/health'
  },
  LOG_LEVELS: {
    INFO: 'INFO',
    WARN: 'WARN',
    ERROR: 'ERROR'
  }
};

/**
 * Classe principale du serveur webhook
 */
class WebhookServer {
  constructor() {
    this.app = express();
    this.isRunning = false;
    
    // Initialiser le serveur
    this.initialize();
  }

  /**
   * Initialise le serveur Express
   */
  initialize() {
    // Middleware pour parser le JSON
    this.app.use(bodyParser.json({ limit: '10mb' }));
    
    // Middleware pour parser les données URL-encoded
    this.app.use(bodyParser.urlencoded({ extended: true, limit: '10mb' }));
    
    // Middleware de logging
    this.app.use(this.loggingMiddleware.bind(this));
    
    // Configuration des routes
    this.setupRoutes();
    
    // Middleware de gestion d'erreurs
    this.app.use(this.errorHandler.bind(this));
    
    console.log('🔧 Serveur webhook initialisé');
  }

  /**
   * Middleware de logging personnalisé
   */
  loggingMiddleware(req, res, next) {
    const timestamp = new Date().toISOString();
    const method = req.method;
    const url = req.url;
    const userAgent = req.get('User-Agent') || 'Unknown';
    
    console.log(`[${timestamp}] ${method} ${url} - ${userAgent}`);
    
    // Ajouter des informations de requête pour le débogage
    req.requestId = this.generateRequestId();
    req.startTime = Date.now();
    
    next();
  }

  /**
   * Configure les routes du serveur
   */
  setupRoutes() {
    // Route de santé du serveur
    this.app.get(SERVER_CONFIG.ENDPOINTS.HEALTH, this.healthCheck.bind(this));
    
    // Route de validation des webhooks (GET)
    this.app.get(SERVER_CONFIG.ENDPOINTS.WEBHOOK, this.handleWebhookValidation.bind(this));
    
    // Route de réception des notifications (POST)
    this.app.post(SERVER_CONFIG.ENDPOINTS.WEBHOOK, this.handleWebhookNotification.bind(this));
    
    // Route par défaut pour les requêtes non reconnues
    this.app.use('*', this.handleNotFound.bind(this));
    
    console.log('🛣️ Routes configurées');
  }

  /**
   * Gère la validation des webhooks Microsoft Graph
   * 
   * @param {Object} req - Objet de requête Express
   * @param {Object} res - Objet de réponse Express
   */
  handleWebhookValidation(req, res) {
    try {
      const validationToken = req.query.validationToken;
      
      if (!validationToken) {
        this.logRequest(SERVER_CONFIG.LOG_LEVELS.WARN, 'Validation token manquant', req);
        return res.status(400).json({
          error: 'Missing validationToken',
          message: 'Le paramètre validationToken est requis pour la validation des webhooks'
        });
      }

      this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, `Validation webhook: ${validationToken}`, req);
      
      // Retourner le token de validation (requis par Microsoft Graph)
      res.status(200)
        .set('Content-Type', 'text/plain')
        .send(validationToken);
        
    } catch (error) {
      this.logRequest(SERVER_CONFIG.LOG_LEVELS.ERROR, `Erreur validation: ${error.message}`, req);
      res.status(500).json({
        error: 'Internal Server Error',
        message: 'Erreur lors de la validation du webhook'
      });
    }
  }

  /**
   * Gère la réception des notifications de changement
   * 
   * @param {Object} req - Objet de requête Express
   * @param {Object} res - Objet de réponse Express
   */
  async handleWebhookNotification(req, res) {
    try {
      const requestBody = req.body;
      const contentType = req.get('Content-Type');
      
      this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, 'Notification webhook reçue', req, {
        contentType,
        bodySize: JSON.stringify(requestBody).length
      });

      // Vérifier que le corps de la requête est présent
      if (!requestBody || Object.keys(requestBody).length === 0) {
        this.logRequest(SERVER_CONFIG.LOG_LEVELS.WARN, 'Corps de notification vide', req);
        return res.status(400).json({
          error: 'Empty notification body',
          message: 'Le corps de la notification est vide'
        });
      }

      // Traiter la notification selon son type
      await this.processNotification(requestBody, req);

      // Répondre avec un statut 202 (Accepted) comme requis par Microsoft Graph
      res.status(202).json({
        status: 'accepted',
        message: 'Notification traitée avec succès',
        timestamp: new Date().toISOString()
      });

    } catch (error) {
      this.logRequest(SERVER_CONFIG.LOG_LEVELS.ERROR, `Erreur notification: ${error.message}`, req);
      
      // En cas d'erreur, on répond quand même avec 202 pour éviter les retry
      // Microsoft Graph retry automatiquement en cas d'échec
      res.status(202).json({
        status: 'accepted_with_errors',
        error: error.message,
        timestamp: new Date().toISOString()
      });
    }
  }

  /**
   * Traite une notification reçue
   * 
   * @param {Object} notification - Données de la notification
   * @param {Object} req - Objet de requête Express
   */
  async processNotification(notification, req) {
    try {
      // Extraire les informations de base de la notification
      const { value, validationToken } = notification;
      
      if (validationToken) {
        this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, `Token de validation reçu: ${validationToken}`, req);
        return;
      }

      if (value && Array.isArray(value)) {
        // Traiter chaque changement individuellement
        for (const change of value) {
          await this.processChange(change, req);
        }
      } else {
        this.logRequest(SERVER_CONFIG.LOG_LEVELS.WARN, 'Format de notification non reconnu', req, {
          notificationKeys: Object.keys(notification)
        });
      }

    } catch (error) {
      this.logRequest(SERVER_CONFIG.LOG_LEVELS.ERROR, `Erreur traitement notification: ${error.message}`, req);
      throw error;
    }
  }

  /**
   * Traite un changement individuel
   * 
   * @param {Object} change - Données du changement
   * @param {Object} req - Objet de requête Express
   */
  async processChange(change, req) {
    try {
      const { changeType, resource, resourceData, tenantId, subscriptionId } = change;
      
      this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, `Changement détecté: ${changeType}`, req, {
        resource,
        subscriptionId,
        tenantId
      });

      // Traiter selon le type de changement
      switch (changeType) {
        case 'created':
          await this.handleResourceCreated(resource, resourceData, req);
          break;
          
        case 'updated':
          await this.handleResourceUpdated(resource, resourceData, req);
          break;
          
        case 'deleted':
          await this.handleResourceDeleted(resource, req);
          break;
          
        default:
          this.logRequest(SERVER_CONFIG.LOG_LEVELS.WARN, `Type de changement non géré: ${changeType}`, req);
      }

    } catch (error) {
      this.logRequest(SERVER_CONFIG.LOG_LEVELS.ERROR, `Erreur traitement changement: ${error.message}`, req);
      // Ne pas faire échouer le traitement des autres changements
    }
  }

  /**
   * Gère la création d'une ressource
   * 
   * @param {string} resource - URI de la ressource
   * @param {Object} resourceData - Données de la ressource
   * @param {Object} req - Objet de requête Express
   */
  async handleResourceCreated(resource, resourceData, req) {
    this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, `Ressource créée: ${resource}`, req, {
      resourceData: resourceData ? Object.keys(resourceData) : 'N/A'
    });
    
    // TODO: Implémenter la logique métier pour la création
    // Exemple: synchroniser avec une base de données, notifier d'autres services, etc.
  }

  /**
   * Gère la mise à jour d'une ressource
   * 
   * @param {string} resource - URI de la ressource
   * @param {Object} resourceData - Données de la ressource
   * @param {Object} req - Objet de requête Express
   */
  async handleResourceUpdated(resource, resourceData, req) {
    this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, `Ressource mise à jour: ${resource}`, req, {
      resourceData: resourceData ? Object.keys(resourceData) : 'N/A'
    });
    
    // TODO: Implémenter la logique métier pour la mise à jour
  }

  /**
   * Gère la suppression d'une ressource
   * 
   * @param {string} resource - URI de la ressource
   * @param {Object} req - Objet de requête Express
   */
  async handleResourceDeleted(resource, req) {
    this.logRequest(SERVER_CONFIG.LOG_LEVELS.INFO, `Ressource supprimée: ${resource}`, req);
    
    // TODO: Implémenter la logique métier pour la suppression
  }

  /**
   * Route de vérification de santé du serveur
   * 
   * @param {Object} req - Objet de requête Express
   * @param {Object} res - Objet de réponse Express
   */
  healthCheck(req, res) {
    const healthStatus = {
      status: 'healthy',
      timestamp: new Date().toISOString(),
      uptime: process.uptime(),
      memory: process.memoryUsage(),
      version: '1.0.0'
    };

    res.status(200).json(healthStatus);
  }

  /**
   * Gère les routes non trouvées
   * 
   * @param {Object} req - Objet de requête Express
   * @param {Object} res - Objet de réponse Express
   */
  handleNotFound(req, res) {
    this.logRequest(SERVER_CONFIG.LOG_LEVELS.WARN, `Route non trouvée: ${req.method} ${req.originalUrl}`, req);
    
    res.status(404).json({
      error: 'Not Found',
      message: 'La route demandée n\'existe pas',
      path: req.originalUrl,
      method: req.method
    });
  }

  /**
   * Middleware de gestion d'erreurs global
   * 
   * @param {Error} error - Erreur capturée
   * @param {Object} req - Objet de requête Express
   * @param {Object} res - Objet de réponse Express
   * @param {Function} next - Fonction next Express
   */
  errorHandler(error, req, res, next) {
    this.logRequest(SERVER_CONFIG.LOG_LEVELS.ERROR, `Erreur serveur: ${error.message}`, req, {
      stack: error.stack
    });

    // Ne pas exposer les détails de l'erreur en production
    const errorResponse = {
      error: 'Internal Server Error',
      message: 'Une erreur interne s\'est produite',
      timestamp: new Date().toISOString()
    };

    // En développement, ajouter plus de détails
    if (process.env.NODE_ENV === 'development') {
      errorResponse.details = error.message;
      errorResponse.stack = error.stack;
    }

    res.status(500).json(errorResponse);
  }

  /**
   * Génère un ID unique pour chaque requête
   * 
   * @returns {string} ID unique de la requête
   */
  generateRequestId() {
    return `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Enregistre une requête avec des informations détaillées
   * 
   * @param {string} level - Niveau de log
   * @param {string} message - Message à logger
   * @param {Object} req - Objet de requête Express
   * @param {Object} [additionalData] - Données supplémentaires à logger
   */
  logRequest(level, message, req, additionalData = {}) {
    const logData = {
      timestamp: new Date().toISOString(),
      level,
      message,
      requestId: req.requestId,
      method: req.method,
      url: req.url,
      userAgent: req.get('User-Agent'),
      ip: req.ip || req.connection.remoteAddress,
      ...additionalData
    };

    // Calculer le temps de traitement si disponible
    if (req.startTime) {
      logData.processingTime = Date.now() - req.startTime;
    }

    // Formater le log selon le niveau
    switch (level) {
      case SERVER_CONFIG.LOG_LEVELS.ERROR:
        console.error('❌', JSON.stringify(logData, null, 2));
        break;
      case SERVER_CONFIG.LOG_LEVELS.WARN:
        console.warn('⚠️', JSON.stringify(logData, null, 2));
        break;
      default:
        console.log('ℹ️', JSON.stringify(logData, null, 2));
    }
  }

  /**
   * Démarre le serveur
   */
  async start() {
    try {
      await new Promise((resolve, reject) => {
        this.server = this.app.listen(SERVER_CONFIG.PORT, (error) => {
          if (error) {
            reject(error);
          } else {
            resolve();
          }
        });
      });

      this.isRunning = true;
      
      console.log(`✅ Serveur webhook démarré sur http://localhost:${SERVER_CONFIG.PORT}`);
      console.log(`🔗 Endpoint webhook: http://localhost:${SERVER_CONFIG.PORT}${SERVER_CONFIG.ENDPOINTS.WEBHOOK}`);
      console.log(`💚 Endpoint santé: http://localhost:${SERVER_CONFIG.PORT}${SERVER_CONFIG.ENDPOINTS.HEALTH}`);
      
    } catch (error) {
      console.error('❌ Erreur lors du démarrage du serveur:', error);
      throw error;
    }
  }

  /**
   * Arrête le serveur
   */
  async stop() {
    if (this.server && this.isRunning) {
      try {
        await new Promise((resolve) => {
          this.server.close(resolve);
        });
        
        this.isRunning = false;
        console.log('🛑 Serveur webhook arrêté');
        
      } catch (error) {
        console.error('❌ Erreur lors de l\'arrêt du serveur:', error);
        throw error;
      }
    }
  }

  /**
   * Obtient le statut du serveur
   * 
   * @returns {Object} Statut du serveur
   */
  getStatus() {
    return {
      isRunning: this.isRunning,
      port: SERVER_CONFIG.PORT,
      uptime: this.isRunning ? process.uptime() : 0,
      endpoints: SERVER_CONFIG.ENDPOINTS
    };
  }
}

// Créer et démarrer le serveur
const webhookServer = new WebhookServer();

// Gestion propre de l'arrêt du serveur
process.on('SIGINT', async () => {
  console.log('\n🔄 Arrêt du serveur...');
  await webhookServer.stop();
  process.exit(0);
});

process.on('SIGTERM', async () => {
  console.log('\n🔄 Arrêt du serveur...');
  await webhookServer.stop();
  process.exit(0);
});

// Démarrer le serveur
webhookServer.start().catch((error) => {
  console.error('❌ Impossible de démarrer le serveur:', error);
  process.exit(1);
});

export default webhookServer;
