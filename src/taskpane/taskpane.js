/**
 * Office Add-in Task Pane - Fichier principal
 * 
 * Ce fichier gère l'interface utilisateur et les interactions principales de l'add-in Outlook.
 * Il intègre les fonctionnalités OpenAI, Microsoft Graph et les actions Outlook.
 * 
 * @author Microsoft Corporation (modifié)
 * @license MIT
 * @version 1.0.0
 */

/* global document, Office */

/**
 * Configuration globale de l'application
 */
const APP_CONFIG = {
  OPENAI_API_URL: 'https://api.openai.com/v1/chat/completions',
  OPENAI_MODEL: 'gpt-4',
  GRAPH_API_BASE: 'https://graph.microsoft.com/v1.0',
  LOCAL_STORAGE_KEYS: {
    OPENAI_API_KEY: 'openai_api_key',
    MSAL_TOKENS: 'msal_tokens',
    MSAL_STATE: 'msal_state'
  },
  TOKEN_REFRESH_THRESHOLD: 5 * 60 * 1000 // 5 minutes
};

/**
 * Gestionnaire principal de l'add-in
 */
class TaskPaneManager {
  constructor() {
    this.elements = {};
    this.graphAuthService = null;
    this.isInitialized = false;
  }

  /**
   * Initialise l'add-in après le chargement d'Office
   */
  async initialize() {
    try {
      // Attendre que Office soit prêt
      await this.waitForOfficeReady();
      
      // Vérifier que nous sommes dans Outlook
      if (Office.context.host !== Office.HostType.Outlook) {
        console.warn('Cet add-in est conçu pour Outlook uniquement');
        return;
      }

      // Initialiser l'interface utilisateur
      this.initializeUI();
      
      // Configurer les gestionnaires d'événements
      this.setupEventHandlers();
      
      // Initialiser le service d'authentification Graph
      this.initializeGraphAuth();
      
      this.isInitialized = true;
      console.log('✅ Add-in initialisé avec succès');
      
    } catch (error) {
      console.error('❌ Erreur lors de l\'initialisation:', error);
    }
  }

  /**
   * Attend que l'API Office soit prête
   */
  waitForOfficeReady() {
    return new Promise((resolve) => {
      Office.onReady((info) => {
        console.log('Office prêt:', info);
        resolve(info);
      });
    });
  }

  /**
   * Initialise l'interface utilisateur
   */
  initializeUI() {
    // Masquer le message de sideload et afficher l'interface principale
    const sideloadMsg = document.getElementById('sideload-msg');
    const appBody = document.getElementById('app-body');
    
    if (sideloadMsg) sideloadMsg.style.display = 'none';
    if (appBody) appBody.style.display = 'flex';

    // Récupérer tous les éléments DOM nécessaires
    this.elements = {
      apiKeyInput: document.getElementById('api-key'),
      promptInput: document.getElementById('prompt-input'),
      sendBtn: document.getElementById('send-prompt'),
      addMailBtn: document.getElementById('add-mail-content'),
      addCategoriesBtn: document.getElementById('add-mail-categories'),
      output: document.getElementById('response-output'),
      autoReplyBtn: document.getElementById('auto-reply'),
      listEventsBtn: document.getElementById('list-events'),
      msalClientIdInput: document.getElementById('msal-client-id'),
      eventsOutput: document.getElementById('events-output'),
      runBtn: document.getElementById('run'),
      itemSubject: document.getElementById('item-subject')
    };

    // Charger les données sauvegardées
    this.loadSavedData();
  }

  /**
   * Charge les données sauvegardées localement
   */
  loadSavedData() {
    try {
      // Charger la clé API OpenAI si disponible
      const savedApiKey = localStorage.getItem(APP_CONFIG.LOCAL_STORAGE_KEYS.OPENAI_API_KEY);
      if (savedApiKey && this.elements.apiKeyInput) {
        this.elements.apiKeyInput.value = savedApiKey;
      }
    } catch (error) {
      console.warn('Impossible de charger les données sauvegardées:', error);
    }
  }

  /**
   * Configure tous les gestionnaires d'événements
   */
  setupEventHandlers() {
    // Gestionnaire pour l'envoi de prompts OpenAI
    if (this.elements.sendBtn) {
      this.elements.sendBtn.addEventListener('click', () => this.handleOpenAIPrompt());
    }

    // Gestionnaire pour la réponse automatique
    if (this.elements.autoReplyBtn) {
      this.elements.autoReplyBtn.addEventListener('click', () => this.handleAutoReply());
    }

    // Gestionnaire pour lister les événements
    if (this.elements.listEventsBtn) {
      this.elements.listEventsBtn.addEventListener('click', () => this.handleListEvents());
    }

    // Gestionnaire pour ajouter le contenu du mail
    if (this.elements.addMailBtn) {
      this.elements.addMailBtn.addEventListener('click', () => this.handleAddMailContent());
    }

    // Gestionnaire pour ajouter les catégories
    if (this.elements.addCategoriesBtn) {
      this.elements.addCategoriesBtn.addEventListener('click', () => this.handleAddCategories());
    }

    // Gestionnaire pour afficher le sujet
    if (this.elements.runBtn) {
      this.elements.runBtn.addEventListener('click', () => this.handleShowSubject());
    }
  }

  /**
   * Initialise le service d'authentification Microsoft Graph
   */
  initializeGraphAuth() {
    try {
      if (typeof MicrosoftGraphAuthService !== 'undefined') {
        this.graphAuthService = new MicrosoftGraphAuthService();
        
        // Configurer les événements du service
        this.graphAuthService.on('onAuthSuccess', (userInfo) => {
          console.log('✅ Authentification Graph réussie:', userInfo);
          this.updateEventsOutput(`Connecté en tant que: ${userInfo.displayName || userInfo.userPrincipalName}`);
        });
        
        this.graphAuthService.on('onAuthError', (error) => {
          console.error('❌ Erreur d\'authentification Graph:', error);
          this.updateEventsOutput(`Erreur d'authentification: ${error.message}`);
        });
        
        this.graphAuthService.on('onTokenRefresh', () => {
          console.log('🔄 Token Graph rafraîchi avec succès');
        });
        
        console.log('✅ Service d\'authentification Graph initialisé');
      }
    } catch (error) {
      console.warn('⚠️ Impossible d\'initialiser le service Graph:', error);
    }
  }

  /**
   * Gère l'envoi de prompts vers OpenAI
   */
  async handleOpenAIPrompt() {
    const apiKey = this.elements.apiKeyInput?.value?.trim();
    const prompt = this.elements.promptInput?.value?.trim();

    if (!this.validateOpenAIInput(apiKey, prompt)) {
      return;
    }

    try {
      // Sauvegarder la clé API localement
      this.saveApiKey(apiKey);
      
      // Afficher le statut
      this.updateOutput('Envoi en cours...');
      
      // Appeler l'API OpenAI
      const response = await this.callOpenAI(apiKey, prompt);
      
      // Afficher la réponse
      this.updateOutput(response);
      
    } catch (error) {
      this.updateOutput(`Erreur: ${error.message}`);
    }
  }

  /**
   * Valide les entrées OpenAI
   */
  validateOpenAIInput(apiKey, prompt) {
    if (!apiKey || !prompt) {
      this.updateOutput('Veuillez saisir la clé API et un prompt.');
      return false;
    }
    return true;
  }

  /**
   * Sauvegarde la clé API localement
   */
  saveApiKey(apiKey) {
    try {
      localStorage.setItem(APP_CONFIG.LOCAL_STORAGE_KEYS.OPENAI_API_KEY, apiKey);
    } catch (error) {
      console.warn('Impossible de sauvegarder la clé API:', error);
    }
  }

  /**
   * Appelle l'API OpenAI
   */
  async callOpenAI(apiKey, prompt) {
    const response = await fetch(APP_CONFIG.OPENAI_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: APP_CONFIG.OPENAI_MODEL,
        messages: [
          { role: 'system', content: 'You are a helpful assistant.' },
          { role: 'user', content: prompt },
        ],
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Erreur API (${response.status}): ${errorText}`);
    }

    const data = await response.json();
    return data?.choices?.[0]?.message?.content || '(Réponse vide)';
  }

  /**
   * Gère la génération de réponse automatique
   */
  async handleAutoReply() {
    const apiKey = this.elements.apiKeyInput?.value?.trim();
    
    if (!apiKey) {
      this.updateOutput('Veuillez saisir la clé API avant de générer la réponse.');
      return;
    }

    try {
      this.updateOutput('Génération de la réponse...');
      
      // Récupérer le contenu du mail
      const mailContent = await this.getMailContent();
      
      // Générer la réponse via OpenAI
      const prompt = this.buildAutoReplyPrompt(mailContent.subject, mailContent.body);
      const response = await this.callOpenAI(apiKey, prompt);
      
      // Ouvrir le formulaire de réponse
      await this.openReplyForm(response);
      
      this.updateOutput('Réponse générée et placée dans un brouillon de réponse.');
      
    } catch (error) {
      this.updateOutput(`Erreur lors de la réponse automatique: ${error.message}`);
    }
  }

  /**
   * Récupère le contenu du mail actuel
   */
  async getMailContent() {
    const item = Office.context.mailbox.item;
    const subject = item?.subject || '';
    
    let bodyText = '';
    await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          bodyText = result.value || '';
          resolve();
        } else {
          reject(result.error);
        }
      });
    });

    return { subject, body: bodyText };
  }

  /**
   * Construit le prompt pour la réponse automatique
   */
  buildAutoReplyPrompt(subject, body) {
    return `Tu es un assistant qui rédige des réponses d'email courtes, polies et en français.

Sujet: ${subject}

Contenu:
${body}

Rédige une réponse appropriée en gardant un ton professionnel, commence par un salut et termine par une formule de politesse.`;
  }

  /**
   * Ouvre le formulaire de réponse avec le texte généré
   */
  async openReplyForm(draftText) {
    const item = Office.context.mailbox.item;
    
    await new Promise((resolve, reject) => {
      item.displayReplyForm(draftText, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(asyncResult.error);
        }
      });
    });
  }

  /**
   * Gère la liste des événements via Microsoft Graph
   */
  async handleListEvents() {
    const clientId = this.elements.msalClientIdInput?.value?.trim();
    
    if (!clientId) {
      this.updateEventsOutput('Veuillez saisir un Client ID (Azure AD) pour appeler Microsoft Graph.');
      return;
    }

    this.updateEventsOutput('Authentification Microsoft Graph...');
    
    try {
      const accessToken = await this.authenticateWithGraph(clientId);
      const events = await this.fetchCalendarEvents(accessToken);
      this.displayEvents(events);
      
    } catch (error) {
      this.updateEventsOutput(`Erreur: ${error.message}`);
    }
  }

  /**
   * Authentifie l'utilisateur avec Microsoft Graph
   */
  async authenticateWithGraph(clientId) {
    return new Promise((resolve, reject) => {
      this.openGraphAuthDialog(clientId, (accessToken, error) => {
        if (error) {
          reject(error);
        } else {
          resolve(accessToken);
        }
      });
    });
  }

  /**
   * Ouvre la fenêtre d'authentification Graph
   */
  openGraphAuthDialog(clientId, onResult) {
    if (!this.graphAuthService) {
      onResult(null, new Error('Service d\'authentification non disponible'));
      return;
    }

    try {
      const redirectUri = `${location.origin}/graph-callback.html`;
      this.graphAuthService.initialize(clientId, redirectUri);
      
      const authUrl = this.graphAuthService.startAuth();
      
      Office.context.ui.displayDialogAsync(
        authUrl,
        { height: 70, width: 50, displayInIframe: false },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            onResult(null, new Error('Impossible d\'ouvrir la fenêtre d\'authentification.'));
            return;
          }

          this.handleGraphAuthDialog(result.value, onResult);
        }
      );
      
    } catch (error) {
      onResult(null, error);
    }
  }

  /**
   * Gère la fenêtre d'authentification Graph
   */
  handleGraphAuthDialog(dialog, onResult) {
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
      try {
        const msg = JSON.parse(arg.message || '{}');
        
        if (msg.type === 'auth_success') {
          try {
            const userInfo = await this.graphAuthService.processAuthResponse(msg.code, msg.state);
            const accessToken = await this.graphAuthService.getValidAccessToken();
            
            dialog.close();
            onResult(accessToken, null);
            
          } catch (error) {
            dialog.close();
            onResult(null, error);
          }
        } else if (msg.type === 'close_auth_window') {
          dialog.close();
        }
      } catch (error) {
        console.error('Erreur lors du traitement du message:', error);
        dialog.close();
        onResult(null, error);
      }
    });

    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      if (arg.errorCode) {
        onResult(null, new Error(`Erreur de dialogue: ${arg.errorCode}`));
      }
    });
  }

  /**
   * Récupère les événements du calendrier
   */
  async fetchCalendarEvents(accessToken) {
    const now = new Date();
    const end = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // +7 jours
    
    const url = `${APP_CONFIG.GRAPH_API_BASE}/me/calendarView?startDateTime=${encodeURIComponent(now.toISOString())}&endDateTime=${encodeURIComponent(end.toISOString())}&$orderby=start/dateTime&$top=10`;
    
    const response = await fetch(url, { 
      headers: { Authorization: `Bearer ${accessToken}` } 
    });
    
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Erreur Graph (${response.status}): ${errorText}`);
    }

    const data = await response.json();
    return data?.value || [];
  }

  /**
   * Affiche les événements récupérés
   */
  displayEvents(events) {
    if (events.length === 0) {
      this.updateEventsOutput('Aucun rendez-vous à venir dans les 7 prochains jours.');
      return;
    }

    const lines = events.map((event) => {
      const subject = event.subject || '(Sans objet)';
      const start = event.start?.dateTime || '';
      const end = event.end?.dateTime || '';
      const location = event.location?.displayName ? ` @ ${event.location.displayName}` : '';
      return `- ${subject} (${start} -> ${end})${location}`;
    });

    this.updateEventsOutput(lines.join('\n'));
  }

  /**
   * Gère l'ajout du contenu du mail au prompt
   */
  async handleAddMailContent() {
    try {
      const mailContent = await this.getMailContent();
      const separator = this.elements.promptInput?.value ? '\n\n' : '';
      
      this.elements.promptInput.value = `${this.elements.promptInput.value || ''}${separator}Sujet: ${mailContent.subject}\n\nCorps:\n${mailContent.body}`.trim();
      
    } catch (error) {
      this.updateOutput(`Erreur lors de l'extraction du mail: ${error.message}`);
    }
  }

  /**
   * Gère l'ajout des catégories du mail
   */
  async handleAddCategories() {
    try {
      const item = Office.context.mailbox.item;
      const categories = item?.categories || [];
      
      // Récupérer les catégories utilisateur si disponible
      let userCategories = [];
      try {
        const requirements = Office.context.requirements;
        const hasApi = requirements && requirements.isSetSupported && requirements.isSetSupported('Mailbox', '1.8');
        
        if (hasApi && Office.context.mailbox?.masterCategories?.getAsync) {
          await new Promise((resolve) => {
            Office.context.mailbox.masterCategories.getAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                userCategories = (result.value || []).map((cat) => cat.displayName).filter(Boolean);
              }
              resolve();
            });
          });
        }
      } catch (error) {
        console.warn('Impossible de récupérer les catégories utilisateur:', error);
      }

      const allCategories = [...new Set([...categories, ...userCategories])];
      
      if (allCategories.length === 0) {
        this.updateOutput('Aucune catégorie trouvée (l\'API Catégories nécessite Outlook supportant Mailbox 1.8).');
        return;
      }

      const toAppend = `\n\nCatégories Outlook: ${allCategories.join(', ')}`;
      this.elements.promptInput.value = `${this.elements.promptInput.value || ''}${toAppend}`.trim();
      
    } catch (error) {
      this.updateOutput(`Erreur lors de la récupération des catégories: ${error.message}`);
    }
  }

  /**
   * Gère l'affichage du sujet du mail
   */
  handleShowSubject() {
    const item = Office.context.mailbox.item;
    const subject = item?.subject || '(Sans objet)';
    
    if (this.elements.itemSubject) {
      this.elements.itemSubject.innerHTML = `
        <b>Sujet: </b>
        <br>
        ${subject}
        <br>
      `;
    }
  }

  /**
   * Met à jour la sortie principale
   */
  updateOutput(text) {
    if (this.elements.output) {
      this.elements.output.textContent = text;
    }
  }

  /**
   * Met à jour la sortie des événements
   */
  updateEventsOutput(text) {
    if (this.elements.eventsOutput) {
      this.elements.eventsOutput.textContent = text;
    }
  }
}

/**
 * Point d'entrée principal de l'add-in
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const taskPaneManager = new TaskPaneManager();
    taskPaneManager.initialize();
  }
});

/**
 * Fonction d'export pour compatibilité avec l'ancien code
 * @deprecated Utiliser TaskPaneManager à la place
 */
export async function run() {
  console.warn('La fonction run() est dépréciée. Utilisez TaskPaneManager.');
  
  const item = Office.context.mailbox.item;
  const insertAt = document.getElementById('item-subject');
  
  if (insertAt && item) {
    const label = document.createElement('b');
    label.appendChild(document.createTextNode('Subject: '));
    insertAt.appendChild(label);
    insertAt.appendChild(document.createElement('br'));
    insertAt.appendChild(document.createTextNode(item.subject));
    insertAt.appendChild(document.createElement('br'));
  }
}
