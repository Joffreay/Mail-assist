/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Enregistrer la fonction pour appliquer la catégorie Orange
    Office.actions.associate("applyOrangeCategory", applyOrangeCategory);
  }
});

/**
 * Applique la catégorie Orange au mail sélectionné
 * @param {Office.AddinCommands.Event} event - L'événement de la commande
 */
async function applyOrangeCategory(event) {
  try {
    // Vérifier que nous avons accès à un élément de mail
    if (!Office.context.mailbox.item) {
      Office.context.ui.messageParent("Aucun mail sélectionné");
      return;
    }

    // Appliquer la catégorie Orange
    await Office.context.mailbox.item.categories.addAsync("Orange");
    
    // Notification de succès
    Office.context.ui.messageParent("Catégorie Orange appliquée avec succès !");
    
    // Marquer l'événement comme terminé
    event.completed();
    
  } catch (error) {
    console.error("Erreur lors de l'application de la catégorie:", error);
    Office.context.ui.messageParent("Erreur lors de l'application de la catégorie: " + error.message);
    event.completed();
  }
}

/**
 * Génère une réponse automatique à l'email actuel
 * @param event {Office.AddinCommands.Event}
 */
async function generateAutoReply(event) {
  try {
    // Récupérer les informations de l'email
    const item = Office.context.mailbox.item;
    
    // Vérifier si un email est sélectionné
    if (!item) {
      // Pas d'email sélectionné (page d'accueil) - ouvrir un nouveau message
      const apiKey = await promptForApiKey();
      if (!apiKey) {
        event.completed();
        return;
      }
      
      // Créer un nouveau message avec une réponse générique
      const prompt = `Tu es un assistant qui rédige des réponses d'email courtes, polies et en français.\n\nRédige un modèle de réponse d'email professionnel et poli en français, commence par un salut et termine par une formule de politesse.`;
      
      const resp = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model: "gpt-5",
          messages: [
            { role: "system", content: "You are a helpful assistant." },
            { role: "user", content: prompt },
          ],
          temperature: 0.6,
        }),
      });

      if (!resp.ok) {
        const errText = await resp.text();
        event.completed();
        return;
      }

      const data = await resp.json();
      const draft = data?.choices?.[0]?.message?.content || "";
      
      // Créer un nouveau message avec le contenu généré
      Office.context.mailbox.displayNewMessageForm({
        subject: "Nouveau message",
        toRecipients: [],
        htmlBody: draft
      });
      
      event.completed();
      return;
    }

    // Afficher un message de notification
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Génération de la réponse automatique...",
      icon: "Icon.80x80",
      persistent: false,
    };

    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "AutoReplyNotification",
      message
    );

    const subject = item.subject || "";
    
    // Récupérer le contenu du corps de l'email
    let bodyText = "";
    await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          bodyText = res.value || "";
          resolve();
        } else {
          reject(res.error);
        }
      });
    });

    // Demander la clé API à l'utilisateur via une boîte de dialogue
    const apiKey = await promptForApiKey();
    if (!apiKey) {
      updateNotification("Réponse automatique annulée - Aucune clé API fournie", "error");
      event.completed();
      return;
    }

    // Générer la réponse via OpenAI
    const prompt = `Tu es un assistant qui rédige des réponses d'email courtes, polies et en français.\n\nSujet: ${subject}\n\nContenu:\n${bodyText}\n\nRédige une réponse appropriée en gardant un ton professionnel, commence par un salut et termine par une formule de politesse.`;

    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "gpt-5",
        messages: [
          { role: "system", content: "You are a helpful assistant." },
          { role: "user", content: prompt },
        ],
        temperature: 0.6,
      }),
    });

    if (!resp.ok) {
      const errText = await resp.text();
      updateNotification(`Erreur API (${resp.status}): ${errText}`, "error");
      event.completed();
      return;
    }

    const data = await resp.json();
    const draft = data?.choices?.[0]?.message?.content || "";

    // Vérifier si nous sommes en mode composition ou lecture
    const itemType = Office.context.mailbox.item.itemType;
    
    if (itemType === Office.MailboxEnums.ItemType.MessageCompose) {
      // En mode composition, insérer directement dans le corps du message
      await new Promise((resolve, reject) => {
        item.body.setAsync(draft, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(asyncResult.error);
          }
        });
      });
      updateNotification("Réponse automatique insérée dans le message en cours de rédaction !", "success");
    } else {
      // En mode lecture, ouvrir une réponse pré-remplie
      await new Promise((resolve, reject) => {
        item.displayReplyForm(draft, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(asyncResult.error);
          }
        });
      });
      updateNotification("Réponse automatique générée et placée dans un brouillon de réponse.", "success");
    }
    
  } catch (error) {
    updateNotification(`Erreur lors de la génération de la réponse: ${error?.message || error}`, "error");
  }

  // Indiquer que la fonction est terminée
  event.completed();
}

/**
 * Demande la clé API à l'utilisateur via une boîte de dialogue
 * @returns {Promise<string>} La clé API saisie ou null si annulé
 */
function promptForApiKey() {
  return new Promise((resolve) => {
    // Essayer de récupérer la clé depuis localStorage (pour le développement)
    let savedApiKey = null;
    try {
      savedApiKey = localStorage.getItem("openai_api_key");
    } catch {}

    if (savedApiKey) {
      // Si une clé est sauvegardée, l'utiliser directement
      resolve(savedApiKey);
      return;
    }

    // Sinon, demander à l'utilisateur de saisir la clé
    Office.context.ui.displayDialogAsync(
      `${location.origin}/auth.html?prompt=api_key`,
      { height: 60, width: 40, displayInIframe: false },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          resolve(null);
          return;
        }

        const dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          try {
            const msg = JSON.parse(arg.message || "{}");
            if (msg.type === "api_key" && msg.apiKey) {
              // Sauvegarder la clé pour une utilisation future
              try {
                localStorage.setItem("openai_api_key", msg.apiKey);
              } catch {}
              dialog.close();
              resolve(msg.apiKey);
            } else if (msg.type === "cancel") {
              dialog.close();
              resolve(null);
            }
          } catch (e) {
            dialog.close();
            resolve(null);
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          resolve(null);
        });
      }
    );
  });
}

/**
 * Met à jour la notification affichée
 * @param {string} message - Le message à afficher
 * @param {string} type - Le type de message (success, error, info)
 */
function updateNotification(message, type = "info") {
  const messageType = type === "error" 
    ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    : type === "success"
    ? Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
    : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

  const notificationMessage = {
    type: messageType,
    message: message,
    icon: "Icon.80x80",
    persistent: false,
  };

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "AutoReplyNotification",
    notificationMessage
  );
}

// Enregistrer la fonction avec Office
Office.actions.associate("generateAutoReply", generateAutoReply);
