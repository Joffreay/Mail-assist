/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Afficher Hello world à l'ouverture
    const hello = document.getElementById("hello-text");
    if (hello) {
      hello.textContent = "Hello world";
    }
    document.getElementById("run").onclick = run;

    const apiKeyInput = document.getElementById("api-key");
    const promptInput = document.getElementById("prompt-input");
    const sendBtn = document.getElementById("send-prompt");
    const addMailBtn = document.getElementById("add-mail-content");
    const addCategoriesBtn = document.getElementById("add-mail-categories");
    const output = document.getElementById("response-output");
    const listEventsBtn = document.getElementById("list-events");
    const msalClientIdInput = document.getElementById("msal-client-id");
    const eventsOutput = document.getElementById("events-output");

    if (apiKeyInput && promptInput && sendBtn && output) {
      // charger la clé si déjà sauvegardée en localStorage (dev uniquement)
      try {
        const saved = localStorage.getItem("openai_api_key");
        if (saved) apiKeyInput.value = saved;
      } catch {}

      sendBtn.addEventListener("click", async () => {
        const apiKey = apiKeyInput.value.trim();
        const prompt = promptInput.value.trim();
        if (!apiKey || !prompt) {
          output.textContent = "Veuillez saisir la clé API et un prompt.";
          return;
        }
        try {
          // sauvegarde locale (dev)
          try { localStorage.setItem("openai_api_key", apiKey); } catch {}
          output.textContent = "Envoi en cours...";

          // Appel API OpenAI (Chat Completions)
          const resp = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "Authorization": `Bearer ${apiKey}`,
            },
            body: JSON.stringify({
              model: "gpt-3.5-turbo",
              messages: [
                { role: "system", content: "You are a helpful assistant." },
                { role: "user", content: prompt },
              ],
              temperature: 0.7,
            }),
          });

          if (!resp.ok) {
            const errText = await resp.text();
            output.textContent = `Erreur API (${resp.status}): ${errText}`;
            return;
          }
          const data = await resp.json();
          const answer = data?.choices?.[0]?.message?.content || "(Réponse vide)";
          output.textContent = answer;
        } catch (e) {
          output.textContent = `Erreur: ${e?.message || e}`;
        }
      });

      // Insérer le contenu du mail en cours dans le prompt
      if (addMailBtn) {
        addMailBtn.addEventListener("click", async () => {
          try {
            const item = Office.context.mailbox.item;
            const subject = item?.subject || "";
            let bodyText = "";
            // Récupère le corps au format texte
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

            const separator = promptInput.value ? "\n\n" : "";
            promptInput.value = `${promptInput.value}${separator}Sujet: ${subject}\n\nCorps:\n${bodyText}`.trim();
          } catch (e) {
            output.textContent = `Erreur lors de l'extraction du mail: ${e?.message || e}`;
          }
        });
      }

      // Ajouter les catégories du mail (Outlook) au prompt
      if (addCategoriesBtn) {
        addCategoriesBtn.addEventListener("click", async () => {
          try {
            const item = Office.context.mailbox.item;
            // Récupère les catégories liées à l'élément (lecture/composition)
            const categories = item?.categories || [];

            // Si vide côté item, tenter de récupérer les catégories définies par l'utilisateur (API mailbox.masterCategories)
            let userCats = [];
            try {
              const req = Office.context.requirements;
              const hasApi = req && req.isSetSupported && req.isSetSupported("Mailbox", "1.8");
              if (hasApi && Office.context.mailbox?.masterCategories?.getAsync) {
                await new Promise((resolve) => {
                  Office.context.mailbox.masterCategories.getAsync((res) => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                      userCats = (res.value || []).map((c) => c.displayName).filter(Boolean);
                    }
                    resolve();
                  });
                });
              }
            } catch {}

            const setFromItem = Array.isArray(categories) ? categories : [];
            const all = [...new Set([...(setFromItem || []), ...(userCats || [])])];
            if (all.length === 0) {
              output.textContent = "Aucune catégorie trouvée (l’API Catégories nécessite Outlook supportant Mailbox 1.8).";
              return;
            }
            const toAppend = `\n\nCatégories Outlook: ${all.join(", ")}`;
            promptInput.value = `${promptInput.value || ""}${toAppend}`.trim();
          } catch (e) {
            output.textContent = `Erreur lors de la récupération des catégories: ${e?.message || e}`;
          }
        });
      }
    }
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}

// Ouvre un Dialog pour authentifier l'utilisateur et renvoyer un token MS Graph
function openAuthDialog(clientId, onResult) {
  const url = `${location.origin}/auth2.html?clientId=${encodeURIComponent(clientId)}`;
  Office.context.ui.displayDialogAsync(
    url,
    { height: 60, width: 40, displayInIframe: false },
    (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        onResult(null, new Error("Impossible d’ouvrir la fenêtre d’authentification."));
        return;
      }
      const dialog = res.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        try {
          const msg = JSON.parse(arg.message || "{}");
          if (msg.type === "token" && msg.token) {
            dialog.close();
            onResult(msg.token, null);
          } else if (msg.type === "error") {
            dialog.close();
            onResult(null, new Error(msg.error || "Erreur d’authentification."));
          }
        } catch (e) {
          dialog.close();
          onResult(null, e);
        }
      });
    }
  );
}

// Attache l’action du bouton "Lister mes prochains RDV"
if (typeof document !== "undefined") {
  document.addEventListener("DOMContentLoaded", () => {
    // Bouton catégories utilisateur (masterCategories)
    const catBtn = document.getElementById("list-categories");
    const catOut = document.getElementById("categories-output");
    if (catBtn && catOut) {
      catBtn.addEventListener("click", () => {
        try {
          const req = Office.context.requirements;
          const hasApi = req && req.isSetSupported && req.isSetSupported("Mailbox", "1.8");
          if (!(hasApi && Office.context.mailbox?.masterCategories?.getAsync)) {
            catOut.textContent = "API non supportée (Mailbox 1.8 requis).";
            return;
          }
          catOut.textContent = "Chargement des catégories...";
          Office.context.mailbox.masterCategories.getAsync((res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              const names = (res.value || []).map((c) => c.displayName).filter(Boolean);
              catOut.textContent = names.length ? names.join("\n") : "Aucune catégorie définie.";
            } else {
              catOut.textContent = `Erreur: ${res.error?.message || res.error || "inconnue"}`;
            }
          });
        } catch (e) {
          catOut.textContent = `Erreur: ${e?.message || e}`;
        }
      });
    }
    const btn = document.getElementById("list-events");
    const clientIdInput = document.getElementById("msal-client-id");
    const out = document.getElementById("events-output");
    // Préremplir depuis localStorage si disponible
    try {
      const savedClientId = localStorage.getItem("msal_client_id");
      if (savedClientId && clientIdInput && !clientIdInput.value) {
        clientIdInput.value = savedClientId;
      }
    } catch {}
    if (btn && clientIdInput && out) {
      btn.addEventListener("click", () => {
        const clientId = clientIdInput.value.trim();
        if (!clientId) {
          out.textContent = "Veuillez saisir un Client ID (Azure AD) pour appeler Microsoft Graph.";
          return;
        }
        // Mémoriser le Client ID pour l’auth dialog
        try { localStorage.setItem("msal_client_id", clientId); } catch {}
        out.textContent = "Authentification...";
        openAuthDialog(clientId, async (accessToken, err) => {
          if (err) {
            out.textContent = `Erreur: ${err.message || err}`;
            return;
          }
          try {
            out.textContent = "Chargement des rendez-vous...";
            const now = new Date();
            const end = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
            const startIso = now.toISOString();
            const endIso = end.toISOString();
            const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${encodeURIComponent(startIso)}&endDateTime=${encodeURIComponent(endIso)}&$orderby=start/dateTime&$top=10`;
            const resp = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
            if (!resp.ok) {
              const t = await resp.text();
              out.textContent = `Erreur Graph (${resp.status}): ${t}`;
              return;
            }
            const data = await resp.json();
            const items = data?.value || [];
            if (items.length === 0) {
              out.textContent = "Aucun rendez-vous à venir dans les 7 prochains jours.";
              return;
            }
            const lines = items.map((it) => {
              const subject = it.subject || "(Sans objet)";
              const start = it.start?.dateTime || "";
              const end = it.end?.dateTime || "";
              const location = it.location?.displayName ? ` @ ${it.location.displayName}` : "";
              return `- ${subject} (${start} -> ${end})${location}`;
            });
            out.textContent = lines.join("\n");
          } catch (e) {
            out.textContent = `Erreur lors de la récupération des événements: ${e?.message || e}`;
          }
        });
      });
    }
  });
}
