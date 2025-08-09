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
