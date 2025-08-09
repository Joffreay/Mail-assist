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
