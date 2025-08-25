import { config } from 'dotenv';
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

config();

// Configuration - À adapter selon votre environnement
const configObj = {
  tenantId: process.env.TENANT_ID || "votre-tenant-id-client",
  clientId: process.env.CLIENT_ID || "votre-client-id",
  clientSecret: process.env.CLIENT_SECRET || "votre-client-secret"
};

async function listUsers() {
  try {
    console.log("🔐 Connexion au tenant...");
    
    // Créer les credentials app-only
    const credential = new ClientSecretCredential(
      configObj.tenantId, 
      configObj.clientId, 
      configObj.clientSecret,
      { additionallyAllowedTenants: ["*"] }
    );

    // Obtenir le token d'accès
    const token = await credential.getToken("https://graph.microsoft.com/.default");
    if (!token) {
      throw new Error("Impossible d'obtenir le token d'accès");
    }

    console.log("✅ Token obtenu avec succès");

    // Initialiser le client Graph
    const graphClient = Client.init({
      authProvider: async (done) => {
        done(null, token.token);
      }
    });

    // Lister les utilisateurs
    console.log("\n📋 Liste des utilisateurs dans le tenant...");
    try {
      const users = await graphClient.api('/users')
        .select("id,displayName,userPrincipalName,mail,accountEnabled")
        .top(50)
        .get();
      
      console.log(`✅ ${users.value.length} utilisateur(s) trouvé(s):`);
      users.value.forEach((user, index) => {
        console.log(`\n${index + 1}. Utilisateur:`);
        console.log(`   🆔 ID: ${user.id}`);
        console.log(`   👤 Nom: ${user.displayName || 'Non défini'}`);
        console.log(`   📧 UPN: ${user.userPrincipalName || 'Non défini'}`);
        console.log(`   📮 Mail: ${user.mail || 'Non défini'}`);
        console.log(`   ✅ Actif: ${user.accountEnabled ? 'Oui' : 'Non'}`);
      });

      // Identifier les utilisateurs avec des boîtes mail
      const usersWithMail = users.value.filter(user => user.mail);
      if (usersWithMail.length > 0) {
        console.log(`\n🎯 ${usersWithMail.length} utilisateur(s) avec boîte mail:`);
        usersWithMail.forEach((user, index) => {
          console.log(`   ${index + 1}. ${user.displayName} (${user.mail})`);
        });
        
        // Suggérer le premier utilisateur avec mail
        const suggestedUser = usersWithMail[0];
        console.log(`\n💡 Suggestion: Utilisez ${suggestedUser.mail} dans votre fichier .env`);
        console.log(`   USER_ID=${suggestedUser.mail}`);
      }

    } catch (error) {
      console.error(`❌ Erreur lors de la récupération des utilisateurs: ${error.message}`);
      
      if (error.response) {
        console.error(`   📡 Détails de l'erreur:`);
        console.error(`      Status: ${error.response.status}`);
        console.error(`      Code: ${error.response.data?.error?.code}`);
        console.error(`      Message: ${error.response.data?.error?.message}`);
      }
    }

    // Vérifier les informations du tenant
    console.log("\n🏢 Informations du tenant...");
    try {
      const tenant = await graphClient.api('/organization').get();
      if (tenant.value && tenant.value.length > 0) {
        const org = tenant.value[0];
        console.log(`   🏢 Nom: ${org.displayName || 'Non défini'}`);
        console.log(`   🆔 ID: ${org.id || 'Non défini'}`);
        console.log(`   🌐 Domaine: ${org.verifiedDomains?.[0]?.name || 'Non défini'}`);
      }
    } catch (error) {
      console.error(`   ❌ Erreur: ${error.message}`);
    }

  } catch (error) {
    console.error("❌ Erreur lors de la connexion:", error);
    
    if (error.response) {
      console.error("📡 Détails de l'erreur Graph:");
      console.error("   Status:", error.response.status);
      console.error("   Message:", error.response.data?.error?.message);
      console.error("   Code:", error.response.data?.error?.code);
    }
    
    throw error;
  }
}

// Exécuter le script
listUsers().catch(console.error);
