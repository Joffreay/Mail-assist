import { config } from 'dotenv';
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

config();

// Configuration avec le bon utilisateur identifié
const configObj = {
  tenantId: process.env.TENANT_ID || "votre-tenant-id-client",
  clientId: process.env.CLIENT_ID || "votre-client-id",
  clientSecret: process.env.CLIENT_SECRET || "votre-client-secret",
  userEmail: "jpelletier_hotmail.fr#EXT#_jpelletierhotmail.onmicrosoREFQZ#EXT#@testNotaire.onmicrosoft.com"
};

async function checkMailbox() {
  try {
    console.log("🔐 Connexion au tenant...");
    console.log(`👤 Utilisateur testé: ${configObj.userEmail}`);
    
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

    // Test 1: Vérifier les informations de l'utilisateur
    console.log("\n📋 Test 1: Informations de l'utilisateur");
    try {
      const user = await graphClient.api(`/users/${configObj.userEmail}`).get();
      console.log(`   ✅ Utilisateur trouvé: ${user.displayName}`);
      console.log(`   📧 UPN: ${user.userPrincipalName}`);
      console.log(`   🆔 ID: ${user.id}`);
      console.log(`   📮 Mail: ${user.mail || 'Non défini'}`);
      console.log(`   🏢 Company: ${user.companyName || 'Non défini'}`);
      console.log(`   📱 Usage Location: ${user.usageLocation || 'Non défini'}`);
    } catch (error) {
      console.error(`   ❌ Erreur: ${error.message}`);
    }

    // Test 2: Vérifier la licence Exchange Online
    console.log("\n📋 Test 2: Licence Exchange Online");
    try {
      const licenses = await graphClient.api(`/users/${configObj.userEmail}/licenseDetails`).get();
      console.log(`   ✅ Licences trouvées: ${licenses.value.length}`);
      
      licenses.value.forEach((license, index) => {
        console.log(`   ${index + 1}. Licence: ${license.skuId}`);
        console.log(`      Nom: ${license.skuPartNumber || 'Non défini'}`);
        console.log(`      État: ${license.capabilityStatus || 'Non défini'}`);
      });
      
      // Chercher Exchange Online
      const exchangeLicense = licenses.value.find(license => 
        license.skuPartNumber && license.skuPartNumber.includes('EXCHANGE')
      );
      
      if (exchangeLicense) {
        console.log(`   🎯 Exchange Online trouvé: ${exchangeLicense.skuPartNumber}`);
      } else {
        console.log(`   ⚠️  Aucune licence Exchange Online trouvée`);
      }
      
    } catch (error) {
      console.error(`   ❌ Erreur: ${error.message}`);
    }

    // Test 3: Vérifier l'accès aux messages
    console.log("\n📋 Test 3: Accès aux messages");
    try {
      const messages = await graphClient.api(`/users/${configObj.userEmail}/messages`)
        .top(1)
        .select("id,subject,from,receivedDateTime")
        .get();
      
      if (messages.value.length > 0) {
        console.log(`   ✅ Accès aux messages OK: ${messages.value.length} message(s) trouvé(s)`);
        console.log(`   📝 Premier message: ${messages.value[0].subject || 'Sans objet'}`);
      } else {
        console.log(`   ⚠️  Aucun message trouvé (boîte vide ou permissions insuffisantes)`);
      }
    } catch (error) {
      console.error(`   ❌ Erreur: ${error.message}`);
      
      if (error.response) {
        console.error(`   📡 Détails de l'erreur:`);
        console.error(`      Status: ${error.response.status}`);
        console.error(`      Code: ${error.response.data?.error?.code}`);
        console.error(`      Message: ${error.response.data?.error?.message}`);
      }
    }

    // Test 4: Vérifier l'accès aux dossiers
    console.log("\n📋 Test 4: Accès aux dossiers");
    try {
      const folders = await graphClient.api(`/users/${configObj.userEmail}/mailFolders`)
        .top(5)
        .select("id,displayName")
        .get();
      
      console.log(`   ✅ Accès aux dossiers OK: ${folders.value.length} dossier(s) trouvé(s)`);
      folders.value.forEach(folder => {
        console.log(`      📁 ${folder.displayName}`);
      });
    } catch (error) {
      console.error(`   ❌ Erreur: ${error.message}`);
    }

    // Test 5: Vérifier l'accès au calendrier
    console.log("\n📋 Test 5: Accès au calendrier");
    try {
      const calendars = await graphClient.api(`/users/${configObj.userEmail}/calendars`)
        .top(5)
        .select("id,name,color")
        .get();
      
      console.log(`   ✅ Accès au calendrier OK: ${calendars.value.length} calendrier(s) trouvé(s)`);
      calendars.value.forEach(calendar => {
        console.log(`      📅 ${calendar.name || 'Sans nom'}`);
      });
    } catch (error) {
      console.error(`   ❌ Erreur: ${error.message}`);
    }

    console.log("\n🎯 Résumé:");
    console.log("   • Si Test 2 échoue: L'utilisateur n'a pas de licence Exchange Online");
    console.log("   • Si Test 3 échoue: La boîte mail n'est pas accessible");
    console.log("   • Si Test 4 échoue: Les dossiers mail ne sont pas accessibles");
    console.log("   • Vérifiez les licences dans Azure AD");

  } catch (error) {
    console.error("❌ Erreur lors de la vérification:", error);
    
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
checkMailbox().catch(console.error);
