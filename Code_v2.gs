/**
 * Script d'automatisation pour le fichier BASE DES DONNÉES E30 FASO
 */

function initialiserToutLeClasseur() {
  var ID_FICHIER = "1Rgsb81FPuyF865Ox0ufSIg6uGEa-3xFunVXXmoOJloA";
  var ss = SpreadsheetApp.openById(ID_FICHIER);
  
  configurerFeuilleUtilisateurs(ss);
  configurerFeuilleMembres(ss);
  configurerFeuilleRapportReunion(ss);
  
  Logger.log("✅ Configuration complète réussie ! Vos 3 feuilles sont prêtes.");
}

/**
 * 1. Configuration de l'onglet 'Utilisateurs'
 */
function configurerFeuilleUtilisateurs(ss) {
  var sheet = ss.getSheetByName("Utilisateurs");
  if (!sheet) return;

  var rangeHeader = sheet.getRange("A1:E1");
  rangeHeader.setBackground("#1A73E8") // Bleu
             .setFontColor("#FFFFFF")
             .setFontWeight("bold");
             
  sheet.autoResizeColumns(1, 5);
}

/**
 * 2. Configuration de l'onglet 'Liste des membres du groupe'
 */
function configurerFeuilleMembres(ss) {
  var sheet = ss.getSheetByName("Liste des membres du groupe");
  if (!sheet) return;

  var rangeHeader = sheet.getRange("A1:M1");
  rangeHeader.setBackground("#34A853") // Vert
             .setFontColor("#FFFFFF")
             .setFontWeight("bold");

  sheet.getRange("D2:D500").setNumberFormat("dd/mm/yyyy");
  sheet.getRange("J2:J500").setNumberFormat("@"); // Format Texte pour numéros
  
  sheet.autoResizeColumns(1, 13);
}

/**
 * 3. Configuration de l'onglet 'Rapport de réunion'
 */
function configurerFeuilleRapportReunion(ss) {
  var sheet = ss.getSheetByName("Rapport de réunion");
  if (!sheet) return;

  var rangeHeader = sheet.getRange("A1:M1");
  rangeHeader.setBackground("#FBBC04") // Jaune / Orange
             .setFontColor("#000000")
             .setFontWeight("bold");

  // Liste déroulante pour le Statut (Colonne M)
  var regleStatut = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Non commencé", "En cours", "Terminé", "En retard"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("M2:M200").setDataValidation(regleStatut);

  // Formule d'ID automatique en colonne A
  sheet.getRange("A2").setFormula('=ARRAYFORMULA(SI(ESTVIDE(B2:B); ""; "R-2026-" & TEXTE(LIGNE(B2:B)-1; "000")))');

  // Formatage des dates
  sheet.getRange("B2:B200").setNumberFormat("dd/mm/yyyy");
  sheet.getRange("L2:L200").setNumberFormat("dd/mm/yyyy");

  sheet.autoResizeColumns(1, 13);
}

/**
 * 4. Gestion de l'affichage Web App (requête GET)
 */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : "accueil";
  var ID_FICHIER = "1Rgsb81FPuyF865Ox0ufSIg6uGEa-3xFunVXXmoOJloA";
  var ss = SpreadsheetApp.openById(ID_FICHIER);

  if (action === "getReunions") {
    var sheet = ss.getSheetByName("Rapport de réunion");
    var data = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }

  var html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f6f9; color: #333; }
          .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); max-width: 500px; margin: auto; }
          h2 { color: #1a73e8; margin-top: 0; }
          .status { display: inline-block; padding: 6px 12px; background: #e6f4ea; color: #137333; border-radius: 4px; font-weight: bold; }
        </style>
      </head>
      <body>
        <div class="card">
          <h2>FasoAgro-Express - Web Service</h2>
          <p>Statut du service : <span class="status">● Opérationnel</span></p>
          <p>L'API Web App est active et prête à recevoir ou transmettre les données de la base <strong>BASE DES DONNÉES E30 FASO</strong>.</p>
        </div>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle("FasoAgro-Express Web Service")
    .setXframeOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 5. Gestion de la réception de données (requête POST)
 */
function doPost(e) {
  try {
    var ID_FICHIER = "1Rgsb81FPuyF865Ox0ufSIg6uGEa-3xFunVXXmoOJloA";
    var ss = SpreadsheetApp.openById(ID_FICHIER);
    var contents = JSON.parse(e.postData.contents);

    if (contents.type === "reunion") {
      var sheet = ss.getSheetByName("Rapport de réunion");
      var nouvelleLigne = [
        "", // L'ID est géré automatiquement par la formule
        contents.date || "",
        contents.heure || "",
        contents.lieu || "",
        contents.animateur || "",
        contents.ordreDuJour || "",
        contents.nbPresents || "",
        contents.pointsCles || "",
        contents.decisions || "",
        contents.action || "",
        contents.responsable || "",
        contents.echeance || "",
        contents.statut || "Non commencé"
      ];
      sheet.appendRow(nouvelleLigne);
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success",
        "message": "Rapport de réunion enregistré avec succès !"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": "Type de données inconnu."
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
