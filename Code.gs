/**
 * Script d'automatisation pour le fichier BASE DES DONNÉES E30 FASO
 * Exécutez la fonction 'initialiserToutLeClasseur' pour tout configurer d'un coup.
 */

function initialiserToutLeClasseur() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  configurerFeuilleUtilisateurs(ss);
  configurerFeuilleMembres(ss);
  configurerFeuilleRapportReunion(ss);
  
  SpreadsheetApp.getUi().alert("✅ Configuration complète réussie ! Vos 3 feuilles sont prêtes.");
}

/**
 * 1. Configuration de l'onglet 'Utilisateurs'
 */
function configurerFeuilleUtilisateurs(ss) {
  var sheet = ss.getSheetByName("Utilisateurs");
  if (!sheet) return;

  var rangeHeader = sheet.getRange("A1:E1");
  rangeHeader.setBackground("#1A73E8") // Bleu Google
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

  // Formatage des colonnes Date et Téléphone
  sheet.getRange("D2:D500").setNumberFormat("dd/mm/yyyy");
  sheet.getRange("J2:J500").setNumberFormat("@"); // Format Texte pour garder le 0 du téléphone
  
  sheet.autoResizeColumns(1, 13);
}

/**
 * 3. Configuration de l'onglet 'Rapport de réunion'
 */
function configurerFeuilleRapportReunion(ss) {
  var sheet = ss.getSheetByName("Rapport de réunion");
  if (!sheet) return;

  // En-têtes
  var rangeHeader = sheet.getRange("A1:M1");
  rangeHeader.setBackground("#FBBC04") // Jaune/Orange
             .setFontColor("#000000")
             .setFontWeight("bold");

  // Liste déroulante pour le Statut (Colonne M)
  var regleStatut = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Non commencé", "En cours", "Terminé", "En retard"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("M2:M200").setDataValidation(regleStatut);

  // Formule automatique pour générer l'ID Réunion en colonne A (si Date saisie en B)
  sheet.getRange("A2").setFormula('=ARRAYFORMULA(SI(ESTVIDE(B2:B); ""; "R-2026-" & TEXTE(LIGNE(B2:B)-1; "000")))');

  // Formatage des dates (Colonnes B et L)
  sheet.getRange("B2:B200").setNumberFormat("dd/mm/yyyy");
  sheet.getRange("L2:L200").setNumberFormat("dd/mm/yyyy");

  sheet.autoResizeColumns(1, 13);
}
