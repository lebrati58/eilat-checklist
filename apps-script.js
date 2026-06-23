// ════════════════════════════════════════════════════════
//  Google Apps Script — Check-out Maison Eilat
//  À coller dans : script.google.com → Nouveau projet
// ════════════════════════════════════════════════════════

const SHEET_NAME = "Départs Eilat";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sheet  = ss.getSheetByName(SHEET_NAME);

    // Créer la feuille + en-têtes si elle n'existe pas encore
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "📅 Date envoi",
        "👤 Nom",
        "📅 Date départ",
        "🔑 Clefs",
        "✅ Tâches",
        "🏊 Piscine",
        "🚰 Eau",
        "🔴 Gaz",
        "🪟 Volets",
        "🕐 Horodatage vérif",
        "📝 Notes"
      ]);
      // Mise en forme de l'en-tête
      sheet.getRange(1, 1, 1, 11)
        .setBackground("#1A6B8A")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Ajouter une ligne avec les données reçues
    sheet.appendRow([
      new Date().toLocaleString("fr-FR"),
      data.nom        || "—",
      data.date       || "—",
      data.clefs      || "—",
      data.taches     || "—",
      data.piscine    ? "✅" : "❌",
      data.eau        ? "✅" : "❌",
      data.gaz        ? "✅" : "❌",
      data.volets     ? "✅" : "❌",
      data.stamp      || "—",
      data.notes      || "—"
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test manuel depuis l'éditeur Apps Script
function test() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        nom: "Lili", date: "2026-06-23", clefs: "chez Rivka",
        taches: "19 / 19", piscine: true, eau: true,
        gaz: true, volets: true, stamp: "lundi 23 juin 2026, 14:30",
        notes: "Test OK"
      })
    }
  };
  doPost(fakeEvent);
  Logger.log("Ligne ajoutée !");
}
