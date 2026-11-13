var SHEET_NAME = "Leads";

function doGet() {
  return json({ ok: true, message: "Endpoint ativo" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var data = JSON.parse((e.postData && e.postData.contents) || "{}");

    if (data.website) {
      return json({ ok: true });
    }

    var nome = clean(data.nome, 120);
    var whatsapp = clean(data.whatsapp, 30);
    var bairro = clean(data.bairro, 120);
    var dia = clean(data.dia, 40);
    var origem = clean(data.origem, 80) || "Perguntas Sem Filtro";

    if (!nome || !whatsapp || !bairro || !dia) {
      return json({ ok: false, error: "Campos obrigatórios." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Data",
        "Nome",
        "WhatsApp",
        "Bairro",
        "Dia livre",
        "Origem",
      ]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([new Date(), nome, whatsapp, bairro, dia, origem]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function clean(value, max) {
  return String(value || "")
    .trim()
    .slice(0, max);
}

function json(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
