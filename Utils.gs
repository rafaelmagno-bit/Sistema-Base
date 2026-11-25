// =============================
// 🧰 UTILITÁRIOS
// =============================

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
}

function getAba(nome) {
  return getSpreadsheet().getSheetByName(nome);
}

function formatarData(data) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );
}

function getUsuario() {
  let email = "";

  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {}

  if (!email) {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch (e) {}
  }

  return email || "Usuário não identificado";
}