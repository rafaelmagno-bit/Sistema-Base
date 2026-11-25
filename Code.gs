// 1. Deixe a versão no topo, fora de qualquer função
var VERSAO_SISTEMA = "1.1.0"; 

function doGet() {
  // Agora o doGet fica bem simples
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("Sistema de Apoio a Gestão Patrimonial - UFPE")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. ATUALIZE a função include para esta versão:
function include(filename) {
  // createTemplate permite que o Google processe os códigos <?= ?> lá dentro
  return HtmlService.createTemplateFromFile(filename)
    .evaluate()
    .getContent();
}

// =============================
// 👤 USUÁRIO
// =============================

function getUsuarioLogado() {
  return Session.getEffectiveUser().getEmail() || "Usuário não identificado";
}