/**
 * ArteDelCantar - Backend Actualizado para servir Web App
 */

// Esta es la función que renderiza tu archivo index.html como Web App
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ArteDelCantar')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Tus endpoints de API se mantienen intactos aquí abajo...
function doPost(e) {
  return ContentService.createTextOutput(JSON.stringify({status: "success"}))
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader("Access-Control-Allow-Origin", "*");
}
