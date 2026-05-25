// ==========================================
// ROUTER.GS - Entry point & HTML include helper
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Azara Course Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: include file HTML lain ke dalam template
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}