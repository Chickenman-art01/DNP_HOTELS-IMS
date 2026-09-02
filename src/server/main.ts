/**
 * DNP HOTELS - INVENTORY MANAGEMENT SYSTEM (IMS)
 * Web App Main Router & Entry Points
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('client/Index');
  return template.evaluate()
    .setTitle('DNP HOTELS | Inventory Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
