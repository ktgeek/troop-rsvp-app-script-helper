// What every print dialog (roster, patrol roster, driving sheet) asks for in its header.
type HeaderSettings = {title: string; subtitle: string}

const DEFAULT_HEADER_SETTINGS: HeaderSettings = {title: '', subtitle: ''}

// Per-printout header settings live in document properties so the next outing starts from the last one's.
function readSettings<T extends object>(key: string, defaults: T): T {
  const saved = PropertiesService.getDocumentProperties().getProperty(key)
  return {...defaults, ...(saved ? JSON.parse(saved) : {})}
}

function writeSettings(key: string, settings: object) {
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(settings))
}

// Lets a dialog page pull in shared markup: <?!= include('PrintShared') ?> in a template evaluates to PrintShared.html.
function include(filename: string): string {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

function showPrintDialog(filename: string, title: string) {
  const dialog = HtmlService.createTemplateFromFile(filename)
    .evaluate()
    .setWidth(1000)
    .setHeight(760)
  SpreadsheetApp.getUi().showModalDialog(dialog, title)
}
