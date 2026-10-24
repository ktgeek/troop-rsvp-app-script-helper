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
