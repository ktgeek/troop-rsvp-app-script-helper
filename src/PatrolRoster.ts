// Spelled-out PoR names for the scout-facing roster, keyed like ROLE_RANK. Anything not listed prints as typed.
const ROLE_NAMES: {[role: string]: string} = {
  spl: 'Senior Patrol Leader',
  aspl: 'Assistant Senior Patrol Leader',
  pl: 'Patrol Leader',
  apl: 'Assistant Patrol Leader',
  qm: 'Quartermaster',
}

const PATROL_ROSTER_SETTINGS_KEY = 'patrolRosterSettings'

type PatrolRosterSettings = {title: string; subtitle: string}

const DEFAULT_PATROL_ROSTER_SETTINGS: PatrolRosterSettings = {title: '', subtitle: ''}

// What scouts see: no contacts or medical details, only what roll call and meal planning need.
type PatrolRosterPerson = {name: string; role: string; dietary: string; notes: string}

type PatrolRosterGroup = {title: string; people: PatrolRosterPerson[]}

function roleDisplayName(role: string): string {
  return ROLE_NAMES[role.toLowerCase().replace(/\./g, '')] ?? role
}

// Allergies are folded in because the cooks need them as much as the dietary answers.
function dietaryLine(person: RosterPerson): string {
  return [person.allergies && `Allergy: ${person.allergies}`, person.dietary]
    .filter(Boolean)
    .join('; ')
}

function patrolRosterGroups(): PatrolRosterGroup[] {
  return rosterGroups().map(group => ({
    title: group.title,
    people: group.people.map(person => ({
      name: person.name,
      role: roleDisplayName(person.role),
      dietary: dietaryLine(person),
      notes: person.patrolNotes,
    })),
  }))
}

function showPatrolRoster() {
  const dialog = HtmlService.createHtmlOutputFromFile('PatrolRosterDialog')
    .setWidth(1000)
    .setHeight(760)
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Printable patrol roster')
}

// Called from PatrolRosterDialog.html via google.script.run.
function getPatrolRosterPage(): {groups: PatrolRosterGroup[]; settings: PatrolRosterSettings} {
  return {
    groups: patrolRosterGroups(),
    settings: readSettings(PATROL_ROSTER_SETTINGS_KEY, DEFAULT_PATROL_ROSTER_SETTINGS),
  }
}

// Called from PatrolRosterDialog.html via google.script.run.
function savePatrolRosterSettings(settings: PatrolRosterSettings) {
  writeSettings(PATROL_ROSTER_SETTINGS_KEY, settings)
}
