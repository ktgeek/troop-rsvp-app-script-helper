const PATROL_ROSTER_SETTINGS_KEY = 'patrolRosterSettings'

// What scouts see: no contacts or medical details, only what roll call and meal planning need.
type PatrolRosterPerson = {name: string; role: string; dietary: string; notes: string}

type PatrolRosterGroup = {title: string; people: PatrolRosterPerson[]}

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
      role: person.role,
      dietary: dietaryLine(person),
      notes: person.patrolNotes,
    })),
  }))
}

function showPatrolRoster() {
  showPrintDialog('PatrolRosterDialog', 'Printable patrol roster')
}

// Called from PatrolRosterDialog.html via google.script.run.
function getPatrolRosterPage(): {groups: PatrolRosterGroup[]; settings: HeaderSettings} {
  return {
    groups: patrolRosterGroups(),
    settings: readSettings(PATROL_ROSTER_SETTINGS_KEY, DEFAULT_HEADER_SETTINGS),
  }
}

// Called from PatrolRosterDialog.html via google.script.run.
function savePatrolRosterSettings(settings: HeaderSettings) {
  writeSettings(PATROL_ROSTER_SETTINGS_KEY, settings)
}
