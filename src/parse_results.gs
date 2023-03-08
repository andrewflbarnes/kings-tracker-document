//# vi: ft=javascript

/** @OnlyCurrentDoc */

const TABLE_START = `<div class=\"league-table-wrapper\">
  <table class=\"ResultsDefault\">
    <thead>
      <tr class=\"tableizer-firstrow\"><th>Position</th><th>Team</th><th>Place</th><th>Points</th><th>Place</th><th>Points</th><th>Place</th><th>Points</th><th>Place</th><th>Points</th><th>Total Points</th><th>Comments</th></tr>
    </thead>
    <tbody>`;
const TABLE_START_ALT = `<div class=\"league-table-wrapper\">
  <table class=\"ResultsAlt\">
    <thead>
      <tr class=\"tableizer-firstrow\"><th>Position</th><th>Team</th><th>Place</th><th>Points</th><th>Place</th><th>Points</th><th>Place</th><th>Points</th><th>Place</th><th>Points</th><th>Total Points</th><th>Comments</th></tr>
    </thead>
    <tbody>`;    
const TABLE_END = `    </tbody>
  </table>
</div>`;

// function onOpen() {
//   SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
//       .createMenu('Kings')
//       .addItem('Generate Site HTML', 'parseResults')
//       .addToUi();
// }

/**
 * For use when deploying as a webapp
 */
function doGet(e) {
  const results = getResultsAsHtml_();

  const htmlResults = "";
  results.forEach((v, k) => {
    htmlResults += `\n${k} Results: \n\n${v}\n`;
  });

  return ContentService.createTextOutput(htmlResults).setMimeType(ContentService.MimeType.TEXT);
}

function parseResults() {
  const ui = SpreadsheetApp.getUi();
  const results = getResultsAsHtml_();

  results.forEach((divisionResults, division) => {
    // Display HTML in alert for copying
    ui.alert(`${division} division`, divisionResults, ui.ButtonSet.OK);

    // Log execution results
    Logger.log(division);
    Logger.log(divisionResults);
  })
}

function getResultsAsHtml_() {
  const results = getResults_();

  const htmlResults = new Map();
  results.forEach((divisionResults, division) => {
    const divisionHtmlRows = divisionResults
      .map((r, index) => asHtml_(index + 1, r))
      .join("\n");

    if (division == 'Ladies') { //Ladies results tables are using alternate CSS class
      htmlResults.set(division, `${TABLE_START_ALT}\n${divisionHtmlRows}\n${TABLE_END}`);
    } else {
      htmlResults.set(division, `${TABLE_START}\n${divisionHtmlRows}\n${TABLE_END}`);
    }
  });

  return htmlResults;
}

function getResults_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Teams');
  let current = sheet.getRange('A6');

  const results = new Map();

  while (!current.isBlank()) {
    const division = current.getValue();

    if (!results.has(division)) {
      results.set(division, [])
    }
    const divisionResults = results.get(division);

    const result = asResult_(current);
    divisionResults.push(result);

    current = current.offset(1, 0);
  }

  results.forEach(divisionResults => divisionResults.sort(teamSort_));

  return results;
}

function teamSort_(a, b) {
  // sort by total points
  let res  = b.total - a.total
  if (res != 0) {
    return res
  }

  // if total points equal sort by ordered points
  for (let i = 0; i < a.ordered.length; i++) {
    res = b.ordered[i] - a.ordered[i]
    if (res != 0) {
      return res
    }
  }

  // if ordered points equal go by strongest finish
  const aLatest = [...a.points].reverse()
  const bLatest = [...b.points].reverse()
  for (let i = 0; i < aLatest.length; i++) {
    res = bLatest[i] - aLatest[i]
    if (res != 0) {
      return res
    }
  }

  // If strongest finish equal go by club ranking
  // This only happens in the case where both teams only got either participation points (1)
  // or didn't attend every round - and had the same result as each other in each round.
  const [aTeam, aTeamIndex] = teamAndIndexFromName_(a.team)
  const [bTeam, bTeamIndex] = teamAndIndexFromName_(b.team)
  res = aTeamIndex - bTeamIndex
  if (res != 0) {
    return res
  }

  // if same club ranking go by name
  if (aTeam == bTeam) {
    return 0
  }

  return aTeam < bTeam ? 1 : -1
}

function teamAndIndexFromName_(name) {
    const atoms = name.split(" ")
    const teamIndex = atoms.length > 1 ? atoms[atoms.length - 1] : 0
    const team = atoms[0].toUpperCase()
    return [team, teamIndex]
}

function asHtml_(position, {
    team,
    r1place,
    r1,
    r2place,
    r2,
    r3place,
    r3,
    r4place,
    r4,
    total,
    comment,
  }) {
  return `      <tr><td>${position}</td><td>${team}</td><td>${r1place}</td><td>${r1}</td><td>${r2place}</td><td>${r2}</td><td>${r3place}</td><td>${r3}</td><td>${r4place}</td><td>${r4}</td><td>${total}</td><td>${comment}</td></tr>`;
}

function asResult_(cell) {
  const team = cell.offset(0, 1).getValue();
  const r1place = cell.offset(0, 2).getValue();
  const r1 = cell.offset(0, 3).getValue();
  const r2place = cell.offset(0, 4).getValue();
  const r2 = cell.offset(0, 5).getValue();
  const r3place = cell.offset(0, 6).getValue();
  const r3 = cell.offset(0, 7).getValue();
  const r4place = cell.offset(0, 8).getValue();
  const r4 = cell.offset(0, 9).getValue();
  const total = cell.offset(0, 10).getValue();
  const comment = cell.offset(0, 11).getValue();
  const points = [r1 ?? 0, r2 ?? 0, r3 ?? 0, r4 ?? 0]
  const ordered = [...points].sort((a, b) => b - a)

  return {
    team,
    r1place,
    r1,
    r2place,
    r2,
    r3place,
    r3,
    r4place,
    r4,
    points,
    ordered,
    total,
    comment,
  }
}
