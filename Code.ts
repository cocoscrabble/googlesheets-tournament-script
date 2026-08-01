// import {blossom} from "./blossom.js"

function onOpen() {
  // Get the UI object.
  var ui = SpreadsheetApp.getUi();

  // Create and add a named menu and its items to the menu bar.
  ui.createMenu('Tournament')
    .addItem('Calculate Standings and Pairings', 'calculateStandings')
    .addToUi();
}

class Results {
  constructor() {
    this.results = []
    this.players = {}
    this.rounds = {}
  }

  updateRound(game_result) {
    var round = game_result.round
    if (this.rounds[round] === undefined) {
      this.rounds[round] = []
    }
    this.rounds[round].push(game_result)
  }

  getPlayerResults(name) {
    if (this.players[name] === undefined) {
      this.players[name] = {
        name: name,
        wins: 0,
        losses: 0,
        ties: 0,
        score: 0,
        spread: 0,
        starts: 0,
        byes: 0,
      }
    }
    return this.players[name];
  }

  updatePlayer(result) {
    // Add the result of a single game to the player's results
    var p = this.getPlayerResults(result.name)
    var spread = result.score - result.opp_score
    p.spread += spread;
    if (spread > 0) {
      p.wins += 1;
    } else if (spread == 0) {
      p.ties += 1;
    } else {
      p.losses += 1;
    }
    p.score = p.wins + 0.5 * p.ties;
    p.starts += result.start;
    if (result.opp.toLowerCase() === "bye") {
      p.byes += 1;
    }
  }

  addGeneratedResult(game_result) {
    // Used for simulations
    this.results.push(game_result);
    this.processResult(game_result);
  }

  processResult(game_result) {
    var winner = winnerResults(game_result);
    var loser = loserResults(game_result);
    this.updatePlayer(winner);
    this.updatePlayer(loser);
    this.updateRound(game_result);
  }

  processResults() {
    for (const game_result of this.results) {
      this.processResult(game_result);
    }
  }

  roundIds() {
    return [0].concat(Object.keys(this.rounds).map(function (i) {
      return parseInt(i);
    }));
  }

  unplayedGames(round) {
    var players = new Set(Object.keys(this.players));
    if (this.rounds[round] === undefined) {
      return players;
    }
    for (const game of this.rounds[round]) {
      players.delete(game.winner);
      players.delete(game.loser);
    }
    for (const p of players) {
      if (p.toLowerCase() == "bye") {
        players.delete(p)
      }
    }
    return players;
  }

  extractPairings(round) {
    var pairings = [];
    console.log("extracting pairings for round:", round)
    // console.log(this.rounds[round])
    for (const game_result of this.rounds[round]) {
      var p = {
        first: {name: game_result.winner},
        second: {name: game_result.loser}
      }
      if (!game_result.winner_first) {
        [p.first, p.second] = [p.second, p.first]
      }
      pairings.push(p);
    }
    return pairings
  }
}

class Entrants {
  constructor(entrants, seeding, tables, fixed_pairings, fixed_starts) {
    this.entrants = entrants
    this.seeding = seeding
    this.tables = tables
    this.fixed_pairings = fixed_pairings
    this.fixed_starts = fixed_starts
    this.settings = new Settings()
  }

  addEntrant(e) {
    this.entrants[e.name] = e.full_name;
    this.seeding.push({ name: e.name, rating: e.rating, seed: e.seed });
    if (e.table != "") {
      this.tables[e.name] = parseInt(e.table);
    }
  }
}

class Repeats {
  constructor() {
    this.matches = {}
  }

  add(name1, name2) {
    // Add a pairing and return the current count
    var key = [name1, name2].sort();
    if (this.matches[key] === undefined) {
      this.matches[key] = 0;
    }
    this.matches[key]++;
    return this.matches[key];
  }

  get(name1, name2) {
    var key = [name1, name2].sort();
    return this.matches[key] || 0;
  }

  max() {
    // The highest number of times any pair has met
    var max = 0;
    for (const key of Object.keys(this.matches)) {
      max = Math.max(max, this.matches[key]);
    }
    return max;
  }
}

class Starts {
  constructor(res, entrants) {
    this.starts = {}  // number of starts
    this.h2h = {}  // for [name1, name2], did name1 start?
    this.recent_starts = {}  // most recent round started
    this.fixed_starts = entrants.fixed_starts;
  }

  init(name) {
    if (this.starts[name] === undefined) {
      this.starts[name] = 0;
    }
    if (this.recent_starts[name] === undefined) {
      this.recent_starts[name] = 0;
    }
  }

  _record(name1, name2, round, p1_starts) {
    if (p1_starts) {
      this.starts[name1]++
      this.recent_starts[name1] = round;
      this.h2h[[name1, name2]] = true;
      this.h2h[[name2, name1]] = false;
    } else {
      this.starts[name2]++
      this.recent_starts[name2] = round;
      this.h2h[[name1, name2]] = false;
      this.h2h[[name2, name1]] = true;
    }
  }

  register(pairing, round) {
    // Register an extracted pairing
    const name1 = pairing.first.name;
    const name2 = pairing.second.name;
    this.init(name1);
    this.init(name2);
    this._record(name1, name2, round, true);
  }

  add(name1, name2, round) {
    this.init(name1);
    this.init(name2);
    var p1_starts;
    // bye always starts
    if (name1.toLowerCase() === "bye") {
      p1_starts = true;
    } else if (name2.toLowerCase() === "bye") {
      p1_starts = false;
    } else if (this.fixed_starts[[round, name1]] === true) {
      p1_starts = true;
    } else if (this.fixed_starts[[round, name2]] === true) {
      p1_starts = false;
    } else {
      var starts1 = this.starts[name1];
      var starts2 = this.starts[name2];
      if (starts1 == starts2) {
        // Whoever went first most recently should go second now.
        if (this.h2h[[name1, name2]] === undefined) {
          if (this.recent_starts[name1] > this.recent_starts[name2]) {
            p1_starts = false;
          } else {
            p1_starts = true;
          }
        } else {
          if (this.h2h[[name1, name2]]) {
            p1_starts = false;
          } else {
            p1_starts = true;
          }
        }
      } else {
        p1_starts = starts1 < starts2;
      }
    }
    this._record(name1, name2, round, p1_starts);
    // console.log("starts:", name1, name2, p1_starts, this.starts[name1], this.starts[name2])
    return p1_starts;
  }
}

class Byes {
  constructor() {
    this.byes = {}
  }

  init(name) {
    if (this.byes[name] === undefined) {
      this.byes[name] = 0;
    }
  }

  add(name) {
    this.init(name);
    this.byes[name]++
  }

  get(name) {
    var ret = this.byes[name];
    return ret === undefined ? 0 : ret;
  }

  update(p) {
    if (isBye(p.first)) {
      this.add(p.second.name)
    }
    if (isBye(p.second)) {
      this.add(p.first.name)
    }
  }

  reset() {
    this.byes = {}
  }
}

class Fixed {
  constructor(pairings, starts) {
    this.pairings = pairings
    this.starts = starts
  }
}

// Global register of byes including ones paired for future rounds
var BYES = new Byes();

function winnerResults(game_result) {
  var start = game_result.winner_first ? 1 : 0;
  return {
    round: game_result.round,
    name: game_result.winner,
    score: game_result.winner_score,
    opp: game_result.loser,
    opp_score: game_result.loser_score,
    start: start,
  }
}

function loserResults(game_result) {
  var start = game_result.winner_first ? 0 : 1;
  return {
    round: game_result.round,
    name: game_result.loser,
    score: game_result.loser_score,
    opp: game_result.winner,
    opp_score: game_result.winner_score,
    start: start,
  }
}

class Settings {
  constructor() {
    this.swiss_weight = 30
    this.swiss_distance = 10
  }
}

class Diags {
  constructor() {
    this.round_status = {}
    this.rounds = {}
    this.pairs = {}
  }
}

// -----------------------------------------------------
// Read data from spreadsheet

function collectSheetData(sheet_label, range) {
  var gs = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = gs.getSheetByName(sheet_label);
  var data_range = sheet.getRange(range);
  var values = data_range.getValues();
  var start = range.split(":")[0]
  var last_row = sheet.getRange(start).getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var data = values.slice(0, last_row - 1);
  return data;
}

function makeResults(rows) {
  // Convert each row into a GameResult object
  var out = []
  for (var entry of rows) {
    // col B = entry[0], C = 1, ...
    var game_result = {
      round: parseInt(entry[0]),
      winner: entry[1],
      winner_score: parseInt(entry[2]),
      loser: entry[3],
      loser_score: parseInt(entry[4]),
      winner_first: entry[5].toLowerCase() == "first" ? true : false,
    }
    // make sure we have a valid round number, and ignore "test player"
    if (!isNaN(game_result.round) && (game_result.winner != "Test Player")) {
      out.push(game_result);
    }
  }
  var res = new Results();
  res.results = out;
  res.processResults();
  return res;
}

function collectResults() {
  var data = collectSheetData("Results", "B2:H");
  return makeResults(data);
}

function entrantFromRow(entry) {
  // col B = entry[0], C = 1, ...
  var name = entry[0];
  var rating = parseInt(entry[1]);
  var table = entry[3];
  var seed = parseInt(entry[4]);
  var full_name = name + ` (#${seed})`;
  if (isNaN(rating)) {
    rating = 0;
  }
  return {name: name, full_name: full_name, rating: rating, table: table, seed: seed}
}

function makeEntrants(rows) {
  var entrants = {}
  var seeding = []
  var tables = {}
  var fixed_pairings = {}
  var fixed_starts = {}
  var ret = new Entrants(
    entrants, seeding, tables, fixed_pairings, fixed_starts
  );
  for (var entry of rows) {
    ret.addEntrant(entrantFromRow(entry));
  }
  ret.seeding.sort(function (i, j) { return i.seed - j.seed });
  return ret;
}

function collectEntrants() {
  var data = collectSheetData("Entrants", "A2:E");
  var entrants = makeEntrants(data);
  console.log("Seeding:", entrants.seeding);
  console.log("Entrants:", entrants.entrants);
  console.log("Tables:", entrants.tables);
  return entrants;
}

function makeRoundPairings(rows) {
  var quads = {}
  var round_robins = {}
  var double_round_robins = {}
  var rounds = []
  for (var entry of rows) {
    // col A = entry[0], B = 1
    var round = parseInt(entry[0]);
    var pairing = entry[1];
    if (pairing.startsWith("Q") || pairing.startsWith("SIX")) {
      if (quads[pairing] === undefined) {
        quads[pairing] = [];
      }
      quads[pairing].push(round);
    } else if (pairing.startsWith("RAND")) {
      rounds[round] = { round: round, type: pairing, start: round };
    } else if (pairing.startsWith("R")) {
      if (round_robins[pairing] === undefined) {
        round_robins[pairing] = [];
      }
      round_robins[pairing].push(round);
    } else if (pairing.startsWith("DR")) {
      if (double_round_robins[pairing] === undefined) {
        double_round_robins[pairing] = [];
      }
      double_round_robins[pairing].push(round);
    } else if (pairing == "CH") {
      rounds[round] = { round: round, type: pairing, start: 0 };
    } else if (pairing == "ST" || pairing == "STPR") {
      rounds[round] = { round: round, type: pairing, start: round - 1 };
    } else {
      rounds[round] = { round: round, type: pairing, start: round };
    }
  }
  for (const q of Object.keys(quads)) {
    const quad = quads[q];
    for (var i = 0; i < quad.length; i++) {
      round = quad[i];
      rounds[round] = { round: round, type: q, start: quad[0], pos: i + 1 }
    }
  }
  for (const r of Object.keys(round_robins)) {
    const rr = round_robins[r];
    for (var i = 0; i < rr.length; i++) {
      round = rr[i];
      rounds[round] = { round: round, type: "R", start: rr[0], pos: i + 1 }
    }
  }
  for (const r of Object.keys(double_round_robins)) {
    const rr = double_round_robins[r];
    for (var i = 0; i < rr.length; i++) {
      round = rr[i];
      rounds[round] = { round: round, type: "DR", start: rr[0], pos: i + 1 }
    }
  }
  return rounds;
}

function collectRoundPairings() {
  var data = collectSheetData("Settings", "A2:B");
  return makeRoundPairings(data);
}

function parseFixedPairing(p) {
  if (p.startsWith("#")) {
    return { standing: parseInt(p.slice(1)) };
  } else {
    const regex = /\s\(.*$/;
    p = p.replace(regex, '');
    return { name: p };
  }
}

function makeFixedPairings(rows) {
  var fp = {};
  var fs = {};
  for (var entry of rows) {
    var round = parseInt(entry[0]);
    var p1 = parseFixedPairing(entry[1]);
    var p2 = parseFixedPairing(entry[2]);
    if (fp[round] === undefined) {
      fp[round] = [];
    }
    fp[round].push({first: p1, second: p2})
    if (`${entry[3]}`.toLowerCase().trim() == "yes") {
      fs[[round, p1.name]] = true;
    }
  }
  return new Fixed(fp, fs);
}

function collectFixedPairings() {
  var data = collectSheetData("FixedPairing", "A2:D");
  return makeFixedPairings(data);
}

function makeSettings(rows) {
  var ret = new Settings();
  for (var entry of rows) {
    const k = entry[0];
    const v = entry[1];
    const val = parseInt(v);
    if (!isNaN(val)) {
      if (k == "swiss_distance") {
        ret.swiss_distance = val
      } else if (k == "swiss_weight") {
        ret.swiss_weight = val
      }
    }
  }
  return ret;
}

function collectSettings() {
  var data = collectSheetData("Settings", "D2:E");
  return makeSettings(data);
}

// -----------------------------------------------------
// Filter data based on round

function standingsAfterRound(res, entrants, round) {
  console.log("standings at round:", round);
  // Calculate standings as of round <round>
  if (round == 0) {
    return entrants.seeding.map(x => res.getPlayerResults(x.name))
  }
  var tmp_res = new Results();
  tmp_res.results = res.results.filter(function (r) { return r.round <= round; });
  tmp_res.processResults();
  var standings = Object.values(tmp_res.players);
  standings.sort(_player_standings_sort);
  standings = getCurrentEntrantsRanking(res, entrants, standings);
  return standings
}

function getFixedPairing(standings, p) {
  if (p.standing !== undefined) {
    return standings[p.standing - 1].name;
  } else {
    return p.name;
  }
}

function removeFixedPairings(standings, entrants, round) {
  var fp = entrants.fixed_pairings[round];
  console.log("round:", round);
  console.log("fp:", fp);
  var remove = {};
  var fixed = [];
  if (fp !== undefined) {
    for (var pair of fp) {
      var p1 = getFixedPairing(standings, pair.first);
      var p2 = getFixedPairing(standings, pair.second);
      if (p1 != p2) {
        console.log("p1, p2:", [p1, p2]);
        [p1, p2] = [p1, p2].sort();
        console.log("sorted:", [p1, p2]);
        remove[p1] = p2;
        remove[p2] = p1;
        fixed.push({first: {name: p1}, second: {name: p2}});
      }
    }
  }
  //console.log("pairing:", remove);
  standings = standings.filter(p => remove[p.name] === undefined);
  var bye_pair = pairBye(standings);
  if (bye_pair.length > 0) {
    fixed.push({first: bye_pair[0], second: bye_pair[1]})
    standings = standings.filter(p => p.name != bye_pair[0].name);
    standings = standings.filter(p => p.name != bye_pair[1].name);
  }
  //console.log("new standings:", standings);
  return [standings, fixed];
}

function isBye(p) {
  return p.name.toLowerCase() == "bye";
}

function _player_byes_sort(x, y) {
  if (x.byes == y.byes) {
    if (x.score == y.score) {
      return (x.spread - y.spread);
    } else {
      return (x.score - y.score);
    }
  } else {
    return (x.byes - y.byes)
  }
}

function pairBye(standings) {
  var bye = standings.filter(isBye);
  if (bye.length == 0) {
    return []
  }
  var candidates = standings.filter((p) => !isBye(p));
  for (var p of candidates) {
    p.byes = BYES.get(p.name);
  }
  candidates.sort(_player_byes_sort);
  console.log("pairing bye:", candidates)
  return [bye[0], candidates[0]]
}

function getCurrentEntrantsRanking(res, entrants, standings) {
  // Get standings only for people in current entrants list
  var ps = standings.filter(p => p.name in entrants.entrants);
  // Add standings for new entrants with no results
  var existing = ps.map(p => p.name);
  var newcomers = entrants.seeding.filter(p => existing.indexOf(p.name) == -1);
  newcomers = newcomers.map(p => res.getPlayerResults(p.name));
  console.log("newcomers:", newcomers);
  var all = ps.concat(newcomers);
  return all
}

function pairingsAfterRound(res, entrants, repeats, round_pairings, round) {
  var standings;
  console.log("round_pairings:", round)
  var pair = round_pairings[round + 1];
  if (pair.type == "K") {
    standings = standingsAfterRound(res, entrants, round);
    return pairKoth(standings, entrants, round)
  } else if (pair.type == "Q") {
    standings = standingsAfterRound(res, entrants, round);
    return pairQoth(standings, entrants, round)
  } else if (pair.type == "RAND") {
    standings = standingsAfterRound(res, entrants, round);
    return pairRandom(standings, entrants, round);
  } else if (pair.type == "RANDNR") {
    return pairRandomNoRepeats(res, entrants, repeats, round);
  } else if (pair.type == "R") {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairRoundRobin(standings, pair.pos)
  } else if (pair.type == "DR") {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairDoubleRoundRobin(standings, pair.pos)
  } else if (pair.type.startsWith("QD")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairDistributedQuads(standings, pair.pos);
  } else if (pair.type.startsWith("QC")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairClusteredQuads(standings, pair.pos);
  } else if (pair.type.startsWith("QE")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairEvansQuads(standings, pair.pos);
  } else if (pair.type.startsWith("SIX")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairSixes(standings, pair.pos);
  } else if (pair.type == "CH") {
    return pairCharlottesville(entrants, round);
  } else if (pair.type == "SPR") {
    return pairSwissPlusRandom(res, entrants, repeats, round, round + 1);
  } else if (pair.type == "STPR") {
    return pairSwissPlusRandom(res, entrants, repeats, round - 1, round + 1);
  } else if (pair.type == "S") {
    return pairSwiss(res, entrants, repeats, round, round + 1);
  } else if (pair.type == "ST") {
    return pairSwiss(res, entrants, repeats, round - 1, round + 1);
  }
}

// -----------------------------------------------------
// Round robin pairing.
// See https://github.com/domino14/liwords/ for strategy

function _pairRR(n, r) {
  // Pair n players at round r
  var init = Array.from({ length: n - 1 }, (_, i) => i + 1);
  var h = n / 2;
  //var start = (r * (n - 3)) % (n - 1);
  var start = n - 1 - r;
  var r1 = init.slice(0, start);
  var r2 = init.slice(start);
  var rotated = [0].concat(r2, r1);
  var h1 = rotated.slice(0, h);
  var h2 = rotated.slice(h).reverse();
  return [h1, h2];
}

function pairRoundRobin(standings, pos) {
  // Pair for game #pos in the round robin
  var n = standings.length;
  var pairings = [];
  var [h1, h2] = _pairRR(n, pos - 1);
  for (var i = 0; i < standings.length / 2; i += 1) {
    pairings.push({ first: standings[h1[i]], second: standings[h2[i]] });
  }
  return pairings
}

function pairDoubleRoundRobin(standings, pos) {
  // Pair for game #pos in a double round robin where everyone plays everyone
  // else twice, with repeated games played consecutively.
  var n = standings.length;
  var pairings = [];
  pos = Math.round(pos / 2);
  var [h1, h2] = _pairRR(n, pos - 1);
  for (var i = 0; i < standings.length / 2; i += 1) {
    pairings.push({ first: standings[h1[i]], second: standings[h2[i]] });
  }
  return pairings
}

// -----------------------------------------------------
// Random pairing.

function shuffleArray(array) {
  // In-place Fisher-Yates variant, see
  // https://stackoverflow.com/questions/2450954/how-to-randomize-shuffle-a-javascript-array
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}

function pairRandom(standings, entrants, round) {
  var fixed
  [standings, fixed] = removeFixedPairings(standings, entrants, round + 1);
  shuffleArray(standings);
  // Now this is just KotH on the shuffled standings
  var pairings = [];
  for (var i = 0; i < standings.length; i += 2) {
    pairings.push({ first: standings[i], second: standings[i + 1] })
  }
  for (var p of fixed) {
    pairings.push(p);
  }
  return pairings
}

function pairRandomNoRepeatsHelper(bracket, repeats) {
  var pairings = []
  var edges = [];
  var names = {};
  var inames = {};
  var i = 0;
  for (var player of bracket) {
    names[player.name] = i;
    inames[i] = player.name;
    i++;
  }
  for (var p1 of bracket) {
    for (var p2 of bracket) {
      if (p1.name < p2.name) {
        let reps = repeats.get(p1.name, p2.name);
        if (reps === undefined) {
          reps = 0;
        }
        let weight = -(10 * reps + Math.random())
        let v1 = names[p1.name];
        let v2 = names[p2.name];
        edges.push([v1, v2, weight]);
      }
    }
  }
  var b = blossom(edges, true)
  for (var i = 0; i < b.length; i++) {
    let v = b[i];
    let p1 = inames[i];
    let p2 = inames[v];
    if (p1 < p2) {
      pairings.push({ first: { name: p1 }, second: { name: p2 } })
    }
  }
  return pairings
}

function pairRandomNoRepeats(results, entrants, repeats, round) {
  console.log("random no repeats pairing for round", round + 1)
  var players = standingsAfterRound(results, entrants, round);
  if (round <= 0) {
    return pairRandom(players, entrants, round);
  }
  var bracket, fixed;
  [bracket, fixed] = removeFixedPairings(players, entrants, round + 1);
  var pairings = pairRandomNoRepeatsHelper(bracket, repeats)
  for (var p of fixed) {
    pairings.push(p);
  }
  return pairings
}

function pairSwissPlusRandom(results, entrants, repeats, round, for_round) {
  console.log("swiss + randnr pairing based on round", round)
  var players = standingsAfterRound(results, entrants, round);
  if (round <= 0) {
    return pairSwissInitial(players, entrants, round);
  }
  var field, fixed;
  [field, fixed] = removeFixedPairings(players, entrants, for_round);
  var n = entrants.settings.swiss_distance;
  var swiss_players = field.slice(0, n);
  var rand_players = field.slice(n);

  console.log("pairing swiss for top", n, "players")
  var swiss_paired = pairSwissHelper(swiss_players, repeats, entrants)
  console.log("pairing random for bottom", field.length - n, "players")
  var random_pairings = pairRandomNoRepeatsHelper(rand_players, repeats)
  var out = []
  for (const group of swiss_paired) {
    for (var p of group) {
      if (p.first.name < p.second.name) {
        p.repeats = repeats.get(p.first.name, p.second.name)
        out.push(p)
      }
    }
  }
  for (var p of random_pairings) {
    out.push(p);
  }
  for (var p of fixed) {
    out.push(p);
  }
  console.log("out:", out)
  return out
}

// -----------------------------------------------------
// King of the hill pairing.

function pairKoth(standings, entrants, round) {
  // Sort by wins and spread
  var fixed
  [standings, fixed] = removeFixedPairings(standings, entrants, round + 1);
  var pairings = [];
  for (var i = 0; i < standings.length; i += 2) {
    pairings.push({ first: standings[i], second: standings[i + 1] })
  }
  for (var p of fixed) {
    pairings.push(p);
  }
  return pairings
}

// -----------------------------------------------------
// Queen of the hill pairing.

function pairQoth(standings, entrants, round) {
  // Sort by wins and spread
  var fixed
  [standings, fixed] = removeFixedPairings(standings, entrants, round + 1);
  var pairings = [];
  var n = standings.length;
  if (n % 4 == 2) {
    var last = n - 6
    for (var i = 0; i < last; i += 4) {
      pairings.push({ first: standings[i], second: standings[i + 2] })
      pairings.push({ first: standings[i + 1], second: standings[i + 3] })
    }
    // Pair the last six players 1-4,2-5,3-6 if we don't have a multiple of 4
    pairings.push({ first: standings[last + 0], second: standings[last + 3] })
    pairings.push({ first: standings[last + 1], second: standings[last + 4] })
    pairings.push({ first: standings[last + 2], second: standings[last + 5] })
  } else {
    for (var i = 0; i < n; i += 4) {
      pairings.push({ first: standings[i], second: standings[i + 2] })
      pairings.push({ first: standings[i + 1], second: standings[i + 3] })
    }
  }
  for (var p of fixed) {
    pairings.push(p);
  }
  return pairings
}

// -----------------------------------------------------
// Quad pairing.

// We assume there are always an even number of players (one of whom might be 'bye'),
// but there might not be a divisible-by-four number. If there are 4n+2 players, we
// divide them into (n-1) quads and a final hex, and pair the hex separately in a
// group of 3 games.

// Quad pairings for four players, 0-3
const Pairings4 = [
  [[0, 3], [1, 2]],
  [[0, 2], [1, 3]],
  [[0, 1], [2, 3]],
  // Repeat if we are doing a hex
  [[0, 3], [1, 2]],
  [[0, 2], [1, 3]],
  [[0, 1], [2, 3]]
]

// Incomplete round robin for 6 players, 0-5
const Pairings6 =
[
  [[0, 5], [1, 4], [2, 3]],
  [[0, 4], [5, 3], [1, 2]],
  [[0, 3], [4, 2], [5, 1]],
  [[0, 2], [3, 1], [4, 5]],
  [[0, 1], [2, 5], [3, 4]]
]

function groupPositionPairs(group, pos) {
  if (group.length == 4) {
    return Pairings4[pos - 1]
  } else {
    return Pairings6[pos - 1]
  }
}

function pairGroupsAtPosition(groups, pos) {
  var pairings = [];
  for (var i = 0; i < groups.length; i++) {
    const group = groups[i];
    var p = groupPositionPairs(group, pos);
    for (let [a, b] of p) {
      pairings.push({ first: group[a], second: group[b] });
    }
  }
  return pairings;
}

function getLastQuadPosition(standings) {
  var leftover = standings.length % 4;
  console.log("leftover:", leftover);
  if (leftover == 0) {
    return standings.length
  } else if (leftover == 2) {
    return standings.length - 6
  }
}

function getLastHexPosition(standings) {
  var leftover = standings.length % 6;
  console.log("leftover:", leftover);
  if (leftover == 0) {
    return standings.length
  } else if (leftover == 2) {
    return standings.length - 8
  } else if (leftover == 4) {
    return standings.length - 4
  }
}

function maybeAddHex(quads, standings, max) {
  // we have a leftover hex, add it to the quads
  if (max < standings.length) {
    quads.push(standings.slice(max, standings.length))
  }
}

function maybeAddQuads(hexes, standings, max) {
  // we have leftover players not in hexes, add them in quads
  var diff = standings.length - max;
  if (diff == 8) {
    hexes.push(standings.slice(max, max + 4));
    hexes.push(standings.slice(max + 4, standings.length));
  } else if (diff == 4) {
    hexes.push(standings.slice(max, max + 4));
  }
}

function pairClusteredQuads(standings, pos) {
  var quads = [];
  var max = getLastQuadPosition(standings);
  for (var i = 0; i < max; i += 4) {
    quads.push(standings.slice(i, i + 4));
  }
  maybeAddHex(quads, standings, max);
  return pairGroupsAtPosition(quads, pos);
}

function distributeQuads(standings, max) {
  // Deal the top `max` players across max/4 quads, so that each quad gets one
  // player from each quarter of the field.
  var quads: any[] = [];
  var stride = max / 4;
  for (var i = 0; i < stride; i++) {
    quads[i] = [];
  }
  for (var i = 0; i < max; i++) {
    var quad = i % stride;
    quads[quad].push(standings[i]);
  }
  return quads;
}

function distributeQuadsWithHex(standings, max) {
  // The distributed counterpart of maybeAddHex: rather than making the hex out
  // of the bottom six players, we spread its members through the field the same
  // way a quad's are. The standings are split into six bands - four of
  // (stride + 1) players that every group draws from, and two single-player
  // bands in mid-field that only the hex draws from.
  //
  // e.g. for 14 players (two quads and a hex) the bands are
  //   [1 2 3] [4 5 6] [7] [8] [9 10 11] [12 13 14]
  // giving quads of 1/4/9/12 and 2/5/10/13, and a hex of 3/6/7/8/11/14.
  var stride = max / 4;
  var groups: any[] = [];
  for (var i = 0; i <= stride; i++) {
    groups[i] = [];
  }
  const bands = [stride + 1, stride + 1, 1, 1, stride + 1, stride + 1];
  var player = 0;
  for (const band of bands) {
    for (var i = 0; i < band; i++) {
      // a full band feeds every group in turn, a single-player band only the hex
      var group = (band == 1) ? stride : i;
      groups[group].push(standings[player]);
      player++;
    }
  }
  return groups;
}

function pairDistributedQuads(standings, pos) {
  var max = getLastQuadPosition(standings);
  var quads = (max < standings.length)
    ? distributeQuadsWithHex(standings, max)
    : distributeQuads(standings, max);
  console.log("quads:", quads)
  console.log("standings:", standings)
  return pairGroupsAtPosition(quads, pos);
}

function pairEvansQuads(standings, pos) {
  // Like distributed quads but flip every other subgroup first,
  // so that the sum of opponent seeds ends up roughly equal.
  // e.g. for 12 people you would make quads from
  // 1 2 3 6 5 4 7 8 9 12 11 10
  var quads = [];
  var max = getLastQuadPosition(standings);
  var stride = max / 4;
  for (var i = 0; i < stride; i++) {
    quads[i] = [];
  }

  // Generate new standings snake-style
  var new_standings = []
  var flip = false;
  for (var i = 0; i < max; i += stride) {
    var slice = standings.slice(i, i + stride);
    if (flip) {
      slice.reverse();
    }
    flip = !flip;
    new_standings = new_standings.concat(slice)
  }

  // Make quads from the new standings
  for (var i = 0; i < max; i++) {
    var quad = i % stride;
    quads[quad].push(new_standings[i]);
  }
  maybeAddHex(quads, standings, max);
  return pairGroupsAtPosition(quads, pos);
}

function pairSixes(standings, pos) {
  // Like evans quads, but groups of six
  var hexes = [];
  var max = getLastHexPosition(standings);
  var stride = max / 6;
  for (var i = 0; i < stride; i++) {
    hexes[i] = [];
  }

  // Generate new standings snake-style
  var new_standings = []
  var flip = false;
  for (var i = 0; i < max; i += stride) {
    var slice = standings.slice(i, i + stride);
    if (flip) {
      slice.reverse();
    }
    flip = !flip;
    new_standings = new_standings.concat(slice)
  }

  // Make hexes from the new standings
  for (var i = 0; i < max; i++) {
    var hex = i % stride;
    hexes[hex].push(new_standings[i]);
  }
  maybeAddQuads(hexes, standings, max);
  return pairGroupsAtPosition(hexes, pos);
}

// -----------------------------------------------------
// Charlottesville pairing.

// Split the field into 2 groups in a snake order.
// Group 1: 1, 4, 5, 8, 9, 12, 13, 16, 17
// Group 2: 2, 3, 6, 7, 10, 11, 14, 15, 18
// For the first 9 rounds, you play a round robin against all the people in the *other* group.

function pairCharlottesville(entrants, round) {
  var g1 = [];
  var g2 = [];
  for (var i = 1; i <= entrants.seeding.length; i += 1) {
    if (i % 4 == 0 || i % 4 == 1) {
      g1.push(i);
    } else {
      g2.push(i);
    }
  }
  // reverse group 2 so the top player plays the second player last
  g2.reverse();
  // rotate group 2 one place per round and pair up with group 1
  var r = round - 1;
  var r1 = g2.slice(0, r);
  var r2 = g2.slice(r);
  var rotated = r2.concat(r1);
  var pairings = [];
  for (var i = 0; i < g1.length; i += 1) {
    var p1 = g1[i] - 1;
    var p2 = rotated[i] - 1;
    pairings.push({ first: entrants.seeding[p1], second: entrants.seeding[p2] });
  }
  return pairings
}

// -----------------------------------------------------
// Swiss pairing.

function calculateScoreGroups(standings) {
  var groups: any[] = []
  console.log("standings:", standings)
  for (var p of standings) {
    var k = p.wins
    if (groups[k] === undefined) {
      groups[k] = []
    }
    groups[k].push(p)
  }
  groups = groups.filter(e => !!e).reverse()
  // Balance groups
  var curr, next
  for (var i = 0; i < groups.length - 1; i++) {
    [curr, next] = [groups[i], groups[i + 1]]
    if (curr.length % 2 != 0) {
      var fst = next.shift();
      curr.push(fst);
    }
  }
  groups = groups.filter(e => e.length != 0)
  return groups;
}

function promote(groups, i, j) {
  if (groups[j] === undefined) {
    // We have run out of groups to promote players out of
    throw new Error("swiss pairing failed: no group " + j + " to promote from, " +
      "groups are " + JSON.stringify(groups.map(g => g.length)));
  }
  var fst = groups[j].shift();
  groups[i].push(fst)
}

function promote2(groups, i) {
  console.log("promoting two into", i);
  var j = i + 1
  promote(groups, i, j);
  if (groups[j].length == 0) {
    promote(groups, i, j + 1)
  } else {
    promote(groups, i, j)
  }
}

function mergeBottom(groups) {
  console.log("merging bottom two groups");
  if (groups.length == 1) {
    console.log("only one group, bailing out!")
    return;
  }
  var last = groups.length - 1;
  groups[last - 1] = groups[last - 1].concat(groups[last]);
  groups[last] = [];
}

function pairSwissInitial(standings) {
  var pairings = [];
  const half = standings.length / 2;
  for (var i = 0; i < half; i++) {
    pairings.push({ first: standings[i], second: standings[i + half] })
  }
  return pairings
}

function pairSwissTop(groups, repeats, nrep) {
  var top = groups[0];
  var candidates = [];
  for (var i = 0; i < top.length; i++) {
    candidates[i] = [];
    for (var j = 0; j < top.length; j++) {
      if (i == j) {
        continue;
      }
      var reps = repeats.get(top[i].name, top[j].name);
      if (reps < nrep) {
        candidates[i].push([reps, Math.abs(i - j), top[j].name, top[i].name])
      }
    }
  }
  for (var i = 0; i < candidates.length; i++) {
    candidates[i] = candidates[i].sort()
  }
  return candidates
}

function pairCandidates(bracket, settings) {
  // console.log("candidates", candidates);
  var edges = [];
  var names = {};
  var inames = {};
  var i = 0;
  for (var player of bracket) {
    let name = player[0][3]
    names[name] = i;
    inames[i] = name;
    i++;
  }
  // console.log("names", names)
  // console.log("inames", inames)
  console.log("swiss settings:", settings)

  for (var player of bracket) {
    for (var m of player) {
      const [repeats, distance, p1, p2] = m;
      // Don't pair candidates too far apart
      if (distance < settings.swiss_distance) {
        let weight = -(settings.swiss_weight * repeats + distance);
        let v1 = names[p1];
        let v2 = names[p2];
        edges.push([v1, v2, weight])
      }
    }
    // var name = player[0][3]
    // console.log(name, player.length, player)
  }
  // console.log("edges:", edges)
  var b = blossom(edges, true)
  // console.log("blossom:", b)
  var pairings = []
  for (var i = 0; i < b.length; i++) {
    let v = b[i];
    let p1 = inames[i];
    let p2 = inames[v];
    pairings.push({ first: { name: p1 }, second: { name: p2 } })
  }
  // console.log("sub pairing:", pairings)
  return pairings
}

function pairSwissHelper(players, repeats, entrants) {
  var groups = calculateScoreGroups(players);
  var dgroups = groups.map(g => g.map(p => [p.name, p.wins]));
  // console.log("groups:", dgroups)
  var candidates;
  var nrep = 1;
  var paired = [];
  // Don't have too small a bottom group. Stop once everything has been merged
  // into a single group, otherwise mergeBottom has nothing left to merge and we
  // spin forever (a field of fewer than six players never reaches the target).
  while (groups.length > 1 && groups[groups.length - 1].length < 6) {
    mergeBottom(groups);
    groups = groups.filter(e => e.length != 0);
  }
  // nrep is the limit on how many times a pair may already have met. Once it is
  // past the highest repeat count in the field every pair is a candidate, so if
  // we still can't pair the group there is nothing left to relax and retrying
  // would loop forever - e.g. a group with an odd number of players can never
  // be matched up no matter how many repeats we allow.
  const max_reps = repeats.max();
  while (groups.length > 0) {
    if (nrep > max_reps + 1) {
      throw new Error("swiss pairing failed: cannot pair group of " +
        groups[0].length + " players (" + groups[0].map(p => p.name).join(", ") + ")");
    }
    dgroups = groups.map(g => g.map(p => [p.name, p.wins]));
    // console.log("groups:", dgroups)
    candidates = pairSwissTop(groups, repeats, nrep)
    //console.log("candidates:", candidates)
    if (candidates.some(e => e.length == 0)) {
      if (groups.length == 1) {
        console.log("failed!")
        nrep += 1;
        console.log("reps:", nrep);
        continue;
      }
      promote2(groups, 0)
      groups = groups.filter(e => e.length != 0)
      if (groups.length == 1) {
        console.log("failed! after promotion")
        nrep += 1;
        console.log("reps:", nrep);
        continue;
      }
    } else {
      var pairs = pairCandidates(candidates, entrants.settings)
      if (pairs.some(e => e.second.name === undefined)) {
        console.log("unpaired!")
        nrep += 1;
        console.log("reps:", nrep);
        continue;
      }
      groups.shift()
      paired.push(pairs);
      if (groups.length == 0) {
        break;
      }
    }
  }
  return paired;
}

function pairSwiss(results, entrants, repeats, round, for_round) {
  console.log("swiss pairing based on round", round)
  if (round <= 0) {
    return pairSwissInitial(entrants.seeding);
  }
  //console.log("repeats for round", round, repeats.matches)
  var players = standingsAfterRound(results, entrants, round);
  var fixed;
  [players, fixed] = removeFixedPairings(players, entrants, for_round);
  var paired = pairSwissHelper(players, repeats, entrants);
  console.log("fixed:", fixed)
  paired.push(fixed);
  var out = []
  for (const group of paired) {
    for (var p of group) {
      if (p.first.name < p.second.name) {
        p.repeats = repeats.get(p.first.name, p.second.name)
        out.push(p)
      } else if (p.first.name == "Bye") {
        // We only have one pairing for this, so always push it.
        p.repeats = repeats.get(p.first.name, p.second.name)
        out.push(p)
      }
    }
  }
  console.log("out:", out)
  return out
}

// -----------------------------------------------------
// Output pairings and standings

function _player_standings_sort(x, y) {
  if (x.score == y.score) {
    return (y.spread - x.spread);
  } else {
    return (y.score - x.score);
  }
}

function outputPlayerStandings(standing_sheet, score_dict, entrants, ratings) {
  // Sort by wins and spread
  var standings = Object.values(score_dict);
  standings.sort(_player_standings_sort);
  standings = standings.filter(x => x.name.toLowerCase() != "bye");
  standings = standings.filter(x => !x.name.includes("bye"));
  var out = standings.map(function (x, index) {
    var full_name = entrants.entrants[x.name] || x.name;
    var rating = ratings[x.name]
    return [
      (index + 1) + ".",
      full_name,
      x.wins + 0.5 * x.ties,
      x.losses + 0.5 * x.ties,
      x.spread,
      rating,
      "",
      x.starts,
    ]
  })

  // Write out standings starting in cell A2
  var outputRow = 2;
  var outputCol = 1;
  if (standings.length == 0) {
    return;
  }
  var outputRange = standing_sheet.getRange(outputRow, outputCol, out.length, out[0].length);
  outputRange.setValues(out);
}

function formatPairings(pairings, entrants, round) {
  var vtable = 1;
  var used = new Set();
  for (var v of Object.values(entrants.tables)) {
    used.add(v)
  }
  var out = pairings.map(function (x, index) {
    var table;
    if (x.first.name in entrants.tables) {
      table = entrants.tables[x.first.name];
    } else if (x.second.name in entrants.tables) {
      table = entrants.tables[x.second.name];
    } else {
      while (used.has(vtable)) {
        vtable++;
      }
      table = vtable;
      vtable++;
    }
    var first = x.first.name;
    var second = x.second.name;
    var rep = x.repeats > 1 ? `(rep ${x.repeats})` : "";
    var start = `${x.first.starts} - ${x.second.starts}`
    return [
      table,
      entrants.entrants[first] || first,
      entrants.entrants[second] || second,
      rep,
      start,
    ]
  })
  out = out.sort((a, b) => parseInt(a) - parseInt(b));
  return out;
}

function outputPairings(pairing_sheet, text_pairing_sheet, pairings, entrants, starts, round, start_row) {
  console.log("pairings:", pairings);
  var out = formatPairings(pairings, entrants, starts, round);
  var ncols = out[0].length
  var text_pairings = []
  var round_header = "ROUND " + (round + 1);
  var pairing_strings = pairings.map(x => `${x.first.name} v. ${x.second.name}`);
  var pairing_string = round_header + ": " + pairing_strings.join(" | ");
  text_pairings.push([pairing_string])
  var header = [
    [round_header, "", "", "", ""],
    ["Table", "1st Player", "2nd Player", "", "Firsts"],
  ];
  out = header.concat(out);
  out = out.concat([["", "", "", "", ""]]);
  // Write out standings starting in cell A2
  var outputRow = start_row;
  var outputCol = 1;
  var outputRange = pairing_sheet.getRange(outputRow, outputCol, out.length, ncols);
  outputRange.setValues(out);
  outputRange.setFontWeight("normal");
  outputRange.setBorder(true, true, true, true, true, true, "white", null);
  outputRange.setHorizontalAlignment("left");
  var header1Range = pairing_sheet.getRange(outputRow, outputCol, 1, ncols);
  var header2Range = pairing_sheet.getRange(outputRow + 1, outputCol, 1, ncols);
  const header1bg = "#531882"
  const header2bg = "#daabff"
  header1Range.setFontWeight("bold");
  header1Range.setFontColor("white");
  header1Range.setBackground(header1bg);
  header1Range.setBorder(true, true, true, true, true, true, header1bg, null);
  header2Range.setFontColor("black");
  header2Range.setFontWeight("bold");
  header2Range.setBackground(header2bg);
  header2Range.setBorder(true, true, true, true, true, true, header2bg, null);
  var textPairingRange = text_pairing_sheet.getRange(round + 1, 1, 1, 1);
  textPairingRange.setValues(text_pairings);
}

function _show_win(g) {
  if (g === undefined) {
    return ""
  }
  return `${g.winner} ${g.winner_score} - ${g.loser} ${g.loser_score}`
}

function _show_loss(g) {
  if (g === undefined) {
    return ""
  }
  return `${g.loser} ${g.loser_score} - ${g.winner} ${g.winner_score}`
}

function _show_high_game(g) {
  if (g === undefined) {
    return ""
  }
  var total = g.winner_score + g.loser_score;
  return `${g.winner} ${g.winner_score} - ${g.loser} ${g.loser_score} (${total})`
}

function _show_spread(g) {
  if (g === undefined) {
    return ""
  }
  var spread = g.winner_score - g.loser_score;
  return `${g.winner} ${g.winner_score} - ${g.loser} ${g.loser_score} (${spread})`
}

function _show_high_spread(g, entrants) {
  if (g === undefined) {
    return ""
  }
  var spread = g.winner_score - g.loser_score;
  var winner_name = entrants.entrants[g.winner] || `${g.winner} (#?)`;
  var loser_name = entrants.entrants[g.loser] || `${g.loser} (#?)`;
  return `${winner_name} ${g.winner_score} - ${loser_name} ${g.loser_score}: (${spread})`
}

function outputStatistics(statistics_sheet, res, entrants) {
  var results = res.results
  results = [...results].filter((x) => x.loser != 'Bye');
  var no_ties = [...results].filter((x) => x.winner_score > x.loser_score);
  var high_game = [...results].sort((x, y) =>
    y.winner_score + y.loser_score - x.winner_score - x.loser_score);
  var high_win = [...results].sort((x, y) => y.winner_score - x.winner_score);
  var low_win = [...results].sort((x, y) => x.winner_score - y.winner_score);
  var high_loss = [...no_ties].sort((x, y) => y.loser_score - x.loser_score);
  var low_loss = [...no_ties].sort((x, y) => x.loser_score - y.loser_score);
  var tie = [...results].filter((x) => x.winner_score == x.loser_score).sort((x, y) => y.loser_score - x.loser_score);
  var high_spread = [...no_ties].sort((x, y) =>
    y.winner_score - y.loser_score - x.winner_score + x.loser_score);
  var low_spread = [...no_ties].sort((x, y) =>
    x.winner_score - x.loser_score - y.winner_score + y.loser_score);
  var out = [["Highest Wins", "Highest Losses", "Lowest Wins", "Lowest Losses"]]
  for (var i = 0; i < 10; i++) {
    out.push([
      _show_win(high_win[i]),
      _show_loss(high_loss[i]),
      _show_win(low_win[i]),
      _show_loss(low_loss[i])
    ].map((x) => x == "" ? "" : `${i + 1}. ${x}`))
  }
  out.push(["", "", "", ""]);
  out.push(["", "", "", ""]);

  out.push(["Highest Games", "Biggest Blowouts", "Nailbiters", "Ties"]);
  for (var i = 0; i < 10; i++) {
    out.push([
      _show_high_game(high_game[i]),
      _show_spread(high_spread[i], entrants),
      _show_spread(low_spread[i]),
      _show_win(tie[i])
    ].map((x) => x == "" ? "" : `${i + 1}. ${x}`))
  }
  console.log(out)

  // Write out standings starting in cell A2
  var outputRow = 1;
  var outputCol = 1;
  statistics_sheet.clearContents();
  var outputRange = statistics_sheet.getRange(outputRow, outputCol, out.length, out[0].length);
  outputRange.setValues(out);
}

function getLastRound(res, round_pairings) {
  // Find the last round we can pair
  var round_ids = res.roundIds();
  console.log("round ids:", round_ids);
  var last_result = Math.max(...round_ids);
  var last_round = 0;
  for (var r of Object.values(round_pairings)) {
    if ((r.start - 1 <= last_result) ||
      (r.type.startsWith("RAND") && r.round - last_round == 1) ||
      (r.type == "R" && r.start - 1 <= last_round)) {
      console.log("Pairing round", r, "based on round", r.start);
      last_round = r.round;
    }
  }
  console.log("Last round", last_round);
  return last_round;
}

function runPairings(res, entrants, round_pairings, starts, diags) {
  var repeats = new Repeats();
  var round_ids = res.roundIds();
  BYES.reset();
  var all_pairings = [];
  var last_round = getLastRound(res, round_pairings);
  for (var i = 0; i < last_round; i++) {
    var round = i + 1;
    var rp = round_pairings[round];
    var pairings;
    var unplayed = res.unplayedGames(round);
    if ((round < round_ids.length) && unplayed.size == 0) {
      diags.round_status[round] = "complete";
      pairings = res.extractPairings(round)
      for (var p of pairings) {
        starts.register(p, round);
        // Add number of starts to the pairing
        p.first.starts = starts.starts[p.first.name];
        p.second.starts = starts.starts[p.second.name];
      }
    } else {
      diags.round_status[round] = `incomplete: ${unplayed}`;
      for (var k of Object.keys(res.players)) {
        var b = BYES.get(k);
        if (b > 0) {
          console.log(`byes for ${k}: ${b}`);
        }
      }
      var orig_pairings = pairingsAfterRound(res, entrants, repeats, round_pairings, i);
      pairings = []
      for (var p0 of orig_pairings) {
        // Copy the object so we don't keep modifying the same reference
        // (pairingsAfterRound pushes player references into the output pairs)
        var p = {first: {name: p0.first.name}, second: {name: p0.second.name}}
        // See who starts in this round
        var p1_first = starts.add(p.first.name, p.second.name, round);
        // Make p.first the first player
        if (!p1_first) {
          [p.first, p.second] = [p.second, p.first]
        }
        // Add number of starts to the pairing
        p.first.starts = starts.starts[p.first.name];
        p.second.starts = starts.starts[p.second.name];
        pairings.push(p);
      }
    }
    for (var p of pairings) {
      BYES.update(p);
      p.repeats = repeats.add(p.first.name, p.second.name);
    }
    all_pairings.push(pairings);
  }
  return all_pairings;
}

function lastResultTimestamp() {
  var gs = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = gs.getSheetByName("Results");
  var row = sheet.getLastRow();
  return sheet.getRange(row, 1).getValues()[0][0];
}

function outputPairingLog(sheet, all_pairings) {
  var last_row = sheet.getLastRow();
  // Write header if sheet is empty
  if (last_row == 0) {
    var header = [["Timestamp", "Last Result", "Round", "Pairings"]];
    var hdr_range = sheet.getRange(last_row + 1, 1, 1, header[0].length);
    hdr_range.setValues(header);
    hdr_range.setFontWeight("bold");
    last_row += 1;
  }
  // Write current pairings
  var outputRow = last_row + 2;
  var outputCol = 1;
  var now = new Date();
  var ts = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var result_ts = lastResultTimestamp();
  var out: any[] = [];
  var round = 0;
  for (var pairings of all_pairings) {
    var round_str = "Round " + (round + 1);
    var pairing_strings = pairings.map(x => `${x.first.name} v. ${x.second.name}`);
    var pairing_str = pairing_strings.join(" | ");
    out.push([ts, result_ts, round_str, pairing_str]);
    round = round + 1;
  }
  var out_range = sheet.getRange(outputRow, outputCol, out.length, out[0].length);
  out_range.setValues(out);
  sheet.getRange(outputRow, outputCol + 1, out.length, 1).setNumberFormat("M/dd/yyyy HH:mm:ss");
}

function outputDiags(diags_sheet, diags) {
  var out = [['Round', 'Status']];
  for (const [round, stat] of Object.entries(diags.round_status)) {
    out.push([`${round}`, stat]);
  }
  // Write out standings starting in cell A1
  var outputRow = 1;
  var outputCol = 1;
  if (out.length == 0) {
    return;
  }
  var outputRange = diags_sheet.getRange(outputRow, outputCol, out.length, out[0].length);
  outputRange.setValues(out);
}

function getOrCreateSheet(name) {
  var gs = SpreadsheetApp.getActiveSpreadsheet();
  var ret = gs.getSheetByName(name);
  if (ret == null) {
    ret = gs.insertSheet(name);
  }
  return ret
}

function processSheet(outputs) {
  var gs = SpreadsheetApp.getActiveSpreadsheet();
  var res = collectResults();

  var entrants = collectEntrants();
  var ratings = {};
  for (var x of entrants.seeding) {
    ratings[x.name] = x.rating
  }

  var fp = collectFixedPairings();
  entrants.fixed_pairings = fp.pairings;
  entrants.fixed_starts = fp.starts;

  var settings = collectSettings();
  console.log("collected settings:", settings)
  entrants.settings = settings;

  console.log("processed results");

  if (entrants.seeding.length % 2 != 0) {
    SpreadsheetApp.getUi().alert("Field has an odd number of players. Add a Bye if needed.");
    return;
  }

  var diags = new Diags();

  // Write out the standings
  var standings_sheet = gs.getSheetByName(outputs.standings);
  outputPlayerStandings(standings_sheet, res.players, entrants, ratings);

  // Write out statistics
  var statistics_sheet = gs.getSheetByName(outputs.statistics);
  outputStatistics(statistics_sheet, res, entrants);

  // Write out the pairings
  var round_pairings = collectRoundPairings();
  console.log("round pairings:", round_pairings);

  var pairing_sheet = gs.getSheetByName(outputs.pairings);
  var text_pairing_sheet = gs.getSheetByName(outputs.text_pairings);

  // Clear pairing sheets
  pairing_sheet.clear();
  text_pairing_sheet.clearContents();

  var starts = new Starts(res, entrants);
  var all_pairings = runPairings(res, entrants, round_pairings, starts, diags);
  var row = 2;
  var i = 0;
  for (var pairings of all_pairings) {
    outputPairings(pairing_sheet, text_pairing_sheet, pairings, entrants, starts, i, row);
    i += 1;
    row += pairings.length + 3;
  }
  pairing_sheet.getDataRange().setFontFamily("Proxima Nova").setFontSize(13);

  // Write audit trail
  var log_pairing_sheet = getOrCreateSheet("PairingsLog");
  outputPairingLog(log_pairing_sheet, all_pairings);

  // Write diagnostics
  var diags_sheet = getOrCreateSheet("Diagnostics");
  outputDiags(diags_sheet, diags);
}

function calculateStandings() {
  var outputs = {
    standings: "Standings",
    pairings: "Pairings",
    text_pairings: "Text Pairings",
    statistics: "Stats"
  }
  processSheet(outputs);
}

// export {
//   makeEntrants, makeRoundPairings, makeResults, makeFixedPairings,
//   pairingsAfterRound, standingsAfterRound, runPairings, formatPairings,
//   Repeats, Starts, Diags, Results
// };

// Version: 2026-07-31
