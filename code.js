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

  extractPairings(round) {
    var pairings = [];
    console.log("extracting pairings for round:", round)
    console.log(this.rounds[round])
    for (const game_result of this.rounds[round]) {
      var pairing = {
        first: {name: game_result.winner, start: game_result.winner_first},
        second: {name: game_result.loser, start: !game_result.winner_first}
      }
      pairings.push(pairing);
    }
    return pairings
  }
}

class Entrants {
  constructor(entrants, seeding, tables, fixed_pairings) {
    this.entrants = entrants
    this.seeding = seeding
    this.tables = tables
    this.fixed_pairings = fixed_pairings
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
}

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

// -----------------------------------------------------
// Read data from spreadsheet

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

function collectResults(result_sheet) {
  // Get the results range within the result sheet
  var result_range = result_sheet.getRange("B2:H");
  var results = result_range.getValues();
  var last_row = result_sheet.getRange("B2").getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var data = results.slice(0, last_row - 1);
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
  var ret = new Entrants(entrants, seeding, tables, fixed_pairings);
  for (var entry of rows) {
    ret.addEntrant(entrantFromRow(entry));
  }
  ret.seeding.sort(function (i, j) { return i.seed - j.seed });
  return ret;
}

function collectEntrants() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var result_sheet = sheet.getSheetByName("Entrants");
  var result_range = result_sheet.getRange("A2:E");
  var results = result_range.getValues();
  var last_row = result_sheet.getRange("A2").getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var data = results.slice(0, last_row - 1);
  var entrants = makeEntrants(data);
  console.log("Seeding:", entrants.seeding);
  console.log("Entrants:", entrants.entrants);
  console.log("Tables:", entrants.tables);
  return entrants;
}

function makeRoundPairings(rows) {
  var quads = {}
  var round_robins = {}
  var rounds = []
  for (var entry of rows) {
    // col A = entry[0], B = 1
    var round = parseInt(entry[0]);
    var pairing = entry[1];
    if (pairing.startsWith("Q")) {
      if (quads[pairing] === undefined) {
        quads[pairing] = [];
      }
      quads[pairing].push(round);
    } else if (pairing.startsWith("R")) {
      if (round_robins[pairing] === undefined) {
        round_robins[pairing] = [];
      }
      round_robins[pairing].push(round);
    } else if (pairing == "CH") {
      rounds[round] = { round: round, type: pairing, start: 0 };
    } else if (pairing == "ST") {
      rounds[round] = { round: round, type: pairing, start: round - 1 };
    } else {
      rounds[round] = { round: round, type: pairing, start: round };
    }
  }
  for (const q of Object.keys(quads)) {
    const quad = quads[q];
    for (i = 0; i < quad.length; i++) {
      round = quad[i];
      rounds[round] = { round: round, type: q, start: quad[0], pos: i + 1 }
    }
  }
  for (const r of Object.keys(round_robins)) {
    const rr = round_robins[r];
    for (i = 0; i < rr.length; i++) {
      round = rr[i];
      rounds[round] = { round: round, type: "R", start: rr[0], pos: i + 1 }
    }
  }
  return rounds;
}

function collectRoundPairings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var result_sheet = sheet.getSheetByName("RoundPairing");
  var result_range = result_sheet.getRange("A2:B");
  var results = result_range.getValues();
  var last_row = result_sheet.getRange("A2").getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var data = results.slice(0, last_row - 1);
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
  for (var entry of rows) {
    var round = parseInt(entry[0]);
    var p1 = parseFixedPairing(entry[1]);
    var p2 = parseFixedPairing(entry[2]);
    if (fp[round] === undefined) {
      fp[round] = [];
    }
    fp[round].push({first: p1, second: p2})
  }
  return fp;
}

function collectFixedPairings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var result_sheet = sheet.getSheetByName("FixedPairing");
  var result_range = result_sheet.getRange("A2:C");
  var results = result_range.getValues();
  var last_row = result_sheet.getRange("A2").getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var data = results.slice(0, last_row - 1);
  return makeFixedPairings(data);
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
  if (fp === undefined) {
    return [standings, []];
  }
  var remove = {};
  var fixed = [];
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
  //console.log("pairing:", remove);
  standings = standings.filter(p => remove[p.name] === undefined);
  //console.log("new standings:", standings);
  return [standings, fixed];
}

function getCurrentEntrantsRanking(res, entrants, standings) {
  // Get standings only for people in current entrants list
  var ps = standings.filter(p => p.name in entrants.entrants);
  // Add standings for new entrants with no results
  var existing = ps.map(p => p.name);
  var newcomers = entrants.seeding.filter(p => existing.indexOf(p.name) == -1);
  newcomers = newcomers.map(p => res.getPlayerResults(p.name));
  console.log("newcomers:", newcomers);
  all = ps.concat(newcomers);
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
  } else if (pair.type == "R") {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairRoundRobin(standings, pair.pos)
  } else if (pair.type.startsWith("QD")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairDistributedQuads(standings, pair.pos);
  } else if (pair.type.startsWith("QC")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairClusteredQuads(standings, pair.pos);
  } else if (pair.type.startsWith("QE")) {
    standings = standingsAfterRound(res, entrants, pair.start - 1);
    return pairEvansQuads(standings, pair.pos);
  } else if (pair.type == "CH") {
    return pairCharlottesville(entrants, round);
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
    last = n - 6
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
  [[0, 1], [2, 3]]
]

// Incomplete round robin for 6 players, 0-5
const Pairings6 = [
  [[0, 1], [2, 3], [4, 5]],
  [[0, 2], [3, 4], [1, 5]],
  [[0, 3], [1, 4], [2, 5]]
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
  for (i = 0; i < groups.length; i++) {
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

function maybeAddHex(quads, standings, max) {
  // we have a leftover hex, add it to the quads
  if (max < standings.length) {
    quads.push(standings.slice(max, standings.length))
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

function pairDistributedQuads(standings, pos) {
  var quads = [];
  var max = getLastQuadPosition(standings);
  var stride = max / 4;
  for (var i = 0; i < stride; i++) {
    quads[i] = [];
  }
  for (var i = 0; i < max; i++) {
    var quad = i % stride;
    quads[quad].push(standings[i]);
  }
  maybeAddHex(quads, standings, max);
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
    p1 = g1[i] - 1;
    p2 = rotated[i] - 1;
    pairings.push({ first: entrants.seeding[p1], second: entrants.seeding[p2] });
  }
  return pairings
}

// -----------------------------------------------------
// Graph matching algorithm
//
// From https://github.com/mattkrick/EdmondsBlossom
function blossom(edges, maxCardinality) {
  if (edges.length === 0) {
    return edges;
  }
  var edmonds = new Edmonds(edges, maxCardinality);
  return edmonds.maxWeightMatching();

};

var Edmonds = function (edges, maxCardinality) {
  this.edges = edges;
  this.maxCardinality = maxCardinality;
  this.nEdge = edges.length;
  this.init();
};

Edmonds.prototype.maxWeightMatching = function () {
  for (var t = 0; t < this.nVertex; t++) {
    //console.log('DEBUG: STAGE ' + t);
    this.label = filledArray(2 * this.nVertex, 0);
    this.bestEdge = filledArray(2 * this.nVertex, -1);
    this.blossomBestEdges = initArrArr(2 * this.nVertex);
    this.allowEdge = filledArray(this.nEdge, false);
    this.queue = [];
    for (var v = 0; v < this.nVertex; v++) {
      if (this.mate[v] === -1 && this.label[this.inBlossom[v]] === 0) {
        this.assignLabel(v, 1, -1);
      }
    }
    var augmented = false;
    while (true) {
      //console.log('DEBUG: SUBSTAGE');
      while (this.queue.length > 0 && !augmented) {
        v = this.queue.pop();
        //console.log('DEBUG: POP ', 'v=' + v);
        //console.assert(this.label[this.inBlossom[v]] == 1);
        for (var ii = 0; ii < this.neighbend[v].length; ii++) {
          var p = this.neighbend[v][ii];
          var k = ~~(p / 2);
          var w = this.endpoint[p];
          if (this.inBlossom[v] === this.inBlossom[w]) continue;
          if (!this.allowEdge[k]) {
            var kSlack = this.slack(k);
            if (kSlack <= 0) {
              this.allowEdge[k] = true;
            }
          }
          if (this.allowEdge[k]) {
            if (this.label[this.inBlossom[w]] === 0) {
              this.assignLabel(w, 2, p ^ 1);
            } else if (this.label[this.inBlossom[w]] === 1) {
              var base = this.scanBlossom(v, w);
              if (base >= 0) {
                this.addBlossom(base, k);
              } else {
                this.augmentMatching(k);
                augmented = true;
                break;
              }
            } else if (this.label[w] === 0) {
              //console.assert(this.label[this.inBlossom[w]] === 2);
              this.label[w] = 2;
              this.labelEnd[w] = p ^ 1;
            }
          } else if (this.label[this.inBlossom[w]] === 1) {
            var b = this.inBlossom[v];
            if (this.bestEdge[b] === -1 || kSlack < this.slack(this.bestEdge[b])) {
              this.bestEdge[b] = k;
            }
          } else if (this.label[w] === 0) {
            if (this.bestEdge[w] === -1 || kSlack < this.slack(this.bestEdge[w])) {
              this.bestEdge[w] = k;
            }
          }
        }
      }
      if (augmented) break;
      var deltaType = -1;
      var delta = [];
      var deltaEdge = [];
      var deltaBlossom = [];
      if (!this.maxCardinality) {
        deltaType = 1;
        delta = getMin(this.dualVar, 0, this.nVertex - 1);
      }
      for (v = 0; v < this.nVertex; v++) {
        if (this.label[this.inBlossom[v]] === 0 && this.bestEdge[v] !== -1) {
          var d = this.slack(this.bestEdge[v]);
          if (deltaType === -1 || d < delta) {
            delta = d;
            deltaType = 2;
            deltaEdge = this.bestEdge[v];
          }
        }
      }
      for (b = 0; b < 2 * this.nVertex; b++) {
        if (this.blossomParent[b] === -1 && this.label[b] === 1 && this.bestEdge[b] !== -1) {
          kSlack = this.slack(this.bestEdge[b]);
          ////console.assert((kSlack % 2) == 0);
          d = kSlack / 2;
          if (deltaType === -1 || d < delta) {
            delta = d;
            deltaType = 3;
            deltaEdge = this.bestEdge[b];
          }
        }
      }
      for (b = this.nVertex; b < this.nVertex * 2; b++) {
        if (this.blossomBase[b] >= 0 && this.blossomParent[b] === -1 && this.label[b] === 2 && (deltaType === -1 || this.dualVar[b] < delta)) {
          delta = this.dualVar[b];
          deltaType = 4;
          deltaBlossom = b;
        }
      }
      if (deltaType === -1) {
        //console.assert(this.maxCardinality);
        deltaType = 1;
        delta = Math.max(0, getMin(this.dualVar, 0, this.nVertex - 1));
      }
      for (v = 0; v < this.nVertex; v++) {
        var curLabel = this.label[this.inBlossom[v]];
        if (curLabel === 1) {
          this.dualVar[v] -= delta;
        } else if (curLabel === 2) {
          this.dualVar[v] += delta;
        }
      }
      for (b = this.nVertex; b < this.nVertex * 2; b++) {
        if (this.blossomBase[b] >= 0 && this.blossomParent[b] === -1) {
          if (this.label[b] === 1) {
            this.dualVar[b] += delta;
          } else if (this.label[b] === 2) {
            this.dualVar[b] -= delta;
          }
        }
      }
      //console.log('DEBUG: deltaType', deltaType, ' delta: ', delta);
      if (deltaType === 1) {
        break;
      } else if (deltaType === 2) {
        this.allowEdge[deltaEdge] = true;
        var i = this.edges[deltaEdge][0];
        var j = this.edges[deltaEdge][1];
        var wt = this.edges[deltaEdge][2];
        if (this.label[this.inBlossom[i]] === 0) {
          i = i ^ j;
          j = j ^ i;
          i = i ^ j;
        }
        //console.assert(this.label[this.inBlossom[i]] == 1);
        this.queue.push(i);
      } else if (deltaType === 3) {
        this.allowEdge[deltaEdge] = true;
        i = this.edges[deltaEdge][0];
        j = this.edges[deltaEdge][1];
        wt = this.edges[deltaEdge][2];
        //console.assert(this.label[this.inBlossom[i]] == 1);
        this.queue.push(i);
      } else if (deltaType === 4) {
        this.expandBlossom(deltaBlossom, false);
      }
    }
    if (!augmented) break;
    for (b = this.nVertex; b < this.nVertex * 2; b++) {
      if (this.blossomParent[b] === -1 && this.blossomBase[b] >= 0 && this.label[b] === 1 && this.dualVar[b] === 0) {
        this.expandBlossom(b, true);
      }
    }
  }
  for (v = 0; v < this.nVertex; v++) {
    if (this.mate[v] >= 0) {
      this.mate[v] = this.endpoint[this.mate[v]];
    }
  }
  for (v = 0; v < this.nVertex; v++) {
    //console.assert(this.mate[v] == -1 || this.mate[this.mate[v]] == v);
  }
  return this.mate;
};

Edmonds.prototype.slack = function (k) {
  var i = this.edges[k][0];
  var j = this.edges[k][1];
  var wt = this.edges[k][2];
  return this.dualVar[i] + this.dualVar[j] - 2 * wt;
};

Edmonds.prototype.blossomLeaves = function (b) {
  if (b < this.nVertex) {
    return [b];
  }
  var leaves = [];
  var childList = this.blossomChilds[b];
  for (var t = 0; t < childList.length; t++) {
    if (childList[t] <= this.nVertex) {
      leaves.push(childList[t]);
    } else {
      var leafList = this.blossomLeaves(childList[t]);
      for (var v = 0; v < leafList.length; v++) {
        leaves.push(leafList[v]);
      }
    }
  }
  return leaves;
};

Edmonds.prototype.assignLabel = function (w, t, p) {
  //console.log('DEBUG: assignLabel(' + w + ',' + t + ',' + p + '}');
  var b = this.inBlossom[w];
  //console.assert(this.label[w] === 0 && this.label[b] === 0);
  this.label[w] = this.label[b] = t;
  this.labelEnd[w] = this.labelEnd[b] = p;
  this.bestEdge[w] = this.bestEdge[b] = -1;
  if (t === 1) {
    this.queue.push.apply(this.queue, this.blossomLeaves(b));
    //console.log('DEBUG: PUSH ' + this.blossomLeaves(b).toString());
  } else if (t === 2) {
    var base = this.blossomBase[b];
    //console.assert(this.mate[base] >= 0);
    this.assignLabel(this.endpoint[this.mate[base]], 1, this.mate[base] ^ 1);
  }
};

Edmonds.prototype.scanBlossom = function (v, w) {
  //console.log('DEBUG: scanBlossom(' + v + ',' + w + ')');
  var path = [];
  var base = -1;
  while (v !== -1 || w !== -1) {
    var b = this.inBlossom[v];
    if ((this.label[b] & 4)) {
      base = this.blossomBase[b];
      break;
    }
    //console.assert(this.label[b] === 1);
    path.push(b);
    this.label[b] = 5;
    //console.assert(this.labelEnd[b] === this.mate[this.blossomBase[b]]);
    if (this.labelEnd[b] === -1) {
      v = -1;
    } else {
      v = this.endpoint[this.labelEnd[b]];
      b = this.inBlossom[v];
      //console.assert(this.label[b] === 2);
      //console.assert(this.labelEnd[b] >= 0);
      v = this.endpoint[this.labelEnd[b]];
    }
    if (w !== -1) {
      v = v ^ w;
      w = w ^ v;
      v = v ^ w;
    }
  }
  for (var ii = 0; ii < path.length; ii++) {
    b = path[ii];
    this.label[b] = 1;
  }
  return base;
};

Edmonds.prototype.addBlossom = function (base, k) {
  var v = this.edges[k][0];
  var w = this.edges[k][1];
  var wt = this.edges[k][2];
  var bb = this.inBlossom[base];
  var bv = this.inBlossom[v];
  var bw = this.inBlossom[w];
  var b = this.unusedBlossoms.pop();
  //console.log('DEBUG: addBlossom(' + base + ',' + k + ')' + ' (v=' + v + ' w=' + w + ')' + ' -> ' + b);
  this.blossomBase[b] = base;
  this.blossomParent[b] = -1;
  this.blossomParent[bb] = b;
  var path = this.blossomChilds[b] = [];
  var endPs = this.blossomEndPs[b] = [];
  while (bv !== bb) {
    this.blossomParent[bv] = b;
    path.push(bv);
    endPs.push(this.labelEnd[bv]);
    //console.assert(this.label[bv] === 2 || (this.label[bv] === 1 && this.labelEnd[bv] === this.mate[this.blossomBase[bv]]));
    //console.assert(this.labelEnd[bv] >= 0);
    v = this.endpoint[this.labelEnd[bv]];
    bv = this.inBlossom[v];
  }
  path.push(bb);
  path.reverse();
  endPs.reverse();
  endPs.push((2 * k));
  while (bw !== bb) {
    this.blossomParent[bw] = b;
    path.push(bw);
    endPs.push(this.labelEnd[bw] ^ 1);
    //console.assert(this.label[bw] === 2 || (this.label[bw] === 1 && this.labelEnd[bw] === this.mate[this.blossomBase[bw]]));
    //console.assert(this.labelEnd[bw] >= 0);
    w = this.endpoint[this.labelEnd[bw]];
    bw = this.inBlossom[w];
  }
  //console.assert(this.label[bb] === 1);
  this.label[b] = 1;
  this.labelEnd[b] = this.labelEnd[bb];
  this.dualVar[b] = 0;
  var leaves = this.blossomLeaves(b);
  for (var ii = 0; ii < leaves.length; ii++) {
    v = leaves[ii];
    if (this.label[this.inBlossom[v]] === 2) {
      this.queue.push(v);
    }
    this.inBlossom[v] = b;
  }
  var bestEdgeTo = filledArray(2 * this.nVertex, -1);
  for (ii = 0; ii < path.length; ii++) {
    bv = path[ii];
    if (this.blossomBestEdges[bv].length === 0) {
      var nbLists = [];
      leaves = this.blossomLeaves(bv);
      for (var x = 0; x < leaves.length; x++) {
        v = leaves[x];
        nbLists[x] = [];
        for (var y = 0; y < this.neighbend[v].length; y++) {
          var p = this.neighbend[v][y];
          nbLists[x].push(~~(p / 2));
        }
      }
    } else {
      nbLists = [this.blossomBestEdges[bv]];
    }
    //console.log('DEBUG: nbLists ' + nbLists.toString());
    for (x = 0; x < nbLists.length; x++) {
      var nbList = nbLists[x];
      for (y = 0; y < nbList.length; y++) {
        k = nbList[y];
        var i = this.edges[k][0];
        var j = this.edges[k][1];
        wt = this.edges[k][2];
        if (this.inBlossom[j] === b) {
          i = i ^ j;
          j = j ^ i;
          i = i ^ j;
        }
        var bj = this.inBlossom[j];
        if (bj !== b && this.label[bj] === 1 && (bestEdgeTo[bj] === -1 || this.slack(k) < this.slack(bestEdgeTo[bj]))) {
          bestEdgeTo[bj] = k;
        }
      }
    }
    this.blossomBestEdges[bv] = [];
    this.bestEdge[bv] = -1;
  }
  var be = [];
  for (ii = 0; ii < bestEdgeTo.length; ii++) {
    k = bestEdgeTo[ii];
    if (k !== -1) {
      be.push(k);
    }
  }
  this.blossomBestEdges[b] = be;
  //console.log('DEBUG: blossomBestEdges[' + b + ']= ' + this.blossomBestEdges[b].toString());
  this.bestEdge[b] = -1;
  for (ii = 0; ii < this.blossomBestEdges[b].length; ii++) {
    k = this.blossomBestEdges[b][ii];
    if (this.bestEdge[b] === -1 || this.slack(k) < this.slack(this.bestEdge[b])) {
      this.bestEdge[b] = k;
    }
  }
  //console.log('DEBUG: blossomChilds[' + b + ']= ' + this.blossomChilds[b].toString());
};

Edmonds.prototype.expandBlossom = function (b, endStage) {
  //console.log('DEBUG: expandBlossom(' + b + ',' + endStage + ') ' + this.blossomChilds[b].toString());
  for (var ii = 0; ii < this.blossomChilds[b].length; ii++) {
    var s = this.blossomChilds[b][ii];
    this.blossomParent[s] = -1;
    if (s < this.nVertex) {
      this.inBlossom[s] = s;
    } else if (endStage && this.dualVar[s] === 0) {
      this.expandBlossom(s, endStage);
    } else {
      var leaves = this.blossomLeaves(s);
      for (var jj = 0; jj < leaves.length; jj++) {
        var v = leaves[jj];
        this.inBlossom[v] = s;
      }
    }
  }
  if (!endStage && this.label[b] === 2) {
    //console.assert(this.labelEnd[b] >= 0);
    var entryChild = this.inBlossom[this.endpoint[this.labelEnd[b] ^ 1]];
    var j = this.blossomChilds[b].indexOf(entryChild);
    if ((j & 1)) {
      j -= this.blossomChilds[b].length;
      var jStep = 1;
      var endpTrick = 0;
    } else {
      jStep = -1;
      endpTrick = 1;
    }
    var p = this.labelEnd[b];
    while (j !== 0) {
      this.label[this.endpoint[p ^ 1]] = 0;
      this.label[this.endpoint[pIndex(this.blossomEndPs[b], j - endpTrick) ^ endpTrick ^ 1]] = 0;
      this.assignLabel(this.endpoint[p ^ 1], 2, p);
      this.allowEdge[~~(pIndex(this.blossomEndPs[b], j - endpTrick) / 2)] = true;
      j += jStep;
      p = pIndex(this.blossomEndPs[b], j - endpTrick) ^ endpTrick;
      this.allowEdge[~~(p / 2)] = true;
      j += jStep;
    }
    var bv = pIndex(this.blossomChilds[b], j);
    this.label[this.endpoint[p ^ 1]] = this.label[bv] = 2;

    this.labelEnd[this.endpoint[p ^ 1]] = this.labelEnd[bv] = p;
    this.bestEdge[bv] = -1;
    j += jStep;
    while (pIndex(this.blossomChilds[b], j) !== entryChild) {
      bv = pIndex(this.blossomChilds[b], j);
      if (this.label[bv] === 1) {
        j += jStep;
        continue;
      }
      leaves = this.blossomLeaves(bv);
      for (ii = 0; ii < leaves.length; ii++) {
        v = leaves[ii];
        if (this.label[v] !== 0) break;
      }
      if (this.label[v] !== 0) {
        //console.assert(this.label[v] === 2);
        //console.assert(this.inBlossom[v] === bv);
        this.label[v] = 0;
        this.label[this.endpoint[this.mate[this.blossomBase[bv]]]] = 0;
        this.assignLabel(v, 2, this.labelEnd[v]);
      }
      j += jStep;
    }
  }
  this.label[b] = this.labelEnd[b] = -1;
  this.blossomEndPs[b] = this.blossomChilds[b] = [];
  this.blossomBase[b] = -1;
  this.blossomBestEdges[b] = [];
  this.bestEdge[b] = -1;
  this.unusedBlossoms.push(b);
};

Edmonds.prototype.augmentBlossom = function (b, v) {
  //console.log('DEBUG: augmentBlossom(' + b + ',' + v + ')');
  var i, j;
  var t = v;
  while (this.blossomParent[t] !== b) {
    t = this.blossomParent[t];
  }
  if (t > this.nVertex) {
    this.augmentBlossom(t, v);
  }
  i = j = this.blossomChilds[b].indexOf(t);
  if ((i & 1)) {
    j -= this.blossomChilds[b].length;
    var jStep = 1;
    var endpTrick = 0;
  } else {
    jStep = -1;
    endpTrick = 1;
  }
  while (j !== 0) {
    j += jStep;
    t = pIndex(this.blossomChilds[b], j);
    var p = pIndex(this.blossomEndPs[b], j - endpTrick) ^ endpTrick;
    if (t >= this.nVertex) {
      this.augmentBlossom(t, this.endpoint[p]);
    }
    j += jStep;
    t = pIndex(this.blossomChilds[b], j);
    if (t >= this.nVertex) {
      this.augmentBlossom(t, this.endpoint[p ^ 1]);
    }
    this.mate[this.endpoint[p]] = p ^ 1;
    this.mate[this.endpoint[p ^ 1]] = p;
  }
  //console.log('DEBUG: PAIR ' + this.endpoint[p] + ' ' + this.endpoint[p^1] + '(k=' + ~~(p/2) + ')');
  this.blossomChilds[b] = this.blossomChilds[b].slice(i).concat(this.blossomChilds[b].slice(0, i));
  this.blossomEndPs[b] = this.blossomEndPs[b].slice(i).concat(this.blossomEndPs[b].slice(0, i));
  this.blossomBase[b] = this.blossomBase[this.blossomChilds[b][0]];
  //console.assert(this.blossomBase[b] === v);
};

Edmonds.prototype.augmentMatching = function (k) {
  var v = this.edges[k][0];
  var w = this.edges[k][1];
  //console.log('DEBUG: augmentMatching(' + k + ')' + ' (v=' + v + ' ' + 'w=' + w);
  //console.log('DEBUG: PAIR ' + v + ' ' + w + '(k=' + k + ')');
  for (var ii = 0; ii < 2; ii++) {
    if (ii === 0) {
      var s = v;
      var p = 2 * k + 1;
    } else {
      s = w;
      p = 2 * k;
    }
    while (true) {
      var bs = this.inBlossom[s];
      //console.assert(this.label[bs] === 1);
      //console.assert(this.labelEnd[bs] === this.mate[this.blossomBase[bs]]);
      if (bs >= this.nVertex) {
        this.augmentBlossom(bs, s);
      }
      this.mate[s] = p;
      if (this.labelEnd[bs] === -1) break;
      var t = this.endpoint[this.labelEnd[bs]];
      var bt = this.inBlossom[t];
      //console.assert(this.label[bt] === 2);
      //console.assert(this.labelEnd[bt] >= 0);
      s = this.endpoint[this.labelEnd[bt]];
      var j = this.endpoint[this.labelEnd[bt] ^ 1];
      //console.assert(this.blossomBase[bt] === t);
      if (bt >= this.nVertex) {
        this.augmentBlossom(bt, j);
      }
      this.mate[j] = this.labelEnd[bt];
      p = this.labelEnd[bt] ^ 1;
      //console.log('DEBUG: PAIR ' + s + ' ' + t + '(k=' + ~~(p/2) + ')');


    }
  }
};


//INIT STUFF//
Edmonds.prototype.init = function () {
  this.nVertexInit();
  this.maxWeightInit();
  this.endpointInit();
  this.neighbendInit();
  this.mate = filledArray(this.nVertex, -1);
  this.label = filledArray(2 * this.nVertex, 0); //remove?
  this.labelEnd = filledArray(2 * this.nVertex, -1);
  this.inBlossomInit();
  this.blossomParent = filledArray(2 * this.nVertex, -1);
  this.blossomChilds = initArrArr(2 * this.nVertex);
  this.blossomBaseInit();
  this.blossomEndPs = initArrArr(2 * this.nVertex);
  this.bestEdge = filledArray(2 * this.nVertex, -1); //remove?
  this.blossomBestEdges = initArrArr(2 * this.nVertex); //remove?
  this.unusedBlossomsInit();
  this.dualVarInit();
  this.allowEdge = filledArray(this.nEdge, false); //remove?
  this.queue = []; //remove?
};
Edmonds.prototype.blossomBaseInit = function () {
  var base = [];
  for (var i = 0; i < this.nVertex; i++) {
    base[i] = i;
  }
  var negs = filledArray(this.nVertex, -1);
  this.blossomBase = base.concat(negs);
};
Edmonds.prototype.dualVarInit = function () {
  var mw = filledArray(this.nVertex, this.maxWeight);
  var zeros = filledArray(this.nVertex, 0);
  this.dualVar = mw.concat(zeros);
};
Edmonds.prototype.unusedBlossomsInit = function () {
  var i, unusedBlossoms = [];
  for (i = this.nVertex; i < 2 * this.nVertex; i++) {
    unusedBlossoms.push(i);
  }
  this.unusedBlossoms = unusedBlossoms;
};
Edmonds.prototype.inBlossomInit = function () {
  var i, inBlossom = [];
  for (i = 0; i < this.nVertex; i++) {
    inBlossom[i] = i;
  }
  this.inBlossom = inBlossom;
};
Edmonds.prototype.neighbendInit = function () {
  var k, i, j;
  var neighbend = initArrArr(this.nVertex);
  for (k = 0; k < this.nEdge; k++) {
    i = this.edges[k][0];
    j = this.edges[k][1];
    neighbend[i].push(2 * k + 1);
    neighbend[j].push(2 * k);
  }
  this.neighbend = neighbend;
};
Edmonds.prototype.endpointInit = function () {
  var p;
  var endpoint = [];
  for (p = 0; p < 2 * this.nEdge; p++) {
    endpoint[p] = this.edges[~~(p / 2)][p % 2];
  }
  this.endpoint = endpoint;
};
Edmonds.prototype.nVertexInit = function () {
  var nVertex = 0;
  for (var k = 0; k < this.nEdge; k++) {
    var i = this.edges[k][0];
    var j = this.edges[k][1];
    if (i >= nVertex) nVertex = i + 1;
    if (j >= nVertex) nVertex = j + 1;
  }
  this.nVertex = nVertex;
};
Edmonds.prototype.maxWeightInit = function () {
  var maxWeight = 0;
  for (var k = 0; k < this.nEdge; k++) {
    var weight = this.edges[k][2];
    if (weight > maxWeight) {
      maxWeight = weight;
    }
  }
  this.maxWeight = maxWeight;
};

//HELPERS//
function filledArray(len, fill) {
  var i, newArray = [];
  for (i = 0; i < len; i++) {
    newArray[i] = fill;
  }
  return newArray;
}

function initArrArr(len) {
  var arr = [];
  for (var i = 0; i < len; i++) {
    arr[i] = [];
  }
  return arr;
}

function getMin(arr, start, end) {
  var min = Infinity;
  for (var i = start; i <= end; i++) {
    if (arr[i] < min) {
      min = arr[i];
    }
  }
  return min;
}

function pIndex(arr, idx) {
  //if idx is negative, go from the back
  return idx < 0 ? arr[arr.length + idx] : arr[idx];
}

// -----------------------------------------------------
// Swiss pairing.

function calculateScoreGroups(standings) {
  var groups = []
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
  var top = groups[j]
  if (top === undefined) {
    console.log("undef!")
    console.log(groups)
    console.log(j)
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

function pairCandidates(bracket) {
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
  console.log("inames", inames)

  for (var player of bracket) {
    for (var m of player) {
      const [repeats, distance, p1, p2] = m;
      // Don't pair candidates too far apart
      if (distance < 11) {
        let weight = -(30 * repeats + distance);
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
  console.log("blossom:", b)
  var pairings = []
  for (var i = 0; i < b.length; i++) {
    let v = b[i];
    let p1 = inames[i];
    let p2 = inames[v];
    pairings.push({ first: { name: p1 }, second: { name: p2 } })
  }
  console.log("sub pairing:", pairings)
  return pairings
}

function pairSwiss(results, entrants, repeats, round, for_round, ) {
  console.log("swiss pairing based on round", round)
  if (round <= 0) {
    return pairSwissInitial(entrants.seeding);
  }
  //console.log("repeats for round", round, repeats.matches)
  var players = standingsAfterRound(results, entrants, round);
  var fixed;
  [players, fixed] = removeFixedPairings(players, entrants, for_round);
  var groups = calculateScoreGroups(players);
  var dgroups = groups.map(g => g.map(p => [p.name, p.wins]));
  console.log("groups:", dgroups)
  var candidates;
  var nrep = 1;
  var paired = [];
  // Don't have too small a bottom group
  if (groups.length > 1) {
    while (groups[groups.length - 1].length < 6) {
      mergeBottom(groups);
      groups = groups.filter(e => e.length != 0);
    }
  }
  while (groups.length > 0) {
    dgroups = groups.map(g => g.map(p => [p.name, p.wins]));
    console.log("groups:", dgroups)
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
      var pairs = pairCandidates(candidates)
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
  console.log("fixed:", fixed)
  paired.push(fixed);
  var out = []
  for (const group of paired) {
    for (var p of group) {
      if (p.first.name < p.second.name) {
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
      rating
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

function outputPairings(pairing_sheet, text_pairing_sheet, pairings, entrants, round, start_row) {
  console.log("pairings:", pairings);
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
    var first = x.first.start ? x.first.name : x.second.name;
    var second = x.first.start ? x.second.name : x.first.name;
    var rep = x.repeats > 1 ? `(rep ${x.repeats})` : "";
    return [
      table,
      entrants.entrants[first] || first,
      entrants.entrants[second] || second,
      rep,
    ]
  })
  out = out.sort((a, b) => parseInt(a) - parseInt(b));
  var ncols = out[0].length
  var text_pairings = []
  var round_header = "ROUND " + (round + 1);
  var pairing_strings = pairings.map(x => `${x.first.name} v. ${x.second.name}`);
  var pairing_string = round_header + ": " + pairing_strings.join(" | ");
  text_pairings.push([pairing_string])
  var header = [
    [round_header, "", "", ""],
  ];
  out = header.concat(out);
  // Write out standings starting in cell A2
  var outputRow = start_row;
  var outputCol = 1;
  var outputRange = pairing_sheet.getRange(outputRow, outputCol, out.length, ncols);
  outputRange.setValues(out);
  outputRange.setFontWeight("normal");
  var headerRange = pairing_sheet.getRange(outputRow, outputCol, header.length, ncols);
  headerRange.setFontWeight("bold");
  var textPairingRange = text_pairing_sheet.getRange(round + 1, 1, 1, 1);
  textPairingRange.setValues(text_pairings);
}

function processSheet(input_sheet_label, standings_sheet_label, pairing_sheet_label, text_pairing_sheet_label) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var result_sheet = sheet.getSheetByName(input_sheet_label);

  var res = collectResults(result_sheet);

  var entrants = collectEntrants();
  var ratings = {};
  for (var x of entrants.seeding) {
    ratings[x.name] = x.rating
  }
  var fp = collectFixedPairings();
  entrants.fixed_pairings = fp;

  console.log("processed results");


  // Write out the standings
  var standings_sheet = sheet.getSheetByName(standings_sheet_label);
  outputPlayerStandings(standings_sheet, res.players, entrants, ratings);

  // Write out the pairings
  var round_pairings = collectRoundPairings();
  console.log("round pairings:", round_pairings);

  // Find the last round we can pair
  var round_ids = res.roundIds();
  console.log("round ids:", round_ids);
  var last_result = Math.max(...round_ids);
  var last_round = 0;
  for (var r of Object.values(round_pairings)) {
    if (r.start - 1 <= last_result) {
      console.log("Pairing round", r, "based on round", r.start);
      last_round = r.round;
    }
  }
  console.log("Last round", last_round);

  var pairing_sheet = sheet.getSheetByName(pairing_sheet_label);
  var text_pairing_sheet = sheet.getSheetByName(text_pairing_sheet_label);

  // Clear pairing sheets
  pairing_sheet.clearContents();
  text_pairing_sheet.clearContents();

  var row = 2;
  var rr_starts = {}
  var repeats = new Repeats();
  for (var i = 0; i < last_round; i++) {
    var rp = round_pairings[i + 1];
    console.log("writing pairings for round:", i + 1, rp);
    var pairings;
    if (i + 1 < round_ids.length) {
      pairings = res.extractPairings(i + 1)
    } else {
      pairings = pairingsAfterRound(res, entrants, repeats, round_pairings, i);
      for (var p of pairings) {
        var p1 = p.first.name;
        var p2 = p.second.name;
        var p1_first;
        if (rp.type == "R" || rp.type == "CH") {
          if (rr_starts[p1] === undefined) {
            rr_starts[p1] = 0
          }
          if (rr_starts[p2] === undefined) {
            rr_starts[p2] = 0
          }
          // Always assign 'bye' the first otherwise the player playing the bye
          // is assigned an extra first andthereby slightly penalised.
          if (p1.toLowerCase() === "bye") {
            p1_first = true;
          } else if (p2.toLowerCase() === "bye") {
            p1_first = false;
          } else {
            p1_first = rr_starts[p1] <= rr_starts[p2];
          }
          if (p1_first) {
            rr_starts[p1]++;
          } else {
            rr_starts[p2]++;
          }
        } else {
          if (p1.toLowerCase() === "bye") {
            p1_first = true;
          } else if (p2.toLowerCase() === "bye") {
            p1_first = false;
          } else {
            p1_first = res.players[p1].starts <= res.players[p2].starts;
          }
        }
        p.first.start = p1_first;
        p.second.start = !p1_first;
      }
    }
    for (var p of pairings) {
      p.repeats = repeats.add(p.first.name, p.second.name);
    }
    outputPairings(pairing_sheet, text_pairing_sheet, pairings, entrants, i, row);
    row += pairings.length + 2;
  }
}

function calculateStandings() {
  processSheet("Results", "Standings", "Pairings", "Text Pairings");
}

// export {
//   makeEntrants, makeRoundPairings, makeResults, pairingsAfterRound,
//   standingsAfterRound
// };
