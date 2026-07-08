/**
 * ─────────────────────────────────────────────────────────────
 *  Golf Ball Tracker — Apps Script Backend (Code.gs)
 * ─────────────────────────────────────────────────────────────
 *
 *  Ball Stock sheet ("Individual_Ball_Stock") columns:
 *    A: Timestamp  B: Brand  C: Model  D: Colour
 *    E: Status     F: Time Lost  G: Current Bag  H: Course  I: Hole
 *
 *  Scores sheet ("Scores") columns — one row per hole played:
 *    A: RoundID  B: Date  C: Course  D: Format(9/18)  E: Hole
 *    F: Score  G: Putts  H: Penalties  I: FairwayHit(Y/N/NA)  J: GIR(Y/N)
 *
 *  Courses sheet ("Courses") columns — one row per hole per course:
 *    A: Course  B: Hole  C: Par
 *
 *  This is a REFERENCE COPY kept in the repo for version control.
 *  The live source of truth is the Apps Script project attached to
 *  the spreadsheet — paste this file's contents there, then
 *  Deploy > Manage deployments > (pencil) > New version > Deploy
 *  to update the existing web app WITHOUT changing its URL.
 */

const SHEET_NAME    = "Individual_Ball_Stock";
const SCORES_SHEET  = "Scores";
const COURSES_SHEET = "Courses";

// ─── CORS helper ─────────────────────────────────────────────
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── doGet ───────────────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  try {
    if (action === "getUniqueBalls")  return corsOutput(getUniqueBalls());
    if (action === "getRecentLost")   return corsOutput(getRecentLost(e.parameter.n || 10));
    if (action === "getStats")        return corsOutput(getStats());
    if (action === "getBagNames")     return corsOutput(getBagNames());
    if (action === "getCourseNames")  return corsOutput(getCourseNames());
    if (action === "getHoleStats")    return corsOutput(getHoleStats(e.parameter.course));
    if (action === "rounds")          return corsOutput(getRounds());
    if (action === "courseParData")   return corsOutput(getCourseParData());
    return corsOutput({ ok: true, message: "Ball Tracker API" });
  } catch(err) {
    return corsOutput({ error: err.message });
  }
}

// ─── doPost ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === "markAsLost")     return corsOutput(markAsLost(body.brand, body.model, body.colour, body.course, body.hole));
    if (action === "restoreBall")    return corsOutput(restoreBall(body.rowIndex));
    if (action === "addBall")        return corsOutput(addBall(body.brand, body.model, body.colour, body.bag));
    if (action === "rounds")         return corsOutput(saveRound(body));
    if (action === "courseParData")  return corsOutput(saveCoursePar(body));
    return corsOutput({ error: "Unknown action" });
  } catch(err) {
    return corsOutput({ error: err.message });
  }
}

// ─── Internal: get a sheet by name ────────────────────────────
function getSheet_() {
  return getNamedSheet_(SHEET_NAME);
}
function getNamedSheet_(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found.`);
  return sheet;
}

// ─── getUniqueBalls ──────────────────────────────────────────
function getUniqueBalls() {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data   = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const counts = {};

  for (const row of data) {
    const brand  = String(row[1] || "").trim();
    const model  = String(row[2] || "").trim();
    const colour = String(row[3] || "White").trim();
    const status = String(row[4] || "").trim();
    const bag    = String(row[6] || "").trim();
    if (!brand || !model || status === "Lost") continue;
    const key = `${brand.toLowerCase()}||${model.toLowerCase()}||${colour.toLowerCase()}||${bag.toLowerCase()}`;
    if (counts[key]) {
      counts[key].count++;
    } else {
      counts[key] = { brand, model, colour, bag, count: 1 };
    }
  }

  return Object.values(counts).sort((a, b) =>
    `${a.brand} ${a.model}`.localeCompare(`${b.brand} ${b.model}`)
  );
}

// ─── getBagNames ─────────────────────────────────────────────
function getBagNames() {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const seen = new Set();

  for (const row of data) {
    const status = String(row[4] || "").trim();
    const bag    = String(row[6] || "").trim();
    if (bag && status !== "Lost") seen.add(bag);
  }

  const order = ["Pro V Bag","Other Titleist Bag","Callaway Bag","Good Bag","Yellow Bag","The Rest"];
  const ordered = order.filter(b => seen.has(b));
  seen.forEach(b => { if (!order.includes(b)) ordered.push(b); });
  return ordered;
}

// ─── getCourseNames ──────────────────────────────────────────
function getCourseNames() {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ["Drax Golf Club"];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const seen = new Set();

  for (const row of data) {
    const status = String(row[4] || "").trim();
    const course = String(row[7] || "").trim();
    if (course && status === "Lost") seen.add(course);
  }

  // Always include Drax as first option
  const result = ["Drax Golf Club"];
  seen.forEach(c => { if (c !== "Drax Golf Club") result.push(c); });
  return result;
}

// ─── markAsLost ──────────────────────────────────────────────
function markAsLost(brand, model, colour, course, hole) {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false };

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  for (let i = 0; i < data.length; i++) {
    const rBrand  = String(data[i][1] || "").trim();
    const rModel  = String(data[i][2] || "").trim();
    const rColour = String(data[i][3] || "White").trim();
    const rStatus = String(data[i][4] || "").trim();

    if (
      rBrand.toLowerCase()  === brand.toLowerCase()  &&
      rModel.toLowerCase()  === model.toLowerCase()  &&
      rColour.toLowerCase() === (colour || "white").toLowerCase() &&
      rStatus !== "Lost"
    ) {
      const sheetRow = i + 2;
      sheet.getRange(sheetRow, 5).setValue("Lost");
      sheet.getRange(sheetRow, 6).setValue(new Date());
      if (course) sheet.getRange(sheetRow, 8).setValue(course);
      if (hole)   sheet.getRange(sheetRow, 9).setValue(Number(hole));
      return { success: true };
    }
  }
  return { success: false, message: "No matching in-stock ball found" };
}

// ─── getRecentLost ────────────────────────────────────────────
function getRecentLost(n) {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const lost = [];

  for (let i = 0; i < data.length; i++) {
    const brand    = String(data[i][1] || "").trim();
    const model    = String(data[i][2] || "").trim();
    const colour   = String(data[i][3] || "White").trim();
    const status   = String(data[i][4] || "").trim();
    const timeLost = data[i][5];
    const bag      = String(data[i][6] || "").trim();
    const course   = String(data[i][7] || "").trim();
    const hole     = data[i][8] ? Number(data[i][8]) : null;
    if (status === "Lost" && timeLost) {
      lost.push({ rowIndex: i + 2, brand, model, colour, bag, course, hole, ts: new Date(timeLost).getTime() });
    }
  }

  lost.sort((a, b) => b.ts - a.ts);

  const tz = Session.getScriptTimeZone();
  return lost.slice(0, n || 10).map(b => ({
    rowIndex:  b.rowIndex,
    brand:     b.brand,
    model:     b.model,
    colour:    b.colour,
    bag:       b.bag,
    course:    b.course,
    hole:      b.hole,
    ts:        b.ts,
    timeLabel: Utilities.formatDate(new Date(b.ts), tz, "d MMM yyyy, HH:mm"),
  }));
}

// ─── restoreBall ─────────────────────────────────────────────
function restoreBall(rowIndex) {
  const sheet = getSheet_();
  sheet.getRange(rowIndex, 5).setValue("");
  sheet.getRange(rowIndex, 6).setValue("");
  sheet.getRange(rowIndex, 8).setValue("");
  sheet.getRange(rowIndex, 9).setValue("");
  return { success: true };
}

// ─── addBall ─────────────────────────────────────────────────
function addBall(brand, model, colour, bag) {
  const sheet = getSheet_();
  sheet.appendRow([new Date(), brand, model, colour || "White", "", "", bag || "", "", ""]);
  return { success: true };
}

// ─── getStats ────────────────────────────────────────────────
function getStats() {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lostToday: 0, totalLost: 0, lastLost: null, avgPerRound: null, worstDay: null, mostLost: null };

  const data     = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const tz       = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "d MMM yyyy");

  let lostToday = 0;
  let totalLost = 0;
  let lastLost  = null;
  const lostCounts = {};
  const dayTotals  = {};

  for (const row of data) {
    const brand    = String(row[1] || "").trim();
    const model    = String(row[2] || "").trim();
    const colour   = String(row[3] || "White").trim();
    const status   = String(row[4] || "").trim();
    const timeLost = row[5];
    if (!brand || !model || status !== "Lost" || !timeLost) continue;

    totalLost++;
    const lostDate = new Date(timeLost);
    const dayKey   = Utilities.formatDate(lostDate, tz, "d MMM yyyy");

    if (dayKey === todayStr) lostToday++;

    if (!lastLost || lostDate.getTime() > lastLost.ts) {
      lastLost = { brand, model, colour, ts: lostDate.getTime() };
    }

    const lkey = `${brand.toLowerCase()}||${model.toLowerCase()}`;
    lostCounts[lkey] = lostCounts[lkey] || { brand, model, count: 0 };
    lostCounts[lkey].count++;

    if (!dayTotals[dayKey]) dayTotals[dayKey] = { dateStr: dayKey, count: 0, ts: lostDate.getTime() };
    dayTotals[dayKey].count++;
    if (lostDate.getTime() > dayTotals[dayKey].ts) dayTotals[dayKey].ts = lostDate.getTime();
  }

  const daysWithLosses = Object.keys(dayTotals).length;
  const avgPerRound = daysWithLosses > 0
    ? Math.round((totalLost / daysWithLosses) * 10) / 10
    : null;

  const dayArr = Object.values(dayTotals).sort((a, b) => b.count - a.count || b.ts - a.ts);
  const worstDay = dayArr.length > 0 ? { dateStr: dayArr[0].dateStr, count: dayArr[0].count } : null;

  const lostArr = Object.values(lostCounts).sort((a, b) => b.count - a.count);
  let mostLost = null;
  if (lostArr.length > 0) {
    const top = lostArr[0];
    const tied = lostArr.filter(l => l.count === top.count);
    mostLost = tied.length === 1 ? top : { tied: true, count: top.count };
  }

  return { lostToday, totalLost, lastLost, avgPerRound, worstDay, mostLost };
}

// ─── getHoleStats ─────────────────────────────────────────────
// Returns per-hole loss data for a given course
function getHoleStats(course) {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { holes: [], roundDates: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const tz   = Session.getScriptTimeZone();

  // Track losses per hole and unique round dates (days with any loss on this course)
  const holeCounts = {};   // hole -> total losses
  const roundDates = new Set();

  for (const row of data) {
    const status   = String(row[4] || "").trim();
    const timeLost = row[5];
    const rowCourse = String(row[7] || "").trim();
    const hole     = row[8] ? Number(row[8]) : null;

    if (status !== "Lost" || !timeLost) continue;
    if (course && rowCourse.toLowerCase() !== course.toLowerCase()) continue;

    if (hole && hole >= 1 && hole <= 18) {
      holeCounts[hole] = (holeCounts[hole] || 0) + 1;
    }

    const dayKey = Utilities.formatDate(new Date(timeLost), tz, "d MMM yyyy");
    roundDates.add(dayKey);
  }

  const numRounds = roundDates.size;

  // Build array for holes 1–18
  const holes = [];
  for (let h = 1; h <= 18; h++) {
    const total = holeCounts[h] || 0;
    const avg   = numRounds > 0 ? Math.round((total / numRounds) * 100) / 100 : 0;
    holes.push({ hole: h, total, avg });
  }

  // Worst hole
  const sorted = [...holes].sort((a, b) => b.total - a.total || b.avg - a.avg);
  const worstHole = sorted[0] && sorted[0].total > 0 ? sorted[0] : null;

  return { holes, worstHole, numRounds };
}

// ─── getRounds ───────────────────────────────────────────────
// Returns every hole-row from the Scores sheet (flat list). The
// frontend groups rows sharing the same roundId back into rounds.
function getRounds() {
  const sheet   = getNamedSheet_(SCORES_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const tz   = Session.getScriptTimeZone();
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  return data
    .map(row => ({
      roundId:    String(row[0] || "").trim(),
      date:       row[1] instanceof Date ? Utilities.formatDate(row[1], tz, "yyyy-MM-dd") : String(row[1] || "").trim(),
      course:     String(row[2] || "").trim(),
      format:     Number(row[3]) || null,
      hole:       Number(row[4]) || null,
      score:      row[5] === "" || row[5] === null ? null : Number(row[5]),
      putts:      row[6] === "" || row[6] === null ? null : Number(row[6]),
      penalties:  row[7] === "" || row[7] === null ? null : Number(row[7]),
      fairwayHit: String(row[8] || "").trim(),
      gir:        String(row[9] || "").trim(),
    }))
    .filter(r => r.roundId && r.hole && r.score !== null);
}

// ─── saveRound ───────────────────────────────────────────────
// body: { course, format(9|18), date("yyyy-MM-dd"), holes: [{hole, score, putts?, penalties?, fairwayHit?, gir?}, ...] }
function saveRound(body) {
  const course = String(body.course || "").trim();
  const format = Number(body.format) || 18;
  const date   = String(body.date || "").trim();
  const holes  = Array.isArray(body.holes) ? body.holes : [];

  if (!course || !date || !holes.length) {
    return { success: false, message: "Missing course, date, or hole scores" };
  }

  const validHoles = holes.filter(h => h && Number(h.hole) && (h.score || h.score === 0));
  if (!validHoles.length) return { success: false, message: "No valid hole scores" };

  const sheet   = getNamedSheet_(SCORES_SHEET);
  const roundId = String(Date.now());

  const rows = validHoles.map(h => [
    roundId,
    date,
    course,
    format,
    Number(h.hole),
    Number(h.score),
    (h.putts === 0 || h.putts) ? Number(h.putts) : "",
    (h.penalties === 0 || h.penalties) ? Number(h.penalties) : "",
    h.fairwayHit || "",
    h.gir || "",
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  return { success: true, roundId };
}

// ─── getCourseParData ──────────────────────────────────────────
// Returns every {course, hole, par} row from the Courses sheet.
function getCourseParData() {
  const sheet   = getNamedSheet_(COURSES_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data
    .map(row => ({
      course: String(row[0] || "").trim(),
      hole:   Number(row[1]) || null,
      par:    row[2] === "" || row[2] === null ? null : Number(row[2]),
    }))
    .filter(r => r.course && r.hole && r.par);
}

// ─── saveCoursePar ─────────────────────────────────────────────
// body: { course, pars: [{hole, par}, ...] }
// Upserts rows into the Courses sheet — updates existing hole/course
// rows in place, appends new ones.
function saveCoursePar(body) {
  const course = String(body.course || "").trim();
  const pars   = Array.isArray(body.pars) ? body.pars : [];
  if (!course || !pars.length) return { success: false, message: "Missing course or par data" };

  const sheet    = getNamedSheet_(COURSES_SHEET);
  const lastRow  = sheet.getLastRow();
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];

  const toAppend = [];
  pars.forEach(p => {
    const hole = Number(p.hole);
    const par  = Number(p.par);
    if (!hole || !par) return;

    let found = false;
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).trim().toLowerCase() === course.toLowerCase() && Number(existing[i][1]) === hole) {
        sheet.getRange(i + 2, 3).setValue(par);
        found = true;
        break;
      }
    }
    if (!found) toAppend.push([course, hole, par]);
  });

  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 3).setValues(toAppend);
  }
  return { success: true };
}
