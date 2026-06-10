/**
 * ─────────────────────────────────────────────────────────────
 *  Golf Ball Tracker — Apps Script Backend (Code.gs)
 * ─────────────────────────────────────────────────────────────
 *
 *  Sheet columns (updated layout):
 *    A: Timestamp  B: Brand  C: Model  D: Colour
 *    E: Status     F: Time Lost
 *
 *  ⚙️  SETUP: Change SHEET_NAME if your tab name differs.
 */

const SHEET_NAME = "Individual_Ball_Stock";

// ─── Serve the web app ────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Ball Tracker")
    .setWidth(400)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Internal: get the sheet ──────────────────────────────────
function getSheet_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found. Update SHEET_NAME in Code.gs.`);
  return sheet;
}

// ─── getUniqueBalls ───────────────────────────────────────────
// Returns unique Brand + Model + Colour combos that have ≥1 in-stock row.
// Also returns a count so the app can show "3 left".
function getUniqueBalls() {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data   = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const counts = {};

  for (const row of data) {
    const brand  = String(row[1] || "").trim();
    const model  = String(row[2] || "").trim();
    const colour = String(row[3] || "White").trim();
    const status = String(row[4] || "").trim();

    if (!brand || !model)  continue;
    if (status === "Lost") continue;

    const key = `${brand.toLowerCase()}||${model.toLowerCase()}||${colour.toLowerCase()}`;
    if (counts[key]) {
      counts[key].count++;
    } else {
      counts[key] = { brand, model, colour, count: 1 };
    }
  }

  return Object.values(counts).sort((a, b) =>
    `${a.brand} ${a.model}`.localeCompare(`${b.brand} ${b.model}`)
  );
}

// ─── markAsLost ───────────────────────────────────────────────
// Marks the first matching in-stock row as Lost.
function markAsLost(brand, model, colour) {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false };

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  for (let i = 0; i < data.length; i++) {
    const rBrand  = String(data[i][1] || "").trim();
    const rModel  = String(data[i][2] || "").trim();
    const rColour = String(data[i][3] || "White").trim();
    const rStatus = String(data[i][4] || "").trim();

    if (
      rBrand.toLowerCase()  === brand.toLowerCase()  &&
      rModel.toLowerCase()  === model.toLowerCase()  &&
      rColour.toLowerCase() === colour.toLowerCase() &&
      rStatus !== "Lost"
    ) {
      const sheetRow = i + 2;
      sheet.getRange(sheetRow, 5).setValue("Lost");     // col E
      sheet.getRange(sheetRow, 6).setValue(new Date()); // col F
      return { success: true };
    }
  }

  return { success: false, message: "No matching in-stock ball found" };
}

// ─── getRecentLost ────────────────────────────────────────────
// Returns the N most recently lost balls (default 5).
function getRecentLost(n) {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const lost = [];

  for (let i = 0; i < data.length; i++) {
    const brand    = String(data[i][1] || "").trim();
    const model    = String(data[i][2] || "").trim();
    const colour   = String(data[i][3] || "White").trim();
    const status   = String(data[i][4] || "").trim();
    const timeLost = data[i][5];

    if (status === "Lost" && timeLost) {
      lost.push({ rowIndex: i + 2, brand, model, colour, ts: new Date(timeLost).getTime() });
    }
  }

  lost.sort((a, b) => b.ts - a.ts);

  const tz = Session.getScriptTimeZone();
  return lost.slice(0, n || 5).map(b => ({
    rowIndex:  b.rowIndex,
    brand:     b.brand,
    model:     b.model,
    colour:    b.colour,
    timeLabel: Utilities.formatDate(new Date(b.ts), tz, "HH:mm, d MMM"),
  }));
}

// ─── restoreBall ─────────────────────────────────────────────
// Clears Status and Time Lost for a given row.
function restoreBall(rowIndex) {
  const sheet = getSheet_();
  sheet.getRange(rowIndex, 5).setValue(""); // col E = Status
  sheet.getRange(rowIndex, 6).setValue(""); // col F = Time Lost
  return { success: true };
}

// ─── addBall ──────────────────────────────────────────────────
// Appends a new in-stock ball row to the sheet.
function addBall(brand, model, colour) {
  const sheet = getSheet_();
  sheet.appendRow([
    new Date(),          // col A: Timestamp
    brand,               // col B: Brand
    model,               // col C: Model
    colour || "White",   // col D: Colour
    "",                  // col E: Status (blank = in stock)
    "",                  // col F: Time Lost
  ]);
  return { success: true };
}

// ─── getStats ────────────────────────────────────────────────
// Returns balls lost today + the most recently lost ball.
function getStats() {
  const sheet   = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lostToday: 0, lastLost: null };

  const data    = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const tz      = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "d MMM yyyy");

  let lostToday = 0;
  let lastLost  = null;

  for (const row of data) {
    const brand    = String(row[1] || "").trim();
    const model    = String(row[2] || "").trim();
    const colour   = String(row[3] || "White").trim();
    const status   = String(row[4] || "").trim();
    const timeLost = row[5];

    if (status !== "Lost" || !timeLost) continue;

    const lostDate = new Date(timeLost);
    if (Utilities.formatDate(lostDate, tz, "d MMM yyyy") === todayStr) lostToday++;

    if (!lastLost || lostDate.getTime() > lastLost.ts) {
      lastLost = {
        brand, model, colour,
        ts:        lostDate.getTime(),
        timeLabel: Utilities.formatDate(lostDate, tz, "HH:mm"),
      };
    }
  }

  return { lostToday, lastLost };
}
