/**
 * Planted SF — Events & CRM Dashboard Builder
 * Artist's Way Spring 2026 | April 7 – June 30
 *
 * HOW TO RUN:
 * 1. Open any Google Sheet (can be a blank one)
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click Run ▶ next to "createPlantedSFDashboard"
 * 5. Click "Review permissions" and authorize when prompted
 * 6. A new spreadsheet will be created — a popup will show the link
 */

// ─── PROGRAM CONFIG ────────────────────────────────────────────────────────────

const SOURCE_SHEET_ID = "182Sa46ev_rFynKE3PhgBz0RnSkncexYuo2XZuj58q8A";
const PROGRAM_NAME    = "Artist's Way — Spring 2026";
const PROGRAM_DATES   = "April 7 – June 30, 2026";
const TARGET          = 30;       // enrollment target
const FAIGHT_RATE     = 10;       // $ per person per yoga class

// Yoga classes per month (April has 3 — kickoff Apr 7 has no yoga)
// June has 4 — closing Jun 30 has no yoga
const YOGA_CLASSES = { apr: 3, may: 4, jun: 4 };

// All 13 session dates (Tuesdays)
const SESSIONS = [
  "Apr 7 — Kickoff",
  "Apr 14", "Apr 21", "Apr 28",
  "May 5",  "May 12", "May 19", "May 26",
  "Jun 2",  "Jun 9",  "Jun 16", "Jun 23",
  "Jun 30 — Closing"
];

// Brand green
const GREEN_DARK  = "#1a3d0a";
const GREEN_LIGHT = "#d9ead3";


// ─── MAIN FUNCTION ─────────────────────────────────────────────────────────────

function createPlantedSFDashboard() {

  // Pull existing participant data from Kate's current sheet
  const existingRows = importExistingData();

  // Create the new spreadsheet
  const ss = SpreadsheetApp.create("Planted SF — Events & CRM");

  buildParticipantsTab(ss, existingRows);
  buildAttendanceTab(ss);
  buildDashboardTab(ss);

  // Move Dashboard to front
  const dash = ss.getSheetByName("📊 Dashboard");
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);

  const url = ss.getUrl();
  Logger.log("✅ Dashboard created: " + url);
  SpreadsheetApp.getUi().alert(
    "✅ Your Planted SF dashboard is ready!\n\nOpen it here:\n" + url +
    "\n\nBookmark this link — it's your main dashboard."
  );
}


// ─── TAB 1: PARTICIPANTS ───────────────────────────────────────────────────────

function buildParticipantsTab(ss, existingRows) {
  const P = ss.getActiveSheet();
  P.setName("Participants");

  // Column headers (A through R)
  // A  Name
  // B  Email
  // C  Phone
  // D  Instagram
  // E  Prior Artist's Way?
  // F  Payment Tier ($/mo)
  // G  Pay Method
  // H  M1 Paid ($)
  // I  M1 Date
  // J  M2 Paid ($)
  // K  M2 Date
  // L  M3 Paid ($)
  // M  M3 Date
  // N  Total Paid ($)
  // O  Program Total ($)
  // P  Balance Due ($)
  // Q  Spot Confirmed?
  // R  Notes

  const headers = [
    "Name", "Email", "Phone", "Instagram",
    "Prior Artist's Way?",
    "Payment Tier ($/mo)", "Pay Method",
    "M1 Paid ($)", "M1 Date",
    "M2 Paid ($)", "M2 Date",
    "M3 Paid ($)", "M3 Date",
    "Total Paid ($)", "Program Total ($)", "Balance Due ($)",
    "Spot Confirmed?", "Notes"
  ];

  P.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(P.getRange(1, 1, 1, headers.length));

  // Dropdowns
  P.getRange("F2:F100").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["100", "112.50", "125", "150", "175"], true)
      .build()
  );
  P.getRange("G2:G100").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Venmo", "Zelle", "Cash", "Other"], true)
      .build()
  );

  // Import existing data: Name(0), Email(1), M1(2), M2(3), M3(4)
  if (existingRows.length > 0) {
    const importData = existingRows.map(row => {
      const m1 = parseMoney(row[2]);
      const m2 = parseMoney(row[3]);
      const m3 = parseMoney(row[4]);
      return [
        row[0] || "", // A Name
        row[1] || "", // B Email
        "",           // C Phone (not in old sheet)
        "",           // D Instagram
        "",           // E Prior AW
        "",           // F Tier
        "",           // G Method
        m1 || "",     // H M1 Paid
        "",           // I M1 Date
        m2 || "",     // J M2 Paid
        "",           // K M2 Date
        m3 || "",     // L M3 Paid
        "",           // M M3 Date
      ];
    });
    P.getRange(2, 1, importData.length, 13).setValues(importData);
  }

  // Formulas for rows 2–100
  for (let r = 2; r <= 100; r++) {
    // N: Total Paid = M1 + M2 + M3
    P.getRange(r, 14).setFormula(
      `=IFERROR(H${r},0)+IFERROR(J${r},0)+IFERROR(L${r},0)`
    );
    // O: Program Total = Tier × 3 months
    P.getRange(r, 15).setFormula(
      `=IF(F${r}<>"",VALUE(F${r})*3,"")`
    );
    // P: Balance Due = Program Total - Total Paid
    P.getRange(r, 16).setFormula(
      `=IF(O${r}<>"",O${r}-N${r},"")`
    );
    // Q: Spot Confirmed
    P.getRange(r, 17).setFormula(
      `=IF(N${r}>0,"✓ Confirmed","⏳ Pending")`
    );
  }

  // Number formatting
  P.getRange("H2:H100").setNumberFormat("$#,##0.00");
  P.getRange("J2:J100").setNumberFormat("$#,##0.00");
  P.getRange("L2:L100").setNumberFormat("$#,##0.00");
  P.getRange("N2:P100").setNumberFormat("$#,##0.00");
  P.getRange("I2:I100").setNumberFormat("mm/dd/yyyy");
  P.getRange("K2:K100").setNumberFormat("mm/dd/yyyy");
  P.getRange("M2:M100").setNumberFormat("mm/dd/yyyy");

  // Conditional formatting on Q (Spot Confirmed)
  const qRange = P.getRange("Q2:Q100");
  P.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Confirmed")
      .setBackground(GREEN_LIGHT).setFontColor("#274e13")
      .setRanges([qRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Pending")
      .setBackground("#fff2cc").setFontColor("#856404")
      .setRanges([qRange]).build()
  ]);

  // Layout
  P.setFrozenRows(1);
  P.setFrozenColumns(2);
  P.setColumnWidth(1, 160);   // Name
  P.setColumnWidth(2, 220);   // Email
  P.setColumnWidth(3, 130);   // Phone
  P.setColumnWidth(4, 140);   // Instagram
  P.setColumnWidth(5, 150);   // Prior AW
  P.setColumnWidth(6, 130);   // Tier
  P.setColumnWidth(7, 110);   // Method
  P.setColumnWidth(18, 220);  // Notes
}


// ─── TAB 2: ATTENDANCE ─────────────────────────────────────────────────────────

function buildAttendanceTab(ss) {
  const A = ss.insertSheet("Attendance");

  const headers = ["Name", ...SESSIONS, "Total Attended", "Attendance %"];
  A.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(A.getRange(1, 1, 1, headers.length));

  // Names linked from Participants tab
  const nameFormulas = [];
  const totalFormulas = [];
  const pctFormulas = [];

  for (let r = 2; r <= 100; r++) {
    nameFormulas.push([`=IF(Participants!A${r}<>"",Participants!A${r},"")`]);
    totalFormulas.push([`=COUNTIF(B${r}:N${r},TRUE)`]);  // B–N = 13 session cols
    pctFormulas.push([`=IF(O${r}>0,O${r}/13,"")`]);      // O = Total Attended
  }

  A.getRange(2, 1, 99, 1).setFormulas(nameFormulas);
  A.getRange(2, 2, 99, 13).insertCheckboxes();            // 13 session columns
  A.getRange(2, 15, 99, 1).setFormulas(totalFormulas);    // col 15 = O
  A.getRange(2, 16, 99, 1).setFormulas(pctFormulas);      // col 16 = P
  A.getRange("P2:P100").setNumberFormat("0%");

  // Layout
  A.setFrozenRows(1);
  A.setFrozenColumns(1);
  A.setColumnWidth(1, 160);
  for (let c = 2; c <= 14; c++) A.setColumnWidth(c, 90);
  A.setColumnWidth(15, 110); // Total Attended
  A.setColumnWidth(16, 100); // Attendance %
}


// ─── TAB 3: DASHBOARD ──────────────────────────────────────────────────────────

function buildDashboardTab(ss) {
  const D = ss.insertSheet("📊 Dashboard");

  // Title
  D.getRange("A1").setValue("🌱 Planted SF — Events & CRM")
    .setFontSize(22).setFontWeight("bold").setFontColor(GREEN_DARK);
  D.getRange("A2").setValue(`${PROGRAM_NAME}  |  ${PROGRAM_DATES}`)
    .setFontSize(12).setFontColor("#666666").setFontStyle("italic");

  // ── ENROLLMENT ──
  sectionHeader(D, "A4", "ENROLLMENT");
  row(D, 5,  "Total Applied",          "=COUNTA(Participants!A2:A1000)");
  row(D, 6,  "✓ Confirmed (Paid)",     '=COUNTIF(Participants!Q2:Q1000,"✓ Confirmed")');
  row(D, 7,  "⏳ Pending (No Payment)",'=COUNTIF(Participants!Q2:Q1000,"⏳ Pending")');
  row(D, 8,  "Target Enrollment",      TARGET);
  row(D, 9,  "Spots Remaining",        `=${TARGET}-B6`);

  // ── REVENUE SUMMARY ──
  sectionHeader(D, "A11", "REVENUE SUMMARY");

  // Column headers
  D.getRange("B12").setValue("Month 1 (Apr)").setFontWeight("bold").setHorizontalAlignment("center");
  D.getRange("C12").setValue("Month 2 (May)").setFontWeight("bold").setHorizontalAlignment("center");
  D.getRange("D12").setValue("Month 3 (Jun)").setFontWeight("bold").setHorizontalAlignment("center");
  D.getRange("E12").setValue("3-Month Total").setFontWeight("bold").setHorizontalAlignment("center");

  row(D, 13, "Confirmed Participants",
    "=B6", "=B6", "=B6", "");
  row(D, 14, "Gross Collected ($)",
    "=SUM(Participants!H2:H1000)",
    "=SUM(Participants!J2:J1000)",
    "=SUM(Participants!L2:L1000)",
    "=B14+C14+D14");
  row(D, 15, `theFaight Payout ($)`,
    `=B6*${FAIGHT_RATE}*${YOGA_CLASSES.apr}`,
    `=B6*${FAIGHT_RATE}*${YOGA_CLASSES.may}`,
    `=B6*${FAIGHT_RATE}*${YOGA_CLASSES.jun}`,
    "=B15+C15+D15");
  row(D, 16, "Kate's Take-Home ($)",
    "=B14-B15", "=C14-C15", "=D14-D15", "=E14-E15");

  // Format revenue rows
  D.getRange("B14:E16").setNumberFormat("$#,##0");
  D.getRange("B16:E16").setFontWeight("bold").setBackground("#e8f5e9");
  D.getRange("A16").setFontWeight("bold");

  // ── REVENUE BY TIER ──
  sectionHeader(D, "A18", "REVENUE BY TIER");
  ["Tier", "# People", "Monthly", "3-Month Total"].forEach((h, i) =>
    D.getRange(19, i + 1).setValue(h).setFontWeight("bold")
  );

  [100, 112.50, 125, 150, 175].forEach((t, i) => {
    const r = 20 + i;
    D.getRange(r, 1).setValue(`$${t}/mo`);
    D.getRange(r, 2).setFormula(`=COUNTIF(Participants!F2:F1000,"${t}")`);
    D.getRange(r, 3).setFormula(`=B${r}*${t}`).setNumberFormat("$#,##0");
    D.getRange(r, 4).setFormula(`=C${r}*3`).setNumberFormat("$#,##0");
  });

  // ── EMAIL LIST ──
  sectionHeader(D, "A27", "EMAIL LIST — Confirmed Participants");
  D.getRange("A28").setValue("Copy the line below → paste into Gmail BCC field:")
    .setFontColor("#666666").setFontStyle("italic");
  D.getRange("A29")
    .setFormula(
      '=IFERROR(TEXTJOIN(", ",TRUE,FILTER(Participants!B2:B1000,Participants!Q2:Q1000="✓ Confirmed")),"No confirmed participants yet — add payment data to Participants tab")'
    )
    .setWrap(true)
    .setBackground("#f8f8f8")
    .setFontColor("#1a1a1a");

  // Layout
  D.setColumnWidth(1, 220);
  D.setColumnWidth(2, 140);
  D.setColumnWidth(3, 140);
  D.setColumnWidth(4, 140);
  D.setColumnWidth(5, 150);
  D.setRowHeight(29, 80);
}


// ─── IMPORT EXISTING DATA ──────────────────────────────────────────────────────

function importExistingData() {
  try {
    const src = SpreadsheetApp.openById(SOURCE_SHEET_ID);
    const sheets = src.getSheets();
    for (const s of sheets) {
      const lastCol = s.getLastColumn();
      if (lastCol < 1) continue;
      const headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.some(h => /payment/i.test(String(h)))) {
        const lastRow = s.getLastRow();
        if (lastRow <= 1) return [];
        return s.getRange(2, 1, lastRow - 1, Math.min(5, lastCol)).getValues();
      }
    }
  } catch (e) {
    Logger.log("Could not import existing data: " + e.message);
  }
  return [];
}


// ─── HELPERS ───────────────────────────────────────────────────────────────────

function styleHeader(range) {
  range
    .setBackground(GREEN_DARK)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11);
}

function sectionHeader(sheet, cell, title) {
  sheet.getRange(cell)
    .setValue(title)
    .setFontWeight("bold")
    .setFontSize(13)
    .setFontColor(GREEN_DARK);
}

function row(sheet, rowNum, label, b, c, d, e) {
  sheet.getRange(rowNum, 1).setValue(label);
  if (b !== undefined && b !== "") setCell(sheet, rowNum, 2, b);
  if (c !== undefined && c !== "") setCell(sheet, rowNum, 3, c);
  if (d !== undefined && d !== "") setCell(sheet, rowNum, 4, d);
  if (e !== undefined && e !== "") setCell(sheet, rowNum, 5, e);
}

function setCell(sheet, r, c, val) {
  const cell = sheet.getRange(r, c);
  if (typeof val === "string" && val.startsWith("=")) cell.setFormula(val);
  else cell.setValue(val);
}

function parseMoney(val) {
  const n = parseFloat(String(val).replace(/[$,\s]/g, ""));
  return isNaN(n) ? 0 : n;
}
