/* ============================================================
   TELEMEDICINE 465 AAM-UPHC — DAILY PERFORMANCE PORTAL
   VERSION 3.0 — XLSM FORMULA-FAITHFUL ENGINE

   RAW INPUTS:
   01. HWCConsultationReferredSummaryReport_export_DD-MM.xlsx
   02. HWCConsultationReceiveSummaryReport_export_DD-MM.xlsx

   IMPORTANT:
   The supplied XLSM is the calculation model. The portal does
   NOT calculate from arbitrary district names. It reproduces the
   Formula sheet's configured A/B values and the exact C/G/J/M
   criteria, then applies the exact D/E/F/H/I/K/L/N/O/R/S logic.
   The Referred input is restricted to the UPHC records that feed
   the XLSM "pivot" sheet before evaluating C/J.
   ============================================================ */
"use strict";

const APP = {
  rows: [], charts: {}, generated: false, reportDate: "",
  referredFile: "", receiveFile: "",
  raw: { referred: [], receive: [] }, config: []
};

const $ = id => document.getElementById(id);
const text = v => String(v ?? "").trim();
const norm = v => text(v).replace(/\s+/g, " ").toLowerCase();
const num = v => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const n = Number(String(v ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
};
const fmt = v => num(v).toLocaleString("en-IN", {maximumFractionDigits: 2});
const fmt0 = v => Math.round(num(v)).toLocaleString("en-IN");
const pct0 = v => `${Math.round(num(v))}%`;
const excelRank = (values, value, descending = true) => {
  const n = num(value);
  return 1 + values.filter(v => descending ? num(v) > n : num(v) < n).length;
};

// Dense ranking: ties share the same rank and the next distinct value
// receives the immediately following rank (e.g. 1,2,3,3,3,4,5,5,5,6).
const denseRank = (values, value, descending = true) => {
  const n = num(value);
  const distinct = [...new Set(values.map(num))];
  distinct.sort((a,b) => descending ? b-a : a-b);
  const idx = distinct.findIndex(v => v === n);
  return idx >= 0 ? idx + 1 : distinct.length + 1;
};
const contains = (value, criterion) =>
  norm(value).includes(norm(String(criterion).replace(/\*/g, "")));
const startsWithText = (value, criterion) =>
  norm(value).startsWith(norm(String(criterion).replace(/\*/g, "")));

/*
 * Exact Formula-sheet configuration extracted from the supplied XLSM.
 * cFormula = [C] Active AAM-UPHCs
 * gFormula = [G] Active Doctors
 * jFormula = [J] Spoke Total Consultations
 * mFormula = [M] HUB Total Consultations
 */
const FORMULA_CONFIG = [
  {
    "sl": 1,
    "district": "Basirhat HD",
    "operational": 5,
    "enabled": 5,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Basirhat\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Basirhat\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"24 PARAGANAS NORTH\",pivot!$B:$B,\"*\"&\" Basirhat\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC BASIRHAT\"&\"*\",rec!F:F)"
  },
  {
    "sl": 2,
    "district": "Jhargram",
    "operational": 1,
    "enabled": 1,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" jhargram\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Jhargram\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"jhargram\",pivot!$B:$B,\"*\"&\" jhargram\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Jhargram\"&\"*\",rec!F:F)"
  },
  {
    "sl": 3,
    "district": "Howrah",
    "operational": 32,
    "enabled": 24,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Howrah\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Howrah\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"howrah\",pivot!$B:$B,\"*\"&\" howrah\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Howrah\"&\"*\",rec!F:F)"
  },
  {
    "sl": 4,
    "district": "Uttar Dinajpur",
    "operational": 6,
    "enabled": 6,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Uttar Dinajpur\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Uttar Dinajpur\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"DINAJPUR UTTAR\",pivot!$B:$B,\"*\"&\" Uttar Dinajpur\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Uttar Dinajpur\"&\"*\",rec!F:F)"
  },
  {
    "sl": 5,
    "district": "Purba Medinipore",
    "operational": 6,
    "enabled": 6,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Purba Medinipur\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Purba Medinipur\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"MEDINIPUR EAST\",pivot!$B:$B,\"*\"&\" PURBA MEDINIPUR\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Purba Medinipur\"&\"*\",rec!F:F)"
  },
  {
    "sl": 6,
    "district": "Rampurhat HD",
    "operational": 2,
    "enabled": 1,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Rampurhat\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Rampurhat\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"BIRBHUM\",pivot!$B:$B,\"*\"&\" RAMPURHAT\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Rampurhat\"&\"*\",rec!F:F)"
  },
  {
    "sl": 7,
    "district": "Bishnupur HD",
    "operational": 2,
    "enabled": 2,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Bishnupur\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Bishnupur\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"BANKURA\",pivot!$B:$B,\"*\"&\" Bishnupur\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC BISHNUPUR\"&\"*\",rec!F:F)"
  },
  {
    "sl": 8,
    "district": "Paschim Medinipore",
    "operational": 11,
    "enabled": 11,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Paschim Medinipur\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Paschim Medinipur\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"MEDINIPUR WEST\",pivot!$B:$B,\"*\"&\" Paschim Medinipur\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Paschim Medinipur\"&\"*\",rec!F:F)"
  },
  {
    "sl": 9,
    "district": "Dakshin Dinajpur",
    "operational": 4,
    "enabled": 4,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Dakshin Dinajpur\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Dakshin Dinajpur\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"DINAJPUR DAKSHIN\",pivot!$B:$B,\"*\"&\" DAKSHIN DINAJPUR\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Dakshin Dinajpur\"&\"*\",rec!F:F)"
  },
  {
    "sl": 10,
    "district": "Paschim Bardhaman",
    "operational": 38,
    "enabled": 38,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Paschim Bardhaman\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Paschim BUrdwan\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"PASCHIM BARDHAMAN\",pivot!$B:$B,\"*\"&\" PASCHIM BARDHAMAN\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Paschim BURDWAN\"&\"*\",rec!F:F)"
  },
  {
    "sl": 11,
    "district": "Nandigram HD",
    "operational": 2,
    "enabled": 2,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Nandigram\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Nandigram\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"MEDINIPUR EAST\",pivot!$B:$B,\"*\"&\" NANDIGRAM\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Nandigram\"&\"*\",rec!F:F)"
  },
  {
    "sl": 12,
    "district": "Purba Bardhaman",
    "operational": 10,
    "enabled": 9,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" purba bardhaman\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Purba Bardhaman\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"Purba Bardhaman\",pivot!$B:$B,\"*\"&\" Purba Bardhaman\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Purba Bardhaman\"&\"*\",rec!F:F)"
  },
  {
    "sl": 13,
    "district": "Birbhum",
    "operational": 5,
    "enabled": 5,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Birbhum\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Birbhum\"&\"*\")-COUNTIFS(rec!B:B,\"HUB UPHC Birbhum\",rec!D:D,\"*\"&\" RPH\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"Birbhum\",pivot!$B:$B,\"*\"&\" Birbhum\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC BIRBHUM\"&\"*\",rec!F:F)"
  },
  {
    "sl": 14,
    "district": "South 24 Parganas",
    "operational": 20,
    "enabled": 20,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" spr\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC South Twenty Four Prgs\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"24 PARAGANAS SOUTH\",pivot!$B:$B,\"*\"&\" SPR\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC South Twenty Four Prgs\"&\"*\",rec!F:F)"
  },
  {
    "sl": 15,
    "district": "Bankura",
    "operational": 3,
    "enabled": 3,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Bankura\"&\"*\",pivot!M:M,\">=10\")-COUNTIFS(pivot!L:L,\"*\"&\" bishnupur\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Bankura\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"BANKURA\",pivot!$B:$B,\"*\"&\" BANKURA\"&\"*\")-SUMIFS(pivot!$F:$F,pivot!$C:$C,\"BANKURA\",pivot!$B:$B,\"*\"&\" Bishnupur\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC BANKURA\"&\"*\",rec!F:F)"
  },
  {
    "sl": 16,
    "district": "Jalpaiguri",
    "operational": 3,
    "enabled": 3,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" jalpaiguri\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Jalpaiguri\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"JALPAIGURI\",pivot!$B:$B,\"*\"&\" jALPAIGURI\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Jalpaiguri\"&\"*\",rec!F:F)"
  },
  {
    "sl": 17,
    "district": "Nadia",
    "operational": 17,
    "enabled": 17,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Nadia\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Nadia\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"NADIA\",pivot!$B:$B,\"*\"&\" NADIA\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Nadia\"&\"*\",rec!F:F)"
  },
  {
    "sl": 18,
    "district": "Murshidabad",
    "operational": 14,
    "enabled": 13,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Murshidabad\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Murshidabad\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"MURSHIDABAD\",pivot!$B:$B,\"*\"&\" MURSHIDABAD\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Murshidabad\"&\"*\",rec!F:F)"
  },
  {
    "sl": 19,
    "district": "KMC",
    "operational": 144,
    "enabled": 144,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" kolkata\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Kolkata\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"KOLKATA\",pivot!$B:$B,\"*\"&\" KOLKATA\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Kolkata\"&\"*\",rec!F:F)"
  },
  {
    "sl": 20,
    "district": "North 24 Parganas",
    "operational": 98,
    "enabled": 98,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" npr\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC North Twenty Four Prgs\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"24 PARAGANAS NORTH\",pivot!$B:$B,\"*\"&\" NPR\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC North Twenty Four Prgs\"&\"*\",rec!F:F)"
  },
  {
    "sl": 21,
    "district": "Darjeeling SMP",
    "operational": 10,
    "enabled": 10,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" DARJEELING SMP\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Darjeeling SMP\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"DARJEELING\",pivot!$B:$B,\"*\"&\" DARJEELING SMP\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Darjeeling SMP\"&\"*\",rec!F:F)"
  },
  {
    "sl": 22,
    "district": "COOCHBEHAR",
    "operational": 2,
    "enabled": 2,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Cooch Behar\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC CoochBehar\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"COOCHBEHAR\",pivot!$B:$B,\"*\"&\" COOCH BEHAR\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC COOCHBEHAR\"&\"*\",rec!F:F)"
  },
  {
    "sl": 23,
    "district": "Hooghly",
    "operational": 29,
    "enabled": 29,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Hooghly\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Hooghly\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"HOOGHLY\",pivot!$B:$B,\"*\"&\" HOOGHLY\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Hooghly\"&\"*\",rec!F:F)"
  },
  {
    "sl": 24,
    "district": "Malda",
    "operational": 6,
    "enabled": 6,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" MALDAH\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Maldah\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"MALDAH\",pivot!$B:$B,\"*\"&\" MALDAH\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Maldah\"&\"*\",rec!F:F)"
  },
  {
    "sl": 25,
    "district": "Darjeeling GTA",
    "operational": 3,
    "enabled": 2,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Darjeeling GTA\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC DARJEELING GTA\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"DARJEELING\",pivot!$B:$B,\"*\"&\" DARJEELING GTA\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC DARJEELING GTA\"&\"*\",rec!F:F)"
  },
  {
    "sl": 26,
    "district": "Alipurduar ",
    "operational": 1,
    "enabled": 1,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Alipurduar\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC ALIPURDUAR\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"ALIPURDUAR\",pivot!$B:$B,\"*\"&\" ALIPURDUAR\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC ALIPURDUAR\"&\"*\",rec!F:F)"
  },
  {
    "sl": 27,
    "district": "Diamond Harbour HD",
    "operational": 1,
    "enabled": 1,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Diamond Harbour\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC Diamond Harbour\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"24 PARAGANAS SOUTH\",pivot!$B:$B,\"*\"&\" Diamond Harbour\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC Diamond Harbour\"&\"*\",rec!F:F)"
  },
  {
    "sl": 28,
    "district": "Purulia",
    "operational": 2,
    "enabled": 2,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" Purulia\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC PURULIA\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"Purulia\",pivot!$B:$B,\"*\"&\" Purulia\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC PURULIA\"&\"*\",rec!F:F)"
  },
  {
    "sl": 29,
    "district": "Kalimpong",
    "operational": 1,
    "enabled": 0,
    "cFormula": "=COUNTIFS(pivot!L:L,\"*\"&\" kalimpong\"&\"*\",pivot!M:M,\">=10\")",
    "gFormula": "=COUNTIF(rec!B:B,\"HUB UPHC KALIMPONG\"&\"*\")",
    "jFormula": "=SUMIFS(pivot!$F:$F,pivot!$C:$C,\"KALIMPONG\",pivot!$B:$B,\"*\"&\" KALIMPONG\"&\"*\")",
    "mFormula": "=SUMIF(rec!B:B,\"HUB UPHC KALIMPONG\"&\"*\",rec!F:F)"
  }
];

const DASHBOARD_ORDER = [
  "Alipurduar",
  "Uttar Dinajpur",
  "Purba Bardhaman",
  "Purba Medinipore",
  "Jhargram",
  "Malda",
  "Birbhum",
  "Rampurhat HD",
  "Nandigram HD",
  "Howrah",
  "Bishnupur HD",
  "Purulia",
  "South 24 Parganas",
  "Bankura",
  "Paschim Medinipore",
  "Jalpaiguri",
  "Murshidabad",
  "COOCHBEHAR",
  "Nadia",
  "Dakshin Dinajpur",
  "North 24 Parganas",
  "KMC",
  "Hooghly",
  "Paschim Bardhaman",
  "Darjeeling GTA",
  "Basirhat HD",
  "Darjeeling SMP",
  "Diamond Harbour HD",
  "Kalimpong"
];
APP.config = FORMULA_CONFIG;

function setStatus(message, type="info") {
  const el = $("status");
  el.className = `status ${type}`;
  el.innerHTML = message;
}
function showLoading(show, message="Processing…") {
  $("loading").classList.toggle("hidden", !show);
  $("loadingText").textContent = message;
}
function escapeHtml(s) {
  return String(s).replace(/[&<>'"]/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;","\"":"&quot;"
  }[c]));
}
function formatDateLong(value) {
  if (!value) return "--";
  const d = new Date(`${value}T00:00:00`);
  return Number.isNaN(d.getTime()) ? value :
    d.toLocaleDateString("en-IN",{day:"2-digit",month:"long",year:"numeric"});
}
function setDefaultDate() {
  const d = new Date();
  const local = new Date(d.getTime() - d.getTimezoneOffset()*60000);
  $("reportDate").value = local.toISOString().slice(0,10);
}
function tickClock() {
  const d = new Date();
  $("liveDate").textContent = d.toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
  $("liveTime").textContent = d.toLocaleTimeString("en-IN",{hour12:false});
}
function parseDateFromFilename(name) {
  const m = text(name).match(/(\d{1,2})-(\d{1,2})(?:\.[^.]+)?$/) ||
            text(name).match(/(\d{1,2})-(\d{1,2})/);
  if (!m) return "";
  const day=Number(m[1]), month=Number(m[2]), year=new Date().getFullYear();
  if(day<1||day>31||month<1||month>12) return "";
  return `${year}-${String(month).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
}

function sheetRows(ws) {
  const matrix = XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  if (!matrix.length) return [];
  const headers = matrix[0].map(v=>text(v));
  return matrix.slice(1)
    .filter(r=>r.some(v=>v!==null&&v!==""))
    .map(r=>{
      const o={};
      headers.forEach((h,i)=>{if(h)o[h]=r[i];});
      return o;
    });
}
function validateRaw(rows,label) {
  if(!rows.length) throw new Error(`${label} is empty.`);
  const required=["Health Facility","District","User","Completed"];
  const keys=Object.keys(rows[0]);
  const missing=required.filter(k=>!keys.includes(k));
  if(missing.length) throw new Error(`${label} is missing column(s): ${missing.join(", ")}`);
}
async function readRawFile(file) {
  const data=await file.arrayBuffer();
  const book=XLSX.read(data,{type:"array",cellDates:true});
  if(!book.SheetNames.length) throw new Error(`${file.name}: no worksheet found.`);
  return {rows:sheetRows(book.Sheets[book.SheetNames[0]]),sheet:book.SheetNames[0]};
}

/*
 * The XLSM Formula sheet's C and J formulas operate on the "pivot"
 * sheet, not on every record in the raw referred export.
 *
 * In the supplied workbook, the pivot source consists of records
 * whose Health Facility begins with "UPHC". Reproduce that filter
 * before C/J evaluation.
 */
function pivotSourceRows(referredRows) {
  return referredRows.filter(r => /^\s*UPHC\b/i.test(text(r["Health Facility"])));
}
function buildPivot(referredRows) {
  const map=new Map();
  for(const r of referredRows){
    const facility=text(r["Health Facility"]);
    if(!facility) continue;
    const key=norm(facility);
    if(!map.has(key)) map.set(key,{facility,completed:0});
    map.get(key).completed += num(r["Completed"]);
  }
  return [...map.values()];
}

/* Excel COUNTIFS(pivot!L:L,"*"&criterion&"*",pivot!M:M,">=10") */
function exactActiveCount(pivotRows, formula) {
  const q=(formula.match(/"([^"]*)"/g)||[]).map(s=>s.slice(1,-1));
  // Normal formula: ["*"," criterion","*",">=10"]
  let count=pivotRows.filter(r=>contains(r.facility,q[1]) && r.completed>=10).length;

  // Bankura formula contains a subtraction:
  // -COUNTIFS(...,"*"&" bishnupur"&"*",...,">=10")
  if(q.length>=8){
    count -= pivotRows.filter(r=>contains(r.facility,q[5]) && r.completed>=10).length;
  }
  return count;
}

/* Excel SUMIFS(pivot!F:F,pivot!C:C,district,pivot!B:B,"*"&facility&"*") */
function exactSpokeTotal(pivotSource, formula) {
  const q=(formula.match(/"([^"]*)"/g)||[]).map(s=>s.slice(1,-1));
  let total=pivotSource
    .filter(r=>norm(r["District"])===norm(q[0]) && contains(r["Health Facility"],q[2]))
    .reduce((s,r)=>s+num(r["Completed"]),0);

  // Bankura formula subtracts the Bishnupur component.
  if(q.length>=8){
    total -= pivotSource
      .filter(r=>norm(r["District"])===norm(q[0]) && contains(r["Health Facility"],q[6]))
      .reduce((s,r)=>s+num(r["Completed"]),0);
  }
  return total;
}

/* Excel COUNTIF(rec!B:B, prefix&"*") with Birbhum exception. */
function exactActiveDoctors(receiveRows, formula) {
  const q=(formula.match(/"([^"]*)"/g)||[]).map(s=>s.slice(1,-1));
  let count=receiveRows.filter(r=>startsWithText(r["Health Facility"],q[0])).length;

  // Exact XLSM exception:
  // -COUNTIFS(rec!B:B,"HUB UPHC Birbhum",rec!D:D,"*"&" RPH")
  if(q.length>=5){
    count -= receiveRows.filter(r =>
      norm(r["Health Facility"])===norm(q[0]) &&
      contains(r["User"],q[4])
    ).length;
  }
  return count;
}

/* Excel SUMIF(rec!B:B, prefix&"*", rec!F:F) */
function exactHubTotal(receiveRows, formula) {
  const q=(formula.match(/"([^"]*)"/g)||[]).map(s=>s.slice(1,-1));
  return receiveRows
    .filter(r=>startsWithText(r["Health Facility"],q[0]))
    .reduce((s,r)=>s+num(r["Completed"]),0);
}

function calculateReport(referredRows,receiveRows){
  const pivotSource=pivotSourceRows(referredRows);
  const pivotRows=buildPivot(pivotSource);

  /*
   * This section mirrors Formula!A4:S33:
   * E=C, F=D, G=E, H=F, I=G, J=H, K=I,
   * L=J, M=K, N=L, O=M, P=N, Q=O, R=Composite, S=Overall.
   */
  const base=APP.config.map(cfg=>{
    const active=exactActiveCount(pivotRows,cfg.cFormula);

    // [D] = IFERROR(E/D*100,0)
    const activePct=cfg.enabled ? active/cfg.enabled*100 : 0;

    // [F] = MAX(1,D/5)
    const expectedDoctors=Math.max(1,cfg.enabled/5);

    // [G] = COUNTIF / special Birbhum exclusion
    const activeDoctors=exactActiveDoctors(receiveRows,cfg.gFormula);

    // [H] = ROUND(IFERROR(I/D,0)*100,0)
    const availability=cfg.enabled ? Math.round(activeDoctors/cfg.enabled*100) : 0;

    // [J] = exact SUMIFS against pivot source
    const spoke=exactSpokeTotal(pivotSource,cfg.jFormula);

    // [K] = IFERROR(J/D,0)
    const spokePerEnabled=cfg.enabled ? spoke/cfg.enabled : 0;

    // [M] = exact SUMIF against receive raw
    const hub=exactHubTotal(receiveRows,cfg.mFormula);

    // [N] = ROUND(IFERROR(M/G,0),0)
    const hubPerDoctor=activeDoctors ? Math.round(hub/activeDoctors) : 0;

    return {
      ...cfg, active, activePct, expectedDoctors, activeDoctors,
      availability, spoke, spokePerEnabled, hub, hubPerDoctor
    };
  });

  const activeRanks=base.map(r=>excelRank(base.map(x=>x.activePct),r.activePct,true));
  const availabilityRanks=base.map(r=>excelRank(base.map(x=>x.availability),r.availability,true));
  const spokeRanks=base.map(r=>excelRank(base.map(x=>x.spokePerEnabled),r.spokePerEnabled,true));
  const hubRanks=base.map(r=>excelRank(base.map(x=>x.hubPerDoctor),r.hubPerDoctor,true));

  /*
   * Formula!R:
   * =(0.25*G5+0.25*N5)+(0.25*K5+0.25*Q5)
   */
  const ranked=base.map((r,i)=>({
    ...r,
    activeRank:activeRanks[i],
    availabilityRank:availabilityRanks[i],
    spokeRank:spokeRanks[i],
    hubRank:hubRanks[i],
    composite:
      0.25*activeRanks[i]+
      0.25*spokeRanks[i]+
      0.25*availabilityRanks[i]+
      0.25*hubRanks[i]
  }));

  /*
   * Overall rank: dense ascending rank of Composite Score.
   * Ties share a rank and the next distinct score gets the next
   * consecutive rank, with no skipped numbers.
   */
  const compositeValues=ranked.map(r=>r.composite);
  const final=ranked.map(r=>({
    ...r,
    // Overall ranking uses dense ranking as requested: ties do not create
    // skipped rank numbers (1,2,3,3,3,4,5,5,5,6...).
    overallRank:denseRank(compositeValues,r.composite,false)
  }));

  /*
   * The XLSM dashboard is physically sorted by Overall Rank.
   * Use its existing tie order as the secondary key so ties render
   * in the same order as the workbook.
   */
  const orderMap=new Map(DASHBOARD_ORDER.map((d,i)=>[norm(d),i]));
  final.sort((a,b)=>{
    const rankDiff=a.overallRank-b.overallRank;
    if(rankDiff) return rankDiff;
    return (orderMap.get(norm(a.district))??999)-
           (orderMap.get(norm(b.district))??999);
  });
  return final;
}

/* Dashboard row colours copied from dashboard sheet:
   Overall Rank 1–9   = #92D050 green
   Overall Rank 10–18 = workbook theme yellow
   Overall Rank 19+   = #FF0000 red
*/
function performanceBand(rank){
  if(rank<=9) return "excel-green";
  if(rank<=18) return "excel-yellow";
  return "excel-red";
}
function topBy(rows, field){
  return [...rows].sort((a,b)=>{
    const d=num(b[field])-num(a[field]);
    return d || (a.sl||0)-(b.sl||0);
  })[0] || null;
}

function lowSpokeCount(referredRows){
  const map=new Map();
  for(const r of pivotSourceRows(referredRows)){
    const facility=text(r["Health Facility"]);
    if(!facility) continue;
    const key=norm(facility);
    if(!map.has(key)) map.set(key,{facility,completed:0});
    map.get(key).completed += num(r["Completed"]);
  }
  return [...map.values()].filter(r=>r.completed>=1 && r.completed<5).length;
}

function renderKpis(rows){
  const bestSpoke=topBy(rows,"spoke");
  const bestHub=topBy(rows,"hub");
  const bestActive=topBy(rows,"activePct");
  const bestDoctor=topBy(rows,"availability");
  const lowCount=lowSpokeCount(APP.raw.referred);

  $("kpiBestSpoke").textContent=bestSpoke?bestSpoke.district:"--";
  $("kpiBestSpokeValue").textContent=bestSpoke?`${fmt0(bestSpoke.spoke)} consultations`:"-- consultations";
  $("kpiBestHub").textContent=bestHub?bestHub.district:"--";
  $("kpiBestHubValue").textContent=bestHub?`${fmt0(bestHub.hub)} consultations`:"-- consultations";
  $("kpiBestActiveDistrict").textContent=bestActive?bestActive.district:"--";
  $("kpiBestActivePct").textContent=bestActive?`${pct0(bestActive.activePct)} active AAM-UPHC`:"0% active AAM-UPHC";
  $("kpiBestDoctorDistrict").textContent=bestDoctor?bestDoctor.district:"--";
  $("kpiBestDoctorPct").textContent=bestDoctor?`${pct0(bestDoctor.availability)} doctor availability`:"0% doctor availability";
  $("kpiLowSpokeCount").textContent=fmt0(lowCount);

  // NEW KPI: AAM UPHCs with 0 consultation = [B] Telemedicine Enabled - [C] Active (state total)
  const totalB=rows.reduce((a,r)=>a+num(r.enabled),0);
  const totalC=rows.reduce((a,r)=>a+num(r.active),0);
  $("kpiZeroAamCount").textContent=fmt0(Math.max(0,totalB-totalC));

  // NEW KPI: Districts with 0 consultations today (Spoke [J] = 0 and HUB [M] = 0)
  const zeroDistricts=rows.filter(r=>num(r.spoke)===0 && num(r.hub)===0);
  $("kpiZeroDistrictCount").textContent=fmt0(zeroDistricts.length);
  $("kpiZeroDistrictNames").textContent=zeroDistricts.length
    ? zeroDistricts.map(r=>r.district.trim()).join(", ")
    : "No district with zero consultations";
}

function renderTable(rows){
  $("reportBody").innerHTML=rows.map((r,i)=>{
    const band=performanceBand(r.overallRank);
    return `<tr class="${band}">
      <td>${i+1}</td>
      <td class="district">${escapeHtml(r.district)}</td>
      <td>${fmt0(r.operational)}</td>
      <td>${fmt0(r.enabled)}</td>
      <td>${fmt0(r.active)}</td>
      <td>${Math.round(r.activePct)}%</td>
      <td>${r.activeRank}</td>
      <td>${fmt0(r.activeDoctors)}</td>
      <td>${Math.round(r.availability)}%</td>
      <td>${r.availabilityRank}</td>
      <td>${fmt0(r.spoke)}</td>
      <td>${Math.round(r.spokePerEnabled)}</td>
      <td>${r.spokeRank}</td>
      <td>${fmt0(r.hub)}</td>
      <td>${fmt0(r.hubPerDoctor)}</td>
      <td>${r.hubRank}</td>
      <td>${r.overallRank}</td>
    </tr>`;
  }).join("");

  const op=rows.reduce((a,r)=>a+r.operational,0);
  const en=rows.reduce((a,r)=>a+r.enabled,0);
  const ac=rows.reduce((a,r)=>a+r.active,0);
  const ex=rows.reduce((a,r)=>a+r.expectedDoctors,0);
  const dr=rows.reduce((a,r)=>a+r.activeDoctors,0);
  const sp=rows.reduce((a,r)=>a+r.spoke,0);
  const hu=rows.reduce((a,r)=>a+r.hub,0);

  // Keep the WEST BENGAL SUMMARY cards synchronized with the generated report.
  // These are the same aggregates used by the total row above.
  $("sumOperational").textContent=fmt0(op);
  $("sumEnabled").textContent=fmt0(en);
  $("sumActive").textContent=fmt0(ac);
  $("sumDoctors").textContent=fmt0(dr);
  $("sumSpoke").textContent=fmt0(sp);
  $("sumHub").textContent=fmt0(hu);

  $("reportFooter").innerHTML=`<tr class="total-row">
    <td colspan="2">WEST BENGAL</td>
    <td>${fmt0(op)}</td><td>${fmt0(en)}</td><td>${fmt0(ac)}</td>
    <td>${en?Math.round(ac/en*100):0}%</td><td>—</td>
    <td>${fmt0(dr)}</td><td>—</td><td>—</td>
    <td>${fmt0(sp)}</td><td>${en?Math.round(sp/en):0}</td><td>—</td>
    <td>${fmt0(hu)}</td><td>${dr?Math.round(hu/dr):0}</td><td>—</td>
    <td>—</td>
  </tr>
  <tr class="footnote-row">
    <td colspan="17"><b>** Active AAM-UPHCs :</b> AAM-UPHCs completed at least ten consultation Vide Memo No. HFW/NCD/462/2021 dated 8th October 2021</td>
  </tr>`;
}

function destroyCharts(){
  Object.values(APP.charts).forEach(c=>{try{c.destroy()}catch{}});
  APP.charts={};
}
function chartBase(){
  return {
    responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false}},
    scales:{x:{grid:{display:false},ticks:{font:{weight:"800",size:11},color:"#172033",maxRotation:35,minRotation:0}},y:{beginAtZero:true,grid:{color:"rgba(15,23,42,.08)"}}}
  };
}
function renderCharts(rows){
  destroyCharts();
  const top=[...rows].sort((a,b)=>b.spoke-a.spoke).slice(0,10);
  const active=[...rows].sort((a,b)=>b.activePct-a.activePct).slice(0,10);
  const doctors=[...rows].sort((a,b)=>b.availability-a.availability).slice(0,10);

  APP.charts.spoke=new Chart($("spokeChart"),{
    type:"bar",
    data:{labels:top.map(r=>r.district),datasets:[{data:top.map(r=>r.spoke),borderRadius:7}]},
    options:chartBase()
  });
  APP.charts.active=new Chart($("activeChart"),{
    type:"bar",
    data:{labels:active.map(r=>r.district),datasets:[{data:active.map(r=>r.activePct),borderRadius:7}]},
    options:{...chartBase(),scales:{...chartBase().scales,y:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"}}}}
  });
  APP.charts.doctor=new Chart($("doctorChart"),{
    type:"bar",
    data:{labels:doctors.map(r=>r.district),datasets:[{data:doctors.map(r=>r.availability),borderRadius:7}]},
    options:{...chartBase(),scales:{...chartBase().scales,y:{beginAtZero:true,max:120,ticks:{callback:v=>v+"%"}}}}
  });

}

async function generateReport(){
  const referred=$("referredFile").files[0], receive=$("receiveFile").files[0];
  if(!referred||!receive){
    setStatus("Please upload <b>both raw Excel files</b> before generating the report.","error");
    return;
  }
  try{
    showLoading(true,"Reading the two raw Excel files…");
    const [ref,rec]=await Promise.all([readRawFile(referred),readRawFile(receive)]);
    validateRaw(ref.rows,"Referred Summary file");
    validateRaw(rec.rows,"Receive Summary file");

    APP.raw.referred=ref.rows;
    APP.raw.receive=rec.rows;
    APP.referredFile=referred.name;
    APP.receiveFile=receive.name;

    const autoDate=parseDateFromFilename(referred.name)||parseDateFromFilename(receive.name);
    if(autoDate) $("reportDate").value=autoDate;
    APP.reportDate=$("reportDate").value;

    showLoading(true,"Reproducing the XLSM Formula sheet calculations…");
    const rows=calculateReport(APP.raw.referred,APP.raw.receive);
    APP.rows=rows;

    $("reportDateText").textContent=`${formatDateLong(APP.reportDate)} | 5:00PM`;
    renderKpis(rows);
    renderTable(rows);
    renderCharts(rows);

    $("reportSection").classList.remove("hidden");
    ["excelBtn","pdfBtn","pngBtn"].forEach(id=>$(id).disabled=false);


    APP.generated=true;
    setStatus(`<b>Report generated Successfully.</b>`,"success");
  }catch(err){
    console.error(err);
    setStatus(`<b>Could not generate report.</b> ${escapeHtml(err.message||String(err))}`,"error");
    $("reportSection").classList.add("hidden");
  }finally{showLoading(false);}
}

function resetApp(){
  APP.rows=[];APP.generated=false;APP.raw={referred:[],receive:[]};
  APP.referredFile="";APP.receiveFile="";destroyCharts();
  $("referredFile").value="";$("receiveFile").value="";
  $("referredName").textContent="No file selected";
  $("receiveName").textContent="No file selected";
  $("reportSection").classList.add("hidden");
  ["excelBtn","pdfBtn","pngBtn"].forEach(id=>$(id).disabled=true);
  setDefaultDate();
  setStatus("Upload both raw Excel files and click <b>Generate Report</b>.","info");
}

function excelBandStyle(rank){
  if(rank<=9) return {fill:"92D050",font:"000000"};
  if(rank<=18) return {fill:"FFF2CC",font:"000000"};
  return {fill:"FF0000",font:"FFFFFF"};
}

function downloadExcel(){
  if(!APP.generated)return;

  const titleRows=[
    ["GOVERNMENT OF WEST BENGAL"],
    ["HEALTH & FAMILY WELFARE DEPARTMENT"],
    ["TELEMEDICINE - AAM-UPHCs"],
    ["DAILY PERFORMANCE REPORT"],
    [`REPORT FOR: ${formatDateLong(APP.reportDate)} | 5:00PM`],
    [""]
  ];

  // Excel layout mirrors the visible dashboard and intentionally omits Composite Score.
  const groupRow=[
    "SL. NO.","DISTRICTS","ACTIVE SPOKES","","","","",
    "DOCTOR'S AVAILABILITY","","",
    "SPOKES PERFORMANCE ANALYSIS & HUB PERFORMANCE ANALYSIS","","","","","",
    "OVERALL RANK [Q]"
  ];
  const subGroupRow=[
    "","","","","","","","","","",
    "SPOKE","","","HUB","","",""
  ];
  const headers=[
    "Sl. No.","DISTRICTS","No of Operational AAM-UPHCs [A]","No. of Telemedicine Enabled AAM-UPHCs [B]","No of active AAM-UPHCs [C]",
    "% Active AAM-UPHCs [D]=[C/B*100]","Rank [E]","Active Doctors [G]","% of Availability [H]","Rank [I]",
    "No of Total Consultations [J]","No of consultations per operational AAM-UPHC K = [J/B]","Rank [L]",
    "Hub_Total Consultations [M]","No. of Consultations per Active Doctor [N]","Rank [O]","Overall Rank [Q]"
  ];

  const data=APP.rows.map((r,i)=>[
    i+1,r.district,r.operational,r.enabled,r.active,
    r.activePct/100,r.activeRank,r.activeDoctors,
    r.availability/100,r.availabilityRank,r.spoke,Math.round(r.spokePerEnabled),
    r.spokeRank,r.hub,Math.round(r.hubPerDoctor),r.hubRank,r.overallRank
  ]);

  const totals=APP.rows.reduce((a,r)=>{
    a.op+=r.operational;a.en+=r.enabled;a.ac+=r.active;
    a.dr+=r.activeDoctors;a.sp+=r.spoke;a.hu+=r.hub;return a;
  },{op:0,en:0,ac:0,ex:0,dr:0,sp:0,hu:0});

  data.push([
    "","WEST BENGAL",totals.op,totals.en,totals.ac,
    totals.en?totals.ac/totals.en:0,"",totals.dr,"","",
    totals.sp,totals.en?Math.round(totals.sp/totals.en):0,"",totals.hu,
    totals.dr?Math.round(totals.hu/totals.dr):0,"",""
  ]);
  data.push(["** Active AAM-UPHCs : AAM-UPHCs completed at least ten consultation Vide Memo No. HFW/NCD/462/2021 dated 8th October 2021"]);

  const aoa=[...titleRows,groupRow,subGroupRow,headers,...data];
  const ws=XLSX.utils.aoa_to_sheet(aoa);

  // Merge government/report headings and dashboard group headers.
  ws["!merges"]=[
    {s:{r:0,c:0},e:{r:0,c:16}},
    {s:{r:1,c:0},e:{r:1,c:16}},
    {s:{r:2,c:0},e:{r:2,c:16}},
    {s:{r:3,c:0},e:{r:3,c:16}},
    {s:{r:4,c:0},e:{r:4,c:16}},
    {s:{r:5,c:0},e:{r:5,c:16}},
    {s:{r:6,c:0},e:{r:8,c:0}},
    {s:{r:6,c:1},e:{r:8,c:1}},
    {s:{r:6,c:2},e:{r:6,c:6}},
    {s:{r:6,c:7},e:{r:6,c:9}},
    {s:{r:6,c:10},e:{r:6,c:15}},
    {s:{r:6,c:16},e:{r:8,c:16}},
    {s:{r:7,c:10},e:{r:7,c:12}},
    {s:{r:7,c:13},e:{r:7,c:15}}
  ];

  // Compact district column; redistribute width to metric columns so
  // grouped headers and sub-headers remain aligned without overlap.
  ws["!cols"]=[
    {wch:6},{wch:18},{wch:15},{wch:17},{wch:15},{wch:16},{wch:8},
    {wch:13},{wch:13},{wch:8},{wch:17},{wch:23},{wch:8},{wch:19},{wch:23},{wch:8},{wch:8}
  ];
  ws["!rows"]=[
    {hpt:22},{hpt:19},{hpt:20},{hpt:28},{hpt:20},{hpt:8},
    {hpt:28},{hpt:23},{hpt:38}
  ];
  ws["!freeze"]="A10";
  ws["!autofilter"]={ref:`A9:Q${10+APP.rows.length}`};
  ws["!sheetViews"]=[{showGridLines:false}];

  const border={
    top:{style:"thin",color:{rgb:"D9E2EA"}},
    bottom:{style:"thin",color:{rgb:"D9E2EA"}},
    left:{style:"thin",color:{rgb:"D9E2EA"}},
    right:{style:"thin",color:{rgb:"D9E2EA"}}
  };
  const center={horizontal:"center",vertical:"center",wrapText:true};

  const titleStyles=[
    {font:{bold:true,color:{rgb:"FFFFFF"},sz:13},fill:{patternType:"solid",fgColor:{rgb:"163E63"}},alignment:center},
    {font:{bold:true,color:{rgb:"FFFFFF"},sz:10},fill:{patternType:"solid",fgColor:{rgb:"163E63"}},alignment:center},
    {font:{bold:true,color:{rgb:"FFFFFF"},sz:11},fill:{patternType:"solid",fgColor:{rgb:"163E63"}},alignment:center},
    {font:{bold:true,color:{rgb:"FFFFFF"},sz:17},fill:{patternType:"solid",fgColor:{rgb:"163E63"}},alignment:center},
    {font:{bold:true,color:{rgb:"163E63"},sz:11},fill:{patternType:"solid",fgColor:{rgb:"EAF2F8"}},alignment:center},
    {fill:{patternType:"solid",fgColor:{rgb:"FFFFFF"}}}
  ];
  titleStyles.forEach((st,r)=>{
    const cell=ws[XLSX.utils.encode_cell({r,c:0})];
    if(cell)cell.s={...st,border};
  });

  const groupStyles={
    active:{font:{bold:true,color:{rgb:"FFFFFF"},sz:10},fill:{patternType:"solid",fgColor:{rgb:"355F7D"}},alignment:center,border},
    doctor:{font:{bold:true,color:{rgb:"FFFFFF"},sz:10},fill:{patternType:"solid",fgColor:{rgb:"4C6D55"}},alignment:center,border},
    perf:{font:{bold:true,color:{rgb:"FFFFFF"},sz:10},fill:{patternType:"solid",fgColor:{rgb:"8A5A31"}},alignment:center,border},
    overall:{font:{bold:true,color:{rgb:"FFFFFF"},sz:10},fill:{patternType:"solid",fgColor:{rgb:"5B4D79"}},alignment:center,border}
  };
  // Apply styles to all cells in rows 7-9, including merged-header anchor cells.
  for(let c=0;c<17;c++){
    const r6=ws[XLSX.utils.encode_cell({r:6,c})];
    const gs=(c<=6)?groupStyles.active:(c<=10?groupStyles.doctor:(c<=16?groupStyles.perf:groupStyles.overall));
    if(r6) r6.s={...gs,border};
    const r7=ws[XLSX.utils.encode_cell({r:7,c})];
    if(r7) r7.s={font:{bold:true,color:{rgb:"17364D"},sz:9},fill:{patternType:"solid",fgColor:{rgb:"EEF4F8"}},alignment:center,border};
    const r8=ws[XLSX.utils.encode_cell({r:8,c})];
    if(r8) r8.s={font:{bold:true,color:{rgb:"17364D"},sz:9},fill:{patternType:"solid",fgColor:{rgb:"F6F9FB"}},alignment:center,border};
  }
  // Re-apply merged group anchor colours.
  ws["A7"].s={...groupStyles.active};
  ws["B7"].s={...groupStyles.active};
  ws["C7"].s={...groupStyles.active};
  ws["H7"].s={...groupStyles.doctor};
  ws["K7"].s={...groupStyles.perf};
  ws["Q7"].s={...groupStyles.overall};
  ws["K8"].s={font:{bold:true,color:{rgb:"17364D"},sz:9},fill:{patternType:"solid",fgColor:{rgb:"F6F9FB"}},alignment:center,border};
  ws["N8"].s={font:{bold:true,color:{rgb:"17364D"},sz:9},fill:{patternType:"solid",fgColor:{rgb:"F6F9FB"}},alignment:center,border};

  const firstDataRow=9;
  for(let i=0;i<APP.rows.length;i++){
    const excelRow=firstDataRow+i;
    const style=excelBandStyle(APP.rows[i].overallRank);
    for(let c=0;c<17;c++){
      const cell=ws[XLSX.utils.encode_cell({r:excelRow,c})];
      if(!cell)continue;
      cell.s={
        fill:{patternType:"solid",fgColor:{rgb:style.fill}},
        font:{color:{rgb:style.font},bold:true,sz:(c===0||c===1)?11:10},
        alignment:{vertical:"center",horizontal:c===1?"left":"center",wrapText:true},
        border
      };
    }
  }

  const totalRow=firstDataRow+APP.rows.length;
  for(let c=0;c<17;c++){
    const cell=ws[XLSX.utils.encode_cell({r:totalRow,c})];
    if(cell)cell.s={font:{bold:true,color:{rgb:"FFFFFF"},sz:11},fill:{patternType:"solid",fgColor:{rgb:"163E63"}},alignment:center,border};
  }
  // Reference-style yellow/red Active AAM-UPHC note directly below the total row.
  const footnoteRow=totalRow+1;
  ws["A"+(footnoteRow+1)].s={
    font:{bold:true,color:{rgb:"D40000"},sz:10},
    fill:{patternType:"solid",fgColor:{rgb:"FFF200"}},
    alignment:{horizontal:"center",vertical:"center",wrapText:true},
    border
  };
  ws["!merges"].push({s:{r:footnoteRow,c:0},e:{r:footnoteRow,c:16}});
  ws["!rows"][footnoteRow]={hpt:24};

  // Number formats: actual percentages in Excel, not decimals.
  for(let r=firstDataRow;r<=totalRow;r++){
    if(ws[`F${r+1}`])ws[`F${r+1}`].z="0%";   // % Active AAM-UPHCs [D]
    if(ws[`I${r+1}`])ws[`I${r+1}`].z="0%";   // % of Availability [H]
    for(const col of ["C","D","E","G","H","J","K","L","M","N","O","P","Q"]){
      if(ws[`${col}${r+1}`])ws[`${col}${r+1}`].z="#,##0";
    }
  }

  const noteRow=totalRow+2;
  ws[XLSX.utils.encode_cell({r:noteRow,c:0})]={v:"Source / calculation note",t:"s",s:{font:{bold:true,color:{rgb:"163E63"},sz:9},alignment:{vertical:"top"}}};
  ws[XLSX.utils.encode_cell({r:noteRow,c:1})]={v:"This report is calculated from the two uploaded raw consultation files using the supplied XLSM Formula-sheet logic. Composite Score is retained internally for ranking but intentionally hidden from the report table.",t:"s",s:{font:{italic:true,color:{rgb:"65758B"},sz:9},alignment:{wrapText:true,vertical:"top"}}};
  ws["!merges"].push({s:{r:noteRow,c:1},e:{r:noteRow,c:16}});
  ws["!rows"][noteRow]={hpt:34};

  const out=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(out,ws,"Daily Report");
  XLSX.writeFile(out,`Telemedicine_AAM_UPHC_Daily_Report_${APP.reportDate||"report"}.xlsx`);
}

function getExportScale(el,targetLongSide,maxArea){
  const rect=el.getBoundingClientRect();
  const w=Math.max(1,rect.width),h=Math.max(1,rect.height);
  // Adaptive scale: reach the target pixel size on the longest side, capped by
  // safe browser canvas limits.
  let scale=targetLongSide/Math.max(w,h);
  scale=Math.min(scale,Math.sqrt(maxArea/(w*h)));
  return Math.max(2,Math.min(scale,8));
}

async function renderReport4K(opts){
  opts=opts||{};
  const source=$("reportPaper");
  if(!source) throw new Error("Generate the report first.");
  const EXPORT_W=opts.width||1700;   // fixed layout width => identical PNG / PDF on every screen

  await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));

  // Render an off-screen clone so the on-screen report is never touched.
  // Charts are not part of the PNG/PDF (header + KPIs + table + summary only).
  const clone=source.cloneNode(true);
  clone.id="reportPaperExport";
  clone.classList.add("export-capture");
  clone.querySelectorAll(".analytics-grid, canvas").forEach(n=>n.remove());

  // Small icons above each KPI label (real elements render more reliably than pseudo-elements).
  const ICONS=["📈","🏥","👥","🩺","📊","⚠️","🗺️"];
  clone.querySelectorAll(".kpi-grid > .kpi").forEach((card,i)=>{
    const label=card.querySelector(":scope > span");
    if(!label) return;
    const ico=document.createElement("em");
    ico.className="exp-ico";
    ico.textContent=ICONS[i]||"•";
    label.insertBefore(ico,label.firstChild);
  });

  const wrap=document.createElement("div");
  wrap.className="export-wrapper";
  wrap.style.cssText=`position:fixed;left:-100000px;top:0;width:${EXPORT_W}px;max-width:none;background:#fff;overflow:visible;padding:0;margin:0;border:0;`;
  wrap.appendChild(clone);
  document.body.appendChild(wrap);

  try{
    await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
    const rect=clone.getBoundingClientRect();
    const width=Math.ceil(Math.max(clone.scrollWidth,rect.width,EXPORT_W));
    const height=Math.ceil(Math.max(clone.scrollHeight,rect.height));
    if(height<300) throw new Error("The report is empty - generate it first.");

    let scale=(opts.targetLongSide||3600)/Math.max(width,height);
    scale=Math.min(scale,Math.sqrt((opts.maxArea||90000000)/(width*height)));
    scale=Math.max(2,Math.min(scale,8));

    return await html2canvas(clone,{
      scale,
      backgroundColor:"#fff",
      useCORS:true,
      allowTaint:false,
      imageTimeout:0,
      logging:false,
      width,
      height,
      windowWidth:width,
      windowHeight:height,
      scrollX:0,
      scrollY:0
    });
  } finally {
    wrap.remove();
  }
}

async function downloadPdf(){
  if(!APP.generated)return;
  const {jsPDF}=window.jspdf;
  setStatus("Preparing compact high-resolution PDF…","info");
  const canvas=await renderReport4K({targetLongSide:6000,maxArea:60000000});
  // Custom page size = report size (297 mm long side, zero margin), so the
  // PDF has no extra white space and the report fills the whole page.
  const ratio=canvas.width/canvas.height;
  const longSide=297;
  const pw=ratio>=1?longSide:longSide*ratio;
  const ph=ratio>=1?longSide/ratio:longSide;
  const pdf=new jsPDF({orientation:pw>=ph?"landscape":"portrait",unit:"mm",format:[pw,ph],compress:true});
  pdf.addImage(canvas.toDataURL("image/jpeg",0.95),"JPEG",0,0,pw,ph,undefined,"FAST");
  pdf.save(`Telemedicine_AAM_UPHC_Daily_Report_${APP.reportDate||"report"}.pdf`);
  setStatus("Compact high-resolution PDF exported successfully.","success");
}
async function downloadPng(){
  if(!APP.generated)return;
  setStatus("Preparing 4K PNG…","info");
  const canvas=await renderReport4K({targetLongSide:9000,maxArea:90000000});
  const a=document.createElement("a");
  a.href=canvas.toDataURL("image/png");
  a.download=`Telemedicine_AAM_UPHC_Daily_Report_${APP.reportDate||"report"}_4K.png`;
  a.click();
  setStatus("4K PNG exported successfully.","success");
}

$("referredFile").addEventListener("change",e=>{
  const f=e.target.files[0];
  $("referredName").textContent=f?`${f.name} • ${(f.size/1024/1024).toFixed(2)} MB`:"No file selected";
  const d=f&&parseDateFromFilename(f.name);
  if(d)$("reportDate").value=d;
});
$("receiveFile").addEventListener("change",e=>{
  const f=e.target.files[0];
  $("receiveName").textContent=f?`${f.name} • ${(f.size/1024/1024).toFixed(2)} MB`:"No file selected";
  const d=f&&parseDateFromFilename(f.name);
  if(d&&!$("reportDate").value)$("reportDate").value=d;
});
$("generateBtn").addEventListener("click",generateReport);
$("resetBtn").addEventListener("click",resetApp);
$("excelBtn").addEventListener("click",downloadExcel);
$("pdfBtn").addEventListener("click",downloadPdf);
$("pngBtn").addEventListener("click",downloadPng);

setDefaultDate();
tickClock();
setInterval(tickClock,1000);
