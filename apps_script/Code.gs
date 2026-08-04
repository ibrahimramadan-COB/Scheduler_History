// ====================================================================
// WebPT Scheduler History Automation — Apps Script
// ====================================================================
// This script is bound to the SAME Google Sheet the GitHub Actions
// pipeline writes into (Master Data / Names / Groups / Emails tabs).
//
// SECURITY NOTE: the GitHub token is read from Script Properties, NOT
// hardcoded here. One-time setup:
//   Project Settings (gear icon) -> Script Properties -> Add property
//     GITHUB_TOKEN = ghp_xxxxxxxxxxxx   (a repo-scoped PAT with
//                                        "Actions: read and write")
// ====================================================================

const GITHUB_USERNAME = "ibrahimramadan-COB";
const GITHUB_REPO = "Scheduler_History";
const GITHUB_WORKFLOW_FILE = "scrape.yml";

const MASTER_TAB = "Master Data";
const NAMES_TAB = "Names";
const GROUPS_TAB = "Groups";
const EMAILS_TAB = "Emails";
const REPORT_RUNS_TAB = "Report Runs";
const CLINICS_TAB = "Clinics";

// Same dedupe key used throughout the pipeline (scraper -> Sheets -> here),
// applied as a defensive final pass on whatever rows a report filters to.
const DEDUPE_KEY_FIELDS = ["APPOINTMENT DATE", "PATIENT", "CREATED TIMESTAMP", "EVENT ACTION", "DATA_SOURCE_TAB"];

const EMAIL_SENDER_NAME = "WebPT Scheduler History Automation";

const EVENT_TYPES = ["Cancelled", "Checked In", "Checked Out", "Created", "Deleted", "Edited", "No Show"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('WebPT Scheduler History Automation')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====================================================================
// GITHUB SYNC — triggers the scraper workflow with a chosen date range
// ====================================================================
function triggerGitHubSync(startDate, endDate, clinics) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    return "❌ GITHUB_TOKEN is not set in Script Properties. Go to Project Settings -> Script Properties and add it.";
  }

  const url = `https://api.github.com/repos/${GITHUB_USERNAME}/${GITHUB_REPO}/actions/workflows/${GITHUB_WORKFLOW_FILE}/dispatches`;
  const inputs = {
    "start_date": startDate,
    "end_date": endDate
  };
  if (clinics && clinics.length > 0) {
    inputs["clinics"] = clinics.join("|");
  }
  const payload = {
    "ref": "main",
    "inputs": inputs
  };
  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + token,
      "Accept": "application/vnd.github.v3+json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      return `✅ GitHub scrape started for ${startDate} → ${endDate}. This can take a while (up to a few hours for all clinics) — check the Actions tab, or come back later and run a report once it's done.`;
    }
    return "❌ GitHub Error: " + response.getContentText();
  } catch (e) {
    return "❌ Apps Script Error: " + e.message;
  }
}

// ====================================================================
// DROPDOWN DATA — Group / Group 2 / Clinic options, read fresh on page load
// ====================================================================
function getFilterOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Group / Group 2
  const groupsSheet = ss.getSheetByName(GROUPS_TAB);
  const group2ByGroup = {};
  const groupsSet = new Set();
  if (groupsSheet) {
    const vals = groupsSheet.getDataRange().getValues();
    const headers = vals[0].map(h => String(h).trim());
    const gIdx = headers.indexOf("Group");
    const g2Idx = headers.indexOf("Group 2");
    for (let i = 1; i < vals.length; i++) {
      const g = String(vals[i][gIdx] || "").trim();
      const g2 = String(vals[i][g2Idx] || "").trim();
      if (!g) continue;
      groupsSet.add(g);
      if (!group2ByGroup[g]) group2ByGroup[g] = [];
      if (g2 && group2ByGroup[g].indexOf(g2) === -1) group2ByGroup[g].push(g2);
    }
  }

  // Clinics — kept fresh automatically by sheets_writer.py on every scrape run
  const clinicsSheet = ss.getSheetByName(CLINICS_TAB);
  let clinics = [];
  if (clinicsSheet && clinicsSheet.getLastRow() > 1) {
    clinics = clinicsSheet.getRange(2, 1, clinicsSheet.getLastRow() - 1, 1)
      .getValues().map(r => String(r[0]).trim()).filter(c => c);
  }

  return {
    groups: Array.from(groupsSet).sort(),
    group2ByGroup: group2ByGroup,
    eventTypes: EVENT_TYPES,
    clinics: clinics
  };
}

// ====================================================================
// MAIN REPORT ENGINE
// ====================================================================
// config = {
//   mode: 'created' | 'vfd' | 'custom',
//   eventTypes: [...] (custom mode only, empty = all types),
//   group: '' (custom mode only),
//   group2: '' (custom mode only),
//   clinics: [...] (any mode, empty = all clinics),
//   startDate: 'YYYY-MM-DD',
//   endDate: 'YYYY-MM-DD',
//   sendEmail: true/false
// }
function runSchedulerReport(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_TAB);
  if (!masterSheet) return "❌ 'Master Data' tab not found — run a GitHub sync first.";

  // Resolve the filter set based on mode
  let eventTypeFilter = [];
  let groupFilter = "";
  let group2Filter = "";
  let reportLabel = "";

  if (config.mode === 'created') {
    eventTypeFilter = ["Created"];
    groupFilter = "CC";
    reportLabel = "Created Appointments — CC Department";
  } else if (config.mode === 'vfd') {
    eventTypeFilter = [];              // all event types
    group2Filter = "VFD";
    reportLabel = "VFD Activity — All Event Types";
  } else {
    eventTypeFilter = config.eventTypes || [];
    groupFilter = config.group || "";
    group2Filter = config.group2 || "";
    reportLabel = "Custom Report";
  }

  const sDate = new Date(config.startDate); sDate.setHours(0, 0, 0, 0);
  const eDate = new Date(config.endDate); eDate.setHours(23, 59, 59, 999);

  const vals = masterSheet.getDataRange().getValues();
  if (vals.length < 2) return "⚠️ 'Master Data' is empty — run a GitHub sync first.";
  const headers = vals[0].map(h => String(h).trim());
  const rows = vals.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });

  const clinicFilter = (config.clinics || []).map(c => c.trim().toLowerCase()).filter(c => c);

  let filtered = rows.filter(r => {
    const d = parseDate(r["APPOINTMENT DATE"]);
    if (!d || d < sDate || d > eDate) return false;
    if (eventTypeFilter.length > 0 && eventTypeFilter.indexOf(String(r["EVENT ACTION"]).trim()) === -1) return false;
    if (groupFilter && String(r["GROUP"]).trim() !== groupFilter) return false;
    if (group2Filter && String(r["GROUP 2"]).trim() !== group2Filter) return false;
    if (clinicFilter.length > 0 && clinicFilter.indexOf(String(r["CLINIC"]).trim().toLowerCase()) === -1) return false;
    return true;
  });

  // Defensive dedupe on the filtered set itself — belt-and-suspenders on top
  // of the dedupe already done when writing to Master Data, in case a range
  // spans rows added across more than one scrape run.
  filtered = dedupeRows(filtered);

  if (filtered.length === 0) {
    logReportRun(config.mode, reportLabel, config.startDate, config.endDate, 0);
    return `⚠️ No rows matched this filter for ${config.startDate} → ${config.endDate}.`;
  }

  const pivot = buildNestedPivot(filtered);
  const html = buildHtmlEmail(pivot, reportLabel, config.startDate, config.endDate);

  let emailStatus = "";
  if (config.sendEmail !== false) {
    const emailSheet = ss.getSheetByName(EMAILS_TAB);
    if (!emailSheet) {
      emailStatus = " (⚠️ 'Emails' tab not found — skipped sending.)";
    } else {
      const emailList = emailSheet.getRange("A1:A" + emailSheet.getLastRow())
        .getValues().map(r => String(r[0]).trim()).filter(e => e.indexOf("@") !== -1).join(",");
      if (emailList) {
        const subject = `${reportLabel} (${config.startDate} to ${config.endDate})`;
        sendThreadedEmail(subject, html, emailList);
        emailStatus = ` Emailed to: ${emailList}.`;
      } else {
        emailStatus = " (⚠️ No valid addresses found in 'Emails' tab — skipped sending.)";
      }
    }
  }

  logReportRun(config.mode, reportLabel, config.startDate, config.endDate, filtered.length);

  return `✅ ${reportLabel}: ${filtered.length} row(s) for ${config.startDate} → ${config.endDate}.${emailStatus}`;
}

// ====================================================================
// DEDUPE — same key as the Python pipeline, applied defensively here too
// ====================================================================
function dedupeRows(rows) {
  const seen = new Set();
  return rows.filter(r => {
    const key = DEDUPE_KEY_FIELDS.map(f => String(r[f])).join("‖");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// ====================================================================
// PIVOT — grouped by DEPARTMENT (Group — Group 2), then by CREATOR USER.
// Values: # of Patients = COUNT DISTINCT of PATIENT within that row;
//         # of Appointments = COUNT (not distinct) of PATIENT, i.e. the
//         number of matching event rows. Each department gets a "Total"
//         subtotal row, and there's one GRAND TOTAL row at the very end.
// ====================================================================
function buildNestedPivot(data) {
  const deptMap = {}; // deptKey -> { creators: {name: {patients:Set, count}}, allPatients:Set, totalCount }

  data.forEach(r => {
    const group = String(r["GROUP"] || "Other").trim() || "Other";
    const group2 = String(r["GROUP 2"] || "Other").trim() || "Other";
    const deptKey = group + " — " + group2;
    const creator = String(r["CREATOR USER"] || "Unknown").trim() || "Unknown";
    const patient = String(r["PATIENT"] || "").trim();

    if (!deptMap[deptKey]) deptMap[deptKey] = { creators: {}, allPatients: new Set(), totalCount: 0 };
    const dept = deptMap[deptKey];
    if (!dept.creators[creator]) dept.creators[creator] = { patients: new Set(), count: 0 };

    dept.creators[creator].patients.add(patient);
    dept.creators[creator].count++;
    dept.allPatients.add(patient);
    dept.totalCount++;
  });

  const outRows = []; // {type:'header'|'data'|'subtotal'|'grandtotal', label, patients, appointments}
  const grandAllPatients = new Set();
  let grandTotalCount = 0;

  Object.keys(deptMap).sort().forEach(deptKey => {
    const dept = deptMap[deptKey];
    outRows.push({ type: 'header', label: deptKey });

    Object.keys(dept.creators).sort().forEach(creator => {
      const c = dept.creators[creator];
      outRows.push({ type: 'data', label: creator, patients: c.patients.size, appointments: c.count });
    });

    outRows.push({ type: 'subtotal', label: 'Total', patients: dept.allPatients.size, appointments: dept.totalCount });
    dept.allPatients.forEach(p => grandAllPatients.add(p));
    grandTotalCount += dept.totalCount;
  });

  outRows.push({ type: 'grandtotal', label: 'GRAND TOTAL', patients: grandAllPatients.size, appointments: grandTotalCount });
  return outRows;
}

// ====================================================================
// REPORT RUN LOG
// ====================================================================
function logReportRun(mode, label, startDate, endDate, rowCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REPORT_RUNS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(REPORT_RUNS_TAB);
    sheet.appendRow(["Timestamp", "Mode", "Report", "Start Date", "End Date", "Row Count"]);
  }
  sheet.appendRow([new Date(), mode, label, startDate, endDate, rowCount]);
}

// ====================================================================
// EMAIL
// ====================================================================
function sendThreadedEmail(subject, htmlBody, emailList) {
  const threads = GmailApp.search('subject:"' + subject + '"');
  if (threads.length > 0) {
    threads[0].reply("", { htmlBody: htmlBody, name: EMAIL_SENDER_NAME, bcc: emailList });
  } else {
    GmailApp.sendEmail(emailList, subject, "", { htmlBody: htmlBody, name: EMAIL_SENDER_NAME });
  }
}

function buildHtmlEmail(pivotRows, title, startDate, endDate) {
  let html = `<div style="font-family: Arial, sans-serif; color: #333;">`;
  html += `<h2 style="color:#1F4E79; margin-bottom:5px;">${title}</h2>`;
  html += `<p style="color:#666; margin-top:0;">${startDate} → ${endDate}</p>`;
  html += `<table border="1" cellpadding="8" style="border-collapse:collapse; font-size:13px; text-align:center; max-width:900px;">`;
  html += `<tr>
             <th style="background-color:#203764; color:white; text-align:left;">CREATOR USER</th>
             <th style="background-color:#203764; color:white;"># of Patients</th>
             <th style="background-color:#203764; color:white;"># of Appointments</th>
           </tr>`;

  pivotRows.forEach(row => {
    if (row.type === 'header') {
      html += `<tr><td colspan="3" style="background-color:#1F4E79; color:white; font-weight:bold; text-align:left;">${row.label}</td></tr>`;
    } else if (row.type === 'data') {
      html += `<tr>
                 <td style="text-align:left; padding-left:25px;">${row.label}</td>
                 <td>${row.patients}</td>
                 <td>${row.appointments}</td>
               </tr>`;
    } else if (row.type === 'subtotal') {
      html += `<tr style="background-color:#DDEBF7; font-weight:bold;">
                 <td style="text-align:left; padding-left:25px;">${row.label}</td>
                 <td>${row.patients}</td>
                 <td>${row.appointments}</td>
               </tr>`;
    } else if (row.type === 'grandtotal') {
      html += `<tr style="background-color:#203764; color:white; font-weight:bold;">
                 <td style="text-align:left;">${row.label}</td>
                 <td>${row.patients}</td>
                 <td>${row.appointments}</td>
               </tr>`;
    }
  });

  html += `</table></div>`;
  return html;
}

// ====================================================================
// HELPERS
// ====================================================================
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}
