const HEADERS = [
  "date", "status", "hours", "wage", "ot",
  "rent", "fuel", "other", "note"
];

// ဒီ function ကို ပထမဆုံး တစ်ကြိမ် Run ပေးပါ။
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Google Sheets ထဲက Extensions → Apps Script မှ ဖွင့်ပါ။");
  }

  PropertiesService.getScriptProperties()
    .setProperty("SPREADSHEET_ID", ss.getId());

  let sheet = ss.getSheetByName("Records");
  if (!sheet) sheet = ss.insertSheet("Records");

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("I:I").setNumberFormat("@");
  sheet.getRange("C:H").setNumberFormat("0.##");
  sheet.autoResizeColumns(1, HEADERS.length);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("နေ့စဉ် အလုပ်နှင့်ငွေစာရင်း")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("SPREADSHEET_ID");

  if (!id) throw new Error("setup function ကို အရင် Run ပေးပါ။");

  const sheet = SpreadsheetApp.openById(id).getSheetByName("Records");
  if (!sheet) throw new Error("Records စာမျက်နှာ မတွေ့ပါ။");

  return sheet;
}

function validDate_(value) {
  const s = String(value || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    throw new Error("ရက်စွဲ မမှန်ပါ။");
  }

  const d = new Date(s + "T00:00:00Z");
  if (isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== s) {
    throw new Error("ရက်စွဲ မမှန်ပါ။");
  }

  return s;
}

function amount_(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    throw new Error("ငွေပမာဏနှင့် နာရီကို သုည သို့မဟုတ် အပေါင်းကိန်း ဖြည့်ပါ။");
  }
  return n;
}

function getRecords(month) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(month))) {
    throw new Error("လရွေးချယ်မှု မမှန်ပါ။");
  }

  const sheet = getSheet_();
  if (sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, HEADERS.length
  ).getValues();

  return values
    .filter(row => String(row[0]).startsWith(month + "-"))
    .map(row => {
      const result = {};
      HEADERS.forEach((key, i) => {
        result[key] = i >= 2 && i <= 7
          ? Number(row[i]) || 0
          : String(row[i] || "");
      });
      return result;
    })
    .sort((a, b) => b.date.localeCompare(a.date));
}

function saveRecord(data) {
  const date = validDate_(data.date);

  if (!["work", "off"].includes(data.status)) {
    throw new Error("အလုပ်ဆင်းမှု ရွေးပေးပါ။");
  }

  const working = data.status === "work";
  const note = String(data.note || "").slice(0, 1000);

  // Spreadsheet formula အဖြစ် မဖတ်စေရန် ကာကွယ်ထားသည်။
  const safeNote = /^[=+\-@]/.test(note) ? "'" + note : note;

  const row = [
    date,
    data.status,
    working ? amount_(data.hours) : 0,
    working ? amount_(data.wage) : 0,
    amount_(data.ot),
    amount_(data.rent),
    amount_(data.fuel),
    amount_(data.other),
    safeNote
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_();
    const last = sheet.getLastRow();

    const dates = last > 1
      ? sheet.getRange(2, 1, last - 1, 1).getValues()
      : [];

    const index = dates.findIndex(r => String(r[0]) === date);

    if (index >= 0) {
      sheet.getRange(index + 2, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return { message: "စာရင်း သိမ်းပြီးပါပြီ။" };
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(date) {
  date = validDate_(date);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_();
    const last = sheet.getLastRow();

    if (last > 1) {
      const dates = sheet.getRange(2, 1, last - 1, 1).getValues();
      const index = dates.findIndex(r => String(r[0]) === date);
      if (index >= 0) sheet.deleteRow(index + 2);
    }

    return { message: "စာရင်း ဖျက်ပြီးပါပြီ။" };
  } finally {
    lock.releaseLock();
  }
}
