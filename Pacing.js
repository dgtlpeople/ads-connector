function readDataRows_(sheetName) {
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return [];
  return sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function rebuildFactTable_() {
  ensureHeader_(SHEETS.FACT, HEADERS.RAW);
  clearDataKeepHeader_(SHEETS.FACT);

  const rows = readDataRows_(SHEETS.RAW_GOOGLE);
  if (rows.length) appendRows_(SHEETS.FACT, rows);
}

function readObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function rebuildPacing_() {
  ensureHeader_(SHEETS.PACE, HEADERS.PACE);
  clearDataKeepHeader_(SHEETS.PACE);

  const plan = readObjects_(SHEETS.PLAN);
  const fact = readObjects_(SHEETS.FACT);
  const today = new Date();
  const rows = [];

  plan.forEach(p => {
    if (String(p.platform).toLowerCase() !== 'google') return;

    const start = new Date(p.start_date);
    const end = new Date(p.end_date);
    const budget = toNumber_(p.budget);

    const daysTotal = Math.max(1, Math.floor((end - start) / 86400000) + 1);

    let daysElapsed = 0;
    if (today < start) {
      daysElapsed = 0;
    } else if (today > end) {
      daysElapsed = daysTotal;
    } else {
      daysElapsed = Math.floor((today - start) / 86400000) + 1;
    }

    const plannedToDate = budget * (daysElapsed / daysTotal);

    const actualToDate = fact
      .filter(f =>
        String(f.platform).toLowerCase() === 'google' &&
        String(f.campaign_id) === String(p.campaign_id)
      )
      .reduce((sum, r) => sum + toNumber_(r.spend), 0);

    const pacePct = plannedToDate > 0 ? actualToDate / plannedToDate : 0;
    const variance = actualToDate - plannedToDate;

    rows.push([
      p.platform,
      p.campaign_id,
      p.campaign_name,
      budget,
      p.start_date,
      p.end_date,
      daysTotal,
      daysElapsed,
      plannedToDate,
      actualToDate,
      pacePct,
      variance
    ]);
  });

  if (rows.length) {
    appendRows_(SHEETS.PACE, rows);
  }
}

function runGooglePipelineYesterday() {
  testGoogleAdsYesterday();
  rebuildFactTable_();
  rebuildPacing_();
  SpreadsheetApp.getUi().alert('Google pipeline finished.');
}