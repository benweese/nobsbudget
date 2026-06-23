/**
 * @OnlyCurrentDoc
 * Dashboard — plain-English "Am I okay?" spanning both open months when present.
 * Robust currency parsing; verdict escalates on multi-day negatives. Mortgage =
 * sum of mortgage-noted entries vs. full escrow. Savings/debt are manual.
 */

const DASH = { ESCROW_CELL: 'C15', SAVINGS_CELL: 'C17', NOVAKS_CELL: 'C18' };
const CUR = '$#,##0.00;[Red]-$#,##0.00';
const ESCROW_DEFAULT = 4120; // full monthly mortgage; fallback if the escrow cell is blank

function money(n) {
  const p = (Math.round(Math.abs(n) * 100) / 100).toFixed(2).split('.');
  p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-$' : '$') + p.join('.');
}

function ordinal(d) {
  const s = ['th', 'st', 'nd', 'rd'];
  const m = d % 100;
  return d + (s[(m - 20) % 10] || s[m] || s[0]);
}

/** Cleans a cell into a number. Strips $/commas; blanks or junk -> null (no data). */
function toNum(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  const n = parseFloat(String(v).replace(/[$,\s]/g, ''));
  return isNaN(n) ? null : n;
}

/** Lowest non-empty value + its 1-based day. Empty cells can't fake a $0 trough. */
function findTrough(arr) {
  let value = Infinity, day = 1;
  arr.forEach((v, i) => { if (v !== null && v < value) { value = v; day = i + 1; } });
  return value === Infinity ? { value: 0, day: 1 } : { value, day };
}

/** Reads the End-of-Day balance row into an array of numbers/nulls. */
function readChecking(sheet, days) {
  return sheet.getRange(LAYOUT.EOD_BALANCE_ROW, 1, 1, days).getValues()[0].map(toNum);
}

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tz = ss.getSpreadsheetTimeZone();
  const dash = ss.getSheetByName('Dashboard');
  if (!dash) { ui.alert('Create a tab named "Dashboard" first.'); return; }

  const now = new Date();
  const curName = Utilities.formatDate(now, tz, 'MMMM yyyy');
  const curLabel = Utilities.formatDate(now, tz, 'MMMM');
  const nextDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const nextName = Utilities.formatDate(nextDate, tz, 'MMMM yyyy');
  const nextLabel = Utilities.formatDate(nextDate, tz, 'MMMM');

  const curSheet = ss.getSheetByName(curName);
  if (!curSheet) { ui.alert(`Monthly sheet "${curName}" not found.`); return; }
  const nextSheet = ss.getSheetByName(nextName);

  const escrowTarget = parseFloat(dash.getRange(DASH.ESCROW_CELL).getValue()) || ESCROW_DEFAULT;
  const savings = parseFloat(dash.getRange(DASH.SAVINGS_CELL).getValue()) || 0;
  const owedNovaks = parseFloat(dash.getRange(DASH.NOVAKS_CELL).getValue()) || 0;

  const daysInCur = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const daysInNext = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
  const numRows = LAYOUT.TRANSACTION_END_ROW - LAYOUT.TRANSACTION_START_ROW + 1;

  let curTxVals, curTxNotes, curChecking;
  try {
    const r = curSheet.getRange(LAYOUT.TRANSACTION_START_ROW, 1, numRows, daysInCur);
    curTxVals = r.getValues();
    curTxNotes = r.getNotes();
    curChecking = readChecking(curSheet, daysInCur);
  } catch (e) {
    Logger.log(`Dashboard read error (current): ${e}`);
    ui.alert(`Could not read "${curName}": ${e}`);
    return;
  }

  // Combined checking runway across both open months (next only if it exists).
  let combined = curChecking.slice();
  if (nextSheet) {
    try { combined = combined.concat(readChecking(nextSheet, daysInNext)); }
    catch (e) { Logger.log(`Next month read failed, current only: ${e}`); }
  }
  const spanningBoth = combined.length > daysInCur;

  const dayLabel = (idx1) => idx1 <= daysInCur
    ? `${curLabel} ${ordinal(idx1)}`
    : `${nextLabel} ${ordinal(idx1 - daysInCur)}`;

  // Mortgage coverage (current month) + biggest expenses.
  let mortgageAllocated = 0; // set-asides add, raids subtract
  const expenses = [];
  for (let col = 0; col < daysInCur; col++) {
    for (let row = 0; row < numRows; row++) {
      const amt = toNum(curTxVals[row][col]);
      if (amt === null || amt === 0) continue;
      const note = String(curTxNotes[row][col]).toLowerCase();
      if (note.includes('mortgage')) {
        mortgageAllocated += -amt;
      } else if (amt < 0) {
        expenses.push({ amount: amt, day: col + 1, vendor: curTxNotes[row][col] || '(no note)' });
      }
    }
  }
  const shortfall = mortgageAllocated - escrowTarget; // negative = short
  expenses.sort((a, b) => a.amount - b.amount);
  const top3 = expenses.slice(0, 3);

  // Headline numbers.
  const scariest = findTrough(combined);
  const daysNegative = combined.filter(v => v !== null && v < 0).length;
  const populated = combined.filter(v => v !== null);
  const runwayEnd = populated.length ? populated[populated.length - 1] : 0;
  const daysRemaining = Math.max(1, (daysInCur - now.getDate() + 1) + (spanningBoth ? daysInNext : 0));
  const safePerDay = Math.max(0, runwayEnd) / daysRemaining;
  const finishLabel = spanningBoth ? nextLabel : curLabel;

  // Verdict — duration matters. 3+ days under zero is a real problem.
  const problems = [];
  if (daysNegative > 2) {
    problems.push({ sev: 3, msg: `You're underwater ${daysNegative} days — lowest is ${money(scariest.value)} on ${dayLabel(scariest.day)}. That's too long to float.` });
  } else if (scariest.value < 0) {
    problems.push({ sev: 2, msg: `You dip below zero for ${daysNegative} day${daysNegative === 1 ? '' : 's'} around ${dayLabel(scariest.day)} (down to ${money(scariest.value)}). Short dip, but watch it.` });
  }
  if (runwayEnd < 0) {
    problems.push({ sev: 3, msg: `You finish at ${money(runwayEnd)} — in the red.` });
  }
  if (shortfall < 0) {
    problems.push({ sev: 2, msg: `The mortgage is ${money(Math.abs(shortfall))} short — ${money(mortgageAllocated)} of ${money(escrowTarget)} set aside.` });
  }

  const maxSev = problems.reduce((m, p) => Math.max(m, p.sev), 0);
  let verdict, bg, fg;
  if (maxSev === 3) { verdict = '🔴 You\u2019re screwed'; bg = '#f4cccc'; fg = '#990000'; }
  else if (maxSev === 2) { verdict = '🟡 It\u2019s tight'; bg = '#fff2cc'; fg = '#7f6000'; }
  else { verdict = '🟢 You\u2019re okay'; bg = '#d9ead3'; fg = '#38761d'; }
  const why = problems.length
    ? problems.map(p => p.msg).join('  ')
    : `Lowest you dip is ${money(scariest.value)} on ${dayLabel(scariest.day)}, the mortgage is covered, and you can spend ${money(safePerDay)} a day. Relax.`;

  // Sentences.
  const safeSentence = `Spend up to this each day on non-bills and still finish ${finishLabel} above zero.`;
  let scarySentence;
  if (daysNegative > 2) {
    scarySentence = `You\u2019re below zero ${daysNegative} days — too long. You\u2019ll need a cushion or to move money around.`;
  } else if (daysNegative > 0) {
    scarySentence = `Below zero ${daysNegative} day${daysNegative === 1 ? '' : 's'} — brief, but cutting it close.`;
  } else {
    scarySentence = 'Even at its lowest, checking stays above zero. Breathing room the whole stretch.';
  }
  const mortgageStatus = shortfall >= 0 ? '✅ Yes' : `⚠️ No — ${money(Math.abs(shortfall))} short`;
  const mortgageSentence = shortfall >= 0
    ? `The full ${money(escrowTarget)} is set aside and ready for the payment.`
    : `Set aside ${money(mortgageAllocated)} of ${money(escrowTarget)} — ${money(Math.abs(shortfall))} to go before month-end.`;

  const title = spanningBoth ? `Am I okay?   —   ${curLabel} + ${nextLabel}` : `Am I okay?   —   ${curLabel}`;

  const rows = [
    [title, ''],                                       // 2
    ['', ''],                                           // 3
    [verdict, ''],                                      // 4
    [why, ''],                                          // 5
    ['', ''],                                            // 6
    ['Safe to spend each day', safePerDay],             // 7
    [safeSentence, ''],                                 // 8
    ['', ''],                                            // 9
    [`Scariest day (${dayLabel(scariest.day)})`, scariest.value], // 10
    [scarySentence, ''],                                // 11
    ['', ''],                                            // 12
    ['Is the mortgage covered?', mortgageStatus],       // 13
    [mortgageSentence, ''],                             // 14
    ['Escrow due at month-end (edit if it changes)', escrowTarget], // 15
    ['', ''],                                            // 16
    ['Savings', savings],                               // 17
    ['Owed to the Novaks', owedNovaks],                 // 18
    ['', ''],                                            // 19
    ['Biggest expenses this month', ''],                // 20
    [top3[0] ? `${top3[0].vendor}  (${ordinal(top3[0].day)})` : '—', top3[0] ? top3[0].amount : ''], // 21
    [top3[1] ? `${top3[1].vendor}  (${ordinal(top3[1].day)})` : '', top3[1] ? top3[1].amount : ''],  // 22
    [top3[2] ? `${top3[2].vendor}  (${ordinal(top3[2].day)})` : '', top3[2] ? top3[2].amount : ''],  // 23
  ];

  dash.getRange('A1:H40').clear();
  dash.getRange(2, 2, rows.length, 2).setValues(rows);

  // Formatting.
  dash.setColumnWidth(1, 30).setColumnWidth(2, 300).setColumnWidth(3, 160);
  ['B2:C2', 'B4:C4', 'B5:C5', 'B8:C8', 'B11:C11', 'B14:C14', 'B20:C20'].forEach(rng => dash.getRange(rng).merge());

  dash.getRange('B2').setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold').setFontSize(15).setHorizontalAlignment('center');
  dash.setRowHeight(2, 34).setRowHeight(4, 46);
  dash.getRange('B4').setBackground(bg).setFontColor(fg).setFontWeight('bold').setFontSize(20).setHorizontalAlignment('center');
  dash.getRange('B5').setFontColor('#666666').setFontStyle('italic').setWrap(true).setFontSize(11);

  ['B7:C7', 'B10:C10', 'B13:C13'].forEach(rng =>
    dash.getRange(rng).setBackground('#f8f9fa').setFontWeight('bold').setFontSize(12)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID));

  ['B8', 'B11', 'B14'].forEach(c =>
    dash.getRange(c).setFontColor('#888888').setFontStyle('italic').setWrap(true).setFontSize(10));

  dash.getRange('C13').setFontColor(shortfall >= 0 ? '#38761d' : '#990000').setHorizontalAlignment('right').setFontWeight('bold');

  ['C7', 'C10', 'C15', 'C17', 'C18', 'C21:C23'].forEach(rng =>
    dash.getRange(rng).setNumberFormat(CUR).setHorizontalAlignment('right'));

  dash.getRange('B20').setBackground('#f1f3f4').setFontWeight('bold').setFontColor('#5f6368').setFontSize(11);
  dash.getRange('B21:C23').setBackground('#f8f9fa');

  ['C15', 'C17', 'C18'].forEach(c =>
    dash.getRange(c).setBackground('#e8f0fe').setBorder(true, true, true, true, true, true, '#a4c2f4', SpreadsheetApp.BorderStyle.SOLID));
  dash.getRange('C15').setNote('Edit me — full monthly mortgage/escrow due');
  dash.getRange('C17').setNote('Edit me — current savings balance');
  dash.getRange('C18').setNote('Edit me — total owed to the Novaks');

  SpreadsheetApp.flush();
  ui.alert(`Dashboard updated.\n${verdict}`);
}