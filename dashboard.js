/**
 * @OnlyCurrentDoc
 * Dashboard engine — the "Am I Screwed?" report.
 * Models checking balance (Line A) vs. true liquidity including the raidable
 * mortgage buffer (Line B), warns if the escrow envelope won't be whole by
 * month-end, and tracks manually-entered savings / debt. All failures surface.
 */

const DASH = { ESCROW_CELL: 'C17', SAVINGS_CELL: 'C21', NOVAKS_CELL: 'C22' };
const CUR = '$#,##0.00;[Red]-$#,##0.00';
const ESCROW_DEFAULT = 4120; // fallback if the escrow cell is blank — never silently disable the warning

/** Formats a number as signed USD. e.g. -1910 -> "-$1,910.00". */
function money(n) {
  const p = (Math.round(Math.abs(n) * 100) / 100).toFixed(2).split('.');
  p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-$' : '$') + p.join('.');
}

/** Ordinal form of a day. e.g. 14 -> "14th". */
function ordinal(d) {
  const s = ['th', 'st', 'nd', 'rd'];
  const m = d % 100;
  return d + (s[(m - 20) % 10] || s[m] || s[0]);
}

/** Lowest value in a daily array + the 1-based day it occurs. */
function findTrough(arr) {
  let value = Infinity, day = 1;
  arr.forEach((v, i) => { if (v < value) { value = v; day = i + 1; } });
  return { value, day };
}
// eslint-disable-next-line no-unused-vars
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const dash = ss.getSheetByName('Dashboard');
  if (!dash) { ui.alert('Create a tab named "Dashboard" first.'); return; }

  const now = new Date();
  const monthName = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'MMMM yyyy');
  const month = ss.getSheetByName(monthName);
  if (!month) { ui.alert(`Monthly sheet "${monthName}" not found.`); return; }

  // Preserve manual inputs BEFORE rebuilding. Escrow falls back to default if blank.
  const escrowTarget = parseFloat(dash.getRange(DASH.ESCROW_CELL).getValue()) || ESCROW_DEFAULT;
  const savings = parseFloat(dash.getRange(DASH.SAVINGS_CELL).getValue()) || 0;
  const owedNovaks = parseFloat(dash.getRange(DASH.NOVAKS_CELL).getValue()) || 0;

  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const todayCol = Math.min(now.getDate(), daysInMonth);
  const daysRemaining = Math.max(1, daysInMonth - now.getDate() + 1);
  const numRows = LAYOUT.TRANSACTION_END_ROW - LAYOUT.TRANSACTION_START_ROW + 1;
  const lastIdx = daysInMonth - 1;

  let txVals, txNotes, eod;
  try {
    const r = month.getRange(LAYOUT.TRANSACTION_START_ROW, 1, numRows, daysInMonth);
    txVals = r.getValues();
    txNotes = r.getNotes();
    eod = month.getRange(LAYOUT.EOD_BALANCE_ROW, 1, 1, daysInMonth).getValues()[0];
  } catch (e) {
    Logger.log(`Dashboard read error: ${e}`);
    ui.alert(`Could not read "${monthName}": ${e}`);
    return;
  }

  // Line A = checking balance as the sheet shows it (mortgage already subtracted).
  const lineA = eod.map(v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; });

  // One in-memory pass: signed mortgage movements, locked last-day payment, non-mortgage expenses.
  const mortgageNet = new Array(daysInMonth).fill(0); // neg = set aside, pos = raid
  let lockedPayment = 0;
  const expenses = [];
  for (let col = 0; col < daysInMonth; col++) {
    for (let row = 0; row < numRows; row++) {
      const amt = parseFloat(txVals[row][col]);
      if (isNaN(amt) || amt === 0) continue;
      const note = String(txNotes[row][col]).toLowerCase();
      if (note.includes('mortgage')) {
        if (col === lastIdx && amt < 0) lockedPayment += Math.abs(amt); // locked, not raidable
        else mortgageNet[col] += amt;
      } else if (amt < 0) {
        expenses.push({ amount: amt, day: col + 1, vendor: txNotes[row][col] || '(no note)' });
      }
    }
  }

  // Line B = checking + cumulative raidable buffer parked to date (floored at 0).
  let cumulative = 0;
  const lineB = lineA.map((bal, i) => { cumulative += -mortgageNet[i]; return bal + Math.max(0, cumulative); });

  // Envelope at month-end = locked last-day payment + everything set aside minus raided.
  const envelopeAtEOM = lockedPayment + cumulative;
  const shortfall = envelopeAtEOM - escrowTarget; // negative = short

  let raidableToDate = 0;
  for (let i = 0; i < todayCol; i++) raidableToDate += -mortgageNet[i];
  raidableToDate = Math.max(0, raidableToDate);

  const tA = findTrough(lineA);
  const tB = findTrough(lineB);
  const monthEnd = lineA[lastIdx];
  const safePerDay = Math.max(0, monthEnd) / daysRemaining;
  expenses.sort((a, b) => a.amount - b.amount); // most negative first
  const top3 = expenses.slice(0, 3);

  // Problems list — every failure surfaces; highest severity drives the banner.
  const problems = [];
  if (tB.value < 0) {
    problems.push({ sev: 3, msg: `Liquidity underwater by ${money(tB.value)} on the ${ordinal(tB.day)} even after raiding the buffer.` });
  } else if (tA.value < 0) {
    problems.push({ sev: 2, msg: `Checking dips to ${money(tA.value)} on the ${ordinal(tA.day)} — floating on the mortgage buffer.` });
  }
  if (monthEnd < 0) {
    problems.push({ sev: 3, msg: `Checking ends the month at ${money(monthEnd)}.` });
  }
  if (shortfall < 0) {
    problems.push({ sev: 3, msg: `Escrow envelope projected at ${money(envelopeAtEOM)} — ${money(Math.abs(shortfall))} short of the ${money(escrowTarget)} due. Repay what you raided before month-end.` });
  }

  const maxSev = problems.reduce((m, p) => Math.max(m, p.sev), 0);
  let verdict, bg, fg;
  if (maxSev === 3) { verdict = '🔴 SCREWED'; bg = '#f4cccc'; fg = '#990000'; }
  else if (maxSev === 2) { verdict = '🟡 TIGHT'; bg = '#fff2cc'; fg = '#7f6000'; }
  else { verdict = '🟢 FINE'; bg = '#d9ead3'; fg = '#38761d'; }
  const why = problems.length
    ? problems.map(p => p.msg).join('  •  ')
    : `Lowest point ${money(tA.value)} on the ${ordinal(tA.day)}, buffer whole, escrow covered. Spend within the daily limit.`;

  // --- LAYOUT (col B label, col C value), starting at row 2 ---
  const rows = [
    [`Am I Screwed?   —   ${monthName}`, ''],            // 2
    ['', ''],                                             // 3
    [verdict, ''],                                        // 4
    [why, ''],                                            // 5
    ['', ''],                                             // 6
    ['CASH FLOW DANGER', ''],                             // 7
    [`Lowest checking  (${ordinal(tA.day)})`, tA.value],  // 8
    [`Lowest w/ float  (${ordinal(tB.day)})`, tB.value],  // 9
    ['Projected month-end', monthEnd],                    // 10
    ['Safe to spend / day left', safePerDay],             // 11
    ['', ''],                                             // 12
    ['MORTGAGE ENVELOPE', ''],                            // 13
    ['Set aside, raidable (to date)', raidableToDate],    // 14
    ['Locked payment (last day)', lockedPayment],         // 15
    ['Projected at month-end', envelopeAtEOM],            // 16
    ['Required (escrow) — EDITABLE', escrowTarget],       // 17
    ['Shortfall / surplus', shortfall],                   // 18
    ['', ''],                                             // 19
    ['MANUAL — edit the blue cells', ''],                 // 20
    ['Current savings', savings],                         // 21
    ['Owed to the Novaks', owedNovaks],                   // 22
    ['Net (savings − Novaks)', savings - owedNovaks],     // 23
    ['', ''],                                             // 24
    ['BIGGEST EXPENSES', ''],                             // 25
    [top3[0] ? `${top3[0].vendor}  (${ordinal(top3[0].day)})` : '—', top3[0] ? top3[0].amount : ''], // 26
    [top3[1] ? `${top3[1].vendor}  (${ordinal(top3[1].day)})` : '', top3[1] ? top3[1].amount : ''],  // 27
    [top3[2] ? `${top3[2].vendor}  (${ordinal(top3[2].day)})` : '', top3[2] ? top3[2].amount : ''],  // 28
  ];

  dash.getRange('A1:H40').clear();
  dash.getRange(2, 2, rows.length, 2).setValues(rows);

  // --- FORMATTING ---
  dash.setColumnWidth(1, 30).setColumnWidth(2, 250).setColumnWidth(3, 200);
  ['B2:C2', 'B4:C4', 'B5:C5', 'B7:C7', 'B13:C13', 'B20:C20', 'B25:C25'].forEach(rng => dash.getRange(rng).merge());

  dash.getRange('B2').setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  dash.setRowHeight(2, 36).setRowHeight(4, 46);
  dash.getRange('B4').setBackground(bg).setFontColor(fg).setFontWeight('bold').setFontSize(20).setHorizontalAlignment('center');
  dash.getRange('B5').setFontColor('#666666').setFontStyle('italic').setWrap(true).setFontSize(10);

  ['B7', 'B13', 'B20', 'B25'].forEach(c =>
    dash.getRange(c).setBackground('#f1f3f4').setFontWeight('bold').setFontColor('#5f6368').setFontSize(11));

  // Currency cells: right-aligned, bold, red negatives.
  ['C8:C11', 'C14:C18', 'C21:C23', 'C26:C28'].forEach(rng =>
    dash.getRange(rng).setNumberFormat(CUR).setHorizontalAlignment('right').setFontWeight('bold'));

  // Card backgrounds + light borders.
  ['B8:C11', 'B14:C18', 'B26:C28'].forEach(rng =>
    dash.getRange(rng).setBackground('#f8f9fa').setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID));

  // Shortfall punchline goes red when the envelope won't be whole.
  if (shortfall < 0) {
    dash.getRange('B18:C18').setBackground('#fce5e0').setFontColor('#990000').setFontWeight('bold');
  }

  // Editable blue cells (applied last so they win over the gray card): escrow, savings, Novaks.
  ['C17', 'C21', 'C22'].forEach(c =>
    dash.getRange(c).setBackground('#e8f0fe').setBorder(true, true, true, true, true, true, '#a4c2f4', SpreadsheetApp.BorderStyle.SOLID));
  dash.getRange('C17').setNote('Edit me — full escrow/mortgage amount due at month-end');
  dash.getRange('C21').setNote('Edit me — current savings balance');
  dash.getRange('C22').setNote('Edit me — total owed to the Novaks');
  dash.getRange('B23:C23').setBackground('#f8f9fa');

  SpreadsheetApp.flush();
  ui.alert(`✅ Dashboard updated for ${monthName}.\nVerdict: ${verdict}`);
}