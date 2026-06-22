/**
 * @OnlyCurrentDoc
 * Dashboard engine — the plain-English "Am I okay this month?" report.
 * Tracks checking balance, daily safe-to-spend, the scariest day, whether the
 * escrow/mortgage is covered, and manual savings/debt. Human-readable by design.
 */

const DASH = { ESCROW_CELL: 'C15', SAVINGS_CELL: 'C17', NOVAKS_CELL: 'C18' };
const CUR = '$#,##0.00;[Red]-$#,##0.00';
const ESCROW_DEFAULT = 4120; // fallback if the escrow cell is blank — never silently disable the check

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

  // Daily checking balance, exactly as the month sheet shows it.
  const checking = eod.map(v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; });

  // One pass: mortgage movements (signed), locked last-day payment, non-mortgage expenses.
  const mortgageNet = new Array(daysInMonth).fill(0); // neg = set aside, pos = raid
  let lockedPayment = 0;
  const expenses = [];
  for (let col = 0; col < daysInMonth; col++) {
    for (let row = 0; row < numRows; row++) {
      const amt = parseFloat(txVals[row][col]);
      if (isNaN(amt) || amt === 0) continue;
      const note = String(txNotes[row][col]).toLowerCase();
      if (note.includes('mortgage')) {
        if (col === lastIdx && amt < 0) lockedPayment += Math.abs(amt);
        else mortgageNet[col] += amt;
      } else if (amt < 0) {
        expenses.push({ amount: amt, day: col + 1, vendor: txNotes[row][col] || '(no note)' });
      }
    }
  }

  // True cash available = checking + mortgage money you're allowed to borrow back.
  let cumulative = 0;
  const withBuffer = checking.map((bal, i) => { cumulative += -mortgageNet[i]; return bal + Math.max(0, cumulative); });
  const envelopeAtEOM = lockedPayment + cumulative;
  const shortfall = envelopeAtEOM - escrowTarget; // negative = short

  const scariest = findTrough(checking);   // lowest checking balance + day
  const worstCase = findTrough(withBuffer); // lowest even after borrowing the buffer
  const monthEnd = checking[lastIdx];
  const safePerDay = Math.max(0, monthEnd) / daysRemaining;
  expenses.sort((a, b) => a.amount - b.amount);
  const top3 = expenses.slice(0, 3);

  // --- VERDICT (plain English) ---
  const problems = [];
  if (worstCase.value < 0) {
    problems.push({ sev: 3, msg: `Even after dipping into the mortgage money, you run out of cash around the ${ordinal(worstCase.day)}.` });
  } else if (scariest.value < 0) {
    problems.push({ sev: 2, msg: `Checking goes negative around the ${ordinal(scariest.day)} — you'll lean on the mortgage money to get by.` });
  }
  if (monthEnd < 0) {
    problems.push({ sev: 3, msg: `You finish the month in the red at ${money(monthEnd)}.` });
  }
  if (shortfall < 0) {
    problems.push({ sev: 3, msg: `The mortgage envelope is ${money(Math.abs(shortfall))} short — put it back before month-end or the payment bounces.` });
  }

  const maxSev = problems.reduce((m, p) => Math.max(m, p.sev), 0);
  let verdict, bg, fg;
  if (maxSev === 3) { verdict = '🔴 You\u2019re screwed'; bg = '#f4cccc'; fg = '#990000'; }
  else if (maxSev === 2) { verdict = '🟡 It\u2019s tight'; bg = '#fff2cc'; fg = '#7f6000'; }
  else { verdict = '🟢 You\u2019re okay'; bg = '#d9ead3'; fg = '#38761d'; }
  const why = problems.length
    ? problems.map(p => p.msg).join('  ')
    : `Lowest you'll dip is ${money(scariest.value)} on the ${ordinal(scariest.day)}, the mortgage is covered, and you can spend ${money(safePerDay)} a day. Relax.`;

  // --- HUMAN-READABLE SENTENCES ---
  const safeSentence = `Spend up to this each day on non-bills and still finish ${monthName} above zero.`;
  const scarySentence = scariest.value < 0
    ? 'Your checking goes below zero here — you\u2019ll need the mortgage buffer or a cushion to cover it.'
    : 'Even at its lowest, checking stays above zero. Breathing room all month.';
  const mortgageStatus = shortfall >= 0 ? '✅ Yes' : `⚠️ No — ${money(Math.abs(shortfall))} short`;
  const mortgageSentence = shortfall >= 0
    ? `On track to have the full ${money(escrowTarget)} ready for the payment at month-end.`
    : `You\u2019ve borrowed from it. Put ${money(Math.abs(shortfall))} back before month-end.`;

  // --- LAYOUT (col B content, col C number), from row 2 ---
  const rows = [
    [`Am I okay this month?   —   ${monthName}`, ''], // 2
    ['', ''],                                          // 3
    [verdict, ''],                                     // 4
    [why, ''],                                         // 5
    ['', ''],                                          // 6
    ['Safe to spend each day', safePerDay],            // 7
    [safeSentence, ''],                                // 8
    ['', ''],                                          // 9
    [`Scariest day (the ${ordinal(scariest.day)})`, scariest.value], // 10
    [scarySentence, ''],                               // 11
    ['', ''],                                          // 12
    ['Is the mortgage covered?', mortgageStatus],      // 13
    [mortgageSentence, ''],                            // 14
    ['Escrow due at month-end (edit if it changes)', escrowTarget], // 15
    ['', ''],                                          // 16
    ['Savings', savings],                              // 17
    ['Owed to the Novaks', owedNovaks],                // 18
    ['', ''],                                          // 19
    ['Biggest expenses this month', ''],               // 20
    [top3[0] ? `${top3[0].vendor}  (${ordinal(top3[0].day)})` : '—', top3[0] ? top3[0].amount : ''], // 21
    [top3[1] ? `${top3[1].vendor}  (${ordinal(top3[1].day)})` : '', top3[1] ? top3[1].amount : ''],  // 22
    [top3[2] ? `${top3[2].vendor}  (${ordinal(top3[2].day)})` : '', top3[2] ? top3[2].amount : ''],  // 23
  ];

  dash.getRange('A1:H40').clear();
  dash.getRange(2, 2, rows.length, 2).setValues(rows);

  // --- FORMATTING ---
  dash.setColumnWidth(1, 30).setColumnWidth(2, 290).setColumnWidth(3, 160);
  ['B2:C2', 'B4:C4', 'B5:C5', 'B8:C8', 'B11:C11', 'B14:C14', 'B20:C20'].forEach(rng => dash.getRange(rng).merge());

  dash.getRange('B2').setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold').setFontSize(15).setHorizontalAlignment('center');
  dash.setRowHeight(2, 34).setRowHeight(4, 46);
  dash.getRange('B4').setBackground(bg).setFontColor(fg).setFontWeight('bold').setFontSize(20).setHorizontalAlignment('center');
  dash.getRange('B5').setFontColor('#666666').setFontStyle('italic').setWrap(true).setFontSize(11);

  // The three "answer" metrics get a light card + bold label/number.
  ['B7:C7', 'B10:C10', 'B13:C13'].forEach(rng =>
    dash.getRange(rng).setBackground('#f8f9fa').setFontWeight('bold').setFontSize(12)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID));

  // Gray italic context sentences under each metric.
  ['B8', 'B11', 'B14'].forEach(c =>
    dash.getRange(c).setFontColor('#888888').setFontStyle('italic').setWrap(true).setFontSize(10));

  // Mortgage status colored by outcome.
  dash.getRange('C13').setFontColor(shortfall >= 0 ? '#38761d' : '#990000').setHorizontalAlignment('right');

  // Money cells: right-aligned, red negatives.
  ['C7', 'C10', 'C15', 'C17', 'C18', 'C21:C23'].forEach(rng =>
    dash.getRange(rng).setNumberFormat(CUR).setHorizontalAlignment('right'));

  // Expenses mini-section.
  dash.getRange('B20').setBackground('#f1f3f4').setFontWeight('bold').setFontColor('#5f6368').setFontSize(11);
  dash.getRange('B21:C23').setBackground('#f8f9fa');

  // Editable blue cells (last, so they win): escrow, savings, Novaks.
  ['C15', 'C17', 'C18'].forEach(c =>
    dash.getRange(c).setBackground('#e8f0fe').setBorder(true, true, true, true, true, true, '#a4c2f4', SpreadsheetApp.BorderStyle.SOLID));
  dash.getRange('C15').setNote('Edit me — full escrow/mortgage amount due at month-end');
  dash.getRange('C17').setNote('Edit me — current savings balance');
  dash.getRange('C18').setNote('Edit me — total owed to the Novaks');

  SpreadsheetApp.flush();
  ui.alert(`Dashboard updated for ${monthName}.\n${verdict}`);
}