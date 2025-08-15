#!/usr/bin/env node
/**
 * Update an existing Excel workbook IN PLACE.
 * - Converts text numbers in Data!G/H to numeric (--fixNumbers)
 * - Reads sprint window from Window!A2:A (or --window or --last)
 * - Rebuilds SprintCalc, KPIs
 * - Colors Data rows for low-predictability sprints
 * - Reasons sheet:
 *     • Sprint Health Summary (per sprint)
 *     • Not Good Items (low sprints)
 *     • Good Items (at/above threshold)
 *     • Per-item Contribution % (to the sprint, vs CommittedPts)
 *
 * Usage:
 *   node kpi-update-inplace.js --file MyWorkbook.xlsx --last 6 --threshold 0.9 --fixNumbers
 *   node kpi-update-inplace.js --file MyWorkbook.xlsx --window "2025.S16,2025.S15" --threshold 0.92
 */

const ExcelJS = require('exceljs');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');

const argv = yargs(hideBin(process.argv))
  .option('file',      { alias: 'f', type: 'string', demandOption: true, describe: 'Path to EXISTING .xlsx' })
  .option('last',      { type: 'number', default: 6, describe: 'If Window is empty, take last N sprints from Data' })
  .option('window',    { type: 'string', describe: 'Comma-separated sprint IDs to use instead of Window' })
  .option('threshold', { type: 'number', default: 0.90, describe: 'Predictability threshold (0..1) for Sprint Health' })
  .option('fixNumbers',{ type: 'boolean', default: true, describe: 'Coerce Data!G/H values to numeric' })
  .help().argv;

// ------- helpers -------
const NBSP = /\u00A0/g;
const clean = v => (v ?? '').toString().replace(NBSP, '').trim();
const lower = v => clean(v).toLowerCase();
const toNum = v => {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const n = Number(clean(v).replace(/,/g,''));
  return Number.isFinite(n) ? n : 0;
};
const sampleStdDev = arr => {
  const n = arr.length; if (n < 2) return 0;
  const m = arr.reduce((a,b)=>a+b,0)/n;
  const sumSq = arr.reduce((s,x)=>s + (x-m)*(x-m), 0);
  return Math.sqrt(sumSq/(n-1));
};
const sampleVar = arr => { const sd = sampleStdDev(arr); return sd*sd; };

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(argv.file);

  // --- Data sheet ---
  const wsData = wb.getWorksheet('Data');
  if (!wsData) { console.error('❌ Sheet "Data" not found.'); process.exit(1); }

  // header map
  const headers = {};
  wsData.getRow(1).eachCell((c, col) => { headers[clean(c.value)] = col; });

  const need = [
    'Team','Sprint','Sprint Start','Sprint End',
    'Item Type','Item ID','Story Points Committed','Story Points','State','Outcome'
  ];
  for (const h of need) {
    if (!headers[h]) { console.error(`❌ Missing Data header: "${h}"`); process.exit(1); }
  }

  const colSprint    = headers['Sprint'];
  const colState     = headers['State'];
  const colOutcome   = headers['Outcome'];
  const colCommitted = headers['Story Points Committed'];
  const colPoints    = headers['Story Points'];
  const colItemId    = headers['Item ID'];

  // read rows & (optionally) coerce numbers in-sheet
  const dataRows = [];
  const lastDataRow = wsData.lastRow.number;

  for (let r = 2; r <= lastDataRow; r++) {
    const row = wsData.getRow(r);
    const sprint = clean(row.getCell(colSprint).value);
    if (!sprint) continue;

    let committed = row.getCell(colCommitted).value;
    let storypts  = row.getCell(colPoints).value;
    const committedNum = toNum(committed);
    const storyptsNum  = toNum(storypts);

    if (argv.fixNumbers) {
      const cCell = row.getCell(colCommitted);
      const sCell = row.getCell(colPoints);
      cCell.value = committedNum; cCell.numFmt = '0';
      sCell.value = storyptsNum;  sCell.numFmt = '0';
    }

    dataRows.push({
      RowIndex: r,
      Team       : clean(row.getCell(headers['Team']).value),
      Sprint     : sprint,
      SprintStart: clean(row.getCell(headers['Sprint Start']).value),
      SprintEnd  : clean(row.getCell(headers['Sprint End']).value),
      ItemType   : clean(row.getCell(headers['Item Type']).value),
      ItemID     : clean(row.getCell(colItemId).value),
      CommittedPts: committedNum,
      StoryPoints  : storyptsNum,
      State      : clean(row.getCell(colState).value),
      Outcome    : clean(row.getCell(colOutcome).value)
    });
  }
  if (!dataRows.length) { console.error('❌ No data rows found.'); process.exit(1); }

  // --- Window sprints ---
  let windowSprints = [];
  if (argv.window) {
    windowSprints = argv.window.split(',').map(s=>clean(s)).filter(Boolean);
  } else {
    const wsWindow = wb.getWorksheet('Window');
    if (wsWindow) {
      for (let r=2; r<=wsWindow.lastRow.number; r++) {
        const v = clean(wsWindow.getCell(`A${r}`).value);
        if (v) windowSprints.push(v);
      }
    }
    if (!windowSprints.length) {
      const seen = new Set();
      for (let i = dataRows.length-1; i>=0 && windowSprints.length < argv.last; i--) {
        const s = dataRows[i].Sprint;
        if (s && !seen.has(s)) { windowSprints.unshift(s); seen.add(s); }
      }
    }
  }
  if (!windowSprints.length) { console.error('❌ No sprint window selected.'); process.exit(1); }

  // --- Sprint aggregation (Accepted-only) ---
  function aggSprint(sprint) {
    const rows = dataRows.filter(d => d.Sprint === sprint);
    const committed = rows.filter(d => lower(d.Outcome) === 'committed')
                          .reduce((sum,d)=>sum+d.CommittedPts,0);
    const accepted  = rows.filter(d => lower(d.State) === 'accepted')
                          .reduce((sum,d)=>sum+d.StoryPoints,0);
    const added     = rows.filter(d => lower(d.Outcome) === 'added')
                          .reduce((sum,d)=>sum+d.StoryPoints,0);
    const removed   = rows.filter(d => lower(d.Outcome) === 'removed')
                          .reduce((sum,d)=>sum+d.StoryPoints,0);
    const pred = committed>0 ? accepted/committed : 0;
    const scope = committed>0 ? (added+removed)/committed : 0;

    // counts & pts for health
    const countAdded   = rows.filter(d => lower(d.Outcome) === 'added').length;
    const countRemoved = rows.filter(d => lower(d.Outcome) === 'removed').length;
    const unfinished   = rows.filter(d => d.CommittedPts>0 && lower(d.State)!=='accepted');
    const countUnfin   = unfinished.length;
    const ptsUnfin     = unfinished.reduce((sum,d)=>sum+d.CommittedPts,0);

    return {
      Sprint: sprint,
      CommittedPts: committed,
      AcceptedPts: accepted,
      AddedPts: added,
      RemovedPts: removed,
      PerSprintPredictability: pred,
      PerSprintScopeChange: scope,
      CountAdded: countAdded,
      CountRemoved: countRemoved,
      CountUnfinishedCommitted: countUnfin,
      PtsUnfinishedCommitted: ptsUnfin
    };
  }
  const sprintAgg = windowSprints.map(aggSprint);

  // --- KPIs over window ---
  const arrAccepted  = sprintAgg.map(s=>s.AcceptedPts);
  const arrCommitted = sprintAgg.map(s=>s.CommittedPts);
  const arrScope     = sprintAgg.map(s=>s.PerSprintScopeChange);
  const avgAccepted  = arrAccepted.length ? arrAccepted.reduce((a,b)=>a+b,0)/arrAccepted.length : 0;
  const sdAccepted   = sampleStdDev(arrAccepted);
  const predictability = avgAccepted>0 ? (1 - sdAccepted/avgAccepted) : 0;
  const committedSum = arrCommitted.reduce((a,b)=>a+b,0);
  const addRemSum    = sprintAgg.reduce((sum,s)=>sum + s.AddedPts + s.RemovedPts, 0);
  const volatility   = committedSum>0 ? (addRemSum/committedSum) : 0;
  const avgCommitted = arrCommitted.length ? arrCommitted.reduce((a,b)=>a+b,0)/arrCommitted.length : 0;
  const avgScopeChange = arrScope.length ? arrScope.reduce((a,b)=>a+b,0)/arrScope.length : 0;
  const varianceAccepted = sampleVar(arrAccepted);

  // --- Determine low & good sprints; color rows in Data ---
  const threshold = typeof argv.threshold === 'number' ? argv.threshold : 0.90;
  const lowSet  = new Set(sprintAgg.filter(s => s.PerSprintPredictability < threshold).map(s=>s.Sprint));
  const goodSet = new Set(sprintAgg.filter(s => s.PerSprintPredictability >= threshold).map(s=>s.Sprint));

  // Colors (ARGB)
  const YELLOW = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF4B3' } }; // light yellow
  const RED    = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFA8A8' } }; // light red
  const ORANGE = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFCC99' } }; // orange

  // Clear previous fills
  for (let r=2; r<=lastDataRow; r++) {
    wsData.getRow(r).eachCell(cell => { cell.fill = undefined; });
  }

  // Color only rows in low sprints
  for (const d of dataRows) {
    if (!lowSet.has(d.Sprint)) continue;
    const out = lower(d.Outcome);
    const st  = lower(d.State);
    if (out === 'added') {
      wsData.getRow(d.RowIndex).eachCell(c => c.fill = YELLOW);
    } else if (out === 'removed') {
      wsData.getRow(d.RowIndex).eachCell(c => c.fill = RED);
    } else if (d.CommittedPts > 0 && st !== 'accepted') {
      wsData.getRow(d.RowIndex).eachCell(c => c.fill = ORANGE);
    }
  }

  // --- SprintCalc (recreate) ---
  let wsSC = wb.getWorksheet('SprintCalc');
  if (wsSC) wb.removeWorksheet(wsSC.id);
  wsSC = wb.addWorksheet('SprintCalc');
  wsSC.columns = [
    { header:'Sprint', key:'Sprint', width:16 },
    { header:'CommittedPts', key:'CommittedPts', width:16 },
    { header:'AcceptedPts', key:'AcceptedPts', width:16 },
    { header:'AddedPts', key:'AddedPts', width:14 },
    { header:'RemovedPts', key:'RemovedPts', width:14 },
    { header:'PerSprint Predictability', key:'PerSprintPredictability', width:26 },
    { header:'PerSprint ScopeChange%', key:'PerSprintScopeChange', width:26 }
  ];
  wsSC.addRows(sprintAgg);
  wsSC.getColumn('PerSprintPredictability').numFmt = '0.00%';
  wsSC.getColumn('PerSprintScopeChange').numFmt = '0.00%';

  // --- KPIs (recreate) ---
  let wsKPI = wb.getWorksheet('KPIs');
  if (wsKPI) wb.removeWorksheet(wsKPI.id);
  wsKPI = wb.addWorksheet('KPIs');
  wsKPI.addRows([
    ['Metric','Value'],
    ['Average Accepted (window)', avgAccepted],
    ['StdDev (Accepted) (window)', sdAccepted],
    ['Predictability % (1 - stdev/avg)', predictability],
    ['Volatility % ((Added+Removed)/Committed)', volatility],
    ['Average Committed (window)', avgCommitted],
    ['Average Scope Change % (per-sprint mean)', avgScopeChange],
    ['Variance (Accepted) (window)', varianceAccepted]
  ]);
  wsKPI.getCell('B4').numFmt = '0.00%';
  wsKPI.getCell('B5').numFmt = '0.00%';
  wsKPI.getCell('B7').numFmt = '0.00%';

  // --- Reasons (recreate): Sprint Health + Not Good Items + Good Items ---
  let wsR = wb.getWorksheet('Reasons');
  if (wsR) wb.removeWorksheet(wsR.id);
  wsR = wb.addWorksheet('Reasons');

  wsR.addRow(['Predictability threshold (0–1)', threshold]);
  wsR.addRow([]);

  // Sprint Health Summary
  wsR.addRow(['Sprint Health Summary']);
  wsR.addRow(['Sprint','Predictability%','Status','Added (items)','Removed (items)','Unfinished (items)','Added Pts','Removed Pts','Unfinished Committed Pts']);
  // build per-sprint detail once
  const perSprintDetail = new Map();
  for (const s of sprintAgg) {
    const rows = dataRows.filter(d => d.Sprint === s.Sprint);
    const addedCount = rows.filter(d => lower(d.Outcome)==='added').length;
    const removedCount = rows.filter(d => lower(d.Outcome)==='removed').length;
    const unfinRows = rows.filter(d => d.CommittedPts>0 && lower(d.State)!=='accepted');
    const unfinCount = unfinRows.length;
    const addedPts = rows.filter(d => lower(d.Outcome)==='added').reduce((sum,d)=>sum+d.StoryPoints,0);
    const removedPts = rows.filter(d => lower(d.Outcome)==='removed').reduce((sum,d)=>sum+d.StoryPoints,0);
    const unfinPts = unfinRows.reduce((sum,d)=>sum+d.CommittedPts,0);
    perSprintDetail.set(s.Sprint, {addedCount, removedCount, unfinCount, addedPts, removedPts, unfinPts});
  }
  const startHealth = wsR.lastRow.number + 1;
  for (const s of sprintAgg) {
    const det = perSprintDetail.get(s.Sprint) || {addedCount:0, removedCount:0, unfinCount:0, addedPts:0, removedPts:0, unfinPts:0};
    const status = s.PerSprintPredictability >= threshold ? 'Good' : 'Not Good';
    const row = wsR.addRow([
      s.Sprint,
      s.PerSprintPredictability,
      status,
      det.addedCount,
      det.removedCount,
      det.unfinCount,
      det.addedPts,
      det.removedPts,
      det.unfinPts
    ]);
    wsR.getCell(`B${row.number}`).numFmt = '0.00%';
  }

  // helper to compute per-item contribution %
  const sprintCommittedMap = new Map(sprintAgg.map(s => [s.Sprint, s.CommittedPts]));
  const makeItemRecord = d => {
    const out = lower(d.Outcome);
    const st  = lower(d.State);
    let reason = '';
    let contribNumerator = 0;
    if (out === 'added') {
      reason = 'Scope added mid-sprint';
      contribNumerator = d.StoryPoints;
    } else if (out === 'removed') {
      reason = 'Scope removed mid-sprint';
      contribNumerator = d.StoryPoints;
    } else if (d.CommittedPts > 0 && st !== 'accepted') {
      reason = 'Committed (planned) but not accepted';
      contribNumerator = d.CommittedPts;
    } else {
      return null; // not a contributing row
    }
    const sprintCommitted = sprintCommittedMap.get(d.Sprint) || 0;
    const contribPct = sprintCommitted > 0 ? (contribNumerator / sprintCommitted) : 0;
    return {
      Sprint: d.Sprint,
      ItemID: d.ItemID,
      Outcome: d.Outcome,
      State: d.State,
      CommittedPts: d.CommittedPts,
      StoryPoints: d.StoryPoints,
      Reason: reason,
      ContributionPct: contribPct
    };
  };

  // NOT GOOD items
  wsR.addRow([]);
  wsR.addRow(['Not Good Items (contributing)']);
  wsR.addRow(['Sprint','Item ID','Outcome','State','CommittedPts','StoryPoints','Reason','Contribution % (of sprint committed)']);
  const notGoodItems = dataRows
    .filter(d => lowSet.has(d.Sprint))
    .map(makeItemRecord)
    .filter(Boolean);
  const ngStart = wsR.lastRow.number + 1;
  wsR.addRows(notGoodItems);
  for (let r = ngStart; r < ngStart + notGoodItems.length; r++) {
    wsR.getCell(`H${r}`).numFmt = '0.00%';
  }

  // GOOD items (for context; same rules, but sprint is >= threshold)
  wsR.addRow([]);
  wsR.addRow(['Good Items (for context)']);
  wsR.addRow(['Sprint','Item ID','Outcome','State','CommittedPts','StoryPoints','Reason','Contribution % (of sprint committed)']);
  const goodItems = dataRows
    .filter(d => goodSet.has(d.Sprint))
    .map(makeItemRecord)
    .filter(Boolean);
  const gdStart = wsR.lastRow.number + 1;
  wsR.addRows(goodItems);
  for (let r = gdStart; r < gdStart + goodItems.length; r++) {
    wsR.getCell(`H${r}`).numFmt = '0.00%';
  }

  await wb.xlsx.writeFile(argv.file);
  console.log(`✅ Updated workbook in place: ${argv.file}`);
})();
