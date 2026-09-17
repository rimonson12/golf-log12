/*
 * Golf Log — Google Sheets receiver
 * ---------------------------------
 * 1. Open a new Google Sheet → Extensions → Apps Script.
 * 2. Delete the sample code, paste this whole file, save.
 * 3. Run "setup" once from the toolbar (choose the function, press Run, accept permissions).
 *    This creates the Rounds / Holes / Shots / Dashboard tabs.
 * 4. Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone → Deploy.
 * 5. Copy the Web app URL (ends in /exec) into Golf Log → Settings → Apps Script web app URL.
 *
 * Every push is an upsert: rows for that round ID are deleted and re-written,
 * so editing a round in the app and pushing again never duplicates data.
 */

const HEADERS = {
  Rounds: ['RoundID','Date','Course','Type','Mode','Holes','Gross','Par','To Par','To Par (18)','Putts','Penalties',
           'FIR Hit','FIR Opp','FIR %','GIR Hit','GIR Opp','GIR %','Scr Hit','Scr Opp','Scr %','3-Putts','Notes','Pushed At','Rating','Slope','Differential','Course Hcp','Net'],
  Holes:  ['RoundID','Date','Course','Type','Hole','Par','Yards','Score','To Par','Putts','Penalties','FIR','GIR','Scramble','GIR Miss'],
  Shots:  ['RoundID','Date','Course','Type','Hole','Par','Shot #','Club','Lie','Result','Strike','Prox (ft)']
};
const CLUBS = ['Dr','3W','Hy','4i','5i','6i','7i','8i','9i','PW','50','54','58'];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.ping) return out({ ok: true });
    setup();
    upsertRound(body);
    return out({ ok: true, roundId: body.round.roundId });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
function doGet() { return out({ ok: true, message: 'Golf Log endpoint is live' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name) {
  let s = ss().getSheetByName(name);
  if (!s) { s = ss().insertSheet(name); }
  if (HEADERS[name] && (s.getLastRow() === 0 || s.getLastColumn() < HEADERS[name].length)) {
    s.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function toDate(s) { const p = String(s).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function ratio(a, b) { return b ? a / b : ''; }

function upsertRound(p) {
  const r = p.round;
  ['Rounds', 'Holes', 'Shots'].forEach(name => deleteRows(sheet(name), r.roundId));
  const d = toDate(r.date);
  sheet('Rounds').appendRow([
    r.roundId, d, r.course, r.type, r.mode, r.holes, r.score, r.par, r.toPar, r.holes ? r.toPar / r.holes * 18 : '',
    r.putts, r.pen, r.firHit, r.firOpp, ratio(r.firHit, r.firOpp), r.girHit, r.girOpp, ratio(r.girHit, r.girOpp),
    r.scrHit, r.scrOpp, ratio(r.scrHit, r.scrOpp), r.threePutts, r.notes || '', new Date(),
    r.rating === '' ? '' : r.rating, r.slope === '' ? '' : r.slope, r.differential === '' ? '' : r.differential
  ]);
  // Course handicap and net score as formulas, driven by the index in Dashboard!B3
  const row = sheet('Rounds').getLastRow();
  sheet('Rounds').getRange(row, 28).setFormula(`=IF(AND(F${row}=18,Y${row}<>"",Z${row}<>"",Dashboard!$B$3<>""),ROUND(Dashboard!$B$3*Z${row}/113+(Y${row}-H${row})),"")`);
  sheet('Rounds').getRange(row, 29).setFormula(`=IF(AB${row}<>"",G${row}-AB${row},"")`);
  if (p.holes.length) {
    sheet('Holes').getRange(sheet('Holes').getLastRow() + 1, 1, p.holes.length, HEADERS.Holes.length).setValues(
      p.holes.map(h => [h.roundId, d, h.course, h.type, h.hole, h.par, h.yds, h.score, h.toPar, h.putts, h.pen, h.fir, h.gir, h.scramble, h.girMiss || '']));
  }
  if (p.shots.length) {
    sheet('Shots').getRange(sheet('Shots').getLastRow() + 1, 1, p.shots.length, HEADERS.Shots.length).setValues(
      p.shots.map(s => [s.roundId, d, s.course, s.type, s.hole, s.par, s.shot, s.club, s.lie, s.result, s.strike, s.prox]));
  }
  ['Rounds', 'Holes', 'Shots'].forEach(name => sheet(name).getRange('B2:B').setNumberFormat('yyyy-mm-dd'));
  sheet('Rounds').getRange('O2:O').setNumberFormat('0%');
  sheet('Rounds').getRange('R2:R').setNumberFormat('0%');
  sheet('Rounds').getRange('U2:U').setNumberFormat('0%');
  sheet('Rounds').getRange('J2:J').setNumberFormat('0.0');
  sheet('Rounds').getRange('AA2:AA').setNumberFormat('0.0');
}

function deleteRows(s, roundId) {
  const last = s.getLastRow();
  if (last < 2) return;
  const ids = s.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === roundId) s.deleteRow(i + 2);
  }
}

/* ---------------- one-time setup: tabs + dashboard ---------------- */
function cond(c) { return c === 'D$4' ? '(Rounds!$D$2:$D<>"")' : `(Rounds!$D$2:$D=${c})`; }
function setup() {
  ['Rounds', 'Holes', 'Shots'].forEach(sheet);
  if (ss().getSheetByName('Dashboard')) return;
  const d = ss().insertSheet('Dashboard', 0);
  d.getRange('A1').setValue('Golf Log').setFontSize(18).setFontWeight('bold');
  d.getRange('A2').setValue('Live from the Rounds, Holes and Shots tabs. Enter your handicap index in B3 for net scores; change the club-table filter below.').setFontColor('#6a716c');
  d.getRange('A3').setValue('Your handicap index').setFontWeight('bold');
  d.getRange('B3').setValue(0).setBackground('#fff2a8').setFontColor('#0000ff').setNumberFormat('0.0');
  d.getRange('C3').setValue('← edit; net = gross − round(index × slope ÷ 113 + rating − par)').setFontColor('#6a716c');
  d.getRange('B4:D4').setValues([['Real', 'Sim', '*']]).setFontColor('#bbbbbb').setFontSize(8);
  d.getRange('A5:D5').setValues([['Metric', 'Real', 'Simulator', 'Combined']]).setFontWeight('bold');

  const rows = [];
  const add = (label, fn, fmt) => rows.push({ label, fn, fmt });
  add('Rounds',                    c => `=COUNTIF(Rounds!$D$2:$D,${c})`, '0');
  add('Avg gross (18-hole rounds)', c => `=IFERROR(AVERAGEIFS(Rounds!$G:$G,Rounds!$D:$D,${c},Rounds!$F:$F,18),"")`, '0.0');
  add('Avg net (18-hole rounds)',  c => `=IFERROR(AVERAGEIFS(Rounds!$AC:$AC,Rounds!$D:$D,${c},Rounds!$AC:$AC,"<>"),"")`, '0.0');
  add('Best gross (18-hole)',      c => `=IFERROR(MINIFS(Rounds!$G:$G,Rounds!$D:$D,${c},Rounds!$F:$F,18),"")`, '0');
  add('Best net (18-hole)',        c => `=IFERROR(MINIFS(Rounds!$AC:$AC,Rounds!$D:$D,${c},Rounds!$AC:$AC,"<>"),"")`, '0');
  add('Avg to par (per 18)',       c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$I:$I)/SUMIF(Rounds!$D:$D,${c},Rounds!$F:$F)*18,"")`, '+0.0;-0.0;0.0');
  add('Putts (per 18)',            c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$K:$K)/SUMIF(Rounds!$D:$D,${c},Rounds!$F:$F)*18,"")`, '0.0');
  add('Penalties (per 18)',        c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$L:$L)/SUMIF(Rounds!$D:$D,${c},Rounds!$F:$F)*18,"")`, '0.0');
  add('Fairways hit',              c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$M:$M)/SUMIF(Rounds!$D:$D,${c},Rounds!$N:$N),"")`, '0%');
  add('Greens in regulation',      c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$P:$P)/SUMIF(Rounds!$D:$D,${c},Rounds!$Q:$Q),"")`, '0%');
  add('Scrambling',                c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$S:$S)/SUMIF(Rounds!$D:$D,${c},Rounds!$T:$T),"")`, '0%');
  add('3-putts per round',         c => `=IFERROR(SUMIF(Rounds!$D:$D,${c},Rounds!$V:$V)/COUNTIF(Rounds!$D$2:$D,${c}),"")`, '0.0');
  add('Par 3 avg to par',          c => `=IFERROR(AVERAGEIFS(Holes!$I:$I,Holes!$D:$D,${c},Holes!$F:$F,3),"")`, '+0.00;-0.00;0.00');
  add('Par 4 avg to par',          c => `=IFERROR(AVERAGEIFS(Holes!$I:$I,Holes!$D:$D,${c},Holes!$F:$F,4),"")`, '+0.00;-0.00;0.00');
  add('Par 5 avg to par',          c => `=IFERROR(AVERAGEIFS(Holes!$I:$I,Holes!$D:$D,${c},Holes!$F:$F,5),"")`, '+0.00;-0.00;0.00');
  add('Avg differential',          c => `=IFERROR(SUMIF(Rounds!$D$2:$D,${c},Rounds!$AA$2:$AA)/COUNTIFS(Rounds!$D$2:$D,${c},Rounds!$AA$2:$AA,"<>"),"")`, '0.0');
  add('Best differential',         c => `=IFERROR(MINIFS(Rounds!$AA:$AA,Rounds!$D:$D,${c},Rounds!$AA:$AA,"<>"),"")`, '0.0');
  const IDX = rows.length, THR = IDX + 1, CNT = IDX + 2;   // positions of the three rows below
  const R = i => 6 + i;
  // WHS-style estimate: best N of the 20 most recent rated rounds, N and adjustment from the small-sample table
  add('Estimated handicap index', c => { const n = c.replace('$4', '$' + R(CNT)), thr = c.replace('$4', '$' + R(THR));
    return `=IF(${n}<3,"",ARRAYFORMULA(AVERAGE(SMALL(IF(${cond(c)}*(Rounds!$AA$2:$AA<>"")*(Rounds!$B$2:$B>=${thr}),Rounds!$AA$2:$AA,""),ROW(INDIRECT("1:"&LOOKUP(${n},{3,4,5,6,7,9,12,15,17,19,20},{1,1,1,2,2,3,4,5,6,7,8})))))+LOOKUP(${n},{3,4,5,6,7},{-2,-1,0,-1,0})))`; }, '0.0');
  add('20th most recent rated date', c => `=IFERROR(ARRAYFORMULA(LARGE(IF(${cond(c)}*(Rounds!$AA$2:$AA<>""),Rounds!$B$2:$B,""),20)),0)`, 'yyyy-mm-dd');
  add('Rated rounds counted',      c => `=SUMPRODUCT(${cond(c)}*(Rounds!$AA$2:$AA<>"")*(Rounds!$B$2:$B>=${c.replace('$4', '$' + R(THR))}))`, '0');

  rows.forEach((row, i) => {
    const r = R(i);
    d.getRange(r, 1).setValue(row.label);
    ['B$4', 'C$4', 'D$4'].forEach((crit, j) => d.getRange(r, 2 + j).setFormula(row.fn(crit)));
    d.getRange(r, 2, 1, 3).setNumberFormat(row.fmt);
  });
  d.getRange(R(THR), 1, 2, 4).setFontColor('#9aa19b').setFontSize(9);

  // club table with its own filter
  const T = R(rows.length) + 1;
  d.getRange(T, 1).setValue('By club').setFontWeight('bold').setFontSize(13);
  d.getRange(T + 1, 1).setValue('Filter');
  d.getRange(T + 1, 2).setValue('*').setBackground('#fff2a8').setFontColor('#0000ff');
  d.getRange(T + 1, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['*', 'Real', 'Sim'], true).build());
  d.getRange(T + 1, 3).setValue('(* = combined)').setFontColor('#6a716c');
  d.getRange(T + 2, 1, 1, 12).setValues([['Club', 'Shots', 'Good %', 'Left', 'Right', 'Short', 'Long', 'Penalty', 'Heel', 'Toe', 'Avg prox (ft)', 'Prox samples']]).setFontWeight('bold');
  const filt = `$B$${T + 1}`;
  CLUBS.forEach((club, i) => {
    const r = T + 3 + i, f = `Shots!$D:$D,${filt},Shots!$H:$H,$A${r}`;
    d.getRange(r, 1).setValue(club);
    d.getRange(r, 2).setFormula(`=COUNTIFS(${f})`);
    d.getRange(r, 3).setFormula(`=IFERROR(COUNTIFS(${f},Shots!$J:$J,"Good")/B${r},"")`).setNumberFormat('0%');
    ['Left', 'Right', 'Short', 'Long', 'Penalty'].forEach((res, j) => d.getRange(r, 4 + j).setFormula(`=COUNTIFS(${f},Shots!$J:$J,"${res}")`));
    d.getRange(r, 9).setFormula(`=COUNTIFS(${f},Shots!$K:$K,"Heel")`);
    d.getRange(r, 10).setFormula(`=COUNTIFS(${f},Shots!$K:$K,"Toe")`);
    d.getRange(r, 11).setFormula(`=IFERROR(AVERAGEIFS(Shots!$L:$L,${f},Shots!$L:$L,">0"),"")`).setNumberFormat('0.0');
    d.getRange(r, 12).setFormula(`=COUNTIFS(${f},Shots!$L:$L,">0")`);
  });
  d.setColumnWidth(1, 220);
  d.setFrozenRows(5);

  // trend chart: score to par per 18, by date
  const chart = d.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(ss().getSheetByName('Rounds').getRange('B1:B500'))
    .addRange(ss().getSheetByName('Rounds').getRange('J1:J500'))
    .setPosition(4, 6, 0, 0).setOption('title', 'Score to par (per 18) by date')
    .setOption('legend', { position: 'none' }).setOption('width', 520).setOption('height', 300).build();
  d.insertChart(chart);
}
