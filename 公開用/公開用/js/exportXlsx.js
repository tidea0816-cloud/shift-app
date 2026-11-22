const FIRST_COL = 7; // G
const MAX_DAYS = 31;
const DOW = '日月火水木金土';
const C = { pink: 'FFF4CCCC', orange: 'FFFCE4D6', blue: 'FFC9DAF8', green: 'FFD9EAD3', yellow: 'FFFFF2CC' };
const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb }, bgColor: { argb } });
const NOFILL = { type: 'pattern', pattern: 'none' };
const STAFF_ROWS = [20, 21, 22, 23, 24];
const PART_ROWS = [29, 31, 33, 35, 37];
const colName = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const lastCol = colName(FIRST_COL + MAX_DAYS - 1); // AK
const clone = (o) => JSON.parse(JSON.stringify(o));
const setFill = (c, fill) => { c.style = { ...clone(c.style), fill: clone(fill) }; };
const isMasterOrPlain = (c) => !(c.isMerged && c.master !== c);

export async function buildXlsx({ month, members, schedule, templateUrl = 'template.xlsx', mustDays = {}, events = {}, help = {} }) {
  const res = await fetch(templateUrl);
  if (!res.ok) throw new Error('template.xlsx を読み込めません: ' + res.status);
  const wb = new globalThis.ExcelJS.Workbook();
  await wb.xlsx.load(await res.arrayBuffer());
  const ws = wb.worksheets[0];
  const days = month.days;
  const col = (d) => FIRST_COL + d - 1;

  const staff = members.filter((m) => m.kind === 'staff').slice(0, STAFF_ROWS.length);
  const parts = members.filter((m) => m.kind === 'part').slice(0, PART_ROWS.length);

  // 名前→テンプレのC列（生年など）
  const tplC = {};
  for (const r of STAFF_ROWS) {
    const n = ws.getCell(r, 4).value;
    if (n && ws.getCell(r, 3).value != null) tplC[n] = ws.getCell(r, 3).value;
  }

  for (const r of ['V10:V12', 'H14:H15', 'N14:N15']) { try { ws.unMergeCells(r); } catch { } }

  // 日付列の全クリア（日付・予算・イベント・社員・バイト）
  const clearRows = [8, 10, 11, 12, 13, 14, 15, 16, 17, ...STAFF_ROWS, ...PART_ROWS.flatMap((r) => [r, r + 1])];
  for (let d = 1; d <= MAX_DAYS; d++) {
    for (const r of [6, 7, ...clearRows]) {
      const c = ws.getCell(r, col(d));
      if (isMasterOrPlain(c)) c.value = null;
      if (r >= 10) setFill(c, NOFILL);
    }
    if (d > days) {
      for (const r of [6, 7]) setFill(ws.getCell(r, col(d)), NOFILL);
      ws.getCell(26, col(d)).value = null;
    }
  }

  ws.getCell('J2').value = month.month;
  for (let d = 1; d <= days; d++) {
    const pink = month.weekendOrHoliday.has(d);
    ws.getCell(6, col(d)).value = d;
    ws.getCell(7, col(d)).value = DOW[month.dow[d]];
    for (const r of [6, 7]) setFill(ws.getCell(r, col(d)), pink ? solid(C.pink) : NOFILL);
  }

  // イベント（10〜13行）・社員予定（14〜17行）：1日1セル（縦4行を結合）、縦書き
  const famOf = (id) => String(members.find((m) => m.id === id)?.name || '').split(/\s+/)[0];
  const evLabel = (e) => {
    const t = String(e.text || '').trim();
    if (e.type !== 'drill' || !t) return t;
    const who = [...new Set([...(e.staffIds || []), ...(e.staffId ? [e.staffId] : [])])].map(famOf).filter(Boolean).join('・');
    return `${t} 8:15${who ? ' ' + who : ''}`;
  };
  const evText = (d, kind) => [...new Set((events[d] || []).filter((e) => e.kind === kind).map(evLabel).filter(Boolean))].join('・');
  const putVertical = (r1, r2, d, text) => {
    if (!text) return;
    ws.mergeCells(r1, col(d), r2, col(d));
    const c = ws.getCell(r1, col(d));
    c.value = text;
    c.style = { ...clone(c.style), alignment: { horizontal: 'center', vertical: 'top', wrapText: true, textRotation: 'vertical' }, font: { ...clone(c.style.font || {}), size: 9 } };
  };
  for (let d = 1; d <= days; d++) { putVertical(10, 13, d, evText(d, 'store')); putVertical(14, 17, d, evText(d, 'staff')); }

  // 応援候補（ヘルプ行の下の名前欄＝39行）
  const fam = (n) => String(n).split(/\s+/)[0];
  let anyHelp = false;
  for (let d = 1; d <= days; d++) {
    const names = [...new Set((help[d] || []).map((h) => fam(h.name)))];
    const c = ws.getCell(39, col(d));
    c.value = names.length ? '候補:' + names.join('/') : null;
    if (names.length) {
      anyHelp = true;
      c.style = { ...clone(c.style), alignment: { horizontal: 'center', vertical: 'top', wrapText: true, textRotation: 'vertical' }, font: { ...clone(c.style.font || {}), size: 8 } };
    }
  }
  if (anyHelp) ws.getRow(39).height = Math.max(40, Math.min(110, 12 * Math.max(...Object.values(help).map((a) => ('候補:' + a.map((h) => fam(h.name)).join('/')).length))));

  // 社員
  const needExtraRow = staff.length > 4;
  if (needExtraRow) {
    for (let c = 1; c <= 47; c++) {
      const src = ws.getCell(23, c), dst = ws.getCell(24, c);
      dst.style = { ...src.style };
    }
    for (const [i, f] of ['COUNTIF(G24:AK24,"A*")', 'COUNTIF(G24:AK24,"D*")', 'COUNTIF(G24:AK24,"B*")',
      'COUNTA(G24:AK24)-SUM(AM24:AO24)-SUM(AQ24:AT24)', 'COUNTIF(G24:AK24,"/")+COUNTIF(G24:AK24,"×")',
      'COUNTIF(G24:AK24,"有")', 'COUNTIF(G24:AK24,"リ")', 'COUNTIF(G24:AK24,"欠")'].entries())
      ws.getCell(24, 39 + i).value = { formula: f };
    ws.getCell(24, 47).value = 0;
  }
  STAFF_ROWS.forEach((r, i) => {
    const m = staff[i];
    ws.getCell(r, 4).value = m ? m.name : null;
    ws.getCell(r, 3).value = m ? (tplC[m.name] ?? null) : null;
    if (!m) return;
    const sch = schedule.staff?.[m.id] || {};
    const must = new Set(mustDays[m.id] || []);
    for (let d = 1; d <= days; d++) {
      const code = sch[d];
      if (code == null || code === '') continue;
      const c = ws.getCell(r, col(d));
      c.value = code;
      if (code === 'リ') setFill(c, solid(C.blue));
      else if (code === 'B15') setFill(c, solid(C.orange));
      else if (code === 'A815' || must.has(d)) setFill(c, solid(C.green));
    }
  });

  // バイト（開始/終了の2段）
  PART_ROWS.forEach((r, i) => {
    const m = parts[i];
    ws.getCell(r, 4).value = m ? m.name : (i === PART_ROWS.length - 1 ? 'ヘルプ' : null);
    if (!m) return;
    const sch = schedule.part?.[m.id] || {};
    for (let d = 1; d <= days; d++) {
      const t = sch[d];
      if (!t) continue;
      ws.getCell(r, col(d)).value = t[0];
      ws.getCell(r + 1, col(d)).value = t[1];
      if (schedule.tentative?.[m.id]?.[d]) { setFill(ws.getCell(r, col(d)), solid(C.yellow)); setFill(ws.getCell(r + 1, col(d)), solid(C.yellow)); }
    }
  });

  // 人員数（社員行の範囲に合わせて常に数式で入れ直す）
  const sEnd = needExtraRow ? 24 : 23;
  for (let d = 1; d <= days; d++) {
    const L = colName(col(d));
    ws.getCell(26, col(d)).value = {
      formula: `COUNTIF(${L}20:${L}${sEnd},"A*")+COUNTIF(${L}20:${L}${sEnd},"B*")+COUNTIF(${L}20:${L}${sEnd},"D*")+0.5*(${PART_ROWS.map((r) => `COUNT(${L}${r})`).join('+')})`,
      result: Object.values(schedule.staff || {}).filter((o) => /^[ABD]/.test(o[d] || '')).length
        + 0.5 * Object.values(schedule.part || {}).filter((o) => o[d]).length,
    };
  }
  ws.addConditionalFormatting({
    ref: `G26:${colName(col(days))}26`,
    rules: [{ type: 'expression', formulae: ['G26<=2.5'], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }],
  });

  wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
