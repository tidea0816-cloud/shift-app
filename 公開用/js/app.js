import { makeMonth } from './calendarUtil.js';

const lib = {
  parse: () => import('./parseText.js'),
  cal: () => import('./calendar.js'),
  solver: () => import('./solver.js'),
  xlsx: () => import('./exportXlsx.js'),
};
const KEY = 'shiftApp.v1';
const STEPS = ['対象月', 'バイト', '社員', '前月末', '生成', '出力'];
const TITLES = ['① 対象月とメンバー', '② バイトの希望', '③ 社員の予定', '④ 前月末の状況', '⑤ 生成・手直し', '⑥ 出力'];
const ROLE = { B: 'Bメイン', A: 'Aメイン', flex: '穴埋め', help: 'ヘルプ' };
const CODES = ['A', 'D', 'B', 'A15', 'D30', 'B15', '/', '×', 'リ', '有'];
const EV = ['', '×', 'リ', '有', 'MUST'];
const EVL = { '': '', '×': '×', 'リ': 'リ', '有': '有', MUST: 'MT' };
const KINDS = [['×', '希望休(×)'], ['リ', 'リフ休(リ)'], ['有', '有給(有)'], ['MUST', '出勤マスト'], ['無視', '無視']];
const DOW = '日月火水木金土';
const BUSY_RE = /セール|フェア|祭|初売|福袋|催事|イベント/;

const DEFAULT_MEMBERS = [
  { id: 'kotani', name: '社員1', kind: 'staff', role: 'B' },
  { id: 'umemura', name: '社員2', kind: 'staff', role: 'flex' },
  { id: 'koyasu', name: '社員3', kind: 'staff', role: 'A' },
  { id: 'suzuki', name: 'ヘルプ社員', kind: 'staff', role: 'help' },
  { id: 'tanaka', name: 'バイト1', kind: 'part' },
  { id: 'honda', name: 'バイト2', kind: 'part' },
  { id: 'ohashi', name: 'バイト3', kind: 'part' },
  { id: 'miura', name: 'バイト4', kind: 'part' },
];

const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtT = t => t == null ? '' : `${String(Math.floor(t + 1e-9)).padStart(2, '0')}:${String(Math.round((t % 1) * 60)).padStart(2, '0')}`;
const parseT = v => { if (!v) return null; const [a, b] = v.split(':').map(Number); return a + b / 60; };
const shortT = t => t == null ? '' : t >= 21.5 ? 'L' : fmtT(t).replace(/^0/, '');
const famName = n => String(n).split(/\s+/)[0];
const ranges = days => {
  const a = [...days].sort((x, y) => x - y), out = [];
  for (let i = 0; i < a.length;) { let j = i; while (a[j + 1] === a[j] + 1) j++; out.push(j > i ? `${a[i]}〜${a[j]}` : `${a[i]}`); i = j + 1; }
  return out.join(', ') + '日';
};
const rgb = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16));
const cdist = (a, b) => { const x = rgb(a), y = rgb(b); return Math.hypot(x[0] - y[0], x[1] - y[1], x[2] - y[2]); };

function fresh() {
  const d = new Date();
  let y = d.getFullYear(), m = d.getMonth() + 2;
  if (m > 12) { m -= 12; y++; }
  return { step: 0, year: y, month: m, members: DEFAULT_MEMBERS.map(x => ({ ...x })), partReq: {}, staffEvents: {}, events: {}, partOff: {}, partDef: {}, prev: {}, schedule: null, colorMap: {} };
}
function load() {
  try { const j = JSON.parse(localStorage.getItem(KEY)); if (j && Array.isArray(j.members)) return { ...fresh(), ...j }; } catch { }
  return fresh();
}
const S = load();
const T = { help: null, cal: null, pending: [], busy: {}, notes: {}, raw: {}, viol: null, head: null, gen: false, err: '', sheet: null, file: null, msg: '', sv: null };
const save = () => { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch { } };

let _m = null;
const mon = () => (_m && _m.year === S.year && _m.month === S.month) ? _m : (_m = makeMonth(S.year, S.month));
const staff = () => S.members.filter(m => m.kind === 'staff');
const parts = () => S.members.filter(m => m.kind === 'part');
const byId = id => S.members.find(m => m.id === id);

function render() {
  const sc = [...document.querySelectorAll('.tbl-wrap')].map(e => e.scrollLeft);
  const y = window.scrollY;
  $('#title').textContent = TITLES[S.step];
  $('#tabs').innerHTML = STEPS.map((n, i) => `<button class="tab${i === S.step ? ' on' : i < S.step ? ' done' : ''}" data-act="go" data-i="${i}"><b>${i + 1}</b><span>${n}</span></button>`).join('');
  $('#main').innerHTML = [v1, v2, v3, v4, v5, v6][S.step](mon());
  $('#nav').innerHTML = `<button data-act="go" data-i="${S.step - 1}" ${S.step === 0 ? 'disabled' : ''}>戻る</button>` +
    (S.step < 5 ? `<button class="pri" data-act="go" data-i="${S.step + 1}">次へ：${STEPS[S.step + 1]}</button>` : '');
  document.querySelectorAll('.tbl-wrap').forEach((e, i) => { if (sc[i]) e.scrollLeft = sc[i]; });
  window.scrollTo(0, y);
  renderSheet();
  save();
  if (S.step === 4 && S.schedule && !T.viol) refreshCheck();
}

/* ---------- ① ---------- */
function v1(m) {
  const yrs = []; for (let y = 2025; y <= 2031; y++) yrs.push(y);
  const first = m.dow[1];
  let cells = DOW.split('').map(c => `<div class="h">${c}</div>`).join('') + '<div></div>'.repeat(first);
  for (let d = 1; d <= m.days; d++) cells += `<div class="${m.holidays.has(d) ? 'hol' : m.dow[d] === 0 ? 'sun' : m.dow[d] === 6 ? 'sat' : ''}">${d}</div>`;
  const hol = Object.entries(m.holidayNames).map(([d, n]) => `${d}日 ${n}`).join('、');
  const rows = S.members.map((x, i) => `<div class="mem"><div class="row">
    <input type="text" class="grow" data-act="mname" data-id="${x.id}" value="${esc(x.name)}" aria-label="名前">
    <button class="sm" data-act="mup" data-id="${x.id}" aria-label="上へ">▲</button><button class="sm" data-act="mdown" data-id="${x.id}" aria-label="下へ">▼</button><button class="sm danger" data-act="mdel" data-id="${x.id}" aria-label="削除">✕</button></div>
    <div class="row" style="margin-top:6px"><select data-act="mkind" data-id="${x.id}"><option value="staff"${x.kind === 'staff' ? ' selected' : ''}>社員</option><option value="part"${x.kind === 'part' ? ' selected' : ''}>バイト</option></select>
    ${x.kind === 'staff' ? `<select data-act="mrole" data-id="${x.id}">${Object.entries(ROLE).map(([k, v]) => `<option value="${k}"${x.role === k ? ' selected' : ''}>${v}</option>`).join('')}</select>` : ''}
    ${x.kind === 'staff' && x.role === 'help' ? `<label class="row" style="gap:4px">月上限<input type="number" inputmode="numeric" min="0" max="31" style="width:64px" value="${x.maxPerMonth ?? 3}" data-act="mmax" data-id="${x.id}">回</label>` : ''}</div>
    ${x.kind === 'part' ? `<div class="row" style="margin-top:6px"><span class="mut">定番</span><input type="time" step="900" value="${fmtT(x.defaultStart)}" data-act="mdst" data-id="${x.id}" aria-label="定番の開始">〜<input type="time" step="900" value="${fmtT(x.defaultEnd)}" data-act="mden" data-id="${x.id}" aria-label="定番の終了">
      <select data-act="mdd" data-id="${x.id}"><option value="weekday"${(x.defaultDays || 'weekday') === 'weekday' ? ' selected' : ''}>平日のみ</option><option value="all"${x.defaultDays === 'all' ? ' selected' : ''}>毎日</option></select></div>` : ''}</div>`).join('');
  return `<section class="card"><h2>対象月</h2><div class="row">
    <select data-act="year">${yrs.map(y => `<option${y === S.year ? ' selected' : ''}>${y}</option>`).join('')}</select>年
    <select data-act="month">${[...Array(12)].map((_, i) => `<option${i + 1 === S.month ? ' selected' : ''}>${i + 1}</option>`).join('')}</select>月
    <span class="mut">${m.days}日間</span></div>
    <div class="cal">${cells}</div>
    <p class="mut">祝日：${hol || 'なし'}</p></section>
  <section class="card"><h2>メンバー</h2>${rows}
    <div class="btns"><button data-act="madd" data-kind="staff">＋社員を追加</button><button data-act="madd" data-kind="part">＋バイトを追加</button></div></section>
  <section class="card"><button class="danger" data-act="reset">入力内容をすべて消す</button></section>`;
}

/* ---------- ② ---------- */
const merge = (id, entries) => { S.partReq[id] = { ...(S.partReq[id] || {}), ...entries }; };
function findPart(g) {
  if (!g) return null;
  return parts().find(p => p.name === g) || parts().find(p => p.name.includes(g) || g.includes(famName(p.name))) || null;
}
function intake(P, text, owner) {
  const r = P.parseRequestText(text, S.year, S.month);
  const n = Object.keys(r.entries || {}).length;
  const g = findPart(P.guessSender(text, parts().map(p => p.name)) || r.name);
  T.raw[owner] = text;
  const note = k => (T.notes[owner] ||= []).push(k);
  if (!n) { note('日付を読み取れませんでした（テキストを確認してください）'); return; }
  if (g) { merge(g.id, r.entries); note(`${g.name}：${n}件${g.id !== owner && owner !== 'any' && owner !== 'paste' ? '（送信者を判定して割り当て）' : ''}`); }
  else if (owner !== 'any' && owner !== 'paste' && byId(owner)) { merge(owner, r.entries); note(`送信者を判定できなかったため ${byId(owner).name} に ${n}件`); }
  else T.pending.push({ entries: r.entries, n });
}
async function readImages(files, owner) {
  T.busy[owner] = true; render();
  try {
    const P = await lib.parse();
    for (let i = 0; i < files.length; i++) {
      const text = await P.ocrImage(files[i], p => { const e = $('#pg-' + owner); if (e) e.textContent = `${i + 1}/${files.length}枚目を読み取り中 ${Math.round(p * 100)}%`; });
      intake(P, text, owner);
    }
  } catch (e) { (T.notes[owner] ||= []).push('エラー：' + (e.message || e)); }
  delete T.busy[owner]; render();
}
async function readText(text) {
  try { intake(await lib.parse(), text, 'paste'); } catch (e) { (T.notes.paste ||= []).push('エラー：' + (e.message || e)); }
  render();
}
function entryRows(id) {
  const q = S.partReq[id] || {};
  return Object.keys(q).map(Number).sort((a, b) => a - b).map(d => {
    const e = q[d], all = e.allDay || (e.start == null && e.end == null);
    return `<div class="erow"><input type="number" inputmode="numeric" min="1" max="31" value="${d}" data-act="eday" data-id="${id}" data-day="${d}" aria-label="日">
      <input type="time" step="900" value="${all ? '' : fmtT(e.start)}" ${all ? 'disabled' : ''} data-act="est" data-id="${id}" data-day="${d}" aria-label="開始">
      <input type="time" step="900" value="${all ? '' : fmtT(e.end ?? 21.5)}" ${all ? 'disabled' : ''} data-act="eend" data-id="${id}" data-day="${d}" aria-label="終了">
      <label><input type="checkbox" ${all ? 'checked' : ''} data-act="eall" data-id="${id}" data-day="${d}">終日</label>
      <button class="sm danger" data-act="edel" data-id="${id}" data-day="${d}" aria-label="削除">✕</button></div>`;
  }).join('') || '<p class="mut">まだありません</p>';
}
function v2() {
  const fileBtn = (id, label) => `<label class="btn pri">${label}<input type="file" accept="image/*" multiple hidden data-act="pfile" data-id="${id}"></label>`;
  const cards = parts().map(p => {
    const n = Object.keys(S.partReq[p.id] || {}).length;
    return `<details class="card"><summary>${esc(p.name)} <span class="mut">${n ? n + '日分' : '未入力'}</span></summary>
      <div class="row">${fileBtn(p.id, 'スクショを選ぶ')}</div>
      ${T.busy[p.id] ? `<p class="mut" id="pg-${p.id}">読み取り中…</p>` : ''}
      ${(T.notes[p.id] || []).map(x => `<p class="mut">${esc(x)}</p>`).join('')}
      ${entryRows(p.id)}
      <div class="btns"><button data-act="eadd" data-id="${p.id}">＋日を追加</button>${n ? `<button class="danger" data-act="eclear" data-id="${p.id}">全消去</button>` : ''}</div>
      ${T.raw[p.id] ? `<details><summary class="mut">読み取った文字</summary><pre>${esc(T.raw[p.id])}</pre></details>` : ''}</details>`;
  }).join('');
  const pend = T.pending.map((q, i) => `<div class="card pend"><b>送信者を判定できませんでした</b><span class="mut">（${q.n}日分）</span>
    <p class="mut" style="margin:4px 0">誰の希望ですか？</p><div class="btns">${parts().map(p => `<button data-act="passign" data-i="${i}" data-id="${p.id}">${esc(p.name)}</button>`).join('')}<button class="danger" data-act="pdrop" data-i="${i}">破棄</button></div></div>`).join('');
  return `<p class="mut">画像を読み取ると、送信者を自動で判定して割り当てます。読み取り結果は必ず確認・修正してください。</p>
  <section class="card"><h2>まとめて読み込み</h2><div class="row">${fileBtn('any', 'スクショを選ぶ（複数可）')}</div>
    ${T.busy.any ? '<p class="mut" id="pg-any">読み取り中…</p>' : ''}${(T.notes.any || []).map(x => `<p class="mut">${esc(x)}</p>`).join('')}
    <h2 style="margin-top:14px">テキストを貼り付け</h2>
    <textarea id="pasteTa" placeholder="「テキスト認識表示」でコピーした文字を貼り付け"></textarea>
    <div class="btns"><button class="pri" data-act="ptext">読み取る</button></div>${(T.notes.paste || []).map(x => `<p class="mut">${esc(x)}</p>`).join('')}</section>
  ${pend}${cards}`;
}

/* ---------- ③ ---------- */
function grid(m, rowsHtml, marks) {
  const head = [...Array(m.days)].map((_, i) => { const d = i + 1; const mk = marks && (S.events[d] || []).length ? `<i class="evm" title="${esc(S.events[d].map(e => e.text).join('・'))}">●</i>` : '';
    return `<th class="${m.weekendOrHoliday.has(d) ? 'we' : ''}${m.dow[d] === 6 && !m.holidays.has(d) ? ' sat' : ''}">${d}${mk}<small>${DOW[m.dow[d]]}</small></th>`; }).join('');
  return `<div class="tbl-wrap"><table class="g"><tr><th class="nm"></th>${head}</tr>${rowsHtml}</table></div>`;
}
function v3(m) {
  const cal = T.cal;
  const colors = cal ? cal.colors.map(c => {
    const to = S.colorMap[c.key]?.to;
    return `<div class="card"><div class="row"><span class="sw" style="background:${esc(c.hex)}"></span><b>この色は誰ですか？</b></div><div class="btns">
       ${S.members.map(s => `<button class="${to === s.id ? 'sel' : ''}" data-act="cmap" data-key="${esc(c.key)}" data-hex="${esc(c.hex)}" data-to="${s.id}">${esc(famName(s.name))}</button>`).join('')}
      <button class="${to === 'event' ? 'sel' : ''}" data-act="cmap" data-key="${esc(c.key)}" data-hex="${esc(c.hex)}" data-to="event">店のイベント（対象外）</button></div></div>`;
  }).join('') : '';
  const bars = cal ? cal.bars.map((b, i) => {
    const to = S.colorMap[b.colorKey]?.to;
    if (!to) return '';
    const hex = cal.colors.find(c => c.key === b.colorKey)?.hex || '#ccc';
    const nm = to === 'event' ? '店のイベント' : famName(byId(to)?.name || '?'), isP = byId(to)?.kind === 'part';
    return `<div class="card bar"><div class="row"><span class="sw" style="background:${esc(hex)}"></span><b>${esc(nm)}</b><span class="mut">${ranges(b.days)}</span></div>
      ${b.kind === 'text' ? `<img src="${esc(b.cropDataUrl)}" alt="帯"><div class="row" style="margin-top:4px"><span class="mut">文字</span><input type="text" class="grow" value="${b.text == null ? '' : esc(b.text)}" placeholder="${b.text == null ? '読み取り中…' : '（なし）'}" data-act="btext" data-i="${i}" aria-label="帯の文字"></div>` : ''}
      ${isP ? `<p class="mut">${b.kind === 'x' ? '×の日は入りません' : '文字の帯は無視します'}。それ以外の日は定番時間で入ります。</p>` : to === 'event' ? (b.kind === 'x' ? '<p class="mut">×の帯は対象外</p>' : '') : `<div class="btns">${KINDS.map(([k, l]) => `<button class="${b.choice === k ? 'sel' : ''}" data-act="bchoice" data-i="${i}" data-k="${k}">${l}</button>`).join('')}</div>`}</div>`;
  }).join('') : '';
  const rows = staff().map(s => `<tr><td class="nm">${esc(famName(s.name))}</td>${[...Array(m.days)].map((_, i) => {
    const d = i + 1, v = S.staffEvents[s.id]?.[d] || '';
    return `<td class="c ${m.weekendOrHoliday.has(d) ? 'we ' : ''}${{ '×': 'x', 'リ': 'ri', '有': 'yu', MUST: 'must' }[v] || ''}" data-act="evc" data-id="${s.id}" data-day="${d}">${EVL[v]}</td>`;
  }).join('')}</tr>`).join('');
  return `<section class="card"><h2>カレンダーのスクショ</h2><div class="row"><label class="btn pri">画像を選ぶ（複数可）<input type="file" accept="image/*" multiple hidden data-act="cfile"></label></div>
    ${T.busy.cal ? `<p class="mut" id="pg-cal">${esc(T.busy.cal === true ? '解析中…' : T.busy.cal)}</p>` : ''}${T.err ? `<p class="err">${esc(T.err)}</p>` : ''}
    <p class="mut">帯の色から社員を判定します。答えた色は保存され、次回から自動で当てはまります。</p></section>
    ${colors}${bars}
    ${cal && bars ? '<div class="card"><button class="pri" data-act="bapply" style="width:100%">読み取り結果を下の表に反映</button></div>' : ''}
    <section class="card"><h2>日付 × 社員</h2><p class="mut">セルをタップで切り替え：なし→×→リ→有→MT（出勤マスト）</p>${grid(m, rows, true)}</section>
    ${evSection(m)}`;
}
const evDefault = t => BUSY_RE.test(t || '');
const evFlat = () => { const o = {}; for (const [d, a] of Object.entries(S.events)) if (a.length) o[d] = a.map(({ src, bset, ...e }) => e); return o; };
function evSection(m) {
  const rows = Object.keys(S.events).map(Number).sort((a, b) => a - b).flatMap(d => S.events[d].map((e, i) => `<div class="erow"><span>${d}日</span>
    <input type="text" class="grow" value="${esc(e.text)}" data-act="evtext" data-day="${d}" data-i="${i}" aria-label="イベント名">
    <span class="mut">${e.kind === 'staff' ? esc(famName(byId(e.staffId)?.name || '社員')) : '店'}</span>
    <button class="sm ${e.busy ? 'sel' : ''}" data-act="evbusy" data-day="${d}" data-i="${i}">繁忙${e.busy ? '○' : '×'}</button>
    <button class="sm danger" data-act="evdel" data-day="${d}" data-i="${i}" aria-label="削除">✕</button></div>`)).join('') || '<p class="mut">まだありません</p>';
  return `<section class="card"><h2>イベント</h2><p class="mut">店のイベントは繁忙にすると目標人数が+1。繁忙は文字（セール・フェア等）から自動判定、タップで切り替え。</p>${rows}
    <div class="row" style="margin-top:8px"><input type="number" inputmode="numeric" min="1" max="${m.days}" placeholder="日" style="width:64px" id="evNewDay">
    <input type="text" class="grow" placeholder="イベント名" id="evNewText">
    <select id="evNewKind"><option value="store">店・ビル</option>${staff().map(s => `<option value="staff:${s.id}">${esc(famName(s.name))}</option>`).join('')}</select>
    <button class="pri" data-act="evadd">追加</button></div></section>`;
}
async function readCalendars(files) {
  T.busy.cal = true; T.err = ''; render();
  try {
    const C = await lib.cal();
    for (let i = 0; i < files.length; i++) {
      T.busy.cal = `${i + 1}/${files.length}枚目を解析中…`; const e = $('#pg-cal'); if (e) e.textContent = T.busy.cal;
      const r = await C.analyzeCalendar(files[i], mon());
      T.cal ||= { colors: [], bars: [] };
      const map = {};
      for (const c of r.colors) {
        const ex = T.cal.colors.find(x => x.key === c.key || cdist(x.hex, c.hex) < 30);
        map[c.key] = ex ? ex.key : c.key;
        if (!ex) T.cal.colors.push(c);
        if (!S.colorMap[map[c.key]]) {
          const near = Object.values(S.colorMap).filter(v => v.hex && cdist(v.hex, c.hex) < 45).sort((a, b) => cdist(a.hex, c.hex) - cdist(b.hex, c.hex))[0];
          if (near) S.colorMap[map[c.key]] = { hex: c.hex, to: near.to };
        }
      }
      for (const b of r.bars) T.cal.bars.push({ ...b, colorKey: map[b.colorKey] || b.colorKey, choice: b.kind === 'x' ? '×' : undefined, text: b.kind === 'text' ? null : undefined });
    }
    T.busy.cal = false; render();
    for (const b of T.cal.bars.filter(b => b.kind === 'text' && b.text === null)) {
      try { b.text = (await C.ocrBar(b.cropDataUrl)) || ''; } catch { b.text = ''; }
      if (/有給|有休/.test(b.text)) b.choice = '有';
      else if (/MT|ミーティング|会議|フォロー/i.test(b.text)) b.choice = 'MUST';
      else if (/リフ/.test(b.text)) b.choice = 'リ';
      if (S.step === 2) render();
    }
  } catch (e) { T.err = '解析に失敗しました：' + (e.message || e); T.busy.cal = false; render(); }
}

/* ---------- ④ ---------- */
function v4() {
  return `<p class="mut">月末時点の連勤数と、最終日のシフトを入れてください。</p>` + staff().map(s => {
    const p = S.prev[s.id] || { run: 0, last: '/' };
    return `<div class="card"><div class="row"><b class="grow">${esc(s.name)}</b>
      <label class="row" style="gap:4px">連勤<input type="number" inputmode="numeric" min="0" max="9" style="width:64px" value="${p.run}" data-act="prun" data-id="${s.id}"></label>
      <label class="row" style="gap:4px">最終<select data-act="plast" data-id="${s.id}">${CODES.map(c => `<option${c === p.last ? ' selected' : ''}>${c}</option>`).join('')}</select></label></div></div>`;
  }).join('');
}

/* ---------- ⑤ ---------- */
function fullPrev() { const o = {}; for (const s of staff()) o[s.id] = S.prev[s.id] || { run: 0, last: '/' }; return o; }
function solveInput() {
  const m = mon(), pr = {};
  for (const p of parts()) {
    const q = { ...(S.partReq[p.id] || {}) };
    if (S.partDef[p.id] && p.defaultStart != null && p.defaultEnd != null)
      for (let d = 1; d <= m.days; d++)
        if (!q[d] && !S.partOff[p.id]?.[d] && (p.defaultDays === 'all' || !m.weekendOrHoliday.has(d))) q[d] = { start: p.defaultStart, end: p.defaultEnd };
    pr[p.id] = q;
  }
  const ev = {}; for (const s of staff()) ev[s.id] = Object.fromEntries(Object.entries(S.staffEvents[s.id] || {}).filter(([, v]) => v));
  return { month: mon(), members: S.members, partRequests: pr, staffEvents: ev, events: evFlat(), prev: fullPrev(), locked: S.schedule?.locked || {} };
}
async function refreshCheck() {
  if (!S.schedule) return;
  try {
    T.sv ||= await lib.solver();
    const inp = solveInput();
    T.viol = T.sv.checkViolations(inp, S.schedule) || [];
    T.head = T.sv.headcount(inp, S.schedule) || {};
    try { T.help = T.sv.suggestHelp ? T.sv.suggestHelp(inp, S.schedule) || {} : {}; } catch { T.help = {}; }
  } catch (e) { T.viol = []; T.head = {}; T.err = String(e.message || e); }
  if (S.step === 4) render();
}
async function gen() {
  T.gen = true; T.err = ''; render();
  try {
    T.sv ||= await lib.solver();
    const inp = solveInput();
    const sch = await T.sv.solve(inp, p => { const e = $('#genbar'); if (e) e.style.width = Math.round(p * 100) + '%'; const t = $('#genpct'); if (t) t.textContent = Math.round(p * 100) + '%'; });
    sch.locked = { ...(inp.locked || {}) };
    S.schedule = sch; T.help = null;
  } catch (e) { T.err = '生成に失敗しました：' + (e.message || e); }
  T.gen = false; T.viol = null; render();
}
function v5(m) {
  const sch = S.schedule;
  const bt = `<section class="card"><button class="pri" style="width:100%" data-act="gen" ${T.gen ? 'disabled' : ''}>${sch ? 'もう一度生成（手直ししたセルは固定）' : 'シフトを生成'}</button>
    ${T.gen ? '<div class="prog"><div id="genbar"></div></div><p class="mut" id="genpct">0%</p>' : ''}${T.err ? `<p class="err">${esc(T.err)}</p>` : ''}</section>`;
  if (!sch) return bt;
  const rows = staff().map(s => `<tr><td class="nm">${esc(famName(s.name))}</td>${[...Array(m.days)].map((_, i) => {
    const d = i + 1, v = sch.staff?.[s.id]?.[d] || '', lk = sch.locked?.[s.id]?.[d];
    const cl = v === '/' ? 'rest' : v === '×' ? 'x' : v === 'リ' ? 'ri' : v === '有' ? 'yu' : v === 'B15' ? 'b15' : '';
    return `<td class="c ${m.weekendOrHoliday.has(d) && !cl ? 'we ' : ''}${cl}${lk ? ' lk' : ''}" data-act="cell" data-id="${s.id}" data-day="${d}">${v === '/' ? '休' : esc(v)}</td>`;
  }).join('')}</tr>`).join('') +
    parts().map(p => `<tr><td class="nm">${esc(famName(p.name))}</td>${[...Array(m.days)].map((_, i) => {
      const d = i + 1, t = sch.part?.[p.id]?.[d];
      return `<td class="pt ${m.weekendOrHoliday.has(d) ? 'we' : ''}">${t ? `${shortT(t[0])}<br>${shortT(t[1])}` : ''}</td>`;
    }).join('')}</tr>`).join('') +
    `<tr><td class="nm">人数</td>${[...Array(m.days)].map((_, i) => { const h = T.head?.[i + 1]; return `<td class="${h != null && h <= 2.5 ? 'low' : ''}">${h ?? ''}</td>`; }).join('')}</tr>`;
  const viol = T.viol ? (T.viol.length ? T.viol.map(v => `<div class="viol ${v.level}"><b>${v.level === 'error' ? '違反' : '注意'}</b> ${v.day ? v.day + '日 ' : ''}${v.staffId ? esc(famName(byId(v.staffId)?.name || '')) + ' ' : ''}${esc(v.msg)}</div>`).join('') : '<p class="ok">ルール違反はありません</p>') : '<p class="mut">チェック中…</p>';
  return `${bt}<section class="card"><h2>シフト表</h2><p class="mut">社員のセルをタップで記号を変更（枠線＝固定）。人数が2.5以下の日は赤。</p>${grid(m, rows, true)}</section>
    ${helpCard()}
    <section class="card"><h2>違反・注意</h2>${viol}</section>`;
}
function helpCard() {
  const h = T.help || {}, days = Object.keys(h).map(Number).filter(d => h[d]?.length).sort((a, b) => a - b);
  return `<section class="card"><h2>応援候補</h2>${days.length ? days.map(d => `<div class="viol warn"><b>${d}日（${DOW[mon().dow[d]]}）</b> ${h[d].map(x => `${esc(x.name)}<span class="mut">（${esc(x.reason)}）</span>`).join('、')}</div>`).join('') : '<p class="mut">人員が足りない日はありません</p>'}</section>`;
}
function renderSheet() {
  const el = $('#sheet');
  if (!T.sheet) { el.className = ''; el.innerHTML = ''; return; }
  const { id, day } = T.sheet, cur = S.schedule?.staff?.[id]?.[day];
  el.className = 'on';
  el.innerHTML = `<div class="in"><h2>${esc(byId(id)?.name || '')}　${S.month}/${day}（${DOW[mon().dow[day]]}）</h2>
    <div class="btns">${CODES.map(c => `<button class="${c === cur ? 'sel' : ''}" data-act="setcode" data-c="${c}">${c === '/' ? '公休' : c}</button>`).join('')}</div>
    <div class="btns"><button data-act="unlock">固定を解除</button><button data-act="closesheet">閉じる</button></div></div>`;
}

/* ---------- ⑥ ---------- */
const fname = () => `${S.month}月シフト表.xlsx`;
function mustDays() {
  const o = {};
  for (const s of staff()) { const d = Object.entries(S.staffEvents[s.id] || {}).filter(([, v]) => v === 'MUST').map(([k]) => +k); if (d.length) o[s.id] = d; }
  return o;
}
function v6() {
  if (!S.schedule) return '<section class="card"><p>先に⑤でシフトを生成してください。</p></section>';
  return `<section class="card"><h2>${esc(fname())}</h2><p class="mut">テンプレートに書き込んだxlsxを作成します。予算欄は空のままです。</p>
    <button class="pri" style="width:100%" data-act="build" ${T.busy.x ? 'disabled' : ''}>${T.busy.x ? '作成中…' : 'xlsxを作成'}</button>
    ${T.msg ? `<p class="${T.msg.startsWith('エラー') ? 'err' : 'ok'}">${esc(T.msg)}</p>` : ''}
    ${T.file ? `<div class="btns"><button class="pri" data-act="share">共有・保存</button><button data-act="dl">ダウンロード</button></div>` : ''}</section>`;
}
async function build() {
  T.busy.x = true; T.msg = ''; T.file = null; render();
  try {
    if (!T.help) await refreshCheck();
    const X = await lib.xlsx();
    const blob = await X.buildXlsx({ month: mon(), members: S.members, schedule: S.schedule, mustDays: mustDays(), events: evFlat(), help: T.help || {}, templateUrl: 'template.xlsx' });
    T.file = new File([blob], fname(), { type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    T.msg = `作成しました（${Math.round(T.file.size / 1024)}KB）`;
  } catch (e) { T.msg = 'エラー：' + (e.message || e); }
  delete T.busy.x; render();
}
function download() {
  const a = document.createElement('a'); a.href = URL.createObjectURL(T.file); a.download = T.file.name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}
async function share() {
  try {
    if (navigator.canShare && navigator.canShare({ files: [T.file] })) { await navigator.share({ files: [T.file], title: T.file.name }); return; }
  } catch (e) { if (e.name === 'AbortError') return; }
  download();
}

/* ---------- events ---------- */
function clearMonthData() { S.partReq = {}; S.staffEvents = {}; S.events = {}; S.partOff = {}; S.partDef = {}; T.help = null; S.prev = {}; S.schedule = null; T.cal = null; T.viol = null; T.head = null; T.notes = {}; T.raw = {}; T.pending = []; T.file = null; }
const hasData = () => Object.keys(S.partReq).length || Object.keys(S.events).length || Object.keys(S.staffEvents).length || S.schedule;
function swapMember(id, dir) {
  const i = S.members.findIndex(m => m.id === id), k = S.members[i].kind;
  let j = i + dir; while (S.members[j] && S.members[j].kind !== k) j += dir;
  if (S.members[j]) [S.members[i], S.members[j]] = [S.members[j], S.members[i]];
}
function setEntry(id, day, f) { const e = S.partReq[id]?.[day]; if (e) f(e); }

function act(el) {
  const d = el.dataset, v = el.type === 'checkbox' ? el.checked : el.value, id = d.id, day = +d.day;
  switch (d.act) {
    case 'go': S.step = Math.max(0, Math.min(5, +d.i)); window.scrollTo(0, 0); break;
    case 'year': case 'month': {
      const ny = d.act === 'year' ? +v : S.year, nm = d.act === 'month' ? +v : S.month;
      if (hasData() && !confirm('対象月を変えると、読み込み済みの希望・予定・生成結果が消えます。よろしいですか？')) break;
      if (hasData()) clearMonthData();
      S.year = ny; S.month = nm; break;
    }
    case 'mname': byId(id).name = v.trim() || byId(id).name; break;
    case 'mkind': { const m = byId(id); m.kind = v; if (v === 'staff') m.role = m.role || 'flex'; else delete m.role; break; }
    case 'mrole': byId(id).role = v; break;
    case 'mup': swapMember(id, -1); break;
    case 'mdown': swapMember(id, 1); break;
    case 'mdel': {
      if (!confirm(`${byId(id).name} を削除しますか？`)) return;
      S.members = S.members.filter(m => m.id !== id);
      delete S.partReq[id]; delete S.staffEvents[id]; delete S.prev[id]; delete S.partOff[id]; delete S.partDef[id];
      if (S.schedule) { S.schedule = null; T.viol = null; }
      break;
    }
    case 'madd': S.members.push({ id: 'm' + Date.now().toString(36), name: d.kind === 'staff' ? '新しい社員' : '新しいバイト', kind: d.kind, ...(d.kind === 'staff' ? { role: 'flex' } : {}) }); break;
    case 'reset': if (!confirm('入力内容をすべて消します。メンバーと色の対応は残ります。')) return; clearMonthData(); S.step = 0; break;
    case 'pfile': { const f = [...el.files]; el.value = ''; if (f.length) readImages(f, id); return; }
    case 'ptext': { const t = $('#pasteTa').value.trim(); if (t) { T.notes.paste = []; readText(t); } return; }
    case 'passign': { const q = T.pending[+d.i]; merge(id, q.entries); T.pending.splice(+d.i, 1); break; }
    case 'pdrop': T.pending.splice(+d.i, 1); break;
    case 'eday': {
      const n = +v, q = S.partReq[id];
      if (n >= 1 && n <= mon().days && !q[n]) { q[n] = q[day]; delete q[day]; }
      break;
    }
    case 'est': setEntry(id, day, e => { e.start = parseT(v); e.allDay = false; if (e.end == null) e.end = 21.5; }); break;
    case 'eend': setEntry(id, day, e => { e.end = parseT(v); }); break;
    case 'eall': setEntry(id, day, e => { if (v) { e.start = null; e.end = null; e.allDay = true; } else { e.start = 17; e.end = 21.5; e.allDay = false; } }); break;
    case 'edel': delete S.partReq[id][day]; break;
    case 'eadd': {
      const q = (S.partReq[id] ||= {}); let n = 1; while (q[n] && n < mon().days) n++;
      q[n] = { start: 17, end: 21.5 }; break;
    }
    case 'eclear': if (!confirm('この人の希望を全部消しますか？')) return; delete S.partReq[id]; break;
    case 'cfile': { const f = [...el.files]; el.value = ''; if (f.length) readCalendars(f); return; }
    case 'cmap': S.colorMap[d.key] = { hex: d.hex, to: d.to }; break;
    case 'bchoice': T.cal.bars[+d.i].choice = d.k; break;
    case 'bapply': {
      for (const d of Object.keys(S.events)) { S.events[d] = S.events[d].filter(e => e.src !== 'cal'); if (!S.events[d].length) delete S.events[d]; }
      const addEv = (day, e) => { const a = (S.events[day] ||= []); if (!a.some(x => x.text === e.text && x.kind === e.kind && x.staffId === e.staffId)) a.push({ ...e, src: 'cal' }); };
      for (const b of T.cal.bars) {
        const to = S.colorMap[b.colorKey]?.to, mem = byId(to), text = (b.text || '').trim();
        if (!to) continue;
        if (mem?.kind === 'part') {
          S.partDef[to] = true;
          if (b.kind === 'x') { const o = (S.partOff[to] ||= {}); for (const x of b.days) o[x] = true; }
        } else if (to === 'event') {
          if (b.kind === 'text' && text) for (const x of b.days) addEv(x, { text, kind: 'store', busy: evDefault(text) });
        } else if (mem && b.choice && b.choice !== '無視') {
          const ev = (S.staffEvents[to] ||= {}); for (const x of b.days) ev[x] = b.choice;
          if (b.kind === 'text' && text) for (const x of b.days) addEv(x, { text, kind: 'staff', staffId: to, busy: false });
        }
      }
      break;
    }
    case 'btext': T.cal.bars[+d.i].text = v; return;
    case 'evtext': { const e = S.events[day][+d.i]; e.text = v.trim(); if (!e.bset && e.kind === 'store') e.busy = evDefault(e.text); delete e.src; break; }
    case 'evbusy': { const e = S.events[day][+d.i]; e.busy = !e.busy; e.bset = true; break; }
    case 'evdel': S.events[day].splice(+d.i, 1); if (!S.events[day].length) delete S.events[day]; break;
    case 'evadd': {
      const dd = parseInt($('#evNewDay').value), tx = $('#evNewText').value.trim(), k = $('#evNewKind').value;
      if (!(dd >= 1 && dd <= mon().days) || !tx) return;
      const [kind, sid] = k.split(':');
      (S.events[dd] ||= []).push({ text: tx, kind, ...(sid ? { staffId: sid } : {}), busy: kind === 'store' && evDefault(tx) });
      break;
    }
    case 'mmax': byId(id).maxPerMonth = Math.max(0, parseInt(v) || 0); break;
    case 'mdst': { const m = byId(id); m.defaultStart = parseT(v); if (m.defaultStart == null) delete m.defaultStart; break; }
    case 'mden': { const m = byId(id); m.defaultEnd = parseT(v); if (m.defaultEnd == null) delete m.defaultEnd; break; }
    case 'mdd': byId(id).defaultDays = v; break;
    case 'evc': {
      const ev = (S.staffEvents[id] ||= {}), cur = ev[day] || '';
      const nx = EV[(EV.indexOf(cur) + 1) % EV.length];
      if (nx) ev[day] = nx; else delete ev[day];
      break;
    }
    case 'prun': (S.prev[id] ||= { run: 0, last: '/' }).run = Math.max(0, Math.min(9, parseInt(v) || 0)); break;
    case 'plast': (S.prev[id] ||= { run: 0, last: '/' }).last = v; break;
    case 'gen': gen(); return;
    case 'cell': T.sheet = { id, day }; break;
    case 'closesheet': T.sheet = null; break;
    case 'setcode': {
      const { id: sid, day: sd } = T.sheet, s = S.schedule;
      (s.staff[sid] ||= {})[sd] = d.c; ((s.locked ||= {})[sid] ||= {})[sd] = true;
      T.sheet = null; T.viol = null; break;
    }
    case 'unlock': { const { id: sid, day: sd } = T.sheet; if (S.schedule.locked?.[sid]) delete S.schedule.locked[sid][sd]; T.sheet = null; break; }
    case 'build': build(); return;
    case 'share': share(); return;
    case 'dl': download(); return;
  }
  render();
}

const isField = e => /^(INPUT|SELECT|TEXTAREA)$/.test(e.tagName);
document.addEventListener('click', e => {
  const el = e.target.closest('[data-act]');
  if (el && !isField(el)) act(el);
  else if (e.target.id === 'sheet') { T.sheet = null; renderSheet(); }
});
document.addEventListener('change', e => { const el = e.target.closest('[data-act]'); if (el && isField(el)) act(el); });

render();
