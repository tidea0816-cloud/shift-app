const CODES = ['A', 'D', 'B', '/', '×', 'リ', '有'];
const codeOf = (s) => {
  if (!s) return 3;
  const i = CODES.indexOf(s[0]);
  return i < 0 ? 3 : i;
};
const isWork = (s) => !!s && 'ADB'.includes(s[0]);
const HELP_MAX = 3;
// 仮決めの設定（変えやすいよう定数化）
export const RULES = {
  managerTitle: '店長', subTitle: '副店長',
  firstDays: [1],            // 月初＝1日のみ
  endRunMax: 4, endRunGoal: 2, endRunPen: 15,   // 店長の月末連勤
  weOffPen: 20, managerWeOffMult: 2,            // 土日祝の公休ペナルティ（店長は倍）
  splitMin: 4, splitPen: 80,                    // 休み1日を挟んだ長い連勤
  aOnOpenerPen: 60, extraAPen: 15, aOverPen: 6,
  drillStart: 8.25,
};
const isMgrM = (m) => m.title === RULES.managerTitle && m.role !== 'help';
const BUSY_RE = /セール|フェア|祭|初売|福袋|催事|イベント/;
const helpMax = (m) => m.maxPerMonth ?? HELP_MAX;
const isBusy = (input, d) => (input.events?.[d] || []).some((e) => e.busy ?? BUSY_RE.test(e.text || ''));

export function targets(input) {
  const N = input.month.days, we = input.month.weekendOrHoliday, TG = new Float64Array(N + 2);
  for (let d = 1; d <= N; d++) {
    let a = d, b = d;
    while (a > 1 && we.has(a - 1)) a--;
    while (b < N && we.has(b + 1)) b++;
    TG[d] = (!we.has(d) ? 3 : (b - a + 1 >= 3 && d !== b ? 5 : 4)) + (isBusy(input, d) ? 1 : 0);
  }
  return TG;
}

// イベントの参加者：{ must: Map(staffId -> イベント名), drill: Set(staffId) }
export function eventPeople(input, d) {
  const must = new Map(), drill = new Set();
  const staff = (input.members || []).filter((m) => m.kind === 'staff');
  for (const e of input.events?.[d] || []) {
    if (!e.type || e.type === 'info') continue;
    const ids = new Set(e.staffIds || []);
    if (e.staffId) ids.add(e.staffId);
    const title = e.type === 'storeMT' ? RULES.managerTitle : e.type === 'subMT' ? RULES.subTitle : null;
    if (title) for (const m of staff) if (m.title === title) ids.add(m.id);
    for (const id of ids) {
      if (!staff.some((m) => m.id === id)) continue;
      if (!must.has(id)) must.set(id, e.text || '');
      if (e.type === 'drill') drill.add(id);
    }
  }
  return { must, drill };
}

function partTimes(r) {
  if (!r) return null;
  if (r.allDay) return [12.5, 21.5];
  if (r.start != null) return [r.start, r.end != null ? r.end : 21.5];
  return null;
}

function buildPart(input) {
  const part = {}, tent = {};
  for (const m of input.members) {
    if (m.kind !== 'part') continue;
    const req = input.partRequests?.[m.id] || {};
    part[m.id] = {}; tent[m.id] = {};
    for (const d of Object.keys(req)) {
      const t = partTimes(req[d]);
      if (!t) continue;
      if (req[d].tentative) tent[m.id][d] = t; else part[m.id][d] = t;
    }
  }
  return { part, tent };
}

export function solveSync(input, onProgress, opts = {}) {
  const { month, members } = input;
  const N = month.days;
  const we = new Uint8Array(N + 2);
  for (const d of month.weekendOrHoliday) we[d] = 1;
  const staff = members.filter((m) => m.kind === 'staff');
  const nS = staff.length;
  const { part, tent } = buildPart(input);
  const pt = new Float64Array(N + 2), opener = new Uint8Array(N + 2);
  for (const id in part) for (const d in part[id]) {
    pt[d] += 0.5;
    if (part[id][d][0] <= 10 && part[id][d][1] >= 13) opener[d] = 1;
  }
  const TG = targets(input), bz = new Uint8Array(N + 2);
  for (let d = 1; d <= N; d++) bz[d] = isBusy(input, d) ? 1 : 0;
  const ppl = [];
  for (let d = 0; d <= N; d++) ppl[d] = d ? eventPeople(input, d) : null;
  const aNeed = N - opener.reduce((a, b) => a + b, 0);
  const nonHelp = staff.filter((m) => m.role !== 'help');
  const wA = (m) => (m.role === 'B' ? 0.6 : 1), wSum = nonHelp.reduce((a, m) => a + wA(m), 0) || 1;
  const isMgr = [], splitOn = [];
  const R = [], fixedMask = [], fixedStr = [], must = [], prevRun = [], prevLast = [], tgt = [], isRoleA = [], isHelp = [], freeDays = [], allowHelp = new Uint8Array(N + 2);
  const ev = input.staffEvents || {};
  for (const m of staff) for (const d in (ev[m.id] || {})) if (m.role !== 'help' && 'リ×有'.includes(ev[m.id][d])) allowHelp[d] = 1;
  const lockedIn = input.locked || {};
  const base = input.base?.staff || {};
  staff.forEach((m, i) => {
    const row = new Int8Array(N + 2), fx = new Uint8Array(N + 2), mu = new Uint8Array(N + 2), fs = {};
    const help = m.role === 'help';
    isHelp[i] = help; isRoleA[i] = m.role === 'A'; isMgr[i] = isMgrM(m); splitOn[i] = !help && !!m.avoidSplitLongRuns;
    let fixedOther = 0;
    for (let d = 1; d <= N; d++) {
      const e = (ev[m.id] || {})[d];
      if (e === 'MUST' || ppl[d].must.has(m.id)) mu[d] = 1;
      if (isMgr[i] && RULES.firstDays.includes(d)) mu[d] = 1;
      if (e && e !== 'MUST') {
        fx[d] = 1; row[d] = codeOf(e);
        if (e === 'リ' || e === '有') fixedOther++;
      } else if (help && !allowHelp[d]) { fx[d] = 1; row[d] = 3; }
      if (ppl[d].drill.has(m.id) && !(e && e !== 'MUST')) { fx[d] = 1; row[d] = 0; fs[d] = 'A815'; }
      if (lockedIn[m.id]?.[d] && base[m.id]?.[d]) {
        fx[d] = 1; row[d] = codeOf(base[m.id][d]); fs[d] = base[m.id][d];
      }
    }
    const fd = [];
    for (let d = 1; d <= N; d++) if (!fx[d]) fd.push(d);
    const work = Math.max(0, N - 8 - fixedOther);
    let ta, td, tb;
    // 社員は遅番寄り：Aは「開店にバイトがいない日」の分だけを社員で分け合う
    ta = help ? 0 : Math.round(aNeed * wA(m) / wSum);
    td = Math.round(work * (m.role === 'B' ? 0.3 : 0.25));
    tb = Math.max(0, work - ta - td);
    const pv = input.prev?.[m.id] || { run: 0, last: '/' };
    R.push(row); fixedMask.push(fx); fixedStr.push(fs); must.push(mu);
    prevRun.push(pv.run || 0); prevLast.push(codeOf(pv.last)); tgt.push([ta, td, tb]); freeDays.push(fd);
  });

  const score = () => {
    let c = 0;
    for (let i = 0; i < nS; i++) {
      const row = R[i];
      if (isHelp[i]) {
        let nb = 0;
        for (let d = 1; d <= N; d++) if (row[d] === 2) nb++;
        c += 20 * nb + 1000 * Math.max(0, nb - helpMax(staff[i]));
        continue;
      }
      let run = prevRun[i], last = prevLast[i], offs = 0, cA = 0, cD = 0, cB = 0, two = 0;
      let segBefore = 0, gap = 0, curGap = 0;
      const mu = must[i], sp = splitOn[i];
      for (let d = 1; d <= N; d++) {
        const x = row[d];
        if (x >= 3) {
          if (sp) { if (run > 0) { segBefore = run; gap = 1; } else gap++; }
          run = 0;
          if (x <= 4) { offs++; if (we[d] && x === 3) c += RULES.weOffPen * (isMgr[i] ? RULES.managerWeOffMult : 1); }
          if (d > 1 && row[d - 1] >= 3 && !(x >= 5 && row[d - 1] >= 5)) two++;
          if (mu[d]) c += 1000;
        } else {
          run++;
          if (sp) { if (run === 1) curGap = gap; if (run === RULES.splitMin && curGap === 1 && segBefore >= RULES.splitMin) c += RULES.splitPen; }
          if (run > 5) c += 1000; else if (run === 5) c += 30;
          if (last === 2 && x === 0) c += 1000;
          if (x === 0) cA++; else if (x === 1) cD++; else cB++;
        }
        last = x;
      }
      const t = tgt[i];
      c += 1000 * Math.abs(offs - 8) + 3 * (Math.abs(cA - t[0]) + Math.abs(cD - t[1]) + Math.abs(cB - t[2]));
      c += 8 * Math.max(0, two - 1) + RULES.aOverPen * Math.max(0, cA - t[0]);
      if (isMgr[i]) { if (run > RULES.endRunMax) c += 1000 * (run - RULES.endRunMax); else if (run > RULES.endRunGoal) c += RULES.endRunPen * (run - RULES.endRunGoal); }
      if (!isRoleA[i] && staff[i].role === 'B') c += 30 * Math.max(0, cA - 4);
    }
    for (let d = 1; d <= N; d++) {
      let nA = 0, nD = 0, nB = 0, bReg = 0, aNon = 0;
      for (let i = 0; i < nS; i++) {
        const x = R[i][d];
        if (x === 0) { nA++; if (!isRoleA[i]) aNon++; }
        else if (x === 1) nD++;
        else if (x === 2) { nB++; if (!isHelp[i]) bReg++; }
      }
      const nW = nA + nD + nB;
      if (bReg === 0) c += 1000;
      if (we[d]) {
        if (nW < 2) c += 1000;
        if (nA === 0 && !opener[d]) c += 60;
        if (nW === 3 && !((nA === 1 && nD === 1) || (nA === 0 && nD === 1) || (nA === 1 && nD === 0))) c += 10;
      } else {
        if (!opener[d] && nA === 0) c += 1000;

        if (nW === 2 && !(nD === 1 && nB === 1)) c += 6;
        if (nW === 3) c += 15;
      }
      if (opener[d]) c += RULES.aOnOpenerPen * nA;
      if (nA > 1) c += RULES.extraAPen * (nA - 1);
      const n = nW + pt[d], t = TG[d];
      c += (we[d] ? 40 : 30) * (bz[d] ? 3 : 1) * Math.max(0, t - n) + 10 * Math.max(0, n - t - 1);
      if (n <= 2) c += 150; else if (n <= 2.5) c += 40;
    }
    return c;
  };

  const trials = opts.trials || 5, iters = opts.iters || 80000, budget = opts.budgetMs || 13000;
  const t0 = Date.now();
  const free = [];
  for (let i = 0; i < nS; i++) for (const d of freeDays[i]) free.push([i, d]);
  const rnd = Math.random;
  let best = null;
  const pool = [0, 1, 2, 3];
  for (let tr = 0; tr < trials && free.length; tr++) {
    for (let i = 0; i < nS; i++) for (const d of freeDays[i]) R[i][d] = isHelp[i] ? 3 : [0, 1, 2, 2, 3][(rnd() * 5) | 0];
    let cur = score(), T = 50;
    const cool = Math.pow(0.3 / 50, 1 / (iters * 0.7));
    let it = 0;
    for (; it < iters; it++) {
      if ((it & 1023) === 0) {
        if (onProgress) onProgress(Math.min(1, (tr + it / iters) / trials));
        if (Date.now() - t0 > budget && best) break;
      }
      const [i, d] = free[(rnd() * free.length) | 0];
      const row = R[i], old = row[d];
      let d2 = 0, o2 = 0;
      if (isHelp[i]) row[d] = old === 3 ? 2 : 3;
      else if (rnd() < 0.3) {
        const fd = freeDays[i];
        d2 = fd[(rnd() * fd.length) | 0]; o2 = row[d2];
        row[d] = o2; row[d2] = old;
      } else {
        let v = pool[(rnd() * 4) | 0];
        if (v === old) v = pool[(v + 1 + ((rnd() * 3) | 0)) & 3];
        row[d] = v;
      }
      const n = score();
      if (n <= cur || rnd() < Math.exp((cur - n) / T)) cur = n;
      else { row[d] = old; if (d2) row[d2] = o2; }
      if (T > 0.3) T *= cool;
    }
    if (!best || cur < best.cur) best = { cur, R: R.map((r) => Int8Array.from(r)) };
    if (Date.now() - t0 > budget) break;
  }
  const out = { staff: {}, part, locked: input.locked || {}, tentative: {} };
  const fin = best ? best.R : R;
  staff.forEach((m, i) => {
    const o = {};
    for (let d = 1; d <= N; d++) o[d] = fixedStr[i][d] || CODES[fin[i][d]];
    out.staff[m.id] = o;
  });
  // △（人手不足のときだけ可）：目標人数に届かない日だけ入れる
  for (let d = 1; d <= N; d++) {
    let n = pt[d];
    for (let i = 0; i < nS; i++) if (fin[i][d] <= 2) n++;
    for (const id in tent) {
      if (n >= TG[d]) break;
      const t = tent[id][d];
      if (!t || part[id][d]) continue;
      part[id][d] = t; (out.tentative[id] ||= {})[d] = true; n += 0.5;
    }
  }
  if (onProgress) onProgress(1);
  out.score = best ? best.cur : null;
  return out;
}

export function solve(input, onProgress) {
  return new Promise((resolve) => {
    let w;
    const fallback = () => setTimeout(() => resolve(solveSync(input, onProgress)), 20);
    try {
      w = new Worker(new URL('./solverWorker.js', import.meta.url), { type: 'module' });
    } catch (e) { return fallback(); }
    w.onmessage = (e) => {
      if (e.data.type === 'progress') onProgress && onProgress(e.data.v);
      else if (e.data.type === 'done') { w.terminate(); resolve(e.data.schedule); }
      else { w.terminate(); fallback(); }
    };
    w.onerror = () => { w.terminate(); fallback(); };
    try { w.postMessage({ input }); } catch (e) { w.terminate(); fallback(); }
  });
}

export function headcount(input, schedule) {
  const r = {};
  for (let d = 1; d <= input.month.days; d++) {
    let n = 0;
    for (const m of input.members) {
      if (m.kind === 'staff') { if (isWork(schedule.staff?.[m.id]?.[d])) n += 1; }
      else if (schedule.part?.[m.id]?.[d]) n += 0.5;
    }
    r[d] = n;
  }
  return r;
}

export function checkViolations(input, schedule) {
  const N = input.month.days, v = [];
  const nm = (m) => m.name || m.id;
  const staff = input.members.filter((m) => m.kind === 'staff');
  const cell = (id, d) => schedule.staff?.[id]?.[d] || '/';
  for (let d = 1; d <= N; d++) {
    if (!staff.some((m) => cell(m.id, d)[0] === 'B')) v.push({ day: d, level: 'error', msg: `${d}日 締めのBが不在` });
  }
  const isOpener = (d) => Object.values(schedule.part || {}).some((o) => o?.[d] && o[d][0] <= 10 && o[d][1] >= 13);
  for (let d = 1; d <= N; d++) {
    const pp = eventPeople(input, d);
    for (const [id, text] of pp.must) {
      const m = staff.find((s) => s.id === id), x = cell(id, d);
      if (!m || input.staffEvents?.[id]?.[d] === 'MUST') continue;
      if (!isWork(x)) v.push({ day: d, staffId: id, level: 'error', msg: `${nm(m)} ${d}日は「${text || 'イベント'}」の参加者なので出勤が必要` });
      else if (pp.drill.has(id) && x !== 'A815') v.push({ day: d, staffId: id, level: 'error', msg: `${nm(m)} ${d}日は防災訓練のため8:15出勤（A815）が必要` });
    }
  }
  const weOffs = {};
  for (const m of staff) { weOffs[m.id] = 0; for (const d of input.month.weekendOrHoliday) if (cell(m.id, d) === '/') weOffs[m.id]++; }
  for (const m of staff) {
    const pv = input.prev?.[m.id] || { run: 0, last: '/' };
    let run = pv.run || 0, last = pv.last || '/', offs = 0, b15 = 0, d30 = 0, work = 0;
    let segBefore = 0, gap = 0, curGap = 0;
    const early = [];
    for (let d = 1; d <= N; d++) {
      const x = cell(m.id, d);
      if (isWork(x)) {
        work++; run++;
        if (run === 1) curGap = gap;
        if (m.avoidSplitLongRuns && run === RULES.splitMin && curGap === 1 && segBefore >= RULES.splitMin)
          v.push({ day: d, staffId: m.id, level: 'warn', msg: `${nm(m)} ${d}日 ${RULES.splitMin}連勤以上→休み1日→${RULES.splitMin}連勤以上の形になっています（体調面で避けたい）` });
        if (x === 'A' && isOpener(d)) early.push(d);
        if (run === 6) v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日で6連勤` });
        if (last[0] === 'B' && x[0] === 'A') v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日 B→A` });
        if (x === 'B15') b15++;
        if (x === 'D30') d30++;
      } else {
        if (run > 0) { segBefore = run; gap = 1; } else gap++;
        run = 0;
        if (x === '/' || x === '×') offs++;
        if (input.staffEvents?.[m.id]?.[d] === 'MUST') v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日は出勤必須` });
      }
      last = x;
    }
    if (m.role === 'help') {
      if (work > helpMax(m)) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} ${work}回（月${helpMax(m)}回まで）` });
      continue;
    }
    if (isMgrM(m)) {
      for (const d of RULES.firstDays) if (d <= N && !isWork(cell(m.id, d))) v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日（月初）は店長の出勤が必要` });
      if (run > RULES.endRunMax) v.push({ day: N, staffId: m.id, level: 'error', msg: `${nm(m)} 月末の連勤が${run}日（翌月1日に出勤しても5連勤を超えないよう${RULES.endRunMax}日以内に）` });
      else if (run > RULES.endRunGoal) v.push({ day: N, staffId: m.id, level: 'warn', msg: `${nm(m)} 月末の連勤が${run}日（${RULES.endRunGoal}日以内が目標）` });
      const others = staff.filter((o) => o.role !== 'help' && o.id !== m.id);
      if (others.length) {
        const avg = others.reduce((a, o) => a + weOffs[o.id], 0) / others.length;
        if (weOffs[m.id] > avg) v.push({ day: null, staffId: m.id, level: 'warn', msg: `${nm(m)} 土日祝の公休が${weOffs[m.id]}日（ほかの社員は平均${avg.toFixed(1)}日。店長は土日祝に出たい）` });
      }
    }
    if (early.length) v.push({ day: null, staffId: m.id, level: 'warn', msg: `${nm(m)} 開店からバイトがいる日に早番(A)が${early.length}回（${early.join('・')}日）。社員は遅番寄りにしたい` });
    if (offs !== 8) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} 公休${offs}日（8日必要）` });
    if (b15 > 3) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} B15が${b15}回（3回まで）` });
    if (d30 > 2) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} D30が${d30}回（2回まで）` });
  }
  const hc = headcount(input, schedule);
  const TG = targets(input);
  for (let d = 1; d <= N; d++) if (hc[d] <= 2.5 || hc[d] < TG[d]) v.push({ day: d, level: 'warn', msg: `${d}日 人員数${hc[d]}（目標${TG[d]}）` });
  return v;
}

export function suggestHelp(input, schedule) {
  const N = input.month.days, hc = headcount(input, schedule), TG = targets(input), out = {};
  const staff = input.members.filter((m) => m.kind === 'staff');
  const cell = (id, d) => schedule.staff?.[id]?.[d] || '/';
  const worked = (id) => { let n = 0; for (let d = 1; d <= N; d++) if (isWork(cell(id, d))) n++; return n; };
  const left = {};
  for (const m of staff) if (m.role === 'help') left[m.id] = helpMax(m) - worked(m.id);
  const canWork = (m, d) => {
    const pv = input.prev?.[m.id] || { run: 0, last: '/' };
    const at = (k) => (k < 1 ? (k === 0 ? pv.last || '/' : '/') : k > N ? '/' : cell(m.id, k));
    let run = 1;
    for (let k = d - 1; k >= 0; k--) {
      if (k === 0) { run += pv.run || 0; break; }
      if (!isWork(cell(m.id, k))) break;
      run++;
    }
    for (let k = d + 1; k <= N && isWork(cell(m.id, k)); k++) run++;
    if (run > 5) return false;
    const pb = at(d - 1)[0] === 'B', na = at(d + 1)[0] === 'A';
    return !(pb && na) || (!pb && !na);
  };
  for (let d = 1; d <= N; d++) {
    if (hc[d] >= TG[d]) continue;
    const L = [];
    for (const m of staff) {
      if (m.role !== 'help' || left[m.id] <= 0 || isWork(cell(m.id, d)) || input.staffEvents?.[m.id]?.[d]) continue;
      L.push({ memberId: m.id, name: m.name, reason: `応援（残り${left[m.id]}回）` });
    }
    for (const m of input.members) {
      if (m.kind !== 'part' || input.partRequests?.[m.id]?.[d] || schedule.part?.[m.id]?.[d]) continue;
      L.push({ memberId: m.id, name: m.name, reason: '希望外・要相談' });
    }
    for (const m of staff) {
      if (m.role === 'help' || cell(m.id, d) !== '/' || !canWork(m, d)) continue;
      L.push({ memberId: m.id, name: m.name, reason: '公休を出勤に変更' });
    }
    out[d] = L;
  }
  return out;
}
