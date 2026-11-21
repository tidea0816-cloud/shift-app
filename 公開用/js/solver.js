const CODES = ['A', 'D', 'B', '/', '×', 'リ', '有'];
const codeOf = (s) => {
  if (!s) return 3;
  const i = CODES.indexOf(s[0]);
  return i < 0 ? 3 : i;
};
const isWork = (s) => !!s && 'ADB'.includes(s[0]);
const HELP_MAX = 3;

function partTimes(r) {
  if (!r) return null;
  if (r.allDay) return [12.5, 21.5];
  if (r.start != null) return [r.start, r.end != null ? r.end : 21.5];
  return null;
}

function buildPart(input) {
  const part = {};
  for (const m of input.members) {
    if (m.kind !== 'part') continue;
    const req = input.partRequests?.[m.id] || {};
    part[m.id] = {};
    for (const d of Object.keys(req)) {
      const t = partTimes(req[d]);
      if (t) part[m.id][d] = t;
    }
  }
  return part;
}

export function solveSync(input, onProgress, opts = {}) {
  const { month, members } = input;
  const N = month.days;
  const we = new Uint8Array(N + 2);
  for (const d of month.weekendOrHoliday) we[d] = 1;
  const staff = members.filter((m) => m.kind === 'staff');
  const nS = staff.length;
  const part = buildPart(input);
  const pt = new Float64Array(N + 2), opener = new Uint8Array(N + 2);
  for (const id in part) for (const d in part[id]) {
    pt[d] += 0.5;
    if (part[id][d][0] <= 10) opener[d] = 1;
  }
  const TG = new Float64Array(N + 2);
  for (let d = 1; d <= N; d++) {
    let a = d, b = d;
    while (a > 1 && we[a - 1]) a--;
    while (b < N && we[b + 1]) b++;
    TG[d] = !we[d] ? 3 : (b - a + 1 >= 3 && d !== b ? 5 : 4);
  }
  const R = [], fixedMask = [], fixedStr = [], must = [], prevRun = [], prevLast = [], tgt = [], isRoleA = [], isHelp = [], freeDays = [], allowHelp = new Uint8Array(N + 2);
  const ev = input.staffEvents || {};
  for (const m of staff) for (const d in (ev[m.id] || {})) if (m.role !== 'help' && 'リ×有'.includes(ev[m.id][d])) allowHelp[d] = 1;
  const lockedIn = input.locked || {};
  const base = input.base?.staff || {};
  staff.forEach((m, i) => {
    const row = new Int8Array(N + 2), fx = new Uint8Array(N + 2), mu = new Uint8Array(N + 2), fs = {};
    const help = m.role === 'help';
    isHelp[i] = help; isRoleA[i] = m.role === 'A';
    let fixedOther = 0;
    for (let d = 1; d <= N; d++) {
      const e = (ev[m.id] || {})[d];
      if (e === 'MUST') mu[d] = 1;
      if (e && e !== 'MUST') {
        fx[d] = 1; row[d] = codeOf(e);
        if (e === 'リ' || e === '有') fixedOther++;
      } else if (help && !allowHelp[d]) { fx[d] = 1; row[d] = 3; }
      if (lockedIn[m.id]?.[d] && base[m.id]?.[d]) {
        fx[d] = 1; row[d] = codeOf(base[m.id][d]); fs[d] = base[m.id][d];
      }
    }
    const fd = [];
    for (let d = 1; d <= N; d++) if (!fx[d]) fd.push(d);
    const work = Math.max(0, N - 8 - fixedOther);
    let ta, td, tb;
    if (m.role === 'A') { ta = Math.round(work * 0.75); td = Math.round(work * 0.15); }
    else if (m.role === 'B') { ta = 2; td = Math.round(work * 0.3); }
    else { ta = Math.round(work * 0.15); td = Math.round(work * 0.25); }
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
        c += 20 * nb + 1000 * Math.max(0, nb - HELP_MAX);
        continue;
      }
      let run = prevRun[i], last = prevLast[i], offs = 0, cA = 0, cD = 0, cB = 0, two = 0;
      const mu = must[i];
      for (let d = 1; d <= N; d++) {
        const x = row[d];
        if (x >= 3) {
          run = 0;
          if (x <= 4) { offs++; if (we[d] && x === 3) c += 10; }
          if (d > 1 && row[d - 1] >= 3 && !(x >= 5 && row[d - 1] >= 5)) two++;
          if (mu[d]) c += 1000;
        } else {
          run++;
          if (run > 5) c += 1000; else if (run === 5) c += 30;
          if (last === 2 && x === 0) c += 1000;
          if (x === 0) cA++; else if (x === 1) cD++; else cB++;
        }
        last = x;
      }
      const t = tgt[i];
      c += 1000 * Math.abs(offs - 8) + 3 * (Math.abs(cA - t[0]) + Math.abs(cD - t[1]) + Math.abs(cB - t[2]));
      c += 8 * Math.max(0, two - 1);
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
        if (nA === 0) c += 60;
        if (nW === 3 && !((nA === 1 && nD === 1) || (nA === 0 && nD === 1) || (nA === 1 && nD === 0))) c += 10;
      } else {
        if (!opener[d] && nA === 0) c += 1000;
        if (opener[d] && aNon > 0) c += 25;
        if (nW === 2 && !(nD === 1 && nB === 1)) c += 6;
        if (nW === 3) c += 15;
      }
      const n = nW + pt[d], t = TG[d];
      c += (we[d] ? 40 : 30) * Math.max(0, t - n) + 10 * Math.max(0, n - t - 1);
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
  const out = { staff: {}, part, locked: input.locked || {} };
  staff.forEach((m, i) => {
    const o = {};
    for (let d = 1; d <= N; d++) o[d] = fixedStr[i][d] || CODES[(best ? best.R[i] : R[i])[d]];
    out.staff[m.id] = o;
  });
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
  for (const m of staff) {
    const pv = input.prev?.[m.id] || { run: 0, last: '/' };
    let run = pv.run || 0, last = pv.last || '/', offs = 0, b15 = 0, d30 = 0, work = 0;
    for (let d = 1; d <= N; d++) {
      const x = cell(m.id, d);
      if (isWork(x)) {
        work++; run++;
        if (run === 6) v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日で6連勤` });
        if (last[0] === 'B' && x[0] === 'A') v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日 B→A` });
        if (x === 'B15') b15++;
        if (x === 'D30') d30++;
      } else {
        run = 0;
        if (x === '/' || x === '×') offs++;
        if (input.staffEvents?.[m.id]?.[d] === 'MUST') v.push({ day: d, staffId: m.id, level: 'error', msg: `${nm(m)} ${d}日は出勤必須` });
      }
      last = x;
    }
    if (m.role === 'help') {
      if (work > HELP_MAX) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} ${work}回（月${HELP_MAX}回まで）` });
      continue;
    }
    if (offs !== 8) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} 公休${offs}日（8日必要）` });
    if (b15 > 3) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} B15が${b15}回（3回まで）` });
    if (d30 > 2) v.push({ day: null, staffId: m.id, level: 'error', msg: `${nm(m)} D30が${d30}回（2回まで）` });
  }
  const hc = headcount(input, schedule);
  for (let d = 1; d <= N; d++) if (hc[d] <= 2.5) v.push({ day: d, level: 'warn', msg: `${d}日 人員数${hc[d]}` });
  return v;
}
