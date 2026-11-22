const L = 21.5;

export async function ocrCanvas(canvas, onProgress, psm) {
  const T = globalThis.Tesseract;
  const worker = await T.createWorker('jpn', 1, {
    logger: (m) => { if (onProgress && m.status === 'recognizing text') onProgress(m.progress); },
  });
  try {
    if (psm) await worker.setParameters({ tessedit_pageseg_mode: String(psm) });
    const { data } = await worker.recognize(canvas);
    return data.text;
  } finally { await worker.terminate(); }
}

export async function loadBitmap(fileOrBlob) {
  return createImageBitmap(fileOrBlob);
}

export async function ocrImage(fileOrBlob, onProgress) {
  const bmp = await loadBitmap(fileOrBlob);
  const c = document.createElement('canvas');
  c.width = bmp.width * 2; c.height = bmp.height * 2;
  const g = c.getContext('2d');
  g.imageSmoothingQuality = 'high';
  g.drawImage(bmp, 0, 0, c.width, c.height);
  const im = g.getImageData(0, 0, c.width, c.height), d = im.data;
  for (let i = 0; i < d.length; i += 4) {
    const v = (d[i] * 299 + d[i + 1] * 587 + d[i + 2] * 114) / 1000;
    d[i] = d[i + 1] = d[i + 2] = v;
  }
  g.putImageData(im, 0, 0);
  return ocrCanvas(c, onProgress);
}

const norm = (s) => s.normalize('NFKC').replace(/[〜~～\-－―ー—–−]/g, '~');
const END_L = /^(?:[LlＬ]|ラスト|ﾗｽﾄ|閉店|クローズ)/;
const TM = /^(\d{1,2})(?::(\d{2})|時(?:(\d{1,2})分?|(半))?)/;

function readTime(s) {
  const m = TM.exec(s);
  if (!m) return null;
  const h = +m[1];
  if (h > 24) return null;
  const min = m[2] ? +m[2] : m[3] ? +m[3] : m[4] ? 30 : 0;
  return { v: h + min / 60, len: m[0].length };
}

function cleanLine(line) {
  let s = norm(line);
  s = s.replace(/(\d)\s+(?=\d?\s*[月日])/g, '$1');
  s = s.replace(/\s+(?=[日月時分曜()\[\]:~\/])/g, '').replace(/(?<=[日月時分曜()\[\]:~\/])\s+/g, '');
  return s.replace(/^[\s・*•●○■□▶>\-\.]+/, '');
}

// OCR由来の崩れ（空白・〜の誤読・時刻の連結）に強い解析
function parseRest(rest) {
  rest = rest.replace(/\s+/g, '');
  if (/いつ|でも|何時|終日|全日|一日中|OK/.test(rest)) return { start: null, end: null, allDay: true };
  const ts = [...rest.matchAll(/(\d{1,2})(?::|時)(\d{2})?/g)].map(m => +m[1] + (m[2] ? +m[2] : 0) / 60).filter(v => v >= 5 && v <= 24);
  if (ts.length >= 2) return { start: ts[0], end: ts[1] };
  if (ts.length === 1) return { start: ts[0], end: L };
  if (!/\d/.test(rest)) return { start: null, end: null, allDay: true };
  return parseRestStrict(rest);
}

function parseRestStrict(rest) {
  rest = rest.replace(/^[(\[][月火水木金土日祝]?[)\]]/, '').replace(/^[月火水木金土日]曜日?/, '');
  const ti = rest.indexOf('~');
  if (ti < 0) {
    const t = readTime(rest.replace(/^[^\d]*/, ''));
    return t ? { start: t.v, end: L } : null;
  }
  const before = rest.slice(0, ti).replace(/^[^\d]*/, ''), after = rest.slice(ti + 1);
  const s = readTime(before);
  let e = null;
  if (END_L.test(after)) e = L;
  else { const t = readTime(after); if (t) e = t.v; }
  if (!s) return e === L ? { start: null, end: null, allDay: true } : e != null ? { start: 9.5, end: e } : null;
  return { start: s.v, end: e ?? L };
}

export function parseRequestText(text, year, month) {
  const days = new Date(year, month, 0).getDate();
  const entries = {};
  let cur = month;
  for (const raw of String(text).split(/\r?\n/)) {
    const s = cleanLine(raw);
    let m, d, rest;
    if ((m = /^(?:\d{4}[\/年])?(\d{1,2})月(\d{1,2})日?(.*)$/.exec(s)) || (m = /^(?:\d{4}\/)?(\d{1,2})\/(\d{1,2})(.*)$/.exec(s))) {
      cur = +m[1]; d = +m[2]; rest = m[3];
    } else if ((m = /^(\d{1,2})日(.*)$/.exec(s))) {
      d = +m[1]; rest = m[2];
    } else continue;
    if (cur !== month || d < 1 || d > days) continue;
    const r = parseRest(rest);
    if (r) entries[d] = r;
  }
  return { name: null, entries };
}

export function guessSender(text, partNames) {
  const t = String(text).normalize('NFKC').replace(/\s+/g, '');
  let best = null, bestPos = Infinity;
  for (const full of partNames) {
    const parts = full.normalize('NFKC').split(/\s+/).filter(Boolean);
    const cands = [parts.join('')];
    if (parts.length > 1) cands.push(parts[0]);
    else if (parts[0].length >= 3) cands.push(parts[0].slice(0, 2));
    for (const c of cands) {
      const p = t.indexOf(c);
      if (c && p >= 0 && p < bestPos) { best = full; bestPos = p; }
    }
  }
  return best;
}
