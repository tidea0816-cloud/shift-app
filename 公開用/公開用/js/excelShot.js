// エクセル（ベースの表）のスクショから、日付ごとのイベント・スケジュールを読む。
// Tesseract の単語ボックス（data.words の bbox）の座標から「列→日」を決める。精度は手直し前提。
import { loadBitmap } from './parseText.js';

export const EVENT_TYPES = [['storeMT', '店長MT'], ['subMT', '副店長MT'], ['drill', '防災訓練'], ['busy', '繁忙'], ['other', 'その他']];
const nz = (s) => String(s ?? '').normalize('NFKC');

export function guessEventType(text) {
  const t = nz(text).replace(/\s+/g, '');
  if (/防災|避難訓練/.test(t)) return 'drill';
  if (/サブ.?MT|副店長/i.test(t)) return 'subMT';
  if (/店長.?MT|店長会議|店長ミーティング/i.test(t)) return 'storeMT';
  if (/セール|フェア|初売|福袋|祭|催事/.test(t)) return 'busy';
  return 'other';
}

function collectWords(data, sc) {
  let ws = data.words;
  if (!ws || !ws.length) {
    ws = [];
    const walk = (o) => {
      if (!o) return;
      if (o.words) ws.push(...o.words);
      for (const k of ['blocks', 'paragraphs', 'lines']) if (o[k]) o[k].forEach(walk);
    };
    walk(data);
  }
  return ws.filter((w) => w && w.bbox && String(w.text || '').trim()).map((w) => ({
    text: w.text.trim(), x0: w.bbox.x0 / sc, y0: w.bbox.y0 / sc, x1: w.bbox.x1 / sc, y1: w.bbox.y1 / sc,
  }));
}

export async function ocrWords(file, onProgress) {
  const T = globalThis.Tesseract;
  const bmp = await loadBitmap(file);
  const sc = Math.max(1, Math.min(2, 4500 / bmp.width));
  const c = document.createElement('canvas');
  c.width = Math.round(bmp.width * sc); c.height = Math.round(bmp.height * sc);
  const g = c.getContext('2d');
  g.imageSmoothingQuality = 'high';
  g.drawImage(bmp, 0, 0, c.width, c.height);
  const im = g.getImageData(0, 0, c.width, c.height), d = im.data;
  for (let i = 0; i < d.length; i += 4) d[i] = d[i + 1] = d[i + 2] = (d[i] * 299 + d[i + 1] * 587 + d[i + 2] * 114) / 1000;
  g.putImageData(im, 0, 0);
  const worker = await T.createWorker('jpn', 1, { logger: (m) => { if (onProgress && m.status === 'recognizing text') onProgress(m.progress); } });
  try {
    await worker.setParameters({ tessedit_pageseg_mode: '11' });
    const { data } = await worker.recognize(c);
    return { words: collectWords(data, sc), width: bmp.width, height: bmp.height };
  } finally { await worker.terminate(); }
}

const median = (a) => { const s = [...a].sort((x, y) => x - y); return s.length ? s[s.length >> 1] : 0; };
const cx = (w) => (w.x0 + w.x1) / 2, cy = (w) => (w.y0 + w.y1) / 2;

// 日付の数字（1〜31）が横一列に並んだ行を探し、x座標から列→日の対応式（x = a + b*日）を作る
function findDateRow(words) {
  const nums = words.map((w) => ({ w, v: /^\d{1,2}$/.test(nz(w.text)) ? +nz(w.text) : 0 })).filter((o) => o.v >= 1 && o.v <= 31);
  if (nums.length < 4) return null;
  const tol = Math.max(6, median(nums.map((o) => o.w.y1 - o.w.y0)) * 0.6);
  nums.sort((p, q) => cy(p.w) - cy(q.w));
  const groups = [];
  for (const o of nums) {
    const g = groups[groups.length - 1];
    if (g && Math.abs(cy(o.w) - g.cy / g.list.length) < tol) { g.list.push(o); g.cy += cy(o.w); } else groups.push({ list: [o], cy: cy(o.w) });
  }
  let best = null;
  for (const g of groups) {
    const L = g.list.sort((p, q) => cx(p.w) - cx(q.w));
    const len = L.map(() => 1), prv = L.map(() => -1);
    for (let i = 0; i < L.length; i++) for (let j = 0; j < i; j++) if (L[j].v < L[i].v && len[j] + 1 > len[i]) { len[i] = len[j] + 1; prv[i] = j; }
    let k = len.indexOf(Math.max(...len));
    const chain = [];
    for (; k >= 0; k = prv[k]) chain.unshift(L[k]);
    if (chain.length >= 4 && (!best || chain.length > best.chain.length)) best = { chain, cy: g.cy / g.list.length };
  }
  if (!best) return null;
  let pts = best.chain, a = 0, b = 0;
  for (let pass = 0; pass < 2; pass++) {
    const n = pts.length, sx = pts.reduce((s, p) => s + p.v, 0), sy = pts.reduce((s, p) => s + cx(p.w), 0);
    const sxx = pts.reduce((s, p) => s + p.v * p.v, 0), sxy = pts.reduce((s, p) => s + p.v * cx(p.w), 0);
    const den = n * sxx - sx * sx;
    if (!den) return null;
    b = (n * sxy - sx * sy) / den; a = (sy - b * sx) / n;
    if (pass === 0) { const keep = pts.filter((p) => Math.abs(cx(p.w) - (a + b * p.v)) < Math.abs(b) * 0.4); if (keep.length >= 4 && keep.length < pts.length) pts = keep; else break; }
  }
  if (!(b > 3)) return null;
  const h = median(pts.map((p) => p.w.y1 - p.w.y0)) || 12;
  return { a, b, y: best.cy, h, days: pts.map((p) => p.v) };
}

// words: [{text,x0,y0,x1,y1}]（元画像のピクセル座標）, width: 画像の幅
export function layoutEvents(words, width) {
  const row = findDateRow(words);
  if (!row) return { items: [], days: [], error: '日付の行を読み取れませんでした（日付の行が写るように撮ってください）' };
  const { a, b } = row;
  const dayOf = (w) => {
    const x = (w.x1 - w.x0) > b * 1.6 ? w.x0 + b / 2 : cx(w);
    return Math.round((x - a) / b);
  };
  const left = words.filter((w) => cy(w) > row.y + row.h && dayOf(w) < 1);
  const find = (re, minY) => { const y = left.filter((w) => re.test(nz(w.text)) && w.y0 > minY).map((w) => w.y0); return y.length ? Math.min(...y) : null; };
  const evY = find(/イベント|自社|ビル/, row.y);
  const scY = find(/スケジュ|ジュール|撮影|研修/, row.y);
  const endY = scY != null ? find(/社員|出勤|シフト|氏名|公休|予算|人員/, scY + 1) : null;
  const top = row.y + row.h * 0.9, bottom = endY ?? Infinity;
  const kindAt = (y) => (scY != null ? (y < scY - row.h * 0.3 ? 'store' : 'staff') : evY != null ? 'store' : 'staff');
  const cell = new Map();
  for (const w of words) {
    const d = dayOf(w);
    if (d < 1 || d > 31 || cy(w) < top || w.y0 >= bottom) continue;
    if (cy(w) < row.y + row.h * 2.6 && /^[日月火水木金土]$/.test(nz(w.text))) continue;
    if (cy(w) < row.y + row.h * 1.9 && /^\d{1,2}$/.test(nz(w.text)) && Math.abs(cy(w) - row.y) < row.h) continue;
    const key = d + '|' + kindAt(cy(w));
    if (!cell.has(key)) cell.set(key, []);
    cell.get(key).push(w);
  }
  const items = [];
  for (const [key, ws] of cell) {
    const [d, kind] = key.split('|');
    const h = median(ws.map((w) => w.y1 - w.y0)) || row.h;
    ws.sort((p, q) => cy(p) - cy(q));
    const lines = [];
    for (const w of ws) {
      const l = lines[lines.length - 1];
      if (l && Math.abs(cy(w) - l.cy) < h * 0.6) { l.ws.push(w); l.y1 = Math.max(l.y1, w.y1); } else lines.push({ ws: [w], cy: cy(w), y0: w.y0, y1: w.y1 });
    }
    const out = [];
    for (const l of lines) {
      const t = l.ws.sort((p, q) => p.x0 - q.x0).map((w) => nz(w.text)).join('');
      const p = out[out.length - 1];
      if (p && t.length <= 2 && (p.n > 1 || p.t.length <= 2) && l.y0 - p.y1 <= h * 1.8) { p.t += t; p.n++; p.y1 = l.y1; }
      else out.push({ t, n: 1, y1: l.y1 });
    }
    for (const o of out) {
      const text = o.t.replace(/\s+/g, '').replace(/^[|¦｜_\-—.,・:;]+|[|¦｜_\-—.,・:;]+$/g, '');
      if (text.length < 2 || !/[A-Za-z぀-ヿ一-鿿]/.test(text)) continue;
      const type = guessEventType(text);
      items.push({ day: +d, kind: type === 'drill' ? 'staff' : kind, text, type });
    }
  }
  items.sort((p, q) => p.day - q.day || (p.kind < q.kind ? 1 : -1));
  const ds = items.map((i) => i.day);
  return { items, days: row.days, pitch: b, scheduleLabel: scY != null, eventLabel: evY != null, maxDay: Math.max(...row.days, ...ds) };
}

export async function readExcelShot(file, onProgress) {
  const { words, width } = await ocrWords(file, onProgress);
  return { ...layoutEvents(words, width), wordCount: words.length };
}
