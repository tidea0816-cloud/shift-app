import { ocrCanvas, loadBitmap } from './parseText.js';

const hex = (c) => '#' + c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('');
const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]);

function findRows(d, W, H) {
  const seps = [];
  for (let y = 0; y < H; y++) {
    let n = 0;
    for (let x = 0; x < W; x += 2) {
      const i = (y * W + x) * 4, r = d[i], g = d[i + 1], b = d[i + 2];
      const mn = Math.min(r, g, b), mx = Math.max(r, g, b);
      if (mn >= 205 && mn <= 242 && mx - mn < 22) n++;
    }
    if (n > (W / 2) * 0.9) { if (seps.length && y - seps[seps.length - 1].e <= 2) seps[seps.length - 1].e = y; else seps.push({ s: y, e: y }); }
  }
  const ys = seps.map((p) => (p.s + p.e) / 2);
  const diffs = ys.slice(1).map((v, i) => v - ys[i]).sort((a, b) => a - b);
  const pitch = diffs.length ? diffs[Math.floor(diffs.length / 2)] : H / 6;
  const ok = ys.filter((y, i) => (i > 0 && Math.abs(y - ys[i - 1] - pitch) < 8) || (i < ys.length - 1 && Math.abs(ys[i + 1] - y - pitch) < 8));
  if (!ok.length) return { top: 0, pitch, rows: 6 };
  return { top: ok[0] - pitch, pitch, rows: ok.length + 1 };
}

function findBars(d, W, H, top) {
  const vis = new Uint8Array(W * H), out = [], stack = new Int32Array(W * H);
  const chroma = (i) => Math.max(d[i], d[i + 1], d[i + 2]) - Math.min(d[i], d[i + 1], d[i + 2]);
  const y0 = Math.max(0, Math.floor(top));
  for (let y = y0; y < H; y++) for (let x = 0; x < W; x++) {
    const p = y * W + x;
    if (vis[p]) continue;
    const i0 = p * 4;
    if (chroma(i0) < 40) continue;
    const seed = [d[i0], d[i0 + 1], d[i0 + 2]];
    let sp = 0, n = 0, x1 = x, x2 = x, y2 = y, sr = 0, sg = 0, sb = 0;
    stack[sp++] = p; vis[p] = 1;
    while (sp) {
      const q = stack[--sp], qx = q % W, qy = (q / W) | 0, i = q * 4;
      n++; sr += d[i]; sg += d[i + 1]; sb += d[i + 2];
      if (qx < x1) x1 = qx; if (qx > x2) x2 = qx; if (qy > y2) y2 = qy;
      for (const [dx, dy] of [[1, 0], [-1, 0], [0, 1], [0, -1]]) {
        const nx = qx + dx, ny = qy + dy;
        if (nx < 0 || ny < y0 || nx >= W || ny >= H) continue;
        const np = ny * W + nx;
        if (vis[np]) continue;
        const ni = np * 4;
        if (chroma(ni) < 40 || dist(seed, [d[ni], d[ni + 1], d[ni + 2]]) > 60) continue;
        vis[np] = 1; stack[sp++] = np;
      }
    }
    const w = x2 - x1 + 1, h = y2 - y + 1, s = W / 868;
    if (w >= 70 * s && h >= 24 * s && h <= 60 * s && n / (w * h) > 0.6) out.push({ x1, x2, y1: y, y2, rgb: [sr / n, sg / n, sb / n] });
  }
  return out;
}

function barKind(d, W, b) {
  const s = W / 868;
  let maxX = -1, cnt = 0;
  for (let y = b.y1 + 3; y <= b.y2 - 3; y++) for (let x = b.x1 + 3; x <= b.x2 - 3; x++) {
    const i = (y * W + x) * 4, mn = Math.min(d[i], d[i + 1], d[i + 2]);
    if (mn > 215 && Math.max(d[i], d[i + 1], d[i + 2]) - mn < 30) { cnt++; if (x > maxX) maxX = x; }
  }
  return cnt === 0 || maxX - b.x1 < 36 * s ? 'x' : 'text';
}

export async function analyzeCalendar(fileOrBlob, monthInfo) {
  const bmp = await loadBitmap(fileOrBlob);
  const W = bmp.width, H = bmp.height;
  const cv = document.createElement('canvas');
  cv.width = W; cv.height = H;
  const g = cv.getContext('2d', { willReadFrequently: true });
  g.drawImage(bmp, 0, 0);
  const d = g.getImageData(0, 0, W, H).data;
  const { top, pitch, rows } = findRows(d, W, H);
  const colW = W / 7, firstDow = monthInfo.dow[1], s = W / 868;

  const clusters = [];
  const raw = findBars(d, W, H, top).map((b) => {
    let c = clusters.find((k) => dist(k.rgb, b.rgb) < 40);
    if (!c) clusters.push((c = { key: 'c' + (clusters.length + 1), rgb: b.rgb }));
    const row = Math.floor(((b.y1 + b.y2) / 2 - top) / pitch);
    const c1 = Math.floor((b.x1 + 4 * s) / colW), c2 = Math.floor((b.x2 - 4 * s) / colW);
    return { ...b, colorKey: c.key, row, c1, c2, kind: barKind(d, W, b), bleed: b.x2 >= W - 3 * s };
  }).filter((b) => b.row >= 0 && b.row < rows).sort((a, b) => a.row - b.row || a.x1 - b.x1);

  const merged = [];
  for (const b of raw) {
    const prev = merged.find((m) => m.colorKey === b.colorKey && m.kind === b.kind && m.last.row + 1 === b.row && (m.kind === 'text' ? m.last.c2 === 6 : m.last.bleed) && b.c1 === 0);
    if (prev) { prev.segs.push(b); prev.last = b; } else merged.push({ colorKey: b.colorKey, kind: b.kind, segs: [b], last: b });
  }

  const bars = merged.map((m) => {
    const days = [];
    for (const b of m.segs) for (let c = Math.max(0, b.c1); c <= Math.min(6, b.c2); c++) {
      const day = b.row * 7 + c - firstDow + 1;
      if (day >= 1 && day <= monthInfo.days) days.push(day);
    }
    if (!days.length) return null;
    const w = m.segs.reduce((a, b) => (b.x2 - b.x1 > a.x2 - a.x1 ? b : a));
    const pad = 2, cw = w.x2 - w.x1 + 1 + pad * 2, ch = w.y2 - w.y1 + 1 + pad * 2;
    const c = document.createElement('canvas');
    c.width = cw; c.height = ch;
    c.getContext('2d').drawImage(cv, w.x1 - pad, w.y1 - pad, cw, ch, 0, 0, cw, ch);
    return { colorKey: m.colorKey, days, kind: m.kind, cropDataUrl: c.toDataURL('image/png') };
  }).filter(Boolean);

  return { colors: clusters.map((k) => ({ key: k.key, hex: hex(k.rgb) })), bars };
}

export async function ocrBar(cropDataUrl) {
  const img = new Image();
  img.src = cropDataUrl;
  await img.decode();
  const k = 3, c = document.createElement('canvas');
  c.width = img.width * k; c.height = img.height * k;
  const g = c.getContext('2d');
  g.imageSmoothingQuality = 'high';
  g.drawImage(img, 0, 0, c.width, c.height);
  const im = g.getImageData(0, 0, c.width, c.height), d = im.data;
  for (let i = 0; i < d.length; i += 4) {
    const mn = Math.min(d[i], d[i + 1], d[i + 2]), t = mn > 170 && Math.max(d[i], d[i + 1], d[i + 2]) - mn < 60;
    d[i] = d[i + 1] = d[i + 2] = t ? 0 : 255;
  }
  g.putImageData(im, 0, 0);
  const pad = document.createElement('canvas');
  pad.width = c.width + 40; pad.height = c.height + 40;
  const pg = pad.getContext('2d');
  pg.fillStyle = '#fff'; pg.fillRect(0, 0, pad.width, pad.height);
  pg.drawImage(c, 20, 20);
  return (await ocrCanvas(pad, null, 7)).trim();
}
