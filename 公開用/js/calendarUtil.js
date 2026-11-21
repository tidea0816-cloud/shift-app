const dowOf = (y, m, d) => new Date(Date.UTC(y, m - 1, d)).getUTCDay();
const nthMon = (y, m, n) => 1 + ((8 - dowOf(y, m, 1)) % 7) + (n - 1) * 7;
const spring = y => Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
const autumn = y => Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
const key = dt => (dt.getUTCMonth() + 1) * 100 + dt.getUTCDate();

export function holidaysOfYear(y) {
  const H = new Map();
  const add = (m, d, n) => H.set(m * 100 + d, n);
  add(1, 1, '元日');
  add(1, nthMon(y, 1, 2), '成人の日');
  if (y >= 1967) add(2, 11, '建国記念の日');
  if (y >= 2020) add(2, 23, '天皇誕生日');
  add(3, spring(y), '春分の日');
  add(4, 29, y >= 2007 ? '昭和の日' : 'みどりの日');
  add(5, 3, '憲法記念日'); add(5, 4, y >= 2007 ? 'みどりの日' : '国民の休日'); add(5, 5, 'こどもの日');
  if (y === 2020) add(7, 23, '海の日'); else if (y === 2021) add(7, 22, '海の日'); else add(7, nthMon(y, 7, 3), '海の日');
  if (y >= 2016) add(8, y === 2020 ? 10 : y === 2021 ? 8 : 11, '山の日');
  add(9, nthMon(y, 9, 3), '敬老の日');
  add(9, autumn(y), '秋分の日');
  if (y === 2020) add(7, 24, 'スポーツの日'); else if (y === 2021) add(7, 23, 'スポーツの日'); else add(10, nthMon(y, 10, 2), y >= 2020 ? 'スポーツの日' : '体育の日');
  add(11, 3, '文化の日');
  add(11, 23, '勤労感謝の日');
  const base = [...H.keys()];
  // 国民の休日：前後が祝日の平日
  for (let t = Date.UTC(y, 0, 2); t < Date.UTC(y, 11, 31); t += 864e5) {
    const dt = new Date(t);
    if (H.has(key(dt)) || dt.getUTCDay() === 0) continue;
    if (H.has(key(new Date(t - 864e5))) && H.has(key(new Date(t + 864e5)))) add(dt.getUTCMonth() + 1, dt.getUTCDate(), '国民の休日');
  }
  // 振替休日：日曜の祝日の次の非祝日
  for (const k of base) {
    const m = Math.floor(k / 100), d = k % 100;
    if (dowOf(y, m, d) !== 0) continue;
    let t = Date.UTC(y, m - 1, d) + 864e5;
    while (H.has(key(new Date(t)))) t += 864e5;
    const dt = new Date(t);
    if (dt.getUTCFullYear() === y) add(dt.getUTCMonth() + 1, dt.getUTCDate(), '振替休日');
  }
  return H;
}

export function makeMonth(year, month) {
  const days = new Date(Date.UTC(year, month, 0)).getUTCDate();
  const H = holidaysOfYear(year);
  const dow = [null], holidays = new Set(), weekendOrHoliday = new Set(), holidayNames = {};
  for (let d = 1; d <= days; d++) {
    dow[d] = dowOf(year, month, d);
    const n = H.get(month * 100 + d);
    if (n) { holidays.add(d); holidayNames[d] = n; }
    if (dow[d] === 0 || dow[d] === 6 || n) weekendOrHoliday.add(d);
  }
  return { year, month, days, dow, holidays, weekendOrHoliday, holidayNames };
}
