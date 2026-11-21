import { solveSync } from './solver.js';

self.onmessage = (e) => {
  try {
    let last = 0;
    const schedule = solveSync(e.data.input, (v) => {
      const t = Date.now();
      if (t - last > 100) { last = t; self.postMessage({ type: 'progress', v }); }
    });
    self.postMessage({ type: 'done', schedule });
  } catch (err) {
    self.postMessage({ type: 'error', message: String(err) });
  }
};
