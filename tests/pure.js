// Extracted pure helpers mirrored from Code.gs (no Apps Script deps)
export function getIdFromLabelish(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/\(([^)]+)\)\s*$/);
  return m ? m[1] : s;
}

function round6(x){ return Math.round((x + Number.EPSILON) * 1e6) / 1e6; }

export function dynCap(T, payments) {
  if (!T || !isFinite(T)) return 0;
  const arr = (payments || []).map(Number).filter(v => v > 0 && isFinite(v));
  if (!arr.length) return 0;
  arr.sort((a,b)=>a-b);
  const n = arr.length;
  const sum = arr.reduce((a,b)=>a+b,0);
  const target = Math.min(T, sum);
  if (target <= 0) return 0;

  let cumsum = 0;
  for (let k = 0; k < n; k++) {
    const next = arr[k];
    const remain = n - k;
    const candidate = (target - cumsum) / remain;
    if (candidate <= next) return round6(candidate);
    cumsum += next;
  }
  return round6((target - (cumsum - arr[n-1])) / 1);
}
