// 숫자 표기 규칙 (요구사항정리 6-3 임시 기준 T-03, T-04)
export const RULES = {
  eokDigits: 2, // 억원 소수 자릿수 (T-03)
  cheonDigits: 1, // 천만원 소수 자릿수 (T-03)
  pctDigits: 2, // % 소수 자릿수 (T-04)
  timesDigits: 2, // 회전율 소수 자릿수 (T-04)
  divZero: '-', // 0으로 나눌 때 (T-04)
  capitalDeficit: '자본잠식', // T-01
  turnaround: '흑자전환', // Q-04 / T-07
};

export const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

export function won(v) {
  if (!isNum(v)) return v ?? '-';
  return Math.round(v).toLocaleString('ko-KR');
}

export function wonUnit(v) {
  return isNum(v) ? `${won(v)}원` : String(v ?? '-');
}

// 요약 카드 단위 (Q-03): 1억 이상 억원, 1억 미만 천만원
export function eok(v, signed = false) {
  if (!isNum(v)) return String(v ?? '-');
  const abs = Math.abs(v);
  const sign = v < 0 ? '-' : signed && v > 0 ? '+' : '';
  if (abs >= 1e8) return `${sign}${(abs / 1e8).toFixed(RULES.eokDigits)}억원`;
  return `${sign}${(abs / 1e7).toFixed(RULES.cheonDigits)}천만원`;
}

export function pct(v, signed = false) {
  if (!isNum(v)) return String(v ?? '-');
  const s = v.toFixed(RULES.pctDigits);
  return `${signed && v > 0 ? '+' : ''}${s}%`;
}

export function times(v) {
  if (!isNum(v)) return String(v ?? '-');
  return `${v.toFixed(RULES.timesDigits)}회`;
}

export function div(a, b) {
  if (!isNum(a) || !isNum(b) || b === 0) return RULES.divZero;
  return a / b;
}

export function fmtDateKo(s) {
  const m = String(s || '').match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/);
  if (!m) return String(s || '');
  return `${m[1]}년 ${Number(m[2])}월 ${Number(m[3])}일`;
}

export function todayIso() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
