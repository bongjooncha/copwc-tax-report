const number = (v) => typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'string' && /^[-+]?\d[\d,]*(?:\.\d+)?$/.test(v.trim()) ? Number(v.replaceAll(',', '')) : null;
const specs = {
  표준대차대조표: ['유동자산', '당좌자산', '재고자산', '비유동자산', '자산총계', '유동부채', '부채총계', '자본금', '자본총계', '부채와자본총계'],
  표준손익계산서: ['매출액', '매출원가', '매출총손익', '판매비와관리비', '영업손익', '영업외수익', '영업외비용', '법인세비용차감전손익', '법인세비용', '당기순손익'],
  법인세신고상황: ['익금조정', '손금조정', '각사업연도소득금액', '과세표준금액', '산출세액', '총부담세액', '기납부세액', '차감납부할세액', '법인지방소득세'],
};
export function validateReport(d) {
  const issues = []; const add = (kind, message) => issues.push({ kind, message });
  if (!d?.법인기본사항?.법인명) add('누락', '법인명: 입력해 주세요.');
  if (!d?.사업연도?.당기) add('누락', '당기 사업연도: 입력해 주세요.');
  for (const y of ['당기', '전기', '전전기']) {
    if (!d?.사업연도?.[y]) continue;
    for (const [table, keys] of Object.entries(specs)) for (const key of keys) if (number(d?.[table]?.[key]?.[y]) === null) add('확인 필요', `${table} · ${key} · ${y}: 금액이 없거나 숫자가 아닙니다.`);
    const get = (table, key) => number(d?.[table]?.[key]?.[y]);
    const check = (label, expected, actual) => { if (expected !== null && actual !== null && Math.abs(expected - actual) > 1) add('불일치', `${y} ${label}: 계산값 ${expected.toLocaleString('ko-KR')}원, 입력값 ${actual.toLocaleString('ko-KR')}원입니다. 원본을 확인해 주세요.`); };
    const plus = (a, b) => a === null || b === null ? null : a + b;
    const minus = (a, b) => a === null || b === null ? null : a - b;
    check('자산 = 부채 + 자본', plus(get('표준대차대조표', '부채총계'), get('표준대차대조표', '자본총계')), get('표준대차대조표', '자산총계'));
    check('매출총손익 = 매출액 − 매출원가', minus(get('표준손익계산서', '매출액'), get('표준손익계산서', '매출원가')), get('표준손익계산서', '매출총손익'));
    check('영업손익 = 매출총손익 − 판관비', minus(get('표준손익계산서', '매출총손익'), get('표준손익계산서', '판매비와관리비')), get('표준손익계산서', '영업손익'));
  }
  for (const w of Array.isArray(d?.확인필요?.경고) ? d.확인필요.경고 : []) add(w.수준 || '자료 확인', [w.항목, w.내용, w.파일 && `출처: ${w.파일} ${w.페이지 || ''}`].filter(Boolean).join(' — '));
  return issues;
}
