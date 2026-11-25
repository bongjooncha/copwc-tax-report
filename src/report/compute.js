// 입력 JSON → 보고서에 필요한 계산값 (비율, 평균, 증감, 코멘트, 특이사항)
import { RULES, isNum, div, won, wonUnit, eok, pct, todayIso } from './format.js';

export const YK = ['당기', '전기', '전전기'];
const PREV = { 당기: '전기', 전기: '전전기', 전전기: null };

const n = (v) => {
  if (isNum(v)) return v;
  if (typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v.replace(/,/g, '')))) {
    return Number(v.replace(/,/g, ''));
  }
  return null;
};

export function compute(raw) {
  const d = raw && typeof raw === 'object' ? raw : {};
  const years = { 당기: null, 전기: null, 전전기: null, ...(d.사업연도 || {}) };
  const active = YK.filter((k) => years[k] !== null && years[k] !== undefined && years[k] !== '');
  const has = (yk) => active.includes(yk);

  const pick = (table) => (key, yk) => {
    if (!has(yk)) return null;
    const row = (d[table] || {})[key];
    const v = row && typeof row === 'object' ? n(row[yk]) : null;
    return v;
  };
  const bs = pick('표준대차대조표');
  const is = pick('표준손익계산서');
  const tx = pick('법인세신고상황');

  // ── Ⅱ-(2) 재무비율 ─────────────────────────────
  const turnoverEndOnly = [];
  const ratios = {};
  for (const yk of YK) {
    if (!has(yk)) { ratios[yk] = null; continue; }
    const equity = bs('자본총계', yk);
    const assets = bs('자산총계', yk);
    const sales = is('매출액', yk);
    const prev = PREV[yk] && has(PREV[yk]) ? PREV[yk] : null;
    if (!prev) turnoverEndOnly.push(yk);
    const avg = (key) => { const current = bs(key, yk); const before = prev ? bs(key, prev) : current; return isNum(current) && isNum(before) ? (current + before) / 2 : null; };
    const eqBad = isNum(equity) && equity < 0; // T-01: 자본잠식 연도만 "-"
    ratios[yk] = {
      부채비율: equity < 0 ? RULES.capitalDeficit : mul100(div(bs('부채총계', yk), equity)),
      유동비율: mul100(div(bs('유동자산', yk), bs('유동부채', yk))),
      당좌비율: mul100(div(bs('당좌자산', yk), bs('유동부채', yk))),
      총자산세전순이익률: mul100(div(is('법인세비용차감전손익', yk), assets)),
      영업이익률: mul100(div(is('영업손익', yk), sales)),
      총자산회전율: div(sales, avg('자산총계')),
      자기자본회전율: eqBad ? RULES.divZero : div(sales, avg('자본총계')),
    };
  }

  // ── Ⅲ 연도별 법인세 신고상황 ─────────────────────
  const taxRows = [
    ['매출액(수입금액)', () => (yk) => is('매출액', yk)],
    ['매출총이익', () => (yk) => is('매출총손익', yk)],
    ['영업이익', () => (yk) => is('영업손익', yk)],
    ['당기순손익', () => (yk) => is('당기순손익', yk)],
    ['(+)익금조정', () => (yk) => tx('익금조정', yk)],
    ['(-)손금조정', () => (yk) => tx('손금조정', yk)],
    ['각사업연도 소득금액', () => (yk) => tx('각사업연도소득금액', yk)],
    ['(-)이월결손금', () => (yk) => tx('이월결손금', yk)],
    ['(-)비과세소득', () => (yk) => tx('비과세소득', yk)],
    ['(-)소득공제', () => (yk) => tx('소득공제', yk)],
    ['과세표준금액', () => (yk) => tx('과세표준금액', yk)],
    ['세율(%)', () => (yk) => tx('세율(%)', yk), { noAvg: true, fmt: 'num' }],
    ['산출세액', () => (yk) => tx('산출세액', yk)],
    ['최저한세 적용대상 공제감면세액', () => (yk) => tx('최저한세적용대상공제감면세액', yk)],
    ['최저한세 적용제외 공제감면세액', () => (yk) => tx('최저한세적용제외공제감면세액', yk)],
    ['총부담세액', () => (yk) => tx('총부담세액', yk)],
    ['기납부세액', () => (yk) => tx('기납부세액', yk)],
    ['차감납부할세액', () => (yk) => tx('차감납부할세액', yk)],
    ['법인지방소득세', () => (yk) => tx('법인지방소득세', yk)],
  ].map(([label, f, opt = {}]) => {
    const get = f();
    const vals = Object.fromEntries(YK.map((yk) => [yk, get(yk)]));
    return { label, fmt: opt.fmt || 'won', vals, avg: opt.noAvg ? null : active.some((k) => !isNum(vals[k])) ? '확인 필요' : average(active.map((k) => vals[k])) };
  });
  const rateRows = [
    ['소득률', (yk) => mul100(div(tx('각사업연도소득금액', yk), is('매출액', yk)))],
    ['매출총이익률', (yk) => mul100(div(is('매출총손익', yk), is('매출액', yk)))],
    ['영업이익률', (yk) => mul100(div(is('영업손익', yk), is('매출액', yk)))],
    ['실효세율(과세표준기준)', (yk) => positiveDiv(tx('총부담세액', yk), tx('과세표준금액', yk))],
    ['실효세율(소득금액기준)', (yk) => positiveDiv(tx('총부담세액', yk), tx('각사업연도소득금액', yk))],
  ].map(([label, f]) => {
    const vals = Object.fromEntries(YK.map((yk) => [yk, has(yk) ? f(yk) : null]));
    return { label, fmt: 'pct', vals, avg: active.some((k) => !isNum(vals[k])) ? '확인 필요' : average(active.map((k) => vals[k])) };
  });

  // ── Ⅳ 참고자료 ─────────────────────────────────
  const prepaid = d.기납부세액 || {};
  const prepaidItems = ['중간예납세액', '원천납부세액', '외국납부세액', '수시부과세액'].map((k) => [k, n(prepaid[k])]);
  const lossRows = arr(d.이월결손금);
  const adj = d.소득금액조정 || {};
  const adjPlus = arr(adj.익금산입및손금불산입);
  const adjMinus = arr(adj.손금산입및익금불산입);
  const credits = arr(d.공제감면세액);
  const inventory = ['당기', '전기'].filter(has).map((yk) => {
    const inv = bs('재고자산', yk);
    const cogs = is('매출원가', yk);
    return { yk, inv, cogs, rate: mul100(div(inv, cogs)) };
  });

  // ── Ⅶ 손익 분석 ────────────────────────────────
  const hasPrev = has('전기');
  const nonOp = (yk) => { const a = is('영업외수익', yk), b = is('영업외비용', yk); return isNum(a) && isNum(b) ? a - b : null; };
  const plRows = [
    ['매출액', (yk) => is('매출액', yk)],
    ['매출총이익', (yk) => is('매출총손익', yk)],
    ['영업손익', (yk) => is('영업손익', yk), true],
    ['영업외손익', nonOp, true],
    ['법인세차감전손익', (yk) => is('법인세비용차감전손익', yk), true],
    ['법인세비용', (yk) => is('법인세비용', yk)],
    ['당기순이익', (yk) => is('당기순손익', yk), true],
  ].map(([label, f, profit]) => {
    const cur = f('당기');
    const prev = hasPrev ? f('전기') : null;
    return { label, prev, cur, diff: hasPrev && isNum(cur) && isNum(prev) ? cur - prev : null, rate: hasPrev ? changeRate(prev, cur, profit) : RULES.divZero };
  });

  const sgaAll = arr(d.판매비와관리비명세)
    .map((r) => ({ name: String(r.계정과목 ?? ''), cur: n(r.당기) ?? 0, prev: n(r.전기) ?? 0 }))
    .filter((r) => r.name);
  const sorted = [...sgaAll].sort((a, b) => b.cur - a.cur);
  const top5 = sorted.slice(0, 5); // T-06 당기 금액 상위 5개
  const rest = sorted.slice(5);
  const salesCur = is('매출액', '당기');
  const salesPrev = hasPrev ? is('매출액', '전기') : null;
  const sgaTotal = { cur: sum(sgaAll.map((r) => r.cur)), prev: sum(sgaAll.map((r) => r.prev)) };
  const etc = { name: '기타', cur: sum(rest.map((r) => r.cur)), prev: sum(rest.map((r) => r.prev)) };
  const sgaRows = [...top5, ...(rest.length ? [etc] : [])].map((r) => ({
    ...r,
    diff: hasPrev ? r.cur - r.prev : null,
    share: mul100(div(r.cur, salesCur)),
  }));
  const sgaNew = hasPrev ? sgaAll.filter((r) => !r.prev && r.cur > 0) : [];
  const sgaGone = hasPrev ? sgaAll.filter((r) => r.prev > 0 && !r.cur) : [];

  // ── Ⅵ 요약 카드 (T-10) ──────────────────────────
  const card = (title, key) => {
    const cur = is(key, '당기');
    const prev = hasPrev ? is(key, '전기') : null;
    return { title, value: eok(cur), delta: hasPrev && isNum(cur) && isNum(prev) ? `전기 대비 ${eok(cur - prev, true)}` : '전기 비교 확인 필요' };
  };
  const debtCur = ratios.당기?.부채비율;
  const debtPrev = hasPrev ? ratios.전기?.부채비율 : null;
  let debtDelta = '전기 자료 없음';
  if (hasPrev) {
    if (isNum(debtCur) && isNum(debtPrev)) debtDelta = `전기 대비 ${pct(debtCur - debtPrev, true)}p`;
    else if (debtPrev === RULES.capitalDeficit && isNum(debtCur)) debtDelta = '전기 자본잠식 → 당기 해소';
    else debtDelta = `전기 ${isNum(debtPrev) ? pct(debtPrev) : debtPrev ?? '-'}`;
  }
  const cards = has('당기')
    ? [
        card('매출액', '매출액'),
        card('영업이익', '영업손익'),
        card('당기순이익', '당기순손익'),
        { title: '부채비율', value: isNum(debtCur) ? pct(debtCur) : String(debtCur ?? '-'), delta: debtDelta },
      ]
    : [];

  const ctx = { d, years, active, has, hasPrev, bs, is, tx, ratios, plRows, sgaAll, sgaRows, sgaTotal, salesCur, salesPrev, sgaNew, sgaGone, credits };
  const canAutoComment = ['매출액', '영업손익', '당기순손익', '영업외수익', '영업외비용'].every((k) => isNum(is(k, '당기'))) && (!hasPrev || ['매출액', '영업손익', '당기순손익', '영업외수익', '영업외비용'].every((k) => isNum(is(k, '전기'))));
  const comments = arr(d.요약코멘트).filter((s) => String(s).trim()).length
    ? arr(d.요약코멘트).map(String)
    : canAutoComment ? autoComments(ctx) : ['기초 재무자료가 누락되어 자동 분석을 보류했습니다. 검토 화면에서 값을 확인해 주세요.'];
  const notables = [...autoNotables(ctx), ...arr(d.특이사항).map(String).filter((s) => s.trim())];

  const info = d.보고서정보 || {};
  return {
    d,
    years,
    active,
    has,
    hasPrev,
    bs,
    is,
    ratios,
    turnoverEndOnly,
    taxRows,
    rateRows,
    prepaidItems,
    lossRows,
    adjPlus,
    adjMinus,
    credits,
    inventory,
    plRows,
    sgaAll,
    sgaRows,
    sgaTotal,
    salesCur,
    salesPrev,
    sgaNew,
    sgaGone,
    cards,
    comments,
    notables,
    report: {
      title: info.보고서명 || '법인세 신고 결과 보고서',
      issueDate: info.발행일 || todayIso(),
      issuer: info.발행기관 || '샘플세무법인',
    },
  };
}

// ── helpers ─────────────────────────────────────
function arr(v) {
  return Array.isArray(v) ? v : [];
}
function sum(a) {
  return a.reduce((s, v) => s + (isNum(v) ? v : 0), 0);
}
function mul100(v) {
  return isNum(v) ? v * 100 : v;
}
function positiveDiv(a, b) {
  if (!isNum(b) || b <= 0) return RULES.divZero;
  return mul100(div(a, b));
}
function average(vals) {
  const nums = vals.filter(isNum);
  return nums.length ? nums.reduce((s, v) => s + v, 0) / nums.length : RULES.divZero;
}
// 증감률 (T-07): 전기 > 0 → 숫자 / 전기 ≤ 0 & 당기 > 0 & 이익항목 → 흑자전환 / 그 외 "-"
export function changeRate(prev, cur, profit = false) {
  if (!isNum(prev) || !isNum(cur)) return RULES.divZero;
  if (prev > 0) return ((cur - prev) / Math.abs(prev)) * 100;
  if (profit && cur > 0) return RULES.turnaround;
  return RULES.divZero;
}

function autoComments(c) {
  if (!c.has('당기')) return [];
  const out = [];
  const sales = c.is('매출액', '당기');
  const op = c.is('영업손익', '당기');
  const net = c.is('당기순손익', '당기');
  const pl = Object.fromEntries(c.plRows.map((r) => [r.label, r]));
  const rateTxt = (r) => (isNum(r.rate) ? `${pct(r.rate)} ${r.diff >= 0 ? '증가' : '감소'}` : r.rate === RULES.turnaround ? '흑자전환' : '');

  // 1. 매출 및 수익성
  if (c.hasPrev) {
    const s = pl['매출액'];
    out.push(
      `[매출 및 수익성] 당기 매출액은 ${eok(sales)}으로 전기(${eok(s.prev)}) 대비 ${eok(s.diff, true)}(${rateTxt(s)})하였습니다. ` +
        `영업손익은 ${eok(op)}(${pl['영업손익'].rate === RULES.turnaround ? '흑자전환' : `영업이익률 ${pct(c.ratios.당기.영업이익률)}`}), 당기순손익은 ${eok(net)}입니다.`,
    );
  } else {
    out.push(`[매출 및 수익성] 당기 매출액은 ${eok(sales)}, 영업손익은 ${eok(op)}, 당기순손익은 ${eok(net)}입니다.`);
  }

  // 2. 판관비 주요항목
  if (c.sgaAll.length) {
    const top = c.sgaRows.filter((r) => r.name !== '기타').slice(0, 3);
    const sgaShare = div(c.sgaTotal.cur, c.salesCur);
    let s = `[판관비 주요항목] 판매비와관리비는 ${wonUnit(c.sgaTotal.cur)}${isNum(sgaShare) ? `(매출액 대비 ${pct(sgaShare * 100)})` : ''}이며, 주요 항목은 ${top
      .map((r) => `${r.name}(${isNum(r.share) ? pct(r.share) : '-'})`)
      .join(', ')}입니다.`;
    if (c.hasPrev) {
      const byDiff = [...c.sgaAll].map((r) => ({ ...r, diff: r.cur - r.prev })).sort((a, b) => b.diff - a.diff);
      const up = byDiff[0];
      const down = byDiff[byDiff.length - 1];
      if (up && up.diff > 0) s += ` 전기 대비 가장 크게 증가한 항목은 ${up.name}(+${won(up.diff)}원)`;
      if (down && down.diff < 0) s += `${up && up.diff > 0 ? ',' : ' 전기 대비'} 가장 크게 감소한 항목은 ${down.name}(${won(down.diff)}원)`;
      if ((up && up.diff > 0) || (down && down.diff < 0)) s += '입니다.';
    }
    out.push(s);
  }

  // 3. 영업외손익
  const nonOpDetail = arr(c.d.영업외손익명세);
  const inc = c.is('영업외수익', '당기');
  const exp = c.is('영업외비용', '당기');
  const incItems = nonOpDetail.filter((r) => r.구분 === '영업외수익' && n(r.당기) > 0).map((r) => `${r.계정과목} ${won(n(r.당기))}원`);
  const expItems = nonOpDetail.filter((r) => r.구분 === '영업외비용' && n(r.당기) > 0).map((r) => `${r.계정과목} ${won(n(r.당기))}원`);
  out.push(
    `[영업외손익] 영업외수익 ${wonUnit(inc)}${incItems.length ? `(${incItems.join(', ')})` : ''}, 영업외비용 ${wonUnit(exp)}${
      expItems.length ? `(${expItems.join(', ')})` : ''
    }이 반영되어 영업외손익은 ${wonUnit(inc - exp)}입니다.`,
  );

  // 4. 전년 대비 주요 변동
  if (c.hasPrev) {
    const ch = (k) => c.bs(k, '당기') - c.bs(k, '전기');
    let s = `[전년 대비 주요 변동] 자산총계 ${eok(ch('자산총계'), true)}, 부채총계 ${eok(ch('부채총계'), true)}, 자본총계 ${eok(ch('자본총계'), true)} 변동하였습니다.`;
    if (c.bs('자본총계', '전기') < 0 && c.bs('자본총계', '당기') >= 0) s += ' 당기순이익 발생으로 전기 자본잠식이 해소되었습니다.';
    out.push(s);
  }

  // 5. 재무비율
  const r = c.ratios.당기;
  let s5 = `[재무비율] 당기 부채비율은 ${isNum(r.부채비율) ? pct(r.부채비율) : r.부채비율}`;
  if (c.hasPrev) {
    const p = c.ratios.전기.부채비율;
    s5 += `(전기 ${isNum(p) ? pct(p) : p})`;
  }
  s5 += `, 유동비율은 ${isNum(r.유동비율) ? pct(r.유동비율) : r.유동비율}입니다.`;
  out.push(s5);

  // 6. 공제·감면
  const cr = c.credits.filter((x) => n(x.공제감면세액));
  out.push(
    cr.length
      ? `[공제·감면] ${cr.map((x) => `${x.구분명}${x.근거법조항 ? `(${x.근거법조항})` : ''} ${won(n(x.공제감면세액))}원`).join(', ')}이 반영되었습니다.`
      : '[공제·감면] 당기 반영된 공제·감면세액은 없습니다.',
  );
  return out;
}

function autoNotables(c) {
  const out = [];
  if (!c.has('당기')) return out;
  // 자본잠식 (2-7)
  for (const yk of ['당기', '전기']) {
    if (!c.has(yk)) continue;
    const eq = c.bs('자본총계', yk);
    const cap = c.bs('자본금', yk);
    if (isNum(eq) && eq < 0 && isNum(cap)) {
      out.push(
        `${c.years[yk]} 사업연도 자본총계 ${wonUnit(eq)}으로 완전자본잠식 상태입니다. 해소를 위해 최소 ${wonUnit(-eq)}(자본금 회복 기준 ${wonUnit(cap - eq)})의 자본 확충(가수금 출자전환 등)이 필요합니다.`,
      );
    }
  }
  if (c.hasPrev && c.bs('자본총계', '전기') < 0 && c.bs('자본총계', '당기') >= 0) {
    out.push(`당기 자본총계 ${wonUnit(c.bs('자본총계', '당기'))}으로 전기 자본잠식이 해소되었습니다.`);
  }
  // 판관비 신규·소멸 (Q-05)
  if (c.sgaNew.length) out.push(`판관비 신규 발생 항목: ${c.sgaNew.map((r) => `${r.name} ${won(r.cur)}원`).join(', ')}`);
  if (c.sgaGone.length) out.push(`판관비 소멸 항목: ${c.sgaGone.map((r) => `${r.name}(전기 ${won(r.prev)}원)`).join(', ')}`);
  // 이월결손금 차이 (Q-08)
  for (const r of arr(c.d.이월결손금)) {
    if (String(r.비고 || '').includes('차이')) out.push(`이월결손금(${r.발생사업연도 ?? ''}) ${won(n(r.발생액))}원: ${r.비고}`);
  }
  return out;
}
