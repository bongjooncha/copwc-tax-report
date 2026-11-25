// 계산값 → 보고서 블록 모델
// 모든 표는 출력예시 양식과 같은 28칸 격자(B~AC)를 기준으로 span을 지정함.
// 같은 모델을 xlsx 렌더러와 HTML(PDF) 렌더러가 함께 사용함.
import { YK } from './compute.js';
import { RULES, fmtDateKo } from './format.js';

export const GRID = 28;
// 출력예시 xlsx 열 너비 (B~AC)
export const COL_WIDTHS = [
  6.83, ...Array(10).fill(5.16), ...Array(11).fill(6.16), 7.33, ...Array(5).fill(8.33),
];

const H = (v, s, extra = {}) => ({ v, s, st: 'head', ...extra });
const L = (v, s, extra = {}) => ({ v, s, st: 'label', ...extra });
const V = (v, s, fmt = 'won', extra = {}) => ({ v, s, st: 'val', fmt, ...extra });
const T = (v, s, extra = {}) => ({ v, s, st: 'val', fmt: 'text', ...extra });

export function buildModel(c) {
  const B = [];
  const yLabel = (yk) => (c.years[yk] ? `${c.years[yk]} 사업연도` : '-');
  const yv = (yk, v) => (c.has(yk) ? (v ?? '확인 필요') : '-');
  const info = c.d.법인기본사항 || {};

  B.push({ t: 'title', text: c.report.title });
  B.push({ t: 'meta', text: `${info.법인명 || ''}${c.years.당기 ? ` · ${c.years.당기} 사업연도` : ''}` });

  // ── Ⅰ. 법인 기본사항 ──
  B.push({ t: 'section', text: 'Ⅰ. 법인 기본사항' });
  B.push({
    t: 'grid',
    rows: [
      [L('법인명', 5), T(info.법인명, 9), L('사업자등록번호', 5), T(info.사업자등록번호, 9)],
      [L('대표자명', 5), T(info.대표자명, 9), L('법인등록번호', 5), T(info.법인등록번호, 9)],
      [L('사업장 소재지', 5), T(info.사업장소재지, 23)],
      [L('업 태', 3), T(info.업태, 6), L('종 목', 3), T(info.종목, 7), L('주업종코드', 3), T(info.주업종코드, 6)],
      [L('법인구분', 3), T(info.법인구분, 6), L('법인종류', 3), T(info.법인종류, 7), L('중소기업여부', 3), T(info.중소기업여부, 6)],
      [L('사업연도', 3), T(info.사업연도, 6), L('시작일자', 3), T(info.시작일자, 7), L('종료일자', 3), T(info.종료일자, 6)],
    ],
  });
  B.push({ t: 'note', mark: '*', lines: ['법인 기본정보는 해당연도 법인세 신고내역을 토대로 작성되었습니다.'] });

  // ── Ⅱ. 연도별 표준재무제표 ──
  B.push({ t: 'section', text: 'Ⅱ. 연도별 표준재무제표' });
  B.push({ t: 'note', mark: '*', lines: ['표준재무제표는 세무조정계산서의 법인세 신고내역을 토대로 작성되었습니다.'] });
  B.push({ t: 'sub', text: '(1) 요약 재무제표' });
  const yearHead = (title) => [H(title, 8), H('코드', 2), ...YK.map((yk) => H(yLabel(yk), 6))];
  const bsRows = [
    ['Ⅰ.유동자산', '유동자산', 1], ['(1) 당좌자산', '당좌자산', 2], ['(2) 재고자산', '재고자산', 44],
    ['(3) 기타유동자산', '기타유동자산', 67], ['Ⅱ.비유동자산', '비유동자산', 80], ['(1) 투자자산', '투자자산', 81],
    ['(2) 유형자산', '유형자산', 110], ['(3) 무형자산', '무형자산', 169], ['(4) 기타비유동자산', '기타비유동자산', 194],
    ['자산총계(Ⅰ+Ⅱ)', '자산총계', 228, true], ['Ⅰ.유동부채', '유동부채', 229], ['Ⅱ.비유동부채', '비유동부채', 284],
    ['부채총계(Ⅰ+Ⅱ)', '부채총계', 333, true], ['Ⅲ.자본금', '자본금', 334], ['Ⅳ.자본잉여금', '자본잉여금', 337],
    ['Ⅴ.자본조정', '자본조정', 348], ['Ⅵ.기타포괄손익누계액', '기타포괄손익누계액', 361], ['Ⅶ.이익잉여금', '이익잉여금', 372],
    ['자본총계(Ⅲ+~Ⅶ)', '자본총계', 382, true], ['부채와자본총계', '부채와자본총계', 383, true],
  ];
  B.push({
    t: 'grid',
    rows: [
      [H('표준대차대조표', 28)],
      yearHead('구 분'),
      ...bsRows.map(([label, key, code, bold]) => [
        T(label, 8, { bold }), T(code, 2, { al: 'center' }), ...YK.map((yk) => V(yv(yk, c.bs(key, yk)), 6, 'won', { bold })),
      ]),
    ],
  });
  const isRows = [
    ['Ⅰ.매출액', '매출액', 1], ['Ⅱ.매출원가', '매출원가', 35], ['Ⅲ.매출총손익', '매출총손익', 66, true],
    ['Ⅳ.판매비와관리비', '판매비와관리비', 67], ['Ⅴ.영업손익', '영업손익', 129, true], ['Ⅵ.영업외수익', '영업외수익', 130],
    ['Ⅶ.영업외비용', '영업외비용', 179], ['Ⅷ.법인세비용차감전손익', '법인세비용차감전손익', 217, true],
    ['Ⅸ.법인세비용', '법인세비용', 218], ['Ⅹ.당기순손익', '당기순손익', 219, true],
  ];
  B.push({
    t: 'grid',
    rows: [
      [H('표준손익계산서', 28)],
      yearHead('구 분'),
      ...isRows.map(([label, key, code, bold]) => [
        T(label, 8, { bold }), T(code, 2, { al: 'center' }), ...YK.map((yk) => V(yv(yk, c.is(key, yk)), 6, 'won', { bold })),
      ]),
    ],
  });

  B.push({ t: 'sub', text: '(2) 재무비율 분석' });
  const ratioDef = [
    ['부채비율', '%', '200이하', 'pct'], ['유동비율', '%', '130이상', 'pct'], ['당좌비율', '%', '100이상', 'pct'],
    ['총자산세전순이익률', '%', '10이상', 'pct'], ['영업이익률', '%', '10이상', 'pct'],
    ['총자산회전율', '회', '1.5회', 'times'], ['자기자본회전율', '회', '3회', 'times'],
  ];
  B.push({
    t: 'grid',
    rows: [
      [H('구 분', 8), H('단위', 2), H('표준비율', 3), ...YK.map((yk) => H(yLabel(yk), 5))],
      ...ratioDef.map(([k, unit, std, fmt]) => [
        T(k, 8), T(unit, 2, { al: 'center' }), T(std, 3, { al: 'center' }),
        ...YK.map((yk) => V(c.has(yk) ? (c.ratios[yk][k] ?? '확인 필요') : '-', 5, fmt, { al: 'center' })),
      ]),
    ],
  });
  const ratioNotes = [
    '「표준비율」은 업종의 특성을 고려하지 아니한 비율로 일반적으로 재무상태가 양호함을 나타내는 비율입니다.',
    `자본총계가 음수인 연도는 부채비율을 「${RULES.capitalDeficit}」, 자기자본회전율을 「-」로 표기합니다.`,
  ];
  if (c.turnoverEndOnly.length) {
    ratioNotes.push(`기초 잔액 없음, 기말 기준 (${c.turnoverEndOnly.map((yk) => c.years[yk]).join(', ')} 사업연도 회전율)`);
  }
  B.push({ t: 'note', mark: '*', lines: ratioNotes });
  B.push({
    t: 'note',
    mark: '**',
    lines: [
      '비율산정공식',
      '부채비율 : 부채/자기자본 * 100',
      '유동비율 : 유동자산/유동부채 * 100',
      '당좌비율 : 당좌자산/유동부채 * 100',
      '총자산세전순이익률 : 세전이익/총자산 * 100',
      '영업이익률 : 영업이익/매출액 * 100',
      '총자산회전율 : 매출액/((기초총자산+기말총자산)/2)',
      '자기자본회전율 : 매출액/((기초자기자본+기말자기자본)/2)',
    ],
  });

  // ── Ⅲ. 연도별 법인세 신고상황 ──
  B.push({ t: 'section', text: 'Ⅲ. 연도별 법인세 신고상황' });
  B.push({ t: 'note', mark: '*', lines: ['세무조정계산서의 법인세 신고내역을 토대로 작성되었습니다.'] });
  const lossDiff = c.lossRows.some((r) => String(r.비고 || '').includes('차이'));
  B.push({
    t: 'grid',
    rows: [
      [H('구 분', 8), ...YK.map((yk) => H(yLabel(yk), 5)), H(`${c.active.length}개년 평균`, 5)],
      ...c.taxRows.map((r) => [
        T(r.label + (lossDiff && r.label === '(-)이월결손금' ? ' ※' : ''), 8),
        ...YK.map((yk) => V(c.has(yk) ? (r.vals[yk] ?? '확인 필요') : '-', 5, r.fmt, r.fmt === 'num' ? { al: 'center' } : {})),
        V(r.avg ?? '', 5, r.fmt),
      ]),
      ...c.rateRows.map((r) => [
        T(r.label, 8),
        ...YK.map((yk) => V(c.has(yk) ? r.vals[yk] : '-', 5, 'pct', { al: 'center', st: 'hl' })),
        V(r.avg, 5, 'pct', { al: 'center', st: 'hl' }),
      ]),
    ],
  });
  B.push({
    t: 'note',
    mark: '*',
    lines: [
      '소득률 = 각사업연도 소득금액/수입금액 * 100 [법인세신고기준]',
      '매출총이익률 = 매출총이익/매출액 * 100 [손익계산서 기준]',
      '영업이익률 = 영업이익/매출액 * 100 [손익계산서 기준]',
      '실효세율(과세표준 기준) = 총부담세액/과세표준 * 100',
      '실효세율(소득금액 기준) = 총부담세액/각사업연도소득금액 * 100',
      `평균은 자료가 있는 연도(${c.active.length}개년) 기준입니다.`,
    ],
  });
  if (lossDiff) B.push({ t: 'note', mark: '※', lines: ['이월결손금: 결산서 결손금과 이월결손금 발생액에 차이 있음 (Ⅳ-(2) 참조)'] });

  // ── Ⅳ. 법인세 신고 참고 자료 ──
  B.push({ t: 'section', text: 'Ⅳ. 법인세 신고 참고 자료' });
  B.push({ t: 'sub', text: `(1) ${c.years.당기 ?? ''}사업연도 기납부세액` });
  const pp = c.prepaidItems;
  const ppSpans = [6, 5, 6, 6, 5];
  B.push({
    t: 'grid',
    rows: [
      [...pp.map(([k], i) => H(k, ppSpans[i])), H('합계', 5)],
      [...pp.map(([, v], i) => V(v ?? '확인 필요', ppSpans[i])), V(pp.every(([, v]) => typeof v === 'number') ? pp.reduce((s, [, v]) => s + v, 0) : '확인 필요', 5)],
    ],
  });

  B.push({ t: 'sub', text: '(2) 이월결손금 현황' });
  const lossSum = (k) => c.lossRows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  B.push({
    t: 'grid',
    rows: [
      [H('발생 사업연도', 6), H('발생액', 6), H('감소액', 6), H('잔액', 5), H('비고', 5)],
      ...(c.lossRows.length
        ? c.lossRows.map((r) => [
            T(r.발생사업연도, 6, { al: 'center' }), V(num(r.발생액), 6), V(num(r.감소액), 6), V(num(r.잔액), 5),
            T(r.비고 || '', 5, { small: true, wrap: true }),
          ])
        : [[T('해당 없음', 28, { al: 'center' })]]),
      [L('합계', 6), V(lossSum('발생액'), 6), V(lossSum('감소액'), 6), V(lossSum('잔액'), 5), T('', 5)],
    ],
  });
  B.push({ t: 'note', mark: '*', lines: [`${c.years.당기 ?? ''}사업연도 법인세 신고내역(자본금과적립금조정명세서(갑))을 토대로 작성되었습니다.`] });

  B.push({ t: 'sub', text: '(3) 해당연도 소득금액조정합계표 현황' });
  const adjN = Math.max(c.adjPlus.length, c.adjMinus.length, 1);
  const adjSum = (a) => a.reduce((s, r) => s + (Number(r.금액) || 0), 0);
  const adjCells = (r) =>
    r ? [T(r.과목, 6), V(num(r.금액), 5), T(r.소득처분, 3, { al: 'center', small: true })] : [T('', 6), T('', 5), T('', 3)];
  const adjCellsR = (r) =>
    r ? [T(r.과목, 6), V(num(r.금액), 5), T(r.소득처분, 3, { al: 'center', small: true })] : [T('', 6), T('', 5), T('', 3)];
  B.push({
    t: 'grid',
    rows: [
      [H('익금산입 및 손금불산입', 14), H('손금산입 및 익금불산입', 14)],
      [H('과 목', 6), H('금 액', 5), H('소득처분', 3), H('과 목', 6), H('금 액', 5), H('소득처분', 3)],
      ...Array.from({ length: adjN }, (_, i) => [...adjCells(c.adjPlus[i]), ...adjCellsR(c.adjMinus[i])]),
      [L('합 계', 6), V(adjSum(c.adjPlus), 5), T('', 3), L('합 계', 6), V(adjSum(c.adjMinus), 5), T('', 3)],
    ],
  });

  B.push({ t: 'sub', text: '(4) 해당연도 신청한 공제·감면세액 내역' });
  const crSum = c.credits.reduce((s, r) => s + (Number(r.공제감면세액) || 0), 0);
  B.push({
    t: 'grid',
    rows: [
      [H('공제·감면 구분명', 11), H('근거 법조항', 11), H('공제·감면세액', 6)],
      ...(c.credits.length
        ? c.credits.map((r) => [T(r.구분명, 11), T(r.근거법조항, 11), V(num(r.공제감면세액), 6)])
        : [[T('해당 없음', 28, { al: 'center' })]]),
      [L('합 계', 22), V(crSum, 6)],
    ],
  });

  B.push({ t: 'sub', text: '(5) 재고율' });
  const inv = c.inventory;
  const invSpan = inv.length === 2 ? [10, 9, 9] : [10, 18];
  B.push({
    t: 'grid',
    rows: [
      [H('구 분', invSpan[0]), ...inv.map((r, i) => H(yLabel(r.yk), invSpan[i + 1]))],
      [T('기말 재고자산', invSpan[0]), ...inv.map((r, i) => V(r.inv ?? '확인 필요', invSpan[i + 1]))],
      [T('매출원가', invSpan[0]), ...inv.map((r, i) => V(r.cogs ?? '확인 필요', invSpan[i + 1]))],
      [T('재고율(%)', invSpan[0]), ...inv.map((r, i) => V(r.rate, invSpan[i + 1], 'pct', { al: 'center' }))],
    ],
  });
  B.push({ t: 'note', mark: '*', lines: ['재고율 = 재고자산/매출원가 * 100 (매출원가가 0인 연도는 「-」)'] });

  B.push({ t: 'sub', text: '(6) 4대보험 회사부담분  (지원예정)' });
  B.push({
    t: 'grid',
    rows: [
      [H('국민연금', 5), H('건강보험(장기요양 포함)', 9), H('고용보험', 4), H('산재보험', 5), H('합 계', 5)],
      [T('', 5), T('', 9), T('', 4), T('', 5), T('', 5)],
    ],
  });
  B.push({ t: 'note', mark: '*', lines: ['4대보험 회사부담분은 향후 지원될 예정입니다.'] });

  // ── Ⅴ. 법인세 신고 결과 보고 ──
  B.push({ t: 'section', text: 'Ⅴ. 법인세 신고 결과 보고' });
  B.push({ t: 'note', mark: '*', lines: ['세무조정계산서의 법인세 신고내역을 토대로 작성된 법인세 신고서식 제출 목록과 수행 업무내역입니다.'] });
  B.push({ t: 'sub', text: '0. 안내사항' });
  const guide = c.d.안내사항 || {};
  for (const k of ['세무조정사항', '회계사항', '감면사항']) {
    const lines = (Array.isArray(guide[k]) ? guide[k] : []).map(String).filter((s) => s.trim());
    B.push({ t: 'para', head: `[${k}]`, lines: lines.length ? lines.map((s) => `- ${s}`) : [k === '회계사항' ? '' : '이슈없음'] });
  }
  const filing = c.d.신고내용 || {};
  B.push({ t: 'sub', text: '(1) 법인세 신고 내용' });
  B.push({
    t: 'para',
    lines: [
      filing.신고서명 || '',
      `과세기간: ${filing.과세기간 || ''}`,
      `신고유형(신고일자): ${filing.신고유형 || ''}${filing.신고일자 ? `(${filing.신고일자})` : ''}`,
    ],
  });
  B.push({ t: 'sub', text: '(2) 신고서식 제출 목록 및 수행업무 명세' });
  const forms = Array.isArray(c.d.제출서식목록) ? c.d.제출서식목록 : [];
  B.push({ t: 'forms', items: forms.map((f) => ({ name: f.서식명 || '', desc: f.수행업무 || '' })) });

  // ── Ⅵ. 주요 재무지표 요약 ──
  B.push({ t: 'section', text: 'Ⅵ. 주요 재무지표 요약', pageBreak: true });
  B.push({ t: 'sub', text: '(1) 핵심지표' });
  B.push({ t: 'cards', cards: c.cards });
  B.push({ t: 'note', mark: '*', lines: ['금액 단위: 1억원 이상은 억원, 1억원 미만은 천만원 (상세 표는 원 단위)'] });
  B.push({ t: 'sub', text: '(2) 요약 코멘트' });
  B.push({ t: 'para', lines: c.comments.map((s, i) => `${i + 1}. ${s}`), wrap: true });
  B.push({ t: 'sub', text: '(3) 그래프' });
  B.push({ t: 'charts', keys: ['pl', 'sgaPie', 'sgaBar', 'bs'] });

  // ── Ⅶ. 손익 분석 ──
  B.push({ t: 'section', text: 'Ⅶ. 손익 분석' });
  B.push({ t: 'sub', text: '(1) 요약손익 비교 (단위: 원)' });
  const plSpan = [7, 5, 5, 5, 6];
  B.push({
    t: 'grid',
    rows: [
      [H('항 목', 7), H(c.hasPrev ? yLabel('전기') : '전기', 5), H(yLabel('당기'), 5), H('증감액', 5), H('증감률(%)', 6)],
      ...c.plRows.map((r) => [
        T(r.label, plSpan[0]), V(c.hasPrev ? r.prev : '-', 5), V(r.cur, 5), V(c.hasPrev ? r.diff : '-', 5),
        V(r.rate, 6, 'pct', { al: 'center', st: 'hl' }),
      ]),
    ],
  });
  B.push({ t: 'note', mark: '*', lines: ['증감률 = (당기 − 전기) ÷ |전기| × 100. 전기 ≤ 0 이고 당기 > 0인 이익 항목은 「흑자전환」, 그 밖에 전기 ≤ 0이면 「-」'] });
  B.push({ t: 'sub', text: '(2) 판관비 주요항목 (단위: 원)' });
  B.push({
    t: 'grid',
    rows: [
      [H('항 목', 7), H(c.hasPrev ? yLabel('전기') : '전기', 5), H(yLabel('당기'), 5), H('증감액', 5), H('매출액 대비(%)', 6)],
      ...(c.sgaRows.length
        ? c.sgaRows.map((r) => [
            T(r.name, 7), V(c.hasPrev ? r.prev : '-', 5), V(r.cur, 5), V(c.hasPrev ? r.diff : '-', 5), V(r.share, 6, 'pct', { al: 'center' }),
          ])
        : [[T('해당 없음', 28, { al: 'center' })]]),
      [
        L('합 계', 7), V(c.hasPrev ? c.sgaTotal.prev : '-', 5), V(c.sgaTotal.cur, 5),
        V(c.hasPrev ? c.sgaTotal.cur - c.sgaTotal.prev : '-', 5),
        V(c.salesCur ? (c.sgaTotal.cur / c.salesCur) * 100 : '-', 6, 'pct', { al: 'center' }),
      ],
    ],
  });
  B.push({ t: 'note', mark: '*', lines: ['당기 금액 상위 5개 항목, 나머지는 「기타」로 합산. 매출액 대비 비중은 당기 매출액 기준'] });

  // ── Ⅷ. 재무제표 주요 특이사항 ──
  B.push({ t: 'section', text: 'Ⅷ. 재무제표 주요 특이사항' });
  B.push({ t: 'para', lines: c.notables.length ? c.notables.map((s) => `- ${s}`) : ['- 특이사항 없음'], wrap: true });

  const warnings = Array.isArray(c.d.확인필요?.경고) ? c.d.확인필요.경고 : [];
  if (warnings.length) {
    B.push({ t: 'sub', text: '(2) 자료 간 차이 및 확인 사항' });
    for (const w of warnings) B.push({ t: 'para', head: `[${w.수준 || '확인'}] ${w.항목 || w.유형 || ''}`, lines: [w.내용, w.표시값 ? `표시값: ${w.표시값}` : '', w.파일 ? `출처: ${w.파일} ${w.페이지 || ''}` : ''].filter(Boolean) });
  }
  const donation = c.d.소득금액조정?.기부금한도초과?.한도초과액;
  if (donation != null) B.push({ t: 'note', mark: '*', lines: [`기부금 한도초과액: ${Number(donation).toLocaleString('ko-KR')}원 (세액조정 시 별도 반영)`] });

  // ── 서명란 (Q-06, T-08) ──
  B.push({ t: 'sign', date: fmtDateKo(c.report.issueDate), issuer: c.report.issuer });
  return B;
}

function num(v) {
  if (typeof v === 'number') return v;
  const x = Number(String(v ?? '').replace(/,/g, ''));
  return String(v ?? '').trim() !== '' && Number.isFinite(x) ? x : v ?? '';
}
