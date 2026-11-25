// 입력 JSON + 출력예시 템플릿(xlsx) → 보고서 워크북
// - 템플릿 1~122행(Ⅰ~Ⅲ, Ⅳ 머리)을 그대로 쓰고, Ⅳ-(1) 부가세·(2) 원천세(템플릿 123~220행)는 빼고(Q-07),
//   템플릿 221~290행(Ⅳ-(3)~Ⅴ-(1))을 이어 붙임. 목록 칸은 항목 수만큼 행을 늘림
// - 비율·합계·평균은 Excel 수식으로 넣고 계산 결과도 함께 저장 (담당자가 값을 고치면 같이 바뀜)
// - Ⅵ·Ⅶ·Ⅷ·서명란 추가 (T-02, Q-09, Q-06/T-08)
// ExcelJS 모듈은 호출하는 쪽에서 넘김 (브라우저: dist 번들, Node 검증 스크립트: 패키지)
import { setCellValue, cellResult } from './formulaResult.js';

const FONT = '맑은 고딕';
const NAVY = 'FF1F2D40';
const THIN = { style: 'thin', color: { argb: 'FF000000' } };
const BOX = { left: THIN, right: THIN, top: THIN, bottom: THIN };
const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const HEAD_FILL = solid('FFE8E4DC');
const NA = '-';
const LAST_COL = 29; // AC
const COLW = [1.1640625, 6.83203125, ...Array(10).fill(5.1640625), ...Array(11).fill(6.1640625), 7.33203125, ...Array(5).fill(8.33203125)];
export const CHART_CM = { w: 15.5, h: 8.2 };

// ── 숫자·문자 도우미 ───────────────────────────────
export const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const num = (v) => {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v.replace(/,/g, '')))) return Number(v.replace(/,/g, ''));
  return null;
};
const arr = (v) => (Array.isArray(v) ? v : []);
export const TBD = '확인 필요';
const isN = (v) => typeof v === 'number' && Number.isFinite(v);
const allN = (...vs) => vs.every(isN);
// 입력값 정규화: 숫자 / "-" / 그 밖의 문자(빈칸은 "확인 필요")
function val(v) {
  const n = num(v);
  if (n !== null) return n;
  if (v === null || v === undefined || String(v).trim() === '') return TBD;
  const s = String(v).trim();
  return s === '-' ? '-' : s;
}
const isTbd = (v) => typeof v === 'string' && v !== '-' && v !== '';
const TBD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const WARN_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
const RED = 'FFC00000';
const noteText = (v) => (typeof v === 'string' ? v : v && typeof v === 'object' ? String(v.내용 ?? v.text ?? '') : '');
const isAlertLine = (s) => /확인 필요|차이/.test(s);
const clone = (o) => (o === undefined ? o : JSON.parse(JSON.stringify(o)));
function group(intStr) {
  return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
function fixed(v, d) {
  const s = Math.abs(v).toFixed(d);
  const [i, f] = s.split('.');
  return `${v < 0 && Number(s) !== 0 ? '-' : ''}${group(i)}${f ? `.${f}` : ''}`;
}
const won = (v) => `${fixed(v, 0)}원`;
function unit(v, sign = false) {
  const s = sign && v > 0 ? '+' : '';
  if (Math.abs(v) >= 1e8) return `${s}${fixed(v / 1e8, 2)}억원`;
  return `${s}${fixed(v / 1e7, 1)}천만원`;
}
const pct = (a, b) => (b ? (a / b) * 100 : null);
function chgRate(cur, prev) {
  if (prev <= 0) return cur > 0 ? '흑자전환' : NA;
  return `${fixed(((cur - prev) / Math.abs(prev)) * 100, 2)}%`;
}
function xround(x, d = 0) {
  const m = 10 ** d;
  return (Math.sign(x) * Math.round(Math.abs(x) * m + 1e-9)) / m;
}
const linesNeeded = (s, per) => String(s).split('\n').reduce((n, p) => n + Math.max(1, Math.ceil([...p].length / per)), 0);
function wrapText(s, width) {
  const out = [];
  let cur = '';
  for (let w of String(s).split(/\s+/).filter(Boolean)) {
    while ([...w].length > width) {
      const room = cur ? width - cur.length - 1 : width;
      if (room > 0 && cur) { out.push(`${cur} ${w.slice(0, room)}`); w = w.slice(room); cur = ''; continue; }
      if (cur) { out.push(cur); cur = ''; }
      out.push(w.slice(0, width));
      w = w.slice(width);
    }
    if (!w) continue;
    if (!cur) cur = w;
    else if (cur.length + 1 + w.length <= width) cur += ` ${w}`;
    else { out.push(cur); cur = w; }
  }
  if (cur) out.push(cur);
  return out.length ? out : [''];
}
function dateKo(s) {
  const m = String(s || '').match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/);
  const d = m ? { y: m[1], m: m[2], d: m[3] } : null;
  const now = new Date();
  const [Y, M, D] = d ? [d.y, d.m, d.d] : [now.getFullYear(), now.getMonth() + 1, now.getDate()];
  return `${Y}년 ${String(M).padStart(2, '0')}월 ${String(D).padStart(2, '0')}일`;
}

const BS_KEYS = {
  1: '유동자산', 2: '당좌자산', 44: '재고자산', 67: '기타유동자산', 80: '비유동자산', 81: '투자자산', 110: '유형자산',
  169: '무형자산', 194: '기타비유동자산', 228: '자산총계', 229: '유동부채', 284: '비유동부채', 333: '부채총계', 334: '자본금',
  337: '자본잉여금', 348: '자본조정', 361: '기타포괄손익누계액', 372: '이익잉여금', 382: '자본총계', 383: '부채와자본총계',
};
const IS_KEYS = {
  1: '매출액', 35: '매출원가', 66: '매출총손익', 67: '판매비와관리비', 129: '영업손익', 130: '영업외수익', 179: '영업외비용',
  217: '법인세비용차감전손익', 218: '법인세비용', 219: '당기순손익',
};
const TAX_ROWS = {
  93: '익금조정', 94: '손금조정', 95: '각사업연도소득금액', 96: '이월결손금', 97: '비과세소득', 98: '소득공제', 99: '과세표준금액',
  100: '세율(%)', 101: '산출세액', 102: '최저한세적용대상공제감면세액', 103: '최저한세적용제외공제감면세액', 104: '총부담세액',
  105: '기납부세액', 106: '차감납부할세액', 107: '법인지방소득세',
};
const YK = ['당기', '전기', '전전기'];
const OLDER = { 당기: '전기', 전기: '전전기', 전전기: null };
const C_BS = { 당기: 'L', 전기: 'R', 전전기: 'X' }; // Ⅱ-(1)
const C_RT = { 당기: 'O', 전기: 'T', 전전기: 'Y' }; // Ⅱ-(2)
const C_TX = { 당기: 'J', 전기: 'O', 전전기: 'T' }; // Ⅲ

/**
 * @returns {{ wb, layout: { chartRow, chartRows, breaks, lastRow, sheetName } }}
 */
export async function buildWorkbook(ExcelJS, templateBuf, data, images = {}) {
  const d = data && typeof data === 'object' ? data : {};
  const tpl = new ExcelJS.Workbook();
  await tpl.xlsx.load(templateBuf);
  const src = tpl.worksheets[0];

  // ── 입력값 ──────────────────────────────────────
  const yrs = d.사업연도 || {};
  const has = (yk) => yrs[yk] !== null && yrs[yk] !== undefined && yrs[yk] !== '';
  const curYear = num(yrs.당기) ?? new Date().getFullYear() - 1;
  const ylabel = (yk) => (has(yk) ? String(yrs[yk]) : String(curYear - (yk === '전기' ? 1 : yk === '전전기' ? 2 : 0)));
  const CUR = ylabel('당기');
  const PRV = ylabel('전기');
  const PP = ylabel('전전기');
  const gisu = d.기수 || {};
  // 금액: 숫자 → 숫자, "-" → "-", 빈칸·문자("확인 필요" 등) → 문자 그대로 (0으로 바꾸지 않음)
  const pick = (table) => (key, yk) => {
    if (!has(yk)) return null;
    const row = (d[table] || {})[key];
    const v = row && typeof row === 'object' ? row[yk] : undefined;
    return val(v);
  };
  const bsk = pick('표준대차대조표');
  const isk = pick('표준손익계산서');
  const txk = pick('법인세신고상황');
  const BS = (code, yk) => bsk(BS_KEYS[code], yk);
  const IS = (code, yk) => isk(IS_KEYS[code], yk);
  const hp = has('전기');
  const info = d.법인기본사항 || {};
  const issuer = d.보고서정보?.발행기관 || '샘플세무법인';

  // 판관비 세부 (당기, 전기)
  const SGA = arr(d.판매비와관리비명세)
    .filter((r) => r && r.계정과목)
    .map((r) => ({ name: String(r.계정과목), cur: num(r.당기) ?? 0, prev: num(r.전기) ?? 0 }));
  const lossRows = arr(d.이월결손금);
  const lossIsPrev = hp && IS(219, '전기') < 0 ? -IS(219, '전기') : null; // 전기 손익계산서 결손금
  const lossCf = lossRows
    .filter((r) => String(r.발생사업연도 ?? '').startsWith(PRV))
    .reduce((s, r) => s + (num(r.발생액) ?? 0), 0);
  const lossDiff = lossIsPrev !== null && lossCf > 0 && lossCf !== lossIsPrev;

  // ── 새 시트 ─────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = 'COPWC';
  wb.created = new Date();
  wb.calcProperties = { fullCalcOnLoad: true };
  const ws = wb.addWorksheet('법인세 신고결과', {
    views: [{ showGridLines: false, topLeftCell: 'A1' }],
    properties: { defaultRowHeight: src.properties.defaultRowHeight || 15 },
  });
  COLW.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  const srcMerges = Object.values(src._merges).map((m) => m.model);
  const rowMap = {}; // 템플릿 행 → 새 시트 행

  function copyBlock(r1, r2, dst0) {
    const off = dst0 - r1;
    for (let r = r1; r <= r2; r += 1) {
      const h = src.getRow(r).height;
      if (h) ws.getRow(r + off).height = h;
      rowMap[r] = r + off;
    }
    for (const m of srcMerges) {
      if (m.top >= r1 && m.bottom <= r2 && m.right <= LAST_COL) ws.mergeCells(m.top + off, m.left, m.bottom + off, m.right);
    }
    for (let r = r1; r <= r2; r += 1) {
      for (let c = 1; c <= LAST_COL; c += 1) {
        const sc = src.getCell(r, c);
        const dc = ws.getCell(r + off, c);
        if (sc.style && Object.keys(sc.style).length) dc.style = clone(sc.style);
        if (sc.isMerged && sc.master !== sc) continue;
        const v = sc.value;
        if (v !== null && v !== undefined) setCellValue(dc, clone(v));
      }
    }
    return r2 + off + 1;
  }
  // 목록 칸: 템플릿 slot 행 스타일로 n행 생성 (값은 비움)
  function copySlots(slotRows, n, dst0) {
    for (let i = 0; i < n; i += 1) {
      const sr = slotRows[Math.min(i, slotRows.length - 1)];
      copyBlock(sr, sr, dst0 + i);
      for (let c = 1; c <= LAST_COL; c += 1) {
        const cell = ws.getCell(dst0 + i, c);
        if (!(cell.isMerged && cell.master !== cell)) setCellValue(cell, null);
      }
    }
    rowMap[slotRows[0]] = dst0;
    return dst0 + n;
  }

  const set = (addr, value, o = {}) => {
    const c = ws.getCell(addr);
    setCellValue(c, value);
    if (o.fmt) c.numFmt = o.fmt;
    if (o.al) c.alignment = { horizontal: o.al, vertical: 'middle' };
    if (o.font) c.font = o.font;
    return c;
  };
  const F = (addr, formula, result, o) => set(addr, result === undefined ? { formula } : { formula, result }, o);
  const font = (o = {}) => ({ name: FONT, size: 10, ...o, color: { argb: o.color || 'FF000000' } });
  // "확인 필요" 칸: 노란 음영 + 빨간 굵은 글씨 (계산 정지 표시)
  const markTbd = (addr, text = TBD) => {
    const c = ws.getCell(addr);
    setCellValue(c, text);
    c.fill = clone(TBD_FILL);
    c.font = { name: FONT, size: 9, bold: true, color: { argb: RED } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    return c;
  };
  // 값 칸: 숫자면 숫자, "-"면 "-", 그 밖의 문자면 "확인 필요" 표시
  const putVal = (addr, v, o = {}) => {
    if (isTbd(v)) return markTbd(addr, v);
    const c = set(addr, v ?? NA, o);
    return c;
  };
  // 수식 칸: 의존값이 모두 숫자일 때만 수식, 아니면 "확인 필요"
  const FX = (addr, formula, deps, result, o) => (allN(...deps) ? F(addr, formula, typeof result === 'function' ? result() : result, o) : markTbd(addr));
  // 주석(d.주석[키]) — 문자열 또는 문자열 배열
  const notes = (key) => {
    const v = (d.주석 || {})[key];
    return (Array.isArray(v) ? v : v ? [v] : []).map(noteText).filter((s) => s.trim());
  };
  const noteFont = (s) => font({ size: 9, color: isAlertLine(s) ? RED : 'FF000000' });
  // 셀표시(d.셀표시): [{ 위치: "표.항목.연도", 종류: "경고"|"확인필요", 메모 }]
  const marks = arr(d.셀표시);
  const applyMark = (addr, m) => {
    const c = ws.getCell(addr);
    c.fill = clone(String(m.종류 || '').includes('경고') ? WARN_FILL : TBD_FILL);
    if (m.메모) c.note = String(m.메모);
  };
  const marksFor = (table) => marks.filter((m) => String(m.위치 || '').startsWith(`${table}.`)).map((m) => {
    const [, key, yk] = String(m.위치).split('.');
    return { key, yk, m };
  });

  // ── 1~122행 (Ⅰ~Ⅳ 머리) ───────────────────────────
  copyBlock(1, 122, 1);

  // 부제
  const basis = d.보고서정보?.기준 ? `  |  ${d.보고서정보.기준}` : '';
  set('B4', `${info.법인명 || ''}  |  ${CUR} 사업연도${gisu.당기 ? `(${gisu.당기})` : ''} 신고분${hp ? `  |  전기 ${PRV} 사업연도${gisu.전기 ? `(${gisu.전기})` : ''} 비교` : ''}${basis}`, {
    font: font({ color: 'FF595959' }),
  });

  // Ⅰ. 법인 기본사항
  for (const [addr, key] of [
    ['G8', '법인명'], ['U8', '사업자등록번호'], ['G9', '대표자명'], ['U9', '법인등록번호'], ['G10', '사업장소재지'],
    ['G11', '업태'], ['P11', '종목'], ['Z11', '주업종코드'], ['G12', '법인구분'], ['P12', '법인종류'], ['Z12', '중소기업여부'],
    ['G13', '사업연도'], ['P13', '시작일자'], ['Z13', '종료일자'],
  ]) {
    const c = ws.getCell(addr);
    c.value = info[key] ?? '';
    c.font = { name: FONT, size: 10 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  }
  const INFO_ADDR = { 법인명: 'G8', 사업자등록번호: 'U8', 대표자명: 'G9', 법인등록번호: 'U9', 사업장소재지: 'G10' };
  for (const { key, m } of marksFor('법인기본사항')) if (INFO_ADDR[key]) applyMark(INFO_ADDR[key], m);
  notes('Ⅰ').slice(0, 2).forEach((s, i) => set(`C${15 + i}`, s, { font: noteFont(s) }));

  // Ⅱ-(1) 요약 재무제표
  for (const hr of [25, 49]) {
    set(`L${hr}`, `${CUR} 사업연도`);
    set(`R${hr}`, `${PRV} 사업연도`);
    set(`X${hr}`, `${PP} 사업연도`);
  }
  const numCell = (addr, v) => {
    if (isTbd(v)) return markTbd(addr, v);
    const c = ws.getCell(addr);
    setCellValue(c, v);
    c.numFmt = '#,##0;[Red]-#,##0';
    c.alignment = { horizontal: 'right', vertical: 'middle' };
    return c;
  };
  const rowOfBs = {};
  const rowOfIs = {};
  for (let r = 26; r <= 45; r += 1) {
    const code = ws.getCell(`J${r}`).value;
    rowOfBs[BS_KEYS[code]] = r;
    for (const yk of YK) numCell(`${C_BS[yk]}${r}`, has(yk) ? BS(code, yk) : NA);
  }
  for (let r = 50; r <= 59; r += 1) {
    const code = ws.getCell(`J${r}`).value;
    rowOfIs[IS_KEYS[code]] = r;
    for (const yk of YK) numCell(`${C_BS[yk]}${r}`, has(yk) ? IS(code, yk) : NA);
  }
  for (const { key, yk, m } of marksFor('표준대차대조표')) if (rowOfBs[key] && C_BS[yk]) applyMark(`${C_BS[yk]}${rowOfBs[key]}`, m);
  for (const { key, yk, m } of marksFor('표준손익계산서')) if (rowOfIs[key] && C_BS[yk]) applyMark(`${C_BS[yk]}${rowOfIs[key]}`, m);

  // Ⅱ-(2) 재무비율 (T-01, T-04, T-05)
  set('O63', `${CUR} 사업연도`);
  set('T63', `${PRV} 사업연도`);
  set('Y63', `${PP} 사업연도`);
  const ratioCell = (addr, formula, result) => {
    F(addr, formula, result, { fmt: '#,##0.00', al: 'center' });
  };
  const firstYears = [];
  for (const yk of YK) {
    const o = C_RT[yk];
    const c = C_BS[yk];
    if (!has(yk)) {
      for (let r = 64; r <= 70; r += 1) set(`${o}${r}`, NA, { fmt: '#,##0.00', al: 'center' });
      continue;
    }
    const g = (code) => BS(code, yk);
    const r2 = (a, b) => (b === 0 ? NA : xround((a / b) * 100, 2));
    const RC = (addr, formula, deps, result) => {
      if (allN(...deps)) ratioCell(addr, formula, result());
      else markTbd(addr);
    };
    RC(`${o}64`, `IF(${c}44<=0,"자본잠식",ROUND(${c}38/${c}44*100,2))`, [g(382), g(333)], () => (g(382) <= 0 ? '자본잠식' : xround((g(333) / g(382)) * 100, 2)));
    RC(`${o}65`, `IF(${c}36=0,"-",ROUND(${c}26/${c}36*100,2))`, [g(1), g(229)], () => r2(g(1), g(229)));
    RC(`${o}66`, `IF(${c}36=0,"-",ROUND(${c}27/${c}36*100,2))`, [g(2), g(229)], () => r2(g(2), g(229)));
    RC(`${o}67`, `IF(${c}35=0,"-",ROUND(${c}57/${c}35*100,2))`, [IS(217, yk), g(228)], () => r2(IS(217, yk), g(228)));
    RC(`${o}68`, `IF(${c}50=0,"-",ROUND(${c}54/${c}50*100,2))`, [IS(129, yk), IS(1, yk)], () => r2(IS(129, yk), IS(1, yk)));
    const old = OLDER[yk] && has(OLDER[yk]) ? OLDER[yk] : null;
    const sales = IS(1, yk);
    if (old) {
      // 기초(=전년 기말) 값이 "확인 필요"면 회전율 계산 보류
      const b = C_BS[old];
      RC(`${o}69`, `ROUND(${c}50/((${b}35+${c}35)/2),2)`, [sales, BS(228, old), g(228)], () => {
        const avgA = (BS(228, old) + g(228)) / 2;
        return avgA ? xround(sales / avgA, 2) : NA;
      });
      RC(`${o}70`, `IF(${c}44<=0,"-",ROUND(${c}50/((${b}44+${c}44)/2),2))`, [sales, BS(382, old), g(382)], () => {
        const avgE = (BS(382, old) + g(382)) / 2;
        return g(382) <= 0 ? NA : xround(sales / avgE, 2);
      });
    } else {
      if (allN(sales, g(228))) firstYears.push(ylabel(yk));
      RC(`${o}69`, `ROUND(${c}50/${c}35,2)`, [sales, g(228)], () => (g(228) ? xround(sales / g(228), 2) : NA));
      RC(`${o}70`, `IF(${c}44<=0,"-",ROUND(${c}50/${c}44,2))`, [sales, g(382)], () => (g(382) <= 0 ? NA : xround(sales / g(382), 2)));
    }
  }
  set('J69', '회');
  set('J70', '회');
  const f9 = font({ size: 9 });
  const n2 = notes('Ⅱ');
  if (n2.length) {
    set('B82', '***', { font: f9 });
    n2.slice(0, 2).forEach((s, i) => set(`C${82 + i}`, s, { font: noteFont(s) }));
  } else {
    if (firstYears.length) {
      set('B82', '***', { font: f9 });
      set('C82', `${firstYears.join(', ')} 사업연도는 설립 첫해로 기초 잔액이 없어 총자산회전율을 기말 총자산 기준으로 산정하였습니다.`, { font: f9 });
    }
    set('C83', '자본총계가 0 이하(자본잠식)인 연도는 부채비율을 『자본잠식』, 자기자본회전율을 『-』로 표기합니다.', { font: f9 });
  }

  // Ⅲ. 연도별 법인세 신고상황 (T-11 평균은 자료가 있는 연도 수로)
  for (const yk of YK) set(`${C_TX[yk]}88`, `${ylabel(yk)} 사업연도`);
  set('Y88', YK.every(has) ? '3개년 평균' : '평균(자료 있는 연도)');
  const LINK = { 89: 50, 90: 52, 91: 54, 92: 59 };
  const tv = {}; // tv[row][col] = 값 (숫자 / "-" / "확인 필요")
  for (let r = 89; r <= 107; r += 1) {
    tv[r] = {};
    for (const yk of YK) {
      const col = C_TX[yk];
      const addr = `${col}${r}`;
      if (!has(yk)) {
        set(addr, NA);
        continue;
      }
      const v = LINK[r] ? IS(Number(ws.getCell(`J${LINK[r]}`).value), yk) : txk(TAX_ROWS[r], yk);
      tv[r][col] = v;
      if (isTbd(v)) markTbd(addr, v);
      else if (LINK[r] && isN(v)) F(addr, `${C_BS[yk]}${LINK[r]}`, v);
      else set(addr, v);
      if (!isTbd(v)) {
        ws.getCell(addr).numFmt = '#,##0;[Red]-#,##0';
        ws.getCell(addr).alignment = { horizontal: 'right', vertical: 'middle' };
      }
    }
    const ys = Object.values(tv[r]);
    if (r === 100) set(`Y${r}`, null);
    else if (ys.some(isTbd)) {
      tv[r].Y = TBD;
      markTbd(`Y${r}`);
    } else {
      const nums = ys.filter(isN);
      tv[r].Y = nums.length ? xround(nums.reduce((s, v) => s + v, 0) / nums.length, 0) : NA;
      F(`Y${r}`, `ROUND(AVERAGE(J${r},O${r},T${r}),0)`, tv[r].Y);
      ws.getCell(`Y${r}`).numFmt = '#,##0;[Red]-#,##0';
      ws.getCell(`Y${r}`).alignment = { horizontal: 'right', vertical: 'middle' };
    }
  }
  const RATE = { 108: [95, 89], 109: [90, 89], 110: [91, 89], 111: [104, 99], 112: [104, 95] };
  for (const [r, [nr, dr]] of Object.entries(RATE)) {
    for (const col of ['J', 'O', 'T', 'Y']) {
      const yk = { J: '당기', O: '전기', T: '전전기' }[col];
      const addr = `${col}${r}`;
      if (yk && !has(yk)) {
        set(addr, NA);
      } else {
        const a = tv[nr][col];
        const b = tv[dr][col];
        if (isTbd(a) || isTbd(b)) {
          markTbd(addr);
          continue;
        }
        F(addr, `IF(${col}${dr}<=0,"-",${col}${nr}/${col}${dr})`, !isN(b) || b <= 0 || !isN(a) ? NA : a / b);
      }
      ws.getCell(addr).numFmt = '0.00%';
      ws.getCell(addr).alignment = { horizontal: 'center', vertical: 'middle' };
    }
  }
  for (const { key, yk, m } of marksFor('법인세신고상황')) {
    const r = Object.entries(TAX_ROWS).find(([, k]) => k === key)?.[0] ?? { 당기순손익: 92, 매출액: 89 }[key];
    if (r && C_TX[yk]) applyMark(`${C_TX[yk]}${r}`, m);
  }
  const n3 = notes('Ⅲ');
  if (n3.length) {
    // 템플릿 118행(**)부터 주석 교체: 1줄은 C118, 2줄째부터 B119 '***'
    set('C118', n3[0], { font: noteFont(n3[0]) });
    n3.slice(1, 3).forEach((s, i) => {
      if (i === 0) set('B119', '***', { font: f9 });
      set(`C${119 + i}`, s, { font: noteFont(s) });
    });
  } else set('C118', '법인지방소득세는 법인지방소득세 과세표준 및 세액신고서상 신고세액입니다. (신고서가 없는 연도는 0)');
  if (lossDiff) {
    // Q-08: 숫자는 그대로 두고 "차이 있음"만 표시
    ws.getCell('J96').fill = solid('FFFCE4D6');
    ws.getCell('J96').note = `차이 있음: ${PRV} 손익계산서 결손금 ${fixed(lossIsPrev, 0)}원 vs 이월결손금 ${fixed(lossCf, 0)}원`;
    const red9 = font({ size: 9, color: 'FFC00000' });
    set('B119', '***', { font: red9 });
    set('C119', `(-)이월결손금 ${CUR} 사업연도 ${fixed(lossCf, 0)}원은 ${PRV} 사업연도 손익계산서상 결손금 ${fixed(lossIsPrev, 0)}원과 ${fixed(lossCf - lossIsPrev, 0)}원 차이 있음 (확인 중)`, { font: red9 });
  }
  set('B122', '*', { font: f9 });
  set('C122', '부가가치세 신고내용(1)과 원천징수이행상황 신고내용(2)은 보고서에서 제외하였습니다.', { font: f9 });

  // ── Ⅳ-(3)~Ⅴ-(1): 템플릿 221~290행 (목록 칸은 항목 수만큼) ──
  let row = 123;
  row = copyBlock(221, 228, row);
  const R = (t) => rowMap[t];
  set(`B${R(221)}`, `(3) ${CUR}사업연도 기납부세액`);
  const pp = d.기납부세액 || {};
  const ppv = ['중간예납세액', '원천납부세액', '외국납부세액', '수시부과세액'].map((k) => (pp[k] === undefined ? 0 : val(pp[k])));
  ['B', 'H', 'M', 'S'].forEach((c, i) => putVal(`${c}${R(224)}`, ppv[i], { fmt: '#,##0' }));
  FX(`Y${R(224)}`, `SUM(B${R(224)}:S${R(224)})`, ppv, () => ppv.reduce((s, v) => s + v, 0), { fmt: '#,##0' });

  // (4) 이월결손금
  const lossFirst = row;
  row = copySlots([229, 230], Math.max(2, lossRows.length), row);
  const lossLast = row - 1;
  if (!lossRows.length) set(`B${lossFirst}`, '해당 없음');
  lossRows.forEach((it, i) => {
    const r = lossFirst + i;
    const a = val(it.발생액);
    const b = val(it.감소액);
    set(`B${r}`, num(it.발생사업연도) ?? it.발생사업연도 ?? '');
    putVal(`I${r}`, a, { fmt: '#,##0' });
    putVal(`P${r}`, b, { fmt: '#,##0' });
    FX(`W${r}`, `I${r}-P${r}`, [a, b], () => a - b, { fmt: '#,##0' });
  });
  row = copyBlock(231, 237, row);
  const lv = (k) => lossRows.map((it) => val(it[k]));
  const lsum = (k) => lv(k).reduce((s, v) => s + (isN(v) ? v : 0), 0);
  FX(`I${R(231)}`, `SUM(I${lossFirst}:O${lossLast})`, lv('발생액'), () => lsum('발생액'), { fmt: '#,##0' });
  FX(`P${R(231)}`, `SUM(P${lossFirst}:V${lossLast})`, lv('감소액'), () => lsum('감소액'), { fmt: '#,##0' });
  FX(`W${R(231)}`, `SUM(W${lossFirst}:AC${lossLast})`, [...lv('발생액'), ...lv('감소액')], () => lsum('발생액') - lsum('감소액'), { fmt: '#,##0' });
  const n44 = notes('Ⅳ-(4)');
  set(`C${R(232)}`, n44[0] || `${CUR}사업연도 자본금과적립금조정명세서(갑) 이월결손금 내역을 토대로 작성되었습니다.`);
  if (n44[1]) {
    set(`B${R(233)}`, '**', { font: f9 });
    set(`C${R(233)}`, n44[1], { font: noteFont(n44[1]) });
  }
  if (lossDiff) {
    const red9 = font({ size: 9, color: 'FFC00000' });
    set(`B${R(233)}`, '**', { font: red9 });
    set(`C${R(233)}`, `발생액 ${fixed(lossCf, 0)}원은 ${PRV} 사업연도 손익계산서상 결손금 ${fixed(lossIsPrev, 0)}원과 ${fixed(lossCf - lossIsPrev, 0)}원 차이 있음 (확인 중)`, { font: red9 });
  }

  // (5) 소득금액조정합계표
  const adj = d.소득금액조정 || {};
  const plus = arr(adj.익금산입및손금불산입);
  const minus = arr(adj.손금산입및익금불산입);
  const adjFirst = row;
  row = copySlots([238, 239], Math.max(2, plus.length, minus.length), row);
  const adjLast = row - 1;
  plus.forEach((it, i) => {
    set(`B${adjFirst + i}`, it.과목 ?? '');
    putVal(`H${adjFirst + i}`, val(it.금액), { fmt: '#,##0' });
    set(`M${adjFirst + i}`, it.소득처분 ?? '');
    if (it.경고) applyMark(`H${adjFirst + i}`, { 종류: '경고', 메모: it.경고 });
  });
  minus.forEach((it, i) => {
    set(`P${adjFirst + i}`, it.과목 ?? '');
    putVal(`V${adjFirst + i}`, val(it.금액), { fmt: '#,##0' });
    set(`AA${adjFirst + i}`, it.소득처분 ?? '');
  });
  row = copyBlock(240, 244, row);
  const av = (a) => a.map((it) => val(it.금액));
  const asum = (a) => av(a).reduce((s, v) => s + (isN(v) ? v : 0), 0);
  FX(`H${R(240)}`, `SUM(H${adjFirst}:L${adjLast})`, av(plus), () => asum(plus), { fmt: '#,##0' });
  FX(`V${R(240)}`, `SUM(V${adjFirst}:Z${adjLast})`, av(minus), () => asum(minus), { fmt: '#,##0' });
  const n45 = notes('Ⅳ-(5)');
  if (n45.length) {
    set(`B${R(241)}`, '*', { font: f9 });
    set(`C${R(241)}`, n45[0], { font: noteFont(n45[0]) });
  }

  // (6) 공제·감면세액
  const credits = arr(d.공제감면세액);
  const crFirst = row;
  row = copySlots([245, 246], Math.max(2, credits.length), row);
  const crLast = row - 1;
  credits.forEach((it, i) => {
    set(`B${crFirst + i}`, it.구분명 ?? '');
    set(`M${crFirst + i}`, it.근거법조항 ?? '');
    putVal(`X${crFirst + i}`, val(it.공제감면세액), { fmt: '#,##0' });
  });
  row = copyBlock(247, 270, row);
  const cv = credits.map((it) => val(it.공제감면세액));
  FX(`X${R(247)}`, `SUM(X${crFirst}:AC${crLast})`, cv, () => cv.reduce((s, v) => s + (isN(v) ? v : 0), 0), { fmt: '#,##0' });
  const n46 = notes('Ⅳ-(6)');
  if (n46.length) {
    set(`B${R(248)}`, '*', { font: noteFont(n46[0]) });
    set(`C${R(248)}`, n46[0], { font: noteFont(n46[0]) });
  }

  // (7) 재고율
  set(`L${R(251)}`, `${CUR} 사업연도`);
  set(`U${R(251)}`, `${PRV} 사업연도`);
  for (const [col, yk] of [['L', '당기'], ['U', '전기']]) {
    for (const [r, v] of [[252, has(yk) ? BS(44, yk) : NA], [253, has(yk) ? IS(35, yk) : NA]]) {
      putVal(`${col}${R(r)}`, v, { fmt: '#,##0', al: 'center' });
    }
    const inv = has(yk) ? BS(44, yk) : NA;
    const cogs = has(yk) ? IS(35, yk) : NA;
    if (!has(yk)) set(`${col}${R(254)}`, NA, { al: 'center' });
    else FX(`${col}${R(254)}`, `IF(${col}${R(253)}=0,"-",ROUND(${col}${R(252)}/${col}${R(253)}*100,2))`, [inv, cogs], () => (!cogs ? NA : xround((inv / cogs) * 100, 2)), {
      fmt: '0.00',
      al: 'center',
    });
  }
  set(`B${R(255)}`, '*');
  set(`C${R(255)}`, '재고율 = 재고자산/매출원가 * 100  (매출원가가 0인 연도는 『-』)');
  // (8) 4대보험: (지원예정), 금액 칸 비움 (T-09)
  set(`Y${R(261)}`, null);

  // Ⅴ-0 안내사항
  const guide = d.안내사항 || {};
  const lines = (k) => arr(guide[k]).map(String).filter((s) => s.trim());
  const tax = lines('세무조정사항');
  const taxFirst = row;
  row = copySlots([271, 272], Math.max(2, tax.length), row);
  (tax.length ? tax : ['이슈없음']).forEach((s, i) => set(`B${taxFirst + i}`, s));
  row = copyBlock(273, 273, row);
  const acc = lines('회계사항'); // T-12: 담당자 직접 입력
  const accFirst = row;
  row = copySlots([274, 275, 276], Math.max(3, acc.length), row);
  acc.forEach((s, i) => set(`B${accFirst + i}`, s));
  row = copyBlock(277, 278, row);
  const red = lines('감면사항');
  const redFirst = row;
  row = copySlots([279, 280], Math.max(2, red.length), row);
  red.forEach((s, i) => {
    const c = set(`B${redFirst + i}`, s);
    if (s.startsWith(TBD)) c.font = { ...(c.font || {}), name: FONT, size: 11, color: { argb: RED } };
  });
  row = copyBlock(281, 290, row);
  const filing = d.신고내용 || {};
  if (filing.신고서명) set(`B${R(285)}`, filing.신고서명);
  set(`B${R(286)}`, `과세기간: ${filing.과세기간 || ''}`);
  set(`B${R(287)}`, `신고유형(신고일자): ${filing.신고유형 || ''}${filing.신고일자 ? `(${filing.신고일자})` : ''}`);

  // ── Ⅴ-(2) 신고서식 제출 목록 ──────────────────────
  const stMark = clone(src.getCell('A291').style);
  const stTitle = clone(src.getCell('B291').style);
  const stDesc = clone(src.getCell('B292').style);
  for (const f of arr(d.제출서식목록)) {
    const a = ws.getCell(row, 1);
    a.value = '○';
    a.style = clone(stMark);
    const t = ws.getCell(row, 2);
    t.value = `[${f.서식명 ?? ''}]`;
    t.style = clone(stTitle);
    const ls = wrapText(f.수행업무 ?? '', 46);
    ls.forEach((ln, k) => {
      const c = ws.getCell(row + 1 + k, 2);
      c.value = ln;
      c.style = clone(stDesc);
    });
    row += 2 + ls.length;
  }
  row += 1;

  // ── 새 섹션 도우미 ─────────────────────────────────
  const SEC_FONT = clone(src.getCell('B6').font);
  const SUB_FONT = clone(src.getCell('B22').font);
  function box(r, a, b, value, o = {}) {
    const { head = false, fmt, h = 'center', bold = false, size = 10, color = 'FF000000', fill, wrap = false, border = true, r2 = r } = o;
    const c1 = colNum(a);
    const c2 = colNum(b);
    if (c2 > c1 || r2 > r) ws.mergeCells(r, c1, r2, c2);
    for (let rr = r; rr <= r2; rr += 1) {
      for (let cc = c1; cc <= c2; cc += 1) {
        const cell = ws.getCell(rr, cc);
        if (border) cell.border = clone(BOX);
        if (head) cell.fill = clone(HEAD_FILL);
        else if (fill) cell.fill = solid(fill);
      }
    }
    const c = ws.getCell(r, c1);
    setCellValue(c, value);
    c.font = { name: FONT, size, bold: bold || head, color: { argb: color } };
    c.alignment = { horizontal: h, vertical: 'middle', wrapText: wrap };
    if (fmt) c.numFmt = fmt;
    return c;
  }
  const text = (r, col, value, size = 9, bold = false, color = 'FF000000') => {
    const c = ws.getCell(`${col}${r}`);
    setCellValue(c, value);
    c.font = { name: FONT, size, bold, color: { argb: color } };
    return c;
  };
  const breaks = [];

  // ── Ⅵ. 주요 재무지표 요약 (T-02, T-10) ─────────────
  const s0 = IS(1, '당기');
  const s1 = hp ? IS(1, '전기') : 0;
  const op0 = IS(129, '당기');
  const op1 = hp ? IS(129, '전기') : 0;
  const ni0 = IS(219, '당기');
  const ni1 = hp ? IS(219, '전기') : 0;
  const eq0 = BS(382, '당기');
  const eq1 = hp ? BS(382, '전기') : null;
  const debt0 = !allN(eq0, BS(333, '당기')) ? TBD : eq0 > 0 ? (BS(333, '당기') / eq0) * 100 : null;
  const debt1 = !hp ? null : !allN(eq1, BS(333, '전기')) ? TBD : eq1 > 0 ? (BS(333, '전기') / eq1) * 100 : null;
  const coreOk = allN(s0, op0, ni0, IS(67, '당기'), IS(130, '당기'), IS(179, '당기'), BS(228, '당기')) &&
    (!hp || allN(s1, op1, ni1, IS(67, '전기'), BS(228, '전기')));

  breaks.push(row - 1);
  row += 1;
  text(row, 'B', 'Ⅵ. 주요 재무지표 요약').font = clone(SEC_FONT);
  row += 1;
  text(row, 'B', '*');
  const basisOf = d.보고서정보?.기준 ? '의 표준재무제표' : '';
  text(row, 'C', hp ? `전기(${PRV})·당기(${CUR}) 세무조정계산서${basisOf}를 토대로 산출한 핵심지표, 요약 코멘트, 그래프입니다.` : `당기(${CUR}) 세무조정계산서${basisOf}를 토대로 산출한 핵심지표, 요약 코멘트, 그래프입니다.`);
  row += 2;
  text(row, 'B', '(1) 핵심지표 비교').font = clone(SUB_FONT);
  row += 2;

  const turn = (c, p) => (hp && p <= 0 && c > 0 ? ' (흑자전환)' : '');
  const star = new Set(arr(d.핵심지표별표).map(String));
  const U = (v) => (isN(v) ? unit(v) : TBD);
  const card = (title, c0, c1) => [
    title,
    U(c0) + (star.has(title) ? ' *' : ''),
    hp ? `전기 ${U(c1)}` : '전기 자료 없음',
    !hp ? '-' : allN(c0, c1) ? `증감 ${unit(c0 - c1, true)}${turn(c0, c1)}` : `증감 ${TBD}`,
    allN(c0, c1) ? c0 - c1 : 0,
  ];
  const cards = [card('매출액', s0, s1), card('영업이익', op0, op1), card('당기순이익', ni0, ni1)];
  {
    const dv = (x) => (x === TBD ? TBD : x === null ? '자본잠식' : `${fixed(x, 2)}%`);
    const v = dv(debt0) + (star.has('부채비율') ? ' *' : '');
    let p = '전기 자료 없음';
    let dl = '-';
    let sign = 0;
    if (hp) {
      p = debt1 === null ? '전기 자본잠식' : `전기 ${dv(debt1)}`;
      if (debt0 === TBD || debt1 === TBD) dl = `증감 ${TBD}`;
      else if (debt1 === null && debt0 !== null) { dl = '자본잠식 해소'; sign = 1; }
      else if (debt0 === null) { dl = '자본잠식'; sign = -1; }
      else { dl = `증감 ${debt0 - debt1 > 0 ? '+' : ''}${fixed(debt0 - debt1, 2)}%p`; sign = debt1 - debt0; }
    }
    cards.push(['부채비율', v, p, dl, sign]);
  }
  const CARD_COLS = [['B', 'H'], ['I', 'O'], ['P', 'V'], ['W', 'AC']];
  CARD_COLS.forEach(([a, b], i) => {
    const [title, val, prev, delta, dd] = cards[i];
    box(row, a, b, title, { fill: NAVY, color: 'FFFFFFFF', bold: true });
    box(row + 1, a, b, val, { size: 16, bold: true, color: NAVY, r2: row + 2 });
    box(row + 3, a, b, prev, { size: 9, color: 'FF595959' });
    box(row + 4, a, b, delta, { size: 9, bold: true, color: dd > 0 ? 'FFC00000' : 'FF0070C0' });
  });
  ws.getRow(row).height = 20;
  row += 5;
  text(row, 'B', '*');
  text(row, 'C', '1억원 이상은 억원(소수 2자리), 1억원 미만은 천만원(소수 1자리) 단위로 반올림하여 표시합니다. 부채비율 = 부채총계/자본총계*100');
  row += 1;
  for (const s of notes('Ⅵ')) {
    text(row, 'C', s, 9, false, isAlertLine(s) ? RED : 'FF000000');
    row += 1;
  }
  row += 1;

  // (2) 요약 코멘트
  text(row, 'B', '(2) 요약 코멘트').font = clone(SUB_FONT);
  row += 2;
  const sgaC = IS(67, '당기');
  const sgaP = hp ? IS(67, '전기') : 0;
  const top5 = [...SGA].sort((a, b) => b.cur - a.cur).slice(0, 5);
  const rest = [...SGA].sort((a, b) => b.cur - a.cur).slice(5);
  const comments = arr(d.요약코멘트).filter((s) => String(s).trim()).length
    ? arr(d.요약코멘트).filter((s) => String(s).trim()).map((s, i) => {
        const m = String(s).match(/^\s*(?:\d+\.\s*)?\[?([^\]:：]{2,20})[\]:：]\s*(.+)$/);
        return m ? [m[1].trim(), m[2].trim()] : [`코멘트 ${i + 1}`, String(s)];
      })
    : coreOk
      ? autoComments()
      : [['자동 코멘트 보류', `기초 재무자료에 『${TBD}』 값이 있어 자동 코멘트를 보류했습니다. 확인 후 다시 출력해 주세요.`]];
  comments.forEach(([head, body], i) => {
    const n = linesNeeded(body, 95);
    box(row, 'B', 'F', `${i + 1}. ${head}`, { head: true, size: 9, wrap: true, r2: row + n - 1 });
    box(row, 'G', 'AC', body, { h: 'left', size: 9, wrap: true, r2: row + n - 1 });
    row += n;
  });
  row += 1;

  function autoComments() {
    const out = [];
    // 1. 매출 및 수익성
    {
      let s = `당기 매출액은 ${won(s0)}(${unit(s0)})으로`;
      if (hp) s += ` 전기 ${won(s1)} 대비 ${won(Math.abs(s0 - s1))}(${chgRate(s0, s1)}) ${s0 >= s1 ? '증가' : '감소'}하였습니다.`;
      else s += ' 집계되었습니다.';
      const opm = s0 ? fixed(pct(op0, s0), 2) : '-';
      if (op0 >= 0) {
        s += ` 영업이익은 ${won(op0)}(영업이익률 ${opm}%)으로`;
        if (hp && op1 < 0) s += ` 전기 영업손실 ${won(-op1)}에서 흑자전환하였고,`;
        else if (hp) s += ` 전기 ${won(op1)} 대비 ${op0 >= op1 ? '증가' : '감소'}하였고,`;
      } else {
        s += ` 영업손실은 ${won(-op0)}(영업이익률 ${opm}%)이며,`;
      }
      s += ni0 >= 0 ? ` 당기순이익은 ${won(ni0)}입니다.` : ` 당기순손실은 ${won(-ni0)}입니다.`;
      out.push(['매출 및 수익성', s]);
    }
    // 2. 판관비 주요항목
    if (SGA.length) {
      let s = `판관비는 ${won(sgaC)}으로 매출액 대비 ${fixed(pct(sgaC, s0) ?? 0, 2)}%${hp && s1 ? `(전기 ${fixed(pct(sgaP, s1), 2)}%)` : ''}입니다. `;
      s += `주요 항목은 ${top5.slice(0, 3).map((r) => `${r.name} ${fixed(pct(r.cur, sgaC) ?? 0, 1)}%`).join(', ')} 순(판관비 대비)`;
      if (hp) {
        const inc = [...SGA].sort((a, b) => b.cur - b.prev - (a.cur - a.prev));
        const dec = [...SGA].sort((a, b) => a.cur - a.prev - (b.cur - b.prev));
        const ups = inc.filter((r) => r.cur - r.prev > 0).slice(0, 2);
        const down = dec.filter((r) => r.cur - r.prev < 0).slice(0, 1);
        const parts = [];
        if (ups.length) parts.push(`${ups.map((r) => `${r.name}(+${won(r.cur - r.prev)})`).join(', ')}가 크게 늘고`);
        if (down.length) parts.push(`${down.map((r) => `${r.name}(${won(r.cur - r.prev)})`).join(', ')}가 줄었습니다.`);
        s += parts.length ? `이며, ${parts.join(' ')}` : '입니다.';
        if (ups.length && !down.length) s = s.replace(/가 크게 늘고$/, '가 크게 늘었습니다.');
      } else s += '입니다.';
      out.push(['판관비 주요항목', s]);
    }
    // 3. 영업외손익
    {
      const det = arr(d.영업외손익명세);
      const inc = det.filter((r) => r.구분 === '영업외수익' && (num(r.당기) ?? 0) !== 0);
      const exp = det.filter((r) => r.구분 === '영업외비용' && (num(r.당기) ?? 0) !== 0);
      const joinKo = (a) => (a.length <= 1 ? a.join('') : `${a.slice(0, -1).join(', ')}과 ${a[a.length - 1]}`);
      const incS = joinKo(inc.map((r) => `${r.계정과목} ${won(num(r.당기))}`));
      let s = `영업외수익 ${won(IS(130, '당기'))}은 ${incS || '발생하지 않았으며'}${incS ? '이며' : ''}, 영업외비용 ${won(IS(179, '당기'))}은 `;
      s += exp.length === 1 ? `${exp[0].계정과목}입니다.` : exp.length ? `${joinKo(exp.map((r) => `${r.계정과목} ${won(num(r.당기))}`))}입니다.` : '발생하지 않았습니다.';
      const names = inc.map((r) => String(r.계정과목));
      const absent = ['단기투자자산처분이익', '외환차익'].filter((k) => !names.includes(k));
      if (absent.length) s += ` ${absent.join('·')} 등 그 밖의 영업외수익은 발생하지 않았습니다.`;
      out.push(['영업외손익 발생 사유', s]);
    }
    // 4. 전년 대비 주요 변동사항
    if (hp) {
      const ap = d.결산부속명세 || {};
      const cash = num(ap.현금성자산?.당기말잔액);
      const ar = num(ap.외상매출금?.당기말잔액);
      const g = ap.대표이사가수금 || {};
      const a0 = BS(228, '당기');
      const a1 = BS(228, '전기');
      const det = [cash !== null ? `현금성자산 ${won(cash)}` : null, ar !== null ? `매출채권 ${won(ar)}` : null].filter(Boolean);
      let s = `자산총계가 ${won(a1)}에서 ${won(a0)}으로 ${a0 >= a1 ? '증가' : '감소'}하였고${det.length ? `(${det.join(', ')})` : ''}, `;
      if (eq1 <= 0 && eq0 > 0) s += `자본총계는 ${won(eq1)}에서 ${won(eq0)}으로 전환되어 자본잠식이 해소되었습니다.`;
      else s += `자본총계는 ${won(eq1)}에서 ${won(eq0)}으로 변동하였습니다.`;
      if (num(g.전기말잔액) !== null && num(g.당기말잔액) !== null) {
        s += ` 대표이사 가수금은 ${won(num(g.전기말잔액))}에서 ${won(num(g.당기말잔액))}으로 ${num(g.당기말잔액) < num(g.전기말잔액) ? '정리되었습니다' : '증가하였습니다'}.`;
      }
      out.push(['전년 대비 주요 변동사항', s]);
    }
    // 5. 재무비율
    {
      const cr = (yk) => (BS(229, yk) ? fixed((BS(1, yk) / BS(229, yk)) * 100, 2) : '-');
      let s = `부채비율은 ${debt0 === null ? '자본잠식' : `${fixed(debt0, 2)}%`}`;
      if (hp) s += `(전기 ${debt1 === null ? '자본잠식' : `${fixed(debt1, 2)}%`})`;
      if (debt0 !== null) s += `로 표준비율(200% 이하)을 ${debt0 <= 200 ? '충족하며' : '초과하며'}`;
      else s += '이며';
      s += `, 유동비율은 ${cr('당기')}%${hp ? `(전기 ${cr('전기')}%)` : ''}입니다.`;
      out.push(['재무비율', s]);
    }
    // 6. 공제·감면
    {
      const cr = credits.filter((x) => num(x.공제감면세액));
      const parts = cr.map((x) => {
        const name = String(x.구분명 || '').replace(/\s*\(최저한세 적용(대상|제외)\)/, '');
        const law = String(x.근거법조항 || '').replace(/「조세특례제한법」\s*/, '조특법 ').trim();
        return `${name}${law ? `(${law})` : ''} ${won(num(x.공제감면세액))}`;
      });
      let s = parts.length ? `${parts.join(', ')}이 반영되었고` : '반영된 공제·감면세액은 없으며';
      const used = txk('이월결손금', '당기');
      if (used > 0) {
        const remain = lossRows.reduce((t, it) => t + ((num(it.발생액) ?? 0) - (num(it.감소액) ?? 0)), 0);
        s += `, ${PRV} 사업연도 이월결손금 ${won(used)}이 ${remain === 0 ? '전액 ' : ''}공제되어`;
      } else s += ',';
      s += ` 총부담세액은 ${won(txk('총부담세액', '당기'))}(법인지방소득세 ${won(txk('법인지방소득세', '당기'))} 별도)입니다.`;
      out.push(['공제·감면 사항', s]);
    }
    return out;
  }

  // (3) 그래프
  text(row, 'B', '(3) 그래프').font = clone(SUB_FONT);
  row += 1;
  text(row, 'B', '*');
  text(row, 'C', `판관비 구성비율·주요 5개 항목은 ${d.보고서정보?.기준 ? '표준손익계산서 ' : ''}당기 금액 상위 5개 기준이며 나머지는 『기타』로 합산하였습니다.`);
  row += 1;
  const chartRow = row + 1;
  const chartRows = 33;
  row = chartRow + 16 * 2 + 1;

  // ── Ⅶ. 손익 분석 ──────────────────────────────────
  breaks.push(row - 1);
  row += 1;
  text(row, 'B', 'Ⅶ. 손익 분석').font = clone(SEC_FONT);
  row += 2;
  text(row, 'B', '(1) 요약손익 비교').font = clone(SUB_FONT);
  text(row, 'AC', '(단위: 원)', 8, false, 'FF002060').alignment = { horizontal: 'right' };
  row += 2;
  for (const [a, b, t] of [['B', 'I', '구 분'], ['J', 'N', `${PRV} 사업연도(전기)`], ['O', 'S', `${CUR} 사업연도(당기)`], ['T', 'X', '증감액'], ['Y', 'AC', '증감률(%)']]) {
    box(row, a, b, t, { head: true });
  }
  row += 1;
  const PL_ROW = {};
  const PL = [
    ['매출액', 'R50', 'L50', (yk) => IS(1, yk)],
    ['매출총이익', 'R52', 'L52', (yk) => IS(66, yk)],
    ['영업손익', 'R54', 'L54', (yk) => IS(129, yk)],
    ['영업외손익', 'R55-R56', 'L55-L56', (yk) => IS(130, yk) - IS(179, yk)],
    ['법인세차감전손익', 'R57', 'L57', (yk) => IS(217, yk)],
    ['법인세비용', 'R58', 'L58', (yk) => IS(218, yk)],
    ['당기순이익', 'R59', 'L59', (yk) => IS(219, yk)],
  ];
  const safe = (fn) => (yk) => {
    try {
      const v = fn(yk);
      return isN(v) ? v : TBD;
    } catch {
      return TBD;
    }
  };
  const plMarks = marksFor('요약손익');
  for (const [name, p, c, f0] of PL) {
    const f = safe(f0);
    PL_ROW[name] = row;
    const pv = hp ? f('전기') : NA;
    const cv = f('당기');
    const cell = (a, b, formula, v, fmt) => (isTbd(v) ? (box(row, a, b, null), markTbd(`${a}${row}`)) : box(row, a, b, isN(v) ? { formula, result: v } : v, { fmt, h: 'right' }));
    box(row, 'B', 'I', name, { h: 'left' });
    cell('J', 'N', p, pv, '#,##0;[Red]-#,##0');
    cell('O', 'S', c, cv, '#,##0;[Red]-#,##0');
    if (allN(cv, pv)) {
      box(row, 'T', 'X', { formula: `O${row}-J${row}`, result: cv - pv }, { fmt: '#,##0;[Red]-#,##0', h: 'right' });
      const rate = pv <= 0 ? (cv > 0 ? '흑자전환' : NA) : xround(((cv - pv) / Math.abs(pv)) * 100, 2);
      box(row, 'Y', 'AC', { formula: `IF(J${row}<=0,IF(O${row}>0,"흑자전환","-"),ROUND((O${row}-J${row})/ABS(J${row})*100,2))`, result: rate }, { fmt: '#,##0.00;[Red]-#,##0.00', h: 'right' });
    } else if (!hp) {
      box(row, 'T', 'X', NA, { h: 'right' });
      box(row, 'Y', 'AC', NA, { h: 'right' });
    } else {
      box(row, 'T', 'X', null);
      markTbd(`T${row}`);
      box(row, 'Y', 'AC', null);
      markTbd(`Y${row}`);
    }
    for (const { key, yk, m } of plMarks) if (key === name) applyMark(`${yk === '전기' ? 'J' : 'O'}${row}`, m);
    row += 1;
  }
  text(row, 'B', '*');
  text(row, 'C', '증감률 = (당기-전기)/|전기|*100.  전기 0 이하 → 당기 흑자이면 『흑자전환』, 그 외 『-』.  영업외손익 = 영업외수익-영업외비용');
  row += 3;

  text(row, 'B', '(2) 판관비 주요항목').font = clone(SUB_FONT);
  text(row, 'AC', '(단위: 원)', 8, false, 'FF002060').alignment = { horizontal: 'right' };
  row += 2;
  // 채점·검색에서 연도를 알 수 있게 머리글에 사업연도를 적음
  for (const [a, b, t] of [['B', 'I', '항 목'], ['J', 'M', `${PRV} 사업연도(전기)`], ['N', 'R', `${CUR} 사업연도(당기)`], ['S', 'V', '증감액'], ['W', 'Y', `${PRV} 매출대비(%)`], ['Z', 'AC', `${CUR} 매출대비(%)`]]) {
    box(row, a, b, t, { head: true, size: 9 });
  }
  row += 1;
  const SGA_FIRST = row;
  const sgaVals = [];
  for (const r of top5) {
    box(row, 'B', 'I', r.name, { h: 'left' });
    box(row, 'J', 'M', r.prev, { fmt: '#,##0', h: 'right' });
    box(row, 'N', 'R', r.cur, { fmt: '#,##0', h: 'right' });
    sgaVals.push([r.prev, r.cur]);
    row += 1;
  }
  const etcP = isN(sgaP) ? sgaP - top5.reduce((s, r) => s + r.prev, 0) : 0;
  const etcC = isN(sgaC) ? sgaC - top5.reduce((s, r) => s + r.cur, 0) : 0;
  box(row, 'B', 'I', '기타', { h: 'left' });
  box(row, 'J', 'M', { formula: `R53-SUM(J${SGA_FIRST}:J${row - 1})`, result: hp ? etcP : NA }, { fmt: '#,##0', h: 'right' });
  box(row, 'N', 'R', { formula: `L53-SUM(N${SGA_FIRST}:N${row - 1})`, result: etcC }, { fmt: '#,##0', h: 'right' });
  sgaVals.push([etcP, etcC]);
  const SGA_ETC = row;
  row += 1;
  box(row, 'B', 'I', '합 계', { head: true });
  box(row, 'J', 'M', { formula: `SUM(J${SGA_FIRST}:J${row - 1})`, result: sgaP }, { fmt: '#,##0', h: 'right', bold: true });
  box(row, 'N', 'R', { formula: `SUM(N${SGA_FIRST}:N${row - 1})`, result: sgaC }, { fmt: '#,##0', h: 'right', bold: true });
  sgaVals.push([sgaP, sgaC]);
  const SGA_TOT = row;
  for (let r = SGA_FIRST; r <= SGA_TOT; r += 1) {
    const [pv, cv] = sgaVals[r - SGA_FIRST];
    const bold = r === SGA_TOT;
    if (allN(cv, pv)) box(r, 'S', 'V', { formula: `N${r}-J${r}`, result: cv - pv }, { fmt: '#,##0;[Red]-#,##0', h: 'right', bold });
    else { box(r, 'S', 'V', null, { bold }); markTbd(`S${r}`); }
    const shareCell = (a, b, formula, v, base) => {
      if (!allN(v, base)) { box(r, a, b, null, { bold }); markTbd(`${a}${r}`); return; }
      box(r, a, b, { formula, result: base ? xround((v / base) * 100, 2) : NA }, { fmt: '0.00', h: 'right', bold });
    };
    shareCell('W', 'Y', `IF(R$50=0,"-",ROUND(J${r}/R$50*100,2))`, pv, s1);
    shareCell('Z', 'AC', `IF(L$50=0,"-",ROUND(N${r}/L$50*100,2))`, cv, s0);
  }
  row += 1;
  text(row, 'B', '*');
  text(row, 'C', `${d.보고서정보?.기준 ? '표준손익계산서 ' : ''}당기 금액 상위 5개 항목 기준이며 나머지는 『기타』로 합산하였습니다. 매출대비(%) = 항목 금액/해당 연도 매출액*100`);
  row += 1;
  for (const ln of wrapText(`기타: ${rest.map((r) => r.name).join(', ')}`, 100)) {
    text(row, 'C', ln);
    row += 1;
  }
  row += 2;

  // ── Ⅷ. 재무제표 주요 특이사항 (Q-05, Q-08, Q-09) ───
  const diffs = arr(d.자료간차이);
  if (diffs.length) breaks.push(row - 1);
  text(row, 'B', 'Ⅷ. 재무제표 주요 특이사항').font = clone(SEC_FONT);
  row += 1;
  const n8 = notes('Ⅷ');
  text(row, 'B', '*');
  text(row, 'C', n8[0] || '결산부속명세서·재무상태표·손익계산서를 토대로 작성하였습니다.');
  row += 2;
  if (diffs.length) {
    text(row, 'B', '(1) 점검 항목').font = clone(SUB_FONT);
    row += 2;
  }
  box(row, 'B', 'G', '점검 항목', { head: true });
  box(row, 'H', 'Y', '내 용', { head: true });
  box(row, 'Z', 'AC', '출 처', { head: true });
  row += 1;
  for (const [head, body, srcName, tbd] of notables()) {
    const n = Math.max(linesNeeded(body, diffs.length ? 66 : 62), linesNeeded(srcName, 14));
    box(row, 'B', 'G', head, { h: 'left', size: 9, wrap: true, r2: row + n - 1 });
    box(row, 'H', 'Y', body, { h: 'left', size: 9, wrap: true, r2: row + n - 1, fill: tbd ? 'FFFFF2CC' : undefined, color: tbd ? RED : 'FF000000' });
    box(row, 'Z', 'AC', srcName, { size: 8, wrap: true, r2: row + n - 1 });
    row += n;
  }
  row += diffs.length ? 2 : 3;

  // (2) 자료 간 차이 (경고) — 표준재무제표 값을 쓰고 다른 서식 값은 덮어쓰지 않고 기록
  if (diffs.length) {
    text(row, 'B', '(2) 자료 간 차이 (경고)').font = clone(SUB_FONT);
    row += 1;
    text(row, 'B', '*');
    text(row, 'C', n8[1] || '표준재무제표 값을 사용하고, 다른 서식 값은 덮어쓰지 않고 아래에 기록합니다. 원인은 확인 중입니다.', 9, false, RED);
    row += 2;
    for (const [a, b, t] of [['B', 'C', 'ID'], ['D', 'G', '항목'], ['H', 'P', '사용 값'], ['Q', 'Y', '다른 서식 값'], ['Z', 'Z', '차이'], ['AA', 'AC', '위치']]) {
      box(row, a, b, t, { head: true, size: 9 });
    }
    row += 1;
    for (const w of diffs) {
      const f = (k) => String(w[k] ?? '');
      const n = Math.max(linesNeeded(f('항목'), 10), linesNeeded(f('사용값'), 26), linesNeeded(f('다른서식값'), 26), linesNeeded(f('위치'), 14));
      box(row, 'B', 'C', f('ID'), { size: 9, r2: row + n - 1, fill: 'FFFCE4D6' });
      box(row, 'D', 'G', f('항목'), { h: 'left', size: 9, wrap: true, r2: row + n - 1 });
      box(row, 'H', 'P', f('사용값'), { h: 'left', size: 9, wrap: true, r2: row + n - 1 });
      box(row, 'Q', 'Y', f('다른서식값'), { h: 'left', size: 9, wrap: true, r2: row + n - 1 });
      box(row, 'Z', 'Z', f('차이'), { size: 8, wrap: true, r2: row + n - 1 });
      box(row, 'AA', 'AC', f('위치'), { size: 8, wrap: true, r2: row + n - 1 });
      row += n;
    }
    row += 1;
    const n82 = notes('Ⅷ-(2)');
    n82.forEach((s, i) => {
      if (i === 0) text(row, 'B', '**');
      text(row, 'C', s);
      row += 1;
    });
    row += 3;
  }

  function notables() {
    const out = [];
    const extra = arr(d.특이사항추가)
      .map((it) => (typeof it === 'string'
        ? (() => { const m = it.match(/^\s*\[([^\]]+)\]\s*(.*)$/); return m ? [m[1], m[2], '', m[2].includes(TBD) || m[1].includes(TBD)] : ['특이사항', it, '', it.includes(TBD)]; })()
        : [String(it?.점검항목 ?? ''), String(it?.내용 ?? ''), String(it?.출처 ?? ''), Boolean(it?.확인필요)]))
      .filter(([h, b]) => h || b);
    if (d.특이사항자동작성 === false) return extra;
    if (!coreOk) return [['자동 점검 보류', `기초 재무자료에 『${TBD}』 값이 있어 자동 점검 항목을 보류했습니다.`, '-', true], ...extra];
    const ap = d.결산부속명세 || {};
    const AR = ap.외상매출금 || {};
    const G = ap.대표이사가수금 || {};
    const arC = num(AR.당기말잔액);
    const arP = num(AR.전기말잔액);
    if (num(AR.당기거래처수) !== null) {
      let s = `외상매출금 거래처가 ${hp ? `전기말 ${num(AR.전기말거래처수) ?? 0}곳에서 ` : ''}당기 거래 ${num(AR.당기거래처수)}곳(기말 잔액 보유 ${num(AR.기말잔액거래처수) ?? 0}곳)으로 ${hp ? '늘었습니다.' : '집계됩니다.'}`;
      const share = pct(num(AR.최대거래처발생액) ?? 0, num(AR.당기발생액) ?? 0);
      if (share !== null && AR.최대거래처명) {
        s += ` 최대 거래처(${AR.최대거래처명}) 발생액이 당기 외상매출 발생액의 ${fixed(share, 2)}%로 집중도는 ${share >= 50 ? '여전히 높습니다' : '높지 않습니다'}.`;
      }
      out.push(['매출처 다변화', s, '외상매출금명세서']);
    }
    if (arC !== null) {
      let s = `기말 매출채권 ${won(arC)}으로 매출액 대비 ${fixed(pct(arC, s0) ?? 0, 2)}%`;
      if (hp && arP !== null) {
        s += `(전기 ${won(arP)}, ${fixed(pct(arP, s1) ?? 0, 2)}%)입니다. `;
        s += arC > arP ? `잔액이 ${won(arC - arP)} 늘어 회수 관리가 필요합니다.` : `잔액이 ${won(arP - arC)} 줄었습니다.`;
      } else s += '입니다.';
      out.push(['매출채권 회수', s, '외상매출금명세서, 재무상태표']);
    }
    const gC = num(G.당기말잔액);
    const gP = num(G.전기말잔액);
    if (gC !== null) {
      let s = `대표이사(${info.대표자명 || ''}) 가수금이 `;
      if (gP !== null) s += `${won(gP)} → ${won(gC)}으로 ${gC <= gP ? `${won(gP - gC)} 정리되었습니다` : `${won(gC - gP)} 늘었습니다`}`;
      else s += `${won(gC)}입니다`;
      if (num(G.당기증가) !== null) s += ` (당기 증가 ${won(num(G.당기증가))}, 감소 ${won(num(G.당기감소) ?? 0)})`;
      out.push(['대표이사 가수금', `${s}.`, '가수금명세서']);
      if (eq0 > 0 && gC > 0) {
        const debtL = BS(333, '당기');
        const repay = ((debtL - gC) / eq0) * 100;
        const conv = ((debtL - gC) / (eq0 + gC)) * 100;
        out.push(['가수금 → 부채비율 영향', `잔여 가수금 ${won(gC)}을 상환하면 부채비율이 ${fixed(debt0, 2)}% → ${fixed(repay, 2)}%, 출자전환하면 ${fixed(conv, 2)}%가 됩니다.`, '재무상태표']);
      }
    }
    {
      const cap1 = hp ? BS(334, '전기') : 0;
      const cap0 = BS(334, '당기');
      let s;
      if (hp && eq1 < 0 && eq0 >= 0) {
        s = `전기 자본총계 ${won(eq1)}(완전잠식, 해소 필요액 ${won(-eq1)} / 부분잠식 기준 ${won(cap1 - eq1)})이었으나 당기순이익으로 자본총계 ${won(eq0)}${eq0 >= cap0 ? '(자본금 초과)' : '(부분잠식)'}가 되어 해소되었습니다. ${eq0 >= cap0 ? '추가 가수금 정리(출자전환)는 필요하지 않습니다.' : `자본금 회복까지 ${won(cap0 - eq0)}이 더 필요합니다.`}`;
      } else if (eq0 < 0) {
        s = `당기 자본총계 ${won(eq0)}로 완전자본잠식 상태입니다. 해소 필요액 ${won(-eq0)} / 부분잠식 기준 ${won(cap0 - eq0)}의 가수금 정리(출자전환 등)가 필요합니다.`;
      } else if (eq0 < cap0) {
        s = `당기 자본총계 ${won(eq0)}로 부분자본잠식 상태입니다. 자본금 회복까지 ${won(cap0 - eq0)}이 필요합니다.`;
      } else {
        s = `당기 자본총계 ${won(eq0)}로 자본잠식이 없습니다.`;
      }
      out.push(['자본잠식 해소', s, '재무상태표']);
    }
    if (ap.가지급금) {
      const v = num(ap.가지급금.당기말잔액) ?? 0;
      out.push(['가지급금', v === 0 ? '재무상태표·결산부속명세서에 가지급금 계정이 없습니다 (잔액 0원, 매출액 대비 0.00%).' : `가지급금 잔액은 ${won(v)}으로 매출액 대비 ${fixed(pct(v, s0) ?? 0, 2)}%입니다.`, '결산부속명세서']);
    }
    if (hp) {
      const nw = SGA.filter((r) => r.prev === 0 && r.cur > 0);
      const gone = SGA.filter((r) => r.cur === 0 && r.prev > 0);
      out.push([
        '판관비 신규·소멸 항목',
        `신규: ${nw.length ? nw.map((r) => `${r.name} ${won(r.cur)}`).join(', ') : '없음'} / 소멸: ${gone.length ? gone.map((r) => `${r.name}(전기 ${won(r.prev)})`).join(', ') : '없음'}`,
        '손익계산서',
      ]);
    }
    if (lossDiff) {
      out.push([
        '이월결손금 차이',
        `${PRV} 손익계산서 결손금 ${won(lossIsPrev)}과 ${CUR} 이월결손금 발생액 ${won(lossCf)}이 ${won(lossCf - lossIsPrev)} 차이 있음 (세무법인 확인 중, 숫자는 원본 그대로 기재).`,
        '손익계산서, 자본금과적립금조정명세서(갑)',
      ]);
    }
    if (hp) {
      const i1 = BS(44, '전기');
      const i0 = BS(44, '당기');
      if (i1 > 0 && i0 === 0) out.push(['재고자산', `전기말 상품 ${won(i1)}이 당기 전액 매출원가로 처리되어 기말 재고는 0원입니다.`, '손익계산서']);
      else if (i0 || i1) out.push(['재고자산', `기말 재고자산은 ${won(i0)}(전기 ${won(i1)})입니다.`, '재무상태표']);
    }
    out.push(...extra);
    return out;
  }

  // ── 서명란 (Q-06, T-08) ────────────────────────────
  box(row, 'B', 'AC', `본 보고서는 ${issuer}에서 발행한 보고서입니다.`, { size: 10, border: false });
  row += 2;
  box(row, 'S', 'W', '발 행 일', { head: true });
  box(row, 'X', 'AC', dateKo(d.보고서정보?.발행일));
  row += 1;
  box(row, 'S', 'W', '발 행 처', { head: true });
  box(row, 'X', 'AC', issuer, { bold: true, size: 11 });
  ws.getRow(row).height = 22;
  const lastRow = row + 1;

  // ── 차트데이터 시트 (숨김, 본문 셀 참조) ───────────
  const cd = wb.addWorksheet('차트데이터', { state: 'hidden' });
  const ref = (addr, result) => ({ formula: `'법인세 신고결과'!${addr}`, result });
  cd.addRow(['주요손익', '전기', '당기']);
  for (const name of ['매출액', '매출총이익', '영업손익', '당기순이익']) {
    const r = PL_ROW[name];
    cd.addRow([name, ref(`J${r}`, cellResult(ws.getCell(`J${r}`))), ref(`O${r}`, cellResult(ws.getCell(`O${r}`)))]);
  }
  cd.addRow([]);
  cd.addRow(['판관비', '전기', '당기']);
  for (let r = SGA_FIRST; r <= SGA_ETC; r += 1) {
    const [pv, cv] = sgaVals[r - SGA_FIRST];
    cd.addRow([ref(`B${r}`, ws.getCell(`B${r}`).value), ref(`J${r}`, pv), ref(`N${r}`, cv)]);
  }
  cd.addRow([]);
  cd.addRow(['재무상태', '전기', '당기']);
  for (const [name, rr] of [['자산총계', 35], ['부채총계', 38], ['자본총계', 44]]) {
    cd.addRow([name, ref(`R${rr}`, ws.getCell(`R${rr}`).value), ref(`L${rr}`, ws.getCell(`L${rr}`).value)]);
  }

  // ── 그래프 이미지 (Chart.js로 그린 PNG) ─────────────
  const pxW = Math.round((CHART_CM.w / 2.54) * 96);
  const pxH = Math.round((CHART_CM.h / 2.54) * 96);
  [['pl', 'B', 0], ['sgaPie', 'P', 0], ['sgaBar', 'B', 16], ['bs', 'P', 16]].forEach(([key, col, dr]) => {
    if (!images[key]) return;
    const id = wb.addImage({ base64: images[key], extension: 'png' });
    ws.addImage(id, { tl: { col: colNum(col) - 1, row: chartRow - 1 + dr }, ext: { width: pxW, height: pxH } });
  });

  // ── 인쇄 설정 ───────────────────────────────────────
  for (const orig of [84, 121, 265]) breaks.push((orig < 123 ? orig : R(orig)) - 1);
  for (const b of [...new Set(breaks)].sort((a, b) => a - b)) ws.getRow(b).addPageBreak();
  ws.pageSetup = {
    ...ws.pageSetup,
    printArea: `A1:AD${lastRow}`,
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
  };
  ws.headerFooter = { oddFooter: `&C&8${issuer}  |  &P / &N` };

  // 안전장치: 계산 결과가 숫자가 아니게 된 칸(NaN·무한대)은 "확인 필요"로 바꿈 (파일 손상 방지)
  wb.eachSheet((sh) => sh.eachRow({ includeEmpty: false }, (rw) => rw.eachCell({ includeEmpty: false }, (cell) => {
    const v = cell.value;
    const x = cellResult(cell);
    if (typeof x === 'number' && !Number.isFinite(x)) {
      if (sh === ws) markTbd(cell.address);
      else cell.value = v && typeof v === 'object' ? { ...v, result: TBD } : TBD;
    }
  })));

  return {
    wb,
    layout: { sheetName: ws.name, chartRow, chartRows, breaks: [...new Set(breaks)].sort((a, b) => a - b), lastRow },
  };
}
