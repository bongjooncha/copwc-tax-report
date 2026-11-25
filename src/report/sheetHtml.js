// 워크북 시트 → A4 페이지 HTML
// Excel의 인쇄 결과(PDF)와 같은 배치를 재현함:
// - 열 너비·행 높이·병합·글꼴·음영·테두리·맞춤·숫자 서식을 셀 그대로 사용
// - 인쇄 설정(A4 세로, 좌우 0.4in·상하 0.5in, 너비 1페이지 맞춤, 가운데 정렬)과 수동 페이지 나눔을 따름
// - 내용이 한 페이지를 넘치면 자동으로 나누되, 병합 칸·그래프는 쪼개지 않음
import { cellResult } from './formulaResult.js';

const PX_PER_IN = 96;
const A4 = { w: (210 / 25.4) * PX_PER_IN, h: (297 / 25.4) * PX_PER_IN };
const MARGIN = { lr: 0.4 * PX_PER_IN, tb: 0.5 * PX_PER_IN, footer: 0.3 * PX_PER_IN };
const FONT_STACK = "'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo','Noto Sans KR',sans-serif";

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
// Excel 열 너비(문자 수) → px. 출력예시 기본 글꼴 기준 실측 계수 (Excel PDF와 대조해 맞춤)
const colPx = (w) => (w ?? 11) * 7.14 + 5;
// Excel은 행 높이를 정수 px로 그림 (17pt → 22px)
const rowPx = (pt) => Math.floor((pt * 4) / 3 + 0.01);
const argb = (c) => (c && typeof c.argb === 'string' && c.argb.length === 8 ? `#${c.argb.slice(2)}` : null);

function group(intStr) {
  return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
function fixed(v, d) {
  const s = Math.abs(v).toFixed(d);
  const [i, f] = s.split('.');
  return `${group(i)}${f ? `.${f}` : ''}`;
}

// 사용하는 숫자 서식만 처리 (#,##0 / #,##0.00 / 0.00 / 0.00% / [Red] 음수 구역 / General)
export function formatNumber(v, fmt) {
  const f = String(fmt || 'General');
  const secs = f.split(';');
  const neg = v < 0 && secs.length > 1;
  const sec = neg ? secs[1] : secs[0];
  const red = /\[Red\]/i.test(sec);
  const body = sec.replace(/\[[^\]]*\]/g, '').replace(/"/g, '');
  let out;
  if (/General/i.test(body) || !/[0#]/.test(body)) {
    out = Number.isInteger(v) ? String(v) : String(Number(v.toPrecision(10)));
  } else {
    const isPct = body.includes('%');
    const x = isPct ? v * 100 : v;
    const dec = (body.split('.')[1] || '').replace(/[^0#]/g, '').length;
    const useGroup = body.includes(',');
    const abs = useGroup ? fixed(Math.abs(x), dec) : Math.abs(x).toFixed(dec);
    const sign = x < 0 && !neg && Number(abs.replace(/,/g, '')) !== 0 ? '-' : '';
    const lead = neg ? body.split(/[0#]/)[0] : '';
    out = `${lead}${sign}${abs}${isPct ? '%' : ''}`;
  }
  return { text: out, red };
}

function cellText(cell) {
  let v = cell.value;
  if (v && typeof v === 'object') {
    if ('result' in v || 'formula' in v || 'sharedFormula' in v) v = cellResult(cell);
    else if (Array.isArray(v.richText)) v = v.richText.map((t) => t.text).join('');
    else if ('text' in v) v = v.text;
  }
  if (v === null || v === undefined) return { text: '', red: false, num: false };
  if (typeof v === 'number') return { ...formatNumber(v, cell.numFmt), num: true };
  return { text: String(v), red: false, num: false };
}

function borderCss(b) {
  if (!b || !b.style) return 'none';
  const w = b.style === 'medium' ? 2 : b.style === 'thick' ? 3 : 1;
  return `${w}px solid ${argb(b.color) || '#000'}`;
}

/**
 * @param ws ExcelJS 워크시트 (buildWorkbook 결과)
 * @param layout buildWorkbook().layout (breaks, lastRow)
 * @param images { key: dataURL } — 시트 이미지(그래프)와 같은 순서·위치로 배치
 * @returns {{ pages: string[], pageSize: {w,h}, scale: number }}
 */
export function sheetPages(ws, layout, { footer = '', imageAlt = null } = {}) {
  const lastRow = layout.lastRow;
  const lastCol = 30; // A~AD (인쇄 영역)
  const widths = Array.from({ length: lastCol }, (_, i) => colPx(ws.getColumn(i + 1).width));
  const totalW = widths.reduce((s, w) => s + w, 0);
  const defH = ws.properties.defaultRowHeight || 17;
  const heights = [0];
  for (let r = 1; r <= lastRow; r += 1) heights[r] = rowPx(ws.getRow(r).height || defH);

  // 너비 1페이지 맞춤 배율 (시트 px → 용지 px)
  // Excel은 '너비 1페이지 맞춤' 배율을 정수 %로 내림함
  const scale = Math.min(1, Math.floor(((A4.w - 2 * MARGIN.lr) / totalW) * 100) / 100);
  const pageW = A4.w / scale; // 시트 단위 용지 크기
  const pageH = A4.h / scale;
  const padX = (pageW - totalW) / 2; // horizontalCentered
  const padT = (MARGIN.tb + 6.5) / scale; // Excel 인쇄는 위 여백보다 약 5pt 아래에서 시작 (실측)
  const capH = pageH - (2 * MARGIN.tb) / scale; // 인쇄 가능 높이

  // 병합 정보
  const merges = Object.values(ws._merges || {}).map((m) => m.model).filter((m) => m.top <= lastRow && m.left <= lastCol);
  const master = new Map();
  const covered = new Set();
  for (const m of merges) {
    master.set(`${m.top}:${m.left}`, m);
    for (let r = m.top; r <= m.bottom; r += 1) for (let c = m.left; c <= Math.min(m.right, lastCol); c += 1) if (r !== m.top || c !== m.left) covered.add(`${r}:${c}`);
  }

  // 그래프 이미지 위치
  const imgs = (ws.getImages ? ws.getImages() : []).map((im) => {
    const media = ws._workbook.getImage(Number(im.imageId));
    const tl = im.range.tl;
    const row = Math.floor(tl.nativeRow ?? tl.row) + 1;
    const col = Math.floor(tl.nativeCol ?? tl.col);
    const ext = im.range.ext || { width: 580, height: 310 };
    const b64 = media?.base64 || '';
    const src = b64.startsWith('data:') ? b64 : `data:image/${media?.extension || 'png'};base64,${b64}`;
    let h = 0;
    let rEnd = row;
    while (h < ext.height && rEnd <= lastRow) h += heights[rEnd++];
    return { row, rowEnd: rEnd - 1, col, w: ext.width, h: ext.height, src };
  });

  // 페이지 나누기: 수동 나눔 + 높이 초과 시 자동 (병합·이미지 범위는 쪼개지 않음)
  const manual = new Set(layout.breaks || []);
  const blockTop = (r) => {
    let top = r;
    for (const m of merges) if (m.top < r && m.bottom >= r) top = Math.min(top, m.top);
    for (const im of imgs) if (im.row < r && im.rowEnd >= r) top = Math.min(top, im.row);
    return top;
  };
  const pages = [];
  let start = 1;
  let acc = 0;
  for (let r = 1; r <= lastRow; r += 1) {
    if (acc + heights[r] > capH && r > start) {
      const cut = Math.max(start + 1, blockTop(r));
      pages.push([start, cut - 1]);
      start = cut;
      acc = 0;
      for (let k = cut; k < r; k += 1) acc += heights[k];
    }
    acc += heights[r];
    if (manual.has(r) && r < lastRow) {
      pages.push([start, r]);
      start = r + 1;
      acc = 0;
    }
  }
  if (start <= lastRow) pages.push([start, lastRow]);
  // 빈 행만 남은 페이지는 인쇄하지 않음 (Excel도 빈 페이지는 출력하지 않음)
  const hasContent = ([a, b]) => {
    if (imgs.some((im) => im.row >= a && im.row <= b)) return true;
    for (let r = a; r <= b; r += 1) {
      for (let c = 1; c <= lastCol; c += 1) {
        const v = ws.getCell(r, c).value;
        if (v !== null && v !== undefined && v !== '') return true;
      }
    }
    return false;
  };
  for (let i = pages.length - 1; i >= 0; i -= 1) if (!hasContent(pages[i])) pages.splice(i, 1);

  const colgroup = `<colgroup>${widths.map((w) => `<col style="width:${w}px">`).join('')}</colgroup>`;
  const html = pages.map(([r1, r2], pi) => {
    const rowsHtml = [];
    for (let r = r1; r <= r2; r += 1) {
      const tds = [];
      for (let c = 1; c <= lastCol; c += 1) {
        if (covered.has(`${r}:${c}`)) continue;
        const m = master.get(`${r}:${c}`);
        const rs = m ? Math.min(m.bottom, r2) - r + 1 : 1;
        const cs = m ? Math.min(m.right, lastCol) - c + 1 : 1;
        const cell = ws.getCell(r, c);
        const st = [];
        const f = cell.font || {};
        if (f.size) st.push(`font-size:${f.size}pt`);
        if (f.bold) st.push('font-weight:700');
        const { text, red, num } = cellText(cell);
        const color = red ? '#FF0000' : argb(f.color);
        if (color && color !== '#000000') st.push(`color:${color}`);
        const fill = cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' ? argb(cell.fill.fgColor) : null;
        if (fill) st.push(`background:${fill}`);
        // 병합 칸 테두리: 위·왼쪽은 기준 칸, 오른쪽·아래는 끝 칸
        const bTop = cell.border?.top;
        const bLeft = cell.border?.left;
        const bRight = (m ? ws.getCell(r, Math.min(m.right, lastCol)) : cell).border?.right;
        const bBottom = (m ? ws.getCell(Math.min(m.bottom, r2), c) : cell).border?.bottom;
        st.push(`border-top:${borderCss(bTop)}`, `border-left:${borderCss(bLeft)}`, `border-right:${borderCss(bRight)}`, `border-bottom:${borderCss(bBottom)}`);
        const al = cell.alignment || {};
        const h = al.horizontal || (num ? 'right' : 'left');
        st.push(`text-align:${h === 'centerContinuous' ? 'center' : h}`);
        st.push(`vertical-align:${al.vertical === 'top' ? 'top' : al.vertical === 'middle' ? 'middle' : 'bottom'}`);
        const wrap = !!al.wrapText;
        // 줄바꿈 없는 글자는 Excel처럼 옆 빈칸으로 넘침 (왼쪽 정렬 → 오른쪽으로, 오른쪽 정렬 → 왼쪽으로)
        const spill = wrap ? '' : h === 'left' ? ' sx-spill' : h === 'right' ? ' sx-spill-r' : '';
        const inner = text ? `<span class="${wrap ? 'sx-w' : 'sx-n'}${spill}">${esc(text)}</span>` : '';
        tds.push(`<td${rs > 1 ? ` rowspan="${rs}"` : ''}${cs > 1 ? ` colspan="${cs}"` : ''} style="${st.join(';')}">${inner}</td>`);
      }
      rowsHtml.push(`<tr style="height:${heights[r].toFixed(2)}px">${tds.join('')}</tr>`);
    }
    // 이 페이지에 속한 그래프
    let y0 = 0;
    const offY = {};
    for (let r = r1; r <= r2; r += 1) {
      offY[r] = y0;
      y0 += heights[r];
    }
    const imgHtml = imgs
      .filter((im) => im.row >= r1 && im.row <= r2)
      .map((im) => {
        const x = widths.slice(0, im.col).reduce((s, w) => s + w, 0);
        // imageAlt: 그래프 이미지 → 그래프 내용 글자 (PDF 투명 글자 층에서 사용)
        const alt = imageAlt?.get(im.src);
        const altAttr = alt ? ` alt="${esc(alt)}" data-alt="${esc(alt)}"` : '';
        return `<img class="sx-img" src="${im.src}"${altAttr} style="left:${x}px;top:${offY[im.row]}px;width:${im.w}px;height:${im.h}px">`;
      })
      .join('');
    const foot = footer ? `<div class="sx-foot" style="bottom:${(MARGIN.footer / scale).toFixed(1)}px">${esc(footer)}  |  ${pi + 1} / ${pages.length}</div>` : '';
    return `<div class="sx-page" style="width:${pageW.toFixed(1)}px;height:${pageH.toFixed(1)}px">
  <div class="sx-body" style="left:${padX.toFixed(1)}px;top:${padT.toFixed(1)}px;width:${totalW}px">
    <table class="sx-table" style="width:${totalW}px">${colgroup}<tbody>${rowsHtml.join('')}</tbody></table>${imgHtml}
  </div>${foot}
</div>`;
  });

  return { pages: html, pageSize: { w: pageW, h: pageH }, scale };
}

export const SHEET_CSS = `
.sx-page{position:relative;background:#fff;box-sizing:border-box;overflow:hidden;font-family:${FONT_STACK};color:#000}
.sx-body{position:absolute}
.sx-table{border-collapse:collapse;table-layout:fixed;font-size:10pt;line-height:1.15}
.sx-table td{padding:0 2px;overflow:visible;position:relative;box-sizing:border-box}
.sx-n{white-space:nowrap}
.sx-spill{position:absolute;left:2px;top:50%;transform:translateY(-50%);white-space:nowrap;z-index:1}
.sx-spill-r{position:absolute;right:2px;top:50%;transform:translateY(-50%);white-space:nowrap;z-index:1}
.sx-w{white-space:normal;word-break:keep-all;overflow-wrap:anywhere;display:block}
.sx-img{position:absolute;z-index:2;box-sizing:border-box;border:1px solid #D9D9D9;border-radius:6px}
.sx-foot{position:absolute;left:0;right:0;text-align:center;font-size:8pt;color:#000}
`;
