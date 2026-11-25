// 블록 모델 → xlsx (셀 단위 편집 가능, 출력예시 양식의 열 너비·색·병합 규칙을 따름)
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import { GRID, COL_WIDTHS } from './model.js';
import { CHART_SIZE } from './charts.js';
import { isNum } from './format.js';

const FONT = '맑은 고딕';
const NAVY = 'FF1F2D40';
const BEIGE = 'FFE8E4DC';
const YELLOW = 'FFFFF2CC';
const LINE = { style: 'thin', color: { argb: 'FF8C8C8C' } };
const FIRST = 2; // B열
const LAST = FIRST + GRID - 1; // AC열
const ROW_H = 17;

const font = (o = {}) => ({ name: FONT, size: 10, color: { argb: 'FF000000' }, ...o });
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

// Excel 열 너비(문자 수) → 픽셀 근사
const colPx = (w) => Math.floor(w * 7 + 5);

export async function renderXlsx(model, images, sheetName = '법인세 신고결과') {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'COPWC';
  wb.created = new Date();
  const ws = wb.addWorksheet(sheetName, {
    views: [{ showGridLines: false }],
    pageSetup: {
      paperSize: 9,
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
    },
  });
  ws.getColumn(1).width = 1.16;
  COL_WIDTHS.forEach((w, i) => (ws.getColumn(FIRST + i).width = w));
  const gridPx = COL_WIDTHS.map(colPx);
  const totalPx = gridPx.reduce((s, v) => s + v, 0);

  let r = 1;
  const merge = (row, c1, c2, row2 = row) => {
    if (c2 > c1 || row2 > row) ws.mergeCells(row, c1, row2, c2);
  };
  const line = (text, opt = {}) => {
    const { col = FIRST, f = {}, al = {}, h = ROW_H, toCol = LAST } = opt;
    const cell = ws.getCell(r, col);
    cell.value = text;
    cell.font = font(f);
    cell.alignment = { vertical: 'middle', ...al };
    merge(r, col, toCol);
    ws.getRow(r).height = h;
    r += 1;
    return cell;
  };
  const gap = (h = 8) => {
    ws.getRow(r).height = h;
    r += 1;
  };
  const wrapHeight = (text, chars, size = 10) => {
    let units = 0;
    for (const ch of String(text)) units += /[ㄱ-힝]/.test(ch) ? 1.7 : 0.95;
    units *= size / 10;
    const lines = Math.max(1, Math.ceil(units / Math.max(chars - 3, 10)));
    return Math.max(ROW_H, lines * 14 + 4);
  };
  const charsBetween = (c1, c2) => COL_WIDTHS.slice(c1 - FIRST, c2 - FIRST + 1).reduce((s, w) => s + w, 0);

  for (const b of model) {
    switch (b.t) {
      case 'title': {
        const cell = line(b.text, { f: { bold: true, size: 22, color: { argb: 'FFFFFFFF' } }, al: { horizontal: 'center' }, h: 44 });
        for (let c = FIRST; c <= LAST; c += 1) ws.getCell(cell.row, c).fill = fill(NAVY);
        break;
      }
      case 'meta':
        line(b.text, { f: { size: 9, color: { argb: 'FF595959' } }, al: { horizontal: 'right' } });
        break;
      case 'section':
        gap(14);
        {
          const cell = line(b.text, { f: { bold: true, size: 12, color: { argb: NAVY } }, h: 22 });
          for (let c = FIRST; c <= LAST; c += 1) ws.getCell(cell.row, c).border = { bottom: { style: 'medium', color: { argb: NAVY } } };
        }
        gap(6);
        break;
      case 'sub':
        gap(6);
        line(b.text, { f: { bold: true } });
        break;
      case 'note':
        b.lines.forEach((t, i) => {
          if (i === 0) {
            const m = ws.getCell(r, FIRST);
            m.value = b.mark;
            m.font = font({ size: 9 });
            m.alignment = { horizontal: 'right', vertical: 'top' };
          }
          const h = wrapHeight(t, charsBetween(FIRST + 1, LAST), 9);
          line(t, { col: FIRST + 1, f: { size: 9, color: { argb: 'FF404040' } }, al: { wrapText: true, vertical: 'top' }, h });
        });
        break;
      case 'grid':
        for (const cells of b.rows) writeGridRow(cells);
        break;
      case 'para':
        if (b.head) line(b.head, { f: { size: 11, bold: true } });
        for (const t of b.lines) {
          const h = wrapHeight(t, charsBetween(FIRST, LAST), 10);
          line(t, { f: { size: 10 }, al: { wrapText: true, vertical: 'top' }, h });
        }
        if (b.head) gap(4);
        break;
      case 'forms':
        for (const it of b.items) {
          const m = ws.getCell(r, FIRST);
          m.value = '○';
          m.font = font({ size: 9 });
          m.alignment = { horizontal: 'center' };
          line(`[${it.name}]`, { col: FIRST + 1, f: { size: 9, bold: true } });
          line(it.desc, { col: FIRST + 1, f: { size: 9 }, al: { wrapText: true, vertical: 'top' }, h: wrapHeight(it.desc, charsBetween(FIRST + 1, LAST), 9) });
          gap(4);
        }
        break;
      case 'cards': {
        const span = Math.floor(GRID / Math.max(b.cards.length, 1));
        const rows = [r, r + 1, r + 2];
        b.cards.forEach((cd, i) => {
          const c1 = FIRST + i * span;
          const c2 = i === b.cards.length - 1 ? LAST : c1 + span - 1;
          const put = (row, v, f, h) => {
            const cell = ws.getCell(row, c1);
            cell.value = v;
            cell.font = font(f);
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            merge(row, c1, c2);
            ws.getRow(row).height = h;
            for (let c = c1; c <= c2; c += 1) {
              const x = ws.getCell(row, c);
              if (row === rows[0]) x.fill = fill(BEIGE);
              x.border = {
                top: row === rows[0] ? LINE : undefined,
                bottom: row === rows[2] ? LINE : undefined,
                left: c === c1 ? LINE : undefined,
                right: c === c2 ? LINE : undefined,
              };
            }
          };
          put(rows[0], cd.title, { bold: true }, 20);
          put(rows[1], cd.value, { bold: true, size: 16, color: { argb: NAVY } }, 32);
          put(rows[2], cd.delta, { size: 9, color: { argb: 'FF595959' } }, 20);
        });
        r += 3;
        break;
      }
      case 'charts': {
        const keys = b.keys.filter((k) => images[k]);
        const halfPx = totalPx / 2;
        const imgW = halfPx - 10;
        const imgH = Math.round((imgW * CHART_SIZE.h) / CHART_SIZE.w);
        const rowsPer = Math.ceil(imgH / 20) + 1;
        for (let i = 0; i < keys.length; i += 2) {
          for (let k = 0; k < rowsPer; k += 1) ws.getRow(r + k).height = 15;
          [keys[i], keys[i + 1]].forEach((key, j) => {
            if (!key) return;
            const id = wb.addImage({ base64: images[key], extension: 'png' });
            ws.addImage(id, { tl: { col: pxToCol(j * halfPx + 4, gridPx) + FIRST - 1, row: r - 1 + 0.2 }, ext: { width: imgW, height: imgH } });
          });
          r += rowsPer;
        }
        break;
      }
      case 'sign': {
        gap(24);
        line('본 보고서는 샘플세무법인에서 발행한 보고서입니다.', { f: { size: 10, color: { argb: 'FF404040' } }, al: { horizontal: 'center' } });
        gap(10);
        line(`발행일 : ${b.date}`, { f: { size: 11 }, al: { horizontal: 'right' } });
        line(b.issuer, { f: { size: 16, bold: true, color: { argb: NAVY } }, al: { horizontal: 'right' }, h: 28 });
        break;
      }
      default:
        break;
    }
  }
  ws.pageSetup.printArea = `A1:AC${r}`;

  function writeGridRow(cells) {
    let c = FIRST;
    let h = ROW_H;
    cells.forEach((cell, idx) => {
      const c1 = c;
      const c2 = Math.min(c + cell.s - 1, LAST);
      const x = ws.getCell(r, c1);
      const st = cell.st;
      const isHead = st === 'head' || st === 'label';
      let v = cell.v;
      let numFmt;
      if (isNum(v) && cell.fmt !== 'text') {
        if (cell.fmt === 'pct') { v /= 100; numFmt = '0.00%'; }
        else if (cell.fmt === 'times') numFmt = '0.00"회"';
        else if (cell.fmt === 'num') numFmt = 'General';
        else numFmt = '#,##0;-#,##0;0';
      } else if (v === null || v === undefined) v = '';
      x.value = v;
      if (numFmt) x.numFmt = numFmt;
      x.font = font({ bold: isHead || !!cell.bold, size: cell.small ? 9 : 10 });
      const numeric = isNum(cell.v) && cell.fmt !== 'text';
      const horizontal = cell.al || (isHead ? 'center' : numeric ? 'right' : typeof v === 'string' && /^(-|자본잠식|흑자전환)$/.test(v) ? 'center' : 'left');
      x.alignment = { horizontal, vertical: 'middle', wrapText: !!cell.wrap, indent: horizontal === 'left' ? 1 : 0 };
      if (cell.wrap && typeof v === 'string') h = Math.max(h, wrapHeight(v, charsBetween(c1, c2), cell.small ? 9 : 10));
      merge(r, c1, c2);
      for (let cc = c1; cc <= c2; cc += 1) {
        const y = ws.getCell(r, cc);
        if (isHead) y.fill = fill(BEIGE);
        else if (st === 'hl') y.fill = fill(YELLOW);
        y.border = {
          top: LINE,
          bottom: LINE,
          left: cc === c1 && idx === 0 ? LINE : cc === c1 ? { style: 'hair', color: { argb: 'FFBFBFBF' } } : undefined,
          right: cc === c2 && idx === cells.length - 1 ? LINE : undefined,
        };
      }
      c = c2 + 1;
    });
    ws.getRow(r).height = h;
    r += 1;
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function pxToCol(px, gridPx) {
  let acc = 0;
  for (let i = 0; i < gridPx.length; i += 1) {
    if (acc + gridPx[i] > px) return i + (px - acc) / gridPx[i];
    acc += gridPx[i];
  }
  return gridPx.length;
}
