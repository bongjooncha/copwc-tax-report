// JSON → 보고서 (xlsx / pdf / 미리보기)
// xlsx·pdf·미리보기 모두 같은 워크북(buildWorkbook: 출력예시.xlsx 템플릿 기반)에서 만듦.
// PDF는 그 워크북 시트를 Excel 인쇄 설정(A4, 너비 1페이지 맞춤, 페이지 나눔)대로 그린 페이지를 이미지로 담고,
// 그 위에 같은 위치의 투명 글자 층을 얹어 텍스트 검색·추출이 되게 함(pdfText.js).
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import html2canvas from 'html2canvas';
import { compute } from './compute.js';
import { renderCharts, chartTexts } from './charts.js';
import { buildWorkbook } from './workbook.js';
import { sheetPages, SHEET_CSS } from './sheetHtml.js';
import { buildTextPdf } from './pdfText.js';

let templateBuf = null;
async function template() {
  if (!templateBuf) {
    const res = await fetch(`${import.meta.env.BASE_URL}report-template.xlsx`);
    if (!res.ok) throw new Error('출력 양식(report-template.xlsx)을 불러오지 못했습니다.');
    templateBuf = await res.arrayBuffer();
  }
  return templateBuf.slice(0);
}

export async function prepare(data) {
  const c = compute(data);
  const images = renderCharts(c);
  const texts = chartTexts(c);
  // 그래프 이미지(dataURL) → 그래프 내용 글자
  const imageAlt = new Map(Object.entries(images).filter(([k]) => texts[k]).map(([k, url]) => [url, texts[k]]));
  const { wb, layout } = await buildWorkbook(ExcelJS, await template(), data, images);
  const ws = wb.getWorksheet(layout.sheetName);
  const issuer = data?.보고서정보?.발행기관 || '샘플세무법인';
  return { c, wb, ws, layout, images, issuer, imageAlt };
}

export function fileBase(c) {
  const name = String(c.d.법인기본사항?.법인명 || '법인')
    .replace(/\(주\)|주식회사|\([^)]*\)/g, '')
    .replace(/[\\/:*?"<>|()\s]/g, '')
    .trim() || '법인';
  return `법인세보고서_${name}_${c.years.당기 ?? ''}`;
}

export async function previewHtml(data, viewWidth = 900) {
  const { ws, layout, issuer, imageAlt } = await prepare(data);
  const { pages, pageSize } = sheetPages(ws, layout, { footer: issuer, imageAlt });
  const zoom = Math.min(1, viewWidth / pageSize.w).toFixed(4);
  return `<style>${SHEET_CSS}.sx-preview{display:flex;flex-direction:column;align-items:center;gap:14px;padding:8px 0}.sx-preview .sx-page{zoom:${zoom};box-shadow:0 1px 6px rgba(0,0,0,.18)}</style><div class="sx-preview">${pages.join('')}</div>`;
}

export async function exportReport(data, { xlsx, pdf }) {
  const { c, wb, ws, layout, issuer, imageAlt } = await prepare(data);
  const base = fileBase(c);
  const done = [];
  if (xlsx) {
    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${base}.xlsx`);
    done.push(`${base}.xlsx`);
  }
  if (pdf) {
    const { pages, pageSize } = sheetPages(ws, layout, { footer: issuer, imageAlt });
    const holder = document.createElement('div');
    holder.style.cssText = 'position:fixed;left:-20000px;top:0;';
    holder.innerHTML = `<style>${SHEET_CSS}</style>${pages.join('')}`;
    document.body.appendChild(holder);
    try {
      if (document.fonts?.ready) await document.fonts.ready;
      const els = [...holder.querySelectorAll('.sx-page')];
      // 용지 px의 약 2배 해상도로 그림
      const scale = Math.max(1, (2 * 794) / pageSize.w);
      const pdfPages = [];
      for (const el of els) {
        const canvas = await html2canvas(el, { scale, backgroundColor: '#ffffff', useCORS: true, logging: false });
        const jpeg = await new Promise((res) => canvas.toBlob((b) => res(b.arrayBuffer()), 'image/jpeg', 0.92));
        pdfPages.push({ el, jpeg: await jpeg });
      }
      const bytes = await buildTextPdf(pdfPages, { title: base });
      if (xlsx) await new Promise((res) => setTimeout(res, 400)); // 연속 다운로드 차단 방지
      download(new Blob([bytes], { type: 'application/pdf' }), `${base}.pdf`);
      done.push(`${base}.pdf`);
    } finally {
      holder.remove();
    }
  }
  return done;
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
