// 미리보기 페이지(.sx-page) → 글자를 읽을 수 있는 PDF
// - 화면 모양: 지금처럼 html2canvas로 그린 페이지 그림(JPEG)을 그대로 씀
// - 글자: 각 줄의 글자를 화면과 같은 위치에 투명하게 겹쳐 넣음 (스캔 문서 OCR 층과 같은 방식)
//   → PDF 뷰어 검색·복사, 텍스트 추출(채점 프로그램 등)이 가능함
// - 한글 글꼴: Pretendard(OFL-1.1). 보고서에 쓴 글자만 PDF에 넣음(subset)
import { PDFDocument } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import fontUrl from 'pretendard/dist/public/static/alternative/Pretendard-Regular.ttf?url';

const A4 = { w: 595.28, h: 841.89 }; // pt

let fontBytes = null;
async function loadFont() {
  if (!fontBytes) {
    const res = await fetch(fontUrl);
    if (!res.ok) throw new Error('PDF 글꼴(Pretendard)을 불러오지 못했습니다.');
    fontBytes = await res.arrayBuffer();
  }
  return fontBytes;
}

/**
 * 페이지 요소 안의 글자를 화면 줄 단위로 모음.
 * 줄바꿈된 칸은 글자마다 위치를 재서 줄별로 나눔.
 * @returns {{ text: string, left: number, top: number, height: number, fontPx: number }[]} (페이지 기준 px)
 */
export function textRuns(pageEl) {
  const base = pageEl.getBoundingClientRect();
  const runs = [];
  const range = document.createRange();
  const walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const s = n.nodeValue;
    if (!s || !s.trim()) continue;
    const cs = getComputedStyle(n.parentElement);
    if (cs.visibility === 'hidden' || cs.display === 'none') continue;
    const fontPx = parseFloat(cs.fontSize) || 12;
    let cur = null;
    for (let i = 0; i < s.length; i += 1) {
      range.setStart(n, i);
      range.setEnd(n, i + 1);
      const r = range.getClientRects()[0];
      if (!r || (r.width === 0 && /\s/.test(s[i]))) {
        if (cur) cur.text += s[i];
        continue;
      }
      const top = r.top - base.top;
      if (cur && Math.abs(top - cur.top) < fontPx * 0.5) {
        cur.text += s[i];
      } else {
        if (cur) runs.push(cur);
        cur = { text: s[i], left: r.left - base.left, top, height: r.height, fontPx };
      }
    }
    if (cur) runs.push(cur);
  }
  // 그래프(그림)는 글자가 없으므로 data-alt의 그래프 내용을 그래프 자리 왼쪽 위부터 줄마다 넣음
  for (const img of pageEl.querySelectorAll('img[data-alt]')) {
    const r = img.getBoundingClientRect();
    const lines = img.dataset.alt.split('\n');
    const fontPx = Math.min(9, (r.height - 4) / Math.max(1, lines.length) / 1.15);
    lines.forEach((text, i) => {
      runs.push({ text, left: r.left - base.left + 3, top: r.top - base.top + 2 + i * fontPx * 1.15, height: fontPx * 1.15, fontPx });
    });
  }
  return runs
    .map((r) => ({ ...r, text: r.text.replace(/\s+/g, ' ').trim() }))
    .filter((r) => r.text);
}

/**
 * @param {{ el: HTMLElement, jpeg: Uint8Array|ArrayBuffer }[]} pages 페이지 요소와 그 그림(JPEG)
 * @param {{ title?: string }} meta
 * @returns {Promise<Uint8Array>}
 */
export async function buildTextPdf(pages, meta = {}) {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(await loadFont(), { subset: true });
  if (meta.title) doc.setTitle(meta.title);
  doc.setLanguage('ko-KR');
  doc.setCreator('COPWC');

  for (const { el, jpeg } of pages) {
    const page = doc.addPage([A4.w, A4.h]);
    const img = await doc.embedJpg(jpeg);
    page.drawImage(img, { x: 0, y: 0, width: A4.w, height: A4.h });
    const s = A4.w / el.offsetWidth; // 페이지 px → pt
    for (const r of textRuns(el)) {
      const size = r.fontPx * s;
      // 글꼴에 없는 글자는 빼지 않고 그대로 넣음(pdf-lib가 대체 글리프로 처리). 투명이라 화면에는 안 보임
      page.drawText(r.text, {
        x: r.left * s,
        y: A4.h - (r.top + r.height * 0.8) * s,
        size,
        font,
        opacity: 0,
      });
    }
  }
  return doc.save();
}
