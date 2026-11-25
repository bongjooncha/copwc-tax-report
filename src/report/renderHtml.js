// 블록 모델 → HTML (미리보기와 PDF 생성에 사용). xlsx와 같은 28칸 격자를 colgroup으로 재현함
import { GRID, COL_WIDTHS } from './model.js';
import { isNum, won, pct, times } from './format.js';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
const total = COL_WIDTHS.reduce((s, w) => s + w, 0);
const colgroup = `<colgroup>${COL_WIDTHS.map((w) => `<col style="width:${((w / total) * 100).toFixed(3)}%">`).join('')}</colgroup>`;

function fmtCell(cell) {
  const v = cell.v;
  if (cell.fmt === 'text' || !isNum(v)) return esc(v ?? '');
  if (cell.fmt === 'pct') return pct(v);
  if (cell.fmt === 'times') return times(v);
  if (cell.fmt === 'num') return String(v);
  return won(v);
}

function gridRow(cells) {
  return `<tr>${cells
    .map((cell) => {
      const numeric = isNum(cell.v) && cell.fmt !== 'text';
      const isHead = cell.st === 'head' || cell.st === 'label';
      const special = typeof cell.v === 'string' && /^(-|자본잠식|흑자전환)$/.test(cell.v);
      const al = cell.al || (isHead ? 'center' : numeric ? 'right' : special ? 'center' : 'left');
      const cls = [`st-${cell.st}`, `al-${al}`, cell.bold ? 'b' : '', cell.small ? 'sm' : '', isNum(cell.v) && cell.v < 0 && cell.fmt !== 'text' ? 'neg' : '']
        .filter(Boolean)
        .join(' ');
      return `<td colspan="${Math.min(cell.s, GRID)}" class="${cls}">${fmtCell(cell)}</td>`;
    })
    .join('')}</tr>`;
}

export function renderHtml(model, images = {}) {
  const out = [];
  // 제목·소제목(과 바로 뒤 주석)은 다음 블록과 묶어 페이지 끝에 홀로 남지 않게 함
  let held = [];
  let parts = out;
  for (const b of model) {
    const heading = b.t === 'section' || b.t === 'sub' || (b.t === 'note' && held.length);
    if (heading) parts = held;
    else if (held.length) {
      parts = [];
    }
    renderBlock(b, parts, images);
    if (!heading && held.length) {
      out.push(`<div class="rp-keep${held[0].includes('rp-break') ? ' rp-break' : ''}">${held.join('').replace(' rp-break', '')}${parts[0] ?? ''}</div>`, ...parts.slice(1));
      held = [];
    }
    parts = out;
  }
  out.push(...held);
  return `<div class="report">${out.join('')}</div>`;
}

function renderBlock(b, parts, images) {
  {
    switch (b.t) {
      case 'title':
        parts.push(`<div class="rp-title rp-blk">${esc(b.text)}</div>`);
        break;
      case 'meta':
        parts.push(`<div class="rp-meta">${esc(b.text)}</div>`);
        break;
      case 'section':
        parts.push(`<h2 class="rp-section${b.pageBreak ? ' rp-break' : ''}">${esc(b.text)}</h2>`);
        break;
      case 'sub':
        parts.push(`<h3 class="rp-sub">${esc(b.text)}</h3>`);
        break;
      case 'note':
        parts.push(
          `<div class="rp-note rp-blk">${b.lines
            .map((t, i) => `<div><span class="mk">${i === 0 ? esc(b.mark) : ''}</span><span>${esc(t)}</span></div>`)
            .join('')}</div>`,
        );
        break;
      case 'grid':
        parts.push(`<table class="rp-grid rp-blk">${colgroup}<tbody>${b.rows.map(gridRow).join('')}</tbody></table>`);
        break;
      case 'para':
        parts.push(
          `<div class="rp-para rp-blk">${b.head ? `<div class="hd">${esc(b.head)}</div>` : ''}${b.lines
            .map((t) => `<p>${esc(t) || '&nbsp;'}</p>`)
            .join('')}</div>`,
        );
        break;
      case 'forms':
        parts.push(
          `<div class="rp-forms">${b.items
            .map((it) => `<div class="rp-blk"><div class="nm"><span>○</span>[${esc(it.name)}]</div><div class="ds">${esc(it.desc)}</div></div>`)
            .join('')}</div>`,
        );
        break;
      case 'cards':
        parts.push(
          `<div class="rp-cards rp-blk">${b.cards
            .map((cd) => `<div class="cd"><div class="t">${esc(cd.title)}</div><div class="v">${esc(cd.value)}</div><div class="d">${esc(cd.delta)}</div></div>`)
            .join('')}</div>`,
        );
        break;
      case 'charts': {
        const keys = b.keys.filter((k) => images[k]);
        for (let i = 0; i < keys.length; i += 2) {
          parts.push(
            `<div class="rp-charts rp-blk">${[keys[i], keys[i + 1]]
              .filter(Boolean)
              .map((k) => `<img src="${images[k]}" alt="${k}">`)
              .join('')}</div>`,
          );
        }
        break;
      }
      case 'sign':
        parts.push(
          `<div class="rp-sign rp-blk"><p class="lead">본 보고서는 샘플세무법인에서 발행한 보고서입니다.</p><p class="dt">발행일 : ${esc(
            b.date,
          )}</p><p class="is">${esc(b.issuer)}</p></div>`,
        );
        break;
      default:
        break;
    }
  }
}
