// JSON → 편집 가능한 표
// - 객체: 항목/값 표 (값이 객체면 안쪽에 다시 표)
// - "항목 → {당기, 전기, 전전기}" 같은 객체의 객체: 행 머리가 있는 격자 표
// - 객체 배열: 격자 표 (행 추가·삭제)
// - 값 배열: 목록 (추가·삭제)
// 값 수정은 원래 자료형을 유지함 (숫자 칸은 콤마 입력 허용, 비우면 null)

const isPlain = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);
const isPrim = (v) => v === null || ['string', 'number', 'boolean'].includes(typeof v);
const isFlatObj = (v) => isPlain(v) && Object.values(v).every(isPrim);

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

export function display(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return Math.abs(v) >= 10000 ? v.toLocaleString('ko-KR', { maximumFractionDigits: 6 }) : String(v);
  return String(v);
}

export function parseInput(text, prev) {
  const t = text.trim();
  const numeric = /^[-+]?[\d,]*\.?\d+$/.test(t) && /\d/.test(t);
  if (typeof prev === 'number' || prev === null || prev === undefined) {
    if (t === '') return null;
    if (numeric) return Number(t.replace(/,/g, ''));
    return text;
  }
  if (typeof prev === 'boolean') return t === 'true' ? true : t === 'false' ? false : text;
  return text;
}

function getIn(obj, path) {
  return path.reduce((o, k) => (o == null ? undefined : o[k]), obj);
}

/**
 * @param {HTMLElement} root
 * @param {object} data   편집 대상 (직접 수정됨)
 * @param {(data:object)=>void} onChange
 * @param {object} templates  빈 배열에 행을 추가할 때 쓸 항목 모양 (경로 → 객체)
 */
export function mountJsonTable(root, data, onChange, templates = {}) {
  const rerender = () => {
    const open = new Set([...root.querySelectorAll('details[open]')].map((d) => d.dataset.key));
    const first = !root.dataset.mounted;
    root.innerHTML = '';
    root.dataset.mounted = '1';
    for (const key of Object.keys(data)) {
      const det = document.createElement('details');
      det.className = 'jt-sec';
      det.dataset.key = key;
      if (first || open.has(key)) det.open = true;
      const v = data[key];
      det.innerHTML = `<summary><span class="jt-key">${esc(key)}</span><span class="jt-meta">${meta(v)}</span></summary>`;
      const body = document.createElement('div');
      body.className = 'jt-body';
      body.appendChild(node(v, [key]));
      det.appendChild(body);
      root.appendChild(det);
    }
  };

  const commit = (structural) => {
    onChange(data);
    if (structural) rerender();
  };

  const setAt = (path, value) => {
    const parent = getIn(data, path.slice(0, -1));
    parent[path[path.length - 1]] = value;
  };

  function input(path) {
    const el = document.createElement('input');
    const v = getIn(data, path);
    el.className = `jt-in${typeof v === 'number' ? ' num' : ''}${typeof v === 'number' && v < 0 ? ' neg' : ''}`;
    el.value = display(v);
    el.placeholder = v === null ? 'null' : '';
    el.spellcheck = false;
    el.addEventListener('input', () => {
      const nv = parseInput(el.value, getIn(data, path));
      setAt(path, nv);
      el.classList.toggle('neg', typeof nv === 'number' && nv < 0);
      el.classList.toggle('num', typeof nv === 'number');
      el.classList.toggle('bad', typeof nv === 'string' && typeof v === 'number');
      commit(false);
    });
    el.addEventListener('blur', () => {
      const cur = getIn(data, path);
      if (typeof cur === 'number') el.value = display(cur);
    });
    return el;
  }

  function delBtn(onClick, label = '삭제') {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'jt-del';
    b.title = label;
    b.setAttribute('aria-label', label);
    b.textContent = '✕';
    b.addEventListener('click', onClick);
    return b;
  }

  function addBtn(text, onClick) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn btn-ghost btn-xs jt-add';
    b.textContent = `+ ${text}`;
    b.addEventListener('click', onClick);
    return b;
  }

  function blankLike(item) {
    if (!isPlain(item)) return typeof item === 'number' ? 0 : '';
    return Object.fromEntries(Object.entries(item).map(([k, v]) => [k, typeof v === 'number' ? 0 : v === null ? null : '']));
  }

  function node(v, path) {
    if (isPrim(v)) return input(path);

    // 객체 배열 → 격자
    if (Array.isArray(v) && (v.length === 0 ? isPlain(templates[path.join('.')]) : v.every(isFlatObj))) {
      const tpl = templates[path.join('.')];
      const cols = [...new Set([...(tpl ? Object.keys(tpl) : []), ...v.flatMap((r) => Object.keys(r))])];
      const wrap = document.createElement('div');
      const tbl = document.createElement('table');
      tbl.className = 'jt-grid';
      tbl.innerHTML = `<thead><tr><th class="jt-idx">#</th>${cols.map((c) => `<th>${esc(c)}</th>`).join('')}<th class="jt-act"></th></tr></thead>`;
      const tb = document.createElement('tbody');
      v.forEach((row, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td class="jt-idx">${i + 1}</td>`;
        cols.forEach((c) => {
          if (!(c in row)) row[c] = tpl && c in tpl ? blankLike({ [c]: tpl[c] })[c] : null;
          const td = document.createElement('td');
          td.appendChild(input([...path, i, c]));
          tr.appendChild(td);
        });
        const td = document.createElement('td');
        td.className = 'jt-act';
        td.appendChild(delBtn(() => { v.splice(i, 1); commit(true); }, `${i + 1}행 삭제`));
        tr.appendChild(td);
        tb.appendChild(tr);
      });
      if (!v.length) tb.innerHTML = `<tr><td class="jt-empty" colspan="${cols.length + 2}">항목 없음</td></tr>`;
      tbl.appendChild(tb);
      wrap.appendChild(scroller(tbl));
      wrap.appendChild(addBtn('행 추가', () => { v.push(blankLike(v[0] || tpl)); commit(true); }));
      return wrap;
    }

    // 값 배열 → 목록
    if (Array.isArray(v) && v.every(isPrim)) {
      const wrap = document.createElement('div');
      wrap.className = 'jt-list';
      v.forEach((_, i) => {
        const row = document.createElement('div');
        row.className = 'jt-li';
        row.innerHTML = `<span class="jt-idx">${i + 1}</span>`;
        const inp = input([...path, i]);
        row.appendChild(inp);
        row.appendChild(delBtn(() => { v.splice(i, 1); commit(true); }, `${i + 1}번 삭제`));
        wrap.appendChild(row);
      });
      if (!v.length) wrap.insertAdjacentHTML('beforeend', '<div class="jt-empty">항목 없음</div>');
      wrap.appendChild(addBtn('항목 추가', () => { v.push(''); commit(true); }));
      return wrap;
    }

    // 객체의 객체 (모두 평평한 객체) → 행 머리 격자
    if (isPlain(v) && Object.keys(v).length && Object.values(v).every(isFlatObj)) {
      const cols = [...new Set(Object.values(v).flatMap((r) => Object.keys(r)))];
      const tbl = document.createElement('table');
      tbl.className = 'jt-grid';
      tbl.innerHTML = `<thead><tr><th>항목</th>${cols.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead>`;
      const tb = document.createElement('tbody');
      for (const [rk, row] of Object.entries(v)) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<th class="jt-rk">${esc(rk)}</th>`;
        cols.forEach((c) => {
          if (!(c in row)) row[c] = null;
          const td = document.createElement('td');
          td.appendChild(input([...path, rk, c]));
          tr.appendChild(td);
        });
        tb.appendChild(tr);
      }
      tbl.appendChild(tb);
      return scroller(tbl);
    }

    // 일반 객체 → 항목/값
    if (isPlain(v)) {
      const tbl = document.createElement('table');
      tbl.className = 'jt-kv';
      const tb = document.createElement('tbody');
      for (const [k, val] of Object.entries(v)) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<th>${esc(k)}</th>`;
        const td = document.createElement('td');
        td.appendChild(node(val, [...path, k]));
        tr.appendChild(td);
        tb.appendChild(tr);
      }
      tbl.appendChild(tb);
      return tbl;
    }

    // 섞인 배열 등: 항목별로
    const wrap = document.createElement('div');
    v.forEach((item, i) => {
      const box = document.createElement('div');
      box.className = 'jt-item';
      box.innerHTML = `<div class="jt-item-h">#${i + 1}</div>`;
      box.appendChild(node(item, [...path, i]));
      wrap.appendChild(box);
    });
    return wrap;
  }

  function scroller(el) {
    const s = document.createElement('div');
    s.className = 'jt-scroll';
    s.appendChild(el);
    return s;
  }

  rerender();
  return { rerender };
}

function meta(v) {
  if (Array.isArray(v)) return `목록 · ${v.length}개`;
  if (isPlain(v)) return `${Object.keys(v).length}개 항목`;
  return typeof v === 'number' ? '숫자' : '값';
}
