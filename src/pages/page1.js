import { validateReport } from '../report/validate.js';
// 1차 페이지: 법인세 분석 보고서 출력 (위) + 조정료 청구서 (아래, 개발 예정)
import { SAMPLE_REPORT, SAMPLE_JSON_TEXT } from '../report/sample.js';
import { mountJsonTable } from '../ui/jsonTable.js';
import { openModal, toast } from '../ui/modal.js';

const STORE_KEY = 'copwc.page1.json';
const store = {
  get() { try { return localStorage.getItem(STORE_KEY) || ''; } catch { return ''; } },
  set(v) { try { localStorage.setItem(STORE_KEY, v); } catch { /* 저장 불가 환경 */ } },
};

// 빈 배열에 행을 추가할 때 쓸 항목 모양
function buildTemplates(obj, path = [], out = {}) {
  if (Array.isArray(obj)) {
    if (obj.length && typeof obj[0] === 'object' && obj[0] && !Array.isArray(obj[0])) out[path.join('.')] = obj[0];
  } else if (obj && typeof obj === 'object') {
    for (const [k, v] of Object.entries(obj)) buildTemplates(v, [...path, k], out);
  }
  return out;
}
const TEMPLATES = {
  ...buildTemplates(SAMPLE_REPORT),
  '소득금액조정.손금산입및익금불산입': { 과목: '', 금액: 0, 소득처분: '' },
};

export function renderPage1(root) {
  root.innerHTML = `
    <section class="card" aria-labelledby="t-report">
      <div class="card-head">
        <div>
          <h2 id="t-report">법인세 분석 보고서</h2>
          <p class="sub">세무조정계산서 자료를 JSON으로 입력하면, 법인세 신고 결과 보고서 양식의 <b>xlsx</b>·<b>pdf</b> 파일로 출력합니다.</p>
        </div>
        <div class="actions">
          <button type="button" class="btn btn-ghost" id="btn-example">형식 예시</button>
          <button type="button" class="btn btn-ghost" id="btn-preview" disabled>미리보기</button>
          <button type="button" class="btn btn-primary" id="btn-export" disabled>출력</button>
        </div>
      </div>

      <section class="review-box" aria-live="polite"><strong>자료 검토</strong><p class="review-summary">JSON을 입력하면 누락과 산식 불일치를 확인합니다.</p><details class="review-details"><summary>확인할 항목 보기</summary><ul></ul></details><p class="sub">확인이 필요한 값은 아래 표 또는 JSON에서 직접 수정하세요. 원본 서식 간 비교와 가려진 값은 원본을 직접 확인해야 합니다.</p></section><div class="workspace">
        <div class="pane pane-json">
          <div class="pane-head">
            <span class="pane-title">JSON 입력</span>
            <div class="pane-tools">
              <label class="btn btn-ghost btn-xs" for="json-file">불러오기</label>
              <input type="file" id="json-file" accept=".json,application/json" hidden>
              <button type="button" class="btn btn-ghost btn-xs" id="btn-save-json">JSON 저장</button>
              <button type="button" class="btn btn-ghost btn-xs" id="btn-format">정렬</button>
              <button type="button" class="btn btn-ghost btn-xs" id="btn-clear">지우기</button>
            </div>
          </div>
          <textarea id="json-input" spellcheck="false" autocomplete="off"
            placeholder='여기에 JSON을 붙여넣으세요.&#10;입력 형식은 [형식 예시] 버튼에서 확인할 수 있습니다.&#10;&#10;{&#10;  "사업연도": { "당기": 2025, "전기": 2024, "전전기": null },&#10;  ...&#10;}'></textarea>
          <div class="status" id="json-status" role="status"><span class="dot"></span><span class="msg">입력 대기</span></div>
        </div>
        <div class="pane pane-table">
          <div class="pane-head">
            <span class="pane-title">표 보기 · 수정</span>
            <span class="pane-hint">표에서 고친 값은 JSON 입력칸에도 바로 반영됩니다</span>
          </div>
          <div class="table-host" id="table-host">
            <div class="empty">
              <div class="empty-icon" aria-hidden="true">{ }</div>
              <p>JSON을 입력하면 이곳에 표로 표시됩니다.</p>
              <button type="button" class="btn btn-ghost btn-sm" id="btn-empty-example">형식 예시 보기</button>
            </div>
          </div>
        </div>
      </div>
    </section>

    <section class="card card-pending" aria-labelledby="t-fee">
      <div class="card-head">
        <div>
          <h2 id="t-fee">조정료 청구서</h2>
          <p class="sub">세무조정 보수 산정과 청구서 출력</p>
        </div>
        <span class="badge">개발 예정</span>
      </div>
      <div class="pending-box">개발 예정</div>
    </section>
  `;

  const ta = root.querySelector('#json-input');
  const status = root.querySelector('#json-status');
  const host = root.querySelector('#table-host');
  const btnExport = root.querySelector('#btn-export');
  const btnPreview = root.querySelector('#btn-preview');
  let data = null;
  let timer = null;

  const setStatus = (kind, msg) => {
    status.className = `status ${kind}`;
    status.querySelector('.msg').textContent = msg;
  };

  const emptyHtml = host.innerHTML;
  const bindEmpty = () => host.querySelector('#btn-empty-example')?.addEventListener('click', showExample);

  const parse = () => {
    const text = ta.value;
    store.set(text);
    if (!text.trim()) {
      data = null;
      setStatus('', '입력 대기');
      host.innerHTML = emptyHtml;
      bindEmpty();
      return toggle();
    }
    try {
      const parsed = JSON.parse(text);
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw new Error('최상위 값은 { } 객체여야 합니다.');
      data = parsed;
      const keys = Object.keys(data).length;
      setStatus('ok', `올바른 JSON · 항목 ${keys}개`);
      host.innerHTML = '';
      const tableRoot = document.createElement('div');
      tableRoot.className = 'jt';
      host.appendChild(tableRoot);
      mountJsonTable(tableRoot, data, (d) => {
        ta.value = JSON.stringify(d, null, 2);
        store.set(ta.value);
        refreshReview();
      }, TEMPLATES);
    } catch (e) {
      data = null;
      setStatus('err', describeJsonError(e, text));
    }
    toggle();
  };

  const refreshReview = () => {
    const summary = root.querySelector('.review-summary'); const list = root.querySelector('.review-details ul'); list.replaceChildren();
    if (!data) { summary.textContent = 'JSON을 입력하면 누락과 산식 불일치를 확인합니다.'; return; }
    const issues = validateReport(data); summary.textContent = issues.length ? `확인할 항목 ${issues.length}건` : '기본 검토에서 누락이나 산식 불일치가 발견되지 않았습니다.';
    for (const issue of issues) { const li = document.createElement('li'); li.textContent = `[${issue.kind}] ${issue.message}`; list.appendChild(li); }
  };
  const toggle = () => {
    refreshReview();
    btnExport.disabled = !data;
    btnPreview.disabled = !data;
  };

  ta.addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(parse, 300);
  });
  ta.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      const { selectionStart: s, selectionEnd: en } = ta;
      ta.setRangeText('  ', s, en, 'end');
    }
  });

  root.querySelector('#json-file').addEventListener('change', async (e) => { const file = e.target.files[0]; if (!file) return; ta.value = await file.text(); parse(); e.target.value = ''; });
  root.querySelector('#btn-save-json').addEventListener('click', () => { if (!data) return toast('먼저 올바른 JSON을 입력해 주세요.', 'err'); const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })); const a = document.createElement('a'); a.href = url; a.download = '보고서입력.json'; a.click(); setTimeout(() => URL.revokeObjectURL(url), 4000); });
  root.querySelector('#btn-format').addEventListener('click', () => {
    if (!data) return toast('올바른 JSON이 아니라 정렬할 수 없습니다.', 'err');
    ta.value = JSON.stringify(data, null, 2);
    store.set(ta.value);
  });
  root.querySelector('#btn-clear').addEventListener('click', () => {
    ta.value = '';
    parse();
    ta.focus();
  });

  function showExample() {
    const body = document.createElement('div');
    body.innerHTML = `
      <ul class="example-notes">
        <li>금액은 <b>원 단위 정수</b>, 손실은 <b>음수</b>로 입력합니다.</li>
        <li>연도 열은 <code>당기</code> · <code>전기</code> · <code>전전기</code>입니다. 자료가 없는 연도는 <code>null</code>로 둡니다.</li>
        <li>재무비율·평균·증감률·그래프는 입력값으로 자동 계산합니다.</li>
        <li><code>요약코멘트</code>를 빈 목록으로 두면 숫자를 바탕으로 자동 작성합니다.</li>
      </ul>
      <pre class="code"><code></code></pre>`;
    body.querySelector('pre.code code').textContent = SAMPLE_JSON_TEXT;
    const foot = document.createElement('div');
    foot.className = 'foot-row';
    foot.innerHTML = `
      <button type="button" class="btn btn-ghost" data-act="copy">복사</button>
      <button type="button" class="btn btn-primary" data-act="use">입력칸에 넣기</button>`;
    const m = openModal({ title: '형식 예시 (JSON)', body, footer: foot, size: 'lg' });
    foot.querySelector('[data-act=copy]').addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(SAMPLE_JSON_TEXT);
        toast('예시 JSON을 복사했습니다.');
      } catch {
        toast('복사 권한이 없어 복사하지 못했습니다.', 'err');
      }
    });
    foot.querySelector('[data-act=use]').addEventListener('click', () => {
      if (ta.value.trim() && !confirm('입력칸의 내용을 예시로 바꿀까요?')) return;
      ta.value = SAMPLE_JSON_TEXT;
      parse();
      m.close();
    });
  }
  root.querySelector('#btn-example').addEventListener('click', showExample);
  bindEmpty();

  btnPreview.addEventListener('click', async () => {
    if (!data) return;
    const body = document.createElement('div');
    body.className = 'preview-wrap';
    body.innerHTML = '<div class="loading">보고서를 만드는 중…</div>';
    openModal({ title: '보고서 미리보기', body, size: 'xl' });
    try {
      const { previewHtml } = await import('../report/export.js');
      body.innerHTML = await previewHtml(structuredClone(data), Math.max(320, (body.clientWidth || 900) - 24));
    } catch (e) {
      console.error(e);
      body.innerHTML = `<div class="loading err">미리보기를 만들지 못했습니다: ${e.message}</div>`;
    }
  });

  btnExport.addEventListener('click', () => {
    if (!data) return;
    const body = document.createElement('div');
    body.innerHTML = `
      <p class="sub">출력할 파일 형식을 선택하세요. 둘 다 선택할 수 있습니다.</p>
      <div class="checks">
        <label class="check"><input type="checkbox" name="fmt" value="xlsx" checked>
          <span class="box" aria-hidden="true"></span>
          <span><b>Excel (xlsx)</b><small>셀 단위로 직접 수정 가능</small></span></label>
        <label class="check"><input type="checkbox" name="fmt" value="pdf">
          <span class="box" aria-hidden="true"></span>
          <span><b>PDF</b><small>A4 세로, 인쇄·전달용</small></span></label>
      </div>`;
    const foot = document.createElement('div');
    foot.className = 'foot-row';
    foot.innerHTML = `
      <button type="button" class="btn btn-ghost" data-act="cancel">취소</button>
      <button type="button" class="btn btn-primary" data-act="go">다운로드</button>`;
    const m = openModal({ title: '보고서 출력', body, footer: foot, size: 'sm' });
    const go = foot.querySelector('[data-act=go]');
    const sync = () => {
      go.disabled = !body.querySelectorAll('input[name=fmt]:checked').length;
    };
    body.addEventListener('change', sync);
    foot.querySelector('[data-act=cancel]').addEventListener('click', m.close);
    go.addEventListener('click', async () => {
      const sel = [...body.querySelectorAll('input[name=fmt]:checked')].map((i) => i.value);
      go.disabled = true;
      go.textContent = '만드는 중…';
      try {
        const { exportReport } = await import('../report/export.js');
        const files = await exportReport(structuredClone(data), { xlsx: sel.includes('xlsx'), pdf: sel.includes('pdf') });
        m.close();
        toast(`다운로드: ${files.join(', ')}`);
      } catch (e) {
        console.error(e);
        toast(`출력하지 못했습니다: ${e.message}`, 'err');
        go.disabled = false;
        go.textContent = '다운로드';
      }
    });
  });

  // 이전 입력 복원
  const saved = store.get();
  if (saved) {
    ta.value = saved;
    parse();
  }
}

function describeJsonError(e, text) {
  const m = /position (\d+)/.exec(e.message);
  if (m) {
    const pos = Number(m[1]);
    const before = text.slice(0, pos);
    const line = before.split('\n').length;
    const col = pos - before.lastIndexOf('\n');
    return `JSON 형식 오류 · ${line}행 ${col}열 근처`;
  }
  const lc = /line (\d+) column (\d+)/.exec(e.message);
  if (lc) return `JSON 형식 오류 · ${lc[1]}행 ${lc[2]}열 근처`;
  return `JSON 형식 오류 · ${e.message}`;
}
