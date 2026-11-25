// 팝업: X 버튼, 바깥 영역 클릭, Esc 키로 닫힘
export function openModal({ title, body, footer, size = 'md', onClose }) {
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `
    <div class="modal modal-${size}" role="dialog" aria-modal="true" aria-label="${title}">
      <div class="modal-head">
        <h3>${title}</h3>
        <button type="button" class="icon-btn modal-x" aria-label="닫기">✕</button>
      </div>
      <div class="modal-body"></div>
      <div class="modal-foot"></div>
    </div>`;
  const bodyEl = overlay.querySelector('.modal-body');
  const footEl = overlay.querySelector('.modal-foot');
  if (typeof body === 'string') bodyEl.innerHTML = body;
  else if (body) bodyEl.appendChild(body);
  if (footer) footEl.appendChild(footer);
  else footEl.remove();

  const close = () => {
    document.removeEventListener('keydown', onKey);
    overlay.classList.add('closing');
    setTimeout(() => overlay.remove(), 120);
    onClose?.();
  };
  const onKey = (e) => {
    if (e.key === 'Escape') close();
  };
  overlay.addEventListener('mousedown', (e) => {
    if (e.target === overlay) close();
  });
  overlay.querySelector('.modal-x').addEventListener('click', close);
  document.addEventListener('keydown', onKey);
  document.body.appendChild(overlay);
  requestAnimationFrame(() => overlay.classList.add('open'));
  return { close, el: overlay, body: bodyEl };
}

export function toast(msg, kind = 'ok') {
  const t = document.createElement('div');
  t.className = `toast toast-${kind}`;
  t.textContent = msg;
  document.body.appendChild(t);
  requestAnimationFrame(() => t.classList.add('show'));
  setTimeout(() => {
    t.classList.remove('show');
    setTimeout(() => t.remove(), 250);
  }, 2600);
}
