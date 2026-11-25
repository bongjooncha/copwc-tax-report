import './style.css';
import './report/report.css';
import { renderPage1 } from './pages/page1.js';

const PAGES = {
  1: { label: '1차', render: renderPage1 },
  2: { label: '2차' },
  3: { label: '3차' },
  4: { label: '4차' },
};

const app = document.querySelector('#app');
app.innerHTML = `
  <header class="topbar">
    <div class="topbar-in">
      <a class="brand" href="#/1" aria-label="COPWC 홈">COPWC</a>
      <nav class="nav" aria-label="페이지">
        ${Object.entries(PAGES).map(([k, p]) => `<a href="#/${k}" data-page="${k}">${p.label}</a>`).join('')}
      </nav>
    </div>
  </header>
  <main class="main" id="page"></main>`;

const pageEl = app.querySelector('#page');

function route() {
  const m = /^#\/(\d)/.exec(location.hash);
  const key = m && PAGES[m[1]] ? m[1] : '1';
  if (!m) history.replaceState(null, '', '#/1');
  app.querySelectorAll('.nav a').forEach((a) => a.classList.toggle('active', a.dataset.page === key));
  document.querySelectorAll('.modal-overlay').forEach((el) => el.remove());
  const page = PAGES[key];
  document.title = `COPWC · ${page.label}`;
  if (page.render) page.render(pageEl);
  else {
    pageEl.innerHTML = `
      <section class="card card-pending">
        <div class="card-head"><div><h2>${page.label} 페이지</h2><p class="sub">이 페이지는 아직 준비 중입니다.</p></div><span class="badge">개발 예정</span></div>
        <div class="pending-box">개발 예정</div>
      </section>`;
  }
  window.scrollTo(0, 0);
}

window.addEventListener('hashchange', route);
route();
