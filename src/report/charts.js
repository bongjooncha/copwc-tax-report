// 그래프 4종 (Q-02) → PNG dataURL. xlsx 이미지와 PDF 양쪽에 같은 이미지를 씀
import {
  Chart, BarController, BarElement, PieController, ArcElement, CategoryScale, LinearScale, Legend, Title, Tooltip,
} from 'chart.js';
import { isNum } from './format.js';

Chart.register(BarController, BarElement, PieController, ArcElement, CategoryScale, LinearScale, Legend, Title, Tooltip);

export const CHART_SIZE = { w: 586, h: 310 }; // 15.5cm × 8.2cm (엑셀 그래프와 같은 크기)
const FONT = "'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo','Noto Sans KR',sans-serif";
const C_PREV = '#A6A6A6';
const C_CUR = '#1F2D40';
const PIE = ['#1F2D40', '#3E5C76', '#748CAB', '#A9B9CF', '#D4C9B5', '#BFBFBF'];

const whiteBg = {
  id: 'whiteBg',
  beforeDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, chart.width, chart.height);
    ctx.restore();
  },
};

// 파이 조각 바깥에 비율(%) 표시 (엑셀 그래프의 데이터 레이블과 같게)
const piePercent = {
  id: 'piePercent',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'pie') return;
    const { ctx } = chart;
    const meta = chart.getDatasetMeta(0);
    const data = chart.data.datasets[0].data;
    const total = data.reduce((s, v) => s + (isNum(v) ? v : 0), 0);
    ctx.save();
    ctx.font = `10px ${FONT}`;
    ctx.fillStyle = '#333';
    meta.data.forEach((arc, i) => {
      if (!total || !isNum(data[i])) return;
      const mid = (arc.startAngle + arc.endAngle) / 2;
      const r = arc.outerRadius + 14;
      const x = arc.x + Math.cos(mid) * r;
      const y = arc.y + Math.sin(mid) * r;
      ctx.textAlign = Math.cos(mid) >= 0 ? 'left' : 'right';
      ctx.textBaseline = 'middle';
      ctx.fillText(`${Math.round((data[i] / total) * 100)}%`, x, y);
    });
    ctx.restore();
  },
};

// 막대 위 값 표시 (options.showValues) — 원 단위 전체 금액(천 단위 쉼표). 막대보다 글자가 넓으면 세로로 씀
// (이전에는 '4.82억'·'3,857만'처럼 줄여 표시할 수 있었으나, 판독·채점을 위해 원 단위로 통일)
const barValues = {
  id: 'barValues',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'bar' || !chart.config.options?.showValues) return;
    const { ctx } = chart;
    ctx.save();
    ctx.font = `9px ${FONT}`;
    ctx.fillStyle = '#333';
    chart.data.datasets.forEach((ds, i) => {
      chart.getDatasetMeta(i).data.forEach((bar, j) => {
        const v = ds.data[j];
        if (!isNum(v)) return;
        const label = wonText(v);
        const up = v >= 0;
        if (ctx.measureText(label).width <= bar.width + 2) {
          ctx.textAlign = 'center';
          ctx.textBaseline = up ? 'bottom' : 'top';
          ctx.fillText(label, bar.x, up ? bar.y - 2 : bar.y + 2);
        } else {
          ctx.save();
          ctx.translate(bar.x, up ? bar.y - 3 : bar.y + 3);
          ctx.rotate(-Math.PI / 2);
          ctx.textAlign = up ? 'left' : 'right';
          ctx.textBaseline = 'middle';
          ctx.fillText(label, 0, 0);
          ctx.restore();
        }
      });
    });
    ctx.restore();
  },
};

const wonText = (v) => `${Math.round(v).toLocaleString('ko-KR')}`;
const yearLabel = (y, kind) => (y ? `${y} 사업연도(${kind})` : kind);

function render(config) {
  const canvas = document.createElement('canvas');
  canvas.width = CHART_SIZE.w * 2;
  canvas.height = CHART_SIZE.h * 2;
  canvas.style.width = `${CHART_SIZE.w}px`;
  canvas.style.height = `${CHART_SIZE.h}px`;
  const holder = document.createElement('div');
  holder.style.cssText = `position:fixed;left:-10000px;top:0;width:${CHART_SIZE.w}px;height:${CHART_SIZE.h}px`;
  holder.appendChild(canvas);
  document.body.appendChild(holder);
  const chart = new Chart(canvas, {
    ...config,
    plugins: [whiteBg, barValues, piePercent],
    options: {
      responsive: false,
      maintainAspectRatio: false,
      animation: false,
      devicePixelRatio: 2,
      layout: { padding: { top: 6, right: 12, bottom: 4, left: 8 } },
      ...config.options,
      plugins: {
        legend: { position: 'bottom', labels: { font: { family: FONT, size: 11 }, boxWidth: 10 } },
        title: { display: true, text: config.title, font: { family: FONT, size: 13, weight: 'bold' }, color: '#000' },
        tooltip: { enabled: false },
        ...(config.options?.plugins || {}),
      },
    },
  });
  chart.resize(CHART_SIZE.w, CHART_SIZE.h);
  chart.draw();
  const url = canvas.toDataURL('image/png');
  chart.destroy();
  holder.remove();
  return url;
}

// 범례에 사업연도, 제목에 단위(원)를 적고 막대 위에 원 단위 금액을 표시함 (판독·채점용)
function barConfig(title, labels, prev, cur, yPrev, yCur) {
  const datasets = [];
  if (prev) datasets.push({ label: yearLabel(yPrev, '전기'), data: prev, backgroundColor: C_PREV, maxBarThickness: 46 });
  datasets.push({ label: yearLabel(yCur, '당기'), data: cur, backgroundColor: C_CUR, maxBarThickness: 46 });
  return {
    type: 'bar',
    title: `${title} (단위: 원)`,
    data: { labels, datasets },
    options: {
      showValues: true,
      scales: {
        x: { ticks: { font: { family: FONT, size: 11 } }, grid: { display: false } },
        y: { grace: '35%', ticks: { font: { family: FONT, size: 10 }, callback: (v) => Number(v).toLocaleString('ko-KR') }, grid: { display: false }, border: { display: true } },
      },
    },
  };
}

export function renderCharts(c) {
  if (!c.has('당기')) return {};
  const yCur = c.years.당기;
  const yPrev = c.years.전기;
  const hp = c.hasPrev; // 전기 데이터가 없으면 당기만 (Q-02)
  const out = {};

  const plKeys = [['매출액', '매출액'], ['매출총이익', '매출총손익'], ['영업이익', '영업손익'], ['당기순이익', '당기순손익']];
  out.pl = render(
    barConfig(
      `주요손익 비교 (${yearLabel(yPrev, '전기')} vs ${yearLabel(yCur, '당기')})`,
      plKeys.map(([l]) => l),
      hp ? plKeys.map(([, k]) => c.is(k, '전기')) : null,
      plKeys.map(([, k]) => c.is(k, '당기')),
      yPrev, yCur,
    ),
  );

  const pieRows = c.sgaRows.filter((r) => r.cur > 0);
  const pieTotal = pieRows.reduce((s, r) => s + r.cur, 0);
  out.sgaPie = render({
    type: 'pie',
    title: `${yearLabel(yCur, '당기')} 판관비 구성비율 (단위: %)`,
    data: {
      labels: pieRows.map((r) => r.name),
      datasets: [{ data: pieRows.map((r) => r.cur), backgroundColor: PIE, borderColor: '#fff', borderWidth: 1 }],
    },
    options: { layout: { padding: { top: 6, right: 12, bottom: 20, left: 30 } }, plugins: { legend: { position: 'right', labels: { font: { family: FONT, size: 11 }, boxWidth: 10 } } } },
  });

  const top = c.sgaRows.filter((r) => r.name !== '기타');
  out.sgaBar = render(
    barConfig(`판관비 주요 5개 항목 (${yearLabel(yPrev, '전기')} vs ${yearLabel(yCur, '당기')})`, top.map((r) => r.name), hp ? top.map((r) => r.prev) : null, top.map((r) => r.cur), yPrev, yCur),
  );

  const bsKeys = [['자산총계', '자산총계'], ['부채총계', '부채총계'], ['자본총계', '자본총계']];
  out.bs = render(
    barConfig(
      `재무상태 비교 (${yearLabel(yPrev, '전기')} vs ${yearLabel(yCur, '당기')})`,
      bsKeys.map(([l]) => l),
      hp ? bsKeys.map(([, k]) => c.bs(k, '전기')) : null,
      bsKeys.map(([, k]) => c.bs(k, '당기')),
      yPrev, yCur,
    ),
  );
  return out;
}

// 그래프 내용(제목·범례·항목·원 단위 금액)을 글자로 — PDF 투명 글자 층과 img alt에 씀.
// 그래프는 그림이라 PDF에서 글자로 읽히지 않으므로, 같은 내용을 그래프 자리에 글자로 겹쳐 넣음.
export function chartTexts(c) {
  if (!c.has('당기')) return {};
  const hp = c.hasPrev;
  const P = yearLabel(c.years.전기, '전기');
  const C = yearLabel(c.years.당기, '당기');
  const w = (v) => (isNum(v) ? `${wonText(v)}원` : '확인 필요');
  const pair = (rows) => rows.map(([name, pv, cv]) => (hp ? `${name}: ${P} ${w(pv)} / ${C} ${w(cv)}` : `${name}: ${C} ${w(cv)}`));
  const pieRows = c.sgaRows.filter((r) => r.cur > 0);
  const tot = pieRows.reduce((s, r) => s + r.cur, 0);
  const out = {
    pl: [`[그래프] 주요손익 비교 (${P} vs ${C}) (단위: 원)`, ...pair(
      [['매출액', '매출액'], ['매출총이익', '매출총손익'], ['영업이익', '영업손익'], ['당기순이익', '당기순손익']]
        .map(([l, k]) => [l, hp ? c.is(k, '전기') : null, c.is(k, '당기')]),
    )],
    sgaPie: [`[그래프] ${C} 판관비 구성비율 (단위: 원, %)`,
      ...pieRows.map((r) => `${r.name}: ${w(r.cur)} (${tot ? ((r.cur / tot) * 100).toFixed(2) : '0.00'}%)`)],
    sgaBar: [`[그래프] 판관비 주요 5개 항목 (${P} vs ${C}) (단위: 원)`,
      ...pair(c.sgaRows.filter((r) => r.name !== '기타').map((r) => [r.name, r.prev, r.cur]))],
    bs: [`[그래프] 재무상태 비교 (${P} vs ${C}) (단위: 원)`, ...pair(
      ['자산총계', '부채총계', '자본총계'].map((k) => [k, hp ? c.bs(k, '전기') : null, c.bs(k, '당기')]),
    )],
  };
  return Object.fromEntries(Object.entries(out).map(([k, v]) => [k, v.join('\n')]));
}
