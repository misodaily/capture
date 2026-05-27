// ===== State =====
let currentJobId = null;
let pollTimer = null;

const $ = (id) => document.getElementById(id);
const els = {
  form: $('genForm'),
  pcUrl: $('pcUrl'),
  moUrl: $('moUrl'),
  searchPcUrl: $('searchPcUrl'),
  searchMoUrl: $('searchMoUrl'),
  landingUrl: $('landingUrl'),
  cardName: $('cardName'),
  submitBtn: $('submitBtn'),
  clearBtn: $('clearBtn'),
  progressPanel: $('progressPanel'),
  progressStep: $('progressStep'),
  progressMeta: $('progressMeta'),
  progressPct: $('progressPct'),
  progressFill: $('progressFill'),
  resultPanel: $('resultPanel'),
  resultTitle: $('resultTitle'),
  resultSub: $('resultSub'),
  resultMeta: $('resultMeta'),
  openBtn: $('openBtn'),
  revealBtn: $('revealBtn'),
  downloadBtn: $('downloadBtn'),
  errorPanel: $('errorPanel'),
  errorMsg: $('errorMsg'),
  retryBtn: $('retryBtn'),
  historyList: $('historyList'),
  toast: $('toast'),
  stages: document.querySelectorAll('.stage'),
};

// ===== Toast =====
let toastTimer = null;
function toast(msg, type = '') {
  els.toast.textContent = msg;
  els.toast.className = 'toast show ' + type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    els.toast.classList.remove('show');
  }, 2400);
}

// ===== Paste buttons =====
document.querySelectorAll('[data-paste]').forEach((btn) => {
  btn.addEventListener('click', async () => {
    const target = $(btn.dataset.paste);
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        target.value = text.trim();
        target.dispatchEvent(new Event('input'));
        toast('붙여넣기 완료', 'success');
      } else {
        toast('클립보드가 비어 있습니다');
      }
    } catch (e) {
      target.focus();
      toast('브라우저가 자동 붙여넣기를 거부했습니다. Ctrl+V 로 직접 붙여넣어 주세요.', 'error');
    }
  });
});

// ===== Auto-fill card name from PC URL =====
let autoFillTimer = null;
function autoExtractCardName() {
  const url = els.pcUrl.value.trim() || els.moUrl.value.trim();
  if (!url || els.cardName.value.trim()) return;
  clearTimeout(autoFillTimer);
  autoFillTimer = setTimeout(async () => {
    try {
      const r = await fetch('/api/extract-card', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url }),
      });
      const j = await r.json();
      if (j.card_name && !els.cardName.value.trim()) {
        els.cardName.value = j.card_name;
        els.cardName.placeholder = j.card_name;
      }
    } catch (e) { /* ignore */ }
  }, 250);
}
els.pcUrl.addEventListener('input', autoExtractCardName);
els.moUrl.addEventListener('input', autoExtractCardName);

// ===== Form submit =====
els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  startGeneration();
});

async function startGeneration() {
  const pc_url = els.pcUrl.value.trim();
  const mo_url = els.moUrl.value.trim();
  const search_pc_url = els.searchPcUrl ? els.searchPcUrl.value.trim() : '';
  const search_mo_url = els.searchMoUrl ? els.searchMoUrl.value.trim() : '';
  const landing_url = els.landingUrl ? els.landingUrl.value.trim() : '';
  const card_name = els.cardName.value.trim();

  if (!pc_url || !mo_url) {
    toast('PC, MO URL 을 모두 입력해 주세요.', 'error');
    return;
  }
  if (!pc_url.includes('card-search.naver.com')) {
    toast('PC URL 형식이 올바르지 않습니다 (card-search.naver.com).', 'error');
    return;
  }
  if (!mo_url.includes('m-card-search.naver.com')) {
    toast('MO URL 형식이 올바르지 않습니다 (m-card-search.naver.com).', 'error');
    return;
  }
  if (search_pc_url && !search_pc_url.includes('search.naver.com')) {
    toast('PC 검색결과 URL 형식이 올바르지 않습니다 (search.naver.com).', 'error');
    return;
  }
  if (search_mo_url && !search_mo_url.includes('m.search.naver.com')) {
    toast('MO 검색결과 URL 형식이 올바르지 않습니다 (m.search.naver.com).', 'error');
    return;
  }
  if (landing_url && !/^https?:\/\//i.test(landing_url)) {
    toast('안내 페이지 URL 은 http:// 또는 https:// 로 시작해야 합니다.', 'error');
    return;
  }

  // 패널 초기화
  els.errorPanel.hidden = true;
  els.resultPanel.hidden = true;
  els.progressPanel.hidden = false;
  setProgress(0, '요청 중...', '');
  setStage(null);
  setSubmitting(true);

  try {
    const r = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ pc_url, mo_url, search_pc_url, search_mo_url, landing_url, card_name }),
    });
    if (!r.ok) {
      const j = await r.json().catch(() => ({}));
      throw new Error(j.error || `HTTP ${r.status}`);
    }
    const j = await r.json();
    currentJobId = j.job_id;
    pollStatus();
  } catch (e) {
    showError(e.message || String(e));
    setSubmitting(false);
  }
}

function setSubmitting(busy) {
  els.submitBtn.disabled = busy;
  els.submitBtn.querySelector('span').textContent = busy ? '생성 중...' : '게재보고 만들기';
}

function setProgress(pct, step, meta) {
  els.progressPct.textContent = `${Math.round(pct)}%`;
  els.progressFill.style.width = `${pct}%`;
  if (step !== undefined) els.progressStep.textContent = step;
  if (meta !== undefined) els.progressMeta.textContent = meta || '';
}

function setStage(active) {
  // active: null | 'pc' | 'mo' | 'ppt' | 'done'
  const order = ['pc', 'mo', 'ppt'];
  els.stages.forEach((stage) => {
    const s = stage.dataset.stage;
    stage.classList.remove('active', 'done');
    if (active === 'done') {
      stage.classList.add('done');
    } else if (active && order.indexOf(s) < order.indexOf(active)) {
      stage.classList.add('done');
    } else if (active === s) {
      stage.classList.add('active');
    }
  });
}

function stageFromProgress(p) {
  if (p >= 88) return 'ppt';
  if (p >= 45) return 'mo';
  if (p >= 10) return 'pc';
  return null;
}

async function pollStatus() {
  if (!currentJobId) return;
  try {
    const r = await fetch(`/api/status/${currentJobId}`);
    if (!r.ok) throw new Error(`status HTTP ${r.status}`);
    const j = await r.json();

    setProgress(j.progress || 0, j.step || '...');
    setStage(stageFromProgress(j.progress || 0));

    if (j.status === 'done') {
      setProgress(100, '완료', '');
      setStage('done');
      showResult(j);
      setSubmitting(false);
      loadHistory();
      return;
    }
    if (j.status === 'error') {
      showError(j.error || '알 수 없는 오류');
      setSubmitting(false);
      return;
    }
    // queued / running → 계속 polling
    pollTimer = setTimeout(pollStatus, 1000);
  } catch (e) {
    showError(e.message || String(e));
    setSubmitting(false);
  }
}

function showResult(j) {
  els.progressPanel.hidden = true;
  els.resultPanel.hidden = false;
  els.errorPanel.hidden = true;
  els.resultSub.textContent = j.file_name || '';
  const meta = [];
  if (j.pc_slides !== undefined) meta.push(`PC ${j.pc_slides}장`);
  if (j.search_pc_slides) meta.push(`PC 검색 ${j.search_pc_slides}장`);
  if (j.landing_pc_slides) meta.push(`PC 안내 ${j.landing_pc_slides}장`);
  if (j.mo_slides !== undefined) meta.push(`MO ${j.mo_slides}장`);
  if (j.search_mo_slides) meta.push(`MO 검색 ${j.search_mo_slides}장`);
  if (j.landing_mo_slides) meta.push(`MO 안내 ${j.landing_mo_slides}장`);
  els.resultMeta.textContent = meta.join(' · ');
  els.downloadBtn.href = `/api/download/${currentJobId}`;
  toast('생성 완료!', 'success');
}

function showError(msg) {
  els.progressPanel.hidden = true;
  els.resultPanel.hidden = true;
  els.errorPanel.hidden = false;
  els.errorMsg.textContent = msg;
}

// ===== Result actions =====
els.openBtn.addEventListener('click', async () => {
  if (!currentJobId) return;
  els.openBtn.disabled = true;
  try {
    const r = await fetch(`/api/open/${currentJobId}`, { method: 'POST' });
    const j = await r.json();
    if (j.ok) toast('PowerPoint 에서 열었습니다', 'success');
    else toast('파일 열기 실패: ' + (j.error || ''), 'error');
  } catch (e) {
    toast('파일 열기 실패', 'error');
  } finally {
    els.openBtn.disabled = false;
  }
});
els.revealBtn.addEventListener('click', async () => {
  if (!currentJobId) return;
  try {
    await fetch(`/api/reveal/${currentJobId}`, { method: 'POST' });
    toast('탐색기를 열었습니다', 'success');
  } catch (e) {
    toast('탐색기 열기 실패', 'error');
  }
});

// ===== Retry =====
els.retryBtn.addEventListener('click', () => {
  els.errorPanel.hidden = true;
  startGeneration();
});

// ===== Clear =====
els.clearBtn.addEventListener('click', () => {
  els.pcUrl.value = '';
  els.moUrl.value = '';
  if (els.searchPcUrl) els.searchPcUrl.value = '';
  if (els.searchMoUrl) els.searchMoUrl.value = '';
  if (els.landingUrl) els.landingUrl.value = '';
  els.cardName.value = '';
  els.cardName.placeholder = '예: 신한카드 Deep Once';
  els.progressPanel.hidden = true;
  els.resultPanel.hidden = true;
  els.errorPanel.hidden = true;
  els.pcUrl.focus();
});

// ===== History =====
async function loadHistory() {
  try {
    const r = await fetch('/api/history');
    const j = await r.json();
    renderHistory(j.items || []);
  } catch (e) {
    /* ignore */
  }
}

function renderHistory(items) {
  if (!items.length) {
    els.historyList.innerHTML = '<div class="history-empty">아직 생성된 보고서가 없습니다.</div>';
    return;
  }
  els.historyList.innerHTML = items.map((it, idx) => {
    const dt = it.finished_at ? formatTime(it.finished_at) : '';
    const card = escapeHtml(it.card_name || '신한카드');
    const fname = escapeHtml(it.file_name || '');
    return `
      <div class="history-item">
        <div class="h-icon">${idx + 1}</div>
        <div class="h-body">
          <div class="h-card" title="${fname}">${card}</div>
          <div class="h-meta">${dt}</div>
        </div>
        <div class="h-act">
          <button data-open="${encodeURIComponent(it.file_path)}" title="PowerPoint 로 열기">열기</button>
          <button data-reveal="${encodeURIComponent(it.file_path)}" title="폴더에서 보기">폴더</button>
        </div>
      </div>
    `;
  }).join('');

  els.historyList.querySelectorAll('[data-open]').forEach((b) => {
    b.addEventListener('click', () => openPath(decodeURIComponent(b.dataset.open)));
  });
  els.historyList.querySelectorAll('[data-reveal]').forEach((b) => {
    b.addEventListener('click', () => revealPath(decodeURIComponent(b.dataset.reveal)));
  });
}

async function openPath(path) {
  try {
    const r = await fetch('/api/open-path', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path }),
    });
    const j = await r.json();
    if (j.ok) toast('PowerPoint 에서 열었습니다', 'success');
    else toast('열기 실패', 'error');
  } catch (e) { toast('열기 실패', 'error'); }
}
async function revealPath(path) {
  try {
    await fetch('/api/reveal-path', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path }),
    });
    toast('탐색기를 열었습니다', 'success');
  } catch (e) { toast('탐색기 열기 실패', 'error'); }
}

function formatTime(iso) {
  try {
    const d = new Date(iso);
    const y = d.getFullYear();
    const m = (d.getMonth() + 1).toString().padStart(2, '0');
    const dd = d.getDate().toString().padStart(2, '0');
    const h = d.getHours().toString().padStart(2, '0');
    const mi = d.getMinutes().toString().padStart(2, '0');
    return `${y}.${m}.${dd} ${h}:${mi}`;
  } catch (e) {
    return iso;
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;',
  }[c]));
}

// ===== Init =====
loadHistory();
els.pcUrl.focus();
