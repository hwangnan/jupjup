'use strict';
/* 줍줍 — front end. Demo mode until a Google Sheets (Apps Script) API is connected in Settings. */

const $ = s => document.querySelector(s), $$ = s => document.querySelectorAll(s);
const DAY = 864e5;
const CATEGORIES = ['의류', '패션잡화', '폰케이스', '생활용품'];
const SOURCES = ['판매완료 캡처', '내 판매기록', '직접 등록'];
const HUE = { '의류': 215, '패션잡화': 28, '폰케이스': 280, '생활용품': 160 };
const GLYPH = {
  '의류': '<path d="M8 3L3 6l2 4 2-1v12h10V9l2 1 2-4-5-3c-.6 1.6-2.1 2.6-4 2.6S8.6 4.6 8 3z"/>',
  '패션잡화': '<path d="M5 8h14l-1.2 12H6.2L5 8zM9 8V6a3 3 0 016 0v2"/>',
  '폰케이스': '<rect x="6.5" y="2.5" width="11" height="19" rx="2.5"/><circle cx="10" cy="6.5" r="1.4"/>',
  '생활용품': '<path d="M7 3h10l-1.3 18H8.3L7 3zM7.4 8h9.2"/>'
};
const IC = {
  bolt: '<path d="M13 2L4 14h7l-1 8 9-12h-7z"/>',
  clock: '<path d="M12 8v5l3 2M12 22a10 10 0 110-20 10 10 0 010 20z"/>',
  check: '<path d="M20 6L9 17l-5-5"/>',
  trend: '<path d="M3 17l6-6 4 4 8-8M15 7h6v6"/>',
  camera: '<rect x="3" y="5" width="18" height="15" rx="3"/><circle cx="12" cy="12.5" r="3.5"/><path d="M8 5l1.5-2h5L16 5"/>',
  link: '<path d="M10 14a5 5 0 007 0l3-3a5 5 0 00-7-7l-1 1M14 10a5 5 0 00-7 0l-3 3a5 5 0 007 7l1-1"/>',
  pen: '<path d="M4 20h4L19 9l-4-4L4 16v4zM14 6l4 4"/>',
  trash: '<path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/>',
  more: '<circle cx="5" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="19" cy="12" r="1.6"/>',
  sheet: '<rect x="3" y="3" width="18" height="18" rx="3"/><path d="M3 9h18M3 15h18M9 3v18"/>',
  phone: '<rect x="6" y="2" width="12" height="20" rx="3"/><path d="M11 18h2"/>',
  spark: '<path d="M12 3l2.5 5.5L20 11l-5.5 2.5L12 19l-2.5-5.5L4 11l5.5-2.5z"/>'
};
const svg = (p, w = 2, extra = '') => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" ${extra}>${p}</svg>`;
const DEFAULT_SET = {
  target: 10000, move: 2000, clean: 1500, feePct: 0, minutes: 40,
  keywords: [{ kw: '케이스티파이', max: 20000 }, { kw: '폴로 니트', max: 30000 }, { kw: '마르헨제이', max: 25000 }, { kw: '스탠리', max: 20000 }, { kw: '칼하트', max: 30000 }]
};

/* ---------- utils ---------- */
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
  del(k) { try { localStorage.removeItem(k); } catch (e) {} }
};
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const won = n => Math.round(n || 0).toLocaleString('ko-KR') + '원';
const man = n => (Math.round((n || 0) / 1000) / 10).toLocaleString('ko-KR') + '만';
const num = v => +String(v ?? '').replace(/[^\d]/g, '') || 0;
const iso = (d = Date.now()) => new Date(d).toISOString();
const med = a => { if (!a.length) return 0; const s = a.slice().sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const daysSince = t => t ? Math.max(0, Math.floor((Date.now() - new Date(t)) / DAY)) : 0;
const daysBetween = (a, b) => a && b ? Math.max(0, Math.round((new Date(b) - new Date(a)) / DAY)) : 0;
const hasJong = w => { const c = String(w).charCodeAt(String(w).length - 1); return c >= 0xAC00 && c <= 0xD7A3 && (c - 0xAC00) % 28 > 0; };
const iga = w => w + (hasJong(w) ? '이' : '가');
const eunneun = w => w + (hasJong(w) ? '은' : '는');
function rel(t) {
  if (!t) return '';
  const m = (Date.now() - new Date(t)) / 6e4;
  if (m < 1) return '방금'; if (m < 60) return Math.floor(m) + '분 전';
  if (m < 1440) return Math.floor(m / 60) + '시간 전'; return Math.floor(m / 1440) + '일 전';
}

/* ---------- state ---------- */
const S = { mode: 'demo', tab: 'home', filter: '전체', invTab: 'watch', set: { ...DEFAULT_SET }, items: [], comps: [], syncedAt: 0 };
const find = id => S.items.find(i => i.id === id);
const compsOf = id => S.comps.filter(c => c.itemId === id);
const costOf = sell => S.set.move + S.set.clean + sell * S.set.feePct / 100;
const soldNet = it => it.soldPrice - it.buyPrice - costOf(it.soldPrice);
function calc(it, buy = it.buyPrice) {
  const cs = compsOf(it.id).map(c => c.price), n = cs.length, sell = med(cs), cost = costOf(sell), net = sell - buy - cost;
  const maxBuy = Math.floor((sell - cost - S.set.target) / 1000) * 1000;
  let v;
  if (n < 2) v = { k: 'unk', l: '시세 확인 필요' };
  else if (net >= S.set.target) v = { k: 'good', l: '사세요' };
  else if (net >= S.set.target * .5) v = { k: 'mid', l: '네고하면 OK' };
  else v = { k: 'bad', l: '패스 추천' };
  return { n, sell, cost, net, maxBuy, rate: buy ? net / buy : 0, hour: net / (S.set.minutes / 60), v, conf: Math.min(n, 3) };
}
const expectedNet = it => it.status === 'selling' && it.listPrice ? it.listPrice - it.buyPrice - costOf(it.listPrice) : compsOf(it.id).length ? calc(it).net : 0;
const imgSrc = it => it.img || (it.imageId ? `https://drive.google.com/thumbnail?id=${encodeURIComponent(it.imageId)}&sz=w800` : '');
const thumb = (it, cls = 'thumb') => { const src = imgSrc(it); return `<div class="${cls}" style="--h:${HUE[it.category] ?? 215}">${svg(GLYPH[it.category] || GLYPH['의류'], 1.7)}${src ? `<img src="${esc(src)}" alt="" loading="lazy" onerror="this.remove()">` : ''}</div>`; };

/* ---------- demo data ---------- */
function rng(a) { return () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
function demoData() {
  const t = Date.now(), mAgo = m => iso(t - m * 6e4), dAgo = d => iso(t - d * DAY);
  const items = [], comps = []; let cid = 0;
  const C = (itemId, price, source, d, cond) => comps.push({ id: 'dc' + (++cid), itemId, price, source, date: dAgo(d), cond });
  const add = o => items.push({ source: '키워드 알림', place: '', cond: '', url: '', retail: null, askPrice: o.buyPrice, ...o });
  add({ id: 'd1', status: 'new', title: '폴로 랄프로렌 케이블 니트 L 네이비', category: '의류', buyPrice: 28000, retail: 259000, cond: '상태 좋음 · 보풀 거의 없음', place: '망원동', createdAt: mAgo(3) });
  C('d1', 58000, '판매완료 캡처', 8, '상태 좋음'); C('d1', 62000, '내 판매기록', 24, '새상품급'); C('d1', 55000, '판매완료 캡처', 20, '사용감 조금');
  add({ id: 'd2', status: 'new', title: '케이스티파이 아이폰15 Pro 미러 케이스', category: '폰케이스', buyPrice: 15000, retail: 82000, cond: '2주 사용 · 박스 있음', place: '합정동', createdAt: mAgo(12) });
  C('d2', 32000, '내 판매기록', 14, '상태 좋음'); C('d2', 35000, '판매완료 캡처', 11, '박스 있음'); C('d2', 30000, '판매완료 캡처', 23, '사용감 조금');
  add({ id: 'd3', status: 'new', title: '마르헨제이 미니 크로스백 아이보리', category: '패션잡화', buyPrice: 19000, retail: 79000, cond: '모서리 살짝 오염', place: '연남동', createdAt: mAgo(25) });
  C('d3', 36000, '판매완료 캡처', 17, '상태 좋음'); C('d3', 33000, '직접 등록', 10, '오염 조금');
  add({ id: 'd4', status: 'new', title: '스탠리 퀜처 텀블러 887ml 크림', category: '생활용품', buyPrice: 22000, retail: 55000, cond: '미사용 · 선물받음', place: '서교동', createdAt: mAgo(41) });
  C('d4', 32000, '판매완료 캡처', 12, '미사용'); C('d4', 34000, '직접 등록', 9, '미사용'); C('d4', 30000, '판매완료 캡처', 21, '사용감 조금');
  add({ id: 'd5', status: 'new', title: '무신사 스탠다드 싱글 블레이저 M', category: '의류', buyPrice: 25000, retail: 79900, cond: '2회 착용', place: '상암동', createdAt: mAgo(62) });
  C('d5', 39000, '직접 등록', 32, '상태 좋음');
  add({ id: 'd6', status: 'new', title: '디스이즈네버댓 로고 후드 L 그레이', category: '의류', buyPrice: 30000, retail: 89000, cond: '목 부분 늘어남', place: '성산동', createdAt: mAgo(130) });
  C('d6', 33000, '판매완료 캡처', 13, '사용감 있음'); C('d6', 35000, '내 판매기록', 35, '상태 좋음');
  add({ id: 'd7', status: 'bought', title: '나이키 ACG 폴라텍 플리스 M', category: '의류', buyPrice: 35000, askPrice: 40000, retail: 139000, createdAt: dAgo(4), boughtAt: dAgo(3) });
  C('d7', 68000, '내 판매기록', 41, '상태 좋음'); C('d7', 72000, '판매완료 캡처', 19, '새상품급');
  add({ id: 'd8', status: 'selling', title: '커버낫 바시티 자켓 L', category: '의류', buyPrice: 32000, retail: 169000, listPrice: 69000, createdAt: dAgo(21), boughtAt: dAgo(20), listedAt: dAgo(16) });
  C('d8', 62000, '판매완료 캡처', 28, '상태 좋음'); C('d8', 65000, '판매완료 캡처', 16, '상태 좋음');
  add({ id: 'd9', status: 'selling', title: '닥터마틴 카드지갑 블랙', category: '패션잡화', buyPrice: 12000, retail: 69000, listPrice: 35000, createdAt: dAgo(9), boughtAt: dAgo(8), listedAt: dAgo(5) });
  C('d9', 33000, '판매완료 캡처', 22, '상태 좋음'); C('d9', 30000, '직접 등록', 12, '사용감');

  const R = rng(11);
  const names = { '의류': ['칼하트 디트로이트 자켓', '폴로 옥스포드 셔츠', '나이키 테크플리스 조거', '파타고니아 신칠라 풀오버', '노스페이스 눕시 베스트'], '패션잡화': ['마뗑킴 레더 숄더백', '칼하트 WIP 비니', '닥터마틴 카드지갑', '마르헨제이 토트백'], '폰케이스': ['케이스티파이 아이폰 케이스', '케이스티파이 갤럭시 S24 케이스', '디자인스킨 카드 케이스'], '생활용품': ['스탠리 텀블러', '르크루제 머그', '발뮤다 미니 램프'] };
  const hold = { '의류': 11, '패션잡화': 7, '폰케이스': 4, '생활용품': 9 }, mult = { '의류': 1.9, '패션잡화': 1.65, '폰케이스': 2.15, '생활용품': 1.4 };
  for (let k = 0; k < 46; k++) {
    const r = R(), cat = r < .4 ? '의류' : r < .65 ? '패션잡화' : r < .9 ? '폰케이스' : '생활용품';
    const soldAgo = R() * 172 + .2, slow = R() < .16, h = Math.max(1, Math.round(hold[cat] * (.5 + R()) + (slow ? 14 : 0)));
    const buy = Math.round((9000 + R() * 26000) / 1000) * 1000;
    const sold = Math.round(buy * mult[cat] * (slow ? .78 : .88 + R() * .28) / 1000) * 1000;
    add({ id: 'h' + k, status: 'sold', title: names[cat][Math.floor(R() * names[cat].length)], category: cat, buyPrice: buy, retail: Math.round(buy * (3 + R() * 3) / 1000) * 1000, soldPrice: sold, listPrice: sold, createdAt: dAgo(soldAgo + h + 1), boughtAt: dAgo(soldAgo + h), listedAt: dAgo(soldAgo + h - .5), soldAt: dAgo(soldAgo) });
  }
  return { items, comps };
}

/* ---------- API ---------- */
const API = {
  cfg() { return store.get('jj-conn', null); },
  async call(action, payload = {}, cfg = this.cfg()) {
    if (!cfg) throw new Error('연결 정보가 없어요');
    const r = await fetch(cfg.url, { method: 'POST', body: JSON.stringify({ action, token: cfg.token, ...payload }) });
    if (!r.ok) throw new Error('서버 응답 ' + r.status);
    let j; try { j = await r.json(); } catch (e) { throw new Error('응답을 읽지 못했어요. 웹 앱 주소와 배포 권한을 확인하세요'); }
    if (!j.ok) throw new Error(j.error || '요청 실패');
    return j.data;
  }
};
const live = () => S.mode === 'live';
function saveCache() { if (live()) store.set('jj-cache', { items: S.items, comps: S.comps, set: S.set, at: S.syncedAt }); }
function applyRemote(d) { S.items = d.items || []; S.comps = d.comps || []; S.set = { ...DEFAULT_SET, ...(d.settings || {}) }; S.syncedAt = Date.now(); saveCache(); }
async function refresh(manual) {
  try { applyRemote(await API.call('bootstrap')); renderAll(); if (manual) toast('최신 상태예요'); }
  catch (e) { toast('불러오지 못했어요: ' + e.message); }
}
async function createItem(data) {
  let it;
  if (live()) it = await API.call('createItem', { item: data });
  else it = { ...data, id: 'l' + Date.now().toString(36), createdAt: iso() };
  S.items.push(it); saveCache(); renderAll(); return it;
}
async function updateItem(id, patch) {
  const it = find(id); if (!it) return false;
  const before = { ...it }; Object.assign(it, patch); renderAll();
  if (!live()) return true;
  try { Object.assign(it, await API.call('updateItem', { id, patch })); saveCache(); return true; }
  catch (e) { for (const k in it) delete it[k]; Object.assign(it, before); renderAll(); toast('저장 못 했어요: ' + e.message); return false; }
}
async function deleteItem(id) {
  const snap = { items: S.items, comps: S.comps };
  S.items = S.items.filter(i => i.id !== id); S.comps = S.comps.filter(c => c.itemId !== id); renderAll();
  if (!live()) return true;
  try { await API.call('deleteItem', { id }); saveCache(); return true; }
  catch (e) { Object.assign(S, snap); renderAll(); toast('삭제 못 했어요: ' + e.message); return false; }
}
async function addComp(comp) {
  const c = live() ? await API.call('addComp', { comp }) : { ...comp, id: 'lc' + Date.now().toString(36) };
  S.comps.push(c); saveCache(); return c;
}
async function deleteComp(id) {
  const snap = S.comps; S.comps = S.comps.filter(c => c.id !== id);
  if (!live()) return true;
  try { await API.call('deleteComp', { id }); saveCache(); return true; }
  catch (e) { S.comps = snap; toast('삭제 못 했어요: ' + e.message); return false; }
}
function saveSettings(patch) {
  Object.assign(S.set, patch); saveCache();
  if (live()) API.call('saveSettings', { settings: patch }).catch(e => toast('설정 저장 실패: ' + e.message));
}
const STAMP = { bought: 'boughtAt', selling: 'listedAt', sold: 'soldAt' };
const STATUS_MSG = { new: '관심에서 뺐어요', watch: '관심 목록에 담았어요', pass: '패스했어요', bought: '매입 완료로 옮겼어요', selling: '판매중으로 옮겼어요', sold: '판매 완료! 수익에 반영했어요' };
function setStatus(id, st, extra = {}) {
  const it = find(id), patch = { status: st, ...extra }, k = STAMP[st];
  if (k && !it[k] && !patch[k]) patch[k] = iso();
  toast(STATUS_MSG[st]);
  return updateItem(id, patch);
}

/* ---------- HOME ---------- */
function renderHome() {
  const list = S.items.filter(i => i.status === 'new' && (S.filter === '전체' || i.category === S.filter)).map(i => ({ i, c: calc(i) }))
    .sort((a, b) => (b.c.n >= 2) - (a.c.n >= 2) || b.c.net - a.c.net || new Date(b.i.createdAt) - new Date(a.i.createdAt));
  const fresh = S.items.filter(i => i.status === 'new'), good = fresh.filter(i => calc(i).v.k === 'good').length;
  $('#home-title').innerHTML = good ? `오늘 살 만한 매물<br><em>${good}개</em>를 찾았어요` : fresh.length ? `새 매물 ${fresh.length}개를<br>살펴보세요` : '새 후보를<br>기다리는 중이에요';
  $('#home-sub').textContent = live() ? `직접 등록하거나 알림으로 들어온 매물이에요${S.syncedAt ? ' · ' + rel(S.syncedAt) + ' 동기화' : ''}` : '당근 키워드 알림으로 들어온 매물을 시세와 비교했어요';
  $('#filters').innerHTML = ['전체', ...CATEGORIES].map(f => `<button class="chip${S.filter === f ? ' on' : ''}" data-filter="${f}">${f}</button>`).join('');
  $('#swipe-hint').hidden = !list.length;
  $('#deals').innerHTML = list.length ? list.map(({ i, c }) => `
   <div class="deal-wrap" data-wrap="${esc(i.id)}"><div class="deal" data-id="${esc(i.id)}">
    ${thumb(i)}
    <div class="d-body">
      <div class="d-meta">${(Date.now() - new Date(i.createdAt)) < 36e5 ? '<span class="live"></span>' : ''}${esc(i.source || '직접 등록')}${i.place ? ' · ' + esc(i.place) : ''} · ${rel(i.createdAt)}</div>
      <div class="d-title">${esc(i.title)}</div>
      <div class="d-price num"><b>${won(i.buyPrice)}</b>${c.n ? ` → 시세 ${man(c.sell)}` : ''}</div>
      <div class="d-foot">
        <span class="profit num ${c.v.k === 'unk' ? 'unk' : c.net < 0 ? 'neg' : c.v.k === 'mid' ? 'mid' : ''}">${c.v.k === 'unk' ? '시세 근거 부족' : (c.net >= 0 ? '+' : '') + won(c.net)}</span>
        <span class="pill ${c.v.k}">${c.v.l}</span>
      </div>
    </div></div></div>`).join('')
    : `<div class="empty">${fresh.length ? '이 카테고리에는 새 매물이 없어요.' : '아직 살펴볼 매물이 없어요.<br>+ 버튼으로 매물을 등록해 보세요.'}${fresh.length ? '' : '<button class="btn pri" data-act="add">매물 등록하기</button>'}</div>`;
  $$('.deal').forEach(bindSwipe);
}
function bindSwipe(el) {
  let x0, y0, dx = 0, on = false, lock = null; const wrap = el.parentElement, id = el.dataset.id;
  el.addEventListener('pointerdown', e => { x0 = e.clientX; y0 = e.clientY; dx = 0; on = true; lock = null; el.style.transition = 'none'; });
  el.addEventListener('pointermove', e => {
    if (!on) return; const ddx = e.clientX - x0, ddy = e.clientY - y0;
    if (!lock && (Math.abs(ddx) > 8 || Math.abs(ddy) > 8)) lock = Math.abs(ddx) > Math.abs(ddy) ? 'x' : 'y';
    if (lock !== 'x') return; dx = ddx; el.style.transform = `translateX(${dx}px) rotate(${dx / 50}deg)`; wrap.dataset.hint = dx > 50 ? 'watch' : dx < -50 ? 'pass' : '';
  });
  const end = () => {
    if (!on) return; on = false; el.style.transition = '';
    if (lock === 'x' && Math.abs(dx) > 100) return swipeOut(id, dx > 0 ? 1 : -1);
    el.style.transform = ''; wrap.dataset.hint = ''; if (!lock) openDetail(id);
  };
  el.addEventListener('pointerup', end); el.addEventListener('pointercancel', () => { lock = 'y'; end(); });
}
function swipeOut(id, dir) {
  const w = document.querySelector(`[data-wrap="${CSS.escape(id)}"]`), d = w.querySelector('.deal');
  d.style.transform = `translateX(${dir * 130}%) rotate(${dir * 8}deg)`;
  setTimeout(() => w.classList.add('gone'), 180);
  setTimeout(() => setStatus(id, dir > 0 ? 'watch' : 'pass'), 520);
}

/* ---------- DETAIL ---------- */
let cur = null;
function openDetail(id) {
  const it = find(id); if (!it) return; cur = it; const p = $('#detail');
  p.innerHTML = detailHTML(it); p.scrollTop = 0; p.classList.remove('scrolled');
  requestAnimationFrame(() => { p.classList.add('on'); $('#cta').classList.add('on'); });
  renderCTA(it); bindSlider(it);
}
function refreshDetail() {
  if (!cur) return; const it = find(cur.id); if (!it) return closeDetail();
  cur = it; const p = $('#detail'), y = p.scrollTop; p.innerHTML = detailHTML(it); p.scrollTop = y; renderCTA(it); bindSlider(it);
}
function closeDetail() { $('#detail').classList.remove('on'); $('#cta').classList.remove('on'); cur = null; }
function detailHTML(it) {
  const c = calc(it), ok = c.n >= 2, cs = compsOf(it.id).sort((a, b) => new Date(b.date) - new Date(a.date));
  const reason = {
    good: `시세보다 ${won(c.sell - it.buyPrice)} 싸게 올라왔어요. 비용을 빼도 ${won(c.net)}이 남아요.`,
    mid: `남긴 하지만 기준(${won(S.set.target)})보다 적어요. ${won(Math.max(c.maxBuy, 0))}까지 깎으면 기준을 넘어요.`,
    bad: '시세와 차이가 작아서 비용을 빼면 거의 남지 않아요.',
    unk: `판매완료 가격이 ${c.n}건뿐이라 판단하기 어려워요. 정가만 보고 사면 위험해요.`
  }[c.v.k];
  const offer = Math.max(Math.floor(Math.min(it.buyPrice * .85, c.maxBuy > 0 ? c.maxBuy : it.buyPrice * .85) / 1000) * 1000, 1000);
  const col = { good: 'brand', mid: 'warn', bad: 'ink3' }[c.v.k];
  return `
  <div class="page-top pad"><button class="icon-btn" data-act="back" aria-label="뒤로">${svg('<path d="M15 5l-7 7 7 7"/>', 2.2, 'width="24" height="24"')}</button><div class="t">${esc(it.title)}</div><button class="icon-btn" data-act="item-menu" aria-label="더 보기">${svg(IC.more, 2.4, 'width="22" height="22"')}</button></div>
  <div class="pad">
    ${thumb(it, 'thumb hero-img')}
    <div class="d-meta" style="margin-top:16px">${esc(it.source || '직접 등록')} · ${esc(it.place || '동네 미입력')} · ${rel(it.createdAt)}</div>
    <div style="font-size:18px;font-weight:700;margin-top:4px;letter-spacing:-.02em">${esc(it.title)}</div>
    <div style="color:var(--ink3);font-size:14px;margin-top:2px">${esc(it.cond || '')}</div>
    ${it.url ? `<a class="ghost-btn" href="${esc(it.url)}" target="_blank" rel="noopener">원본 매물 보기 ›</a>` : ''}

    <div class="verdict"><span class="pill ${c.v.k}">${c.v.l}</span>
      <h2>${c.v.k === 'unk' ? '시세를 먼저 확인하세요' : c.net >= 0 ? `사면 <span style="color:var(--${col})">${won(c.net)}</span> 남아요` : `사면 ${won(-c.net)} 손해예요`}</h2>
      <p>${reason}</p></div>

    ${ok ? `<div class="card"><h3>얼마까지 사도 될까 <small>매입가를 움직여 보세요</small></h3>
      <div class="slider-val num" id="sv"></div><div class="slider-cap" id="sc"></div>
      <input type="range" id="buy" min="${Math.max(1000, Math.round(c.sell * .2 / 1000) * 1000)}" max="${Math.max(Math.round(c.sell / 1000) * 1000, it.buyPrice)}" step="1000" value="${it.buyPrice}" aria-label="매입가">
      <div class="scale-marks" id="marks"></div>
      <div id="calc"></div></div>` : ''}

    <div class="card"><h3>시세 근거 <span class="conf">${[1, 2, 3].map(n => `<i class="${n <= c.conf ? 'f' : ''}"></i>`).join('')}<small style="margin-left:6px">${['없음', '부족', '보통', '충분'][c.conf]}</small></span></h3>
      ${it.retail ? `<div class="retail"><span>정가</span><b class="num">${won(it.retail)}</b></div><div class="note">정가는 참고만 해요. 수익 계산은 아래 판매완료 가격만 써요.</div>` : ''}
      ${cs.map(x => `<button class="comp" data-comp="${esc(x.id)}"><span class="src ${x.source === '직접 등록' ? 'manual' : ''}">${esc(x.source)}</span><span class="c-mid">${esc(x.cond || '상태 미입력')} · ${new Date(x.date).toLocaleDateString('ko-KR', { month: 'numeric', day: 'numeric' })}</span><b class="num">${won(x.price)}</b></button>`).join('')}
      ${!ok ? `<div class="warnbox" style="margin-top:${cs.length ? 12 : 0}px"><b>판매완료 가격이 2건 이상 필요해요.</b><br>비슷한 상품이 실제로 팔린 가격을 넣으면 바로 다시 계산해요.</div>` : `<div class="row" style="margin-top:4px"><span>중앙값</span><b class="num">${won(c.sell)}</b></div>`}
      <button class="ghost-btn" data-act="add-comp">+ 판매완료 가격 추가</button>
    </div>

    ${live() ? '' : `<div class="card"><h3>사기 전에 확인할 것</h3>
      ${[['사진상 택·로고 위치가 정품과 같아요', 1], ['오염·보풀은 사진만으로 확인이 어려워요', 0], ['판매자에게 실측 사이즈를 물어보세요', 0]].map(([t, g]) => `<div class="check">${svg(g ? IC.check : '<path d="M12 8v5M12 17h.01"/><circle cx="12" cy="12" r="10"/>', 2.4, `style="color:var(--${g ? 'brand' : 'warn'})"`)}${t}</div>`).join('')}
      <div class="note" style="margin-bottom:0">AI 분석 연결 전이라 예시 문구예요.</div></div>`}

    <div class="card"><h3>네고 메시지</h3>
      <div class="msg" id="nego">안녕하세요! ${esc(it.title)} 아직 판매 중인가요? 혹시 ${won(offer)}에 가능하시면 오늘 바로 받으러 갈게요.</div>
      <button class="ghost-btn" data-act="copy">메시지 복사</button>
    </div>
  </div>`;
}
function bindSlider(it) {
  const r = $('#buy'); if (!r) return; const c0 = calc(it);
  const pct = v => (v - r.min) / (r.max - r.min) * 100;
  const marks = [[it.askPrice || it.buyPrice, '올린 가격'], [c0.maxBuy, '최대 매입가']].filter(m => m[0] >= +r.min && m[0] <= +r.max);
  $('#marks').innerHTML = marks.map(([v, l]) => `<span style="left:calc(13px + (100% - 26px) * ${pct(v) / 100})">${l}</span>`).join('');
  const up = () => {
    const b = +r.value, c = calc(it, b); r.style.setProperty('--p', pct(b) + '%');
    $('#sv').textContent = won(b);
    $('#sc').textContent = c.net >= S.set.target ? '이 가격이면 기준 이상 남아요' : c.net > 0 ? `남지만 기준보다 ${won(S.set.target - c.net)} 적어요` : '이 가격이면 손해예요';
    $('#calc').innerHTML = `
      <div class="row" style="margin-top:6px"><span>예상 판매가</span><b class="num">${won(c.sell)}</b></div>
      <div class="row"><span>매입가</span><b class="num">− ${won(b)}</b></div>
      <div class="row"><span>이동·세탁·포장${S.set.feePct ? '·수수료' : ''}</span><b class="num">− ${won(c.cost)}</b></div>
      <div class="row total"><span>순이익</span><b class="num ${c.net < 0 ? 'neg' : ''}">${c.net >= 0 ? '+' : ''}${won(c.net)}</b></div>
      <div class="kpis"><div class="kpi"><small>수익률</small><b class="num">${Math.round(c.rate * 100)}%</b></div><div class="kpi"><small>시간당 (${S.set.minutes}분 기준)</small><b class="num">${won(c.hour)}</b></div></div>`;
  };
  r.addEventListener('input', up); up();
}
function renderCTA(it) {
  if (it.status !== 'new' && it.status !== 'watch') { $('#cta').innerHTML = `<button class="btn pri" data-act="row" data-id="${esc(it.id)}">상태 바꾸기</button>`; return; }
  const w = it.status === 'watch';
  $('#cta').innerHTML = `<button class="btn sec icon${w ? ' on' : ''}" data-act="watch" aria-label="${w ? '관심 빼기' : '관심 담기'}"><svg width="24" height="24" viewBox="0 0 24 24" fill="${w ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M12 21s-8-5.2-8-11a4.5 4.5 0 018-2.8A4.5 4.5 0 0120 10c0 5.8-8 11-8 11z"/></svg></button><button class="btn sec" data-act="pass">패스</button><button class="btn pri" data-act="buy">샀어요</button>`;
}

/* ---------- FORMS ---------- */
let F = {};
function itemForm(opts = {}) {
  const it = opts.item || {};
  F = { edit: it.id || null, category: it.category || '의류', img: opts.img || null, imgChanged: !!opts.img, source: { link: '링크', shot: '사진', manual: '직접 입력' }[opts.mode] || it.source || '직접 입력' };
  const photo = F.img || imgSrc(it);
  const urlField = `<label class="field"><span>매물 링크${opts.mode === 'link' ? '' : ' (선택)'}</span><input id="f-url" type="url" inputmode="url" placeholder="https://www.daangn.com/..." value="${esc(it.url || '')}"></label>`;
  sheet(`<h3>${F.edit ? '매물 수정' : '매물 등록'}</h3><p class="sub">${F.edit ? '바뀐 내용만 고치면 돼요' : '상품명과 가격만 있으면 바로 계산해요'}</p>
   ${opts.mode === 'link' ? urlField : ''}
   <button class="photo-slot" data-act="pick-photo" id="photo-slot">${photo ? `<img src="${esc(photo)}" alt="">` : `<span>${svg(IC.camera, 2, 'width="26" height="26"')}사진 추가 (선택)</span>`}</button>
   <label class="field"><span>상품명</span><input id="f-title" autocomplete="off" placeholder="예) 폴로 케이블 니트 L 네이비" value="${esc(it.title || '')}"></label>
   <div class="field" style="margin-bottom:6px"><span>카테고리</span></div>
   <div class="pick">${CATEGORIES.map(c => `<button class="chip${c === F.category ? ' on' : ''}" data-cat="${c}">${c}</button>`).join('')}</div>
   <div class="field-2"><label class="field"><span>${F.edit ? '매입가' : '올라온 가격'}</span><input id="f-buy" inputmode="numeric" placeholder="28,000" value="${it.buyPrice ? it.buyPrice.toLocaleString('ko-KR') : ''}"></label>
   <label class="field"><span>정가 (선택)</span><input id="f-retail" inputmode="numeric" placeholder="259,000" value="${it.retail ? it.retail.toLocaleString('ko-KR') : ''}"></label></div>
   ${opts.mode === 'link' ? '' : urlField}
   <div class="field-2"><label class="field"><span>동네 (선택)</span><input id="f-place" placeholder="망원동" value="${esc(it.place || '')}"></label>
   <label class="field"><span>상태 메모 (선택)</span><input id="f-cond" placeholder="보풀 없음" value="${esc(it.cond || '')}"></label></div>
   <button class="btn pri full" data-act="save-item">${F.edit ? '저장' : '등록하고 계산하기'}</button>`);
}
async function saveItem(btn) {
  const title = $('#f-title').value.trim(), buy = num($('#f-buy').value), url = $('#f-url').value.trim();
  if (!title) return toast('상품명을 적어주세요');
  if (!buy) return toast('가격을 적어주세요');
  if (url && !/^https?:\/\//i.test(url)) return toast('링크는 https://로 시작해야 해요');
  const data = { title, category: F.category, buyPrice: buy, retail: num($('#f-retail').value) || null, url, place: $('#f-place').value.trim(), cond: $('#f-cond').value.trim() };
  btn.disabled = true; btn.textContent = '저장하는 중…';
  try {
    if (F.imgChanged && F.img) {
      if (live()) data.imageId = (await API.call('uploadImage', { dataUrl: F.img })).imageId;
      else data.img = F.img;
    }
    if (F.edit) {
      if (await updateItem(F.edit, data)) { closeSheet(); refreshDetail(); toast('수정했어요'); }
      else { btn.disabled = false; btn.textContent = '다시 시도'; }
    } else {
      const it = await createItem({ ...data, askPrice: buy, status: 'new', source: F.source });
      closeSheet(); S.filter = '전체'; switchTab('home'); setTimeout(() => openDetail(it.id), 250); toast('등록했어요');
    }
  } catch (e) { toast('저장 못 했어요: ' + e.message); btn.disabled = false; btn.textContent = '다시 시도'; }
}
function shrink(file, max = 1000) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => {
      const k = Math.min(1, max / Math.max(img.width, img.height)), c = document.createElement('canvas');
      c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(img.src); res(c.toDataURL('image/jpeg', .82));
    };
    img.onerror = () => rej(new Error('이미지를 읽지 못했어요')); img.src = URL.createObjectURL(file);
  });
}
function compForm() {
  F = { src: SOURCES[0] };
  const today = new Date(Date.now() - new Date().getTimezoneOffset() * 6e4).toISOString().slice(0, 10);
  sheet(`<h3>판매완료 가격 추가</h3><p class="sub">실제로 팔린 가격만 넣어주세요. 판매 중인 가격은 시세가 아니에요.</p>
   <label class="field"><span>팔린 가격</span><input id="c-price" inputmode="numeric" placeholder="55,000"></label>
   <div class="field" style="margin-bottom:6px"><span>어디서 확인했나요</span></div>
   <div class="pick">${SOURCES.map(s => `<button class="chip${s === F.src ? ' on' : ''}" data-src="${s}">${s}</button>`).join('')}</div>
   <div class="field-2"><label class="field"><span>상태 (선택)</span><input id="c-cond" placeholder="상태 좋음"></label>
   <label class="field"><span>팔린 날짜</span><input id="c-date" type="date" value="${today}" max="${today}"></label></div>
   <button class="btn pri full" data-act="save-comp">추가하기</button>`);
}
async function saveComp(btn) {
  const price = num($('#c-price').value); if (!price) return toast('팔린 가격을 적어주세요');
  const d = $('#c-date').value; btn.disabled = true;
  try {
    await addComp({ itemId: cur.id, price, source: F.src, cond: $('#c-cond').value.trim(), date: d ? iso(new Date(d + 'T12:00:00')) : iso() });
    closeSheet(); refreshDetail(); renderHome(); toast('시세에 반영했어요');
  } catch (e) { toast('추가 못 했어요: ' + e.message); btn.disabled = false; }
}

/* ---------- INVENTORY ---------- */
const TABS = [['watch', '관심'], ['bought', '매입'], ['selling', '판매중'], ['sold', '완료']];
function renderInv() {
  const by = k => S.items.filter(i => i.status === k);
  const stock = [...by('bought'), ...by('selling')], exp = stock.reduce((s, i) => s + expectedNet(i), 0);
  $('#inv-strip').innerHTML = `<div><small>재고</small><b class="num">${stock.length}개</b></div><div><small>묶인 돈</small><b class="num">${man(stock.reduce((s, i) => s + i.buyPrice, 0))}</b></div><div><small>예상 수익</small><b class="num" style="color:var(--brand)">${exp >= 0 ? '+' : ''}${man(exp)}</b></div>`;
  const seg = $('#seg'); seg.querySelectorAll('button').forEach(b => b.remove());
  TABS.forEach(([k, l]) => seg.insertAdjacentHTML('beforeend', `<button data-inv="${k}" class="${S.invTab === k ? 'on' : ''}">${l}<em>${by(k).length}</em></button>`));
  $('#seg-ind').style.transform = `translateX(${TABS.findIndex(t => t[0] === S.invTab) * 100}%)`;
  let L = by(S.invTab);
  if (S.invTab === 'sold') L = L.sort((a, b) => new Date(b.soldAt) - new Date(a.soldAt));
  $('#inv-list').innerHTML = L.length ? `<div class="inv">${L.map(i => {
    const c = calc(i); let right = '', sub = '';
    if (i.status === 'watch') { sub = `${won(i.buyPrice)}${i.place ? ' · ' + esc(i.place) : ''}`; right = c.n >= 2 ? `<span class="profit num" style="font-size:15px">${c.net >= 0 ? '+' : ''}${won(c.net)}</span><span class="pill ${c.v.k}" style="font-size:11px">${c.v.l}</span>` : `<span class="pill unk" style="font-size:11px">시세 확인 필요</span>`; }
    if (i.status === 'bought') { const d = daysSince(i.boughtAt); sub = `매입 ${won(i.buyPrice)}`; right = `<span class="age${d >= 7 ? ' old' : ''}">${d}일째</span>${c.n ? `<span style="font-size:13px;color:var(--ink3)">시세 ${man(c.sell)}</span>` : ''}`; }
    if (i.status === 'selling') { const d = daysSince(i.listedAt); sub = `매입 ${won(i.buyPrice)} → 판매가 ${won(i.listPrice)}`; right = `<span class="age${d >= 14 ? ' old' : ''}">${d}일째</span>${d >= 14 ? '<span style="font-size:12px;color:var(--bad);font-weight:600">가격 내려보세요</span>' : ''}`; }
    if (i.status === 'sold') { const n = soldNet(i); sub = `${won(i.buyPrice)} → ${won(i.soldPrice)} · ${daysBetween(i.boughtAt || i.createdAt, i.soldAt)}일 만에`; right = `<span class="profit num ${n < 0 ? 'neg' : ''}" style="font-size:16px">${n >= 0 ? '+' : ''}${won(n)}</span>`; }
    return `<button class="inv-row" data-row="${esc(i.id)}">${thumb(i)}<div class="d-body"><div class="d-title">${esc(i.title)}</div><div class="d-meta num">${sub}</div></div><div class="r">${right}</div></button>`;
  }).join('')}</div>`
    : `<div class="empty" style="margin-top:14px">${{ watch: '발굴 탭에서 매물을 오른쪽으로 밀면<br>여기에 모여요.', bought: '산 물건이 아직 없어요.', selling: '판매 중인 물건이 없어요.', sold: '아직 판매 완료된 물건이 없어요.' }[S.invTab]}</div>`;
}
function rowSheet(id) {
  const it = find(id); if (!it) return; const c = calc(it), dt = `<button class="btn sec full" data-act="detail" data-id="${esc(id)}">자세히 보기</button>`;
  if (it.status === 'watch') return sheet(`<h3>${esc(it.title)}</h3><p class="sub">${won(it.buyPrice)}${c.n >= 2 ? ` · 예상 순이익 ${won(c.net)}` : ''}</p>
    <div class="btn-stack"><button class="btn pri full" data-act="buy" data-id="${esc(id)}">샀어요</button>${dt}<button class="btn sec full" data-act="unwatch" data-id="${esc(id)}">관심 빼기</button></div>`);
  if (it.status === 'bought') return sheet(`<h3>판매 등록했나요?</h3><p class="sub">올린 가격을 넣으면 판매중으로 옮길게요.${c.n ? ` 시세 중앙값은 ${won(c.sell)}이에요.` : ''}</p>
    <label class="field"><span>판매 등록가</span><input id="f-list" inputmode="numeric" value="${c.n ? (Math.round(c.sell * 1.08 / 1000) * 1000).toLocaleString('ko-KR') : ''}"></label>
    <div class="btn-stack"><button class="btn pri full" data-act="do-list" data-id="${esc(id)}">판매중으로 옮기기</button>${dt}</div>`);
  if (it.status === 'selling') return sheet(`<h3>${esc(it.title)}</h3><p class="sub">${daysSince(it.listedAt)}일째 판매 중 · 등록가 ${won(it.listPrice)}</p>
    <label class="field"><span>실제로 판 가격</span><input id="f-sold" inputmode="numeric" value="${(it.listPrice || '').toLocaleString('ko-KR')}"></label>
    <div class="btn-stack"><button class="btn pri full" data-act="do-sold" data-id="${esc(id)}">팔렸어요</button>
    <button class="btn sec full" data-act="do-drop" data-id="${esc(id)}">가격 10% 내리기 (${won(Math.round(it.listPrice * .9 / 1000) * 1000)})</button>${dt}</div>`);
  const n = soldNet(it);
  sheet(`<h3>${n >= 0 ? won(n) + ' 남겼어요' : won(-n) + ' 손해였어요'}</h3><p class="sub">${won(it.buyPrice)}에 사서 ${won(it.soldPrice)}에 팔았어요 · ${daysBetween(it.boughtAt || it.createdAt, it.soldAt)}일 걸림</p>
    <div class="kpis" style="margin-top:0"><div class="kpi"><small>수익률</small><b class="num">${Math.round(n / it.buyPrice * 100)}%</b></div><div class="kpi"><small>정가 대비 매입</small><b class="num">${it.retail ? Math.round(it.buyPrice / it.retail * 100) + '%' : '-'}</b></div></div><div style="height:16px"></div>
    <div class="btn-stack">${dt}<button class="btn sec full" data-act="close-sheet">닫기</button></div>`);
}

/* ---------- STATS ---------- */
function renderStats(animate) {
  const el = $('#stats'), sold = S.items.filter(i => i.status === 'sold' && i.soldAt && i.soldPrice);
  if (!sold.length) { el.innerHTML = `<div class="empty" style="margin-top:16px">아직 판매 완료된 물건이 없어요.<br>재고 탭에서 '팔렸어요'를 누르면 여기에 수익이 쌓여요.</div>`; return; }
  const now = new Date(), months = [];
  for (let k = 5; k >= 0; k--) { const m = new Date(now.getFullYear(), now.getMonth() - k, 1); months.push({ y: m.getFullYear(), m: m.getMonth(), label: (m.getMonth() + 1) + '월', v: 0 }); }
  const inM = (i, b) => { const t = new Date(i.soldAt); return t.getFullYear() === b.y && t.getMonth() === b.m; };
  sold.forEach(i => { const b = months.find(b => inM(i, b)); if (b) b.v += soldNet(i); });
  const cm = months[5], pm = months[4], curList = sold.filter(i => inM(i, cm));
  const prevSame = sold.filter(i => inM(i, pm) && new Date(i.soldAt).getDate() <= now.getDate()).reduce((s, i) => s + soldNet(i), 0);
  const delta = cm.v - prevSame, buySum = curList.reduce((s, i) => s + i.buyPrice, 0);
  const avgRate = buySum ? cm.v / buySum : 0, hourly = curList.length ? cm.v / (curList.length * S.set.minutes / 60) : 0;
  const byCat = {}; sold.forEach(i => (byCat[i.category] ??= []).push(i));
  const rateOf = a => a.reduce((s, i) => s + soldNet(i), 0) / Math.max(1, a.reduce((s, i) => s + i.buyPrice, 0));
  const cats = Object.entries(byCat).map(([c, a]) => [c, rateOf(a)]).sort((a, b) => b[1] - a[1]), cmax = Math.max(...cats.map(c => c[1]), .01);
  el.innerHTML = `<div class="sub" style="margin-top:12px">${cm.m + 1}월 순이익</div>
    <div class="hero-num num${cm.v < 0 ? ' neg' : ''}" id="month-num">${won(cm.v)}</div>
    <div class="delta${delta < 0 ? ' down' : ''}">지난달 같은 날보다 ${delta >= 0 ? '+' : '−'}${won(Math.abs(delta))}</div>
    <div class="strip"><div><small>판매</small><b class="num">${curList.length}건</b></div><div><small>평균 수익률</small><b class="num">${curList.length ? Math.round(avgRate * 100) + '%' : '-'}</b></div><div><small>시간당</small><b class="num">${curList.length ? man(hourly) + '원' : '-'}</b></div></div>
    <div class="card chart"><h3>월별 순이익 <small>최근 6개월</small></h3>${chartSVG(months)}</div>
    <div class="card"><h3>품목별 수익률 <small>매입가 대비</small></h3>${cats.map(([k, v]) => `<div class="bar-row"><span>${k}</span><div class="track"><div class="fill" style="--v:${animate ? 0 : Math.max(v, 0) / cmax}" data-v="${Math.max(v, 0) / cmax}"></div></div><b class="num">${Math.round(v * 100)}%</b></div>`).join('')}</div>
    <div class="card"><h3>내 거래에서 찾은 패턴</h3>${insights(sold)}</div>`;
  if (animate) animateStats(cm.v);
}
function chartSVG(months) {
  const W = 340, H = 180, L = 30, top = 24, bot = 26, bw = 28, gap = (W - L - bw * 6) / 6;
  const mx = Math.max(...months.map(m => m.v), 1), step = [10000, 20000, 50000, 100000, 200000, 500000, 1e6, 2e6].find(s => mx / s <= 3) || 5e6, max = Math.ceil(mx / step) * step;
  const y = v => top + (H - top - bot) * (1 - Math.max(v, 0) / max);
  let g = '';
  for (let v = step; v <= max; v += step) g += `<line x1="${L}" x2="${W}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)" stroke-dasharray="3 4"/><text x="0" y="${y(v) + 4}" font-size="10" fill="var(--ink3)">${v / 10000}만</text>`;
  g += months.map((m, i) => {
    const x = L + gap / 2 + i * (bw + gap), last = i === 5, h = H - bot - y(m.v);
    return `<rect x="${x}" y="${y(m.v)}" width="${bw}" height="${h}" rx="7" fill="${last ? 'var(--brand)' : 'var(--mute-soft)'}" data-y="${y(m.v)}" data-h="${h}"/>
    ${last ? `<text x="${x + bw / 2}" y="${y(m.v) - 8}" text-anchor="middle" font-size="12" font-weight="700" fill="var(--brand-ink)">${man(m.v)}</text>` : ''}
    <text x="${x + bw / 2}" y="${H - 6}" text-anchor="middle" font-size="12" fill="${last ? 'var(--ink)' : 'var(--ink3)'}" font-weight="${last ? 700 : 400}">${m.label}</text>`;
  }).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="최근 6개월 월별 순이익 막대 차트">${g}</svg>`;
}
function animateStats(to) {
  const rects = $$('#chart-anim rect, #stats svg rect'), base = 180 - 26;
  rects.forEach(r => { r.style.transition = 'none'; r.setAttribute('y', base); r.setAttribute('height', 0); });
  requestAnimationFrame(() => requestAnimationFrame(() => {
    rects.forEach((r, i) => { r.style.transition = `all .8s var(--ease) ${i * 60}ms`; r.setAttribute('y', r.dataset.y); r.setAttribute('height', r.dataset.h); });
    $$('#stats .fill').forEach(f => f.style.setProperty('--v', f.dataset.v));
  }));
  const el = $('#month-num'), t0 = performance.now();
  const step = t => { const k = Math.min((t - t0) / 900, 1), e = 1 - Math.pow(1 - k, 3); el.textContent = won(to * e); if (k < 1) requestAnimationFrame(step); };
  requestAnimationFrame(step);
}
function insights(sold) {
  if (sold.length < 5) return `<p class="muted" style="margin:0">판매 기록이 5건 이상 쌓이면 잘 남는 품목과 매입 조건을 찾아드려요. 지금 ${sold.length}건이에요.</p>`;
  const out = [], rate = a => a.reduce((s, i) => s + soldNet(i), 0) / Math.max(1, a.reduce((s, i) => s + i.buyPrice, 0));
  const byCat = {}; sold.forEach(i => (byCat[i.category] ??= []).push(i));
  const speed = Object.entries(byCat).filter(([, a]) => a.length >= 2).map(([c, a]) => [c, a.reduce((s, i) => s + daysBetween(i.boughtAt || i.createdAt, i.soldAt), 0) / a.length]).sort((a, b) => a[1] - b[1]);
  if (speed.length >= 2) out.push(['bolt', `<b>${iga(speed[0][0])} 가장 빨리 팔려요.</b> 평균 ${Math.round(speed[0][1])}일, ${eunneun(speed.at(-1)[0])} ${Math.round(speed.at(-1)[1])}일이 걸렸어요.`]);
  const held = i => daysBetween(i.listedAt || i.boughtAt || i.createdAt, i.soldAt);
  const old = sold.filter(i => held(i) >= 14), fresh = sold.filter(i => held(i) < 14);
  if (old.length >= 2 && fresh.length >= 2 && rate(old) < rate(fresh)) out.push(['clock', `<b>2주 넘게 걸린 물건은 수익률이 ${Math.round(rate(old) * 100)}%예요.</b> 2주 안에 팔린 물건은 ${Math.round(rate(fresh) * 100)}%였어요. 오래 걸리면 일찍 가격을 내리는 게 나아요.`]);
  const cheap = sold.filter(i => i.retail && i.buyPrice / i.retail <= .25);
  if (cheap.length >= 3) out.push(['check', `<b>정가의 25% 이하로 산 물건</b>은 ${cheap.length}건 중 ${cheap.filter(i => soldNet(i) >= S.set.target).length}건이 기준(${man(S.set.target)}원) 이상 남았어요.`]);
  if (!out.length) { const best = Object.entries(byCat).map(([c, a]) => [c, rate(a)]).sort((a, b) => b[1] - a[1])[0]; out.push(['trend', `<b>${eunneun(best[0])} 수익률이 가장 높아요.</b> 매입가 대비 평균 ${Math.round(best[1] * 100)}%가 남았어요.`]); }
  return out.map(([ic, t]) => `<div class="insight"><div class="ib">${svg(IC[ic], 2.2)}</div><div>${t}</div></div>`).join('');
}

/* ---------- SETTINGS ---------- */
const COSTS = [['move', '이동·교통비', '원'], ['clean', '세탁·포장', '원'], ['feePct', '판매 수수료', '%'], ['minutes', '건당 소요 시간', '분']];
function renderMe() {
  const s = S.set, conn = API.cfg();
  $('#me').innerHTML = `
  <div class="card"><h3>화면 테마</h3><div class="seg seg3"><div class="seg-ind" id="theme-ind"></div>${[['light', '라이트'], ['dark', '다크'], ['system', '기기 설정']].map(([k, l]) => `<button data-theme-set="${k}">${l}</button>`).join('')}</div></div>
  <div class="card"><h3>알림 키워드 <small>당근 앱에 등록한 키워드</small></h3>
    <div class="kw">${(s.keywords || []).map((k, i) => `<span>${esc(k.kw)}${k.max ? `<small>~${man(k.max)}</small>` : ''}<button class="x" data-act="del-kw" data-i="${i}" aria-label="${esc(k.kw)} 삭제">×</button></span>`).join('')}<button class="add" data-act="add-kw">+ 추가</button></div>
    <p class="note" style="margin-bottom:0">MacroDroid 알림을 받을 때 이 키워드와 최대 가격으로 먼저 걸러요.</p></div>
  <div class="card"><h3>매입 기준</h3><div class="slider-cap">이만큼은 남아야 '사세요'라고 알려드려요</div>
    <div class="slider-val num" id="target-val"></div><input type="range" id="target" min="3000" max="30000" step="1000" value="${s.target}" aria-label="최소 순이익"></div>
  <div class="card"><h3>건당 기본 비용 <small>순이익 계산에 빠져요</small></h3>
    ${COSTS.map(([k, l, u]) => `<button class="row" data-act="edit-cost" data-k="${k}"><span>${l}</span><b class="num">${u === '원' ? won(s[k]) : s[k] + u}</b></button>`).join('')}</div>
  <div class="card"><h3>연결</h3>
    <button class="conn" data-act="conn-sheets"><span class="oi">${svg(IC.sheet)}</span><div><b>Google Sheets</b><small>${conn ? '매물·거래 기록을 저장하고 있어요' : '연결하면 실제 기록으로 바뀌어요'}</small></div><span class="status${conn ? ' ok' : ''}">${conn ? '연결됨' : '연결하기'}</span></button>
    <div class="conn"><span class="oi">${svg(IC.phone)}</span><div><b>MacroDroid 알림 수집</b><small>당근 알림을 앱으로 전달</small></div><span class="status">다음 단계</span></div>
    <div class="conn"><span class="oi">${svg(IC.spark)}</span><div><b>AI 매물 분석</b><small>사진·제목에서 상품 식별</small></div><span class="status">다음 단계</span></div></div>
  ${live() ? '' : '<p class="note">지금은 예시 데이터를 보여주고 있어요. Google Sheets를 연결하면 실제 기록으로 바뀌어요.</p>'}`;
  const tg = $('#target');
  const up = () => { $('#target-val').textContent = `순이익 ${won(+tg.value)} 이상`; tg.style.setProperty('--p', (tg.value - 3000) / 27000 * 100 + '%'); };
  tg.addEventListener('input', () => { S.set.target = +tg.value; up(); });
  tg.addEventListener('change', () => { saveSettings({ target: +tg.value }); renderHome(); toast('기준을 바꿨어요. 발굴 목록에 바로 반영돼요'); });
  up(); paintThemeSeg();
}
function costSheet(k) {
  const [, l, u] = COSTS.find(c => c[0] === k);
  sheet(`<h3>${l}</h3><p class="sub">매물 1건을 사고팔 때 드는 평균 ${u === '분' ? '시간' : '비용'}이에요</p>
    <label class="field"><span>${l} (${u})</span><input id="cost-v" inputmode="${u === '%' ? 'decimal' : 'numeric'}" value="${S.set[k]}"></label>
    <button class="btn pri full" data-act="save-cost" data-k="${k}">저장</button>`);
}
function kwSheet() {
  sheet(`<h3>키워드 추가</h3><p class="sub">당근 앱에 등록한 키워드와 같은 이름으로 적어주세요</p>
    <label class="field"><span>키워드</span><input id="kw-v" placeholder="예) 케이스티파이"></label>
    <label class="field"><span>최대 가격 (선택)</span><input id="kw-max" inputmode="numeric" placeholder="20,000"></label>
    <button class="btn pri full" data-act="save-kw">추가하기</button>`);
}
function connSheet() {
  const c = API.cfg();
  sheet(`<h3>Google Sheets 연결</h3><p class="sub">Apps Script 웹 앱 주소와 API 토큰을 넣어주세요. 만드는 방법은 SETUP.md에 있어요.</p>
    <label class="field"><span>웹 앱 주소</span><input id="k-url" type="url" autocomplete="off" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(c?.url || '')}"></label>
    <label class="field"><span>API 토큰</span><input id="k-token" type="password" autocomplete="off" value="${esc(c?.token || '')}"></label>
    <div class="btn-stack"><button class="btn pri full" data-act="do-conn">${c ? '다시 연결하기' : '연결하기'}</button>${c ? '<button class="btn danger full" data-act="do-disconn">연결 해제하고 예시 데이터 보기</button>' : ''}</div>`);
}
async function doConnect(btn) {
  const url = $('#k-url').value.trim(), token = $('#k-token').value.trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) return toast('…/exec로 끝나는 웹 앱 주소를 넣어주세요');
  if (!token) return toast('API 토큰을 넣어주세요');
  btn.disabled = true; btn.textContent = '확인하는 중…';
  try {
    const d = await API.call('bootstrap', {}, { url, token });
    store.set('jj-conn', { url, token }); setMode('live'); applyRemote(d);
    closeSheet(); renderAll(); renderMe(); toast('연결했어요');
  } catch (e) { toast('연결 못 했어요: ' + e.message); btn.disabled = false; btn.textContent = '다시 시도'; }
}
function disconnect() {
  store.del('jj-conn'); store.del('jj-cache'); setMode('demo');
  Object.assign(S, demoData()); S.set = { ...DEFAULT_SET }; S.syncedAt = 0;
  closeSheet(); renderAll(); renderMe(); toast('예시 데이터로 바꿨어요');
}
function setMode(m) { S.mode = m; document.body.classList.toggle('live', m === 'live'); $('#refresh-btn').hidden = m !== 'live'; }

/* ---------- THEME ---------- */
const mq = matchMedia('(prefers-color-scheme: dark)');
let themePref = store.get('jj-theme', 'light');
function applyTheme(p, anim) {
  themePref = p; store.set('jj-theme', p);
  const dark = p === 'dark' || (p === 'system' && mq.matches);
  const go = () => { document.documentElement.classList.toggle('dark', dark); const m = $('meta[name="theme-color"]'); if (m) m.content = dark ? '#0A0C0F' : '#F2F4F6'; };
  if (anim && document.startViewTransition && !matchMedia('(prefers-reduced-motion: reduce)').matches) document.startViewTransition(go); else go();
  paintThemeSeg();
}
function paintThemeSeg() {
  const ind = $('#theme-ind'); if (!ind) return;
  ind.style.transform = `translateX(${['light', 'dark', 'system'].indexOf(themePref) * 100}%)`;
  $$('[data-theme-set]').forEach(b => b.classList.toggle('on', b.dataset.themeSet === themePref));
}
mq.addEventListener?.('change', () => { if (themePref === 'system') applyTheme('system', true); });

/* ---------- SHEET / TOAST / PUSH ---------- */
function sheet(html) { $('#sheet').innerHTML = '<div class="grab"></div>' + html; $('#scrim').classList.add('on'); $('#sheet').classList.add('on'); }
function closeSheet() { $('#scrim').classList.remove('on'); $('#sheet').classList.remove('on'); }
let tt; function toast(m) { const t = $('#toast'); t.textContent = m; t.classList.add('on'); clearTimeout(tt); tt = setTimeout(() => t.classList.remove('on'), 2200); }
function simPush() {
  const id = 'p' + Date.now().toString(36);
  const it = { id, status: 'new', title: '폴로 랄프로렌 옥스포드 셔츠 M 화이트', category: '의류', buyPrice: 18000, askPrice: 18000, retail: 189000, cond: '택 없음 · 상태 좋음', place: '망원동', source: '키워드 알림', createdAt: iso() };
  S.items.push(it);
  [[42000, '판매완료 캡처', 6], [39000, '판매완료 캡처', 15], [45000, '내 판매기록', 29]].forEach(([p, s, d], k) => S.comps.push({ id: id + k, itemId: id, price: p, source: s, cond: '상태 좋음', date: iso(Date.now() - d * DAY) }));
  const c = calc(it), p = $('#push');
  p.innerHTML = `${thumb(it)}<div style="min-width:0"><small>당근 키워드 알림 · 방금</small><b>${esc(it.title)}</b><small class="num">${won(it.buyPrice)} · 예상 순이익 <span style="color:var(--brand);font-weight:700">+${won(c.net)}</span></small></div>`;
  p.onclick = () => { p.classList.remove('on'); switchTab('home'); openDetail(id); };
  if (!cur) renderHome();
  p.classList.add('on'); setTimeout(() => p.classList.remove('on'), 4800);
}
function addSheet() {
  sheet(`<h3>매물 등록</h3><p class="sub">알림으로 못 받은 매물도 바로 계산해 드려요</p>
   ${[['link', '링크 붙여넣기', '당근 매물 주소를 함께 저장해요', IC.link], ['shot', '사진으로 등록', '매물 사진이나 캡처를 올려요', IC.camera], ['manual', '직접 입력', '상품명과 가격만 적어요', IC.pen]]
     .map(([k, b, s, p]) => `<button class="opt" data-act="add-${k}"><span class="oi">${svg(p)}</span><span><b>${b}</b><small>${s}</small></span></button>`).join('')}`);
}

/* ---------- NAV ---------- */
const SCR = { home: '#s-home', inv: '#s-inv', stats: '#s-stats', me: '#s-me' };
function switchTab(t) {
  if (cur) closeDetail();
  const changed = S.tab !== t; S.tab = t;
  Object.entries(SCR).forEach(([k, s]) => $(s).classList.toggle('on', k === t));
  $$('.tab').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  if (t === 'home') renderHome(); if (t === 'inv') renderInv();
  if (t === 'stats') renderStats(changed); if (t === 'me') renderMe();
}
function renderAll() { renderHome(); renderInv(); if (S.tab === 'stats') renderStats(false); }

let fileTarget = 'new';
const ACTIONS = {
  back: closeDetail, 'close-sheet': closeSheet, add: addSheet, 'sim-push': simPush, refresh: () => refresh(true),
  'theme-flip': () => applyTheme(document.documentElement.classList.contains('dark') ? 'light' : 'dark', true),
  watch: () => { setStatus(cur.id, cur.status === 'watch' ? 'new' : 'watch').then(() => cur && renderCTA(find(cur.id))); },
  unwatch: b => { closeSheet(); setStatus(b.dataset.id, 'new'); },
  pass: () => { const id = cur.id; closeDetail(); setStatus(id, 'pass'); },
  row: b => rowSheet(b.dataset.id),
  buy: b => {
    const it = find(b.dataset.id || cur.id);
    sheet(`<h3>얼마에 샀나요?</h3><p class="sub">네고했다면 실제 금액으로 바꿔주세요</p><label class="field"><span>실제 매입가</span><input id="f-buy2" inputmode="numeric" value="${it.buyPrice.toLocaleString('ko-KR')}"></label><button class="btn pri full" data-act="do-buy" data-id="${esc(it.id)}">매입 완료</button>`);
  },
  'do-buy': b => { const id = b.dataset.id, v = num($('#f-buy2').value) || find(id).buyPrice; closeSheet(); if (cur) closeDetail(); setStatus(id, 'bought', { buyPrice: v }); S.invTab = 'bought'; switchTab('inv'); },
  'do-list': b => { const v = num($('#f-list').value); if (!v) return toast('판매 등록가를 적어주세요'); closeSheet(); S.invTab = 'selling'; setStatus(b.dataset.id, 'selling', { listPrice: v }); },
  'do-sold': b => { const v = num($('#f-sold').value); if (!v) return toast('판 가격을 적어주세요'); closeSheet(); S.invTab = 'sold'; setStatus(b.dataset.id, 'sold', { soldPrice: v }); },
  'do-drop': b => { const it = find(b.dataset.id), v = Math.round(it.listPrice * .9 / 1000) * 1000; closeSheet(); updateItem(it.id, { listPrice: v }).then(ok => ok && toast(`판매가를 ${won(v)}으로 내렸어요`)); },
  detail: b => { closeSheet(); const id = b.dataset.id; switchTab('home'); setTimeout(() => openDetail(id), 200); },
  'item-menu': () => sheet(`<h3>${esc(cur.title)}</h3><div style="height:8px"></div>
     <button class="opt" data-act="edit-item"><span class="oi">${svg(IC.pen)}</span><span><b>수정하기</b><small>상품명·가격·사진을 고쳐요</small></span></button>
     <button class="opt" data-act="ask-del"><span class="oi" style="color:var(--bad)">${svg(IC.trash)}</span><span><b>삭제하기</b><small>시세 기록도 함께 지워져요</small></span></button>`),
  'edit-item': () => itemForm({ item: cur }),
  'ask-del': () => sheet(`<h3>이 매물을 삭제할까요?</h3><p class="sub">${esc(cur.title)}과 시세 기록 ${compsOf(cur.id).length}건이 지워지고 되돌릴 수 없어요.</p><div class="btn-stack"><button class="btn danger full" data-act="do-del">삭제</button><button class="btn sec full" data-act="close-sheet">취소</button></div>`),
  'do-del': () => { const id = cur.id; closeSheet(); closeDetail(); deleteItem(id).then(ok => ok && toast('삭제했어요')); },
  'add-link': () => itemForm({ mode: 'link' }), 'add-manual': () => itemForm({ mode: 'manual' }),
  'add-shot': () => { fileTarget = 'new'; $('#file').click(); },
  'pick-photo': () => { fileTarget = 'form'; $('#file').click(); },
  'save-item': b => saveItem(b),
  'add-comp': () => compForm(), 'save-comp': b => saveComp(b),
  'del-comp': b => { const id = b.dataset.id; closeSheet(); deleteComp(id).then(ok => { if (ok) { refreshDetail(); renderHome(); toast('시세에서 뺐어요'); } }); },
  'edit-cost': b => costSheet(b.dataset.k),
  'save-cost': b => { const k = b.dataset.k, raw = $('#cost-v').value, v = k === 'feePct' ? Math.min(50, Math.max(0, parseFloat(raw) || 0)) : num(raw); if (k === 'minutes' && !v) return toast('1분 이상으로 적어주세요'); saveSettings({ [k]: v }); closeSheet(); renderMe(); renderAll(); toast('저장했어요'); },
  'add-kw': kwSheet,
  'save-kw': () => { const kw = $('#kw-v').value.trim(); if (!kw) return toast('키워드를 적어주세요'); saveSettings({ keywords: [...(S.set.keywords || []), { kw, max: num($('#kw-max').value) || null }] }); closeSheet(); renderMe(); },
  'del-kw': b => { const k = (S.set.keywords || []).slice(); k.splice(+b.dataset.i, 1); saveSettings({ keywords: k }); renderMe(); },
  'conn-sheets': connSheet, 'do-conn': b => doConnect(b), 'do-disconn': disconnect,
  toast: b => toast(b.dataset.msg),
  copy: () => {
    const n = $('#nego'), t = n.textContent;
    const sel = () => { const r = document.createRange(); r.selectNodeContents(n); getSelection().removeAllRanges(); getSelection().addRange(r); toast('메시지를 선택했어요. 길게 눌러 복사하세요'); };
    if (navigator.clipboard) navigator.clipboard.writeText(t).then(() => toast('메시지를 복사했어요'), sel); else sel();
  }
};
document.addEventListener('click', e => {
  const b = e.target.closest('[data-tab],[data-filter],[data-act],[data-inv],[data-row],[data-theme-set],[data-cat],[data-src],[data-comp]'); if (!b) return;
  if (b.dataset.tab) return switchTab(b.dataset.tab);
  if (b.dataset.filter) { S.filter = b.dataset.filter; return renderHome(); }
  if (b.dataset.inv) { S.invTab = b.dataset.inv; return renderInv(); }
  if (b.dataset.row) return rowSheet(b.dataset.row);
  if (b.dataset.themeSet) return applyTheme(b.dataset.themeSet, true);
  if (b.dataset.cat || b.dataset.src) {
    if (b.dataset.cat) F.category = b.dataset.cat; else F.src = b.dataset.src;
    b.parentElement.querySelectorAll('.chip').forEach(x => x.classList.toggle('on', x === b)); return;
  }
  if (b.dataset.comp) {
    const c = S.comps.find(x => x.id === b.dataset.comp); if (!c) return;
    return sheet(`<h3>이 시세를 뺄까요?</h3><p class="sub">${won(c.price)} · ${esc(c.source)} · ${esc(c.cond || '상태 미입력')}</p><div class="btn-stack"><button class="btn danger full" data-act="del-comp" data-id="${esc(c.id)}">시세에서 빼기</button><button class="btn sec full" data-act="close-sheet">취소</button></div>`);
  }
  ACTIONS[b.dataset.act]?.(b);
});
$('#file').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const img = await shrink(f);
    if (fileTarget === 'new') itemForm({ mode: 'shot', img });
    else { F.img = img; F.imgChanged = true; $('#photo-slot').innerHTML = `<img src="${img}" alt="">`; }
  } catch (err) { toast(err.message); }
});
$('#detail').addEventListener('scroll', e => e.target.classList.toggle('scrolled', e.target.scrollTop > 180));

/* ---------- BOOT ---------- */
(function boot() {
  applyTheme(themePref, false);
  if (API.cfg()) {
    setMode('live');
    const c = store.get('jj-cache', null);
    if (c) { S.items = c.items || []; S.comps = c.comps || []; S.set = { ...DEFAULT_SET, ...(c.set || {}) }; S.syncedAt = c.at || 0; }
    renderAll(); refresh(false);
  } else {
    setMode('demo'); Object.assign(S, demoData()); renderAll();
    setTimeout(() => { if (!live() && S.tab === 'home' && !cur) simPush(); }, 3500);
  }
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
