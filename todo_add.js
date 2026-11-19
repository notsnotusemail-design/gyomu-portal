// TODO（旧：引き継ぎ）の追加画面。トップ・TODO一覧・日次スケジュールで共通に使う。
// 使い方: openTodoAdd({ date: 'YYYY-MM-DD'（省略可）, onAdded: () => {...} })
// 日付はプルダウン（スマホだとダイヤル）を使わず、全部ボタンで選ぶ。
(function () {
  const SERVER = window.location.origin;
  const WD = ['日', '月', '火', '水', '木', '金', '土'];
  const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const addDays = (s, n) => { const [y, m, d] = s.split('-').map(Number); return ymd(new Date(y, m - 1, d + n)); };
  const today = () => ymd(new Date());
  const md = s => { const [y, m, d] = s.split('-').map(Number); const w = WD[new Date(y, m - 1, d).getDay()]; return `${m}/${d}(${w})`; };

  const css = `
.tda-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.28);z-index:900;align-items:center;justify-content:center;padding:16px}
.tda-ov.open{display:flex}
.tda{background:#fff;border-radius:16px;padding:22px 20px 18px;width:400px;max-width:100%;max-height:92vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.18);font-family:-apple-system,BlinkMacSystemFont,"Hiragino Sans","Yu Gothic",sans-serif;color:#1a1a1a}
.tda h3{font-size:15px;font-weight:700;margin:0 0 4px}
.tda .tda-note{font-size:11px;color:#999;margin-bottom:14px;line-height:1.5}
.tda .tda-lb{font-size:12px;font-weight:600;color:#666;margin:12px 0 6px;display:block}
.tda textarea{width:100%;box-sizing:border-box;padding:9px 10px;border:1.5px solid #e2e2e2;border-radius:9px;font-size:14px;font-family:inherit;resize:vertical;background:#fafafa}
.tda textarea:focus,.tda input[type=date]:focus{outline:none;border-color:#4f6ef7;background:#fff}
.tda-btns{display:flex;flex-wrap:wrap;gap:6px}
.tda-b{flex:1 1 0;min-width:0;min-height:40px;padding:8px 6px;border:1.5px solid #e0e0e0;border-radius:9px;background:#fff;color:#444;font-size:13px;font-weight:600;font-family:inherit;cursor:pointer;transition:all .12s;white-space:nowrap}
.tda-b:hover{background:#f5f5f5}
.tda-b.on{background:#1a1a1a;border-color:#1a1a1a;color:#fff}
.tda-b.t-todo.on{background:#dbeafe;border-color:#93c5fd;color:#1e3a8a}
.tda-b.t-memo.on{background:#fef3c7;border-color:#fcd34d;color:#78350f}
.tda-b.w-medium.on{background:#ede9fe;border-color:#c4b5fd;color:#5b21b6}
.tda-days .tda-b{flex:0 0 calc((100% - 18px)/4);min-width:0;font-size:12px;padding:8px 2px}
.tda-sub{margin-top:8px;padding:10px;border-radius:10px;background:#f7f7f5}
.tda-sub .tda-lb{margin-top:0}
.tda input[type=date]{width:100%;box-sizing:border-box;padding:8px 10px;border:1.5px solid #e2e2e2;border-radius:8px;font-size:14px;font-family:inherit;background:#fff}
.tda-pair{display:flex;gap:8px;margin-top:8px}
.tda-pair>div{flex:1}
.tda-pair label{font-size:10px;color:#999;display:block;margin-bottom:3px}
.tda-when{font-size:12px;color:#4f6ef7;font-weight:600;margin-top:10px;min-height:16px}
.tda-foot{display:flex;gap:8px;margin-top:16px}
.tda-foot button{flex:1;padding:11px;border-radius:9px;font-size:14px;font-weight:600;font-family:inherit;cursor:pointer}
.tda-cancel{background:#fff;border:1.5px solid #e0e0e0;color:#555}
.tda-ok{background:#1a1a1a;border:none;color:#fff}
.tda-ok:disabled{opacity:.5;cursor:default}
.tda-err{font-size:12px;color:#dc2626;margin-top:8px;min-height:0}
`;

  let st = { type: 'todo', when: 'tomorrow', date: '', start: '', end: '', onAdded: null };
  let root;

  function build() {
    if (root) return;
    const s = document.createElement('style'); s.textContent = css; document.head.appendChild(s);
    root = document.createElement('div');
    root.className = 'tda-ov';
    root.innerHTML = `
<div class="tda" role="dialog" aria-modal="true">
  <h3>✅ TODO を追加</h3>
  <div class="tda-note">「いつ」で選んだ日のトップと日次スケジュールに出ます。TODO は完了にするまで毎日残ります。</div>
  <div id="tdaTypeRow">
    <span class="tda-lb">種類</span>
    <div class="tda-btns">
      <button type="button" class="tda-b t-todo" data-type="todo">☐ TODO</button>
      <button type="button" class="tda-b t-memo" data-type="memo">📋 メモ</button>
    </div>
  </div>
  <span class="tda-lb">内容</span>
  <textarea id="tdaText" rows="2" placeholder="やることを入力…"></textarea>
  <span class="tda-lb">いつ</span>
  <div class="tda-btns">
    <button type="button" class="tda-b" data-when="today">今日</button>
    <button type="button" class="tda-b" data-when="tomorrow">明日</button>
    <button type="button" class="tda-b" data-when="pick">別の日</button>
    <button type="button" class="tda-b w-medium" data-when="medium">📅 中期</button>
  </div>
  <div class="tda-sub" id="tdaPick" style="display:none">
    <div class="tda-btns tda-days" id="tdaDays"></div>
    <div style="margin-top:8px"><label class="tda-lb" style="font-size:11px;margin:0 0 4px">それより先の日</label><input type="date" id="tdaDate"></div>
  </div>
  <div class="tda-sub" id="tdaMedium" style="display:none">
    <span class="tda-lb" style="font-size:11px">出しておく期間（開始日から）</span>
    <div class="tda-btns" id="tdaSpan">
      <button type="button" class="tda-b" data-span="7">1週間</button>
      <button type="button" class="tda-b" data-span="14">2週間</button>
      <button type="button" class="tda-b" data-span="30">1か月</button>
    </div>
    <div class="tda-pair">
      <div><label>開始日</label><input type="date" id="tdaStart"></div>
      <div><label>終了日</label><input type="date" id="tdaEnd"></div>
    </div>
  </div>
  <div class="tda-when" id="tdaWhen"></div>
  <div class="tda-err" id="tdaErr"></div>
  <div class="tda-foot">
    <button type="button" class="tda-cancel" id="tdaCancel">キャンセル</button>
    <button type="button" class="tda-ok" id="tdaOk">追加</button>
  </div>
</div>`;
    document.body.appendChild(root);
    root.addEventListener('click', e => { if (e.target === root) close(); });
    root.querySelectorAll('[data-type]').forEach(b => b.onclick = () => { st.type = b.dataset.type; paint(); });
    root.querySelectorAll('[data-when]').forEach(b => b.onclick = () => {
      st.when = b.dataset.when;
      if (st.when === 'pick' && !st.date) st.date = addDays(today(), 2);
      paint();
    });
    root.querySelectorAll('[data-span]').forEach(b => b.onclick = () => {
      st.start = st.start || today(); st.end = addDays(st.start, +b.dataset.span); paint();
    });
    q('#tdaDate').onchange = e => { st.date = e.target.value; paint(); };
    q('#tdaStart').onchange = e => { st.start = e.target.value; paint(); };
    q('#tdaEnd').onchange = e => { st.end = e.target.value; paint(); };
    q('#tdaCancel').onclick = close;
    q('#tdaOk').onclick = submit;
    q('#tdaText').addEventListener('keydown', e => { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && root.classList.contains('open')) close(); });
  }
  const q = s => root.querySelector(s);

  function paint() {
    const isMed = st.when === 'medium';
    q('#tdaTypeRow').style.display = isMed ? 'none' : '';
    root.querySelectorAll('[data-type]').forEach(b => b.classList.toggle('on', b.dataset.type === st.type));
    root.querySelectorAll('[data-when]').forEach(b => b.classList.toggle('on', b.dataset.when === st.when));
    q('#tdaPick').style.display = st.when === 'pick' ? '' : 'none';
    q('#tdaMedium').style.display = isMed ? '' : 'none';
    // 明後日から1週間ぶんの日付ボタン
    const t = today(), days = [];
    for (let i = 2; i <= 9; i++) days.push(addDays(t, i));
    q('#tdaDays').innerHTML = days.map(d => `<button type="button" class="tda-b${d === st.date ? ' on' : ''}" data-day="${d}">${md(d)}</button>`).join('');
    q('#tdaDays').querySelectorAll('[data-day]').forEach(b => b.onclick = () => { st.date = b.dataset.day; paint(); });
    q('#tdaDate').value = st.when === 'pick' ? st.date : '';
    q('#tdaStart').value = st.start; q('#tdaEnd').value = st.end;
    root.querySelectorAll('[data-span]').forEach(b => b.classList.toggle('on', st.start && st.end === addDays(st.start, +b.dataset.span)));
    q('#tdaWhen').textContent = isMed
      ? (st.start && st.end ? `${md(st.start)} 〜 ${md(st.end)} のあいだ毎日出ます` : '')
      : `${md(target())} に追加します`;
  }
  function target() {
    if (st.when === 'today') return today();
    if (st.when === 'tomorrow') return addDays(today(), 1);
    return st.date || addDays(today(), 2);
  }

  function open(opts) {
    build();
    opts = opts || {};
    const t = today(), tm = addDays(t, 1), d = opts.date || '';
    st = { type: 'todo', when: 'tomorrow', date: '', start: t, end: addDays(t, 7), onAdded: opts.onAdded || null };
    if (d === t) st.when = 'today';
    else if (d === tm) st.when = 'tomorrow';
    else if (d) { st.when = 'pick'; st.date = d; st.start = d; st.end = addDays(d, 7); }
    q('#tdaText').value = ''; q('#tdaErr').textContent = ''; q('#tdaOk').disabled = false;
    paint();
    root.classList.add('open');
    setTimeout(() => q('#tdaText').focus(), 60);
  }
  function close() { root && root.classList.remove('open'); }

  async function submit() {
    const text = q('#tdaText').value.trim();
    const err = q('#tdaErr');
    if (!text) { err.textContent = '内容を入力してください'; return; }
    const payload = { text };
    if (st.when === 'medium') {
      if (!st.start || !st.end) { err.textContent = '開始日と終了日を選んでください'; return; }
      if (st.end < st.start) { err.textContent = '終了日が開始日より前になっています'; return; }
      payload.type = 'medium'; payload.start_date = st.start; payload.end_date = st.end;
    } else {
      payload.type = st.type; payload.date = target();
    }
    q('#tdaOk').disabled = true; err.textContent = '';
    try {
      const r = await fetch(`${SERVER}/api/handover/add`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload), signal: AbortSignal.timeout(8000)
      });
      const d = await r.json().catch(() => ({}));
      if (!r.ok || d.ok === false) throw new Error(d.error || 'サーバーエラー');
      close();
      if (st.onAdded) await st.onAdded(d.item);
    } catch (e) {
      err.textContent = '追加できませんでした（' + e.message + '）';
      q('#tdaOk').disabled = false;
    }
  }

  window.openTodoAdd = open;
})();
