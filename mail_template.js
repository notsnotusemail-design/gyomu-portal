/* ============================================================
   メール定型文 共通スクリプト
   - 差し込み変数の置換
   - 他ツールから呼び出す「定型文ピッカー」モーダル
   使い方: <script src="/mail_template.js"></script>
           MailTpl.openPicker({ name:'山内様', no:'302', invoiceName:'' })
   ============================================================ */
(function () {
  const SERVER = window.location.origin;

  // 差し込み変数の一覧（定型文ツールのボタンもここから生成する）
  const VARS = [
    { key: '顧客名',     desc: '顧客情報の「名前」' },
    { key: 'お客様No',   desc: 'お客様No.' },
    { key: '請求書宛名', desc: '未設定なら顧客名が入る' },
    { key: '案件名',     desc: '送信時に手入力' },
    { key: '請求月',     desc: '送信時に手入力（例: 8月）' },
    { key: '今日の日付', desc: '例: 2026年9月18日' },
    { key: '担当者名',   desc: '自分の名前（設定で変更可）' },
  ];

  function todayJa() {
    const d = new Date();
    return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
  }

  function senderName() {
    return localStorage.getItem('mailtpl_sender') || '野津 欧';
  }

  /* 本文・件名の {{変数}} を実データに置き換える。
     対応表に無い変数、値が空の変数はそのまま残す
     （空のままコピーして送ってしまわないよう、警告で拾えるようにする）。 */
  function fill(text, ctx) {
    ctx = ctx || {};
    const map = {
      '顧客名':     ctx.customerName || '',
      'お客様No':   ctx.customerNo   || '',
      '請求書宛名': ctx.invoiceName  || ctx.customerName || '',
      '案件名':     ctx.caseName     || '',
      '請求月':     ctx.billingMonth || '',
      '今日の日付': ctx.today || todayJa(),
      '担当者名':   ctx.sender || senderName(),
    };
    return String(text || '').replace(/\{\{\s*([^}]+?)\s*\}\}/g,
      (m, k) => (map[k] ? map[k] : m));
  }

  // 未入力のまま残っている変数を拾う（コピー前の警告用）
  function missingVars(text) {
    const found = String(text || '').match(/\{\{\s*[^}]+?\s*\}\}/g) || [];
    return [...new Set(found)];
  }

  async function loadTemplates() {
    const r = await fetch(SERVER + '/api/mail-templates', { signal: AbortSignal.timeout(15000) });
    const d = await r.json();
    if (!d.ok) throw new Error(d.error || '取得失敗');
    return d.templates;
  }

  /* クリップボードコピー。https以外でも動くようフォールバックを持つ */
  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.cssText = 'position:fixed;top:-2000px;left:-2000px;';
      document.body.appendChild(ta);
      ta.select();
      let ok = false;
      try { ok = document.execCommand('copy'); } catch (_) {}
      document.body.removeChild(ta);
      return ok;
    }
  }

  /* コピーが拒否された環境用。該当テキストを選択状態にして手動コピーできるようにする */
  function selectElementText(el) {
    if (!el) return;
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function esc(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function injectStyle() {
    if (document.getElementById('mailtpl-style')) return;
    const st = document.createElement('style');
    st.id = 'mailtpl-style';
    st.textContent = `
      .mtp-overlay { position:fixed; inset:0; background:rgba(0,0,0,0.42); z-index:9000;
        display:flex; align-items:center; justify-content:center; padding:20px; }
      .mtp-modal { background:#fff; border-radius:14px; width:100%; max-width:720px;
        max-height:88vh; display:flex; flex-direction:column; overflow:hidden;
        box-shadow:0 12px 40px rgba(0,0,0,0.22);
        font-family:-apple-system,BlinkMacSystemFont,"Hiragino Sans","Yu Gothic",sans-serif; }
      .mtp-head { display:flex; align-items:center; gap:10px; padding:14px 18px;
        border-bottom:1px solid #eee; }
      .mtp-title { font-size:15px; font-weight:700; }
      .mtp-sub { font-size:12px; color:#999; }
      .mtp-x { margin-left:auto; border:none; background:#f1f1f1; width:28px; height:28px;
        border-radius:8px; cursor:pointer; font-size:14px; color:#666; }
      .mtp-x:hover { background:#e4e4e4; }
      .mtp-body { padding:16px 18px; overflow-y:auto; }
      .mtp-row { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:12px; }
      .mtp-sel, .mtp-in { padding:8px 10px; border:1.5px solid #e2e2e2; border-radius:8px;
        font-size:13px; font-family:inherit; background:#fafafa; }
      .mtp-sel:focus, .mtp-in:focus { outline:none; border-color:#4f6ef7; background:#fff; }
      .mtp-sel { flex:1; min-width:200px; cursor:pointer; }
      .mtp-in { width:130px; }
      .mtp-label { font-size:11px; font-weight:600; color:#999; margin:12px 0 5px; }
      .mtp-prev { background:#fafafa; border:1.5px solid #eee; border-radius:9px;
        padding:10px 12px; font-size:13px; line-height:1.75; white-space:pre-wrap;
        word-break:break-word; min-height:40px; }
      .mtp-prev.subj { white-space:normal; font-weight:600; min-height:0; }
      .mtp-warn { font-size:11.5px; color:#b45309; background:#fffbeb; border:1px solid #fde68a;
        border-radius:7px; padding:7px 10px; margin-top:8px; }
      .mtp-foot { display:flex; gap:8px; padding:12px 18px; border-top:1px solid #eee;
        background:#fafafa; flex-wrap:wrap; }
      .mtp-btn { padding:8px 14px; border-radius:8px; font-size:13px; font-weight:600;
        cursor:pointer; font-family:inherit; border:1.5px solid #e2e2e2; background:#fff;
        color:#555; transition:all 0.15s; }
      .mtp-btn:hover { background:#f2f2f2; }
      .mtp-btn.primary { background:#4f6ef7; border-color:#4f6ef7; color:#fff; }
      .mtp-btn.primary:hover { background:#3b5ce0; }
      .mtp-btn.ghost { margin-left:auto; }
      .mtp-empty { text-align:center; color:#aaa; font-size:13px; padding:30px 10px; }
      .mtp-empty a { color:#4f6ef7; }
    `;
    document.head.appendChild(st);
  }

  /* 顧客情報ツールなどから呼ぶ定型文ピッカー。
     customer は /api/customers-all の1件（name/no/invoiceName）をそのまま渡せる */
  async function openPicker(customer) {
    injectStyle();
    customer = customer || {};
    const ov = document.createElement('div');
    ov.className = 'mtp-overlay';
    ov.innerHTML = `
      <div class="mtp-modal">
        <div class="mtp-head">
          <div>
            <div class="mtp-title">✉️ メール定型文</div>
            <div class="mtp-sub">${esc(customer.no || '')} ${esc(customer.name || '宛先未選択')}</div>
          </div>
          <button class="mtp-x" title="閉じる">✕</button>
        </div>
        <div class="mtp-body"><div class="mtp-empty">読み込み中…</div></div>
      </div>`;
    document.body.appendChild(ov);
    const close = () => ov.remove();
    ov.querySelector('.mtp-x').onclick = close;
    ov.onclick = (e) => { if (e.target === ov) close(); };
    document.addEventListener('keydown', function onEsc(e) {
      if (e.key === 'Escape') { close(); document.removeEventListener('keydown', onEsc); }
    });

    let templates = [];
    try {
      templates = (await loadTemplates()).filter(t => t.enabled);
    } catch (e) {
      ov.querySelector('.mtp-body').innerHTML =
        `<div class="mtp-empty">⚠️ 取得に失敗しました: ${esc(e.message)}</div>`;
      return;
    }
    if (!templates.length) {
      ov.querySelector('.mtp-body').innerHTML =
        `<div class="mtp-empty">定型文がまだ登録されていません<br>
         <a href="/メール定型文ツール.html">メール定型文ツールで登録する →</a></div>`;
      return;
    }

    const modal = ov.querySelector('.mtp-modal');
    modal.querySelector('.mtp-body').innerHTML = `
      <div class="mtp-row">
        <select class="mtp-sel" id="mtpSel">
          ${templates.map((t, i) => `<option value="${i}">${esc(t.category ? '[' + t.category + '] ' : '')}${esc(t.name)}</option>`).join('')}
        </select>
        <input class="mtp-in" id="mtpCase"  placeholder="案件名">
        <input class="mtp-in" id="mtpMonth" placeholder="請求月 例:8月">
      </div>
      <div class="mtp-label">件名</div>
      <div class="mtp-prev subj" id="mtpSubj"></div>
      <div class="mtp-label">本文</div>
      <div class="mtp-prev" id="mtpBody"></div>
      <div id="mtpWarn"></div>`;
    modal.insertAdjacentHTML('beforeend', `
      <div class="mtp-foot">
        <button class="mtp-btn primary" id="mtpCopyBody">本文をコピー</button>
        <button class="mtp-btn" id="mtpCopySubj">件名をコピー</button>
        <button class="mtp-btn" id="mtpMail">メールを開く</button>
        <button class="mtp-btn ghost" id="mtpEdit">定型文を編集</button>
      </div>`);

    const $ = (id) => modal.querySelector('#' + id);
    const ctx = () => ({
      customerName: customer.name || customer.pageName || '',
      customerNo:   customer.no || '',
      invoiceName:  customer.invoiceName || '',
      caseName:     $('mtpCase').value.trim(),
      billingMonth: $('mtpMonth').value.trim(),
    });
    let cur = { subject: '', body: '' };

    function render() {
      const t = templates[Number($('mtpSel').value)] || {};
      cur = { subject: fill(t.subject, ctx()), body: fill(t.body, ctx()) };
      $('mtpSubj').textContent = cur.subject || '（件名なし）';
      $('mtpBody').textContent = cur.body || '（本文なし）';
      const miss = missingVars(cur.subject + '\n' + cur.body);
      $('mtpWarn').innerHTML = miss.length
        ? `<div class="mtp-warn">未入力の変数があります: ${esc(miss.join(' '))}</div>` : '';
    }
    $('mtpSel').onchange = render;
    $('mtpCase').oninput = render;
    $('mtpMonth').oninput = render;
    render();

    async function flash(btn, text) {
      const old = btn.textContent;
      btn.textContent = text;
      setTimeout(() => { btn.textContent = old; }, 1300);
    }
    async function copyOr(btn, text, previewEl) {
      if (await copyText(text)) { flash(btn, '✅ コピーしました'); return; }
      selectElementText(previewEl);   // 選択しておけば ⌘C で拾える
      flash(btn, '⌘Cでコピーしてください');
    }
    $('mtpCopyBody').onclick = (e) => copyOr(e.target, cur.body, $('mtpBody'));
    $('mtpCopySubj').onclick = (e) => copyOr(e.target, cur.subject, $('mtpSubj'));
    $('mtpMail').onclick = () => {
      window.location.href =
        `mailto:?subject=${encodeURIComponent(cur.subject)}&body=${encodeURIComponent(cur.body)}`;
    };
    $('mtpEdit').onclick = () => { window.location.href = '/メール定型文ツール.html'; };
  }

  window.MailTpl = { VARS, fill, missingVars, loadTemplates, copyText,
                     selectElementText, openPicker, todayJa, senderName };
})();
