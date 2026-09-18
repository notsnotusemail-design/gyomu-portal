/* ============================================================
   メール定型文 共通スクリプト
   - 差し込み変数の置換
   - 送信時の入力欄（日付ダイヤル・自由記載欄）の生成
   - 他ツールから呼び出す「定型文ピッカー」モーダル
   使い方: <script src="/mail_template.js"></script>
           MailTpl.openPicker({ name:'山内様', no:'302', invoiceName:'' })

   変数の考え方:
     顧客名などの「自動で埋まる変数」以外は、本文に {{◯◯}} と書いた時点で
     送信時の入力欄になる。テンプレート側を固定したまま、送るときに文章を
     整えられるようにするため。名前が「日付」で終わるものはダイヤルになる。
   ============================================================ */
(function () {
  const SERVER = window.location.origin;

  // 顧客データ・システムから自動で埋まる変数
  const AUTO_VARS = [
    { key: '顧客名',     desc: '顧客情報の「名前」' },
    { key: 'お客様No',   desc: 'お客様No.' },
    { key: '請求書宛名', desc: '未設定なら顧客名が入る' },
    { key: '今日の日付', desc: '送る日（自動）' },
    { key: '担当者名',   desc: '自分の名前（設定で変更可）' },
  ];
  const AUTO_KEYS = AUTO_VARS.map(v => v.key);

  // 送信時に手入力する変数の見本。これ以外の名前を書いても入力欄になる
  const MANUAL_VARS = [
    { key: '日付',     desc: '任意の日付をダイヤルで選ぶ' },
    { key: '案件名',   desc: '送信時に入力' },
    { key: '請求月',   desc: '送信時に入力' },
    { key: '自由記載', desc: '好きな文章を送信時に入力' },
  ];

  const VARS = AUTO_VARS.concat(MANUAL_VARS);   // 変数ボタン用

  const DATE_FORMATS = [
    { id: 'ymd',   label: '2026年9月18日' },
    { id: 'ymd_w', label: '2026年9月18日(木)' },
    { id: 'md',    label: '9月18日' },
    { id: 'md_w',  label: '9月18日(木)' },
    { id: 'slash', label: '9/18' },
    { id: 'iso',   label: '2026/09/18' },
  ];
  const WD = ['日', '月', '火', '水', '木', '金', '土'];

  function formatDate(y, m, d, fmt) {
    const w = WD[new Date(y, m - 1, d).getDay()];
    switch (fmt) {
      case 'ymd_w': return `${y}年${m}月${d}日(${w})`;
      case 'md':    return `${m}月${d}日`;
      case 'md_w':  return `${m}月${d}日(${w})`;
      case 'slash': return `${m}/${d}`;
      case 'iso':   return `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}`;
      default:      return `${y}年${m}月${d}日`;
    }
  }

  function todayJa() {
    const d = new Date();
    return formatDate(d.getFullYear(), d.getMonth() + 1, d.getDate(), 'ymd');
  }

  function senderName() {
    return localStorage.getItem('mailtpl_sender') || '野津 欧';
  }

  function dateFormat() {
    return localStorage.getItem('mailtpl_datefmt') || 'ymd';
  }

  /* 本文・件名に出てくる {{変数}} の名前を、出てきた順に返す */
  function placeholders(text) {
    const names = [];
    const re = /\{\{\s*([^}]+?)\s*\}\}/g;
    let m;
    while ((m = re.exec(String(text || ''))) !== null) {
      if (!names.includes(m[1])) names.push(m[1]);
    }
    return names;
  }

  /* 送信時に手入力が必要な変数（＝自動で埋まらないもの）。
     名前が「日付」で終わるものはダイヤルで入れる */
  function manualVars(text) {
    return placeholders(text)
      .filter(k => !AUTO_KEYS.includes(k))
      .map(k => ({ key: k, type: /日付$/.test(k) ? 'date' : 'text' }));
  }

  /* 本文・件名の {{変数}} を実データに置き換える。
     対応表に無い変数、値が空の変数はそのまま残す
     （空のままコピーして送ってしまわないよう、警告で拾えるようにする）。 */
  function fill(text, ctx) {
    ctx = ctx || {};
    const vals = ctx.values || {};
    const map = {
      '顧客名':     ctx.customerName || '',
      'お客様No':   ctx.customerNo   || '',
      '請求書宛名': ctx.invoiceName  || ctx.customerName || '',
      '今日の日付': ctx.today || todayJa(),
      '担当者名':   ctx.sender || senderName(),
    };
    return String(text || '').replace(/\{\{\s*([^}]+?)\s*\}\}/g,
      (m, k) => map[k] || vals[k] || m);
  }

  // 未入力のまま残っている変数を拾う（コピー前の警告用）
  function missingVars(text) {
    const found = String(text || '').match(/\{\{\s*[^}]+?\s*\}\}/g) || [];
    return [...new Set(found)];
  }

  /* ---------- 案件名の候補 ----------
     請求書の「詳細」に載るのと同じ文字列を候補にする。
     案件名（Notionの指定案件ファイル名）が空のときだけ 備考 → 案件番号 で代用する。
     サーバーの customer フィルターは部分一致なので（302 が 3302 も拾う）、
     こちら側でお客様No.の完全一致に絞り直す。 */
  const _caseCache = {};

  async function loadCases(customerNo) {
    const no = String(customerNo || '').trim();
    if (!no) return [];
    if (_caseCache[no]) return _caseCache[no];
    try {
      const r = await fetch(SERVER + '/api/cases/list?customer=' + encodeURIComponent(no),
                            { signal: AbortSignal.timeout(20000) });
      const d = await r.json();
      if (!d.ok) return [];
      const seen = new Set();
      const list = (d.cases || [])
        .filter(c => String(c.customer || '').trim() === no)
        .sort((a, b) => (b.date || '').localeCompare(a.date || ''))
        .map(c => ({
          label: (c.filename || '').trim() || (c.note || '').trim() || (c.number || '').trim(),
          number: c.number || '', date: c.date || '',
        }))
        .filter(c => c.label && !seen.has(c.label) && seen.add(c.label))
        .slice(0, 60);
      _caseCache[no] = list;
      return list;
    } catch (e) {
      return [];   // 案件が引けなくても手入力はできる
    }
  }

  const isCaseVar = (key) => /案件名$/.test(key) || key === '案件';

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

  /* ---------- 送信時の入力欄 ----------
     テンプレートが使っている変数だけを見て、必要な入力欄を組み立てる。
     container: 差し込む要素 / text: 件名＋本文 / state: 値の入れ物 / onChange: 再描画 */
  function renderVarInputs(container, text, state, onChange) {
    injectStyle();
    state.values = state.values || {};
    state.dates  = state.dates  || {};
    const vars = manualVars(text);

    if (!vars.length) {
      container.innerHTML = '<div class="mtp-noinput">この定型文に入力が必要な項目はありません</div>';
      return;
    }

    container.innerHTML = vars.map(v => {
      if (v.type === 'date') {
        const d = state.dates[v.key] || (state.dates[v.key] = defaultDate());
        return `
          <div class="mtp-field">
            <label class="mtp-flabel">${esc(v.key)}</label>
            <div class="mtp-dial" data-var="${esc(v.key)}">
              <input class="mtp-dial-num" type="number" data-part="y" min="2020" max="2035" value="${d.y}" style="width:60px">
              <span class="mtp-dial-sep">年</span>
              <input class="mtp-dial-num" type="number" data-part="m" min="0" max="13" value="${d.m}" style="width:40px">
              <span class="mtp-dial-sep">月</span>
              <input class="mtp-dial-num" type="number" data-part="d" min="0" max="32" value="${d.d}" style="width:40px">
              <span class="mtp-dial-sep">日</span>
              <button class="mtp-dial-today" type="button" title="今日に戻す">今日</button>
            </div>
            <select class="mtp-fmt" data-var="${esc(v.key)}">
              ${DATE_FORMATS.map(f => `<option value="${f.id}"${f.id === dateFormat() ? ' selected' : ''}>${esc(f.label)}</option>`).join('')}
            </select>
          </div>`;
      }
      if (isCaseVar(v.key)) {
        const cases = state.cases || [];
        return `
          <div class="mtp-field">
            <label class="mtp-flabel">${esc(v.key)}</label>
            ${comboHTML(v.key, state.values[v.key], cases,
                        cases.length ? '候補から選ぶか直接入力' : '案件名を入力')}
          </div>`;
      }
      const isLong = /自由記載|本文|メモ|備考|内容/.test(v.key);
      return `
        <div class="mtp-field">
          <label class="mtp-flabel">${esc(v.key)}</label>
          ${isLong
            ? `<textarea class="mtp-fin mtp-ftext" data-var="${esc(v.key)}" rows="3"
                 placeholder="${esc(v.key)}を入力（改行できます）">${esc(state.values[v.key] || '')}</textarea>`
            : `<input class="mtp-fin" data-var="${esc(v.key)}" value="${esc(state.values[v.key] || '')}"
                 placeholder="${esc(v.key)}を入力">`}
        </div>`;
    }).join('');

    wireCombo(container, state.cases || [], onChange);

    // テキスト欄
    container.querySelectorAll('.mtp-fin').forEach(el => {
      el.addEventListener('input', () => {
        state.values[el.dataset.var] = el.value;
        onChange();
      });
    });

    // 日付ダイヤル
    container.querySelectorAll('.mtp-dial').forEach(dial => {
      const key = dial.dataset.var;
      const fmtSel = container.querySelector(`.mtp-fmt[data-var="${CSS.escape(key)}"]`);
      const apply = () => {
        const d = state.dates[key];
        state.values[key] = formatDate(d.y, d.m, d.d, fmtSel.value);
        onChange();
      };
      dial.querySelectorAll('.mtp-dial-num').forEach(inp => {
        const handler = () => {
          const d = state.dates[key];
          const part = inp.dataset.part;
          d[part] = parseInt(inp.value, 10);
          normalizeDate(d);
          dial.querySelector('[data-part=y]').value = d.y;
          dial.querySelector('[data-part=m]').value = d.m;
          dial.querySelector('[data-part=d]').value = d.d;
          apply();
        };
        inp.addEventListener('change', handler);
        inp.addEventListener('input',  handler);
      });
      dial.querySelector('.mtp-dial-today').addEventListener('click', () => {
        state.dates[key] = defaultDate();
        const d = state.dates[key];
        dial.querySelector('[data-part=y]').value = d.y;
        dial.querySelector('[data-part=m]').value = d.m;
        dial.querySelector('[data-part=d]').value = d.d;
        apply();
      });
      fmtSel.addEventListener('change', () => {
        localStorage.setItem('mailtpl_datefmt', fmtSel.value);
        apply();
      });
      apply();   // 初期値をすぐ反映する
    });
  }

  /* ---------- 選択と自由入力を兼ねるコントロール ----------
     候補は ◀▶ のダイヤルで送れるが、欄に直接打ち込んでもいい。
     候補が無いときはただのテキスト欄として振る舞う。 */
  let _comboSeq = 0;

  function comboHTML(varKey, value, candidates, placeholder) {
    const id = 'mtpdl' + (++_comboSeq);
    const has = candidates && candidates.length;
    return `
      <div class="mtp-combo" data-var="${esc(varKey)}" data-list="${id}">
        ${has ? '<button class="mtp-combo-arw" type="button" data-step="-1" title="前の候補">◀</button>' : ''}
        <input class="mtp-fin mtp-combo-in" data-var="${esc(varKey)}" value="${esc(value || '')}"
               placeholder="${esc(placeholder || '')}" ${has ? `list="${id}"` : ''}>
        ${has ? '<button class="mtp-combo-arw" type="button" data-step="1" title="次の候補">▶</button>' : ''}
        ${has ? `<span class="mtp-combo-count">0/${candidates.length}</span>` : ''}
      </div>
      ${has ? `<datalist id="${id}">${candidates.map(c =>
          `<option value="${esc(c.label !== undefined ? c.label : c)}">${esc(c.hint || '')}</option>`).join('')}</datalist>` : ''}`;
  }

  /* comboHTML で作った要素にダイヤル操作を付ける */
  function wireCombo(scope, candidates, onChange) {
    scope.querySelectorAll('.mtp-combo').forEach(combo => {
      const input = combo.querySelector('.mtp-combo-in');
      const count = combo.querySelector('.mtp-combo-count');
      const list  = (typeof candidates === 'function' ? candidates(combo.dataset.var) : candidates) || [];
      const labels = list.map(c => (c.label !== undefined ? c.label : c));
      const sync = () => {
        if (!count) return;
        const i = labels.indexOf(input.value);
        count.textContent = `${i < 0 ? 0 : i + 1}/${labels.length}`;
      };
      combo.querySelectorAll('.mtp-combo-arw').forEach(btn => {
        btn.addEventListener('click', () => {
          if (!labels.length) return;
          const step = Number(btn.dataset.step);
          let i = labels.indexOf(input.value);
          i = i < 0 ? (step > 0 ? 0 : labels.length - 1)
                    : (i + step + labels.length) % labels.length;
          input.value = labels[i];
          sync();
          input.dispatchEvent(new Event('input'));
        });
      });
      input.addEventListener('input', sync);
      sync();
    });
  }

  function defaultDate() {
    const n = new Date();
    return { y: n.getFullYear(), m: n.getMonth() + 1, d: n.getDate() };
  }

  /* ダイヤルを回して 12月の次・1月の前・月末を越えたときに繰り上げ／繰り下げる */
  function normalizeDate(d) {
    if (!Number.isFinite(d.y)) d.y = new Date().getFullYear();
    if (!Number.isFinite(d.m)) d.m = 1;
    if (!Number.isFinite(d.d)) d.d = 1;
    if (d.m > 12) { d.m = 1;  d.y++; }
    if (d.m < 1)  { d.m = 12; d.y--; }
    const last = new Date(d.y, d.m, 0).getDate();
    if (d.d > last) { d.d = 1; d.m++; if (d.m > 12) { d.m = 1; d.y++; } }
    else if (d.d < 1) {
      d.m--; if (d.m < 1) { d.m = 12; d.y--; }
      d.d = new Date(d.y, d.m, 0).getDate();
    }
    d.y = Math.min(2035, Math.max(2020, d.y));
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
      .mtp-sel { flex:1; min-width:200px; cursor:pointer; padding:8px 10px;
        border:1.5px solid #e2e2e2; border-radius:8px; font-size:13px;
        font-family:inherit; background:#fafafa; }
      .mtp-sel:focus { outline:none; border-color:#4f6ef7; background:#fff; }
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

      /* 送信時の入力欄 */
      .mtp-inputs { display:grid; gap:10px; }
      .mtp-field { display:flex; align-items:flex-start; gap:8px; flex-wrap:wrap; }
      .mtp-flabel { font-size:11px; font-weight:600; color:#999; width:76px;
        flex-shrink:0; padding-top:9px; word-break:break-all; }
      .mtp-fin { flex:1; min-width:180px; padding:8px 11px; border:1.5px solid #e2e2e2;
        border-radius:8px; font-size:13px; font-family:inherit; background:#fafafa;
        color:#1a1a1a; }
      .mtp-fin:focus { outline:none; border-color:#4f6ef7; background:#fff; }
      .mtp-ftext { line-height:1.7; resize:vertical; }
      .mtp-noinput { font-size:12px; color:#bbb; padding:4px 0; }

      /* 日付ダイヤル（請求書ツールの月ダイヤルと同じ操作感） */
      .mtp-dial { display:flex; align-items:center; gap:4px; background:#fff;
        border:1.5px solid #d1d5db; border-radius:10px; padding:4px 10px; user-select:none; }
      .mtp-dial-num { font-size:15px; font-weight:700; color:#111827; border:none;
        outline:none; background:transparent; font-family:inherit; text-align:center;
        -moz-appearance:textfield; }
      .mtp-dial-num::-webkit-inner-spin-button,
      .mtp-dial-num::-webkit-outer-spin-button { opacity:1; height:22px; cursor:pointer; }
      .mtp-dial-num:focus { color:#4f6ef7; }
      .mtp-dial-sep { font-size:13px; color:#6b7280; }
      .mtp-dial-today { border:1.5px solid #e2e2e2; background:#fafafa; border-radius:7px;
        font-size:11px; color:#666; padding:3px 8px; cursor:pointer; font-family:inherit;
        margin-left:4px; }
      .mtp-dial-today:hover { background:#eef1ff; border-color:#4f6ef7; color:#4f6ef7; }
      .mtp-fmt { padding:7px 8px; border:1.5px solid #e2e2e2; border-radius:8px;
        font-size:12px; font-family:inherit; background:#fafafa; cursor:pointer; color:#555; }
      .mtp-fmt:focus { outline:none; border-color:#4f6ef7; background:#fff; }

      /* 選択と自由入力を兼ねる欄（案件名・顧客・カテゴリ） */
      .mtp-combo { display:flex; align-items:center; gap:5px; flex:1; min-width:200px; }
      .mtp-combo-in { flex:1; min-width:120px; }
      .mtp-combo-arw { border:1.5px solid #d1d5db; background:#fff; border-radius:8px;
        width:28px; height:33px; cursor:pointer; font-size:11px; color:#6b7280;
        font-family:inherit; flex-shrink:0; }
      .mtp-combo-arw:hover { background:#eef1ff; border-color:#4f6ef7; color:#4f6ef7; }
      .mtp-combo-count { font-size:11px; color:#bbb; flex-shrink:0; min-width:32px; }
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
      </div>
      <div class="mtp-inputs" id="mtpInputs"></div>
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
    const state = { values: {}, dates: {}, cases: [] };
    let cur = { subject: '', body: '' };

    // その顧客の案件を案件名の候補にする（引けなくても手入力はできる）
    loadCases(customer.no).then(list => {
      state.cases = list;
      rebuild();
    });

    const ctx = () => ({
      customerName: customer.name || customer.pageName || '',
      customerNo:   customer.no || '',
      invoiceName:  customer.invoiceName || '',
      values:       state.values,
    });

    function render() {
      const t = templates[Number($('mtpSel').value)] || {};
      cur = { subject: fill(t.subject, ctx()), body: fill(t.body, ctx()) };
      $('mtpSubj').textContent = cur.subject || '（件名なし）';
      $('mtpBody').textContent = cur.body || '（本文なし）';
      const miss = missingVars(cur.subject + '\n' + cur.body);
      $('mtpWarn').innerHTML = miss.length
        ? `<div class="mtp-warn">未入力の項目があります: ${esc(miss.join(' '))}</div>` : '';
    }
    function rebuild() {
      const t = templates[Number($('mtpSel').value)] || {};
      renderVarInputs($('mtpInputs'), (t.subject || '') + '\n' + (t.body || ''), state, render);
      render();
    }
    $('mtpSel').onchange = rebuild;
    rebuild();

    function flash(btn, text) {
      const old = btn.textContent;
      btn.textContent = text;
      setTimeout(() => { btn.textContent = old; }, 1400);
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

  window.MailTpl = {
    VARS, AUTO_VARS, MANUAL_VARS, AUTO_KEYS, DATE_FORMATS,
    fill, missingVars, placeholders, manualVars, renderVarInputs,
    formatDate, todayJa, senderName, dateFormat,
    loadTemplates, loadCases, isCaseVar, comboHTML, wireCombo,
    copyText, selectElementText, injectStyle, openPicker,
  };
})();
