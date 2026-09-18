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
    { key: '今日の日付', desc: '送る日（自動・年あり）' },
    { key: '今日の月日', desc: '送る日（自動・月日だけ）' },
    { key: '担当者名',   desc: '自分の名前（設定で変更可）' },
  ];
  const AUTO_KEYS = AUTO_VARS.map(v => v.key);

  // 送信時に手入力する変数の見本。これ以外の名前を書いても入力欄になる
  const MANUAL_VARS = [
    { key: '年月日',   desc: '年・月・日をダイヤルかカレンダーで選ぶ' },
    { key: '月日',     desc: '月・日だけ（年は出ない）' },
    { key: '請求月',   desc: '1〜12月をダイヤルで選ぶ' },
    { key: '案件名',   desc: 'その顧客の案件から選ぶ（直接入力も可）' },
    { key: '自由記載', desc: '好きな文章を送信時に入力' },
  ];

  const VARS = AUTO_VARS.concat(MANUAL_VARS);   // 変数ボタン用

  // 年ありの変数（{{日付}}）と、月日だけの変数（{{月日}}）で選べる形式を分ける
  const DATE_FORMATS = [
    { id: 'ymd',   label: '2026年9月18日' },
    { id: 'ymd_w', label: '2026年9月18日(木)' },
    { id: 'iso',   label: '2026/09/18' },
    { id: 'md',    label: '9月18日' },
    { id: 'md_w',  label: '9月18日(木)' },
    { id: 'slash', label: '9/18' },
  ];
  const MD_FORMAT_IDS = ['md', 'md_w', 'slash'];
  const MD_FORMATS = DATE_FORMATS.filter(f => MD_FORMAT_IDS.includes(f.id));
  const formatsFor  = (kind) => (kind === 'md' ? MD_FORMATS : DATE_FORMATS);
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

  function todayJa(kind) {
    const d = new Date();
    return formatDate(d.getFullYear(), d.getMonth() + 1, d.getDate(), dateFormat(kind));
  }

  function senderName() {
    return localStorage.getItem('mailtpl_sender') || '野津 欧';
  }

  function dateFormat(kind) {
    const saved = localStorage.getItem('mailtpl_datefmt_' + (kind === 'md' ? 'md' : 'ymd'));
    if (saved && formatsFor(kind).some(f => f.id === saved)) return saved;
    return kind === 'md' ? 'md' : 'ymd';
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

  /* 出てきた順に、重複も含めて全部返す */
  function placeholderList(text) {
    const list = [];
    const re = /\{\{\s*([^}]+?)\s*\}\}/g;
    let m;
    while ((m = re.exec(String(text || ''))) !== null) list.push(m[1]);
    return list;
  }

  /* 同じ名前の変数が複数あっても、既定では1つずつ独立して扱う。
     「同じ値にする」を入れた名前だけ、値を共有する。 */
  function instKeyOf(name, inst, linked) {
    return (linked && linked[name]) ? name : `${name}#${inst}`;
  }

  function countNames(text) {
    const n = {};
    placeholderList(text).forEach(k => { n[k] = (n[k] || 0) + 1; });
    return n;
  }

  /* 送信時に手入力が必要な変数（＝自動で埋まらないもの）。
     名前が「日付」で終わるものはダイヤルで入れる */
  function manualVars(text, linked) {
    const total = countNames(text);
    const seen  = {};
    const out   = [];
    placeholderList(text).forEach(name => {
      if (AUTO_KEYS.includes(name)) return;
      const inst = (seen[name] = (seen[name] || 0) + 1);
      const isLinked = !!(linked && linked[name]);
      if (isLinked && inst > 1) return;          // 連動中は入力欄を1つにまとめる
      const kind = slotKind(name);
      const type = (kind === 'ymd' || kind === 'md') ? 'date'
                 : (kind === 'month') ? 'month' : 'text';
      out.push({ key: name, inst, kind, type,
                 vkey: instKeyOf(name, inst, linked),
                 total: total[name], linked: isLinked });
    });
    return out;
  }

  /* 本文・件名の {{変数}} を実データに置き換える。
     対応表に無い変数、値が空の変数はそのまま残す
     （空のままコピーして送ってしまわないよう、警告で拾えるようにする）。 */
  function resolveVar(key, ctx) {
    ctx = ctx || {};
    const auto = {
      '顧客名':     ctx.customerName || '',
      'お客様No':   ctx.customerNo   || '',
      '請求書宛名': ctx.invoiceName  || ctx.customerName || '',
      '今日の日付': ctx.today   || todayJa(),
      '今日の月日': ctx.todayMd || todayJa('md'),
      '担当者名':   ctx.sender || senderName(),
    };
    // 自動で埋まる変数でも、確認ビューで選び直したらそちらを優先する
    // （宛名を「◯◯様」に変える、日付を今日以外にする、といった調整のため）
    return (ctx.values || {})[key] || auto[key] || '';
  }

  /* 何番目の差し込みかを見て値を引く */
  function resolveSlot(name, inst, ctx) {
    ctx = ctx || {};
    const vkey = instKeyOf(name, inst, ctx.linked);
    const v = (ctx.values || {})[vkey];
    if (v) return v;
    return resolveVar(name, Object.assign({}, ctx, { values: {} }));
  }

  // 顧客が変わったら、顧客由来の手直しは捨てる
  const CUSTOMER_VARS = ['顧客名', 'お客様No', '請求書宛名'];
  function clearCustomerOverrides(values) {
    CUSTOMER_VARS.forEach(k => { delete values[k]; });
  }

  function fill(text, ctx) {
    return String(text || '').replace(/\{\{\s*([^}]+?)\s*\}\}/g,
      (m, k) => resolveVar(k, ctx) || m);
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
        .map(c => {
          const label = (c.filename || '').trim() || (c.note || '').trim() || (c.number || '').trim();
          return { label, number: c.number || '', date: c.date || '',
                   hint: [c.date || '', c.number || ''].filter(Boolean).join('  ') };
        })
        .filter(c => c.label && !seen.has(c.label) && seen.add(c.label))
        .slice(0, 60);
      _caseCache[no] = list;
      return list;
    } catch (e) {
      return [];   // 案件が引けなくても手入力はできる
    }
  }

  const isCaseVar = (key) => /案件名$/.test(key) || key === '案件';

  /* ---------- 案件候補の期間しぼり ----------
     過去の案件が多い顧客だと一覧から探すのが大変なので、既定では直近1ヶ月だけを
     ダイヤルに載せる。それ以前は「3ヶ月」「すべて」で広げるか、年月を指定して
     その月の案件だけに絞る。 */
  const CASE_RANGES = [
    { id: '1m',    label: '直近1ヶ月' },
    { id: '3m',    label: '3ヶ月' },
    { id: 'all',   label: 'すべて' },
    { id: 'month', label: '年月で選ぶ' },
  ];

  /* 年月モードの初期値。案件が1件も無ければ今月から始める */
  function newestCaseMonth(cases) {
    const d = (cases || []).map(c => c.date).filter(Boolean).sort().pop();
    if (d) return { y: Number(d.slice(0, 4)), m: Number(d.slice(5, 7)), d: 1 };
    const n = new Date();
    return { y: n.getFullYear(), m: n.getMonth() + 1, d: 1 };
  }

  function monthsAgo(n) {
    const d = new Date();
    d.setMonth(d.getMonth() - n);
    return d.toISOString().slice(0, 10);
  }

  function filterCases(cases, range) {
    const list = cases || [];
    if (!range || range.mode === 'all') return list;
    if (range.mode === 'month') {
      const md = range.month || {};
      if (!md.y) return list;
      const key = `${md.y}-${String(md.m).padStart(2, '0')}`;
      return list.filter(c => (c.date || '').startsWith(key));
    }
    const from = monthsAgo(range.mode === '3m' ? 3 : 1);
    return list.filter(c => (c.date || '') >= from);
  }

  /* 直近に案件が無い顧客で空のダイヤルを出さないよう、
     中身のある一番狭い期間まで自動で広げる */
  function initialRange(cases) {
    for (const mode of ['1m', '3m']) {
      if (filterCases(cases, { mode }).length) return { mode };
    }
    return { mode: 'all' };
  }


  /* ---------- 確認ビュー（プレビュー）の変数スロット ----------
     差し込んだ箇所を span で包み、そこをクリックしたらカレンダーや候補を出す。
     手で打ち替えられたスロットは「ただの文字」に降格させる（＝変数ではなくなる）。 */
  const SLOT_RE = /\{\{\s*([^}]+?)\s*\}\}/g;

  // 判定の順番が大事：年月日→月日→月 の順に見る
  function slotKind(key) {
    if (/年月日$|日付$/.test(key)) return 'ymd';
    if (/月日$/.test(key))         return 'md';
    if (/月$/.test(key))           return 'month';
    if (isCaseVar(key))            return 'case';
    return 'text';
  }

  const MONTHS = Array.from({ length: 12 }, (_, i) => `${i + 1}月`);

  const SLOT_HINT = {
    ymd:   'クリックでカレンダー',
    md:    'クリックでカレンダー',
    month: 'クリックで月を選ぶ',
    case:  'クリックで案件を選ぶ',
    text:  'クリックで書き込む',
  };

  function slotSpan(key, inst, value, total) {
    const kind  = slotKind(key);
    const shown = value || `{{${key}}}`;
    const num   = total > 1 ? `（${inst}つ目）` : '';
    const hint  = `${key}${num}：${SLOT_HINT[kind]}／Deleteで丸ごと消す`;
    return `<span class="mtp-slot${value ? '' : ' empty'}" data-var="${esc(key)}"` +
           ` data-inst="${inst}" data-kind="${kind}" data-shown="${esc(shown)}"` +
           ` title="${esc(hint)}">${esc(shown)}</span>`;
  }

  /* テンプレート本文を、変数のところだけスロットにしたHTMLにする */
  function renderPreviewInto(el, text, ctx) {
    let out = '', last = 0, m;
    SLOT_RE.lastIndex = 0;
    const src   = String(text || '');
    const total = countNames(src);
    const seen  = {};
    while ((m = SLOT_RE.exec(src)) !== null) {
      const name = m[1];
      const inst = (seen[name] = (seen[name] || 0) + 1);
      out += esc(src.slice(last, m.index));
      out += slotSpan(name, inst, resolveSlot(name, inst, ctx), total[name]);
      last = m.index + m[0].length;
    }
    out += esc(src.slice(last));
    el.innerHTML = out;
    autoGrow(el);
  }

  /* 画面に出ている文章をそのまま取り出す（コピー・メール用） */
  function previewText(el) {
    return el ? (el.innerText !== undefined ? el.innerText : el.textContent) : '';
  }

  /* 変数の値が変わったとき、まだ変数のままのスロットだけ書き換える */
  function setSlotValue(el, key, inst, value) {
    // inst を省いたら、その名前の全部を書き換える（連動しているとき）
    const sel = `.mtp-slot[data-var="${CSS.escape(key)}"]` +
                (inst ? `[data-inst="${inst}"]` : '');
    el.querySelectorAll(sel).forEach(sp => {
      const shown = value || `{{${key}}}`;
      sp.textContent = shown;
      sp.dataset.shown = shown;
      sp.classList.toggle('empty', !value);
    });
    autoGrow(el);
  }

  /* 確認ビューで選んだ値を状態に反映する。
     ダイヤル側の初期値も一緒に揃えないと、入力欄を組み直した拍子に
     選んだ値が上書きされてしまう。 */
  function commitVarValue(state, name, vkey, value, fmt) {
    state.values = state.values || {};
    if (value && typeof value === 'object') {          // カレンダーで選んだ日付
      state.dates = state.dates || {};
      state.dates[vkey] = value;
      const kind = slotKind(name) === 'md' ? 'md' : 'ymd';
      state.values[vkey] = formatDate(value.y, value.m, value.d, fmt || dateFormat(kind));
    } else {
      state.values[vkey] = value;
      const m = /^(\d{1,2})月$/.exec(String(value || ''));
      if (m) {                                         // 月を選んだらダイヤルも合わせる
        state.months = state.months || {};
        state.months[vkey] = Number(m[1]);
      }
    }
    return state.values[vkey];
  }

  /* 手直し後でも、変数のまま残っているスロットには値の変更を届ける */
  function refreshSlots(el, text, ctx) {
    const seen = {};
    placeholderList(text).forEach(name => {
      const inst = (seen[name] = (seen[name] || 0) + 1);
      setSlotValue(el, name, inst, resolveSlot(name, inst, ctx));
    });
  }

  /* 打ち替えられたスロットを普通の文字に降格させる */
  function detachEditedSlots(el) {
    let changed = false;
    el.querySelectorAll('.mtp-slot').forEach(sp => {
      if (sp.textContent !== sp.dataset.shown) {
        sp.classList.remove('mtp-slot', 'empty');
        sp.removeAttribute('title');
        changed = true;
      }
    });
    return changed;
  }

  /* ---------- スロットの操作（クリックでカレンダー／候補） ----------
     opts = { candidates(key), dateOf(key), onPick(key, value), onEdit() } */
  function wirePreview(el, opts) {
    injectStyle();
    if (el._mtpWired) { el._mtpOpts = opts; return; }
    el._mtpWired = true;
    el._mtpOpts = opts;

    el.addEventListener('click', (e) => {
      const sp = e.target.closest && e.target.closest('.mtp-slot');
      if (!sp || !el.contains(sp)) return;
      openSlotEditor(el, sp);
    });

    el.addEventListener('input', () => {
      if (detachEditedSlots(el)) { /* 変数ではなくなった */ }
      autoGrow(el);
      const o = el._mtpOpts || {};
      if (o.onEdit) o.onEdit();
    });

    /* 差し込み箇所は1つのまとまりとして扱う。
       Delete / Backspace を押したら、文字を1つずつではなく丸ごと消す。 */
    el.addEventListener('keydown', (e) => {
      if (e.key !== 'Backspace' && e.key !== 'Delete') return;
      const slot = slotAtCaret(el, e.key);
      if (!slot) return;
      e.preventDefault();
      slot.remove();
      const o = el._mtpOpts || {};
      autoGrow(el);
      if (o.onEdit) o.onEdit();
    });

    // 件名に改行は入れさせない
    if (el.classList.contains('subj')) {
      el.addEventListener('keydown', (e) => { if (e.key === 'Enter') e.preventDefault(); });
    }

    // 書式付きで貼られると崩れるので、貼り付けは常に文字だけにする
    el.addEventListener('paste', (e) => {
      e.preventDefault();
      const t = (e.clipboardData || window.clipboardData).getData('text/plain');
      document.execCommand('insertText', false, t);
    });
  }

  /* キャレットが差し込み箇所の中／すぐ隣にあるかを見る */
  function slotAtCaret(el, key) {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const r = sel.getRangeAt(0);
    if (!el.contains(r.startContainer)) return null;

    const asSlot = (node) => {
      if (!node) return null;
      const e = node.nodeType === 1 ? node : node.parentElement;
      const sp = e && e.closest ? e.closest('.mtp-slot') : null;
      return sp && el.contains(sp) ? sp : null;
    };

    // 選択している／中にキャレットがある
    const inside = asSlot(r.startContainer) || asSlot(r.endContainer);
    if (inside) return inside;
    if (r.collapsed) {
      const node = r.startContainer;
      const at   = r.startOffset;
      if (node.nodeType === 3) {                       // テキストの端にいるとき
        if (key === 'Backspace' && at === 0) return asSlot(node.previousSibling);
        if (key === 'Delete' && at === node.length)    return asSlot(node.nextSibling);
      } else {                                         // 要素の子の境目にいるとき
        const kids = node.childNodes;
        if (key === 'Backspace') return asSlot(kids[at - 1]);
        if (key === 'Delete')    return asSlot(kids[at]);
      }
    }
    return null;
  }

  /* 開いている選択肢・入力枠・カレンダーを閉じる。
     外クリック用のリスナーが前回分も残っていると、開いた直後の枠まで
     閉じてしまうので、閉じるときに必ず外しておく。 */
  let _slotOff = null;

  function closeSlotMenu() {
    document.querySelectorAll('.mtp-slotmenu, .mtp-slotdate').forEach(m => m.remove());
    if (_slotOff) {
      document.removeEventListener('click', _slotOff, true);
      document.removeEventListener('keydown', _slotOff, true);
      _slotOff = null;
    }
  }

  /* 外をクリック／Escで閉じる。開いた枠は常にこれ1つだけ */
  function closeOnOutside(node) {
    if (_slotOff) {
      document.removeEventListener('click', _slotOff, true);
      document.removeEventListener('keydown', _slotOff, true);
    }
    _slotOff = (ev) => {
      if (ev.type === 'keydown') { if (ev.key === 'Escape') closeSlotMenu(); return; }
      if (!node.contains(ev.target)) closeSlotMenu();
    };
    setTimeout(() => {
      if (!_slotOff) return;
      document.addEventListener('click', _slotOff, true);
      document.addEventListener('keydown', _slotOff, true);
    }, 0);
  }

  function openSlotEditor(el, sp) {
    closeSlotMenu();
    const o    = el._mtpOpts || {};
    const key  = sp.dataset.var;
    const inst = Number(sp.dataset.inst) || 1;
    const kind = sp.dataset.kind;
    const vkey = (o.vkeyOf && o.vkeyOf(key, inst)) || key;
    const commit = (value) => {
      if (o.onPick) o.onPick(key, inst, value);   // 入力欄と連動先もまとめて更新
      else setSlotValue(el, key, inst, value);
    };

    if (kind === 'ymd' || kind === 'md') {
      const d = (o.dateOf && o.dateOf(vkey)) || defaultDate();
      const inp = document.createElement('input');
      inp.type = 'date';
      inp.className = 'mtp-slotdate';
      inp.value = `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
      placeNear(sp, inp);

      // カレンダーを閉じたのに入力欄だけ残る、を防ぐ
      inp.addEventListener('change', () => {
        if (inp.value) {
          const [y, m, dd] = inp.value.split('-').map(Number);
          commit({ y, m, d: dd });
        }
        closeSlotMenu();
      });
      inp.addEventListener('blur', () => setTimeout(closeSlotMenu, 200));
      closeOnOutside(inp);

      try {
        if (typeof inp.showPicker === 'function') {
          inp.classList.add('hidden');   // ネイティブのカレンダーだけ見せる
          inp.showPicker();
          return;
        }
      } catch (err) {
        inp.classList.remove('hidden');  // 開けない環境では入力欄をそのまま使ってもらう
      }
      inp.focus();
      return;
    }

    if (kind === 'month') {
      openSlotMenu(el, sp, MONTHS.map(m => ({ label: m })), commit);
      return;
    }

    const cands = (o.candidates && o.candidates(key)) || [];
    if (cands.length) {
      openSlotMenu(el, sp, cands, commit);
      return;
    }

    // 候補が無い変数は、その場に入力枠を出して書いてもらう
    openSlotInput(el, sp, key, vkey, commit);
  }

  /* 候補から選ぶメニュー（案件名・月） */
  function openSlotMenu(el, sp, cands, commit) {
    const menu = document.createElement('div');
    menu.className = 'mtp-slotmenu';
    menu.innerHTML = cands.map(c => {
      const label = c.label !== undefined ? c.label : c;
      return `<button type="button" class="mtp-slotmenu-item" data-val="${esc(label)}">
                <span>${esc(label)}</span>${c.hint ? `<em>${esc(c.hint)}</em>` : ''}</button>`;
    }).join('') + '<div class="mtp-slotmenu-note">そのまま打ち替えると変数ではなくなります</div>';
    placeNear(sp, menu);
    menu.querySelectorAll('.mtp-slotmenu-item').forEach(b => {
      b.addEventListener('click', () => { commit(b.dataset.val); closeSlotMenu(); });
    });
    closeOnOutside(menu);
  }

  /* 自由記載など、候補の無い変数の入力枠。書いた内容はそのまま反映する */
  function openSlotInput(el, sp, key, vkey, commit) {
    const o    = el._mtpOpts || {};
    const cur  = (o.valueOf && o.valueOf(vkey)) || '';
    const long = /自由記載|本文|メモ|備考|内容|理由|原因/.test(key);
    const box  = document.createElement('div');
    box.className = 'mtp-slotmenu mtp-slotinput';
    box.innerHTML = `
      <div class="mtp-slotinput-head">${esc(key)}</div>
      ${long ? `<textarea class="mtp-slotinput-in" rows="4"
                  placeholder="ここに書いた内容がそのまま入ります（改行できます）"></textarea>`
             : `<input class="mtp-slotinput-in" placeholder="${esc(key)}を入力">`}
      <div class="mtp-slotinput-foot">
        <button type="button" class="mtp-slotinput-clear">空にする</button>
        <span class="mtp-slotmenu-note">${long ? '⌘Enter' : 'Enter'}かEscで閉じる／このまま Delete で丸ごと消せます</span>
      </div>`;
    placeNear(sp, box);
    const inp = box.querySelector('.mtp-slotinput-in');
    inp.value = cur;
    inp.focus();
    inp.setSelectionRange(cur.length, cur.length);
    inp.addEventListener('input', () => commit(inp.value));
    // Enterで閉じる。複数行の欄では改行を優先し、⌘/Ctrl+Enterで閉じる
    inp.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      if (long && !(e.metaKey || e.ctrlKey)) return;
      e.preventDefault();
      closeSlotMenu();
      el.focus();
    });
    box.querySelector('.mtp-slotinput-clear').addEventListener('click', () => {
      inp.value = ''; commit(''); inp.focus();
    });
    closeOnOutside(box);
  }

  /* スロットのすぐ下に出す */
  function placeNear(sp, node) {
    const r = sp.getBoundingClientRect();
    node.style.position = 'absolute';
    node.style.left = (window.scrollX + r.left) + 'px';
    node.style.top  = (window.scrollY + r.bottom + 4) + 'px';
    node.style.zIndex = 9600;
    document.body.appendChild(node);
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
    if (el.select) { el.focus(); el.select(); return; }   // input / textarea
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  /* 本文の入力欄は中身に合わせて伸ばす（スクロールバーの中で書かせない） */
  function autoGrow(el) {
    if (!el || el.tagName !== 'TEXTAREA') return;
    el.style.height = 'auto';
    el.style.height = Math.min(el.scrollHeight + 2, 640) + 'px';
  }

  function esc(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /* ---------- 送信時の入力欄 ----------
     テンプレートが使っている変数だけを見て、必要な入力欄を組み立てる。
     container: 差し込む要素 / text: 件名＋本文 / state: 値の入れ物 / onChange: 再描画 */
  /* ---------- 選択と自由入力を兼ねるコントロール ----------
     候補は ◀▶ のダイヤルで送れるが、欄に直接打ち込んでもいい。
     候補が無いときはただのテキスト欄として振る舞う。 */
  let _comboSeq = 0;

  function comboHTML(varKey, value, candidates, placeholder, note) {
    const id = 'mtpdl' + (++_comboSeq);
    const list = candidates || [];
    // 候補が0件でもダイヤルは出したままにする（消えると壊れて見えるため）
    const dis = list.length ? '' : ' disabled';
    return `
      <div class="mtp-combo" data-var="${esc(varKey)}" data-list="${id}">
        <button class="mtp-combo-arw" type="button" data-step="-1" title="前の候補"${dis}>◀</button>
        <input class="mtp-fin mtp-combo-in" data-var="${esc(varKey)}" value="${esc(value || '')}"
               placeholder="${esc(placeholder || '')}" list="${id}">
        <button class="mtp-combo-arw" type="button" data-step="1" title="次の候補"${dis}>▶</button>
        <span class="mtp-combo-count"${note !== undefined ? ` data-note="${esc(note)}"` : ''}
              >${note !== undefined ? esc(note) : '0/' + list.length}</span>
      </div>
      <datalist id="${id}">${list.map(c =>
          `<option value="${esc(c.label !== undefined ? c.label : c)}">${esc(c.hint || '')}</option>`).join('')}</datalist>`;
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
        // 候補が無いときは呼び出し側の文言（読み込み中…／候補なし）を残す
        if (!labels.length) { count.textContent = count.dataset.note || '候補なし'; return; }
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

  function renderVarInputs(container, text, state, onChange) {
    injectStyle();
    state.values = state.values || {};
    state.dates  = state.dates  || {};
    state.ranges = state.ranges || {};
    state.months = state.months || {};
    state.linked = state.linked || {};
    container._mtpArgs = { text, state, onChange };   // 組み直すときのため
    const vars = manualVars(text, state.linked);

    if (!vars.length) {
      container.innerHTML = '<div class="mtp-noinput">この定型文に入力が必要な項目はありません</div>';
      return;
    }

    /* 同じ名前が何回か出てくるときだけ、番号と「連動」ボタンを出す */
    const label = (v) => {
      const num = (v.total > 1 && !v.linked) ? `<span class="mtp-fnum">${v.inst}つ目</span>` : '';
      const link = (v.total > 1 && (v.inst === 1 || v.linked))
        ? `<button type="button" class="mtp-link${v.linked ? ' on' : ''}" data-link="${esc(v.key)}"
             title="${v.linked ? 'それぞれ別の値に戻す' : v.key + 'の' + v.total + '箇所を同じ値にする'}"
             >${v.linked ? '🔗 連動中' : '🔗 連動'}</button>` : '';
      return `<label class="mtp-flabel">${esc(v.key)}${num}${link}</label>`;
    };

    container.innerHTML = vars.map(v => {
      if (v.type === 'date') {
        const d = state.dates[v.vkey] || (state.dates[v.vkey] = defaultDate());
        return `
          <div class="mtp-field">
            ${label(v)}
            <div class="mtp-dial mtp-datedial" data-vkey="${esc(v.vkey)}">
              <input class="mtp-dial-num" type="number" data-part="y" min="2020" max="2035" value="${d.y}" style="width:60px">
              <span class="mtp-dial-sep">年</span>
              <input class="mtp-dial-num" type="number" data-part="m" min="0" max="13" value="${d.m}" style="width:40px">
              <span class="mtp-dial-sep">月</span>
              <input class="mtp-dial-num" type="number" data-part="d" min="0" max="32" value="${d.d}" style="width:40px">
              <span class="mtp-dial-sep">日</span>
              <button class="mtp-dial-today" type="button" title="今日に戻す">今日</button>
              <span class="mtp-cal">
                <button class="mtp-cal-btn" type="button" title="カレンダーから選ぶ">📅</button>
                <input class="mtp-cal-in" type="date" tabindex="-1"
                       value="${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}">
              </span>
            </div>
            <select class="mtp-fmt" data-vkey="${esc(v.vkey)}" data-kind="${esc(v.kind || 'ymd')}">
              ${formatsFor(v.kind).map(f =>
                  `<option value="${f.id}"${f.id === dateFormat(v.kind) ? ' selected' : ''}>${esc(f.label)}</option>`).join('')}
            </select>
          </div>`;
      }

      if (isCaseVar(v.key)) {
        const all = state.cases || [];
        const range = state.ranges[v.vkey] || (state.ranges[v.vkey] = initialRange(all));
        const shown = filterCases(all, range);
        const note = state.casesLoading ? '読み込み中…'
                   : (all.length ? `${shown.length}件` : '候補なし');
        const md = range.month || (range.month = newestCaseMonth(all));
        return `
          <div class="mtp-field">
            ${label(v)}
            <div class="mtp-casewrap">
              ${comboHTML(v.vkey, state.values[v.vkey], shown,
                          all.length ? '候補から選ぶか直接入力' : '案件名を入力', note)}
              <div class="mtp-range" data-vkey="${esc(v.vkey)}">
                ${CASE_RANGES.map(r => {
                    const n = r.id === 'month' ? null : filterCases(all, { mode: r.id }).length;
                    return `<button type="button" class="mtp-range-btn${r.id === range.mode ? ' on' : ''}"
                              data-range="${r.id}">${esc(r.label)}${n === null ? '' : ` <span class="mtp-range-n">${n}</span>`}</button>`;
                  }).join('')}
                ${range.mode === 'month' ? `
                  <span class="mtp-dial mtp-monthdial">
                    <input class="mtp-dial-num" type="number" data-part="y" min="2020" max="2035" value="${md.y}" style="width:60px">
                    <span class="mtp-dial-sep">年</span>
                    <input class="mtp-dial-num" type="number" data-part="m" min="0" max="13" value="${md.m}" style="width:40px">
                    <span class="mtp-dial-sep">月</span>
                    <span class="mtp-cal">
                      <button class="mtp-cal-btn" type="button" title="カレンダーから選ぶ">📅</button>
                      <input class="mtp-cal-in" type="month" tabindex="-1"
                             value="${md.y}-${String(md.m).padStart(2, '0')}">
                    </span>
                  </span>` : ''}
              </div>
            </div>
          </div>`;
      }

      if (v.type === 'month') {
        const m = state.months[v.vkey] || (state.months[v.vkey] = new Date().getMonth() + 1);
        return `
          <div class="mtp-field">
            ${label(v)}
            <span class="mtp-dial mtp-monthonly" data-vkey="${esc(v.vkey)}">
              <input class="mtp-dial-num" type="number" data-part="m" min="0" max="13"
                     value="${m}" style="width:44px">
              <span class="mtp-dial-sep">月</span>
            </span>
          </div>`;
      }

      const isLong = /自由記載|本文|メモ|備考|内容/.test(v.key);
      return `
        <div class="mtp-field">
          ${label(v)}
          ${isLong
            ? `<textarea class="mtp-fin mtp-ftext" data-vkey="${esc(v.vkey)}" rows="3"
                 placeholder="${esc(v.key)}を入力（改行できます）">${esc(state.values[v.vkey] || '')}</textarea>`
            : `<input class="mtp-fin" data-vkey="${esc(v.vkey)}" value="${esc(state.values[v.vkey] || '')}"
                 placeholder="${esc(v.key)}を入力">`}
        </div>`;
    }).join('');

    wireCombo(container, (vkey) => isCaseVar(vkey.split('#')[0])
      ? filterCases(state.cases, state.ranges[vkey])
      : [], onChange);

    // 「連動」ボタン：同じ名前の箇所をまとめる／それぞれ別に戻す。
    // 切り替えたときに入力済みの値が消えないよう、値を引き継いでおく。
    container.querySelectorAll('.mtp-link').forEach(btn => {
      btn.addEventListener('click', () => {
        const name  = btn.dataset.link;
        const total = countNames(text)[name] || 1;
        const on    = !state.linked[name];
        const copy  = (from, to) => {
          if (state.values[from] !== undefined && state.values[to] === undefined)
            state.values[to] = state.values[from];
          if (state.dates[from]  && !state.dates[to])  state.dates[to]  = Object.assign({}, state.dates[from]);
          if (state.months[from] && !state.months[to]) state.months[to] = state.months[from];
        };
        if (on) {
          // 値が入っている一番早い箇所にそろえる（1つ目が空のこともある）
          let src = `${name}#1`;
          for (let i = 1; i <= total; i++) {
            if (state.values[`${name}#${i}`]) { src = `${name}#${i}`; break; }
          }
          delete state.values[name];
          delete state.dates[name];
          delete state.months[name];
          copy(src, name);
        } else {
          for (let i = 1; i <= total; i++) copy(name, `${name}#${i}`);
        }
        state.linked[name] = on;
        rerenderVarInputs(container);
      });
    });

    // 期間ボタン
    container.querySelectorAll('.mtp-range-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const vkey = btn.closest('.mtp-range').dataset.vkey;
        state.ranges[vkey] = Object.assign({}, state.ranges[vkey], { mode: btn.dataset.range });
        rerenderVarInputs(container);
      });
    });

    // 年月ダイヤル（その月の案件だけに絞る）
    container.querySelectorAll('.mtp-monthdial').forEach(dial => {
      const vkey = dial.closest('.mtp-range').dataset.vkey;
      dial.querySelectorAll('.mtp-dial-num').forEach(inp => {
        const handler = () => {
          const md = state.ranges[vkey].month;
          md[inp.dataset.part] = parseInt(inp.value, 10);
          normalizeDate(md);
          rerenderVarInputs(container);
        };
        inp.addEventListener('change', handler);
      });
      wireCalendar(dial, (cal, picked) => {
        state.ranges[vkey].month = { y: picked.y, m: picked.m, d: 1 };
        rerenderVarInputs(container);
      });
    });

    // 月だけのダイヤル（12の次は1、1の前は12）
    container.querySelectorAll('.mtp-monthonly').forEach(dial => {
      const vkey = dial.dataset.vkey;
      const inp = dial.querySelector('[data-part=m]');
      const apply = () => {
        let m = parseInt(inp.value, 10);
        if (!Number.isFinite(m)) m = new Date().getMonth() + 1;
        if (m > 12) m = 1;
        if (m < 1)  m = 12;
        inp.value = m;
        state.months[vkey] = m;
        state.values[vkey] = `${m}月`;
        onChange();
      };
      inp.addEventListener('change', apply);
      inp.addEventListener('input',  apply);
      apply();
    });

    // テキスト欄
    container.querySelectorAll('.mtp-fin').forEach(el => {
      el.addEventListener('input', () => {
        state.values[el.dataset.vkey] = el.value;
        onChange();
      });
    });

    // 日付ダイヤル（年月しぼり・月だけのダイヤルは見た目が同じなだけなので拾わない）
    container.querySelectorAll('.mtp-datedial').forEach(dial => {
      const vkey = dial.dataset.vkey;
      const fmtSel = container.querySelector(`.mtp-fmt[data-vkey="${CSS.escape(vkey)}"]`);
      const apply = () => {
        const d = state.dates[vkey];
        state.values[vkey] = formatDate(d.y, d.m, d.d, fmtSel.value);
        onChange();
      };
      dial.querySelectorAll('.mtp-dial-num').forEach(inp => {
        const handler = () => {
          const d = state.dates[vkey];
          d[inp.dataset.part] = parseInt(inp.value, 10);
          normalizeDate(d);
          showOnDial();
          apply();
        };
        inp.addEventListener('change', handler);
        inp.addEventListener('input',  handler);
      });
      const showOnDial = () => {
        const d = state.dates[vkey];
        dial.querySelector('[data-part=y]').value = d.y;
        dial.querySelector('[data-part=m]').value = d.m;
        dial.querySelector('[data-part=d]').value = d.d;
        const cal = dial.querySelector('.mtp-cal-in');
        if (cal) cal.value = `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
      };
      dial.querySelector('.mtp-dial-today').addEventListener('click', () => {
        state.dates[vkey] = defaultDate();
        showOnDial();
        apply();
      });
      // カレンダーから選んでもダイヤルと表示を合わせる
      wireCalendar(dial, (cal, picked) => {
        state.dates[vkey] = picked;
        showOnDial();
        apply();
      });
      fmtSel.addEventListener('change', () => {
        // 形式の好みは「年あり」「月日だけ」で別々に覚える
        localStorage.setItem('mailtpl_datefmt_' + (fmtSel.dataset.kind === 'md' ? 'md' : 'ymd'),
                             fmtSel.value);
        apply();
      });
      apply();   // 初期値をすぐ反映する
    });
  }


  /* 期間を変えたときなど、同じ引数で入力欄を組み直す */
  function rerenderVarInputs(container) {
    const a = container._mtpArgs;
    if (!a) return;
    renderVarInputs(container, a.text, a.state, a.onChange);
    a.onChange();
  }

  /* 📅 を押したらネイティブのカレンダーを開く。
     showPicker が無い環境では、その場で日付入力欄そのものを出して選んでもらう。 */
  function wireCalendar(scope, onPick) {
    scope.querySelectorAll('.mtp-cal').forEach(cal => {
      const btn = cal.querySelector('.mtp-cal-btn');
      const inp = cal.querySelector('.mtp-cal-in');
      btn.addEventListener('click', () => {
        try {
          if (typeof inp.showPicker === 'function') { inp.showPicker(); return; }
        } catch (e) { /* ユーザー操作以外からは開けない等。下のフォールバックへ */ }
        cal.classList.add('open');   // 入力欄を出して直接触ってもらう
        inp.focus();
      });
      inp.addEventListener('change', () => {
        if (!inp.value) return;
        const [y, m, d] = inp.value.split('-').map(Number);
        onPick(cal, { y, m, d: d || 1 });
      });
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
      /* プレビューは出力したあとに手で直せる */
      .mtp-prev { display:block; width:100%; background:#fafafa; border:1.5px solid #eee;
        border-radius:9px; padding:10px 12px; font-size:13px; line-height:1.75;
        font-family:inherit; color:#1a1a1a; white-space:pre-wrap; word-break:break-word;
        min-height:40px; max-height:50vh; overflow-y:auto; }
      .mtp-prev:focus { outline:none; border-color:#4f6ef7; background:#fff; }
      .mtp-prev.subj { white-space:normal; font-weight:600; min-height:0; resize:none; }
      .mtp-editable { font-size:10px; color:#ccc; font-weight:400; margin-left:6px; }

      /* 確認ビューの変数スロット */
      .mtp-slot { background:#eef1ff; border-radius:4px; padding:1px 3px; cursor:pointer;
        box-shadow:inset 0 -1px 0 #c7d2fe; transition:background 0.12s, box-shadow 0.12s; }
      .mtp-slot:hover { background:#4f6ef7; color:#fff; box-shadow:inset 0 -1px 0 #3b5ce0; }
      .mtp-slot.empty { background:#fffbeb; color:#b45309; box-shadow:inset 0 -1px 0 #fde68a; }
      .mtp-slot.empty:hover { background:#f59e0b; color:#fff; }
      /* 自由記載など「ここを実情に合わせて書く」場所は、書ける枠だと分かる見た目にする */
      .mtp-slot[data-kind="text"] { cursor:text; border-bottom:1px dashed #a5b4fc; }
      .mtp-slot[data-kind="text"]:hover { background:#e0e7ff; color:#3b5ce0;
        box-shadow:inset 0 0 0 1px #a5b4fc; border-bottom-color:transparent; }
      .mtp-slot[data-kind="text"].empty:hover { background:#fef3c7; color:#92400e;
        box-shadow:inset 0 0 0 1px #fcd34d; }

      .mtp-slotinput { padding:10px; min-width:280px; }
      .mtp-slotinput-head { font-size:11px; font-weight:600; color:#999; margin-bottom:6px; }
      .mtp-slotinput-in { width:100%; border:1.5px solid #e2e2e2; border-radius:8px;
        padding:8px 10px; font-size:13px; font-family:inherit; line-height:1.7;
        background:#fafafa; color:#1a1a1a; resize:vertical; }
      .mtp-slotinput-in:focus { outline:none; border-color:#4f6ef7; background:#fff; }
      .mtp-slotinput-foot { display:flex; align-items:center; gap:8px; margin-top:6px; }
      .mtp-slotinput-clear { border:1.5px solid #e2e2e2; background:#fff; border-radius:7px;
        font-size:11px; color:#666; padding:3px 9px; cursor:pointer; font-family:inherit; }
      .mtp-slotinput-clear:hover { background:#f5f5f5; }
      .mtp-slotinput .mtp-slotmenu-note { border:none; margin:0; padding:0; }
      .mtp-slotdate { border:1.5px solid #4f6ef7; border-radius:8px; padding:5px 8px;
        font-size:13px; font-family:inherit; background:#fff; }
      .mtp-slotdate.hidden { opacity:0; width:1px; height:1px; padding:0; border:none;
        pointer-events:none; }
      .mtp-slotmenu { background:#fff; border:1.5px solid #e2e2e2; border-radius:10px;
        box-shadow:0 8px 24px rgba(0,0,0,0.16); padding:5px; max-height:260px; overflow-y:auto;
        min-width:200px; max-width:340px;
        font-family:-apple-system,BlinkMacSystemFont,"Hiragino Sans","Yu Gothic",sans-serif; }
      .mtp-slotmenu-item { display:flex; align-items:baseline; gap:8px; width:100%;
        border:none; background:none; text-align:left; font-size:13px; font-family:inherit;
        color:#1a1a1a; padding:6px 9px; border-radius:7px; cursor:pointer; }
      .mtp-slotmenu-item:hover { background:#eef1ff; color:#3b5ce0; }
      .mtp-slotmenu-item em { font-style:normal; font-size:10.5px; color:#bbb; margin-left:auto;
        white-space:nowrap; }
      .mtp-slotmenu-note { font-size:10.5px; color:#bbb; padding:5px 9px 3px; border-top:1px solid #f0f0f0;
        margin-top:3px; }
      .mtp-edited { font-size:11.5px; color:#3b5ce0; background:#eef1ff; border:1px solid #c7d2fe;
        border-radius:7px; padding:7px 10px; margin-top:8px;
        display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
      .mtp-relink { border:1.5px solid #c7d2fe; background:#fff; color:#3b5ce0; border-radius:7px;
        font-size:11px; padding:3px 9px; cursor:pointer; font-family:inherit; margin-left:auto; }
      .mtp-relink:hover { background:#4f6ef7; border-color:#4f6ef7; color:#fff; }
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
      .mtp-fnum { display:block; font-size:10px; color:#c4c4c4; font-weight:400; }
      .mtp-link { display:block; margin-top:3px; border:1.5px solid #e8e8e8; background:#fff;
        border-radius:20px; font-size:10px; color:#999; padding:2px 7px; cursor:pointer;
        font-family:inherit; }
      .mtp-link:hover { background:#f5f5f5; color:#666; }
      .mtp-link.on { background:#eef1ff; border-color:#4f6ef7; color:#4f6ef7; font-weight:600; }

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
      .mtp-combo-count { font-size:11px; color:#bbb; flex-shrink:0; min-width:44px; }
      .mtp-combo-arw:disabled { opacity:0.4; cursor:default; }
      .mtp-combo-arw:disabled:hover { background:#fff; border-color:#d1d5db; color:#6b7280; }

      /* 案件候補の期間しぼり */
      .mtp-casewrap { flex:1; min-width:220px; }
      .mtp-range { display:flex; align-items:center; gap:5px; flex-wrap:wrap; margin-top:6px; }
      .mtp-range-btn { border:1.5px solid #e8e8e8; background:#fff; border-radius:20px;
        font-size:11px; color:#888; padding:3px 10px; cursor:pointer; font-family:inherit; }
      .mtp-range-btn:hover { background:#f5f5f5; color:#555; }
      .mtp-range-btn.on { background:#eef1ff; border-color:#4f6ef7; color:#4f6ef7; font-weight:600; }
      .mtp-range-n { color:#bbb; font-size:10px; }
      .mtp-range-btn.on .mtp-range-n { color:#8ea0f7; }
      /* カレンダー（ネイティブの日付ピッカーを開くだけ） */
      .mtp-cal { position:relative; display:inline-flex; align-items:center; margin-left:2px; }
      .mtp-cal-btn { border:1.5px solid #e2e2e2; background:#fafafa; border-radius:7px;
        font-size:12px; padding:3px 7px; cursor:pointer; font-family:inherit; line-height:1.4; }
      .mtp-cal-btn:hover { background:#eef1ff; border-color:#4f6ef7; }
      .mtp-cal-in { position:absolute; right:0; bottom:0; width:1px; height:1px;
        opacity:0; border:none; padding:0; }
      /* showPicker が使えない環境では入力欄そのものを出す */
      .mtp-cal.open .mtp-cal-in { position:static; width:auto; height:auto; opacity:1;
        border:1.5px solid #4f6ef7; border-radius:7px; padding:4px 6px; font-size:12px;
        font-family:inherit; margin-left:4px; }
      .mtp-cal.open .mtp-cal-btn { display:none; }

      .mtp-monthdial { padding:2px 8px; }
      .mtp-monthonly { padding:3px 10px; }
      .mtp-monthdial .mtp-dial-num { font-size:13px; }
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
      <div class="mtp-label">件名<span class="mtp-editable">色の付いた所をクリックすると選び直せます</span></div>
      <div class="mtp-prev subj" id="mtpSubj" contenteditable="true"></div>
      <div class="mtp-label">本文</div>
      <div class="mtp-prev" id="mtpBody" contenteditable="true"></div>
      <div id="mtpEdited"></div>
      <div id="mtpWarn"></div>`;
    modal.insertAdjacentHTML('beforeend', `
      <div class="mtp-foot">
        <button class="mtp-btn primary" id="mtpCopyBody">本文をコピー</button>
        <button class="mtp-btn" id="mtpCopySubj">件名をコピー</button>
        <button class="mtp-btn" id="mtpMail">メールを開く</button>
        <button class="mtp-btn ghost" id="mtpEdit">定型文を編集</button>
      </div>`);

    const $ = (id) => modal.querySelector('#' + id);
    const state = { values: {}, dates: {}, months: {}, linked: {},
                    cases: [], ranges: {}, casesLoading: true };


    // その顧客の案件を案件名の候補にする（引けなくても手入力はできる）
    loadCases(customer.no).then(list => {
      state.cases = list;
      state.ranges = {};
      state.casesLoading = false;
      rebuild();
    });

    const ctx = () => ({
      customerName: customer.name || customer.pageName || '',
      customerNo:   customer.no || '',
      invoiceName:  customer.invoiceName || '',
      values:       state.values,
      linked:       state.linked,
    });

    // プレビューを手で直したら、変数を変えてもその編集を上書きしない
    let edited = false;

    function syncWarn() {
      const miss = missingVars(previewText($('mtpSubj')) + '\n' + previewText($('mtpBody')));
      $('mtpWarn').innerHTML = miss.length
        ? `<div class="mtp-warn">未入力の項目があります: ${esc(miss.join(' '))}</div>` : '';
      $('mtpEdited').innerHTML = edited
        ? `<div class="mtp-edited">手で直した内容を表示しています（変数を変えても反映されません）
             <button type="button" class="mtp-relink">テンプレートから作り直す</button></div>` : '';
      const relink = $('mtpEdited').querySelector('.mtp-relink');
      if (relink) relink.onclick = () => { edited = false; render(); };
    }

    function render() {
      const t = templates[Number($('mtpSel').value)] || {};
      if (edited) {
        // 手直しした文章は残したまま、変数のままの箇所だけ更新する
        refreshSlots($('mtpSubj'), t.subject, ctx());
        refreshSlots($('mtpBody'), t.body, ctx());
      } else {
        renderPreviewInto($('mtpSubj'), t.subject, ctx());
        renderPreviewInto($('mtpBody'), t.body, ctx());
      }
      syncWarn();
    }
    function rebuild() {
      const t = templates[Number($('mtpSel').value)] || {};
      renderVarInputs($('mtpInputs'), (t.subject || '') + '\n' + (t.body || ''), state, render);
      render();
    }

    /* 確認ビューで選び直したとき。同じ変数の箇所と入力欄をまとめて更新する */
    function applyVarValue(name, inst, value) {
      const vkey = instKeyOf(name, inst, state.linked);
      const fmtSel = $('mtpInputs').querySelector(`.mtp-fmt[data-vkey="${CSS.escape(vkey)}"]`);
      commitVarValue(state, name, vkey, value, fmtSel && fmtSel.value);
      // 連動しているときは同じ名前の全部、していないときはその箇所だけ
      const only = state.linked[name] ? null : inst;
      setSlotValue($('mtpSubj'), name, only, state.values[vkey]);
      setSlotValue($('mtpBody'), name, only, state.values[vkey]);
      const t = templates[Number($('mtpSel').value)] || {};
      renderVarInputs($('mtpInputs'), (t.subject || '') + '\n' + (t.body || ''), state, render);
      syncWarn();
    }

    const previewOpts = {
      vkeyOf:     (name, inst) => instKeyOf(name, inst, state.linked),
      candidates: (name) => isCaseVar(name.split('#')[0])
        ? filterCases(state.cases, state.ranges[name]) : [],
      dateOf:     (vkey) => state.dates[vkey],
      valueOf:    (vkey) => state.values[vkey] || '',
      onPick:     applyVarValue,
      onEdit:     () => { edited = true; syncWarn(); },
    };
    wirePreview($('mtpSubj'), previewOpts);
    wirePreview($('mtpBody'), previewOpts);
    // 定型文を選び直したら、その定型文の内容から作り直す
    $('mtpSel').onchange = () => { edited = false; rebuild(); };
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
    $('mtpCopyBody').onclick = (e) => copyOr(e.target, previewText($('mtpBody')), $('mtpBody'));
    $('mtpCopySubj').onclick = (e) => copyOr(e.target, previewText($('mtpSubj')), $('mtpSubj'));
    $('mtpMail').onclick = () => {
      window.location.href = `mailto:?subject=${encodeURIComponent(previewText($('mtpSubj')))}` +
                             `&body=${encodeURIComponent(previewText($('mtpBody')))}`;
    };
    $('mtpEdit').onclick = () => { window.location.href = '/メール定型文ツール.html'; };
  }

  window.MailTpl = {
    VARS, AUTO_VARS, MANUAL_VARS, AUTO_KEYS, DATE_FORMATS,
    fill, missingVars, placeholders, manualVars, renderVarInputs,
    formatDate, todayJa, senderName, dateFormat, formatsFor, MD_FORMATS, MONTHS,
    loadTemplates, loadCases, isCaseVar, comboHTML, wireCombo,
    filterCases, initialRange, rerenderVarInputs, CASE_RANGES,
    copyText, selectElementText, injectStyle, autoGrow, wireCalendar, openPicker,
    resolveVar, slotKind, renderPreviewInto, previewText, setSlotValue, refreshSlots,
    clearCustomerOverrides,
    detachEditedSlots, wirePreview, closeSlotMenu, slotAtCaret, commitVarValue,
    placeholderList, instKeyOf, resolveSlot, countNames,
  };
})();
