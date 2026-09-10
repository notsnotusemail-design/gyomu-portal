/*
 * customer_link.js — 全ページ共通：お客様No / 顧客名 を Notion の顧客ページへのリンクにする
 *
 * 使い方: 各HTMLの </body> 直前で  <script src="/customer_link.js" defer></script>
 *
 * 仕組み
 *  1. /api/customer-links から {お客様No: {name, url}} を取得（sessionStorage に5分キャッシュ）
 *  2. 明示指定:  <span data-cust-no="302">…</span>  → その要素をリンク化
 *                 data-cust-url="https://…" があればそれを優先（Noが無い顧客にも対応）
 *  3. 自動検出:  画面上のテキストから「登録済みのお客様No」「顧客名」を見つけてリンク化
 *                 - 案件番号 302-26-0907 のような文字列は先頭の 302 だけをリンクにする
 *                 - 金額(302,000 / ¥302 / 302円)・日付・ID等の数字の並びには反応しない
 *                 - 入力欄 / ボタン / onclick 付き要素 / .no-cust-link の中は触らない
 *  4. 描画後に追加された要素も MutationObserver で追いかける
 *
 *  ホバーで「Notion ↗ 名前（No）」の吹き出し、クリックで新しいタブに Notion を開く。
 *  window.CustLink.refresh() で再取得、CustLink.scan(root) で手動再走査できる。
 */
(function () {
  'use strict';
  if (window.__custLinkLoaded) return;
  window.__custLinkLoaded = true;

  var API_PATH  = '/api/customer-links';
  var CACHE_KEY = 'custLinks:v1';
  var CACHE_TTL = 5 * 60 * 1000;
  var SKIP_SEL  = 'script,style,textarea,input,select,option,button,a,code,pre,' +
                  '[contenteditable],[onclick],.no-cust-link,.cust-link,[data-cust-no]';

  /* ---------- スタイル ---------- */
  var css = [
    '.cust-link{color:inherit;text-decoration:underline dotted;text-decoration-color:#8b5cf6;',
    '  text-decoration-thickness:1.5px;text-underline-offset:3px;cursor:pointer;position:relative;',
    '  border-radius:3px;transition:background .12s,color .12s}',
    '.cust-link:hover,.cust-link.cust-hover{color:#6d28d9;background:rgba(139,92,246,.10);text-decoration-style:solid}',
    '.cust-link:hover::after,.cust-link.cust-hover::after{content:attr(data-tip);position:absolute;left:0;bottom:calc(100% + 7px);',
    '  background:#1f2937;color:#fff;font-size:11px;font-weight:500;line-height:1.3;padding:5px 9px;',
    '  border-radius:6px;white-space:nowrap;z-index:99999;pointer-events:none;',
    '  box-shadow:0 4px 14px rgba(0,0,0,.22)}',
    '.cust-link:hover::before,.cust-link.cust-hover::before{content:"";position:absolute;left:12px;bottom:calc(100% + 2px);',
    '  border:5px solid transparent;border-top-color:#1f2937;z-index:99999;pointer-events:none}'
  ].join('\n');
  var styleEl = document.createElement('style');
  styleEl.textContent = css;
  (document.head || document.documentElement).appendChild(styleEl);

  /* ---------- リンク表の取得 ---------- */
  var LINKS = null;      // {no: {name, url}}
  var matchers = null;   // 構築済みの正規表現
  var loading = null;

  function readCache() {
    try {
      var raw = sessionStorage.getItem(CACHE_KEY);
      if (!raw) return null;
      var obj = JSON.parse(raw);
      if (!obj || !obj.at || Date.now() - obj.at > CACHE_TTL) return null;
      return obj.links || null;
    } catch (e) { return null; }
  }
  function writeCache(links) {
    try { sessionStorage.setItem(CACHE_KEY, JSON.stringify({ at: Date.now(), links: links })); } catch (e) {}
  }
  function loadLinks(force) {
    if (!force) {
      var c = readCache();
      if (c) { LINKS = c; buildMatchers(); return Promise.resolve(LINKS); }
    }
    if (loading) return loading;
    loading = fetch(API_PATH + (force ? '?refresh=1' : ''), { cache: 'no-store' })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d || !d.ok || !d.links) throw new Error('no links');
        LINKS = d.links; writeCache(LINKS); buildMatchers(); return LINKS;
      })
      .catch(function () { LINKS = LINKS || {}; buildMatchers(); return LINKS; })
      .then(function (l) { loading = null; return l; });
    return loading;
  }

  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  function buildMatchers() {
    // 自動検出に使うNo: 3桁以上の数字で始まる、または英字を含むもの。
    // 「00」「10」のような2桁は 10月 / 10:00 / 13件 に誤爆するので、data-cust-no の明示指定でだけ効く。
    var nos = Object.keys(LINKS || {}).filter(function (k) {
      return k && (/^\d{3,}/.test(k) || /[A-Za-z]/.test(k)) && !/\s/.test(k);
    }).sort(function (a, b) { return b.length - a.length; });
    var nameToNo = {};
    nos.forEach(function (no) {
      var nm = (LINKS[no].name || '').trim();
      if (nm.length >= 3 && !nameToNo[nm]) nameToNo[nm] = no;
    });
    var names = Object.keys(nameToNo).sort(function (a, b) { return b.length - a.length; });

    // 直後にこれが来たら No ではなく 数量・金額・時刻 と見なす
    var UNIT = '[0-9A-Za-z,.¥￥円万千億月日時分秒件人回本個枚台歳年％%:：／/]';
    var noRe = null, noReNoLB = null;
    if (nos.length) {
      var alt = nos.map(escRe).join('|');
      // 前後が数字・英字・カンマ・小数点・通貨・円 なら金額/日付/IDの一部と見なして除外。
      // 「302-26-0907」のように直後が「-」の案件番号は先頭のNoだけ拾う。
      try {
        noRe = new RegExp('(?<![0-9A-Za-z,.¥￥\\-])(' + alt + ')(?!' + UNIT + ')', 'g');
      } catch (e) {
        noReNoLB = new RegExp('(' + alt + ')(?!' + UNIT + ')', 'g');   // 後読み非対応ブラウザ用
      }
    }
    matchers = {
      nos: nos, noRe: noRe, noReNoLB: noReNoLB, nameToNo: nameToNo,
      nameRe: names.length ? new RegExp('(' + names.map(escRe).join('|') + ')', 'g') : null,
      quick: new RegExp('[0-9A-Za-z]' + (names.length ? '|' + names.map(escRe).join('|') : ''))
    };
  }

  /* ---------- 文字列から お客様No を解決 ---------- */
  function resolveNo(str) {
    if (!str || !LINKS) return null;
    var s = String(str).trim();
    if (LINKS[s]) return s;
    var m = findMatches(s);
    for (var i = 0; i < m.length; i++) if (m[i].no) return m[i].no;
    return null;
  }

  function tipFor(no, url) {
    var nm = (LINKS[no] && LINKS[no].name) || '';
    return 'Notion ↗ ' + (nm ? nm + '（' + no + '）' : no);
  }
  function urlFor(no) { return LINKS[no] && LINKS[no].url; }

  /* ---------- テキスト中のマッチ列挙（重なりは先頭優先） ---------- */
  function findMatches(text) {
    var out = [];
    if (!matchers) return out;
    var m;
    if (matchers.noRe) {
      matchers.noRe.lastIndex = 0;
      while ((m = matchers.noRe.exec(text))) out.push({ i: m.index, len: m[1].length, no: m[1] });
    } else if (matchers.noReNoLB) {
      matchers.noReNoLB.lastIndex = 0;
      while ((m = matchers.noReNoLB.exec(text))) {
        var prev = m.index > 0 ? text[m.index - 1] : '';
        if (!/[0-9A-Za-z,.¥￥\-]/.test(prev)) out.push({ i: m.index, len: m[1].length, no: m[1] });
      }
    }
    if (matchers.nameRe) {
      matchers.nameRe.lastIndex = 0;
      while ((m = matchers.nameRe.exec(text))) out.push({ i: m.index, len: m[1].length, no: matchers.nameToNo[m[1]] });
    }
    out.sort(function (a, b) { return a.i - b.i || b.len - a.len; });
    var res = [], end = -1;
    out.forEach(function (x) { if (x.i >= end) { res.push(x); end = x.i + x.len; } });
    return res;
  }

  /* ---------- DOM 変換 ---------- */
  var busy = false;

  function makeAnchor(no, text, url) {
    var a = document.createElement('a');
    a.className = 'cust-link';
    a.href = url || urlFor(no);
    a.target = '_blank';
    a.rel = 'noopener';
    a.setAttribute('data-tip', tipFor(no));
    a.setAttribute('data-cust-linked', no);
    a.textContent = text;
    return a;
  }

  function isSkipped(el) {
    for (var n = el; n && n.nodeType === 1; n = n.parentNode) {
      if (n.matches && n.matches(SKIP_SEL)) return true;
    }
    return false;
  }

  function scanTextNodes(root) {
    if (!matchers || !matchers.nos.length && !matchers.nameRe) return 0;
    var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: function (node) {
        var t = node.nodeValue;
        if (!t || t.length < 2 || !matchers.quick.test(t)) return NodeFilter.FILTER_REJECT;
        if (!node.parentNode || isSkipped(node.parentNode)) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    var nodes = [], n;
    while ((n = walker.nextNode())) nodes.push(n);
    var count = 0;
    nodes.forEach(function (node) {
      var text = node.nodeValue;
      var ms = findMatches(text);
      if (!ms.length) return;
      var frag = document.createDocumentFragment(), pos = 0;
      ms.forEach(function (x) {
        if (!urlFor(x.no)) return;
        if (x.i > pos) frag.appendChild(document.createTextNode(text.slice(pos, x.i)));
        frag.appendChild(makeAnchor(x.no, text.slice(x.i, x.i + x.len)));
        pos = x.i + x.len; count++;
      });
      if (pos === 0) return;
      if (pos < text.length) frag.appendChild(document.createTextNode(text.slice(pos)));
      node.parentNode.replaceChild(frag, node);
    });
    return count;
  }

  function enhanceExplicit(root) {
    var els = (root.querySelectorAll ? root : document).querySelectorAll('[data-cust-no]:not([data-cust-linked])');
    var count = 0;
    Array.prototype.forEach.call(els, function (el) {
      var raw = el.getAttribute('data-cust-no') || el.textContent;
      var no  = resolveNo(raw);
      var url = el.getAttribute('data-cust-url') || (no && urlFor(no));
      if (!url) return;
      el.classList.add('cust-link');
      el.setAttribute('data-cust-linked', no || '');
      el.setAttribute('data-url', url);
      el.setAttribute('data-tip', no ? tipFor(no) : 'Notion ↗ ' + (el.textContent || '').trim());
      if (el.tagName === 'A') { el.href = url; el.target = '_blank'; el.rel = 'noopener'; }
      count++;
    });
    return count;
  }

  function scan(root) {
    root = root || document.body;
    if (!root || !LINKS) return 0;
    busy = true;
    var c = 0;
    try { c = enhanceExplicit(root) + scanTextNodes(root); }
    finally { busy = false; }
    return c;
  }

  /* クリック：明示指定の要素（a以外）は新しいタブで開く。親の onclick（カード開閉など）へは伝えない */
  document.addEventListener('click', function (e) {
    var el = e.target && e.target.closest ? e.target.closest('.cust-link') : null;
    if (!el) return;
    e.stopPropagation();
    if (el.tagName !== 'A') {
      var url = el.getAttribute('data-url');
      if (url) { e.preventDefault(); window.open(url, '_blank', 'noopener'); }
    }
  }, true);

  /* ホバー：:hover が効かない環境（タッチ・一部の埋め込みブラウザ）でも吹き出しが出るようクラスでも制御 */
  document.addEventListener('mouseover', function (e) {
    var el = e.target && e.target.closest ? e.target.closest('.cust-link') : null;
    if (el) el.classList.add('cust-hover');
  }, true);
  document.addEventListener('mouseout', function (e) {
    var el = e.target && e.target.closest ? e.target.closest('.cust-link') : null;
    if (el && !(e.relatedTarget && el.contains(e.relatedTarget))) el.classList.remove('cust-hover');
  }, true);

  /* 後から描画された分も追いかける */
  var timer = null;
  var mo = new MutationObserver(function () {
    if (busy) return;
    clearTimeout(timer);
    timer = setTimeout(function () { scan(document.body); }, 150);
  });

  function start() {
    loadLinks(false).then(function () {
      scan(document.body);
      if (document.body) mo.observe(document.body, { childList: true, subtree: true, characterData: true });
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();

  window.CustLink = {
    refresh: function () { return loadLinks(true).then(function () { return scan(document.body); }); },
    scan: scan,
    resolveNo: resolveNo,
    url: function (no) { return urlFor(no) || ''; },
    html: function (no, text) {   // テンプレート文字列から使う用
      var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (ch) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]; }); };
      return '<span data-cust-no="' + esc(no) + '">' + esc(text == null ? no : text) + '</span>';
    }
  };
})();
