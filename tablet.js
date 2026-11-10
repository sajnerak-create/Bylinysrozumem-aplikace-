/*
 * Byliny s rozumem – tabletová aplikace.
 * Vzhled je stejný soubor aplikace.html jako na počítači. Tento skript ho spustí v tabletu
 * a místo přímého spojení s tabulkou použije data uložená v tabletu:
 *  - data se stahují z Apps Scriptu (tabletová implementace) a ukládají se pro čtení bez internetu,
 *  - při spuštění s internetem se automaticky zkontrolují novější data, ručně tlačítkem „Aktualizovat“,
 *  - e-mail a PDF posílá Apps Script (jen s internetem), upravovat na tabletu nejde.
 */
(function BSR() {
  'use strict';
  var K_CFG = 'bsr-tablet', K_DATA = 'bsr-data', K_UI = 'bsr-ui';

  /**
   * Odkaz na tabulku zveřejněnou na webu (Soubor → Sdílet → Publikovat na webu → Celý dokument → .xlsx).
   * Když ho tablet má uložený, tablet čte tabulku napřímo a nepotřebuje Apps Script ani klíč.
   * E-mail pak připraví v poštovní aplikaci tabletu a PDF vytvoří přes tisk v tabletu.
   */
  var K_ZDROJ = 'bsr-zdroj', K_TAB = 'bsr-tabulka';
  // Odkaz se do tabletu uloží z adresy …#zdroj=<odkaz> nebo z úvodní obrazovky. V kódu (veřejném) není.
  (function () {
    var h = new URLSearchParams(location.hash.slice(1));
    if (h.get('zdroj') || h.get('tabulka')) {
      try {
        if (h.get('zdroj')) localStorage.setItem(K_ZDROJ, JSON.stringify(h.get('zdroj')));
        if (h.get('tabulka')) localStorage.setItem(K_TAB, JSON.stringify(h.get('tabulka')));
      } catch (e) {}
      history.replaceState(null, '', location.pathname + location.search);
    }
  })();
  var TABULKA = (function () { try { return JSON.parse(localStorage.getItem(K_TAB) || '""'); } catch (e) { return ''; } })();
  var ZDROJ = (function () { try { return JSON.parse(localStorage.getItem(K_ZDROJ) || '""'); } catch (e) { return ''; } })();
  var PRIMO = !!(ZDROJ || TABULKA);

  /* ---------- ukládání v tabletu ---------- */
  function get(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } }
  function put(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
  function cfg() { var c = get(K_CFG); return c && c.url && c.klic ? c : null; }

  /* ---------- spojení s Apps Scriptem ---------- */
  function call(akce, extra, ms) {
    var c = cfg();
    if (!c) return Promise.reject(new Error('Tablet zatím není propojený s tabulkou.'));
    if (navigator.onLine === false) return Promise.reject(new Error('Tablet je bez internetu.'));
    var body = { akce: akce, klic: c.klic };
    Object.keys(extra || {}).forEach(function (k) { body[k] = extra[k]; });
    var ctl = window.AbortController ? new AbortController() : null;
    var t = ctl && setTimeout(function () { ctl.abort(); }, ms || 90000);
    return fetch(c.url, {
      method: 'POST', redirect: 'follow', signal: ctl && ctl.signal,
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (r) {
      return r.text().then(function (txt) {
        var j; try { j = JSON.parse(txt); } catch (e) {
          throw new Error('Tabulka neodpověděla správně. Zkontrolujte, že tabletová implementace má přístup „Kdokoli“.');
        }
        if (!j.ok) throw new Error(j.chyba || 'Neznámá chyba.');
        return j;
      });
    }, function (e) {
      throw new Error(e && e.name === 'AbortError' ? 'Tabulka neodpověděla včas.' : 'Nepodařilo se spojit s tabulkou.');
    }).finally(function () { if (t) clearTimeout(t); });
  }

  /* ---------- přímé čtení zveřejněné tabulky ---------- */
  var TYPY = { 'Obtíže': 'ailments', 'Byliny': 'herbs', 'Recepty': 'recipes', 'Letáky': 'leaflets', 'Monografie Dr. Svátek': 'monographs' };
  var JSON_SL = { ingredients: 1, sections: 1 };

  function nactiXlsx() {
    if (window.XLSX) return Promise.resolve();
    return new Promise(function (res, rej) {
      var sc = document.createElement('script');
      sc.src = 'xlsx.mini.min.js'; sc.onload = function () { res(); };
      sc.onerror = function () { rej(new Error('Nepodařilo se načíst čtečku tabulky.')); };
      document.head.appendChild(sc);
    });
  }

  /** Převede sešit na stejná data, jaká dává Apps Script (getAll). */
  function zXlsx(wb) {
    return wb.SheetNames.map(function (n) {
      return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) };
    });
  }

  /* Záložní cesta bez omezení prohlížeče: listy po jednom přes Google Visualization (JSONP). */
  var LISTY = ['Obtíže', 'Byliny', 'Recepty', 'Letáky', 'Monografie Dr. Svátek', 'Průvodci', 'Nastavení'];
  var jsonpN = 0;
  function gvizList(key, name, ms) {
    return new Promise(function (res, rej) {
      var cb = '__bsrG' + (++jsonpN) + '_' + Date.now();
      var sc = document.createElement('script');
      var t = setTimeout(function () { done(); rej(new Error('Tabulka neodpověděla včas.')); }, 25000);
      function done() { clearTimeout(t); try { delete window[cb]; } catch (e) { window[cb] = undefined; } sc.remove(); }
      window[cb] = function (r) { done(); res(zGviz(name, r)); };
      sc.onerror = function () { done(); rej(new Error('Záložní spojení s tabulkou selhalo.')); };
      sc.src = gvizUrl(name, '&tqx=responseHandler:' + cb);
      document.head.appendChild(sc);
    });
  }
  function gvizUrl(name, extra) {
    return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(TABULKA) + '/gviz/tq?headers=1' + extra +
      '&sheet=' + encodeURIComponent(name) + '&t=' + Date.now();
  }
  function zGviz(name, r) {
    if (!r || r.status === 'error' || !r.table) return null;
    var cols = r.table.cols.map(function (c) { return String(c.label || '').trim(); });
    return { name: name, rows: [cols].concat(r.table.rows.map(function (row) {
      return (row.c || []).map(function (c) { return c && c.v != null ? c.v : ''; });
    })) };
  }
  /** Rozebere CSV (uvozovky, čárky a zalomení řádků uvnitř buněk). */
  function zCsv(txt) {
    var rows = [], row = [], cell = '', q = false;
    for (var i = 0; i < txt.length; i++) {
      var ch = txt[i];
      if (q) {
        if (ch === '"') { if (txt[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && txt[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  /* Cesta 2: list po listu jako CSV přes fetch (Google u tohoto rozhraní stahování z jiných stránek povoluje). */
  function gvizFetch(name, ms) {
    var ctl = window.AbortController ? new AbortController() : null;
    var t = ctl && setTimeout(function () { ctl.abort(); }, ms || 60000);
    return fetch(gvizUrl(name, '&tqx=out:csv'), { signal: ctl && ctl.signal, cache: 'no-store', credentials: 'omit' })
      .then(function (r) {
        return r.text().then(function (txt) {
          var ukazka = name + ': ' + r.status + ' ' + txt.slice(0, 120).replace(/\s+/g, ' ');
          if (!r.ok || /^\s*</.test(txt)) { gvizFetch.chyby.push(ukazka); return null; }
          var rows = zCsv(txt.replace(/^\uFEFF/, ''));
          if (!rows.length) return null;
          rows[0] = rows[0].map(function (h) { return h.replace(/^\uFEFF/, '').trim(); });
          if (rows[0][0] !== 'id') gvizFetch.chyby.push(ukazka);
          return { name: name, rows: rows };
        });
      })
      .finally(function () { if (t) clearTimeout(t); });
  }
  function stahnoutGviz(ms) {
    gvizFetch.chyby = [];
    return Promise.all(LISTY.map(function (n) { return gvizFetch(n, ms); }))
      .catch(function (e2) {
        return Promise.all(LISTY.map(function (n) { return gvizList(TABULKA, n, ms); })).catch(function (e3) {
          throw new Error((e2 && e2.message || e2) + ' / ' + (e3 && e3.message || e3));
        });
      })
      .then(function (listy) {
        listy = listy.filter(Boolean);
        var ok = listy.some(function (l) { return l.rows[0] && l.rows[0][0] === 'id'; });
        if (!ok && gvizFetch.chyby.length) throw new Error('Odpověď tabulky: ' + gvizFetch.chyby.slice(0, 2).join(' / '));
        return listy;
      });
  }

  function zTabulky(listy) {
    var out = {};
    listy.forEach(function (l) {
      var n = l.name, rows = l.rows;
      if (!rows.length) return;
      var head = rows[0].map(function (h) { return String(h == null ? '' : h).trim(); });
      if (head[0] !== 'id') return;
      var type = TYPY[n.trim()] || (head.indexOf('guide') >= 0 ? 'guides' : head.indexOf('value') >= 0 && head.length <= 4 ? 'settings' : n.trim());
      rows = rows.slice(1).filter(function (r) { return r[0] !== '' && r[0] != null; });
      var json = {};
      head.forEach(function (c, i) {
        if (!c) return;
        var vals = rows.map(function (r) { return r[i]; }).filter(function (v) { return v !== '' && v != null; });
        json[c] = JSON_SL[c] || (vals.length > 0 && vals.every(function (v) { return typeof v === 'string' && /^\s*[\[{]/.test(v); }));
        if (c === 'number' || c === 'order') rows.forEach(function (r) { if (typeof r[i] === 'number' && r[i] === Math.floor(r[i])) r[i] = String(r[i]); });
      });
      out[type] = rows.map(function (r) {
        var rec = {};
        head.forEach(function (c, i) {
          if (!c) return;
          var v = r[i];
          if (json[c]) { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
          else if (c === 'ailmentIds' || /Ids$/.test(c)) v = String(v == null ? '' : v).split(',').map(function (x) { return x.trim(); }).filter(String);
          else if (c === 'baseAmount') v = typeof v === 'number' ? v : Number(String(v).replace(/\s/g, '').replace(',', '.')) || 0;
          else v = v == null ? '' : String(v);
          rec[c] = v;
        });
        return rec;
      });
    });
    return out;
  }

  function otisk(str) {
    var h = 2166136261;
    for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
    return (h >>> 0).toString(36) + '-' + str.length;
  }

  function stahnoutXlsx(ms) {
    var ctl = window.AbortController ? new AbortController() : null;
    var t = ctl && setTimeout(function () { ctl.abort(); }, ms || 90000);
    var url = ZDROJ + (ZDROJ.indexOf('?') < 0 ? '?' : '&') + 't=' + Date.now();
    return nactiXlsx().then(function () {
      return fetch(url, { signal: ctl && ctl.signal, cache: 'no-store' });
    }).then(function (r) {
      if (!r.ok) throw new Error('Tabulka není dostupná (' + r.status + '). Je zveřejněná na webu?');
      return r.arrayBuffer();
    }, function (e) {
      throw new Error(e && e.name === 'AbortError' ? 'Tabulka neodpověděla včas.' : 'Prohlížeč nepustil stažení tabulky (' + ((e && e.message) || e) + ').');
    }).then(function (buf) {
      return zXlsx(XLSX.read(new Uint8Array(buf), { type: 'array' }));
    }).finally(function () { if (t) clearTimeout(t); });
  }

  function stahnoutPrimo(ms) {
    if (navigator.onLine === false) return Promise.reject(new Error('Tablet je bez internetu.'));
    var cesty = [];
    if (TABULKA) cesty.push(stahnoutGviz);
    if (ZDROJ) cesty.push(stahnoutXlsx);
    var chyby = [];
    var zkus = function (i) {
      if (i >= cesty.length) return Promise.reject(new Error(chyby.join(' | ') || 'Chybí odkaz na tabulku.'));
      return cesty[i](ms).then(function (listy) {
        var data = zTabulky(listy);
        if (!data.herbs && !data.recipes) throw new Error('Z tabulky se nepodařilo přečíst listy Byliny a Recepty.');
        return data;
      }).catch(function (e) { chyby.push((e && e.message) || String(e)); return zkus(i + 1); });
    };
    return zkus(0).then(function (data) { return { verze: otisk(JSON.stringify(data)), data: data }; });
  }

  /** Stáhne data; vrátí {zmena:true/false, data}. Uloží je do tabletu. */
  function stahnout(ms) {
    var old = get(K_DATA);
    if (PRIMO) return stahnoutPrimo(ms).then(function (j) {
      var now = new Date().toISOString();
      if (old && old.verze === j.verze) { old.kontrola = now; put(K_DATA, old); return { zmena: false, data: old.data }; }
      if (!put(K_DATA, { verze: j.verze, data: j.data, stazeno: now, kontrola: now })) throw new Error('Data se do tabletu nevešla.');
      return { zmena: true, data: j.data };
    });
    return call('data', { verze: old && old.verze }, ms).then(function (j) {
      var now = new Date().toISOString();
      if (j.stejne && old) { old.kontrola = now; put(K_DATA, old); return { zmena: false, data: old.data }; }
      var rec = { verze: j.verze, data: j.data, stazeno: now, kontrola: now };
      if (!put(K_DATA, rec)) throw new Error('Data se do tabletu nevešla.');
      if (j.aplikace && j.aplikace.indexOf('</html>') > 0) put(K_UI, j.aplikace);
      return { zmena: true, data: j.data };
    });
  }

  function kdy(iso) {
    if (!iso) return '';
    var d = new Date(iso);
    return d.toLocaleDateString('cs-CZ', { day: 'numeric', month: 'numeric', year: 'numeric' }) + ' ' +
      d.toLocaleTimeString('cs-CZ', { hour: '2-digit', minute: '2-digit' });
  }
  function esc(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  /* =====================================================================
   * Režim A: běží uvnitř aplikace (aplikace.html) – náhrada google.script.run
   * ===================================================================== */
  if (window.BSR_UI) {
    var stav = { text: '', nove: false };
    var nacist = function () {
        var old = get(K_DATA);
        var ma = PRIMO || cfg();
        if (!ma || navigator.onLine === false) {
          if (old) { stav.text = navigator.onLine === false ? 'bez internetu' : ''; return Promise.resolve(old.data); }
          return Promise.reject(new Error(ma ? 'Data zatím nejsou stažená. Připojte tablet k internetu.' : 'Tablet není propojený s tabulkou.'));
        }
        var p = stahnout(old ? 120000 : 180000);
        if (!old) return p.then(function (r) { return r.data; });
        // Máme uložená data: počkáme chvíli na nová, jinak ukážeme uložená a nová nabídneme po stažení.
        return new Promise(function (res) {
          var done = false;
          var t = setTimeout(function () { done = true; stav.text = 'kontroluji novější data…'; bar(); res(old.data); }, 8000);
          p.then(function (r) {
            if (!done) { clearTimeout(t); done = true; res(r.data); }
            else if (r.zmena) { stav.nove = true; stav.text = ''; bar(); }
            else { stav.text = ''; bar(); }
          }, function () {
            stav.text = 'nová data se nepodařilo zkontrolovat';
            if (!done) { clearTimeout(t); done = true; res(old.data); } else bar();
          });
        });
    };
    var naText = function (html) {
      var d = new DOMParser().parseFromString(String(html)
        .replace(/<(style|script)[\s\S]*?<\/\1>/gi, '')
        .replace(/<\/t[dh]>/gi, '  ').replace(/<br\s*\/?>/gi, '\n').replace(/<li[^>]*>/gi, '\n• ')
        .replace(/<\/(p|div|h[1-6]|li|tr|table|ul|ol|section)>/gi, '\n'), 'text/html');
      return (d.body ? d.body.textContent : '').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    };
    var mail = function (to, name, html) {
      var a = document.createElement('a');
      a.href = 'mailto:' + encodeURIComponent(to) + '?subject=' + encodeURIComponent(name) + '&body=' + encodeURIComponent(naText(html));
      document.body.appendChild(a); a.click(); a.remove();
      return Promise.resolve(true);
    };
    var tisk = function (name, html) {
      var d = new DOMParser().parseFromString(String(html), 'text/html');
      var css = [].map.call(d.querySelectorAll('style'), function (x) { return x.textContent; }).join('\n').replace(/\bbody\b/g, ':host');
      var el = document.getElementById('bsr-print');
      if (!el) { el = document.createElement('div'); el.id = 'bsr-print'; document.body.appendChild(el); el.attachShadow({ mode: 'open' }); }
      el.shadowRoot.innerHTML = '<style>' + css + '</style>' + (d.body ? d.body.innerHTML : '');
      var t0 = document.title; document.title = name;
      setTimeout(function () { window.print(); document.title = t0; }, 50);
      return Promise.resolve({ url: '#', name: name });
    };
    var H = {
      getAll: function () { return nacist().then(function (r) { setTimeout(bar, 0); return r; }); },
      sendHtmlEmail: function (to, name, html) { if (PRIMO) return mail(to, name, html); return call('email', { email: to, name: name, html: html }).then(function () { return true; }); },
      makePdfHtml: function (name, html) { if (PRIMO) return tisk(name, html); return call('pdf', { name: name, html: html }).then(function (j) { return { url: j.url, name: j.name }; }); },
      saveRecord: function () { return Promise.reject(new Error('na tabletu se jen čte. Upravujte v tabulce nebo v aplikaci na počítači.')); },
      deleteRecord: function () { return Promise.reject(new Error('na tabletu se jen čte. Upravujte v tabulce nebo v aplikaci na počítači.')); }
    };
    var runner = function (ok, fail) {
      return new Proxy({}, {
        get: function (_, k) {
          if (k === 'withSuccessHandler') return function (f) { return runner(f, fail); };
          if (k === 'withFailureHandler') return function (f) { return runner(ok, f); };
          if (k === 'withUserObject') return function () { return runner(ok, fail); };
          return function () {
            var a = arguments;
            Promise.resolve().then(function () {
              if (!H[k]) throw new Error('Tato funkce na tabletu není dostupná.');
              return H[k].apply(null, a);
            }).then(function (r) { if (ok) ok(r); }, function (e) { if (fail) fail(e); });
          };
        }
      });
    };
    window.google = { script: { run: runner(), host: { close: function () {} } } };

    /* Lišta se stářím dat, tlačítkem Aktualizovat a nastavením. */
    var bar = function () {
      var el = document.getElementById('bsr-bar');
      if (!el) {
        if (!document.body) return;
        el = document.createElement('div'); el.id = 'bsr-bar';
        document.body.insertBefore(el, document.body.firstChild);
        el.addEventListener('click', function (e) {
          if (e.target.closest('[data-bsr-upd]')) aktualizovat(e.target);
          if (e.target.closest('[data-bsr-show]')) location.reload();
          if (e.target.closest('[data-bsr-set]')) nastaveni();
        });
      }
      var d = get(K_DATA);
      var info = d ? 'Data z tabulky: ' + kdy(d.kontrola || d.stazeno) : 'Data zatím nejsou stažená';
      if (navigator.onLine === false) info += ' · bez internetu';
      else if (stav.text) info += ' · ' + stav.text;
      el.innerHTML = '<span>' + esc(info) + '</span>' +
        (stav.nove ? '<button type="button" class="bsr-b bsr-p" data-bsr-show>Zobrazit nová data</button>' : '') +
        '<button type="button" class="bsr-b" data-bsr-upd>Aktualizovat</button>' +
        (PRIMO ? '' : '<button type="button" class="bsr-b" data-bsr-set aria-label="Nastavení tabletu" title="Nastavení tabletu">⚙</button>');
    };
    var aktualizovat = function (btn) {
      if (navigator.onLine === false) { stav.text = ''; bar(); alert('Tablet je bez internetu. Data se aktualizují, až bude připojený.'); return; }
      btn.disabled = true; btn.textContent = 'Stahuji…';
      stahnout().then(function (r) {
        if (r.zmena) { location.reload(); return; }
        stav.text = 'data jsou aktuální'; stav.nove = false; bar();
      }, function (e) { stav.text = ''; bar(); alert('Aktualizace se nepodařila: ' + e.message); });
    };
    var nastaveni = function () {
      var c = cfg() || {};
      var ov = document.createElement('div'); ov.className = 'overlay';
      ov.innerHTML = '<form class="form"><h2>Nastavení tabletu</h2>' +
        '<p>Tablet čte data z vaší Google Tabulky. Upravovat se dá jen v tabulce nebo v aplikaci na počítači.</p>' +
        '<label class="bsr-l">Adresa tabletové implementace<input name="url" value="' + esc(c.url || '') + '" autocomplete="off"></label>' +
        '<label class="bsr-l">Klíč tabletu<input name="klic" value="' + esc(c.klic || '') + '" autocomplete="off"></label>' +
        '<div class="formbar"><button type="button" class="btn ghost" data-x>Zavřít</button><button class="btn primary">Uložit</button></div></form>';
      document.body.appendChild(ov);
      ov.addEventListener('click', function (e) { if (e.target === ov || e.target.closest('[data-x]')) ov.remove(); });
      ov.querySelector('form').addEventListener('submit', function (e) {
        e.preventDefault();
        var f = e.target;
        put(K_CFG, { url: f.url.value.trim(), klic: f.klic.value.trim() });
        location.reload();
      });
    };
    document.addEventListener('DOMContentLoaded', bar);
    window.addEventListener('online', bar);
    window.addEventListener('offline', bar);
    return;
  }

  /* =====================================================================
   * Režim B: úvodní stránka – propojení, offline režim, spuštění aplikace
   * ===================================================================== */
  // Odkaz z e-mailu: …#nastavit=<adresa>&klic=<klíč>
  var h = new URLSearchParams(location.hash.slice(1));
  if (h.get('nastavit') && h.get('klic')) {
    put(K_CFG, { url: h.get('nastavit'), klic: h.get('klic') });
    history.replaceState(null, '', location.pathname + location.search);
  }
  if ('serviceWorker' in navigator) {
    var mel = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(function () {});
    // Nová verze aplikace převzala řízení: jednou obnovit, ať běží nová.
    navigator.serviceWorker.addEventListener('controllerchange', function () {
      if (mel && !sessionStorage.getItem('bsr-obnoveno')) { try { sessionStorage.setItem('bsr-obnoveno', '1'); } catch (e) {} location.reload(); }
    });
  }
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});

  var PRIDAT = '<script>window.BSR_UI=1;(' + BSR.toString() + ')()<\/script>' +
    '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">' +
    '<link rel="manifest" href="manifest.webmanifest"><meta name="theme-color" content="#2e7d46">' +
    '<style>#add,[data-edit],[data-del]{display:none!important}' +
    '#bsr-bar{display:flex;flex-wrap:wrap;gap:6px 10px;align-items:center;justify-content:flex-end;' +
    'padding:6px max(16px,env(safe-area-inset-right,0px)) 6px max(16px,env(safe-area-inset-left,0px));' +
    'padding-top:calc(6px + env(safe-area-inset-top,0px));background:var(--accent-soft,#e6f3e9);color:var(--muted,#4f7f5e);font-size:14px}' +
    '#bsr-bar span{margin-right:auto}' +
    '.bsr-b{border:1px solid var(--line,#cfe3d5);background:#fff;color:var(--ink,#1d5c34);border-radius:999px;padding:5px 12px;font:inherit;cursor:pointer}' +
    '.bsr-p{background:var(--accent,#2e7d46);color:#fff;border-color:transparent}' +
    '.bsr-l{display:block;margin:12px 0}.bsr-l input{display:block;width:100%;margin-top:4px}' +
    '@media screen{#bsr-print{display:none}}@media print{body>*:not(#bsr-print){display:none!important}#bsr-print{display:block}}</style>';

  /* Texty u e-mailu a PDF, když je posílá tablet sám (bez Apps Scriptu). */
  var TEXTY = [
    ['PDF je uložené na Disku ve složce „Byliny s rozumem – PDF“. <a href="${esc(r.url)}" target="_blank" rel="noopener">Otevřít ${esc(r.name)}</a>', 'Otevřel se tisk. Vyberte tiskárnu, nebo „Uložit jako PDF“.'],
    ['PDF je uložené na Disku. <a href="${esc(r.url)}" target="_blank" rel="noopener">Otevřít ${esc(r.name)}</a>', 'Otevřel se tisk. Vyberte tiskárnu, nebo „Uložit jako PDF“.'],
    [";window.open(r.url,'_blank')", ''],
    ['Vytvářím PDF na Disku…', 'Připravuji tisk…'],
    ['Odesláno na ${to} (s PDF v příloze).', 'E-mail pro ${to} je připravený v poštovní aplikaci, stačí ho odeslat.'],
    ['Odesláno na ${to}: ', 'E-mail pro ${to} je připravený v poštovní aplikaci: '],
    [' v jednom e-mailu (s PDF v příloze).', '. Stačí ho odeslat.'],
    ['Odesílám…', 'Otevírám poštu…']
  ];

  function spustit(html) {
    html = html.replace(/<base[^>]*>/i, '');
    if (PRIMO) TEXTY.forEach(function (t) { html = html.split(t[0]).join(t[1]); });
    html = /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, function (m) { return m + PRIDAT; }) : PRIDAT + html;
    document.open(); document.write(html); document.close();
  }

  function uvod(chyba) {
    var box = document.getElementById('boot');
    box.innerHTML = '<img src="icon-192.png" alt=""><h1>Byliny s rozumem</h1>' +
      '<p class="muted">Vložte odkaz na vaši tabulku zveřejněnou na webu ' +
      '(v tabulce Soubor → Sdílet → Publikovat na webu → Celý dokument → Microsoft Excel). Odkaz končí <b>output=xlsx</b>.</p>' +
      (chyba ? '<p class="err">' + esc(chyba) + '</p>' : '') +
      '<form><label for="z">Odkaz na tabulku</label><input id="z" name="zdroj" value="' + esc(ZDROJ) + '" placeholder="https://docs.google.com/spreadsheets/d/e/…/pub?output=xlsx" autocomplete="off">' +
      '<button>Stáhnout data</button></form>';
    box.querySelector('form').addEventListener('submit', function (e) {
      e.preventDefault();
      var z = e.target.zdroj.value.trim();
      if (!/^https:\/\/docs\.google\.com\/spreadsheets\/.+/.test(z)) { uvod('Tohle nevypadá jako odkaz na Google Tabulku.'); return; }
      put(K_ZDROJ, z); ZDROJ = z; PRIMO = true;
      var b = e.target.querySelector('button'); b.disabled = true; b.textContent = 'Stahuji data…';
      stahnout(180000).then(start, function (err) { uvod(err.message); });
    });
  }

  function start() {
    if (!PRIMO && !cfg()) { uvod(); return; }
    var ui = PRIMO ? null : get(K_UI);
    if (ui) { spustit(ui); return; }
    fetch('aplikace.html').then(function (r) { return r.text(); }).then(spustit, function () {
      uvod('Aplikaci se nepodařilo načíst. Připojte tablet k internetu.');
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
