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

  /** Stáhne data; vrátí {zmena:true/false, data}. Uloží je do tabletu. */
  function stahnout(ms) {
    var old = get(K_DATA);
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
        if (!cfg() || navigator.onLine === false) {
          if (old) { stav.text = navigator.onLine === false ? 'bez internetu' : ''; return Promise.resolve(old.data); }
          return Promise.reject(new Error(cfg() ? 'Data zatím nejsou stažená. Připojte tablet k internetu.' : 'Tablet není propojený s tabulkou.'));
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
    var H = {
      getAll: function () { return nacist().then(function (r) { setTimeout(bar, 0); return r; }); },
      sendHtmlEmail: function (to, name, html) { return call('email', { email: to, name: name, html: html }).then(function () { return true; }); },
      makePdfHtml: function (name, html) { return call('pdf', { name: name, html: html }).then(function (j) { return { url: j.url, name: j.name }; }); },
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
        '<button type="button" class="bsr-b" data-bsr-set aria-label="Nastavení tabletu" title="Nastavení tabletu">⚙</button>';
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
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(function () {});
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
    '.bsr-l{display:block;margin:12px 0}.bsr-l input{display:block;width:100%;margin-top:4px}</style>';

  function spustit(html) {
    html = html.replace(/<base[^>]*>/i, '');
    html = /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, function (m) { return m + PRIDAT; }) : PRIDAT + html;
    document.open(); document.write(html); document.close();
  }

  function uvod(chyba) {
    var c = cfg() || {};
    var box = document.getElementById('boot');
    box.innerHTML = '<img src="icon-192.png" alt=""><h1>Byliny s rozumem</h1>' +
      '<p class="muted">Propojte tablet s vaší Google Tabulkou. Nejjednodušší je otevřít na tabletu odkaz z e-mailu ' +
      '„Byliny s rozumem – nastavení tabletu“, který si pošlete z tabulky (menu Bylinky → Tablet: poslat odkaz pro nastavení). ' +
      'Nebo údaje vložte sem.</p>' + (chyba ? '<p class="err">' + esc(chyba) + '</p>' : '') +
      '<form><label for="u">Adresa tabletové implementace</label><input id="u" name="url" value="' + esc(c.url || '') + '" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off">' +
      '<label for="k">Klíč tabletu</label><input id="k" name="klic" value="' + esc(c.klic || '') + '" autocomplete="off">' +
      '<button>Propojit a stáhnout data</button></form>';
    box.querySelector('form').addEventListener('submit', function (e) {
      e.preventDefault();
      var f = e.target, b = f.querySelector('button');
      put(K_CFG, { url: f.url.value.trim(), klic: f.klic.value.trim() });
      b.disabled = true; b.textContent = 'Stahuji data… (může to trvat i minutu)';
      stahnout(180000).then(start, function (err) { uvod(err.message); });
    });
  }

  function start() {
    if (!get(K_DATA) && !cfg()) { uvod(); return; }
    var ui = get(K_UI);
    if (ui) { spustit(ui); return; }
    fetch('aplikace.html').then(function (r) { return r.text(); }).then(spustit, function () {
      uvod('Aplikaci se nepodařilo načíst. Připojte tablet k internetu.');
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
