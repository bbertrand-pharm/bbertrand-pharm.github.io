/* Moteur commun : lecture des extractions, calculs, chiffrement et affichage du tableau de bord. */
(function () {
  'use strict';
  const E = (window.EFF = {});

  E.TITRE = 'Le pharmacien clinicien au service de l’efficience';
  E.SOUS_TITRE = 'Bon usage des médicaments et produits de santé';

  /* ---------- Formats ---------- */
  const fEur = new Intl.NumberFormat('fr-FR', {style: 'currency', currency: 'EUR', maximumFractionDigits: 0});
  const fNum = new Intl.NumberFormat('fr-FR', {maximumFractionDigits: 0});
  const fPct = new Intl.NumberFormat('fr-FR', {maximumFractionDigits: 1});
  E.eur = (v) => (v == null || isNaN(v) ? '—' : fEur.format(v));
  E.num = (v) => (v == null || isNaN(v) ? '—' : fNum.format(v));
  E.pct = (v) => (v == null || isNaN(v) ? '—' : fPct.format(v) + ' %');
  const sum = (arr, f) => arr.reduce((s, x) => s + (+f(x) || 0), 0);
  E.MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];

  /* ---------- Lecture Excel (SheetJS) ---------- */
  E.readWorkbook = async (file) => XLSX.read(await file.arrayBuffer(), {type: 'array'});
  const rowsOf = (wb) => {
    const out = [];
    for (const name of wb.SheetNames) out.push(...XLSX.utils.sheet_to_json(wb.Sheets[name], {header: 1, defval: null, blankrows: false}));
    return out;
  };
  const norm = (s) => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();

  /* Liste code + libellé (hors GHS, LPPR, correspondance UF) : 2 premières colonnes, en-tête ignoré. */
  E.parseList = (wb) => {
    const m = new Map();
    for (const r of rowsOf(wb)) {
      const code = parseInt(r[0], 10);
      if (!Number.isFinite(code)) continue;
      if (!m.has(code)) m.set(code, String(r[1] == null ? '' : r[1]).trim());
    }
    return [...m.entries()].map(([code, libelle]) => ({code, libelle}));
  };

  /* Extraction « Analyse_conso » : repère les colonnes par leur nom. */
  E.parseConso = (wb) => {
    const rows = rowsOf(wb);
    const hi = rows.findIndex((r) => r.some((c) => norm(c) === 'code produit'));
    if (hi < 0) return rows.length < 2 ? {lignes: []} : {lignes: [], erreur: 'colonne « Code produit » introuvable'};
    const h = rows[hi].map(norm);
    const col = (name, from) => h.findIndex((c, i) => i >= (from || 0) && c === name);
    const cCode = col('code produit');
    const idx = {
      exercice: col('exercice'), uf: col('uf'), ufLib: col('libelle'), code: cCode, lib: col('libelle', cCode + 1),
      compte: col('compte de stock'), q: col('periodes/quantite'), v: col('periodes/valeur'), de: col('de la periode'), a: col('a la periode')
    };
    if (idx.q < 0) idx.q = col('annuel/quantite');
    if (idx.uf < 0 || idx.compte < 0) return {lignes: [], erreur: 'ce n\u2019est pas une extraction \u00ab Analyse_conso \u00bb (colonnes UF ou compte de stock absentes)'};
    if (idx.v < 0) idx.v = col('annuel/valeur');
    const lignes = [];
    let exercice = null, ufLib = null;
    for (const r of rows.slice(hi + 1)) {
      const code = parseInt(r[idx.code], 10);
      if (!Number.isFinite(code) || !Number.isFinite(parseInt(r[idx.uf], 10))) continue;
      exercice = exercice || r[idx.exercice];
      ufLib = ufLib || r[idx.ufLib];
      lignes.push({uf: parseInt(r[idx.uf], 10), code, libelle: String(r[idx.lib] || '').trim(), compte: String(r[idx.compte] || '').trim(),
        q: +r[idx.q] || 0, v: Math.round((+r[idx.v] || 0) * 100) / 100});
    }
    return {lignes, exercice: exercice ? +exercice : null, ufLib: ufLib ? String(ufLib).trim() : null};
  };

  /* ---------- Classements ---------- */
  E.famille = (compte) => {
    if (compte === 'H942110') return 'Solutés';
    if (compte.startsWith('H9421')) return 'Médicaments';
    if (compte.startsWith('H9422')) return 'DMS';
    return 'Autres';
  };
  E.FAMILLES = ['Médicaments', 'Solutés', 'DMS'];

  /* Regroupements saisis sous la forme « 1751e = 1751, 1753, 1754 » (une ligne par regroupement). */
  E.parseRegroupements = (txt) => {
    const out = [];
    for (const line of String(txt || '').split(/\n/)) {
      const m = line.match(/^\s*([^=:]+?)\s*[=:]\s*(.+)$/);
      if (!m) continue;
      const ufs = m[2].split(/[,;\s]+/).map((x) => parseInt(x, 10)).filter(Number.isFinite);
      if (ufs.length) out.push({code: m[1].trim(), ufs});
    }
    return out;
  };

  /* Jeu de données d'un service (chiffré tel quel) : configuration posée par l'administrateur
     + consommations brutes déposées par le service. */
  E.defaultPeriode = () => { const n = new Date(); return {exercice: n.getFullYear(), mois: Math.max(1, n.getMonth())}; };
  E.buildService = (svc, ref, github, prev) => {
    const regs = E.parseRegroupements(svc.regroupements);
    const ufNom = new Map((ref.uf || []).map((r) => [r.code, r.libelle]));
    const ufs = new Set(regs.flatMap((g) => g.ufs));
    prev = prev || {};
    const fichiers = {};
    Object.entries(prev.fichiers || {}).forEach(([u, f]) => { if (ufs.has(+u)) fichiers[u] = f; });
    return {
      version: 2, id: svc.id, service: svc.nom, genere: prev.genere || new Date().toISOString(),
      periode: prev.periode || E.defaultPeriode(),
      regroupements: regs.map((g) => ({code: g.code, ufs: g.ufs.map((u) => ({uf: u, nom: ufNom.get(u) || ('UF ' + u)}))})),
      ref: {horsGHS: (ref.horsGHS || []).map((r) => r.code), lppr: (ref.lppr || []).map((r) => r.code)},
      conso: (prev.conso || []).filter((l) => ufs.has(l.uf)),
      recues: (prev.recues || []).filter((u) => ufs.has(u)),
      fichiers, github
    };
  };
  E.serviceUFs = (d) => d.regroupements.flatMap((g) => g.ufs.map((u) => u.uf));

  /* Intègre des extractions dans le jeu d'un service : seules ses UF sont acceptées ; une UF déposée remplace ses lignes. */
  E.ingest = (d, results) => {
    const allowed = new Set(E.serviceUFs(d)), msgs = [];
    for (const {name, r} of results) {
      if (r.erreur) { msgs.push({ok: false, t: `${name} : ${r.erreur}`}); continue; }
      let ufs = [...new Set(r.lignes.map((l) => l.uf))];
      if (!r.lignes.length) {
        const m = name.match(/(\d{3,5})(?!.*\d{3,5})/);
        if (!m) { msgs.push({ok: false, t: `${name} : fichier vide, UF non identifiable`}); continue; }
        ufs = [+m[1]];
      }
      const hors = ufs.filter((u) => !allowed.has(u));
      if (hors.length) { msgs.push({ok: false, t: `${name} : UF ${hors.join(', ')} hors de ce service, fichier ignoré`}); continue; }
      d.conso = d.conso.filter((l) => !ufs.includes(l.uf)).concat(r.lignes);
      ufs.forEach((u) => { if (!d.recues.includes(u)) d.recues.push(u); d.fichiers[u] = {fichier: name, date: new Date().toISOString(), lignes: r.lignes.filter((l) => l.uf === u).length}; });
      if (r.exercice) d.periode.exercice = r.exercice;
      msgs.push({ok: true, t: `${name} : UF ${ufs.join(', ')}, ${r.lignes.length} lignes` + (r.lignes.length ? '' : ' (aucune consommation)')});
    }
    return msgs;
  };

  /* Lignes classées par famille (compte de stock) et financement (listes hors GHS / LPPR). */
  E.lignes = (d) => {
    const hg = new Set(d.ref.horsGHS), lp = new Set(d.ref.lppr), ufs = new Set(E.serviceUFs(d));
    return d.conso.filter((l) => ufs.has(l.uf)).map((l) => ({...l, famille: E.famille(l.compte), fin: hg.has(l.code) ? 'Hors GHS' : (lp.has(l.code) ? 'LPPR' : 'Intra-GHS')}));
  };

  /* ---------- Publication GitHub (API Contents) ---------- */
  E.github = (g) => {
    const [owner, repo] = String(g.repo || '').split('/');
    const b64utf8 = (s) => { const bytes = new TextEncoder().encode(s); let bin = ''; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(bin); };
    const call = async (method, path, body) => {
      const url = `https://api.github.com/repos/${owner}/${repo}/contents/${g.base || ''}${path}` + (method === 'GET' ? `?ref=${encodeURIComponent(g.branch || 'main')}` : '');
      const r = await fetch(url, {method, cache: 'no-store', headers: {Authorization: 'Bearer ' + g.token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28'}, body: body ? JSON.stringify(body) : undefined});
      if (method === 'GET' && r.status === 404) return null;
      if (!r.ok) {
        let m = ''; try { m = (await r.json()).message; } catch (e) {}
        const err = new Error(`${r.status} ${m}`); err.status = r.status; throw err;
      }
      return r.json();
    };
    const api = {
      get: (path) => call('GET', path),
      raw: async (path) => {
        const url = `https://api.github.com/repos/${owner}/${repo}/contents/${g.base || ''}${path}?ref=${encodeURIComponent(g.branch || 'main')}`;
        const r = await fetch(url, {cache: 'no-store', headers: {Authorization: 'Bearer ' + g.token, Accept: 'application/vnd.github.raw+json', 'X-GitHub-Api-Version': '2022-11-28'}});
        if (r.status === 404) return null;
        if (!r.ok) { const err = new Error(String(r.status)); err.status = r.status; throw err; }
        return r.text();
      },
      put: async (path, text, message) => { const cur = await call('GET', path); return call('PUT', path, {message, content: b64utf8(text), branch: g.branch || 'main', sha: cur ? cur.sha : undefined}); },
      del: async (path, message) => { const cur = await call('GET', path); if (cur) await call('DELETE', path, {message, sha: cur.sha, branch: g.branch || 'main'}); }
    };
    return api;
  };
  E.ghError = (e) => (e.status === 401 || e.status === 403 ? 'accès GitHub refusé : le jeton est invalide, expiré ou sans droit d’écriture sur le dépôt.' : (e.status === 409 ? 'publication simultanée en cours, réessayez dans un instant.' : e.message));
  /* Publie le jeu d'un service et met à jour sa date dans la liste publique. */
  E.publishService = async (d, password) => {
    const gh = E.github(d.github);
    d.genere = new Date().toISOString();
    await gh.put(`data/${d.id}.enc.json`, JSON.stringify(await E.encrypt(d, password)), `Efficience : mise à jour ${d.service}`);
    let m = {titre: E.TITRE, services: []};
    try { const cur = await gh.raw('data/manifest.json'); if (cur) m = JSON.parse(cur); } catch (e) {}
    m.services = (m.services || []).filter((s) => s.id !== d.id).concat({id: d.id, nom: d.service, maj: d.genere});
    m.maj = d.genere;
    await gh.put('data/manifest.json', JSON.stringify(m, null, 2), 'Efficience : liste des services');
  };

  /* ---------- Calculs ---------- */
  E.compute = (d) => {
    const F = 12 / d.periode.mois;
    const L = E.lignes(d);
    const recus = new Set([...(d.recues || []), ...d.conso.map((l) => l.uf)]);
    const RG = d.regroupements.map((g) => ({code: g.code, ufs: g.ufs.map((u) => ({...u, recue: recus.has(u.uf)}))}));
    const isIntra = (l) => l.fin === 'Intra-GHS';
    const tot = sum(L, (l) => l.v) * F, intra = sum(L.filter(isIntra), (l) => l.v) * F;
    const bloc = (rows) => {
      const s = (f) => Math.round(sum(rows.filter(f), (l) => l.v) * F);
      const o = {med: s((l) => isIntra(l) && l.famille === 'Médicaments'), sol: s((l) => isIntra(l) && l.famille === 'Solutés'),
        dms: s((l) => isIntra(l) && l.famille === 'DMS'), autres: s((l) => isIntra(l) && l.famille === 'Autres'), sus: s((l) => !isIntra(l))};
      o.total = o.med + o.sol + o.dms + o.autres + o.sus;
      return o;
    };
    const ufs = [];
    const regs = RG.map((g) => {
      const rows = L.filter((l) => g.ufs.some((u) => u.uf === l.uf));
      g.ufs.forEach((u) => ufs.push({uf: u.uf, nom: u.nom, recue: u.recue, reg: g.code, ...bloc(L.filter((l) => l.uf === u.uf))}));
      let cles = null;
      if (g.ufs.length > 1) {
        const intraG = rows.filter(isIntra);
        const T = (fam, rr) => sum(rr.filter((l) => !fam || l.famille === fam), (l) => l.v);
        const share = (x, t) => (t ? Math.round(x / t * 1000) / 10 : null);
        cles = g.ufs.map((u) => {
          const mine = intraG.filter((l) => l.uf === u.uf);
          return {uf: u.uf, nom: u.nom, recue: u.recue, med: share(T('Médicaments', mine), T('Médicaments', intraG)),
            sol: share(T('Solutés', mine), T('Solutés', intraG)), dms: share(T('DMS', mine), T('DMS', intraG)),
            glob: share(T(null, mine), T(null, intraG)), eur: Math.round(T(null, mine) * F)};
        });
      }
      return {code: g.code, ufs: g.ufs, nom: g.ufs.length === 1 ? g.ufs[0].nom : g.code + ' — ' + g.ufs.map((u) => u.uf).join(', '), cles, ...bloc(rows)};
    });
    const top = (perim, fam) => {
      const rows = L.filter((l) => isIntra(l) && l.famille === fam && (perim === '*' || regs.find((g) => g.code === perim).ufs.some((u) => u.uf === l.uf)));
      const m = new Map();
      for (const l of rows) {
        const p = m.get(l.code) || {code: l.code, libelle: l.libelle, q: 0, v: 0};
        p.q += l.q; p.v += l.v; m.set(l.code, p);
      }
      const prods = [...m.values()].filter((p) => p.v > 0).sort((a, b) => b.v - a.v);
      const t = sum(prods, (p) => p.v);
      let cum = 0;
      return prods.map((p, i) => {
        const prev = cum; cum += p.v;
        return {rang: i + 1, code: p.code, libelle: p.libelle, q: Math.round(p.q * F), v: Math.round(p.v * F), cumul: Math.round(cum / t * 1000) / 10,
          classe: prev / t < 0.8 ? 'A' : (prev / t < 0.95 ? 'B' : 'C')};
      });
    };
    return {F, total: Math.round(tot), intra: Math.round(intra), sus: Math.round(tot - intra), partIntra: tot ? Math.round(intra / tot * 1000) / 10 : null,
      refs: new Set(L.map((l) => l.code)).size, regs: regs.sort((a, b) => b.total - a.total), ufs, top};
  };

  /* ---------- Chiffrement (AES-GCM 256, clé dérivée PBKDF2-SHA256) ---------- */
  const ITER = 310000;
  const b64 = (buf) => { let s = ''; const a = new Uint8Array(buf); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); };
  const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
  const deriveKey = async (password, salt, iter) => {
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256'}, base, {name: 'AES-GCM', length: 256}, false, ['encrypt', 'decrypt']);
  };
  E.encrypt = async (obj, password) => {
    const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
    const key = await deriveKey(password, salt, ITER);
    const ct = await crypto.subtle.encrypt({name: 'AES-GCM', iv}, key, new TextEncoder().encode(JSON.stringify(obj)));
    return {format: 'eff-aesgcm-v1', kdf: 'PBKDF2-SHA256', iter: ITER, salt: b64(salt), iv: b64(iv), data: b64(ct)};
  };
  E.decrypt = async (env, password) => {
    const key = await deriveKey(password, unb64(env.salt), env.iter || ITER);
    const pt = await crypto.subtle.decrypt({name: 'AES-GCM', iv: unb64(env.iv)}, key, unb64(env.data));
    return JSON.parse(new TextDecoder().decode(pt));
  };
  E.genPassword = () => {
    const alpha = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    const r = crypto.getRandomValues(new Uint32Array(15));
    const s = [...r].map((x) => alpha[x % alpha.length]).join('');
    return s.slice(0, 5) + '-' + s.slice(5, 10) + '-' + s.slice(10);
  };
  E.slug = (s) => norm(s).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);

  /* ---------- Affichage ---------- */
  const el = (tag, attrs, ...kids) => {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === 'class') n.className = v; else if (k === 'text') n.textContent = v; else if (k.startsWith('on')) n.addEventListener(k.slice(2), v); else if (v != null) n.setAttribute(k, v);
    }
    for (const k of kids.flat()) if (k != null) n.append(k.nodeType ? k : document.createTextNode(String(k)));
    return n;
  };
  E.el = el;
  const COLORS = ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)'];
  const SERIES = [['med', 'Médicaments intra-GHS'], ['sol', 'Solutés'], ['dms', 'DMS intra-GHS'], ['sus', 'Hors GHS / LPPR']];

  const table = (cols, rows, total) => {
    const t = el('table', {class: 'eff-table'},
      el('thead', {}, el('tr', {}, cols.map((c) => el('th', {class: c.n ? 'n' : null, text: c.label})))),
      el('tbody', {}, rows.map((r) => el('tr', {class: r._muted ? 'muted' : null}, cols.map((c) => el('td', {class: c.n ? 'n' : null, text: c.fmt ? c.fmt(r[c.key], r) : (r[c.key] == null ? '—' : r[c.key])}))))));
    if (total) t.append(el('tfoot', {}, el('tr', {}, cols.map((c) => el('td', {class: c.n ? 'n' : null, text: total[c.key] == null ? '' : total[c.key]})))));
    return el('div', {class: 'eff-tw'}, t);
  };

  const stackedBars = (items) => {
    const wrap = el('div', {class: 'eff-chart'});
    const tip = el('div', {class: 'eff-tip', hidden: ''});
    const draw = () => {
      wrap.querySelectorAll('svg').forEach((s) => s.remove());
      const W = Math.max(280, wrap.clientWidth || 600), narrow = W < 560, rowH = narrow ? 52 : 40, lw = narrow ? 0 : 230, rw = 86;
      const H = items.length * rowH + 8, max = Math.max(1, ...items.map((i) => i.total));
      const ns = 'http://www.w3.org/2000/svg';
      const svg = document.createElementNS(ns, 'svg');
      svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('width', '100%'); svg.setAttribute('role', 'img');
      svg.setAttribute('aria-label', 'Coûts annualisés empilés par famille de produits');
      const x = (v) => lw + (W - lw - rw) * v / max;
      const add = (tag, a, txt) => { const n = document.createElementNS(ns, tag); for (const [k, v] of Object.entries(a)) n.setAttribute(k, v); if (txt != null) n.textContent = txt; svg.append(n); return n; };
      items.forEach((it, i) => {
        const y = i * rowH + (narrow ? 18 : 6);
        add('text', {x: 0, y: narrow ? y - 5 : y + 17, class: 'lbl'}, it.nom.length > 34 && !narrow ? it.nom.slice(0, 33) + '…' : it.nom);
        let acc = 0;
        SERIES.forEach(([k, label], j) => {
          const w = x(acc + it[k]) - x(acc);
          if (w > 0) {
            const r = add('rect', {x: x(acc), y, width: w, height: 24, rx: 2, style: 'fill:' + COLORS[j]});
            r.addEventListener('mousemove', (ev) => {
              const b = wrap.getBoundingClientRect();
              tip.hidden = false; tip.textContent = it.nom + ' · ' + label + ' : ' + E.eur(it[k]);
              tip.style.left = Math.min(ev.clientX - b.left + 12, b.width - 220) + 'px'; tip.style.top = (ev.clientY - b.top + 14) + 'px';
            });
            r.addEventListener('mouseleave', () => { tip.hidden = true; });
          }
          acc += it[k];
        });
        add('text', {x: x(it.total) + 6, y: y + 17, class: 'val'}, E.eur(it.total));
      });
      wrap.prepend(svg);
    };
    wrap.append(tip);
    requestAnimationFrame(draw);
    if (window.ResizeObserver) { let w0 = 0; new ResizeObserver(() => { if (Math.abs(wrap.clientWidth - w0) > 4) { w0 = wrap.clientWidth; draw(); } }).observe(wrap); }
    const leg = el('div', {class: 'eff-leg'}, SERIES.map(([, label], j) => el('span', {}, el('i', {style: 'background:' + COLORS[j]}), label)));
    return el('div', {}, wrap, leg);
  };

  E.renderDashboard = (root, d) => {
    const c = E.compute(d);
    const p = d.periode;
    root.replaceChildren();
    const finMois = E.MOIS[p.mois - 1];
    root.append(
      el('header', {class: 'eff-head'},
        el('p', {class: 'eff-kicker', text: E.TITRE}),
        el('h1', {text: d.service}),
        el('p', {class: 'eff-sub', text: `Exercice ${p.exercice} · consommations arrêtées de janvier à ${finMois} (${p.mois} mois), extrapolées sur 12 mois · valorisation au PMP`}),
        el('p', {class: 'eff-lead'}, 'Sur une année, le service consommerait ', el('b', {text: E.eur(c.total)}), ' de produits pharmaceutiques, dont ',
          el('b', {text: E.pct(c.partIntra)}), ' à la charge de l’hôpital (intra-GHS).', c.regs.length > 1 ? [' Premier poste : ', el('b', {text: c.regs[0].nom}), '.'] : null)),
      el('section', {class: 'eff-tiles'},
        [['Coût annualisé total', E.eur(c.total)], ['À la charge de l’hôpital (intra-GHS)', E.eur(c.intra)], ['Remboursé en sus (hors GHS + LPPR)', E.eur(c.sus)], ['Références consommées', E.num(c.refs)]]
          .map(([l, v]) => el('div', {class: 'eff-card'}, el('div', {class: 'eff-tl', text: l}), el('div', {class: 'eff-tv', text: v}))))
    );

    root.append(el('section', {id: 'couts'}, el('h2', {text: 'Coûts annualisés par ' + (c.regs.length > 1 ? 'regroupement' : 'UF')}),
      el('p', {class: 'eff-note', text: 'Médicaments, solutés et DMS à la charge de l’hôpital, et produits remboursés en sus du GHS (pris en charge par l’Assurance maladie).'}),
      stackedBars(c.regs.length > 1 ? c.regs : c.ufs.map((u) => ({...u, nom: u.uf + ' — ' + u.nom})))));

    const avecCles = c.regs.filter((g) => g.cles);
    if (avecCles.length) {
      const host = el('div');
      const sel = el('select', {class: 'eff-select', 'aria-label': 'Regroupement'}, avecCles.map((g) => el('option', {value: g.code, text: g.nom})));
      const show = () => {
        const g = avecCles.find((x) => x.code === sel.value);
        const manquantes = g.cles.filter((r) => !r.recue).map((r) => r.uf);
        const vides = g.cles.filter((r) => r.recue && !r.eur).map((r) => r.uf);
        host.replaceChildren(table([
          {key: 'uf', label: 'UF'}, {key: 'nom', label: 'Service'}, {key: 'med', label: 'Médicaments', n: 1, fmt: E.pct}, {key: 'sol', label: 'Solutés', n: 1, fmt: E.pct},
          {key: 'dms', label: 'DMS', n: 1, fmt: E.pct}, {key: 'glob', label: 'Global', n: 1, fmt: E.pct}, {key: 'eur', label: 'Montant annualisé', n: 1, fmt: E.eur}
        ], g.cles.map((r) => ({...r, _muted: !r.recue})), {uf: 'Total', med: '100 %', sol: '100 %', dms: '100 %', glob: '100 %', eur: E.eur(sum(g.cles, (r) => r.eur))}),
        (manquantes.length || vides.length) ? el('p', {class: 'eff-warn'}, [
          vides.length ? `UF sans consommation imputée sur la période : ${vides.join(', ')}. ` : '',
          manquantes.length ? `Extraction non transmise : ${manquantes.join(', ')}. ` : '',
          'Une UF approvisionnée sur le stock d’une autre apparaît à 0 % : pour le stock partagé, une clé fondée sur l’activité (journées, séjours, passages) peut être plus juste.'
        ].join('')) : null);
      };
      sel.addEventListener('change', show);
      root.append(el('section', {id: 'cles'}, el('h2', {text: 'Clés de répartition des UF regroupées'}),
        el('p', {class: 'eff-note', text: 'Part de chaque UF dans la consommation intra-GHS de son regroupement, par famille. Les produits remboursés en sus, rattachés au patient, sont exclus.'}),
        avecCles.length > 1 ? el('div', {class: 'eff-ctl'}, sel) : null, host));
      show();
    }

    const ufRows = c.ufs.map((u) => ({...u, _muted: !u.recue, nom: u.nom + (u.recue ? '' : ' (non transmise)')}));
    root.append(el('section', {id: 'detail-uf'}, el('h2', {text: 'Détail par UF'}), el('p', {class: 'eff-note', text: 'Montants annualisés.'}),
      table([
        {key: 'uf', label: 'UF'}, {key: 'nom', label: 'Service'}, c.regs.length > 1 ? {key: 'reg', label: 'Regroupement'} : null,
        {key: 'med', label: 'Médicaments', n: 1, fmt: E.eur}, {key: 'sol', label: 'Solutés', n: 1, fmt: E.eur}, {key: 'dms', label: 'DMS', n: 1, fmt: E.eur},
        {key: 'sus', label: 'Hors GHS / LPPR', n: 1, fmt: E.eur}, {key: 'total', label: 'Total', n: 1, fmt: E.eur}
      ].filter(Boolean), ufRows, {uf: 'Total', med: E.eur(sum(c.ufs, (u) => u.med)), sol: E.eur(sum(c.ufs, (u) => u.sol)), dms: E.eur(sum(c.ufs, (u) => u.dms)), sus: E.eur(sum(c.ufs, (u) => u.sus)), total: E.eur(c.total)})));

    const hostTop = el('div');
    const selP = el('select', {class: 'eff-select', 'aria-label': 'Périmètre'}, el('option', {value: '*', text: 'Tout le service'}), c.regs.length > 1 ? c.regs.map((g) => el('option', {value: g.code, text: g.nom})) : null);
    const selF = el('select', {class: 'eff-select', 'aria-label': 'Famille'}, E.FAMILLES.map((f) => el('option', {value: f, text: f})));
    const showTop = () => {
      const rows = c.top(selP.value, selF.value);
      hostTop.replaceChildren(rows.length ? table([
        {key: 'rang', label: 'Rang', n: 1}, {key: 'libelle', label: 'Produit'}, {key: 'q', label: 'Quantité / an', n: 1, fmt: E.num}, {key: 'v', label: 'Coût / an', n: 1, fmt: E.eur},
        {key: 'cumul', label: '% cumulé', n: 1, fmt: E.pct}, {key: 'classe', label: 'Classe'}
      ], rows.slice(0, 20)) : el('p', {class: 'eff-note', text: 'Aucune consommation intra-GHS pour ce choix.'}),
      rows.length ? el('p', {class: 'eff-note', text: `${rows.filter((r) => r.classe === 'A').length} produits de classe A sur ${rows.length} références ; les 20 premiers sont affichés.`}) : null);
    };
    selP.addEventListener('change', showTop); selF.addEventListener('change', showTop);
    root.append(el('section', {id: 'efficience'}, el('h2', {text: 'Produits à plus fort impact budgétaire'}),
      el('p', {class: 'eff-note', text: 'Produits intra-GHS classés par coût annualisé. Classe A : les produits qui cumulent les premiers 80 % de la dépense, cibles prioritaires des actions de pharmacie clinique (relais per os, alternatives, bon usage, conditionnements).'}),
      el('div', {class: 'eff-ctl'}, c.regs.length > 1 ? selP : null, selF), hostTop));
    showTop();

    root.append(el('section', {id: 'dotations'}, el('h2', {text: 'Dotations du service'}),
      el('p', {class: 'eff-note', text: 'Ce volet comparera la dotation du service à 15 jours de consommation, arrondie au conditionnement. Il sera ajouté avec l’état des dotations.'})));

    if (c.regs.length > 1) {
      const intraTot = sum(c.regs, (g) => g.med + g.sol + g.dms + g.autres);
      const part = (v, t) => (t ? Math.round(v / t * 1000) / 10 : null);
      const cards = c.regs.map((g) => {
        const gi = g.med + g.sol + g.dms + g.autres;
        const ufs = c.ufs.filter((u) => u.reg === g.code);
        return el('div', {class: 'eff-card eff-syn'},
          el('div', {class: 'eff-syn-h'}, el('b', {text: g.code}), el('span', {class: 'eff-pill', text: E.pct(part(g.total, c.total)) + ' du service'})),
          el('div', {class: 'eff-tv', text: E.eur(g.total)}),
          el('p', {class: 'eff-small', text: `dont ${E.eur(gi)} à la charge de l’hôpital et ${E.eur(g.sus)} remboursés en sus`}),
          el('div', {class: 'eff-syn-bar', role: 'img', 'aria-label': 'Répartition par famille'},
            SERIES.map(([k, label], j) => g.total && g[k] > 0 ? el('i', {style: `width:${g[k] / g.total * 100}%;background:${COLORS[j]}`, title: label + ' : ' + E.eur(g[k])}) : null)),
          el('ul', {class: 'eff-syn-list'}, ufs.map((u) => el('li', {class: u.recue ? null : 'muted'},
            el('span', {text: `${u.uf} · ${u.nom}`}),
            el('span', {class: 'n', text: ufs.length > 1 ? (u.recue ? `${E.eur(u.med + u.sol + u.dms + u.autres)} · ${E.pct(part(u.med + u.sol + u.dms + u.autres, gi))}` : 'non transmise') : (u.recue ? E.eur(u.total) : 'non transmise')})))),
          ufs.length > 1 ? el('p', {class: 'eff-small', text: 'Montant intra-GHS annualisé de l’UF et sa part dans le regroupement (clé globale).'}) : null);
      });
      root.append(el('section', {id: 'synthese-regroupements'}, el('h2', {text: 'Synthèse des regroupements'}),
        el('p', {class: 'eff-note', text: `${c.regs.length} regroupements, ${c.ufs.length} UF. Montants annualisés ; chaque carte rappelle les UF regroupées sous un même code et leur poids.`}),
        el('div', {class: 'eff-syn-grid'}, cards),
        el('div', {style: 'margin-top:14px'}, table([
          {key: 'code', label: 'Regroupement'}, {key: 'nb', label: 'UF', n: 1}, {key: 'med', label: 'Médicaments', n: 1, fmt: E.eur}, {key: 'sol', label: 'Solutés', n: 1, fmt: E.eur},
          {key: 'dms', label: 'DMS', n: 1, fmt: E.eur}, {key: 'sus', label: 'Hors GHS / LPPR', n: 1, fmt: E.eur}, {key: 'total', label: 'Total', n: 1, fmt: E.eur},
          {key: 'pIntra', label: 'Part de l’intra-GHS', n: 1, fmt: E.pct}
        ], c.regs.map((g) => ({...g, nb: g.ufs.length, pIntra: part(g.med + g.sol + g.dms + g.autres, intraTot)})),
        {code: 'Total', nb: c.ufs.length, med: E.eur(sum(c.regs, (g) => g.med)), sol: E.eur(sum(c.regs, (g) => g.sol)), dms: E.eur(sum(c.regs, (g) => g.dms)), sus: E.eur(c.sus), total: E.eur(c.total), pIntra: '100 %'}))));
    }

    root.append(el('details', {class: 'eff-method'}, el('summary', {text: 'Méthode'}), el('ul', {},
      el('li', {text: 'Source : extractions « Analyse_conso » du logiciel de gestion, une par UF, valorisées au PMP de la dernière balance définitive ; les retours sont déduits.'}),
      el('li', {text: `Annualisation : montant de la période × 12 / ${p.mois}.`}),
      el('li', {text: 'Familles selon le compte de stock : H942110 = solutés et électrolytes, autres H9421… = médicaments, H9422… = DMS.'}),
      el('li', {text: 'Financement : produit inscrit sur la liste des médicaments hors GHS ou des DMS LPPR = remboursé en sus ; sinon intra-GHS.'}),
      el('li', {text: 'Mise à jour : ' + new Date(d.genere).toLocaleDateString('fr-FR', {day: 'numeric', month: 'long', year: 'numeric'}) + '.'}))));
  };
})();
