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
      compte: col('compte de stock'), lieu: col('lieu de stockage'), q: col('periodes/quantite'), v: col('periodes/valeur'), de: col('de la periode'), a: col('a la periode')
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
      lignes.push({uf: parseInt(r[idx.uf], 10), code, libelle: String(r[idx.lib] || '').trim(), compte: String(r[idx.compte] || '').trim(), lieu: idx.lieu >= 0 ? String(r[idx.lieu] || '').trim() : undefined,
        q: +r[idx.q] || 0, v: Math.round((+r[idx.v] || 0) * 100) / 100});
    }
    return {lignes, exercice: exercice ? +exercice : null, ufLib: ufLib ? String(ufLib).trim() : null};
  };

  /* ---------- Classements ---------- */
  /* Solutés = poches et flacons de perfusion du compte H942110 (code « O » de la colonne « Lieu de stockage ») ;
     les ampoules et la nutrition parentérale de ce compte restent en médicaments. */
  const POCHES = /\b(POCHE|FLAC|FLACON|VIAFLO|VERSABLE)\b/i;
  const NUTRITION = /(OLIMEL|INTRALIPID|SMOFLIPID|KABIVEN|CLINIMIX|CLINOMEL|NUMETA|NUTRIFLEX|OMEGAVEN|NUTRYELT|CERNEVIT|SOLUVIT|VITALIPID)/i;
  E.famille = (compte, lieu, libelle) => {
    if (compte === 'H942110') {
      const poche = lieu !== undefined && lieu !== '' ? lieu === 'O' : POCHES.test(libelle || '');
      return poche && !NUTRITION.test(libelle || '') ? 'Solutés' : 'Médicaments';
    }
    if (compte.startsWith('H9421')) return lieu === 'L' ? 'Stupéfiants' : 'Médicaments';
    if (compte.startsWith('H9422')) return 'DMS';
    return 'Autres';
  };
  E.FAMILLES = ['Médicaments', 'Solutés', 'Stupéfiants', 'DMS'];
  /* Famille d'une référence de dotation jamais consommée : rangement, puis code (2xxxxx = DMS). */
  E.familleDotation = (x) => {
    const r = String(x.rangement || '').toUpperCase(), lib = String(x.libelle || '').toUpperCase();
    if (/STUP/.test(r)) return 'Stupéfiants';
    if (x.code >= 200000) return 'DMS';
    if (/NACL|GLUCOSE|BIONOLYTE|SOLUTE|GROS VOLUME|MANNITOL|BICARBONATE/.test(r) && /POCHE|FLAC/.test(lib) && !NUTRITION.test(lib)) return 'Solutés';
    return 'Médicaments';
  };

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
      fichiers, github, pistes: prev.pistes || {},
      dotations: Object.fromEntries(Object.entries(prev.dotations || {}).filter(([k]) => regs.some((g) => g.code === k)))
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

  /* ---------- Dotations ---------- */
  /* Fichier « Code / Rangement / Libellé / Dotation » (une ligne par référence). */
  E.parseDotation = (wb) => {
    const rows = rowsOf(wb);
    const hi = rows.findIndex((r) => r.some((c) => norm(c) === 'dotation') && r.some((c) => norm(c) === 'code'));
    if (hi < 0) return {lignes: [], erreur: 'colonnes « Code » et « Dotation » introuvables'};
    const h = rows[hi].map(norm), ix = (n) => h.indexOf(n);
    const iC = ix('code'), iR = ix('rangement'), iL = ix('libelle'), iD = ix('dotation');
    const m = new Map();
    for (const r of rows.slice(hi + 1)) {
      const code = parseInt(String(r[iC] == null ? '' : r[iC]).trim(), 10), dot = +r[iD];
      if (!Number.isFinite(code) || !Number.isFinite(dot)) continue;
      m.set(code, {code, rangement: iR >= 0 ? String(r[iR] || '').trim() : '', libelle: iL >= 0 ? String(r[iL] || '').trim() : '', dotation: dot});
    }
    return {lignes: [...m.values()]};
  };
  E.ingestDotation = (d, code, name, r) => {
    d.dotations = d.dotations || {};
    d.dotations[code] = {fichier: name, date: new Date().toISOString(), lignes: r.lignes};
  };
  /* Conditionnement lu dans le libellé (« CARTON DE 60 », « BTE/50 », « BT/100 », « 5/SACH ») ; 1 à défaut. */
  E.conditionnement = (lib) => {
    const s = String(lib || '').toUpperCase();
    const m = s.match(/(?:CARTON|BOITE|BOÎTE|BTE|BT|SACHET|PAQUET|CAISSE)\s*(?:DE\s*|\/\s*)(\d{1,4})/) || s.match(/(\d{1,4})\s*\/\s*SACH/);
    const n = m ? parseInt(m[1], 10) : 1;
    return n > 0 ? n : 1;
  };
  E.JOURS_DOTATION = 15;
  E.COUV_MAX = 30; E.COUV_MIN = 10;
  E.alerteCouv = (r) => (r.couv == null ? null : (r.couv > E.COUV_MAX ? 'sur' : (r.couv < E.COUV_MIN ? 'sous' : null)));
  E.STATUTS = ['À augmenter', 'À réduire', 'Sans consommation', 'Hors dotation', 'Adaptée'];

  /* Lignes classées par famille  /* Lignes classées par famille (compte de stock) et financement (listes hors GHS / LPPR). */
  E.lignes = (d) => {
    const hg = new Set(d.ref.horsGHS), lp = new Set(d.ref.lppr), ufs = new Set(E.serviceUFs(d));
    return d.conso.filter((l) => ufs.has(l.uf)).map((l) => ({...l, famille: E.famille(l.compte, l.lieu, l.libelle), fin: hg.has(l.code) ? 'Hors GHS' : (lp.has(l.code) ? 'LPPR' : 'Intra-GHS')}));
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
        stu: s((l) => isIntra(l) && l.famille === 'Stupéfiants'), dms: s((l) => isIntra(l) && l.famille === 'DMS'), autres: s((l) => isIntra(l) && l.famille === 'Autres'), sus: s((l) => !isIntra(l))};
      o.total = s(() => true);
      o.fam = {};
      for (const f of ['Médicaments', 'Solutés', 'Stupéfiants', 'DMS', 'Autres']) o.fam[f] = {total: s((l) => l.famille === f), sus: s((l) => l.famille === f && !isIntra(l))};
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
            sol: share(T('Solutés', mine), T('Solutés', intraG)), stu: share(T('Stupéfiants', mine), T('Stupéfiants', intraG)), dms: share(T('DMS', mine), T('DMS', intraG)),
            glob: share(T(null, mine), T(null, intraG)), eur: Math.round(T(null, mine) * F)};
        });
      }
      return {code: g.code, ufs: g.ufs, nom: g.ufs.length === 1 ? g.ufs[0].nom : g.ufs.map((u) => u.nom).join(', '), cles, ...bloc(rows)};
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
    /* Références d'un regroupement par coût décroissant : famille ('*' = toutes), avec ou sans les produits remboursés en sus. */
    const produits = (code, fam, avecSus) => {
      const g = regs.find((x) => x.code === code);
      const rows = L.filter((l) => g.ufs.some((u) => u.uf === l.uf) && (fam === '*' || l.famille === fam) && (avecSus || isIntra(l)));
      const m = new Map();
      for (const l of rows) {
        const p = m.get(l.code) || {code: l.code, libelle: l.libelle, famille: l.famille, fin: l.fin, q: 0, v: 0, ufs: new Set()};
        p.q += l.q; p.v += l.v; p.ufs.add(l.uf); m.set(l.code, p);
      }
      const prods = [...m.values()].filter((p) => p.v !== 0 || p.q !== 0).sort((a, b) => b.v - a.v || b.q - a.q);
      const t = sum(prods, (p) => p.v);
      let cum = 0;
      return prods.map((p, i) => {
        cum += p.v;
        return {rang: i + 1, code: p.code, libelle: p.libelle, famille: p.famille, fin: p.fin, ufs: [...p.ufs].sort().join(', '),
          q: Math.round(p.q * F), v: Math.round(p.v * F), vr: p.v * F, cumul: t ? Math.round(cum / t * 1000) / 10 : null};
      });
    };
    /* Adéquation dotation / consommation d'un regroupement : cible = 15 jours de consommation annualisée, arrondie au conditionnement supérieur. */
    const pu = new Map();
    { const acc = new Map(); for (const l of L) { const a = acc.get(l.code) || [0, 0]; if (l.q > 0 && l.v > 0) { a[0] += l.q; a[1] += l.v; } acc.set(l.code, a); }
      acc.forEach(([q, v], k) => { if (q > 0) pu.set(k, v / q); }); }
    const dotation = (code) => {
      const src = ((d.dotations || {})[code] || {}).lignes;
      if (!src) return null;
      const g = regs.find((x) => x.code === code);
      const conso = new Map();
      for (const l of L) if (g.ufs.some((u) => u.uf === l.uf)) { const p = conso.get(l.code) || {q: 0, libelle: l.libelle, famille: l.famille}; p.q += l.q; conso.set(l.code, p); }
      const J = E.JOURS_DOTATION, rows = [];
      const mk = (code, libelle, rangement, dot, famille) => {
        const c = conso.get(code), qAn = c ? Math.max(0, c.q * F) : 0, besoin = qAn * J / 365, condLu = E.conditionnement(libelle), cond = condLu <= Math.max(besoin, 1) ? condLu : 1;
        const cible = besoin > 0 ? Math.ceil(besoin / cond) * cond : 0;
        let statut;
        if (dot == null) statut = 'Hors dotation';
        else if (qAn <= 0) statut = 'Sans consommation';
        else if (Math.abs(dot - cible) <= Math.max(cond, 0.2 * cible)) statut = 'Adaptée';
        else statut = dot < cible ? 'À augmenter' : 'À réduire';
        const p = pu.get(code);
        return {code, libelle, rangement: rangement || '', famille: famille || (c && c.famille) || E.familleDotation({code, rangement, libelle}), dot, qAn: Math.round(qAn), besoin: Math.round(besoin * 10) / 10, cond, cible,
          ecart: dot == null ? cible : cible - dot, statut, couv: dot != null && qAn > 0 ? Math.round(dot / (qAn / 365) * 10) / 10 : null, pu: p == null ? null : p, valDot: p != null && dot != null ? dot * p : null, valCible: p != null ? cible * p : null};
      };
      const inDot = new Set();
      for (const x of src) { inDot.add(x.code); rows.push(mk(x.code, x.libelle || (conso.get(x.code) || {}).libelle, x.rangement, x.dotation)); }
      for (const [k, c] of conso) if (!inDot.has(k) && c.q * F * J / 365 >= 1) rows.push(mk(k, c.libelle, '', null, c.famille));
      const sumv = (f) => Math.round(sum(rows.filter((r) => r.dot != null && r.pu != null), f));
      return {rows, fichier: d.dotations[code].fichier, date: d.dotations[code].date, valActuelle: sumv((r) => r.valDot), valProposee: sumv((r) => (r.statut === 'Sans consommation' ? 0 : r.valCible)),
        sansPrix: rows.filter((r) => r.dot != null && r.pu == null).length};
    };
    return {F, total: Math.round(tot), intra: Math.round(intra), sus: Math.round(tot - intra), partIntra: tot ? Math.round(intra / tot * 1000) / 10 : null,
      refs: new Set(L.map((l) => l.code)).size, regs, ufs, top, produits, dotation};
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
    for (const k of kids.flat(Infinity)) if (k != null) n.append(k.nodeType ? k : document.createTextNode(String(k)));
    return n;
  };
  E.el = el;
  const COLORS = ['var(--c1)', 'var(--c2)', 'var(--c5)', 'var(--c3)', 'var(--c4)'];
  const SERIES = [['med', 'Médicaments intra-GHS'], ['sol', 'Solutés'], ['stu', 'Stupéfiants'], ['dms', 'DMS intra-GHS'], ['sus', 'Hors GHS / LPPR']];

  const table = (cols, rows, total) => {
    const t = el('table', {class: 'eff-table'},
      el('thead', {}, el('tr', {}, cols.map((c) => el('th', {class: c.n ? 'n' : null, text: c.label})))),
      el('tbody', {}, rows.map((r) => el('tr', {class: r._muted ? 'muted' : null}, cols.map((c) => c.render ? el('td', {class: c.n ? 'n' : null, 'data-k': c.key}, c.render(r)) : el('td', {class: c.n ? 'n' : null, 'data-k': c.key, text: c.fmt ? c.fmt(r[c.key], r) : (r[c.key] == null ? '—' : r[c.key])}))))));
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

  E.renderDashboard = (root, d, opts) => {
    opts = opts || {};
    const c = E.compute(d);
    const p = d.periode;
    const multi = c.regs.length > 1;
    const anchor = (g) => 'reg-' + E.slug(g.code);
    root.replaceChildren();
    root.append(
      el('header', {class: 'eff-head'},
        el('p', {class: 'eff-kicker', text: E.TITRE}),
        el('h1', {text: d.service}),
        el('p', {class: 'eff-sub', text: `Exercice ${p.exercice} · consommations arrêtées de janvier à ${E.MOIS[p.mois - 1]} (${p.mois} mois), extrapolées sur 12 mois · valorisation au PMP`}),
        el('p', {class: 'eff-lead'}, 'Sur une année, le service consommerait ', el('b', {text: E.eur(c.total)}), ' de produits pharmaceutiques, dont ',
          el('b', {text: E.pct(c.partIntra)}), ' à la charge de l’hôpital (intra-GHS) et ', el('b', {text: E.eur(c.sus)}), ' remboursés en sus du séjour.'),
        multi ? el('nav', {class: 'eff-tabs', role: 'tablist', 'aria-label': 'Affichage'},
          el('button', {type: 'button', role: 'tab', class: 'eff-tab', 'data-view': 'ensemble'}, el('b', {text: 'Vue d’ensemble'}), el('span', {text: E.eur(c.total)})),
          c.regs.map((g) => el('button', {type: 'button', role: 'tab', class: 'eff-tab', 'data-view': anchor(g)}, el('b', {text: g.code}), el('span', {text: (g.ufs.length === 1 ? g.ufs[0].nom : g.ufs.length + ' UF') + ' · ' + E.eur(g.total)})))) : null));
    const blocs = [];

    c.regs.forEach((g, gi) => {
      const sec = el('section', {class: 'eff-reg', id: anchor(g)});
      const titre = g.ufs.length === 1 ? `${g.code} — ${g.ufs[0].nom}` : `${g.code} — regroupement de ${g.ufs.length} UF`;
      sec.append(el('h2', {class: 'eff-reg-h', text: titre}),
        el('p', {class: 'eff-note'}, g.ufs.map((u, i) => [i ? ' · ' : '', el('span', {class: u.recue ? null : 'eff-off', text: `${u.uf} ${u.nom}` + (u.recue ? '' : ' (non transmise)')})])));

      /* 1. Coût global par type de produit */
      const famTile = (f, label, susLabel) => el('div', {class: 'eff-card'}, el('div', {class: 'eff-tl', text: label}), el('div', {class: 'eff-tv', text: E.eur(g.fam[f].total)}),
        el('div', {class: 'eff-small', text: g.fam[f].sus ? `dont ${E.eur(g.fam[f].sus)} ${susLabel}` : 'entièrement à la charge de l’hôpital'}));
      sec.append(el('h3', {text: '1. Coût global annualisé par type de produit'}),
        el('div', {class: 'eff-tiles'},
          el('div', {class: 'eff-card eff-card-main'}, el('div', {class: 'eff-tl', text: 'Coût global annualisé'}), el('div', {class: 'eff-tv', text: E.eur(g.total)}),
            el('div', {class: 'eff-small', text: `${E.eur(g.total - g.sus)} à la charge de l’hôpital · ${E.eur(g.sus)} remboursés en sus`})),
          famTile('Médicaments', 'Médicaments', 'hors GHS'), famTile('Solutés', 'Solutés', 'remboursés en sus'), famTile('Stupéfiants', 'Stupéfiants', 'remboursés en sus'), famTile('DMS', 'Dispositifs médicaux stériles', 'LPPR')),
        el('div', {class: 'eff-syn-bar', role: 'img', 'aria-label': 'Répartition du coût par type de produit'},
          SERIES.map(([k, label], j) => g.total && g[k] > 0 ? el('i', {style: `width:${g[k] / g.total * 100}%;background:${COLORS[j]}`, title: label + ' : ' + E.eur(g[k])}) : null)),
        el('div', {class: 'eff-leg'}, SERIES.map(([k, label], j) => el('span', {}, el('i', {style: 'background:' + COLORS[j]}), `${label} ${E.pct(g.total ? Math.round(g[k] / g.total * 1000) / 10 : null)}`))));

      /* 2. Références par coût décroissant */
      const selF = el('select', {class: 'eff-select', 'aria-label': 'Type de produit'}, [['*', 'Tous les produits'], ...E.FAMILLES.map((f) => [f, f])].map(([v, l]) => el('option', {value: v, text: l})));
      const chk = el('input', {type: 'checkbox', checked: ''});
      const info = el('p', {class: 'eff-small'}), host = el('div'), more = el('button', {class: 'eff-btn ghost', type: 'button'});
      let all = false;
      const N = 25;
      const showRefs = () => {
        const rows = c.produits(g.code, selF.value, chk.checked);
        const tot = Math.round(sum(rows, (r) => r.vr));
        info.textContent = rows.length ? `${rows.length} références · ${E.eur(tot)} par an` : '';
        host.replaceChildren(rows.length ? table([
          {key: 'rang', label: 'Rang', n: 1}, {key: 'libelle', label: 'Référence'}, {key: 'code', label: 'Code'},
          selF.value === '*' ? {key: 'famille', label: 'Type'} : null,
          chk.checked ? {key: 'fin', label: 'Financement', fmt: (v) => (v === 'Intra-GHS' ? 'Intra-GHS' : v + ' (en sus)')} : null,
          g.ufs.length > 1 ? {key: 'ufs', label: 'UF'} : null,
          {key: 'q', label: 'Quantité / an', n: 1, fmt: E.num}, {key: 'v', label: 'Coût / an', n: 1, fmt: E.eur},
          {key: 'pGlob', label: '% du total', n: 1, fmt: (v) => (v != null && v < 0.1 && v > 0 ? '<\u00a00,1\u00a0%' : E.pct(v))}, {key: 'cumul', label: '% cumulé', n: 1, fmt: E.pct}
        ].filter(Boolean), (all ? rows : rows.slice(0, N)).map((r) => ({...r, pGlob: tot ? Math.round(r.vr / tot * 1000) / 10 : null, _muted: r.fin !== 'Intra-GHS'}))) : el('p', {class: 'eff-small', text: 'Aucune référence pour ce choix.'}));
        more.hidden = rows.length <= N;
        more.textContent = all ? 'Afficher les 25 premières seulement' : `Afficher les ${rows.length} références`;
      };
      selF.addEventListener('change', showRefs); chk.addEventListener('change', showRefs);
      more.addEventListener('click', () => { all = !all; showRefs(); });
      sec.append(el('h3', {text: '2. Références par coût global décroissant'}),
        el('p', {class: 'eff-note', text: 'Les deux pourcentages se recalculent selon le filtre : % du total = part de la référence dans le total du tableau affiché (indiqué au-dessus du tableau) ; % cumulé = part atteinte de la première référence jusqu’à la ligne.'}),
        el('div', {class: 'eff-ctl'}, selF, el('label', {class: 'eff-check'}, chk, 'Inclure les références remboursées en sus du séjour (hors GHS, LPPR)')),
        info, host, el('div', {class: 'eff-row'}, more));
      showRefs();
      const notepadHost = el('div');
      const drawNotepad = () => {
        const pistes = (d.pistes || {})[g.code] || [];
        notepadHost.replaceChildren(pistes.length ? el('aside', {class: 'eff-notepad', 'aria-label': 'Pistes d’efficience'},
          el('p', {class: 'eff-notepad-h', text: 'Pistes d’efficience'}),
          el('p', {class: 'eff-notepad-sub', text: 'Idées à discuter avec l’équipe médicale et soignante, sans diminuer la qualité de la prise en charge. Montants et gains estimés à partir des consommations annualisées.'}),
          el('ul', {}, pistes.map((x) => el('li', {},
            el('b', {text: x.titre}), ' — ', x.texte,
            x.gain ? el('span', {class: 'eff-notepad-gain', text: 'Gain estimé : ' + x.gain}) : null,
            (x.sources || []).length ? el('span', {class: 'eff-notepad-src'}, 'Sources : ', x.sources.map((s, i) => [i ? ' · ' : '', el('a', {href: s.url, target: '_blank', rel: 'noopener', text: s.label})])) : null)))) : '');
      };
      drawNotepad();
      sec.append(notepadHost);
      if (opts.chat) sec.append(opts.chat({g, c, d, redraw: drawNotepad}));

      /* 3. Dotation */
      sec.append(el('h3', {text: '3. Adéquation de la dotation aux consommations annualisées'}));
      const dt = c.dotation(g.code);
      if (!dt) {
        sec.append(el('div', {class: 'eff-card eff-pending'},
          el('p', {text: 'Ce volet compare, référence par référence, la dotation de ' + (g.ufs.length === 1 ? 'l’UF' : 'ce regroupement') + ' à 15 jours de consommation annualisée, arrondie au conditionnement : à augmenter, à réduire, sans consommation, hors dotation.'}),
          el('p', {class: 'eff-small', text: 'En attente de l’état de la dotation (colonnes Code, Rangement, Libellé, Dotation), à déposer dans « Mettre à jour les données du service ».'})));
      } else {
        const cnt = (s) => dt.rows.filter((r) => r.statut === s).length;
        const nbDot = dt.rows.filter((r) => r.dot != null).length;
        sec.append(el('p', {class: 'eff-note', text: `Dotation de ${nbDot} références (${dt.fichier}). Cible : ${E.JOURS_DOTATION} jours de consommation annualisée, arrondie au conditionnement supérieur (lu dans le libellé : carton, boîte…), ou à l’unité quand le conditionnement dépasse le besoin de 15 jours. Écart toléré : un conditionnement ou 20 %.`}),
          el('div', {class: 'eff-tiles eff-tiles-dot'},
            [['À augmenter', 'warn'], ['À réduire', 'info'], ['Sans consommation', 'bad'], ['Hors dotation', ''], ['Adaptée', 'ok']].map(([s, k]) => el('div', {class: 'eff-card eff-st-' + (k || 'none')}, el('div', {class: 'eff-tl', text: s}), el('div', {class: 'eff-tv', text: E.num(cnt(s))})))),
          el('p', {class: 'eff-small'}, 'Valeur de la dotation : ', el('b', {text: E.eur(dt.valActuelle)}), ' aujourd’hui → ', el('b', {text: E.eur(dt.valProposee)}), ' avec les ajustements proposés (références sans consommation retirées)' + (dt.sansPrix ? ` ; ${dt.sansPrix} références sans prix connu non valorisées.` : '.')));
        const nSur = dt.rows.filter((r) => E.alerteCouv(r) === 'sur').length, nSous = dt.rows.filter((r) => E.alerteCouv(r) === 'sous').length;
        sec.append(el('div', {class: 'eff-couv'},
          el('div', {class: 'eff-couv-item eff-couv-sur'}, el('b', {text: E.num(nSur)}), ` références dont la dotation couvre plus d’un mois (> ${E.COUV_MAX} jours) de consommation moyenne`),
          el('div', {class: 'eff-couv-item eff-couv-sous'}, el('b', {text: E.num(nSous)}), ` références dont la dotation couvre moins de ${E.COUV_MIN} jours de consommation moyenne`)),
          el('p', {class: 'eff-small', text: 'Dans le tableau, ces références sont surlignées en bleu (plus d’un mois de stock) ou en rose (moins de 10 jours). Cliquez sur un bandeau pour les afficher seules.'}));
        const selS = el('select', {class: 'eff-select', 'aria-label': 'Statut'}, el('option', {value: 'actions', text: 'Ajustements proposés'}), el('option', {value: '*', text: 'Toutes les références'}),
          el('option', {value: 'sur', text: `Plus d’un mois de stock (> ${E.COUV_MAX} j)`}), el('option', {value: 'sous', text: `Moins de ${E.COUV_MIN} jours de stock`}), E.STATUTS.map((s) => el('option', {value: s, text: s})));
        const selFD = el('select', {class: 'eff-select', 'aria-label': 'Type de produit'}, [['*', 'Tous les produits'], ...E.FAMILLES.map((f) => [f, f])].map(([v, l]) => el('option', {value: v, text: l})));
        d.decisions = d.decisions || {};
        const dec = d.decisions[g.code] = d.decisions[g.code] || {};
        const decidee = (r) => (dec[r.code] != null ? dec[r.code] : (r.statut === 'Adaptée' ? r.dot : null));
        const exportD = el('button', {class: 'eff-btn ghost', type: 'button', text: 'Exporter la dotation décidée (CSV)'});
        exportD.addEventListener('click', () => {
          const q = (s) => '"' + String(s == null ? '' : s).replace(/"/g, '""') + '"';
          const lines = [['Code', 'Référence', 'Type', 'Rangement', 'Dotation actuelle', 'Consommation annuelle', 'Dotation proposée', 'Statut', 'Dotation décidée'].map(q).join(';')]
            .concat(dt.rows.map((r) => [r.code, r.libelle, r.famille, r.rangement, r.dot == null ? '' : r.dot, r.qAn, r.cible, r.statut, decidee(r) == null ? '' : decidee(r)].map(q).join(';')));
          const a = el('a', {href: URL.createObjectURL(new Blob(['\ufeff' + lines.join('\r\n')], {type: 'text/csv;charset=utf-8'})), download: `dotation-${g.code}-${new Date().toISOString().slice(0, 10)}.csv`});
          document.body.append(a); a.click(); a.remove();
        });
        const hostD = el('div'), infoD = el('p', {class: 'eff-small'}), moreD = el('button', {class: 'eff-btn ghost', type: 'button'});
        let allD = false;
        moreD.addEventListener('click', () => { allD = !allD; showD(); });
        const ordre = (s) => E.STATUTS.indexOf(s);
        const showD = () => {
          const v = selS.value;
          const rows = dt.rows.filter((r) => (selFD.value === '*' || r.famille === selFD.value) && (v === '*' || (v === 'actions' ? r.statut !== 'Adaptée' : (v === 'sur' || v === 'sous' ? E.alerteCouv(r) === v : r.statut === v))))
            .sort(v === 'sur' ? (a, b) => b.couv - a.couv : v === 'sous' ? (a, b) => a.couv - b.couv :
              (a, b) => ordre(a.statut) - ordre(b.statut) || Math.abs(b.ecart * (b.pu || 0)) - Math.abs(a.ecart * (a.pu || 0)) || b.qAn - a.qAn);
          infoD.textContent = `${rows.length} références`;
          hostD.classList.add('eff-dot-table');
          hostD.replaceChildren(rows.length ? table([
            {key: 'libelle', label: 'Référence'}, {key: 'rangement', label: 'Rangement'},
            {key: 'dot', label: 'Dotation actuelle', n: 1, fmt: (v) => (v == null ? '—' : E.num(v))}, {key: 'qAn', label: 'Conso / an', n: 1, fmt: E.num},
            {key: 'cond', label: 'Cond.', n: 1}, {key: 'cible', label: 'Dotation proposée', n: 1, fmt: E.num},
            {key: 'statut', label: 'Statut', fmt: (v) => v},
            {key: 'decidee', label: 'Dotation décidée', n: 1, render: (r) => {
              const val = decidee(r), pre = dec[r.code] == null && val != null;
              if (!opts.editable) return el('span', {class: pre ? 'eff-pre' : null, text: val == null ? '—' : E.num(val)});
              const inp = el('input', {class: 'eff-dec' + (pre ? ' eff-pre' : ''), type: 'number', min: '0', step: '1', inputmode: 'numeric', value: val == null ? '' : String(val), 'aria-label': 'Dotation décidée pour ' + r.libelle,
                title: pre ? 'Prérempli : dotation actuelle jugée adaptée' : ''});
              inp.addEventListener('change', () => {
                const n = inp.value === '' ? null : Math.max(0, Math.round(+inp.value));
                if (n == null) delete dec[r.code]; else dec[r.code] = n;
                inp.classList.remove('eff-pre'); inp.title = '';
                if (opts.onChange) opts.onChange();
              });
              return inp;
            }}
          ], (allD ? rows : rows.slice(0, 25)).map((r) => ({...r, _muted: r.statut === 'Adaptée'}))) : el('p', {class: 'eff-small', text: 'Aucune référence pour ce choix.'}));
          moreD.hidden = rows.length <= 25; moreD.textContent = allD ? 'Afficher les 25 premières seulement' : `Afficher les ${rows.length} références`;
          hostD.querySelectorAll('tbody tr').forEach((tr, i) => { const al = E.alerteCouv(rows[i]); if (al) tr.classList.add('eff-row-' + al); const td = tr.querySelector('td[data-k=statut]'); if (td) td.className = 'eff-st eff-st-' + ({'À augmenter': 'warn', 'À réduire': 'info', 'Sans consommation': 'bad', 'Hors dotation': 'none', 'Adaptée': 'ok'})[rows[i].statut]; });
        };
        selS.addEventListener('change', () => { allD = false; showD(); });
        selFD.addEventListener('change', () => { allD = false; showD(); });
        sec.querySelectorAll('.eff-couv-item').forEach((it) => { it.addEventListener('click', () => { selS.value = it.classList.contains('eff-couv-sur') ? 'sur' : 'sous'; allD = false; showD(); }); });
        const nDec = () => dt.rows.filter((r) => decidee(r) != null).length;
        sec.append(el('div', {class: 'eff-ctl'}, selFD, selS), infoD, hostD, el('div', {class: 'eff-row'}, moreD, exportD),
          el('p', {class: 'eff-small', text: 'Dotation décidée : préremplie avec la dotation actuelle quand elle est jugée adaptée (en italique) ; à saisir pour les autres références' + (opts.editable ? ', puis « Publier » en bas de page pour l’enregistrer.' : '.')}),
          el('p', {class: 'eff-small', text: 'À relire avec l’équipe avant toute modification : un produit d’urgence ou à rotation lente peut rester en dotation même sans consommation récente ; « Hors dotation » = référence consommée au moins une fois par quinzaine en moyenne mais absente de la dotation (à ajouter ou à laisser en délivrance nominative).'}));
        showD();
      }
      blocs.push(sec);
    });

    const globalSec = el('section', {class: 'eff-reg eff-global', id: 'ensemble'});
    if (c.regs.length > 1) {
      const intraTot = sum(c.regs, (g) => g.med + g.sol + g.stu + g.dms + g.autres);
      const famG = (f) => ({total: sum(c.regs, (g) => g.fam[f].total), sus: sum(c.regs, (g) => g.fam[f].sus)});
      const G = {total: c.total, sus: c.sus, med: sum(c.regs, (g) => g.med), sol: sum(c.regs, (g) => g.sol), stu: sum(c.regs, (g) => g.stu), dms: sum(c.regs, (g) => g.dms)};
      const famTileG = (f, label, susLabel) => { const v = famG(f); return el('div', {class: 'eff-card'}, el('div', {class: 'eff-tl', text: label}), el('div', {class: 'eff-tv', text: E.eur(v.total)}),
        el('div', {class: 'eff-small', text: v.sus ? `dont ${E.eur(v.sus)} ${susLabel}` : 'entièrement à la charge de l’hôpital'})); };
      globalSec.append(el('h2', {class: 'eff-reg-h', text: 'Vue d’ensemble du service'}),
        el('p', {class: 'eff-note', text: `${c.regs.length} UF ou regroupements, ${c.ufs.length} UF · ${E.num(c.refs)} références consommées.`}),
        el('h3', {text: '1. Bilan financier global annualisé'}),
        el('div', {class: 'eff-tiles'},
          el('div', {class: 'eff-card eff-card-main'}, el('div', {class: 'eff-tl', text: 'Coût global annualisé'}), el('div', {class: 'eff-tv', text: E.eur(c.total)}),
            el('div', {class: 'eff-small', text: `${E.eur(c.intra)} à la charge de l’hôpital (${E.pct(c.partIntra)}) · ${E.eur(c.sus)} remboursés en sus`})),
          famTileG('Médicaments', 'Médicaments', 'hors GHS'), famTileG('Solutés', 'Solutés', 'remboursés en sus'), famTileG('Stupéfiants', 'Stupéfiants', 'remboursés en sus'), famTileG('DMS', 'Dispositifs médicaux stériles', 'LPPR')),
        el('div', {class: 'eff-syn-bar', role: 'img', 'aria-label': 'Répartition du coût global par type de produit'},
          SERIES.map(([k, label], j) => G.total && G[k] > 0 ? el('i', {style: `width:${G[k] / G.total * 100}%;background:${COLORS[j]}`, title: label + ' : ' + E.eur(G[k])}) : null)),
        el('div', {class: 'eff-leg'}, SERIES.map(([k, label], j) => el('span', {}, el('i', {style: 'background:' + COLORS[j]}), `${label} ${E.pct(G.total ? Math.round(G[k] / G.total * 1000) / 10 : null)}`))),
        el('h3', {text: '2. Comparaison des ' + c.regs.length + ' UF ou regroupements'}),
        el('p', {class: 'eff-note', text: 'Coûts annualisés empilés par type de produit. Cliquez sur un onglet, ou sur une carte plus bas, pour afficher le détail d’une UF ou d’un regroupement.'}),
        stackedBars(c.regs.map((g) => ({...g, nom: g.code + ' — ' + (g.ufs.length === 1 ? g.ufs[0].nom : g.ufs.length + ' UF')}))));
      const part = (v, t) => (t ? Math.round(v / t * 1000) / 10 : null);
      const cards = c.regs.map((g) => {
        const gi = g.med + g.sol + g.stu + g.dms + g.autres;
        const ufs = c.ufs.filter((u) => u.reg === g.code);
        return el('div', {class: 'eff-card eff-syn'},
          el('div', {class: 'eff-syn-h'}, el('b', {text: g.code}), el('span', {class: 'eff-pill', text: E.pct(part(g.total, c.total)) + ' du service'})),
          el('div', {class: 'eff-tv', text: E.eur(g.total)}),
          el('p', {class: 'eff-small', text: `dont ${E.eur(gi)} à la charge de l’hôpital et ${E.eur(g.sus)} remboursés en sus`}),
          el('div', {class: 'eff-syn-bar', role: 'img', 'aria-label': 'Répartition par famille'},
            SERIES.map(([k, label], j) => g.total && g[k] > 0 ? el('i', {style: `width:${g[k] / g.total * 100}%;background:${COLORS[j]}`, title: label + ' : ' + E.eur(g[k])}) : null)),
          el('ul', {class: 'eff-syn-list'}, ufs.map((u) => el('li', {class: u.recue ? null : 'muted'},
            el('span', {text: `${u.uf} · ${u.nom}`}),
            el('span', {class: 'n', text: ufs.length > 1 ? (u.recue ? `${E.eur(u.med + u.sol + u.stu + u.dms + u.autres)} · ${E.pct(part(u.med + u.sol + u.stu + u.dms + u.autres, gi))}` : 'non transmise') : (u.recue ? E.eur(u.total) : 'non transmise')})))),
          ufs.length > 1 ? el('p', {class: 'eff-small', text: 'Montant intra-GHS annualisé de l’UF et sa part dans le regroupement (clé globale).'}) : null);
      });
      globalSec.append(el('div', {style: 'margin-top:14px'}, table([
          {key: 'code', label: 'Regroupement'}, {key: 'nb', label: 'UF', n: 1}, {key: 'med', label: 'Médicaments intra-GHS', n: 1, fmt: E.eur}, {key: 'sol', label: 'Solutés', n: 1, fmt: E.eur}, {key: 'stu', label: 'Stupéfiants', n: 1, fmt: E.eur},
          {key: 'dms', label: 'DMS intra-GHS', n: 1, fmt: E.eur}, {key: 'sus', label: 'Hors GHS / LPPR', n: 1, fmt: E.eur}, {key: 'total', label: 'Total', n: 1, fmt: E.eur},
          {key: 'pIntra', label: 'Part de l’intra-GHS', n: 1, fmt: E.pct}
        ], c.regs.map((g) => ({...g, nb: g.ufs.length, pIntra: part(g.med + g.sol + g.stu + g.dms + g.autres, intraTot)})),
        {code: 'Total', nb: c.ufs.length, med: E.eur(sum(c.regs, (g) => g.med)), sol: E.eur(sum(c.regs, (g) => g.sol)), stu: E.eur(sum(c.regs, (g) => g.stu)), dms: E.eur(sum(c.regs, (g) => g.dms)), sus: E.eur(c.sus), total: E.eur(c.total), pIntra: '100 %'})),
        el('h3', {text: '3. Bilan du regroupement des UF'}),
        el('p', {class: 'eff-note', text: 'Chaque carte rappelle les UF regroupées sous un même code et leur poids ; cliquez pour ouvrir le détail.'}),
        el('div', {class: 'eff-syn-grid'}, cards));
      cards.forEach((cd, i) => { cd.classList.add('eff-click'); cd.addEventListener('click', () => { show(anchor(c.regs[i])); window.scrollTo({top: 0, behavior: 'smooth'}); }); });
    }

    if (c.regs.length > 1) c.regs.filter((g) => g.cles).forEach((g) => {
      globalSec.append(el('h3', {text: `Clés de répartition par type de produit — ${g.code}`}),
        el('p', {class: 'eff-note', text: 'Part de chaque UF dans la consommation intra-GHS du regroupement. Une UF approvisionnée sur le stock d’une autre apparaît à 0 % : pour le stock partagé, une clé fondée sur l’activité (journées, séjours, passages) peut être plus juste.'}),
        table([
          {key: 'uf', label: 'UF'}, {key: 'nom', label: 'Service'}, {key: 'med', label: 'Médicaments', n: 1, fmt: E.pct}, {key: 'sol', label: 'Solutés', n: 1, fmt: E.pct}, {key: 'stu', label: 'Stupéfiants', n: 1, fmt: E.pct},
          {key: 'dms', label: 'DMS', n: 1, fmt: E.pct}, {key: 'glob', label: 'Global', n: 1, fmt: E.pct}, {key: 'eur', label: 'Intra-GHS / an', n: 1, fmt: E.eur}
        ], g.cles.map((r) => ({...r, _muted: !r.recue, nom: r.nom + (r.recue ? '' : ' (non transmise)')})), {uf: 'Total', med: '100 %', sol: '100 %', stu: '100 %', dms: '100 %', glob: '100 %', eur: E.eur(sum(g.cles, (r) => r.eur))}));
    });

    const views = el('div', {class: 'eff-views'}, multi ? globalSec : null, blocs);
    root.append(views);
    function show(view) {
      const ids = [...(multi ? ['ensemble'] : []), ...blocs.map((b) => b.id)];
      const ok = ids.includes(view) ? view : ids[0];
      globalSec.hidden = ok !== 'ensemble';
      blocs.forEach((b) => { b.hidden = b.id !== ok; });
      root.querySelectorAll('.eff-tab').forEach((b) => { const on = b.dataset.view === ok; b.classList.toggle('on', on); b.setAttribute('aria-selected', on ? 'true' : 'false'); });
      if (multi && history.replaceState) history.replaceState(null, '', location.pathname + location.search + '#' + ok);
      requestAnimationFrame(() => window.dispatchEvent(new Event('resize')));
    }
    root.querySelectorAll('.eff-tab').forEach((b) => b.addEventListener('click', () => show(b.dataset.view)));
    show(decodeURIComponent((location.hash || '').slice(1)));

    root.append(el('details', {class: 'eff-method'}, el('summary', {text: 'Méthode'}), el('ul', {},
      el('li', {text: 'Source : extractions « Analyse_conso » du logiciel de gestion, une par UF, valorisées au PMP de la dernière balance définitive ; les retours sont déduits.'}),
      el('li', {text: `Annualisation : montant de la période × 12 / ${p.mois}.`}),
      el('li', {text: 'Familles selon le compte de stock : solutés = poches et flacons de perfusion du compte H942110 (hors nutrition parentérale) ; autres produits des comptes H9421… = médicaments (ampoules d’électrolytes et nutrition parentérale comprises) ; H9422… = DMS.'}),
      el('li', {text: 'Financement : produit inscrit sur la liste des médicaments hors GHS ou des DMS LPPR = remboursé en sus ; sinon intra-GHS.'}),
      el('li', {text: 'Mise à jour : ' + new Date(d.genere).toLocaleDateString('fr-FR', {day: 'numeric', month: 'long', year: 'numeric'}) + '.'}))));
  };
})();
