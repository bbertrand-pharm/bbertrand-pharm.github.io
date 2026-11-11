/* Espace pharmacien : discuter des pistes d'efficience avec Claude, via l'abonnement claude.ai du pharmacien.
   La page prépare le contexte (copié dans le presse-papiers et ouvert dans claude.ai), puis applique la liste
   de pistes que Claude renvoie entre les balises <pistes>…</pistes>. Aucune clé API, aucun coût supplémentaire. */
(function () {
  'use strict';
  const E = window.EFF, el = E.el;
  let ACTIVE = false;
  E.pharmaActive = () => ACTIVE;
  E.pharmaOpen = () => { ACTIVE = true; };
  E.pharmaLock = () => { ACTIVE = false; };

  const fmtN = (v) => new Intl.NumberFormat('fr-FR', {maximumFractionDigits: 0}).format(v);
  const contexte = ({g, c, d}) => {
    const p = d.periode;
    const lignes = c.produits(g.code, '*', true).slice(0, 45).map((r) => `${r.rang}. ${r.libelle} | ${r.famille} | ${r.fin} | ${fmtN(r.q)} /an | ${fmtN(r.v)} €/an | cumul ${r.cumul} %`);
    const fam = Object.entries(g.fam).filter(([, v]) => v.total).map(([f, v]) => `${f} : ${fmtN(v.total)} € (dont ${fmtN(v.sus)} € remboursés en sus)`).join(' ; ');
    let dot = 'Dotation : non transmise.';
    const dt = c.dotation(g.code);
    if (dt) {
      const cnt = {}; dt.rows.forEach((r) => { cnt[r.statut] = (cnt[r.statut] || 0) + 1; });
      const ex = dt.rows.filter((r) => r.statut === 'À augmenter' || r.statut === 'Sans consommation').slice(0, 15).map((r) => `${r.libelle} : dotation ${r.dot ?? '—'}, conso ${fmtN(r.qAn)}/an, proposée ${r.cible} (${r.statut})`);
      dot = `Dotation : ${Object.entries(cnt).map(([k, v]) => `${k} ${v}`).join(', ')}. Exemples : ${ex.join(' ; ')}`;
    }
    const pistes = JSON.stringify((d.pistes || {})[g.code] || [], null, 1);
    return `Je suis pharmacien hospitalier clinicien (France). J'analyse les consommations de médicaments, solutés, stupéfiants et dispositifs médicaux stériles d'un service de soins pour proposer des actions d'efficience sans diminuer la qualité de la prise en charge. Aide-moi à affiner les pistes ci-dessous.

Consignes : réponds en français, de façon concise et argumentée ; appuie-toi sur les données fournies sans les inventer, et sur des données probantes (essais, recommandations HAS, sociétés savantes) en les nommant ; signale ce qui reste à vérifier localement ; distingue ce qui est à la charge de l'hôpital (intra-GHS) de ce qui est remboursé en sus. Chaque fois que je te demande de modifier, ajouter, retirer ou réécrire des pistes, termine ta réponse par la liste COMPLÈTE des pistes mises à jour, en JSON strict entre les balises <pistes> et </pistes> : un tableau d'objets {"titre": "...", "texte": "...", "gain": "..." (facultatif), "sources": [{"label": "...", "url": "https://..."}] (facultatif, seulement des URL dont tu es sûr)}.

DONNÉES
Service : ${d.service}. Bloc : ${g.code} (${g.ufs.map((u) => u.uf + ' ' + u.nom).join(', ')}).
Période : exercice ${p.exercice}, de janvier à la fin du mois n° ${p.mois}, extrapolée sur 12 mois.
Coût global annualisé : ${fmtN(g.total)} €, dont ${fmtN(g.sus)} € remboursés en sus (hors GHS / LPPR). ${fam}.
Références par coût décroissant (45 premières) :
${lignes.join('\n')}
${dot}

PISTES ACTUELLES (JSON)
${pistes}

Ma première demande : `;
  };

  const copier = async (txt) => {
    try { await navigator.clipboard.writeText(txt); return true; }
    catch (e) {
      const ta = el('textarea', {style: 'position:absolute;left:-9999px'}); ta.value = txt; document.body.append(ta); ta.select();
      let ok = false; try { ok = document.execCommand('copy'); } catch (x) {} ta.remove(); return ok;
    }
  };

  /* Lit la liste de pistes dans une réponse collée : balises <pistes>, ou à défaut un tableau JSON. */
  const extraire = (txt) => {
    const m = txt.match(/<pistes>([\s\S]*?)<\/pistes>/i);
    let src = m ? m[1] : (txt.match(/\[\s*\{[\s\S]*\}\s*\]/) || [null])[0];
    if (!src) return null;
    src = src.replace(/^```(?:json)?/m, '').replace(/```\s*$/m, '').trim();
    try { const l = JSON.parse(src); return Array.isArray(l) && l.length && l.every((x) => x && x.titre && x.texte) ? l : null; } catch (e) { return null; }
  };

  E.chatPanel = (args, onPistes) => {
    const {g} = args;
    const msg = el('p', {class: 'eff-small', role: 'status'});
    const go = el('button', {class: 'eff-btn', type: 'button', text: 'Copier le contexte et ouvrir Claude'});
    go.addEventListener('click', async () => {
      const ok = await copier(contexte(args));
      window.open('https://claude.ai/new', '_blank', 'noopener');
      msg.className = ok ? 'eff-ok' : 'eff-err';
      msg.textContent = ok ? 'Contexte copié. Dans la conversation Claude qui s’ouvre, collez-le (Ctrl+V ou ⌘V), complétez votre demande à la fin, puis envoyez.' : 'La copie automatique a échoué : utilisez « Afficher le contexte » et copiez-le à la main.';
    });
    const voir = el('details', {class: 'eff-chat-ctx'}, el('summary', {text: 'Afficher le contexte envoyé'}));
    voir.addEventListener('toggle', () => { if (voir.open && !voir.querySelector('pre')) voir.append(el('pre', {text: contexte(args)})); });
    const ta = el('textarea', {class: 'eff-input eff-chat-in', rows: '4', placeholder: 'Collez ici la réponse de Claude qui contient la nouvelle liste de pistes (entre <pistes> et </pistes>).'});
    const apply = el('button', {class: 'eff-btn ghost', type: 'button', text: 'Appliquer à l’encadré'});
    const res = el('p', {class: 'eff-small', role: 'status'});
    const pub = el('button', {class: 'eff-btn', type: 'button', text: 'Publier les pistes', hidden: ''});
    pub.addEventListener('click', () => {
      const b = document.getElementById('publish'), maj = document.getElementById('maj');
      if (!b) return; if (maj) maj.open = true; b.click();
      const log = document.getElementById('pub-log'); if (log) log.scrollIntoView({behavior: 'smooth', block: 'center'});
    });
    apply.addEventListener('click', () => {
      const liste = extraire(ta.value);
      if (!liste) { res.className = 'eff-err'; res.textContent = 'Aucune liste de pistes lisible dans le texte collé. Demandez à Claude de redonner la liste complète entre <pistes> et </pistes>.'; return; }
      args.d.pistes = args.d.pistes || {};
      args.d.pistes[g.code] = liste.map((x) => ({titre: String(x.titre), texte: String(x.texte), gain: x.gain ? String(x.gain) : undefined,
        sources: Array.isArray(x.sources) ? x.sources.filter((s) => s && s.label && /^https:\/\//.test(s.url)).map((s) => ({label: String(s.label), url: String(s.url)})) : undefined}));
      args.redraw(); ta.value = '';
      res.className = 'eff-ok'; res.textContent = `${liste.length} pistes appliquées à l’encadré. Relisez-les, puis « Publier les pistes » pour les enregistrer.`;
      if (onPistes) onPistes();
      pub.hidden = !document.getElementById('publish');
    });
    ta.addEventListener('input', () => { res.textContent = ''; });
    return el('section', {class: 'eff-chat', 'aria-label': 'Discuter des pistes avec Claude'},
      el('div', {class: 'eff-chat-h'}, el('b', {text: 'Discuter des pistes avec Claude'}), el('span', {class: 'eff-pill', text: 'Espace pharmacien'})),
      el('ol', {class: 'eff-chat-steps'},
        el('li', {}, 'Cliquez sur le bouton : le contexte de ce bloc (coûts, références, dotation, pistes actuelles) est copié et une conversation Claude s’ouvre avec votre abonnement.'),
        el('li', {}, 'Collez le contexte, ajoutez votre demande (« le génotypage RHD est déjà systématique, retire cette piste », « détaille un protocole albumine »…) et dialoguez librement.'),
        el('li', {}, 'Quand la liste vous convient, copiez la réponse de Claude et collez-la ci-dessous pour mettre l’encadré à jour.')),
      el('div', {class: 'eff-row'}, go), msg, voir, ta, el('div', {class: 'eff-row'}, apply, pub), res);
  };
})();
