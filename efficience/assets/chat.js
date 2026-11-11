/* Espace pharmacien : dialogue avec Claude sous les pistes d'efficience.
   La clé API reste sur l'appareil du pharmacien, chiffrée avec son code ; elle n'est jamais publiée. */
(function () {
  'use strict';
  const E = window.EFF, el = E.el;
  const STORE = 'eff-pharma-v1';
  const MODEL = 'claude-sonnet-5-5';
  let SESSION = null; // {apiKey}
  const histories = {};

  E.pharmaConfigured = () => { try { return !!localStorage.getItem(STORE); } catch (e) { return false; } };
  E.pharmaUnlock = async (code) => {
    const env = JSON.parse(localStorage.getItem(STORE));
    SESSION = await E.decrypt(env, code);
    return SESSION;
  };
  E.pharmaSetup = async (apiKey, code) => {
    localStorage.setItem(STORE, JSON.stringify(await E.encrypt({apiKey: apiKey.trim()}, code)));
    SESSION = {apiKey: apiKey.trim()};
  };
  E.pharmaForget = () => { try { localStorage.removeItem(STORE); } catch (e) {} SESSION = null; };
  E.pharmaActive = () => !!SESSION;
  E.pharmaLock = () => { SESSION = null; };

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
    const pistes = ((d.pistes || {})[g.code] || []).map((x, i) => `${i + 1}. ${x.titre} — ${x.texte}${x.gain ? ' (gain : ' + x.gain + ')' : ''}`).join('\n');
    return `Service : ${d.service}. Bloc : ${g.code} (${g.ufs.map((u) => u.uf + ' ' + u.nom).join(', ')}).
Période : exercice ${p.exercice}, janvier à fin du mois ${p.mois}, extrapolée sur 12 mois.
Coût global annualisé : ${fmtN(g.total)} €, dont ${fmtN(g.sus)} € remboursés en sus (hors GHS / LPPR). ${fam}.
Références par coût décroissant (45 premières) :
${lignes.join('\n')}
${dot}
Pistes d'efficience actuellement affichées :
${pistes || '(aucune)'}`;
  };
  const SYSTEME = `Tu es Claude, assistant d'un pharmacien hospitalier clinicien (France) qui analyse les consommations de médicaments, solutés, stupéfiants et dispositifs médicaux stériles d'un service de soins, pour proposer des actions d'efficience sans diminuer la qualité de la prise en charge.
Réponds en français, de façon concise et argumentée. Appuie-toi sur les données fournies (ne les invente pas) et sur des données probantes connues (essais, recommandations HAS, sociétés savantes) en les nommant ; signale clairement les incertitudes, et ce qui reste à vérifier localement. Distingue ce qui est à la charge de l'hôpital (intra-GHS) de ce qui est remboursé en sus.
Quand le pharmacien demande de modifier, ajouter, retirer ou réécrire des pistes, termine ta réponse par la liste COMPLÈTE des pistes mises à jour, au format JSON strict entre les balises <pistes> et </pistes> : un tableau d'objets {"titre": "...", "texte": "...", "gain": "..." (facultatif), "sources": [{"label": "...", "url": "https://..."}] (facultatif, uniquement des URL dont tu es sûr)}. Sinon, n'ajoute pas ce bloc.`;

  const appel = async (messages, ctx) => {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {'content-type': 'application/json', 'x-api-key': SESSION.apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true'},
      body: JSON.stringify({model: MODEL, max_tokens: 2000, system: SYSTEME + '\n\nDonnées du service :\n' + ctx, messages})
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      const m = (j.error && j.error.message) || r.status;
      if (r.status === 401) throw new Error('clé API refusée (invalide ou révoquée).');
      if (r.status === 429) throw new Error('limite d’usage atteinte, réessayez dans un instant.');
      if (r.status === 400 && /credit|balance/i.test(m)) throw new Error('crédit API épuisé : rechargez votre compte sur console.anthropic.com.');
      throw new Error(m);
    }
    return (j.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('\n');
  };

  /* Panneau de dialogue d'un bloc. */
  E.chatPanel = (args, onPistes) => {
    const {g} = args;
    const key = args.d.id + '|' + g.code;
    const hist = histories[key] = histories[key] || [];
    const log = el('div', {class: 'eff-chat-log', role: 'log', 'aria-live': 'polite'});
    const ta = el('textarea', {class: 'eff-input eff-chat-in', rows: '2', placeholder: 'Ex. : « Détaille la piste sur l’albumine avec un protocole type », « Ajoute une piste sur les antibiotiques », « Le génotypage RHD est déjà systématique, retire cette piste ».'});
    const send = el('button', {class: 'eff-btn', type: 'button', text: 'Envoyer'});
    const err = el('p', {class: 'eff-err'});
    const bulle = (role, text) => {
      const b = el('div', {class: 'eff-chat-msg eff-chat-' + role});
      const m = role === 'assistant' ? text.match(/<pistes>([\s\S]*?)<\/pistes>/) : null;
      const visible = m ? text.replace(m[0], '').trim() : text;
      visible.split(/\n{2,}/).forEach((para) => b.append(el('p', {text: para})));
      if (m) {
        let liste = null;
        try { liste = JSON.parse(m[1]); } catch (e) {}
        if (Array.isArray(liste) && liste.every((x) => x && x.titre && x.texte)) {
          const btn = el('button', {class: 'eff-btn ghost', type: 'button', text: `Appliquer ces ${liste.length} pistes à l’encadré`});
          btn.addEventListener('click', () => {
            args.d.pistes = args.d.pistes || {};
            args.d.pistes[g.code] = liste.map((x) => ({titre: String(x.titre), texte: String(x.texte), gain: x.gain ? String(x.gain) : undefined,
              sources: Array.isArray(x.sources) ? x.sources.filter((s) => s && s.label && /^https:\/\//.test(s.url)).map((s) => ({label: String(s.label), url: String(s.url)})) : undefined}));
            args.redraw(); btn.disabled = true; btn.textContent = 'Pistes appliquées';
            if (onPistes) onPistes();
          });
          b.append(el('div', {class: 'eff-chat-apply'}, el('span', {class: 'eff-small', text: `Proposition de ${liste.length} pistes mises à jour.`}), btn));
        }
      }
      log.append(b); log.scrollTop = log.scrollHeight;
    };
    hist.forEach((h) => bulle(h.role, h.content));
    const go = async () => {
      const q = ta.value.trim(); if (!q) { err.textContent = 'Écrivez une question d’abord.'; return; }
      err.textContent = ''; ta.value = '';
      hist.push({role: 'user', content: q}); bulle('user', q);
      const wait = el('div', {class: 'eff-chat-msg eff-chat-assistant eff-chat-wait', text: 'Claude réfléchit…'}); log.append(wait);
      send.disabled = true;
      try {
        const a = await appel(hist.slice(-12), contexte(args));
        wait.remove(); hist.push({role: 'assistant', content: a}); bulle('assistant', a);
      } catch (e) { wait.remove(); hist.pop(); err.textContent = 'Échec : ' + e.message; ta.value = q; }
      finally { send.disabled = false; ta.focus(); }
    };
    send.addEventListener('click', go);
    ta.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); go(); } });
    ta.addEventListener('input', () => { err.textContent = ''; });
    return el('section', {class: 'eff-chat', 'aria-label': 'Dialogue avec Claude'},
      el('div', {class: 'eff-chat-h'}, el('b', {text: 'Discuter des pistes avec Claude'}), el('span', {class: 'eff-pill', text: 'Espace pharmacien'})),
      el('p', {class: 'eff-small', text: 'Claude connaît les consommations, la dotation et les pistes de ce bloc. Demandez-lui de nuancer, détailler, ajouter ou retirer des pistes ; il pourra proposer une nouvelle version de l’encadré à appliquer. Ctrl + Entrée pour envoyer.'}),
      log, ta, err, el('div', {class: 'eff-row'}, send));
  };
})();
