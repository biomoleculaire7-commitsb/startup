import React, { useState } from 'react'
import ReactDOM from 'react-dom/client'

// ── CSS Injection ──────────────────────────────────────────────────────────
;(function(){
  const s = document.createElement('style')
  s.textContent = "/* src/index.css — StartupAI Global Styles */\n*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }\n\n:root {\n  --navy: #0A1172;\n  --navy-m: #1a237e;\n  --gold: #C9A84C;\n  --gold-l: #f0d080;\n  --white: #ffffff;\n  --bg: #080818;\n  --bg-m: #0f0f2a;\n  --text: #e0e0f0;\n  --text-m: #888;\n  --border: rgba(201,168,76,.2);\n  --surface: rgba(255,255,255,.03);\n}\n\nhtml, body { height: 100%; }\nbody {\n  background: linear-gradient(160deg, #080818, #0f0f2a 50%, #0a0a1a);\n  background-attachment: fixed;\n  color: var(--text);\n  font-family: 'Segoe UI', 'Cairo', Arial, sans-serif;\n  direction: rtl;\n  min-height: 100vh;\n}\n#root { min-height: 100vh; display: flex; flex-direction: column; }\n\n/* ── Animations ─── */\n@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.35} }\n@keyframes fadeUp { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }\n@keyframes spin { to{transform:rotate(360deg)} }\n.fade-up { animation: fadeUp .45s ease forwards; }\n.dots span { animation: pulse 1.2s infinite; }\n.dots span:nth-child(2) { animation-delay:.2s; }\n.dots span:nth-child(3) { animation-delay:.4s; }\n.spin { display:inline-block; animation: spin 1s linear infinite; }\n\n/* ── Scrollbar ─── */\n::-webkit-scrollbar { width: 5px; }\n::-webkit-scrollbar-track { background: #0a0a1a; }\n::-webkit-scrollbar-thumb { background: #3a3a5c; border-radius: 3px; }\n\n/* ── App layout ─── */\n.app { display: flex; flex-direction: column; min-height: 100vh; }\n\n/* ── Header ─── */\n.header {\n  background: rgba(255,255,255,.03);\n  border-bottom: 1px solid rgba(201,168,76,.15);\n  padding: .75rem 1.5rem;\n  display: flex;\n  align-items: center;\n  justify-content: space-between;\n  position: sticky; top: 0; z-index: 100;\n  backdrop-filter: blur(10px);\n}\n.header-brand { display:flex; align-items:center; gap:.65rem; }\n.brand-logo {\n  width: 2.1rem; height: 2.1rem;\n  background: linear-gradient(135deg, var(--gold), var(--gold-l));\n  border-radius: .42rem;\n  display: flex; align-items: center; justify-content: center;\n  font-size: .95rem; font-weight: 800; color: var(--navy);\n}\n.brand-name { color: var(--gold); font-weight: 800; font-size: .9rem; line-height: 1; }\n.brand-sub { color: #3a3a55; font-size: .62rem; }\n.header-actions { display:flex; gap:.4rem; align-items:center; }\n.badge-ai {\n  background: rgba(201,168,76,.1); color: var(--gold);\n  padding: .2rem .6rem; border-radius: 2rem;\n  font-size: .67rem; border: 1px solid rgba(201,168,76,.28); font-weight: 700;\n}\n.btn-reset {\n  background: rgba(255,80,80,.1); color: #ff6060;\n  border: 1px solid rgba(255,80,80,.28); border-radius: .32rem;\n  padding: .22rem .6rem; cursor: pointer; font-size: .67rem; font-family: inherit;\n  transition: all .2s;\n}\n.btn-reset:hover { background: rgba(255,80,80,.2); }\n\n/* ── Main ─── */\n.main { flex: 1; max-width: 900px; margin: 0 auto; padding: 1.5rem 1.1rem; width: 100%; }\n\n/* ── Hero ─── */\n.hero { text-align: center; margin-bottom: 1.6rem; }\n.hero-badge {\n  display: inline-block;\n  background: rgba(201,168,76,.1); border: 1px solid rgba(201,168,76,.28);\n  border-radius: 2rem; padding: .28rem .85rem; font-size: .72rem;\n  color: var(--gold); margin-bottom: .65rem; font-weight: 700;\n}\n.hero-title {\n  font-size: clamp(1.45rem,4vw,2.3rem); font-weight: 900;\n  background: linear-gradient(135deg, #fff, var(--gold));\n  -webkit-background-clip: text; -webkit-text-fill-color: transparent;\n  margin-bottom: .45rem; line-height: 1.2;\n}\n.hero-sub { color: #666; font-size: .83rem; margin-bottom: .7rem; }\n.hero-sub strong { color: var(--gold); }\n.axes-badges { display: flex; flex-wrap: wrap; gap: .3rem; justify-content: center; }\n.axis-badge {\n  background: rgba(201,168,76,.07); border: 1px solid rgba(201,168,76,.15);\n  border-radius: .32rem; padding: .15rem .5rem; font-size: .63rem; color: #806830;\n}\n\n/* ── Form ─── */\n.form-card {\n  background: var(--surface); border: 1px solid var(--border);\n  border-radius: 1rem; padding: 1.4rem; margin-bottom: 1.2rem;\n}\n.form-section { margin-bottom: 1rem; }\n.section-label { display: block; color: var(--gold); font-size: .8rem; font-weight: 700; margin-bottom: .4rem; }\n.label-hint { color: #555; font-weight: 400; font-size: .68rem; }\n.idea-input {\n  width: 100%; background: rgba(255,255,255,.05);\n  border: 1px solid rgba(201,168,76,.18); border-radius: .5rem;\n  padding: .65rem .85rem; color: var(--text); font-size: .86rem;\n  font-family: inherit; resize: vertical; outline: none; line-height: 1.7;\n  transition: border-color .3s;\n}\n.idea-input:focus { border-color: rgba(201,168,76,.55); }\n.char-count { display: block; font-size: .68rem; margin-top: .25rem; }\n\n.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: .55rem; margin-bottom: .8rem; }\n.grid-3 { display: grid; grid-template-columns: 2fr 1.3fr 1.3fr; gap: .35rem; }\n.flex-1 { flex: 1; }\n\n.field label { display: block; color: var(--gold); font-size: .76rem; font-weight: 700; margin-bottom: .25rem; }\n.field.small label { font-size: .68rem; }\n.field input {\n  width: 100%; background: rgba(255,255,255,.05);\n  border: 1px solid rgba(201,168,76,.18); border-radius: .4rem;\n  padding: .45rem .75rem; color: var(--text); font-size: .83rem;\n  font-family: inherit; outline: none; transition: border-color .3s;\n}\n.field.small input { padding: .32rem .55rem; font-size: .77rem; }\n.field input:focus { border-color: rgba(201,168,76,.55); }\n\n.members-section { border-top: 1px solid rgba(201,168,76,.12); padding-top: .85rem; }\n.member-row { display: flex; align-items: flex-start; gap: .5rem; margin-bottom: .4rem; }\n.member-avatar {\n  width: 2rem; height: 2rem; border-radius: 50%;\n  display: flex; align-items: center; justify-content: center;\n  font-size: 1rem; flex-shrink: 0; margin-top: 1.5rem;\n}\n\n.form-footer { display: flex; justify-content: flex-end; margin-top: 1rem; }\n.btn-generate {\n  background: #1a1a2e; color: #2a2a3e; border: none;\n  border-radius: .5rem; padding: .72rem 2rem;\n  font-family: inherit; font-weight: 800; font-size: .92rem;\n  cursor: not-allowed; transition: all .3s;\n}\n.btn-generate.ready {\n  background: linear-gradient(135deg, var(--gold), var(--gold-l));\n  color: var(--navy); cursor: pointer;\n  box-shadow: 0 4px 18px rgba(201,168,76,.3);\n}\n.btn-generate.ready:hover { transform: translateY(-1px); box-shadow: 0 6px 24px rgba(201,168,76,.4); }\n\n/* ── Error ─── */\n.error-box {\n  background: rgba(255,80,80,.09); border: 1px solid rgba(255,80,80,.28);\n  border-radius: .55rem; padding: .75rem 1rem; color: #ff8080;\n  margin-bottom: 1rem; font-size: .8rem;\n}\n\n/* ── Progress ─── */\n.progress-section { margin-bottom: 1.2rem; }\n.step-badge {\n  display: flex; align-items: center; gap: .5rem;\n  padding: .58rem .92rem; border-radius: .55rem; margin-bottom: .35rem;\n  border: 1px solid #1e1e30; background: rgba(255,255,255,.02);\n  transition: all .4s;\n}\n.step-badge.done { background: rgba(80,200,120,.07); border-color: rgba(80,200,120,.3); }\n.step-badge.active { background: rgba(201,168,76,.1); border-color: rgba(201,168,76,.35); }\n.step-icon { font-size: .9rem; }\n.step-label { color: #444; font-size: .8rem; }\n.step-badge.done .step-label { color: #6de06d; font-weight: 700; }\n.step-badge.active .step-label { color: var(--gold); font-weight: 700; }\n.step-sub { color: #333348; font-size: .68rem; }\n\n/* ── Download banners ─── */\n.dl-banner {\n  border-radius: .55rem; padding: .65rem 1rem;\n  margin-bottom: 1rem; font-size: .8rem;\n  display: flex; align-items: center; gap: .5rem;\n}\n.dl-banner.pptx { background: rgba(26,35,126,.15); border: 1px solid rgba(26,35,126,.4); color: #8888ff; }\n.dl-banner.docx { background: rgba(21,101,192,.12); border: 1px solid rgba(21,101,192,.35); color: #60a0ff; }\n\n/* ── Results ─── */\n.results { display: flex; flex-direction: column; gap: 0; }\n.project-banner {\n  background: linear-gradient(135deg, rgba(201,168,76,.1), rgba(10,17,114,.12));\n  border: 1px solid rgba(201,168,76,.28);\n  border-radius: .9rem; padding: 1.1rem 1.3rem; margin-bottom: 1.1rem;\n}\n.project-info { margin-bottom: .9rem; }\n.project-title { color: var(--gold); font-weight: 900; font-size: 1.1rem; }\n.project-tagline { color: #666; font-size: .76rem; margin-top: .18rem; }\n.project-uni { color: #2a2a45; font-size: .66rem; margin-top: .12rem; }\n\n.download-buttons { display: flex; gap: .6rem; flex-wrap: wrap; margin-bottom: .8rem; }\n.btn-dl {\n  display: flex; align-items: center; gap: .5rem;\n  border-radius: .48rem; padding: .65rem 1rem;\n  cursor: pointer; font-family: inherit; font-weight: 700;\n  font-size: .8rem; border: 1px solid rgba(201,168,76,.35);\n  transition: all .25s; text-align: right;\n}\n.btn-dl:disabled { opacity: .55; cursor: not-allowed; }\n.btn-dl .dl-icon { font-size: 1.2rem; }\n.btn-dl .dl-sub { font-size: .65rem; opacity: .75; margin-top: .08rem; }\n.btn-dl.pptx { background: linear-gradient(135deg,#1a237e,#283593); color: #fff; box-shadow: 0 3px 12px rgba(26,35,126,.3); }\n.btn-dl.docx { background: linear-gradient(135deg,#1565C0,#0D47A1); color: #fff; box-shadow: 0 3px 12px rgba(21,101,192,.28); }\n.btn-dl.json { background: rgba(201,168,76,.12); color: var(--gold); }\n.btn-dl:not(:disabled):hover { transform: translateY(-2px); filter: brightness(1.1); }\n\n.info-note {\n  padding: .5rem .75rem; background: rgba(201,168,76,.06);\n  border-radius: .4rem; border: 1px solid rgba(201,168,76,.12);\n  color: #806830; font-size: .68rem;\n}\n\n/* ── Card ─── */\n.card {\n  background: var(--surface); border: 1px solid rgba(201,168,76,.18);\n  border-radius: .9rem; overflow: hidden; margin-bottom: 1rem;\n}\n.card-header {\n  display: flex; align-items: center; justify-content: space-between;\n  padding: .8rem 1.2rem; cursor: pointer;\n  background: rgba(201,168,76,.055);\n  border-bottom: 1px solid rgba(201,168,76,.12);\n}\n.card-title-group { display: flex; align-items: center; gap: .55rem; }\n.card-icon { font-size: 1.1rem; }\n.card-title { color: var(--gold); font-weight: 800; font-size: .85rem; }\n.card-sub { color: #555; font-size: .66rem; }\n.card-toggle { color: var(--gold); font-size: .7rem; }\n.card-body { padding: 1.1rem; }\n\n/* ── Axes list ─── */\n.axis-row {\n  display: flex; align-items: center; justify-content: space-between;\n  padding: .44rem .7rem; background: rgba(10,17,114,.1);\n  border: 1px solid #1a2050; border-radius: .4rem; margin-bottom: .3rem;\n}\n.ax-num { color: var(--gold); font-weight: 700; font-size: .76rem; }\n.ax-fr { color: #777; font-size: .73rem; }\n.ax-done { color: #4a4; font-size: .7rem; font-weight: 700; }\n\n/* ── Team ─── */\n.team-grid { display: flex; gap: .6rem; flex-wrap: wrap; }\n.team-card {\n  background: rgba(255,255,255,.04); border-width: 1px; border-style: solid;\n  border-radius: .6rem; padding: .6rem .9rem; min-width: 150px; flex: 1;\n}\n.team-avatar {\n  width: 2rem; height: 2rem; border-radius: 50%;\n  display: flex; align-items: center; justify-content: center;\n  font-size: 1rem; margin-bottom: .35rem;\n}\n.team-name { color: var(--text); font-weight: 700; font-size: .82rem; }\n.team-role { color: #888; font-size: .72rem; margin-top: .1rem; }\n.team-bg { color: #555; font-size: .68rem; }\n\n/* ── Script ─── */\n.script-box {\n  background: rgba(0,0,0,.28); border-radius: .55rem;\n  padding: 1rem; max-height: 350px; overflow-y: auto;\n  line-height: 1.95; font-size: .82rem; color: #bbbbd0;\n}\n.script-box pre { font-family: inherit; white-space: pre-wrap; word-break: break-word; }\n.btn-copy {\n  margin-top: .65rem; background: rgba(201,168,76,.12); color: var(--gold);\n  border: 1px solid rgba(201,168,76,.28); border-radius: .38rem;\n  padding: .38rem .85rem; cursor: pointer; font-family: inherit;\n  font-size: .74rem; font-weight: 700; transition: all .2s;\n}\n.btn-copy:hover { background: rgba(201,168,76,.2); }\n\n/* ── Done banner ─── */\n.done-banner {\n  background: linear-gradient(135deg, rgba(201,168,76,.09), rgba(201,168,76,.03));\n  border: 1px solid rgba(201,168,76,.3); border-radius: .9rem;\n  padding: 1.3rem; text-align: center; margin-top: .5rem;\n}\n.done-emoji { font-size: 1.7rem; margin-bottom: .3rem; }\n.done-banner h3 { color: var(--gold); margin-bottom: .3rem; font-weight: 800; font-size: 1.05rem; }\n.done-banner p { color: #666; font-size: .78rem; margin-bottom: .25rem; }\n.done-sub { color: #333348 !important; font-size: .68rem !important; margin-bottom: .85rem !important; }\n.btn-new {\n  background: linear-gradient(135deg, var(--gold), var(--gold-l));\n  color: var(--navy); border: none; border-radius: .45rem;\n  padding: .55rem 1.4rem; cursor: pointer; font-family: inherit;\n  font-weight: 800; font-size: .83rem; transition: all .25s;\n}\n.btn-new:hover { transform: translateY(-1px); box-shadow: 0 4px 16px rgba(201,168,76,.35); }\n\n/* ── Footer ─── */\n.footer {\n  text-align: center; padding: 1rem;\n  color: #1a1a2e; font-size: .65rem;\n  border-top: 1px solid #0d0d1e;\n}\n\n/* ── Responsive ─── */\n@media (max-width: 640px) {\n  .grid-2 { grid-template-columns: 1fr; }\n  .grid-3 { grid-template-columns: 1fr; }\n  .download-buttons { flex-direction: column; }\n  .btn-dl { width: 100%; }\n  .member-avatar { display: none; }\n}\n"
  document.head.appendChild(s)
})()

// ── Prompts ────────────────────────────────────────────────────────────────
// src/prompts.js — System prompts for Claude

const SYS_JSON = `IMPORTANT: Tu dois répondre UNIQUEMENT avec du JSON valide. Commence directement par { sans aucun texte avant. Termine par } sans aucun texte après. Pas de markdown, pas de backticks.

Tu es expert en projets Diplôme/Startup (Arrêté Ministériel 1275 - CNCSIIU/MHESR Algérie).
Réponds UNIQUEMENT avec un JSON valide et complet, sans backticks ni texte supplémentaire.

Structure JSON exacte:
{
  "projectTitle": "Nom court et percutant",
  "tagline": "Slogan accrocheur",
  "guide": {
    "ax1_domain": "domaine activité précis",
    "ax1_idea": "description 3-4 phrases détaillées",
    "ax1_how": "comment réalisé concrètement",
    "ax1_who_where": "qui et où",
    "ax1_values": "valeurs: modernité, performance, flexibilité, accessibilité, réduction coûts",
    "ax1_obj_short": "objectif 1 an avec chiffres",
    "ax1_obj_mid": "objectif 3 ans avec part de marché",
    "ax1_obj_long": "objectif 5 ans leader national",
    "ax1_market_share": "X% du marché",
    "ax1_timeline": [["1","Étude & Validation"],["2","Développement MVP"],["3","Tests utilisateurs"],["4","Lancement officiel"],["5","Croissance & Scale"]],
    "ax2_innovation_type": "type innovation",
    "ax2_innovation_justification": "pourquoi pertinent en Algérie",
    "ax2_domains": "domaines innovation",
    "ax2_specifics": "innovations concrètes distinctives",
    "ax3_potential_market": "marché potentiel: qui, combien, où",
    "ax3_target_market": "segment cible précis",
    "ax3_segment_why": "justification",
    "ax3_key_clients": "clients-clés potentiels",
    "ax3_competitors": [{"name":"Concurrent 1","forces":"","weaknesses":""},{"name":"Concurrent 2","forces":"","weaknesses":""},{"name":"Indirect","forces":"","weaknesses":""}],
    "ax3_comp_shares": "parts de marché estimées",
    "ax3_mix_0": "description produit",
    "ax3_mix_1": "stratégie prix",
    "ax3_mix_2": "canaux distribution",
    "ax3_mix_3": "actions promotionnelles",
    "ax3_visibility": "visibilité digitale et physique",
    "ax3_distribution": "distribution détaillée",
    "ax4_process": [{"step":"Étape 1","desc":"","result":""},{"step":"Étape 2","desc":"","result":""},{"step":"Étape 3","desc":"","result":""},{"step":"Étape 4","desc":"","result":""},{"step":"Étape 5","desc":"","result":""}],
    "ax4_purchase_policy": "politique approvisionnement",
    "ax4_suppliers": "fournisseurs identifiés",
    "ax4_payment_policy": "modalités paiement",
    "ax4_workforce": [{"post":"Directeur","nb":"1","qual":"Gestion","loc":"Siège"},{"post":"Développeur","nb":"2","qual":"Informatique","loc":"Sur site"},{"post":"Commercial","nb":"1","qual":"Marketing","loc":"Terrain"}],
    "ax6_prototype_type": "numérique/physique/vidéo",
    "ax6_description": "description prototype",
    "ax6_steps": [{"step":"Spécifications","desc":""},{"step":"Maquette","desc":""},{"step":"Développement","desc":""},{"step":"Tests","desc":""},{"step":"Présentation jury","desc":""}],
    "ax6_link": "URL démo",
    "ax6_jury_note": "comment présenter au jury"
  },
  "slides": {
    "slide2_description": "phrase percutante vision",
    "slide3": {"problem":"","solution":"","status":"","product":"","businessModel":"","customers":""},
    "slide5": {"mainProblem":"","subProblems":["","",""],"affectedCount":"","stats":""},
    "slide6": {"solutionDesc":"","benefits":["","",""],"steps":["","","",""],"mediaNote":""},
    "slide7": {"features":["","","","",""],"digitalProduct":"","whatsapp":"+213 ..."},
    "slide8": {"productionSteps":["","","","",""],"productionNote":"","partners":["incubateur","fournisseurs","financement","collectivités"]},
    "slide9": {"segment":"","marketSize":"TAM: X | SAM: X | SOM: X (DZD)","growth":""},
    "slide10": {"competitors":"","advantages":"","innovation":""},
    "slide11": {"visibility":"","distribution":"","pricing":"","mixMarketing":{"product":"","price":"","place":"","promotion":""}},
    "slide12": {"challenges":"","awards":"","media":"","social":""},
    "slide13": {"phases":[{"period":"M1-M2","label":"Validation","desc":"Étude"},{"period":"M3-M4","label":"MVP","desc":"Prototype"},{"period":"M5-M6","label":"Beta","desc":"Tests"},{"period":"M7-M9","label":"Lancement","desc":"Go-to-Market"},{"period":"M10-M12","label":"Croissance","desc":"Scale"}]},
    "slide14": {"revenues":[500000,1500000,4000000],"initialInvestment":"X DZD","ca1":"X DZD","breakEven":"Mois X","needs":["Financement","RH","Infrastructure","Légal"]},
    "slide15": {"industrialPartner":"partenaire industriel"}
  },
  "contact": {"phone":"+213 ...","email":"contact@projet.dz","website":"www.projet.dz"}
}
Utilise des chiffres crédibles pour le marché algérien.`;

const SYS_SCRIPT = `أنت خبير في كتابة سكريبتات الفيديو التسويقية للشركات الناشئة الجزائرية.
اكتب سكريبت فيديو احترافي مدته 2-3 دقائق بالعربية الحيوية الشبابية.

قسّمه هكذا:
[HOOK — 0:00-0:15] [بصري: ...] [صوت: ...]
[PROBLÈME — 0:15-0:40] [بصري: ...] [صوت: ...]
[SOLUTION — 0:40-1:10] [بصري: ...] [صوت: ...]
[FONCTIONNALITÉS — 1:10-1:45] [بصري: ...] [صوت: ...]
[PREUVE SOCIALE — 1:45-2:05] [بصري: ...] [صوت: ...]
[CTA — 2:05-2:20] [بصري: ...] [صوت: ...]

اجعل النص ملهماً ومحفزاً يخاطب الشباب والطلبة الجزائريين.`;


// ── API ────────────────────────────────────────────────────────────────────
// src/api.js — API calls

const BACKEND = import.meta.env.VITE_BACKEND_URL || '';

// ── Claude API (called directly from browser) ──────────────────────────────
async function callClaude(system, userMsg, stream = false, onChunk = null) {
  const body = {
    max_tokens: 4000,
    system,
    messages: [{ role: 'user', content: userMsg }],
  };

  if (!stream) {
    // Non-streaming — use backend proxy
    const res = await fetch('/claude', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const errText = await res.text().catch(() => 'Unknown error');
      throw new Error('API Error ' + res.status + ': ' + errText.slice(0, 200));
    }
    const d = await res.json();
    if (d.error) throw new Error(typeof d.error === 'string' ? d.error : JSON.stringify(d.error));
    if (!d.content || !d.content.length) throw new Error('Empty response from AI');
    return (d.content || []).map(b => b.text || '').join('');
  }

  // Streaming — use backend SSE proxy
  const res = await fetch('/claude/stream', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const errText = await res.text().catch(() => '');
    throw new Error('Stream Error ' + res.status + ': ' + errText.slice(0, 200));
  }
  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let full = '';
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    for (const line of dec.decode(value).split('\n')) {
      if (line.startsWith('data: ')) {
        try {
          const d = JSON.parse(line.slice(6));
          if (d.type === 'content_block_delta' && d.delta?.text) {
            full += d.delta.text;
            if (onChunk) onChunk(full);
          }
        } catch {}
      }
    }
  }
  return full;
}

// ── Backend file generation ────────────────────────────────────────────────
async function callBackend(endpoint, data) {
  const res = await fetch(`${BACKEND}${endpoint}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }));
    throw new Error(err.error || `Erreur serveur ${res.status}`);
  }
  return res.blob();
}

const generatePPTX = data => callBackend('/generate-pptx', data);
const generateDOCX = data => callBackend('/generate-docx', data);

// ── Download helper ────────────────────────────────────────────────────────
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}


// ── App ────────────────────────────────────────────────────────────────────

const AXES = [
  { id:1, fr:'Présentation du projet',           ar:'عرض المشروع',               icon:'📁' },
  { id:2, fr:'Aspects innovants',                ar:'الجوانب الابتكارية',         icon:'💡' },
  { id:3, fr:'Analyse stratégique du marché',    ar:'التحليل الاستراتيجي للسوق',  icon:'📊' },
  { id:4, fr:'Plan de production et org.',       ar:'خطة الإنتاج والتنظيم',       icon:'⚙️' },
  { id:5, fr:'Plan financier',                   ar:'الخطة المالية',              icon:'💰' },
  { id:6, fr:'Prototype expérimental',           ar:'النموذج التجريبي',           icon:'🔬' },
]

const MEMBER_COLORS = ['#4472C4','#7030A0','#C00000','#375623']

function Spinner() {
  return <span className="spin">⏳</span>
}

function Dots() {
  return <span className="dots"><span>.</span><span>.</span><span>.</span></span>
}

function Field({ label, value, onChange, placeholder, disabled = false, small = false }) {
  return (
    <div className={`field ${small ? 'small' : ''}`}>
      <label>{label}</label>
      <input
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        disabled={disabled}
      />
    </div>
  )
}

function StepBadge({ done, active, label, sub }) {
  return (
    <div className={`step-badge ${done ? 'done' : active ? 'active' : ''}`}>
      <span className="step-icon">{done ? '✅' : active ? '⏳' : '⬜'}</span>
      <div>
        <div className="step-label">{label} {active && <Dots />}</div>
        {sub && <div className="step-sub">{sub}</div>}
      </div>
    </div>
  )
}

function Card({ title, subtitle, icon, children, defaultOpen = true }) {
  const [open, setOpen] = useState(defaultOpen)
  return (
    <div className="card">
      <div className="card-header" onClick={() => setOpen(o => !o)}>
        <div className="card-title-group">
          <span className="card-icon">{icon}</span>
          <div>
            <div className="card-title">{title}</div>
            {subtitle && <div className="card-sub">{subtitle}</div>}
          </div>
        </div>
        <span className="card-toggle">{open ? '▲' : '▼'}</span>
      </div>
      {open && <div className="card-body">{children}</div>}
    </div>
  )
}

export default function App() {
  const [idea, setIdea]             = useState('')
  const [university, setUniversity] = useState('')
  const [incubator, setIncubator]   = useState('')
  const [faculty, setFaculty]       = useState('')
  const [department, setDepartment] = useState('')
  const [speciality, setSpeciality] = useState('')
  const [year, setYear]             = useState('2025/2026')
  const [members, setMembers]       = useState([
    { name: '', role: '', background: '' },
    { name: '', role: '', background: '' },
    { name: '', role: '', background: '' },
  ])

  const [phase, setPhase]           = useState(null)
  const [loading, setLoading]       = useState(false)
  const [error, setError]           = useState('')
  const [projectData, setProjectData] = useState(null)
  const [script, setScript]         = useState('')
  const [showForm, setShowForm]     = useState(true)

  const ready = idea.trim().length > 50 && university.trim() && incubator.trim()
  const updMember = (i, f, v) => setMembers(m => m.map((mem, idx) => idx === i ? { ...mem, [f]: v } : mem))

  const generate = async () => {
    if (!ready || loading) return
    setLoading(true); setError(''); setProjectData(null); setScript(''); setShowForm(false)
    try {
      setPhase('json')
      const memberStr = members.filter(m => m.name).map(m => `${m.name} (${m.role || 'Membre'})`).join(', ') || 'À définir'
      const rawJSON = await callClaude(
        SYS_JSON,
        `Idée: ${idea}\nUniversité: ${university}\nIncubateur: ${incubator}\nFaculté: ${faculty || '...'}\nDépartement: ${department || '...'}\nSpécialité: ${speciality || '...'}\nAnnée: ${year}\nMembres: ${memberStr}`
      )
      let data
      try {
        // Clean response: remove markdown, find JSON boundaries
        let clean = rawJSON.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim()
        const start = clean.indexOf('{')
        const end   = clean.lastIndexOf('}')
        if (start !== -1 && end !== -1) clean = clean.slice(start, end + 1)
        data = JSON.parse(clean)
      } catch {
        throw new Error('Erreur JSON — réessayez avec plus de détails sur votre idée.')
      }

      data.universityName  = university
      data.incubatorName   = incubator
      data.facultyName     = faculty
      data.departmentName  = department
      data.speciality      = speciality
      data.academicYear    = year
      if (members.some(m => m.name)) {
        data.teamMembers = members.filter(m => m.name).map(m => ({
          name: m.name, role: m.role || 'Membre',
          background: m.background || speciality || 'Informatique',
          faculty, contact: '+213 ...',
        }))
      }
      setProjectData(data)

      setPhase('script')
      const sc = await callClaude(
        SYS_SCRIPT,
        `اسم المشروع: ${data.projectTitle}\nالفكرة: ${idea}\nالوصف: ${data.slides?.slide2_description || ''}`,
        true, setScript
      )
      setScript(sc)
      setPhase('done')
    } catch (e) {
      setError(e.message || 'خطأ غير متوقع — تحقق من Render Logs')
      setPhase(null)
    }
    setLoading(false)
  }

  const dlPPTX = async () => {
    if (!projectData) return
    setPhase('dl_pptx')
    try {
      const blob = await generatePPTX(projectData)
      downloadBlob(blob, `PitchDeck_${(projectData.projectTitle || 'startup').replace(/\s+/g, '_')}.pptx`)
    } catch (e) { setError('خطأ PPTX: ' + e.message) }
    setPhase('done')
  }

  const dlDOCX = async () => {
    if (!projectData) return
    setPhase('dl_docx')
    try {
      const blob = await generateDOCX(projectData)
      downloadBlob(blob, `Guide_Projet_${(projectData.projectTitle || 'startup').replace(/\s+/g, '_')}.docx`)
    } catch (e) { setError('خطأ DOCX: ' + e.message) }
    setPhase('done')
  }

  const reset = () => {
    setIdea(''); setPhase(null); setLoading(false); setError('')
    setProjectData(null); setScript(''); setShowForm(true)
  }

  return (
    <div className="app">
      <header className="header">
        <div className="header-brand">
          <div className="brand-logo">S</div>
          <div>
            <div className="brand-name">StartupAI</div>
            <div className="brand-sub">Diplôme/Startup 1275 • CNCSIIU • MHESR</div>
          </div>
        </div>
        <div className="header-actions">
          <span className="badge-ai">✦ Claude AI</span>
          {phase && (
            <button className="btn-reset" onClick={reset}>↺ مشروع جديد</button>
          )}
        </div>
      </header>

      <main className="main">

        {/* Hero */}
        {showForm && (
          <div className="hero fade-up">
            <div className="hero-badge">🎓 وفق المرسوم الوزاري 1275 — Diplôme/Startup</div>
            <h1 className="hero-title">من فكرة إلى مشروع Startup رسمي</h1>
            <p className="hero-sub">
              دليل المشروع <strong>Word قابل للتعديل</strong> (6 محاور) +
              Pitch Deck <strong>PowerPoint</strong> (16 شريحة) +
              سكريبت الفيديو
            </p>
            <div className="axes-badges">
              {AXES.map(a => (
                <span key={a.id} className="axis-badge">{a.icon} Axe {a.id}: {a.fr}</span>
              ))}
            </div>
          </div>
        )}

        {/* Form */}
        {showForm && (
          <div className="form-card fade-up">
            <div className="form-section">
              <label className="section-label">💡 فكرة مشروعك *</label>
              <textarea
                value={idea}
                onChange={e => setIdea(e.target.value)}
                disabled={loading}
                placeholder="مثال: تطبيق يربط طلبة الجزائر بأساتذة معتمدين للدروس الخصوصية عبر نظام حجز رقمي ودفع إلكتروني... (اكتب على الأقل 2-3 جمل لتفاصيل أفضل)"
                rows={3}
                className="idea-input"
              />
              <span className="char-count" style={{ color: idea.length > 50 ? '#C9A84C' : '#C06030' }}>
                {idea.length} حرف • {Math.round(idea.trim().split(/\s+/).filter(Boolean).length)} كلمة {idea.trim().length < 50 ? `— أضف ${50 - idea.trim().length} حرف أو أكثر` : '✓ جاهز'}
              </span>
            </div>

            <div className="grid-2">
              <Field label="🏛️ اسم الجامعة *" value={university} onChange={setUniversity} placeholder="Université de Béjaïa..." disabled={loading} />
              <Field label="🏢 اسم الحاضنة *" value={incubator} onChange={setIncubator} placeholder="Incubateur UBMA..." disabled={loading} />
              <Field label="📚 الكلية" value={faculty} onChange={setFaculty} placeholder="Faculté des Sciences Exactes" disabled={loading} />
              <Field label="🔬 القسم" value={department} onChange={setDepartment} placeholder="Département Informatique" disabled={loading} />
              <Field label="🎯 التخصص" value={speciality} onChange={setSpeciality} placeholder="Master 2 Génie Logiciel" disabled={loading} />
              <Field label="📅 السنة الجامعية" value={year} onChange={setYear} placeholder="2025/2026" disabled={loading} />
            </div>

            <div className="members-section">
              <div className="section-label">👥 أعضاء فريق المشروع <span className="label-hint">(ستظهر صورهم في شريحة الفريق)</span></div>
              {members.map((m, i) => (
                <div key={i} className="member-row">
                  <div className="member-avatar" style={{ background: MEMBER_COLORS[i] }}>👤</div>
                  <div className="grid-3 flex-1">
                    <Field small label={`الطالب ${i + 1}`} value={m.name} onChange={v => updMember(i, 'name', v)} placeholder={`Étudiant 0${i + 1} — Nom Prénom`} disabled={loading} />
                    <Field small label="الدور" value={m.role} onChange={v => updMember(i, 'role', v)} placeholder="Chef de projet..." disabled={loading} />
                    <Field small label="التخصص" value={m.background} onChange={v => updMember(i, 'background', v)} placeholder="Génie logiciel..." disabled={loading} />
                  </div>
                </div>
              ))}
            </div>

            <div className="form-footer">
              <button
                className={`btn-generate ${ready && !loading ? 'ready' : ''}`}
                onClick={generate}
                disabled={!ready || loading}
              >
                {loading ? <><Dots /> جاري التوليد...</> : '⚡ توليد المشروع الكامل'}
              </button>
            </div>
          </div>
        )}

        {/* Error */}
        {error && <div className="error-box">⚠️ {error}</div>}

        {/* Progress */}
        {phase && !['done', 'dl_pptx', 'dl_docx'].includes(phase) && (
          <div className="progress-section fade-up">
            <StepBadge done={['script', 'done'].includes(phase)} active={phase === 'json'}
              label="🤖 توليد البيانات الهيكلية (6 محاور رسمية)"
              sub="JSON شامل — دليل المشروع + Pitch Deck" />
            <StepBadge done={phase === 'done'} active={phase === 'script'}
              label="🎬 كتابة سكريبت الفيديو التسويقي"
              sub="بالعربية مع توجيهات المخرج" />
          </div>
        )}

        {/* Download status */}
        {phase === 'dl_pptx' && <div className="dl-banner pptx"><Dots /> جاري إنشاء Pitch Deck PowerPoint (16 شريحة)...</div>}
        {phase === 'dl_docx' && <div className="dl-banner docx"><Dots /> جاري إنشاء دليل المشروع Word (6 محاور)...</div>}

        {/* Results */}
        {projectData && (
          <div className="results fade-up">

            {/* Project banner */}
            <div className="project-banner">
              <div className="project-info">
                <div className="project-title">{projectData.projectTitle}</div>
                <div className="project-tagline">«{projectData.tagline}»</div>
                <div className="project-uni">{university} — {incubator}</div>
              </div>
              <div className="download-buttons">
                <button
                  className="btn-dl pptx"
                  onClick={dlPPTX}
                  disabled={loading || ['dl_pptx', 'dl_docx'].includes(phase)}
                >
                  <span className="dl-icon">📊</span>
                  <div>
                    <div>تحميل Pitch Deck</div>
                    <div className="dl-sub">.pptx — 16 شريحة رسمية</div>
                  </div>
                </button>
                <button
                  className="btn-dl docx"
                  onClick={dlDOCX}
                  disabled={loading || ['dl_pptx', 'dl_docx'].includes(phase)}
                >
                  <span className="dl-icon">📄</span>
                  <div>
                    <div>تحميل دليل المشروع</div>
                    <div className="dl-sub">.docx — 6 محاور — قابل للتعديل</div>
                  </div>
                </button>
                <button
                  className="btn-dl json"
                  onClick={() => navigator.clipboard.writeText(JSON.stringify(projectData, null, 2))}
                >
                  <span className="dl-icon">📋</span>
                  <div>
                    <div>نسخ البيانات</div>
                    <div className="dl-sub">JSON كامل</div>
                  </div>
                </button>
              </div>
              <div className="info-note">
                ℹ️ الملفات تُنشأ من الخادم. Word قابل للفتح والتعديل في Microsoft Word. PowerPoint يحتوي على أيقونات الأعضاء والجامعة والحاضنة.
              </div>
            </div>

            {/* Axes summary */}
            <Card title="محاور دليل المشروع الرسمي" subtitle="وفق المرسوم الوزاري 1275 — CNCSIIU" icon="📋">
              {AXES.map(a => (
                <div key={a.id} className="axis-row">
                  <span><span className="ax-num">{a.icon} Axe {a.id}:</span> <span className="ax-fr">{a.fr}</span></span>
                  <span className="ax-done">✓ généré</span>
                </div>
              ))}
            </Card>

            {/* Team */}
            {projectData.teamMembers?.length > 0 && (
              <Card title="فريق المشروع — شريحة 4 في Pitch Deck" subtitle="مع أيقونات ملونة لكل عضو" icon="👥">
                <div className="team-grid">
                  {projectData.teamMembers.map((m, i) => (
                    <div key={i} className="team-card" style={{ borderColor: MEMBER_COLORS[i] + '44' }}>
                      <div className="team-avatar" style={{ background: MEMBER_COLORS[i] }}>👤</div>
                      <div className="team-name">{m.name}</div>
                      <div className="team-role">{m.role}</div>
                      <div className="team-bg">{m.background}</div>
                    </div>
                  ))}
                </div>
              </Card>
            )}

            {/* Script */}
            {script && (
              <Card title="سكريبت الفيديو التسويقي" subtitle="جاهز للتصوير مع توجيهات المخرج" icon="🎬">
                <div className="script-box">
                  <pre>{script}</pre>
                </div>
                <button className="btn-copy" onClick={() => navigator.clipboard.writeText(script)}>
                  📋 نسخ السكريبت
                </button>
              </Card>
            )}

            {/* Done */}
            {phase === 'done' && (
              <div className="done-banner fade-up">
                <div className="done-emoji">🎉</div>
                <h3>مشروعك Diplôme/Startup جاهز!</h3>
                <p>حمّل دليل المشروع Word + Pitch Deck PowerPoint + سكريبت الفيديو</p>
                <p className="done-sub">وفق المرسوم الوزاري 1275 • CNCSIIU • MHESR • الجزائر</p>
                <button className="btn-new" onClick={reset}>+ مشروع جديد</button>
              </div>
            )}
          </div>
        )}
      </main>

      <footer className="footer">
        StartupAI • Diplôme/Startup 1275 • CNCSIIU • MHESR • الجزائر
      </footer>
    </div>
  )
}


// ── Mount ──────────────────────────────────────────────────────────────────
ReactDOM.createRoot(document.getElementById('root')).render(
  React.createElement(React.StrictMode, null, React.createElement(App, null))
)
