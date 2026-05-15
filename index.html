<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Painel de Controle de Lotes</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,400&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

:root{
  --ink:     #01223E;
  --blue:    #113D63;
  --mid:     #1C5491;
  --sky:     #7BB6E0;
  --ice:     #C8E4F5;
  --bg:      #EDF2F7;
  --surface: #FFFFFF;
  --s100:    #E8EEF5;
  --s200:    #CFD9E6;
  --s300:    #A0B4C8;
  --s500:    #5A7A96;
  --s700:    #2D4A61;
  --red:     #C8001A;
  --green:   #057A55;
  --amber:   #B45309;
  --purple:  #6427C2;
  --r:       10px;
  --sh:      0 1px 3px rgba(1,34,62,.07), 0 6px 18px rgba(1,34,62,.08);
  --shd:     0 2px 8px rgba(1,34,62,.10), 0 16px 40px rgba(1,34,62,.12);
}

html{scroll-behavior:smooth}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--ink);min-height:100vh;font-size:14px;line-height:1.55}

/* ══ HEADER ═══════════════════════════════════════════════════════ */
.app-hdr{
  background:linear-gradient(135deg,var(--ink) 0%,var(--blue) 55%,var(--mid) 100%);
  box-shadow:0 3px 24px rgba(1,34,62,.35);
  position:sticky;top:0;z-index:200;
}
.hdr-top{
  display:flex;align-items:center;gap:18px;
  padding:16px 28px 14px;
  border-bottom:1px solid rgba(255,255,255,.09);
}
.logo{
  width:52px;height:52px;border-radius:12px;flex-shrink:0;
  background:rgba(123,182,224,.15);border:1.5px solid rgba(123,182,224,.25);
  display:flex;align-items:center;justify-content:center;
}
.hdr-brand h1{font-family:'Syne',sans-serif;font-size:1.35rem;font-weight:800;color:#fff;letter-spacing:-.02em;line-height:1.15}
.hdr-brand p{font-size:.73rem;color:var(--ice);margin-top:2px;opacity:.8}
.hdr-meta{margin-left:auto;display:flex;flex-direction:column;align-items:flex-end;gap:5px}
.hdr-pill{
  background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.16);
  border-radius:20px;padding:3px 12px;font-size:.7rem;color:rgba(255,255,255,.7);
  font-family:'DM Mono',monospace;white-space:nowrap;
}
.hdr-pill strong{color:#fff}
.hdr-nav{display:flex;padding:0 28px;overflow-x:auto;gap:2px}
.ntab{
  display:flex;align-items:center;gap:6px;
  padding:9px 16px;border:none;background:none;cursor:pointer;
  font-family:'DM Sans',sans-serif;font-size:.75rem;font-weight:600;
  color:rgba(255,255,255,.45);border-bottom:2px solid transparent;
  transition:all .2s;white-space:nowrap;
}
.ntab:hover{color:rgba(255,255,255,.8)}
.ntab.on{color:#fff;border-bottom-color:var(--sky)}
.ntab svg{opacity:.7}
.ntab.on svg{opacity:1}

/* ══ TOOLBAR ═══════════════════════════════════════════════════════ */
.toolbar{
  background:var(--surface);border-bottom:1px solid var(--s200);
  padding:10px 28px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;
}
.btn{
  display:inline-flex;align-items:center;gap:6px;
  padding:7px 16px;border-radius:7px;border:none;cursor:pointer;
  font-family:'DM Sans',sans-serif;font-size:.78rem;font-weight:600;
  transition:all .14s;white-space:nowrap;line-height:1;
}
.btn-ink{background:var(--ink);color:#fff}
.btn-ink:hover{background:#011528;transform:translateY(-1px);box-shadow:0 4px 12px rgba(1,34,62,.28)}
.btn-sky{background:var(--sky);color:var(--ink)}
.btn-sky:hover{background:#5fa3cf;color:#fff;transform:translateY(-1px)}
.btn-green{background:var(--green);color:#fff}
.btn-green:hover{background:#046242;transform:translateY(-1px)}
.btn-red{background:var(--red);color:#fff}
.btn-red:hover{background:#a3001500;transform:translateY(-1px);background:#a30015}
.btn-ghost{background:transparent;color:var(--s500);border:1.5px solid var(--s200)}
.btn-ghost:hover{background:var(--s100)}
.btn-share{background:linear-gradient(135deg,#7BB6E0,#113D63);color:#fff}
.btn-share:hover{filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 4px 14px rgba(17,61,99,.3)}
#file-input{display:none}
.upload-lbl{cursor:pointer}
.ml-auto{margin-left:auto}

/* ══ FILTERS ════════════════════════════════════════════════════════ */
.fbar{background:var(--surface);border-bottom:1px solid var(--s200)}
.finner{max-width:1360px;margin:0 auto;padding:14px 28px}
.fhead{display:flex;align-items:center;gap:8px;margin-bottom:12px}
.fhead h3{font-family:'Syne',sans-serif;font-size:.88rem;font-weight:700;color:var(--blue)}
.ftoggle{margin-left:auto;background:none;border:none;cursor:pointer;font-family:'DM Sans',sans-serif;font-size:.73rem;color:var(--s300);font-weight:500}
.ftoggle:hover{color:var(--blue)}
.fgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(148px,1fr));gap:10px}
.ff label{display:block;font-size:.66rem;font-weight:600;color:var(--s300);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px}
.ff input,.ff select{
  width:100%;border:1.5px solid var(--s200);border-radius:6px;
  padding:6px 10px;font-size:.78rem;font-family:'DM Sans',sans-serif;
  color:var(--ink);background:var(--bg);outline:none;transition:border-color .15s;
}
.ff input:focus,.ff select:focus{border-color:var(--sky);background:#fff}
.faction{display:flex;gap:8px;margin-top:10px}

/* ══ MAIN ══════════════════════════════════════════════════════════ */
.main{max-width:1360px;margin:0 auto;padding:24px 28px 56px}

/* ── STRIP ── */
.strip{
  background:linear-gradient(135deg,var(--ink),var(--blue));
  border-radius:12px;padding:18px 26px;margin-bottom:22px;
  display:flex;flex-wrap:wrap;gap:0;align-items:center;
  box-shadow:var(--sh);
}
.si{flex:1;min-width:100px;text-align:center;padding:0 12px}
.si+.si{border-left:1px solid rgba(255,255,255,.12)}
.si-lbl{font-size:.65rem;color:var(--ice);text-transform:uppercase;letter-spacing:.07em;font-weight:600;opacity:.75}
.si-val{font-size:1.6rem;font-weight:700;color:#fff;font-family:'DM Mono',monospace;line-height:1.1;margin-top:2px}
.si-val.sm{font-size:1rem;margin-top:4px}

/* ── SECTION ── */
.sec{margin-bottom:44px}
.sec-hd{display:flex;align-items:center;gap:12px;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid var(--s200)}
.sec-ico{width:36px;height:36px;border-radius:10px;flex-shrink:0;display:flex;align-items:center;justify-content:center}
.ico-m{background:#DBEAFE} .ico-f{background:#FEF3C7}
.ico-e{background:#D1FAE5} .ico-l{background:#EDE9FE}
.ico-t{background:var(--s100)}
.sec-hd h2{font-family:'Syne',sans-serif;font-size:1rem;font-weight:700;color:var(--ink)}
.sec-badge{margin-left:auto;background:var(--s100);border-radius:20px;padding:3px 11px;font-size:.7rem;font-family:'DM Mono',monospace;color:var(--s500)}

/* ── GRID ── */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
@media(max-width:1100px){.g4{grid-template-columns:1fr 1fr}}
@media(max-width:780px){.g2,.g3,.g4{grid-template-columns:1fr}}

/* ── CARD ── */
.card{
  background:var(--surface);border-radius:12px;box-shadow:var(--sh);
  border:1px solid var(--s100);padding:20px 22px;transition:box-shadow .2s;
}
.card:hover{box-shadow:var(--shd)}
.ctitle{font-family:'Syne',sans-serif;font-size:.82rem;font-weight:700;color:var(--blue);margin-bottom:14px;display:flex;align-items:center;gap:7px}

/* ── KPI ── */
.kpi{
  background:var(--surface);border-radius:12px;border:1px solid var(--s100);
  box-shadow:var(--sh);padding:18px 20px;position:relative;overflow:hidden;
  transition:box-shadow .2s,transform .2s;
}
.kpi:hover{box-shadow:var(--shd);transform:translateY(-2px)}
.kpi::after{
  content:'';position:absolute;left:0;top:0;bottom:0;width:3px;
  background:var(--kc,var(--sky));border-radius:3px 0 0 3px;
}
.klbl{font-size:.65rem;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--s300);margin-bottom:6px}
.kval{font-size:1.85rem;font-weight:700;color:var(--ink);line-height:1;font-family:'DM Mono',monospace}
.kval.sm{font-size:1rem;margin-top:4px;font-family:'DM Sans',sans-serif;font-weight:600}
.ksub{font-size:.68rem;color:var(--s300);margin-top:5px}
.kpi.dark{background:linear-gradient(135deg,var(--blue),var(--ink));border-color:transparent}
.kpi.dark .klbl{color:var(--ice);opacity:.7}
.kpi.dark .kval{color:#fff}
.kpi.dark .ksub{color:rgba(255,255,255,.4)}

/* ── RANK ── */
.rl{display:flex;flex-direction:column;gap:8px}
.ri{
  display:flex;align-items:center;gap:10px;
  padding:9px 12px;border-radius:8px;background:var(--bg);
  border:1px solid var(--s100);transition:background .15s;
}
.ri:hover{background:var(--ice)}
.rpos{
  width:22px;height:22px;border-radius:50%;flex-shrink:0;
  background:var(--blue);color:#fff;font-size:.63rem;font-weight:700;
  display:flex;align-items:center;justify-content:center;font-family:'DM Mono',monospace;
}
.rpos.p1{background:#B7950B}.rpos.p2{background:var(--s300)}.rpos.p3{background:#8B4513}
.rname{font-size:.8rem;font-weight:600;color:var(--ink);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.rbar-wrap{width:70px;height:5px;background:var(--s200);border-radius:3px;flex-shrink:0}
.rbar{height:100%;border-radius:3px;background:var(--sky);transition:width .4s}
.rbadge{
  flex-shrink:0;background:var(--blue);color:#fff;
  border-radius:20px;padding:3px 11px;font-size:.7rem;font-weight:700;
  font-family:'DM Mono',monospace;
}
.rbadge.am{background:var(--amber)}.rbadge.gr{background:var(--green)}.rbadge.pu{background:var(--purple)}

/* ── PROGRESS ── */
.pl{display:flex;flex-direction:column;gap:10px}
.pi-hd{display:flex;justify-content:space-between;margin-bottom:4px}
.pi-nm{font-size:.78rem;font-weight:600;color:var(--ink)}
.pi-vl{font-size:.7rem;color:var(--s300);font-family:'DM Mono',monospace}
.pi-track{height:8px;background:var(--s200);border-radius:4px;overflow:hidden}
.pi-fill{height:100%;border-radius:4px;transition:width .5s}

/* ── CHART ── */
.cw{position:relative}
.cw canvas{display:block}
.cempty{
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  height:100%;gap:8px;color:var(--s300);font-size:.78rem;
}

/* ── TABLE ── */
.twrap{overflow-x:auto;border-radius:8px;border:1px solid var(--s200)}
table{width:100%;border-collapse:collapse;font-size:.77rem}
thead th{
  background:var(--blue);color:#fff;padding:10px 14px;text-align:left;
  font-weight:600;white-space:nowrap;font-size:.7rem;letter-spacing:.03em;
  font-family:'Syne',sans-serif;cursor:pointer;user-select:none;
}
thead th:hover{background:var(--mid)}
tbody tr:nth-child(even){background:var(--bg)}
tbody tr:hover{background:var(--ice)}
tbody td{padding:8px 14px;color:var(--s700);border-bottom:1px solid var(--s100);white-space:nowrap}
.te{text-align:center;padding:40px;color:var(--s300);font-size:.82rem}

/* ── BADGE ── */
.badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:12px;font-size:.65rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em}
.bm{background:#DBEAFE;color:#1E40AF}
.bf{background:#FEF3C7;color:#92400E}
.be{background:#D1FAE5;color:#065F46}
.bl{background:#EDE9FE;color:#5B21B6}

/* ── EMPTY ── */
.empty{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:32px;gap:10px;color:var(--s300)}
.empty svg{opacity:.3}
.empty p{font-size:.8rem}

/* ── TOAST ── */
#toast{
  position:fixed;bottom:24px;right:24px;z-index:9999;
  background:var(--ink);color:#fff;padding:12px 18px;
  border-radius:10px;font-size:.78rem;font-weight:600;
  box-shadow:var(--shd);opacity:0;transform:translateY(16px);
  transition:all .25s;pointer-events:none;display:flex;align-items:center;gap:8px;max-width:340px;
}
#toast.show{opacity:1;transform:translateY(0)}
#toast.err{background:var(--red)}

/* ── MT ── */
.mt{margin-top:16px}

/* ── PRINT ── */
@media print{
  .toolbar,.fbar,.hdr-nav,.noprint{display:none!important}
  body{background:#fff}
  .main{padding:8px}
  .card,.kpi{box-shadow:none;border:1px solid #ccc}
  .app-hdr{position:static}
}

/* ── ANIM ── */
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
.anim{animation:fadeUp .35s ease both}
.d1{animation-delay:.04s}.d2{animation-delay:.09s}.d3{animation-delay:.14s}.d4{animation-delay:.19s}
</style>
</head>
<body>

<!-- HEADER -->
<header class="app-hdr">
  <div class="hdr-top">
    <div class="logo">
      <svg width="26" height="26" viewBox="0 0 28 28" fill="none">
        <rect x="2" y="9" width="24" height="16" rx="3" stroke="#7BB6E0" stroke-width="2"/>
        <path d="M9 9V7a5 5 0 0 1 10 0v2" stroke="#7BB6E0" stroke-width="2" stroke-linecap="round"/>
        <circle cx="14" cy="17" r="2.5" fill="#7BB6E0"/>
        <path d="M14 19.5V22" stroke="#7BB6E0" stroke-width="2" stroke-linecap="round"/>
      </svg>
    </div>
    <div class="hdr-brand">
      <h1>Painel de Controle de Lotes</h1>
      <p>Controle de Saídas de Materiais, Ferramentas, EPIs e Logística</p>
    </div>
    <div class="hdr-meta">
      <div class="hdr-pill">📅 <span id="hdate">—</span></div>
      <div class="hdr-pill">Exibindo <strong id="hcount">0</strong> registros</div>
    </div>
  </div>
  <nav class="hdr-nav">
    <button class="ntab on" onclick="goTo('sm')">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 7H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2Z"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>
      Materiais
    </button>
    <button class="ntab" onclick="goTo('sf')">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/></svg>
      Ferramentas
    </button>
    <button class="ntab" onclick="goTo('se')">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
      EPIs
    </button>
    <button class="ntab" onclick="goTo('sl')">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="1" y="3" width="15" height="13" rx="2"/><path d="M16 8h4l3 3v5h-7V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>
      Logística
    </button>
    <button class="ntab" onclick="goTo('st')">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>
      Registros
    </button>
  </nav>
</header>

<!-- TOOLBAR -->
<div class="toolbar noprint">
  <input type="file" id="file-input" accept=".xlsx,.xls">
  <label for="file-input" class="upload-lbl">
    <span class="btn btn-ink">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/></svg>
      Importar XLSX
    </span>
  </label>
  <button class="btn btn-sky" onclick="dlTemplate()">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
    Baixar Template
  </button>
  <button class="btn btn-green" onclick="loadDemo()">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>
    Demo
  </button>
  <button class="btn btn-ghost" onclick="clearAll()">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>
    Limpar
  </button>
  <button class="btn btn-share noprint" onclick="openShare()">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
    Compartilhar
  </button>
  <button class="btn btn-red ml-auto noprint" onclick="window.print()">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 9V3a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v6"/><rect x="6" y="14" width="12" height="8" rx="1"/></svg>
    Imprimir
  </button>
</div>

<!-- MODAL COMPARTILHAR -->
<div id="share-overlay" onclick="closeShare()" style="display:none;position:fixed;inset:0;background:rgba(1,34,62,.55);z-index:1000;align-items:center;justify-content:center;backdrop-filter:blur(3px)">
  <div onclick="event.stopPropagation()" style="background:#fff;border-radius:16px;padding:28px 28px 24px;max-width:520px;width:calc(100% - 32px);box-shadow:0 24px 64px rgba(1,34,62,.25);position:relative">
    <button onclick="closeShare()" style="position:absolute;top:14px;right:14px;background:none;border:none;cursor:pointer;font-size:1.2rem;color:#A0B4C8;line-height:1">✕</button>
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:18px">
      <div style="width:40px;height:40px;border-radius:10px;background:#EDF2F7;display:flex;align-items:center;justify-content:center;flex-shrink:0">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#113D63" stroke-width="2.5"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
      </div>
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:1rem;font-weight:700;color:#01223E">Compartilhar Painel</div>
        <div style="font-size:.75rem;color:#A0B4C8;margin-top:2px">Os dados ficam embutidos no link — sem servidor necessário</div>
      </div>
    </div>

    <div id="share-status" style="font-size:.78rem;color:#5A7A96;margin-bottom:12px;min-height:20px"></div>

    <div style="display:flex;gap:8px;margin-bottom:16px">
      <input id="share-url" readonly style="flex:1;border:1.5px solid #CFD9E6;border-radius:7px;padding:9px 12px;font-size:.75rem;font-family:'DM Mono',monospace;color:#01223E;background:#EDF2F7;outline:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" placeholder="Gerando link...">
      <button class="btn btn-ink" id="copy-btn" onclick="copyLink()" style="flex-shrink:0">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
        Copiar
      </button>
    </div>

    <div id="share-warn" style="display:none;background:#FEF3C7;border:1px solid #FDE68A;border-radius:8px;padding:10px 14px;font-size:.75rem;color:#92400E;line-height:1.5;margin-bottom:14px">
      ⚠️ <strong>Link muito grande.</strong> Alguns navegadores ou apps de mensagem podem truncar URLs longas.<br>
      Nesse caso, use <strong>Exportar CSV</strong> e envie o arquivo junto com o link do painel.
    </div>

    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <button class="btn" style="background:#25D366;color:#fff;font-size:.75rem" onclick="shareWA()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
        WhatsApp
      </button>
      <button class="btn btn-ghost" style="font-size:.75rem" onclick="closeShare()">Fechar</button>
    </div>
  </div>
</div>

<!-- FILTERS -->
<div class="fbar noprint">
  <div class="finner">
    <div class="fhead">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#113D63" stroke-width="2.5"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>
      <h3>Filtros</h3>
      <button class="ftoggle" id="ftbtn" onclick="toggleF()">▲ Ocultar</button>
    </div>
    <div id="fbody">
      <div class="fgrid">
        <div class="ff"><label>Data início</label><input type="date" id="fds" onchange="applyF()"></div>
        <div class="ff"><label>Data fim</label><input type="date" id="fde" onchange="applyF()"></div>
        <div class="ff"><label>Setor (Mat./Fer.)</label>
          <select id="fset" onchange="applyF()"><option value="">Todos</option></select></div>
        <div class="ff"><label>Colaborador (Mat.)</label>
          <select id="fcm" onchange="applyF()"><option value="">Todos</option></select></div>
        <div class="ff"><label>Colaborador (EPI)</label>
          <select id="fce" onchange="applyF()"><option value="">Todos</option></select></div>
        <div class="ff"><label>Unidade (EPI)</label>
          <select id="fun" onchange="applyF()"><option value="">Todas</option></select></div>
        <div class="ff"><label>Unidade (Logística)</label>
          <select id="flu" onchange="applyF()"><option value="">Todas</option></select></div>
        <div class="ff"><label>Tipo de Envio</label>
          <select id="fte" onchange="applyF()"><option value="">Todos</option></select></div>
      </div>
      <div class="faction">
        <button class="btn btn-ghost" onclick="resetF()" style="padding:5px 12px;font-size:.72rem">✕ Limpar filtros</button>
      </div>
    </div>
  </div>
</div>

<!-- MAIN -->
<main class="main">

  <!-- STRIP -->
  <div class="strip anim">
    <div class="si"><div class="si-lbl">Total Registros</div><div class="si-val" id="ss0">0</div></div>
    <div class="si"><div class="si-lbl">Saídas Materiais</div><div class="si-val" id="ss1">0</div></div>
    <div class="si"><div class="si-lbl">Qtd. Materiais</div><div class="si-val" id="ss2">0</div></div>
    <div class="si"><div class="si-lbl">Saídas Ferramentas</div><div class="si-val" id="ss3">0</div></div>
    <div class="si"><div class="si-lbl">Saídas EPIs</div><div class="si-val" id="ss4">0</div></div>
    <div class="si"><div class="si-lbl">Gasto EPI</div><div class="si-val sm" id="ss5">R$ 0,00</div></div>
    <div class="si"><div class="si-lbl">Envios Logística</div><div class="si-val" id="ss6">0</div></div>
  </div>

  <!-- ██ MATERIAIS █████████████████████████████████████████████████ -->
  <div class="sec anim d1" id="sm">
    <div class="sec-hd">
      <div class="sec-ico ico-m">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#1E40AF" stroke-width="2.2"><path d="M20 7H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2Z"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>
      </div>
      <h2>Controle de Saídas de Materiais</h2>
      <span class="sec-badge" id="cm-cnt">0 registros</span>
    </div>
    <div class="g4" style="margin-bottom:16px">
      <div class="kpi" style="--kc:#1E40AF">
        <div class="klbl">Saídas (registros)</div>
        <div class="kval" id="mk1">0</div><div class="ksub">Linhas da aba Dados</div>
      </div>
      <div class="kpi" style="--kc:#7BB6E0">
        <div class="klbl">Quantidade total</div>
        <div class="kval" id="mk2">0</div><div class="ksub">Unidades retiradas</div>
      </div>
      <div class="kpi" style="--kc:#059669">
        <div class="klbl">Item mais retirado</div>
        <div class="kval sm" id="mk3">—</div><div class="ksub" id="mk3s"></div>
      </div>
      <div class="kpi" style="--kc:#D97706">
        <div class="klbl">Colaborador top</div>
        <div class="kval sm" id="mk4">—</div><div class="ksub" id="mk4s"></div>
      </div>
    </div>
    <div class="g2">
      <div class="card">
        <div class="ctitle">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#113D63" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>
          Top 7 — Itens mais retirados
        </div>
        <div class="rl" id="m-items"><div class="empty"><p>Sem dados</p></div></div>
      </div>
      <div class="card">
        <div class="ctitle">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#113D63" stroke-width="2.5"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
          Top 7 — Colaboradores com mais retiradas
        </div>
        <div class="rl" id="m-colab"><div class="empty"><p>Sem dados</p></div></div>
      </div>
    </div>
    <div class="card mt">
      <div class="ctitle">Saídas por Setor (Materiais)</div>
      <div class="cw" style="height:200px"><canvas id="ch-m-set"></canvas><div class="cempty" id="ce-m-set"><svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M3 3v18h18"/><path d="M18 9l-5 5-2-2-5 5"/></svg>Sem dados</div></div>
    </div>
  </div>

  <!-- ██ FERRAMENTAS ███████████████████████████████████████████████ -->
  <div class="sec anim d2" id="sf">
    <div class="sec-hd">
      <div class="sec-ico ico-f">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#92400E" stroke-width="2.2"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/></svg>
      </div>
      <h2>Controle de Saídas de Ferramentas</h2>
      <span class="sec-badge" id="cf-cnt">0 registros</span>
    </div>
    <div class="g4" style="margin-bottom:16px">
      <div class="kpi" style="--kc:#92400E">
        <div class="klbl">Saídas (registros)</div>
        <div class="kval" id="fk1">0</div><div class="ksub">Linhas da aba Dados</div>
      </div>
      <div class="kpi" style="--kc:#FDE68A">
        <div class="klbl">Quantidade total</div>
        <div class="kval" id="fk2">0</div><div class="ksub">Ferramentas retiradas</div>
      </div>
      <div class="kpi" style="--kc:#D97706">
        <div class="klbl">Ferramenta top</div>
        <div class="kval sm" id="fk3">—</div><div class="ksub" id="fk3s"></div>
      </div>
      <div class="kpi" style="--kc:#B45309">
        <div class="klbl">Colaborador top</div>
        <div class="kval sm" id="fk4">—</div><div class="ksub" id="fk4s"></div>
      </div>
    </div>
    <div class="g2">
      <div class="card">
        <div class="ctitle">Top 7 — Ferramentas com mais saídas</div>
        <div class="rl" id="f-items"><div class="empty"><p>Sem dados</p></div></div>
      </div>
      <div class="card">
        <div class="ctitle">Top 7 — Colaboradores que mais retiraram ferramentas</div>
        <div class="rl" id="f-colab"><div class="empty"><p>Sem dados</p></div></div>
      </div>
    </div>
  </div>

  <!-- ██ EPI ███████████████████████████████████████████████████████ -->
  <div class="sec anim d3" id="se">
    <div class="sec-hd">
      <div class="sec-ico ico-e">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#065F46" stroke-width="2.2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
      </div>
      <h2>Controle de Carregamentos de EPI</h2>
      <span class="sec-badge" id="ce-cnt">0 registros</span>
    </div>
    <div class="g4" style="margin-bottom:16px">
      <div class="kpi dark" style="--kc:var(--sky)">
        <div class="klbl">Gasto Total EPI</div>
        <div class="kval" style="font-size:1.2rem" id="ek1">R$ 0,00</div><div class="ksub">Preço × Saídas</div>
      </div>
      <div class="kpi" style="--kc:#057A55">
        <div class="klbl">Saídas registradas</div>
        <div class="kval" id="ek2">0</div><div class="ksub">Linhas da aba EPIs</div>
      </div>
      <div class="kpi" style="--kc:#BBF7D0">
        <div class="klbl">EPI mais retirado</div>
        <div class="kval sm" id="ek3">—</div><div class="ksub" id="ek3s"></div>
      </div>
      <div class="kpi" style="--kc:#057A55">
        <div class="klbl">Setor com + saídas</div>
        <div class="kval sm" id="ek4">—</div><div class="ksub" id="ek4s"></div>
      </div>
    </div>
    <div class="g2">
      <div class="card">
        <div class="ctitle">EPIs com mais saídas (unidades)</div>
        <div class="pl" id="e-prog"><div class="empty"><p>Sem dados</p></div></div>
      </div>
      <div class="card">
        <div class="ctitle">Saídas por Setor</div>
        <div class="cw" style="height:220px"><canvas id="ch-e-set"></canvas><div class="cempty" id="ce-e-set">Sem dados</div></div>
      </div>
    </div>
    <div class="g2 mt">
      <div class="card">
        <div class="ctitle">Distribuição por Unidade</div>
        <div class="cw" style="height:220px"><canvas id="ch-e-uni"></canvas><div class="cempty" id="ce-e-uni">Sem dados</div></div>
      </div>
      <div class="card">
        <div class="ctitle">Gasto EPI por dia (R$)</div>
        <div class="cw" style="height:220px"><canvas id="ch-e-dia"></canvas><div class="cempty" id="ce-e-dia">Sem dados</div></div>
      </div>
    </div>
    <div class="card mt">
      <div class="ctitle">Top 7 — Colaboradores com mais EPIs</div>
      <div class="rl" id="e-colab"><div class="empty"><p>Sem dados</p></div></div>
    </div>
  </div>

  <!-- ██ LOGÍSTICA ████████████████████████████████████████████████ -->
  <div class="sec anim d4" id="sl">
    <div class="sec-hd">
      <div class="sec-ico ico-l">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#5B21B6" stroke-width="2.2"><rect x="1" y="3" width="15" height="13" rx="2"/><path d="M16 8h4l3 3v5h-7V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>
      </div>
      <h2>Logística</h2>
      <span class="sec-badge" id="cl-cnt">0 registros</span>
    </div>
    <div class="g4" style="margin-bottom:16px">
      <div class="kpi dark" style="--kc:#A78BFA">
        <div class="klbl">Unidade com + envios</div>
        <div class="kval sm" id="lk1">N/D</div><div class="ksub" id="lk1s">0 envios</div>
      </div>
      <div class="kpi" style="--kc:#6427C2">
        <div class="klbl">Total de envios</div>
        <div class="kval" id="lk2">0</div><div class="ksub">Linhas da aba Logística</div>
      </div>
      <div class="kpi" style="--kc:#DDD6FE">
        <div class="klbl">Material mais enviado</div>
        <div class="kval sm" id="lk3">—</div><div class="ksub" id="lk3s"></div>
      </div>
      <div class="kpi" style="--kc:#A78BFA">
        <div class="klbl">Tipo de envio top</div>
        <div class="kval sm" id="lk4">—</div><div class="ksub" id="lk4s"></div>
      </div>
    </div>
    <div class="g2">
      <div class="card">
        <div class="ctitle">Unidades com mais envios</div>
        <div class="rl" id="l-uni"><div class="empty"><p>Sem dados</p></div></div>
      </div>
      <div class="card">
        <div class="ctitle">Materiais enviados por tipo</div>
        <div class="cw" style="height:220px"><canvas id="ch-l-mat"></canvas><div class="cempty" id="ce-l-mat">Sem dados</div></div>
      </div>
    </div>
    <div class="card mt">
      <div class="ctitle">Tipos de envio (VOO, BARCO, etc.)</div>
      <div class="cw" style="height:180px"><canvas id="ch-l-tipo"></canvas><div class="cempty" id="ce-l-tipo">Sem dados</div></div>
    </div>
  </div>

  <!-- ██ TABELA ████████████████████████████████████████████████████ -->
  <div class="sec" id="st">
    <div class="sec-hd">
      <div class="sec-ico ico-t">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#475569" stroke-width="2.2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>
      </div>
      <h2>Registros Detalhados</h2>
      <span class="sec-badge" id="ct-cnt">0</span>
      <button class="btn btn-sky noprint" onclick="exportCSV()" style="margin-left:8px;padding:5px 12px;font-size:.72rem">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
        Exportar CSV
      </button>
    </div>
    <div class="card" style="padding:0;overflow:hidden">
      <div class="twrap">
        <table>
          <thead><tr>
            <th onclick="srt('tipo')">Tipo ↕</th>
            <th onclick="srt('data')">Data ↕</th>
            <th onclick="srt('item')">Item / EPI / Material ↕</th>
            <th onclick="srt('colab')">Colaborador ↕</th>
            <th onclick="srt('funcao')" >Função</th>
            <th onclick="srt('setor')">Setor ↕</th>
            <th onclick="srt('unidade')">Unidade</th>
            <th onclick="srt('qtd')" style="text-align:right">Qtd ↕</th>
            <th onclick="srt('preco')" style="text-align:right">Preço Unit.</th>
            <th onclick="srt('total')" style="text-align:right">Total ↕</th>
          </tr></thead>
          <tbody id="tbody"><tr><td colspan="10" class="te">Importe um XLSX ou clique em <strong>Demo</strong> para visualizar.</td></tr></tbody>
        </table>
      </div>
    </div>
  </div>

</main>

<div id="toast"><span id="tmsg"></span></div>

<script>
/* ══ STATE ══════════════════════════════════════════════════════════ */
const $ = id => document.getElementById(id);
let RAW = {mat:[], fer:[], epi:[], log:[]};
let FIL = {mat:[], fer:[], epi:[], log:[]};
let FLAT = [];
let sortK = 'data', sortD = 1;
const CH = {};

/* ══ UTILS ══════════════════════════════════════════════════════════ */
const fmt   = n => Number(n||0).toLocaleString('pt-BR');
const fmtR  = n => 'R$ ' + Number(n||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const clamp = (v,a,b) => Math.min(b,Math.max(a,v));

function toast(msg, err=false){
  const t=$('toast'); $('tmsg').textContent=msg;
  t.className = err?'err show':'show';
  clearTimeout(t._t); t._t=setTimeout(()=>t.className='',3000);
}

/* ══ DATE ══════════════════════════════════════════════════════════ */
$('hdate').textContent = new Date().toLocaleDateString('pt-BR');

/* ══ NAV ═══════════════════════════════════════════════════════════ */
function goTo(id){
  document.getElementById(id)?.scrollIntoView({behavior:'smooth',block:'start'});
}
const SECS = ['sm','sf','se','sl','st'];
const TABS = document.querySelectorAll('.ntab');
window.addEventListener('scroll',()=>{
  let cur=0;
  SECS.forEach((id,i)=>{const el=document.getElementById(id);if(el&&el.getBoundingClientRect().top<=130)cur=i;});
  TABS.forEach((t,i)=>t.classList.toggle('on',i===cur));
},{passive:true});

/* ══ PARSE DATE ════════════════════════════════════════════════════ */
function pd(v){
  if(!v)return null;
  if(v instanceof Date && !isNaN(v))return v;
  const s=String(v).trim();
  const m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if(m)return new Date(+m[3],+m[2]-1,+m[1]);
  const d=new Date(s); return isNaN(d.getTime())?null:d;
}

/* ══ IMPORT XLSX ════════════════════════════════════════════════════
   Aba "Dados"  cols: DATA | MAT-ITEM | MAT-SAIDAS | COLAB-MAT | FUNCAO | SETOR | DATA | FERR | FERR-SAIDAS | COLAB-FERR | FUNCAO | SETOR
   Aba "EPIs"   cols: DATA | EPI-NOME | PRECO | EPI-SAIDAS | SETOR | COLABORADOR | FUNCAO | UNIDADE
   Aba "Logística" cols: DATA | UNIDADE | MATERIAIS ENVIADOS | TIPO DE ENVIO
══════════════════════════════════════════════════════════════════ */
$('file-input').addEventListener('change',function(e){
  const file=e.target.files[0]; if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:'binary',cellDates:true});
      parseSheets(wb);
      toast(`✓ Importado: ${RAW.mat.length} materiais · ${RAW.fer.length} ferramentas · ${RAW.epi.length} EPIs · ${RAW.log.length} logística`);
    }catch(err){toast('Erro ao ler XLSX. Verifique o arquivo.',true);}
  };
  reader.readAsBinaryString(file);
  e.target.value='';
});

function parseSheets(wb){
  RAW={mat:[],fer:[],epi:[],log:[]};

  /* ─ Aba Dados ─ */
  const dadosName = wb.SheetNames.find(n=>/dados/i.test(n));
  if(dadosName){
    const ws=wb.Sheets[dadosName];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
    // skip header row (row 0)
    for(let i=1;i<rows.length;i++){
      const r=rows[i];
      // Materiais: cols 0-5
      if(r[1]!=null){
        RAW.mat.push({
          data:    pd(r[0]),
          item:    String(r[1]||'').trim(),
          qtd:     Number(r[2])||0,
          colab:   String(r[3]||'').trim(),
          funcao:  String(r[4]||'').trim(),
          setor:   String(r[5]||'').trim(),
        });
      }
      // Ferramentas: cols 6-11
      if(r[7]!=null){
        RAW.fer.push({
          data:    pd(r[6]),
          item:    String(r[7]||'').trim(),
          qtd:     Number(r[8])||0,
          colab:   String(r[9]||'').trim(),
          funcao:  String(r[10]||'').trim(),
          setor:   String(r[11]||'').trim(),
        });
      }
    }
  }

  /* ─ Aba EPIs ─ */
  const episName = wb.SheetNames.find(n=>/epi/i.test(n));
  if(episName){
    const ws=wb.Sheets[episName];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
    for(let i=1;i<rows.length;i++){
      const r=rows[i];
      if(r[1]!=null){
        RAW.epi.push({
          data:    pd(r[0]),
          item:    String(r[1]||'').trim(),
          preco:   Number(r[2])||0,
          qtd:     Number(r[3])||0,
          setor:   String(r[4]||'').trim(),
          colab:   String(r[5]||'').trim(),
          funcao:  String(r[6]||'').trim(),
          unidade: String(r[7]||'').trim(),
        });
      }
    }
  }

  /* ─ Aba Logística ─ */
  const logName = wb.SheetNames.find(n=>/log/i.test(n));
  if(logName){
    const ws=wb.Sheets[logName];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
    for(let i=1;i<rows.length;i++){
      const r=rows[i];
      if(r[1]!=null){
        RAW.log.push({
          data:     pd(r[0]),
          unidade:  String(r[1]||'').trim(),
          material: String(r[2]||'').trim(),
          tipo:     String(r[3]||'').trim(),
        });
      }
    }
  }

  populateDDs();
  applyF();
}

/* ══ POPULATE DROPDOWNS ════════════════════════════════════════════ */
function populateDDs(){
  const uniq = (arr,fn) => [...new Set(arr.map(fn).filter(Boolean))].sort();
  const fill = (id,opts) => {
    const cur=$(id).value;
    $(id).innerHTML='<option value="">Todos</option>'+opts.map(v=>`<option value="${v}"${v===cur?' selected':''}>${v}</option>`).join('');
  };
  const setors = uniq([...RAW.mat,...RAW.fer], r=>r.setor);
  fill('fset',setors);
  fill('fcm',uniq(RAW.mat,r=>r.colab));
  fill('fce',uniq(RAW.epi,r=>r.colab));
  fill('fun',uniq(RAW.epi,r=>r.unidade));
  fill('flu',uniq(RAW.log,r=>r.unidade));
  fill('fte',uniq(RAW.log,r=>r.tipo));
}

/* ══ FILTERS ════════════════════════════════════════════════════════ */
let filOn=true;
function toggleF(){
  filOn=!filOn;
  $('fbody').style.display=filOn?'':'none';
  $('ftbtn').textContent=filOn?'▲ Ocultar':'▼ Mostrar';
}

function applyF(){
  const ds=$('fds').value, de=$('fde').value;
  const fset=$('fset').value, fcm=$('fcm').value, fce=$('fce').value;
  const fun=$('fun').value, flu=$('flu').value, fte=$('fte').value;

  const dateOk=(d,ds,de)=>{
    if(!d)return true;
    if(ds && d<new Date(ds))return false;
    if(de && d>new Date(de+'T23:59:59'))return false;
    return true;
  };

  FIL.mat = RAW.mat.filter(r=>
    dateOk(r.data,ds,de) &&
    (!fset||r.setor===fset) &&
    (!fcm||r.colab===fcm)
  );
  FIL.fer = RAW.fer.filter(r=>
    dateOk(r.data,ds,de) &&
    (!fset||r.setor===fset)
  );
  FIL.epi = RAW.epi.filter(r=>
    dateOk(r.data,ds,de) &&
    (!fce||r.colab===fce) &&
    (!fun||r.unidade===fun)
  );
  FIL.log = RAW.log.filter(r=>
    dateOk(r.data,ds,de) &&
    (!flu||r.unidade===flu) &&
    (!fte||r.tipo===fte)
  );

  const total = FIL.mat.length+FIL.fer.length+FIL.epi.length+FIL.log.length;
  $('hcount').textContent = fmt(total);
  renderAll();
}

function resetF(){
  ['fds','fde'].forEach(id=>$(id).value='');
  ['fset','fcm','fce','fun','flu','fte'].forEach(id=>$(id).value='');
  applyF();
}

/* ══ TOP-N ══════════════════════════════════════════════════════════ */
function topN(arr,kFn,vFn=()=>1,n=7){
  const m={};
  arr.forEach(r=>{const k=kFn(r);if(k&&k!=='-')m[k]=(m[k]||0)+vFn(r);});
  return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,n);
}

/* ══ RENDER RANK LIST ═══════════════════════════════════════════════ */
function rRank(id,entries,bc=''){
  const el=$(id);
  if(!entries.length){el.innerHTML='<div class="empty"><p>Sem dados</p></div>';return;}
  const max=entries[0][1];
  el.innerHTML=entries.map(([nm,vl],i)=>`
    <div class="ri">
      <div class="rpos ${i===0?'p1':i===1?'p2':i===2?'p3':''}">${i+1}</div>
      <div class="rname" title="${nm}">${nm}</div>
      <div class="rbar-wrap"><div class="rbar" style="width:${clamp(vl/max*100,2,100)}%"></div></div>
      <span class="rbadge ${bc}">${fmt(vl)}</span>
    </div>`).join('');
}

/* ══ RENDER PROGRESS ════════════════════════════════════════════════ */
function rProg(id,entries,color='#057A55'){
  const el=$(id);
  if(!entries.length){el.innerHTML='<div class="empty"><p>Sem dados</p></div>';return;}
  const max=entries[0][1];
  el.innerHTML=entries.map(([nm,vl])=>`
    <div>
      <div class="pi-hd"><span class="pi-nm" title="${nm}">${nm}</span><span class="pi-vl">${fmt(vl)} unid.</span></div>
      <div class="pi-track"><div class="pi-fill" style="width:${clamp(vl/max*100,2,100)}%;background:${color}"></div></div>
    </div>`).join('');
}

/* ══ CHARTS ══════════════════════════════════════════════════════════ */
const PAL=['#7BB6E0','#113D63','#01223E','#057A55','#D97706','#C8001A','#6427C2','#0EA5E9','#84CC16'];

function dCh(id){if(CH[id]){CH[id].destroy();delete CH[id];}}

function rBar(cid,eid,labels,data,color='#113D63'){
  const c=$(cid),e=$(eid);
  if(!labels.length){c.style.display='none';e.style.display='flex';dCh(cid);return;}
  c.style.display='block';e.style.display='none';dCh(cid);
  CH[cid]=new Chart(c,{type:'bar',data:{labels,datasets:[{data,backgroundColor:color,borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{x:{grid:{display:false},ticks:{font:{size:10},maxRotation:35}},
               y:{grid:{color:'#E8EEF5'},ticks:{font:{size:10,family:"'DM Mono'"}}}}
    }});
}

function rDonut(cid,eid,labels,data){
  const c=$(cid),e=$(eid);
  if(!labels.length){c.style.display='none';e.style.display='flex';dCh(cid);return;}
  c.style.display='block';e.style.display='none';dCh(cid);
  CH[cid]=new Chart(c,{type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:PAL.slice(0,data.length),borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{legend:{position:'right',labels:{font:{size:10},boxWidth:12,padding:10}}}}
  });
}

function rLine(cid,eid,labels,data,color='#057A55'){
  const c=$(cid),e=$(eid);
  if(!labels.length){c.style.display='none';e.style.display='flex';dCh(cid);return;}
  c.style.display='block';e.style.display='none';dCh(cid);
  CH[cid]=new Chart(c,{type:'line',
    data:{labels,datasets:[{data,borderColor:color,backgroundColor:color+'22',fill:true,tension:.4,pointRadius:3,borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{x:{grid:{display:false},ticks:{font:{size:10},maxRotation:35}},
               y:{grid:{color:'#E8EEF5'},ticks:{font:{size:10,family:"'DM Mono'"}}}}
    }});
}

/* ══ RENDER ALL ══════════════════════════════════════════════════════ */
function renderAll(){
  const M=FIL.mat, F=FIL.fer, E=FIL.epi, L=FIL.log;

  /* ─ Strip ─ */
  $('ss0').textContent=fmt(M.length+F.length+E.length+L.length);
  $('ss1').textContent=fmt(M.length);
  $('ss2').textContent=fmt(M.reduce((s,r)=>s+r.qtd,0));
  $('ss3').textContent=fmt(F.length);
  $('ss4').textContent=fmt(E.length);
  const gastoEPI=E.reduce((s,r)=>s+r.preco*r.qtd,0);
  $('ss5').textContent=fmtR(gastoEPI);
  $('ss6').textContent=fmt(L.length);

  /* ─ Materiais ─ */
  $('cm-cnt').textContent=fmt(M.length)+' registros';
  $('mk1').textContent=fmt(M.length);
  $('mk2').textContent=fmt(M.reduce((s,r)=>s+r.qtd,0));
  const mtI=topN(M,r=>r.item,r=>r.qtd||1,1), mtC=topN(M,r=>r.colab,r=>r.qtd||1,1);
  $('mk3').textContent=mtI[0]?.[0]||'—'; $('mk3s').textContent=mtI[0]?fmt(mtI[0][1])+' unid.':'';
  $('mk4').textContent=mtC[0]?.[0]||'—'; $('mk4s').textContent=mtC[0]?fmt(mtC[0][1])+' retiradas':'';
  rRank('m-items',topN(M,r=>r.item,r=>r.qtd||1));
  rRank('m-colab',topN(M,r=>r.colab,r=>r.qtd||1));
  const mSet=topN(M,r=>r.setor,r=>r.qtd||1,8);
  rBar('ch-m-set','ce-m-set',mSet.map(e=>e[0]),mSet.map(e=>e[1]),'#113D63');

  /* ─ Ferramentas ─ */
  $('cf-cnt').textContent=fmt(F.length)+' registros';
  $('fk1').textContent=fmt(F.length);
  $('fk2').textContent=fmt(F.reduce((s,r)=>s+r.qtd,0));
  const ftI=topN(F,r=>r.item,r=>r.qtd||1,1), ftC=topN(F,r=>r.colab,r=>r.qtd||1,1);
  $('fk3').textContent=ftI[0]?.[0]||'—'; $('fk3s').textContent=ftI[0]?fmt(ftI[0][1])+' unid.':'';
  $('fk4').textContent=ftC[0]?.[0]||'—'; $('fk4s').textContent=ftC[0]?fmt(ftC[0][1])+' retiradas':'';
  rRank('f-items',topN(F,r=>r.item,r=>r.qtd||1),'am');
  rRank('f-colab',topN(F,r=>r.colab,r=>r.qtd||1),'am');

  /* ─ EPIs ─ */
  $('ce-cnt').textContent=fmt(E.length)+' registros';
  $('ek1').textContent=fmtR(gastoEPI);
  $('ek2').textContent=fmt(E.length);
  const etI=topN(E,r=>r.item,r=>r.qtd||1,1), etS=topN(E,r=>r.setor,r=>r.qtd||1,1);
  $('ek3').textContent=etI[0]?.[0]||'—'; $('ek3s').textContent=etI[0]?fmt(etI[0][1])+' unid.':'';
  $('ek4').textContent=etS[0]?.[0]||'—'; $('ek4s').textContent=etS[0]?fmt(etS[0][1])+' saídas':'';
  rProg('e-prog',topN(E,r=>r.item,r=>r.qtd||1));
  const eSet=topN(E,r=>r.setor,r=>r.qtd||1,7);
  rBar('ch-e-set','ce-e-set',eSet.map(e=>e[0]),eSet.map(e=>e[1]),'#057A55');
  const eUni=topN(E,r=>r.unidade,r=>r.qtd||1,6);
  rDonut('ch-e-uni','ce-e-uni',eUni.map(e=>e[0]),eUni.map(e=>e[1]));
  const byDay={};
  E.forEach(r=>{if(r.data&&r.preco){const k=r.data.toLocaleDateString('pt-BR');byDay[k]=(byDay[k]||0)+r.preco*r.qtd;}});
  const dks=Object.keys(byDay).sort((a,b)=>{const pa=a.split('/'),pb=b.split('/');return new Date(+pa[2],+pa[1]-1,+pa[0])-new Date(+pb[2],+pb[1]-1,+pb[0]);});
  rLine('ch-e-dia','ce-e-dia',dks,dks.map(k=>+byDay[k].toFixed(2)));
  rRank('e-colab',topN(E,r=>r.colab,r=>r.qtd||1),'gr');

  /* ─ Logística ─ */
  $('cl-cnt').textContent=fmt(L.length)+' registros';
  const ltU=topN(L,r=>r.unidade,()=>1,1), ltM=topN(L,r=>r.material,()=>1,1), ltT=topN(L,r=>r.tipo,()=>1,1);
  $('lk1').textContent=ltU[0]?.[0]||'N/D'; $('lk1s').textContent=(ltU[0]?fmt(ltU[0][1]):0)+' envios';
  $('lk2').textContent=fmt(L.length);
  $('lk3').textContent=ltM[0]?.[0]||'—'; $('lk3s').textContent=ltM[0]?fmt(ltM[0][1])+' envios':'';
  $('lk4').textContent=ltT[0]?.[0]||'—'; $('lk4s').textContent=ltT[0]?fmt(ltT[0][1])+' vezes':'';
  rRank('l-uni',topN(L,r=>r.unidade,()=>1),'pu');
  const lMat=topN(L,r=>r.material,()=>1,7);
  rBar('ch-l-mat','ce-l-mat',lMat.map(e=>e[0]),lMat.map(e=>e[1]),'#6427C2');
  const lTip=topN(L,r=>r.tipo,()=>1,6);
  rDonut('ch-l-tipo','ce-l-tipo',lTip.map(e=>e[0]),lTip.map(e=>e[1]));

  /* ─ Tabela ─ */
  buildFlat();
  renderTable();
}

/* ══ FLAT TABLE ═════════════════════════════════════════════════════ */
function buildFlat(){
  FLAT=[];
  FIL.mat.forEach(r=>FLAT.push({tipo:'Material',data:r.data,item:r.item,colab:r.colab,funcao:r.funcao,setor:r.setor,unidade:'—',qtd:r.qtd,preco:0,total:0}));
  FIL.fer.forEach(r=>FLAT.push({tipo:'Ferramenta',data:r.data,item:r.item,colab:r.colab,funcao:r.funcao,setor:r.setor,unidade:'—',qtd:r.qtd,preco:0,total:0}));
  FIL.epi.forEach(r=>FLAT.push({tipo:'EPI',data:r.data,item:r.item,colab:r.colab,funcao:r.funcao,setor:r.setor,unidade:r.unidade,qtd:r.qtd,preco:r.preco,total:r.preco*r.qtd}));
  FIL.log.forEach(r=>FLAT.push({tipo:'Logística',data:r.data,item:r.material,colab:r.tipo,funcao:'—',setor:'—',unidade:r.unidade,qtd:0,preco:0,total:0}));
  $('ct-cnt').textContent=fmt(FLAT.length)+(FLAT.length>500?' (500 exibidos)':'');
}

function renderTable(){
  const sorted=[...FLAT].sort((a,b)=>{
    let av=a[sortK],bv=b[sortK];
    if(av instanceof Date)av=av?.getTime()||0;
    if(bv instanceof Date)bv=bv?.getTime()||0;
    if(typeof av==='string')av=av.toLowerCase();
    if(typeof bv==='string')bv=bv.toLowerCase();
    return(av>bv?1:av<bv?-1:0)*sortD;
  });
  const rows=sorted.slice(0,500);
  if(!rows.length){$('tbody').innerHTML='<tr><td colspan="10" class="te">Nenhum registro.</td></tr>';return;}
  const bc={Material:'bm',Ferramenta:'bf',EPI:'be','Logística':'bl'};
  $('tbody').innerHTML=rows.map(r=>`<tr>
    <td><span class="badge ${bc[r.tipo]||'bm'}">${r.tipo}</span></td>
    <td style="font-family:'DM Mono',monospace;font-size:.72rem">${r.data?r.data.toLocaleDateString('pt-BR'):'—'}</td>
    <td title="${r.item}">${r.item||'—'}</td>
    <td>${r.colab||'—'}</td>
    <td style="color:var(--s500);font-size:.72rem">${r.funcao||'—'}</td>
    <td>${r.setor||'—'}</td>
    <td>${r.unidade||'—'}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace">${r.qtd||'—'}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace">${r.preco?fmtR(r.preco):'—'}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600">${r.total?fmtR(r.total):'—'}</td>
  </tr>`).join('');
}

function srt(k){if(sortK===k)sortD*=-1;else{sortK=k;sortD=1;}renderTable();}

/* ══ EXPORT CSV ═════════════════════════════════════════════════════ */
function exportCSV(){
  const hdr=['Tipo','Data','Item','Colaborador','Função','Setor','Unidade','Qtd','Preço Unit.','Total'];
  const lines=[hdr,...FLAT.map(r=>[
    r.tipo, r.data?r.data.toLocaleDateString('pt-BR'):'', r.item, r.colab, r.funcao,
    r.setor, r.unidade, r.qtd||'', r.preco?r.preco.toFixed(2).replace('.',','):'',
    r.total?r.total.toFixed(2).replace('.',','):''
  ])].map(r=>r.map(v=>`"${String(v||'').replace(/"/g,'""')}"`).join(';')).join('\n');
  const a=document.createElement('a');
  a.href='data:text/csv;charset=utf-8,\uFEFF'+encodeURIComponent(lines);
  a.download='painel_lotes_'+new Date().toISOString().slice(0,10)+'.csv';
  a.click();
  toast('CSV exportado.');
}

/* ══ CLEAR ══════════════════════════════════════════════════════════ */
function clearAll(){
  RAW={mat:[],fer:[],epi:[],log:[]};
  FIL={mat:[],fer:[],epi:[],log:[]};
  FLAT=[];
  populateDDs();
  renderAll();
  $('hcount').textContent='0';
  toast('Dados limpos.');
}

/* ══ TEMPLATE ═══════════════════════════════════════════════════════ */
function dlTemplate(){
  const wb=XLSX.utils.book_new();

  // Aba Dados
  const d=[
    ['DATA','MATERIAIS - ITEM','MATERIAIS - SAÍDAS','COLABORADOR MATERIAIS - NOME','FUNÇÃO','SETOR','DATA','FERRAMENTAS - FERRAMENTA','FERRAMENTAS - SAÍDAS','COLABORADOR FERRAMENTAS - NOME','FUNÇÃO','SETOR'],
    ['22/04/2026','TRAPO BRANCO',25,'ALEXANDRO','HIDROJATO','HIDROJATO','04/04/2026','CHAVE MANDRIL',1,'CARLOS CAETANO','PINTOR','PINTURA'],
    ['22/04/2026','LIXA DE 100',50,'ALEXANDRO','HIDROJATO','HIDROJATO','','','','','',''],
    ['22/04/2026','TRINCHA 1"',24,'CARLOS CAETANO','PINTOR','PINTURA','','','','','',''],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(d),'Dados');

  // Aba EPIs
  const e=[
    ['DATA','EPI - NOME','preço','EPI - SAÍDAS','SETOR','COLABORADOR','FUNÇÃO','UNIDADE'],
    ['04/05/2026','MACACÃO DE SEGURANÇA - RF',12.85,1,'PRODUÇÃO','SIDINEI SILVA MENEZES','SOLDADOR','-'],
    ['06/05/2026','MASCARA RESPIRATORIA PFF2 - 3M',14.85,5,'PRODUÇÃO','GENIVALDO SILVA DA CRUZ','OP.PONTE ROLANTE','-'],
    ['14/05/2026','OCULOS DE SEG. INCOLOR ESPELHADO',22.85,2,'OFFSHORE','-','-','POLVO A'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(e),'EPIs');

  // Aba Logística
  const l=[
    ['DATA','UNIDADE','MATERIAIS ENVIADOS','TIPO DE ENVIO'],
    ['04/05/2026','POLVO A','EPI','VOO'],
    ['05/05/2026','POLVO A','FERRAMENTA','VOO'],
    ['06/05/2026','ALPHA STAR','FERRAMENTA','BARCO'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(l),'Logística');

  XLSX.writeFile(wb,'template_painel_lotes.xlsx');
  toast('Template baixado!');
}

/* ══ COMPARTILHAR (URL com dados embutidos) ══════════════════════════
   Serializa RAW para JSON → comprime com LZ-string → codifica em Base64
   → coloca no hash da URL. Quem abre o link, decodifica e restaura os dados.
═══════════════════════════════════════════════════════════════════════ */

/* LZ-String minificado (compressão leve para URLs) */
var LZString=(function(){function o(r,o){if(!t[r]){t[r]={};for(var n=0;n<r.length;n++)t[r][r.charAt(n)]=n}return t[r][o]}var r=String.fromCharCode,n="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",e="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+-$",t={},i={compressToBase64:function(r){if(null==r)return"";var o=i._compress(r,6,function(r){return n.charAt(r)});switch(o.length%4){default:case 0:return o;case 1:return o+"===";case 2:return o+"==";case 3:return o+"="}},decompressFromBase64:function(r){return null==r?"":""==r?null:i._decompress(r.length,32,function(o){return o(n,r)})},compressToEncodedURIComponent:function(r){return null==r?"":i._compress(r,6,function(r){return e.charAt(r)})},decompressFromEncodedURIComponent:function(r){return null==r?"":""==r?null:(r=r.replace(/ /g,"+"),i._decompress(r.length,32,function(o){return o(e,r)}))},compress:function(r){return i._compress(r,16,function(r){return String.fromCharCode(r)})},decompress:function(r){return null==r?"":""==r?null:i._decompress(r.length,32768,function(o){return o(r,0)})},_compress:function(r,o,n){if(null==r)return"";var e,t,i,s={},p={},u="",a="",c="",l=2,f=3,h=2,d=[],m=0,v=0;for(i=0;i<r.length;i+=1){u=r.charAt(i);Object.prototype.hasOwnProperty.call(s,u)||(s[u]=f++,p[u]=!0);a=c+u;if(Object.prototype.hasOwnProperty.call(s,a))c=a;else{if(Object.prototype.hasOwnProperty.call(p,c)){if(c.charCodeAt(0)<256){for(e=0;e<h;e++)m=m<<1,v==o-1?(v=0,d.push(n(m)),m=0):v++;for(t=c.charCodeAt(0),e=0;e<8;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}else{for(t=1,e=0;e<h;e++)m=m<<1|t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t=0;for(t=c.charCodeAt(0),e=0;e<16;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}l--,0==l&&(l=Math.pow(2,h),h++);delete p[c]}else for(t=s[c],e=0;e<h;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1;l--,0==l&&(l=Math.pow(2,h),h++);s[a]=f++,c=String(u)}}if(""!==c){if(Object.prototype.hasOwnProperty.call(p,c)){if(c.charCodeAt(0)<256){for(e=0;e<h;e++)m=m<<1,v==o-1?(v=0,d.push(n(m)),m=0):v++;for(t=c.charCodeAt(0),e=0;e<8;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}else{for(t=1,e=0;e<h;e++)m=m<<1|t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t=0;for(t=c.charCodeAt(0),e=0;e<16;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1}l--,0==l&&(l=Math.pow(2,h),h++);delete p[c]}else for(t=s[c],e=0;e<h;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1;l--,0==l&&(l=Math.pow(2,h),h++)}for(t=2,e=0;e<h;e++)m=m<<1|1&t,v==o-1?(v=0,d.push(n(m)),m=0):v++,t>>=1;for(;;){if(m=m<<1,v==o-1){d.push(n(m));break}v++}return d.join("")},_decompress:function(r,o,n){var e,t,i,s,p,u,a,c=[],l=4,f=4,h=3,d="",m=[],v={val:n(0),position:o,index:1};for(t=0;t<3;t+=1)c[t]=t;for(i=0,p=Math.pow(2,2),u=1;u!=p;)s=v.val&v.position,v.position>>=1,0==v.position&&(v.position=o,v.val=n(v.index++)),i|=(s>0?1:0)*u,u<<=1;switch(e=i){case 0:for(i=0,p=Math.pow(2,8),u=1;u!=p;)s=v.val&v.position,v.position>>=1,0==v.position&&(v.position=o,v.val=n(v.index++)),i|=(s>0?1:0)*u,u<<=1;a=r(i);break;case 1:for(i=0,p=Math.pow(2,16),u=1;u!=p;)s=v.val&v.position,v.position>>=1,0==v.position&&(v.position=o,v.val=n(v.index++)),i|=(s>0?1:0)*u,u<<=1;a=r(i);break;case 2:return""}for(c[3]=a,d=a,m.push(a);;){if(v.index>r)return"";for(i=0,p=Math.pow(2,h),u=1;u!=p;)s=v.val&v.position,v.position>>=1,0==v.position&&(v.position=o,v.val=n(v.index++)),i|=(s>0?1:0)*u,u<<=1;switch(a=i){case 0:for(i=0,p=Math.pow(2,8),u=1;u!=p;)s=v.val&v.position,v.position>>=1,0==v.position&&(v.position=o,v.val=n(v.index++)),i|=(s>0?1:0)*u,u<<=1;c[f++]=r(i),a=f-1,l--;break;case 1:for(i=0,p=Math.pow(2,16),u=1;u!=p;)s=v.val&v.position,v.position>>=1,0==v.position&&(v.position=o,v.val=n(v.index++)),i|=(s>0?1:0)*u,u<<=1;c[f++]=r(i),a=f-1,l--;break;case 2:return m.join("")}if(0==l&&(l=Math.pow(2,h),h++),c[a])d=c[a];else{if(a!==f)return null;d=e+e.charAt(0)}m.push(d),c[f++]=e+d.charAt(0),l--,0==l&&(l=Math.pow(2,h),h++),e=d}return m.join("")}};return i})();

function rawToJSON(){
  // Serializa datas como string ISO
  const ser=(arr)=>arr.map(r=>{
    const o={...r};
    if(o.data instanceof Date && !isNaN(o.data)) o.data=o.data.toISOString();
    return o;
  });
  return JSON.stringify({mat:ser(RAW.mat),fer:ser(RAW.fer),epi:ser(RAW.epi),log:ser(RAW.log)});
}

function jsonToRaw(json){
  const deser=(arr)=>arr.map(r=>{
    const o={...r};
    if(o.data && typeof o.data==='string') o.data=new Date(o.data);
    return o;
  });
  const p=JSON.parse(json);
  return {mat:deser(p.mat||[]),fer:deser(p.fer||[]),log:deser(p.log||[]),epi:deser(p.epi||[])};
}

function gerarLink(){
  const total=RAW.mat.length+RAW.fer.length+RAW.epi.length+RAW.log.length;
  const status=$('share-status'), urlInput=$('share-url'), warn=$('share-warn');
  if(total===0){
    status.textContent='⚠️ Nenhum dado para compartilhar. Importe um XLSX primeiro.';
    urlInput.value='';
    return;
  }
  status.textContent='⏳ Comprimindo dados...';
  setTimeout(()=>{
    try{
      const json=rawToJSON();
      const compressed=LZString.compressToEncodedURIComponent(json);
      const base=window.location.origin+window.location.pathname;
      const url=base+'#d='+compressed;
      urlInput.value=url;
      const kb=Math.round(url.length/1024*10)/10;
      status.textContent=`✅ Link gerado! ${total} registros · ${kb} KB`;
      warn.style.display = url.length>8000 ? 'block' : 'none';
    }catch(e){
      status.textContent='❌ Erro ao gerar link.';
    }
  },50);
}

function openShare(){
  const ov=$('share-overlay');
  ov.style.display='flex';
  $('share-status').textContent='';
  $('share-url').value='';
  $('share-warn').style.display='none';
  gerarLink();
}

function closeShare(){ $('share-overlay').style.display='none'; }

function copyLink(){
  const url=$('share-url').value;
  if(!url)return;
  navigator.clipboard.writeText(url).then(()=>{
    const btn=$('copy-btn');
    btn.textContent='✓ Copiado!';
    btn.style.background='var(--green)';
    setTimeout(()=>{btn.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Copiar';btn.style.background='';},2500);
    toast('Link copiado para a área de transferência!');
  }).catch(()=>{
    $('share-url').select();
    document.execCommand('copy');
    toast('Link copiado!');
  });
}

function shareWA(){
  const url=$('share-url').value;
  if(!url){toast('Gere o link primeiro.',true);return;}
  const txt=encodeURIComponent('Painel de Controle de Lotes:\n'+url);
  window.open('https://wa.me/?text='+txt,'_blank');
}

/* ── Ao carregar: verifica se há dados no hash da URL ── */
function loadFromURL(){
  const hash = window.location.hash; // ex: #d=xyz...
  if(!hash || !hash.startsWith('#d=')) return;
  const compressed = hash.slice(3); // remove '#d='
  if(!compressed) return;
  try{
    const json=LZString.decompressFromEncodedURIComponent(compressed);
    if(!json)throw new Error('decompress failed');
    RAW=jsonToRaw(json);
    populateDDs();
    applyF();
    const total=RAW.mat.length+RAW.fer.length+RAW.epi.length+RAW.log.length;
    // limpa o hash da URL sem recarregar a página
    history.replaceState(null,'',window.location.pathname);
    toast(`✅ Dados carregados do link! (${total} registros)`);
  }catch(e){
    toast('Não foi possível carregar os dados do link.',true);
  }
}
// Roda depois que tudo estiver inicializado
window.addEventListener('DOMContentLoaded', loadFromURL);

/* ══ DEMO DATA ══════════════════════════════════════════════════════ */
function loadDemo(){
  const itMat=['TRAPO BRANCO','LIXA DE 100','TRINCHA 1"','ROLO DE ESPUMA 5CM','VARETA ER70S-3 3,2MM','ARREBITE 3,2MM','DISCO DE CORTE 4.5"','PANO DE LIMPEZA','FITA ADESIVA','TINTA EPÓXI'];
  const itFer=['CHAVE MANDRIL','ALICATE DE BICO','FURADEIRA 3/8"','PARAFUSADEIRA','ESMERILHADEIRA','TORQUÍMETRO','MULTÍMETRO','NÍVEL A LASER'];
  const itEpi=['MACACÃO DE SEGURANÇA - RF','MASCARA RESPIRATORIA PFF2 - 3M','LUVA DE SEG. PIGMENTADA - KALIPSO','OCULOS DE SEG. INCOLOR - CARBOGRAFITE','PROTETOR AURICULAR TIPO PLUG - 3M','LUVA DE SEG. ANTI IMPACTO - VOLK','CAPACETE DE SEGURANÇA ABA FRONTAL'];
  const colM=['ALEXANDRO','CARLOS CAETANO','VANILDO','LUIZ CARLOS','SIDINEI','GENIVALDO'];
  const colF=['CARLOS CAETANO','ALEXANDRO','JONAS','LUIZ CARLOS'];
  const colE=['SIDINEI SILVA MENEZES','GENIVALDO SILVA DA CRUZ','JONAS GUIMARÃES CAMPOS','ALEXANDRO','LUIZ CARLOS'];
  const setM=['HIDROJATO','PINTURA','PRODUÇÃO','CALDEIRARIA'];
  const setE=['PRODUÇÃO','HIDROJATO','OFFSHORE'];
  const unis=['POLVO A','ALPHA STAR','FRADE','HUNTER'];
  const precos=[12.85,14.85,15.85,22.85,21.85,19.85,29.90];

  RAW.mat=[];RAW.fer=[];RAW.epi=[];RAW.log=[];

  for(let i=0;i<30;i++){
    const d=new Date(2026,3,Math.floor(Math.random()*22)+1);
    RAW.mat.push({data:d,item:itMat[i%itMat.length],qtd:Math.floor(Math.random()*100)+5,colab:colM[i%colM.length],funcao:'OPERADOR',setor:setM[i%setM.length]});
  }
  for(let i=0;i<12;i++){
    const d=new Date(2026,3,Math.floor(Math.random()*22)+1);
    RAW.fer.push({data:d,item:itFer[i%itFer.length],qtd:Math.floor(Math.random()*5)+1,colab:colF[i%colF.length],funcao:'TÉCNICO',setor:setM[i%setM.length]});
  }
  for(let i=0;i<17;i++){
    const idx=i%itEpi.length, d=new Date(2026,4,i+4);
    RAW.epi.push({data:d,item:itEpi[idx],preco:precos[idx%precos.length],qtd:Math.floor(Math.random()*8)+1,setor:setE[i%setE.length],colab:colE[i%colE.length],funcao:'OPERADOR',unidade:i<10?'-':unis[i%unis.length]});
  }
  const mEnv=['EPI','FERRAMENTA','CONSUMO','OBRA'];
  const tipoEnv=['VOO','BARCO'];
  for(let i=0;i<7;i++){
    RAW.log.push({data:new Date(2026,4,i+4),unidade:unis[i%unis.length],material:mEnv[i%mEnv.length],tipo:tipoEnv[i%tipoEnv.length]});
  }

  populateDDs();applyF();
  toast(`✓ Demo: ${RAW.mat.length} mat · ${RAW.fer.length} fer · ${RAW.epi.length} EPIs · ${RAW.log.length} log`);
}
</script>
</body>
</html>
