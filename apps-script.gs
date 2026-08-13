/**
 * Dashboard do Funil do Quiz — Ciclo Fértil
 * Cole este código inteiro no editor do Apps Script (Code.gs).
 * O quiz continua enviando os dados igual (trackEtapa / trackCheckout / enviarRespostas).
 */

const SHEET_NAME  = 'Eventos';   // aba onde os eventos são gravados
const PITCH_START = 30;          // pitch começa na tela 30
const BODY_INI    = 3;           // corpo monitorado: telas 3..29
const BODY_FIM    = 29;
const THRESH      = 5;           // alerta se a queda passar de 5 pontos
const TOTAL_TELAS = 34;
const TZ          = 'America/Sao_Paulo';

const NOMES = [
  '(abertura/idade)',
  '(prova social)',
  'Seu ciclo menstrual é…',
  'Qual é o seu maior desejo nesse momento?',
  'Em qual área você sente algo travado?',
  'Há quanto tempo está tentando engravidar?',
  'Você nota mudanças no muco ao longo do mês?',
  'Você já notou muco tipo clara de ovo?',
  'Usa sabonete íntimo/ducha/lenços com frequência?',
  'Já mediu a temperatura ao acordar?',
  'Você sabe quando está ovulando?',
  'Sente dor de um lado do ventre no meio do ciclo?',
  'Como você identifica seu período fértil?',
  'Como descreveria seu estresse hoje?',
  'Tem um momento do dia só pra você?',
  'Fica ansiosa quando a menstruação vai chegar?',
  'Intensidade da pressão de engravidar?',
  'Seu humor muda bastante no dia?',
  'Como descreveria seu sono?',
  'Frequência de atividade física?',
  'Como descreveria sua alimentação?',
  'Sente inchaço/cansaço/dores sem causa?',
  'Ainda acredita que vai conseguir engravidar?',
  'Tem prática de meditação/oração/respiração?',
  'Consegue ter momentos de alegria e leveza?',
  'Você se culpa quando o ciclo não vem?',
  'Quantos dias dura seu ciclo?',
  'Quantos dias dura sua menstruação?',
  'Em quantos dias seu desejo sexual aumenta?',
  '(RESULTADO diagnóstico — início do pitch)',
  'Quando começou sua última menstruação?',
  '(loading)',
  'gráfico fascinations',
  'Página de vendas'
];

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  if (p.etapa || p.action === 'respostas') {
    ingest_(p);
    return ContentService.createTextOutput('ok');
  }
  if (p.aba === 'criativos') {
    return renderCriativos_(resolveRange_(p));
  }
  return renderDashboard_(resolveRange_(p));
}

function ingest_(p) {
  const sh = getSheet_();
  const tipo = (p.action === 'respostas') ? 'respostas' : 'etapa';
  sh.appendRow([new Date(), String(p.sessao || ''), tipo, String(p.etapa || ''), String(p.data || '')]);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['ts', 'sessao', 'tipo', 'etapa', 'data']);
  }
  return sh;
}

function zerarDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (sh) { ss.deleteSheet(sh); }
  getSheet_();
  return true;
}

/* ---------------- FILTRO POR DATA ---------------- */
function fmt_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function mkDate_(key) { const p = String(key).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function today_() { return fmt_(new Date()); }
function addDays_(key, n) { const d = mkDate_(key); d.setDate(d.getDate() + n); return fmt_(d); }
function weekStart_(key) { const d = mkDate_(key); const dow = (d.getDay() + 6) % 7; d.setDate(d.getDate() - dow); return fmt_(d); }
function dKey_(v) { const d = (v instanceof Date) ? v : new Date(v); return isNaN(d) ? '' : Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function dBR_(k) { const p = String(k).split('-'); return p.length === 3 ? (p[2] + '/' + p[1] + '/' + p[0]) : k; }
function R_(from, to, label, preset) { return { from: from, to: to, label: label, preset: preset }; }

function resolveRange_(p) {
  const t = today_();
  const preset = p.preset || ((p.from || p.to) ? 'custom' : 'todos');
  if (preset === 'hoje')  return R_(t, t, 'Hoje', 'hoje');
  if (preset === 'ontem') { const y = addDays_(t, -1); return R_(y, y, 'Ontem', 'ontem'); }
  if (preset === '7d')    return R_(addDays_(t, -6),  t, 'Últimos 7 dias',  '7d');
  if (preset === '14d')   return R_(addDays_(t, -13), t, 'Últimos 14 dias', '14d');
  if (preset === '28d')   return R_(addDays_(t, -27), t, 'Últimos 28 dias', '28d');
  if (preset === '30d')   return R_(addDays_(t, -29), t, 'Últimos 30 dias', '30d');
  if (preset === 'semana') { const ws = weekStart_(t); return R_(ws, t, 'Esta semana', 'semana'); }
  if (preset === 'semana_passada') { const ws = weekStart_(t); return R_(addDays_(ws, -7), addDays_(ws, -1), 'Semana passada', 'semana_passada'); }
  if (preset === 'mes')   return R_(t.slice(0, 8) + '01', t, 'Este mês', 'mes');
  if (preset === 'mes_passado') { const d = mkDate_(t.slice(0, 8) + '01'); const end = fmt_(new Date(d.getFullYear(), d.getMonth(), 0)); return R_(end.slice(0, 8) + '01', end, 'Mês passado', 'mes_passado'); }
  if (preset === 'custom') { const f = p.from || '', to2 = p.to || f; return R_(f, to2, (f === to2 ? dBR_(f) : (dBR_(f) + ' – ' + dBR_(to2))), 'custom'); }
  return R_('', '', 'Máximo (todos)', 'todos');
}

function computeFunnel_(from, to) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const sess = {};
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, 5).getValues();
    rows.forEach(function (r) {
      const id = String(r[1] || '');
      if (!id) return;
      if (from || to) {
        const k = dKey_(r[0]);
        if (!k) return;
        if (from && k < from) return;
        if (to && k > to) return;
      }
      const et = parseInt(r[3], 10);
      if (!sess[id]) sess[id] = { max: 0, checkout: false };
      if (!isNaN(et)) {
        if (et === 35) sess[id].checkout = true;
        else if (et >= 1 && et <= TOTAL_TELAS && et > sess[id].max) sess[id].max = et;
      }
    });
  }
  const reached = new Array(TOTAL_TELAS + 1).fill(0);
  let visitantes = 0, compraram = 0, fim = 0;
  Object.keys(sess).forEach(function (id) {
    const s = sess[id];
    if (s.max >= 1) visitantes++;
    if (s.checkout) compraram++;
    if (s.max >= TOTAL_TELAS) fim++;
    for (let n = 1; n <= TOTAL_TELAS; n++) if (s.max >= n) reached[n]++;
  });
  const base = reached[1] || 1;
  const etapas = [];
  let maiorQueda = { tela: 0, drop: 0 };
  for (let n = 1; n <= TOTAL_TELAS; n++) {
    const ret = Math.round((reached[n] / base) * 1000) / 10;
    const retPrev = n > 1 ? Math.round((reached[n - 1] / base) * 1000) / 10 : 100;
    const drop = n > 1 ? Math.round((retPrev - ret) * 10) / 10 : 0;
    const isBody = n >= BODY_INI && n <= BODY_FIM;
    const alerta = isBody && drop > THRESH;
    if (isBody && drop > maiorQueda.drop) maiorQueda = { tela: n, drop: drop };
    etapas.push({ n: n, nome: NOMES[n - 1] || ('Tela ' + n), acessos: reached[n], ret: ret, drop: drop, alerta: alerta, pitch: n >= PITCH_START });
  }
  return {
    etapas: etapas, visitantes: visitantes, compraram: compraram, fim: fim,
    pctFim: visitantes ? Math.round((fim / visitantes) * 1000) / 10 : 0,
    pctCompra: visitantes ? Math.round((compraram / visitantes) * 1000) / 10 : 0,
    maiorQueda: maiorQueda,
    alertas: etapas.filter(function (e) { return e.alerta; }).map(function (e) { return e.n; })
  };
}

function renderDashboard_(range) {
  const d = computeFunnel_(range.from, range.to);
  return HtmlService.createHtmlOutput(buildHtml_(d, range))
    .setTitle('Funil do Quiz — Ciclo Fértil')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function buildHtml_(d, range) {
  const dataJson = JSON.stringify(d);
  const URL = ScriptApp.getService().getUrl();
  const T = today_();
  const on = function (k) { return range.preset === k ? ' active' : ''; };
  const curFrom = range.preset === 'custom' ? range.from : '';
  const curTo = range.preset === 'custom' ? range.to : '';
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;}
  body{background:#f4f5f4;color:#1a1a1a;padding:24px;}
  .wrap{max-width:1100px;margin:0 auto;}
  h1{font-size:22px;font-weight:600;}
  .sub{color:#7a7a7a;font-size:13px;margin-bottom:16px;}
  .btn{float:right;border:1px solid #d9534f;color:#d9534f;background:#fff;border-radius:8px;padding:8px 14px;font-size:13px;cursor:pointer;}
  .tabs{display:flex;gap:6px;margin-bottom:16px;}
  .tabs a{padding:8px 18px;border-radius:10px;font-size:13.5px;font-weight:600;text-decoration:none;color:#666;background:#fff;border:1px solid #e2e2e2;}
  .tabs a.active{background:#1a7a4f;color:#fff;border-color:#1a7a4f;}
  .filterbar{position:relative;margin-bottom:20px;}
  .rangeBtn{border:1px solid #d5d5d5;background:#fff;border-radius:10px;padding:10px 16px;font-size:13.5px;font-weight:600;color:#333;cursor:pointer;display:inline-flex;align-items:center;gap:8px;}
  .rangeBtn .car{color:#999;font-size:11px;}
  .rangePanel{position:absolute;top:48px;left:0;z-index:60;background:#fff;border:1px solid #e2e2e2;border-radius:14px;box-shadow:0 10px 34px rgba(0,0,0,.14);display:flex;overflow:hidden;}
  .presets{width:196px;border-right:1px solid #eee;padding:8px;}
  .presets a{display:block;padding:9px 12px;border-radius:8px;font-size:13px;color:#444;text-decoration:none;cursor:pointer;}
  .presets a:hover{background:#f1f4f2;}
  .presets a.active{background:#e9f3ee;color:#2f6b4f;font-weight:600;}
  .presets .div{height:1px;background:#eee;margin:6px 8px;}
  .calarea{padding:14px 16px;width:278px;}
  .calnav{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;font-weight:600;font-size:14px;color:#333;}
  .calnav button{border:none;background:none;font-size:20px;line-height:1;cursor:pointer;color:#666;padding:2px 10px;border-radius:6px;}
  .calnav button:hover{background:#f0f0f0;}
  .calgrid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}
  .caldow{font-size:10.5px;color:#aaa;text-align:center;padding:4px 0;text-transform:uppercase;}
  .calday{text-align:center;font-size:12.5px;padding:8px 0;border-radius:7px;cursor:pointer;color:#333;}
  .calday:hover{background:#eef2f0;}
  .calday.muted{visibility:hidden;}
  .calday.inrange{background:#e9f3ee;border-radius:0;}
  .calday.edge{background:#1a7a4f;color:#fff;border-radius:7px;}
  .calfoot{display:flex;align-items:center;gap:10px;margin-top:14px;}
  .calsel{flex:1;font-size:12px;color:#666;}
  .cbtn{border:1px solid #dcdcdc;background:#fff;border-radius:8px;padding:8px 16px;font-size:12.5px;cursor:pointer;}
  .cbtn.primary{background:#1a7a4f;color:#fff;border-color:#1a7a4f;font-weight:600;}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:18px;}
  .kpi{background:#fff;border-radius:12px;padding:16px 18px;}
  .kpi .lbl{font-size:12px;color:#8a8a8a;text-transform:uppercase;letter-spacing:.04em;}
  .kpi .val{font-size:30px;font-weight:600;margin-top:4px;}
  .card{background:#fff;border-radius:14px;padding:20px;margin-bottom:18px;}
  .alertbox{background:#fdecec;color:#b3261e;border-radius:12px;padding:12px 16px;margin-bottom:18px;font-size:13.5px;}
  .curveCard{background:#0d1512;border-radius:14px;padding:18px 16px 10px;margin-bottom:18px;}
  .curveHead{display:flex;justify-content:space-between;color:#e8f3ec;font-size:14px;font-weight:500;margin-bottom:8px;}
  .row{display:flex;align-items:center;gap:10px;padding:5px 8px;border-radius:8px;}
  .row.alert{background:#fdecec;}
  .row.pitch{opacity:.55;}
  .row .num{width:24px;text-align:right;font-size:12px;color:#9a9a9a;}
  .row .nome{flex:0 0 260px;font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .row .bar{flex:1;height:16px;background:#eee;border-radius:4px;overflow:hidden;}
  .row .bar > i{display:block;height:100%;}
  .row .pct{width:52px;text-align:right;font-size:12px;color:#555;}
  .row .drop{width:90px;text-align:right;font-size:12px;}
  .drop.on{color:#d9534f;font-weight:600;}
  .drop.off{color:#b0b0b0;}
  .sec-title{font-size:15px;font-weight:600;margin-bottom:10px;}
  .legend{font-size:12px;color:#8a8a8a;margin-bottom:10px;}
  @media(max-width:560px){.rangePanel{flex-direction:column;}.presets{width:100%;border-right:none;border-bottom:1px solid #eee;}.calarea{width:100%;}}
</style></head><body>
<div class="wrap">
  <button class="btn" onclick="if(confirm('Zerar TODOS os dados do funil?')){google.script.run.withSuccessHandler(function(){location.reload();}).zerarDados();}">Zerar dados</button>
  <h1>Funil do Quiz — Ciclo Fértil</h1>
  <div class="sub">Atualizado em tempo real via Google Sheets · funil por sessão (sempre decrescente)</div>
  <div class="tabs">
    <a class="active" href="${URL}" target="_top">Funil</a>
    <a href="${URL}?aba=criativos" target="_top">Criativos</a>
  </div>
  <div class="filterbar">
    <button class="rangeBtn" onclick="var p=document.getElementById('rangePanel');p.style.display=(p.style.display==='none'||!p.style.display)?'flex':'none';">
      <span>📅</span><span>${range.label}</span><span class="car">▼</span>
    </button>
    <div id="rangePanel" class="rangePanel" style="display:none;">
      <div class="presets">
        <a class="${on('hoje')}"  href="${URL}?preset=hoje"  target="_top">Hoje</a>
        <a class="${on('ontem')}" href="${URL}?preset=ontem" target="_top">Ontem</a>
        <a class="${on('7d')}"    href="${URL}?preset=7d"    target="_top">Últimos 7 dias</a>
        <a class="${on('14d')}"   href="${URL}?preset=14d"   target="_top">Últimos 14 dias</a>
        <a class="${on('28d')}"   href="${URL}?preset=28d"   target="_top">Últimos 28 dias</a>
        <a class="${on('30d')}"   href="${URL}?preset=30d"   target="_top">Últimos 30 dias</a>
        <div class="div"></div>
        <a class="${on('semana')}"         href="${URL}?preset=semana"         target="_top">Esta semana</a>
        <a class="${on('semana_passada')}" href="${URL}?preset=semana_passada" target="_top">Semana passada</a>
        <a class="${on('mes')}"            href="${URL}?preset=mes"            target="_top">Este mês</a>
        <a class="${on('mes_passado')}"    href="${URL}?preset=mes_passado"    target="_top">Mês passado</a>
        <div class="div"></div>
        <a class="${on('todos')}" href="${URL}?preset=todos" target="_top">Máximo (todos)</a>
      </div>
      <div class="calarea">
        <div class="calnav"><button onclick="calMove(-1)">‹</button><span id="calTitle"></span><button onclick="calMove(1)">›</button></div>
        <div id="calGrid" class="calgrid"></div>
        <div class="calfoot">
          <span id="calSel" class="calsel">Selecione as datas</span>
          <button class="cbtn primary" onclick="calApply()">Aplicar</button>
        </div>
      </div>
    </div>
  </div>
  <div class="kpis">
    <div class="kpi"><div class="lbl">Visitantes</div><div class="val">${d.visitantes}</div></div>
    <div class="kpi"><div class="lbl">Chegaram ao fim</div><div class="val">${d.pctFim}%</div></div>
    <div class="kpi"><div class="lbl">Clicaram em comprar</div><div class="val">${d.pctCompra}%</div></div>
    <div class="kpi"><div class="lbl">Maior queda no corpo</div><div class="val">${d.maiorQueda.tela ? ('Tela ' + d.maiorQueda.tela) : '—'}</div></div>
  </div>
  <div id="alertBox" class="alertbox" style="display:none;"></div>
  <div class="curveCard">
    <div class="curveHead"><span>Retenção do funil</span><span style="color:#5f8a72;">${TOTAL_TELAS} telas</span></div>
    <div style="position:relative;width:100%;height:240px;"><canvas id="ret"></canvas></div>
  </div>
  <div class="card">
    <div class="sec-title">Etapas do Quiz</div>
    <div class="legend">Alerta (vermelho) = queda &gt; ${THRESH} pts numa tela do corpo (telas ${BODY_INI}–${BODY_FIM}). Telas 1–2 e pitch (${PITCH_START}+) são ignoradas.</div>
    <div id="list"></div>
  </div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>
  var URL = "${URL}", TODAY = "${T}";
  var selFrom = "${curFrom}", selTo = "${curTo}";
  var MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  var vd; (function(){ var base = selFrom || TODAY; var p = base.split('-'); vd = new Date(+p[0], +p[1]-1, 1); })();
  function pad2(n){ return (n<10?'0':'')+n; }
  function keyOf(y,m,d){ return y+'-'+pad2(m+1)+'-'+pad2(d); }
  function brdate(k){ var p=k.split('-'); return p[2]+'/'+p[1]+'/'+p[0]; }
  function renderCal(){
    var y=vd.getFullYear(), m=vd.getMonth();
    document.getElementById('calTitle').textContent = MESES[m]+' '+y;
    var startDow=(new Date(y,m,1).getDay()+6)%7, dias=new Date(y,m+1,0).getDate();
    var html='';
    ['seg','ter','qua','qui','sex','sáb','dom'].forEach(function(d){ html+='<div class="caldow">'+d+'</div>'; });
    for(var i=0;i<startDow;i++) html+='<div class="calday muted"></div>';
    for(var d=1;d<=dias;d++){
      var k=keyOf(y,m,d), cls='calday';
      if(selFrom && selTo && k>=selFrom && k<=selTo) cls+=' inrange';
      if(k===selFrom || k===selTo) cls+=' edge';
      html+='<div class="'+cls+'" data-k="'+k+'">'+d+'</div>';
    }
    document.getElementById('calGrid').innerHTML=html;
    var sel=document.getElementById('calSel');
    if(selFrom && selTo) sel.textContent=brdate(selFrom)+' – '+brdate(selTo);
    else if(selFrom) sel.textContent=brdate(selFrom)+' – ...';
    else sel.textContent='Selecione as datas';
  }
  function pickDay(k){
    if(!selFrom || (selFrom && selTo)){ selFrom=k; selTo=''; }
    else if(k>=selFrom){ selTo=k; }
    else { selTo=selFrom; selFrom=k; }
    renderCal();
  }
  function calMove(n){ vd=new Date(vd.getFullYear(), vd.getMonth()+n, 1); renderCal(); }
  function calApply(){ if(!selFrom){ alert('Selecione uma data no calendário'); return; } window.top.location.href = URL+'?from='+selFrom+'&to='+(selTo||selFrom); }
  document.getElementById('calGrid').addEventListener('click', function(e){ var k=e.target.getAttribute('data-k'); if(k) pickDay(k); });
  renderCal();

  var D = ${dataJson};
  var list = document.getElementById('list');
  D.etapas.forEach(function(e){
    var barColor = e.alerta ? '#d84a3a' : (e.pitch ? '#9a9a9a' : '#6b9e7e');
    var dropHtml = e.n>1 ? (e.alerta
        ? '<span class="drop on">&#9660; '+e.drop+' pts</span>'
        : '<span class="drop off">'+(e.drop>0?('-'+e.drop+' pts'):'')+'</span>') : '';
    var div = document.createElement('div');
    div.className = 'row' + (e.alerta?' alert':'') + (e.pitch?' pitch':'');
    div.innerHTML = '<span class="num">'+e.n+'</span><span class="nome">'+e.nome+'</span>'+
      '<span class="bar"><i style="width:'+e.ret+'%;background:'+barColor+';"></i></span>'+
      '<span class="pct">'+e.ret+'%</span>'+dropHtml;
    list.appendChild(div);
  });
  if (D.alertas.length){
    var ab = document.getElementById('alertBox'); ab.style.display='block';
    ab.innerHTML = '&#9888; '+D.alertas.length+' tela(s) do corpo com queda acima de ${THRESH} pts: <b>tela '+D.alertas.join(', ')+'</b>. Vale revisar a pergunta.';
  }
  var labels=D.etapas.map(function(e){return e.n;}), data=D.etapas.map(function(e){return e.ret;}), nomes=D.etapas.map(function(e){return e.nome;});
  var ctx=document.getElementById('ret').getContext('2d');
  var g=ctx.createLinearGradient(0,0,0,240); g.addColorStop(0,'rgba(53,208,127,0.45)'); g.addColorStop(1,'rgba(53,208,127,0.02)');
  new Chart(ctx,{type:'line',data:{labels:labels,datasets:[{data:data,borderColor:'#35d07f',borderWidth:2,backgroundColor:g,fill:true,tension:0.4,pointRadius:0,pointHoverRadius:5,pointHoverBackgroundColor:'#35d07f'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{title:function(t){var i=t[0].dataIndex;return 'Tela '+labels[i]+' — '+nomes[i];},label:function(c){return c.parsed.y+'% ainda no funil';}}}},
      scales:{y:{min:0,max:100,ticks:{color:'#6b8f7b',callback:function(v){return v+'%';}},grid:{color:'rgba(255,255,255,0.06)'}},x:{ticks:{color:'#6b8f7b',maxTicksLimit:10},grid:{display:false}}}}});
</script>
</body></html>`;
}

/* ════════════════════════════════════════════════════
   ABA CRIATIVOS — controle de criativos do Meta
   Planilha 'Criativos': 1 linha por criativo.
   Etiquetas (formato, angulo, hook, emocao, estrutura) você edita direto na planilha.
   Métricas: atualizadas via MCP do Meta (Claude) de tempos em tempos.
   ════════════════════════════════════════════════════ */
const CRIATIVOS_SHEET = 'Criativos';

const CRIATIVOS_HEADERS = ['ad_id','anuncio','campanha','status','formato','angulo','hook','emocao','estrutura','arquetipo','segmentacao','amplificador','prova','gasto','impressoes','ctr_pct','hook_rate_pct','hold_rate_pct','views25_pct','lpv','custo_lpv','checkouts','custo_checkout','vendas','custo_venda','faturamento','roas','roi_pct','atualizado'];

// Faturamento vem do valor de compra que o Meta reporta (já inclui order bump).
// PRECO_VENDA só é usado como reserva quando a API não devolve valor.
const PRECO_VENDA   = 37.90;  // preço do produto principal
const PRECO_BUMP    = 17.00;  // preço do order bump
const LIQUIDO_VENDA = 33.50;  // o que cai na conta por venda do principal
const LIQUIDO_BUMP  = 14.48;  // o que cai na conta por order bump

/* Quantos bumps teve dá pra deduzir do valor total, porque os dois preços são fixos:
   valor = vendas × 37,90 + bumps × 17,00  →  bumps = (valor − vendas × 37,90) ÷ 17,00 */
function liquido_(faturamento, vendas) {
  const f = Number(faturamento) || 0;
  const v = Number(vendas) || 0;
  if (!v) return 0;
  const bumps = Math.max(0, Math.round((f - v * PRECO_VENDA) / PRECO_BUMP));
  return v * LIQUIDO_VENDA + bumps * LIQUIDO_BUMP;
}

// aba só do robô: uma linha por anúncio por dia (base do filtro de data)
const CRIATIVOS_DIARIO = 'Criativos_Diario';
const DIARIO_HEADERS = ['data','ad_id','anuncio','campanha','gasto','impressoes','clicks','v3s','v2s','thruplay','p25','lpv','checkouts','vendas','valor'];

// 💎 destaque de criativo com métrica excelente (ajuste os limites aqui)
// Regra: precisa das TRÊS ao mesmo tempo.
const DIAMANTE_CTR_MIN = 5;      // CTR em %
const DIAMANTE_CPC_MAX = 1.0;    // CPC em R$
const DIAMANTE_GASTO_MIN = 100;  // gasto mínimo em R$ pra valer o selo

function ehDiamante_(r) {
  const ctr = Number(r.ctr_pct) || 0;
  const impr = Number(r.impressoes) || 0;
  const gasto = Number(r.gasto) || 0;
  const cliques = impr * ctr / 100;
  const cpc = cliques ? gasto / cliques : null;
  return gasto >= DIAMANTE_GASTO_MIN && ctr >= DIAMANTE_CTR_MIN && cpc !== null && cpc <= DIAMANTE_CPC_MAX;
}

function getCriativosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CRIATIVOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CRIATIVOS_SHEET);
    sh.appendRow(CRIATIVOS_HEADERS);
    return sh;
  }
  // migração: planilha antiga sem as colunas novas (arquetipo..prova depois de estrutura)
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (h.indexOf('arquetipo') === -1) {
    const posEstrutura = h.indexOf('estrutura') + 1; // 1-based
    sh.insertColumnsAfter(posEstrutura, 4);
    sh.getRange(1, posEstrutura + 1, 1, 4).setValues([['arquetipo','segmentacao','amplificador','prova']]);
  }
  // migração: faturamento/roas/roi entram depois de custo_venda
  const h2 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (h2.indexOf('faturamento') === -1) {
    const posCustoVenda = h2.indexOf('custo_venda') + 1; // 1-based
    sh.insertColumnsAfter(posCustoVenda, 3);
    sh.getRange(1, posCustoVenda + 1, 1, 3).setValues([['faturamento','roas','roi_pct']]);
  }
  return sh;
}

function lerCriativos_() {
  const sh = getCriativosSheet_();
  const vals = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[1]) continue;
    const o = {};
    CRIATIVOS_HEADERS.forEach(function (h, j) { o[h] = r[j]; });
    rows.push(o);
  }
  return rows;
}

function agruparPor_(rows, campo) {
  const g = {};
  rows.forEach(function (r) {
    const k = String(r[campo] || '(sem etiqueta)');
    if (!g[k]) g[k] = { n: 0, gasto: 0, checkouts: 0, vendas: 0, faturamento: 0 };
    g[k].n++;
    g[k].gasto += Number(r.gasto) || 0;
    g[k].checkouts += Number(r.checkouts) || 0;
    g[k].vendas += Number(r.vendas) || 0;
    g[k].faturamento += Number(r.faturamento) || 0;
  });
  return Object.keys(g).map(function (k) {
    const x = g[k];
    return {
      nome: k, n: x.n, gasto: x.gasto, checkouts: x.checkouts, vendas: x.vendas, faturamento: x.faturamento,
      custoCheckout: x.checkouts ? x.gasto / x.checkouts : null,
      custoVenda: x.vendas ? x.gasto / x.vendas : null
    };
  }).sort(function (a, b) { return (a.custoVenda || 1e9) - (b.custoVenda || 1e9); });
}

function moeda_(v) { return (v === null || v === '' || isNaN(v)) ? '—' : 'R$' + Number(v).toFixed(2).replace('.', ','); }

function hookCurto_(s) {
  s = String(s || '');
  if (!s) return '—';
  const corte = s.search(/[.?!,]/);
  const lim = corte > 0 ? Math.min(corte, 45) : 45;
  if (s.length <= lim) return s;
  let t = s.slice(0, lim);
  const sp = t.lastIndexOf(' ');
  if (sp > 20) t = t.slice(0, sp);
  return t + '…';
}

/* Recalcula as métricas de cada linha a partir da aba diária, dentro do período.
   Mantém todas as etiquetas. Linhas sem entrega no período ficam zeradas. */
function aplicarPeriodo_(rows, range) {
  const ag = lerDiarioAgregado_(range.from, range.to);
  if (!ag) return rows; // aba diária ainda não existe: usa o que está gravado
  return rows.map(function (r) {
    const a = ag[String(r.anuncio)];
    const o = {};
    Object.keys(r).forEach(function (k) { o[k] = r[k]; });
    if (!a) {
      ['gasto','impressoes','ctr_pct','hook_rate_pct','hold_rate_pct','views25_pct','lpv','custo_lpv','checkouts','custo_checkout','vendas','custo_venda','faturamento','roas','roi_pct']
        .forEach(function (k) { o[k] = ''; });
      o.gasto = 0; o.impressoes = 0; o.checkouts = 0; o.vendas = 0; o.faturamento = 0;
      return o;
    }
    const base3s = a.v3s || a.v2s;
    const fat = a.valor > 0 ? a.valor : a.compras * PRECO_VENDA;
    const liq = liquido_(fat, a.compras);
    o.gasto = round2_(a.spend);
    o.impressoes = a.impr;
    o.ctr_pct = a.impr ? Math.round(10000 * a.clicks / a.impr) / 100 : '';
    o.hook_rate_pct = (a.impr && base3s) ? round1_(100 * base3s / a.impr) : '';
    o.hold_rate_pct = base3s ? round1_(100 * a.thruplay / base3s) : '';
    o.views25_pct = a.impr ? round1_(100 * a.p25 / a.impr) : '';
    o.lpv = a.lpv || '';
    o.custo_lpv = a.lpv ? round2_(a.spend / a.lpv) : '';
    o.checkouts = a.ic || 0;
    o.custo_checkout = a.ic ? round2_(a.spend / a.ic) : '';
    o.vendas = a.compras || 0;
    o.custo_venda = a.compras ? round2_(a.spend / a.compras) : '';
    o.faturamento = round2_(fat);
    o.roas = a.spend ? Math.round(100 * fat / a.spend) / 100 : '';
    o.roi_pct = a.spend ? round1_(100 * (liq - a.spend) / a.spend) : '';
    return o;
  });
}

function renderCriativos_(range) {
  range = range || R_('', '', 'Máximo (todos)', 'todos');
  const rows = aplicarPeriodo_(lerCriativos_(), range);
  const rowsJson = JSON.stringify(rows);
  const URL = ScriptApp.getService().getUrl();
  const on = function (k) { return range.preset === k ? ' active' : ''; };
  const T = today_();
  const curFrom = range.preset === 'custom' ? range.from : '';
  const curTo = range.preset === 'custom' ? range.to : '';
  const totGasto = rows.reduce(function (s, r) { return s + (Number(r.gasto) || 0); }, 0);
  const totVendas = rows.reduce(function (s, r) { return s + (Number(r.vendas) || 0); }, 0);
  const totCheckouts = rows.reduce(function (s, r) { return s + (Number(r.checkouts) || 0); }, 0);
  const testados = rows.filter(function (r) { return String(r.anuncio || '').trim() !== ''; }).length;
  const totFaturamento = rows.reduce(function (s, r) { return s + (Number(r.faturamento) || 0); }, 0);
  const totLiquido = liquido_(totFaturamento, totVendas);
  const roasGeral = totGasto ? totFaturamento / totGasto : 0;
  const roiGeral = totGasto ? 100 * (totLiquido - totGasto) / totGasto : 0;

  const linha = function (r) {
    return '<tr>' +
      '<td class="nm">' + (ehDiamante_(r) ? '💎 ' : '') + r.anuncio + '</td>' +
      '<td>' + (r.formato || '—') + '</td>' +
      '<td>' + (r.estrutura || '—') + '</td>' +
      '<td>' + (r.angulo || '—') + '</td>' +
      '<td class="hk" title="' + String(r.hook || '').replace(/"/g, '&quot;') + '">' + hookCurto_(r.hook) + '</td>' +
      '<td>' + (r.emocao || '—') + '</td>' +
      '<td>' + (r.amplificador || '—') + '</td>' +
      '<td>' + (r.arquetipo || '—') + '</td>' +
      '<td>' + (r.segmentacao || '—') + '</td>' +
      '<td>' + (r.prova || '—') + '</td>' +
      '<td class="num">' + moeda_(r.gasto) + '</td>' +
      '<td class="num">' + (r.ctr_pct !== '' ? Number(r.ctr_pct).toFixed(2) + '%' : '—') + '</td>' +
      '<td class="num">' + (r.hook_rate_pct !== '' ? Number(r.hook_rate_pct).toFixed(1) + '%' : '—') + '</td>' +
      '<td class="num">' + (r.hold_rate_pct !== '' ? Number(r.hold_rate_pct).toFixed(1) + '%' : '—') + '</td>' +
      '<td class="num">' + (r.views25_pct !== '' ? Number(r.views25_pct).toFixed(1) + '%' : '—') + '</td>' +
      '<td class="num">' + (r.checkouts !== '' ? r.checkouts : '—') + '</td>' +
      '<td class="num">' + moeda_(r.custo_checkout) + '</td>' +
      '<td class="num">' + (r.vendas !== '' ? r.vendas : '—') + '</td>' +
      '<td class="num vend">' + moeda_(r.custo_venda) + '</td>' +
      '<td class="num">' + (Number(r.faturamento) ? moeda_(r.faturamento) : '—') + '</td>' +
      '<td class="num" style="color:' + (Number(r.roas) >= 1 ? '#2e7d52' : '#c0392b') + '">' + (r.roas !== '' && r.roas !== null ? Number(r.roas).toFixed(2) + 'x' : '—') + '</td>' +
      '<td class="num" style="color:' + (Number(r.roi_pct) >= 0 ? '#2e7d52' : '#c0392b') + '">' + (r.roi_pct !== '' && r.roi_pct !== null ? Number(r.roi_pct).toFixed(0) + '%' : '—') + '</td>' +
      '</tr>';
  };

  const celulasRoi_ = function (vendas, gasto, faturamento) {
    const fat = Number(faturamento) || 0;
    const liq = liquido_(fat, vendas);
    const g2 = Number(gasto) || 0;
    const roas = g2 ? fat / g2 : null;
    const roi = g2 ? 100 * (liq - g2) / g2 : null;
    return '<td class="num">' + (fat ? moeda_(fat) : '—') + '</td>'
      + '<td class="num" style="color:' + (roas !== null && roas >= 1 ? '#2e7d52' : '#c0392b') + '">' + (roas !== null ? roas.toFixed(2) + 'x' : '—') + '</td>'
      + '<td class="num" style="color:' + (roi !== null && roi >= 0 ? '#2e7d52' : '#c0392b') + '">' + (roi !== null ? roi.toFixed(0) + '%' : '—') + '</td>';
  };

  const grupoHtml = function (titulo, campo) {
    const gs = agruparPor_(rows, campo);
    let h = '<div class="card"><div class="sec-title">' + titulo + '</div><table class="tb"><tr><th>' + titulo.replace('Por ', '') + '</th><th class="num">Criativos</th><th class="num">Gasto</th><th class="num">Checkouts</th><th class="num">R$/Checkout</th><th class="num">Vendas</th><th class="num">R$/Venda</th><th class="num">Faturamento</th><th class="num">ROAS</th><th class="num">ROI</th></tr>';
    gs.forEach(function (g) {
      h += '<tr><td class="nm">' + g.nome + '</td><td class="num">' + g.n + '</td><td class="num">' + moeda_(g.gasto) + '</td><td class="num">' + g.checkouts + '</td><td class="num">' + moeda_(g.custoCheckout) + '</td><td class="num">' + g.vendas + '</td><td class="num vend">' + moeda_(g.custoVenda) + '</td>' + celulasRoi_(g.vendas, g.gasto, g.faturamento) + '</tr>';
    });
    return h + '</table></div>';
  };

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;}
  body{background:#f4f5f4;color:#1a1a1a;padding:24px;}
  .wrap{max-width:1100px;margin:0 auto;}
  h1{font-size:22px;font-weight:600;}
  .sub{color:#7a7a7a;font-size:13px;margin-bottom:16px;}
  .tabs{display:flex;gap:6px;margin-bottom:16px;}
  .tabs a{padding:8px 18px;border-radius:10px;font-size:13.5px;font-weight:600;text-decoration:none;color:#666;background:#fff;border:1px solid #e2e2e2;}
  .tabs a.active{background:#1a7a4f;color:#fff;border-color:#1a7a4f;}
  .filterbar{position:relative;margin-bottom:20px;}
  .rangeBtn{border:1px solid #d5d5d5;background:#fff;border-radius:10px;padding:10px 16px;font-size:13.5px;font-weight:600;color:#333;cursor:pointer;display:inline-flex;align-items:center;gap:8px;}
  .rangeBtn .car{color:#999;font-size:11px;}
  .rangePanel{position:absolute;top:48px;left:0;z-index:60;background:#fff;border:1px solid #e2e2e2;border-radius:14px;box-shadow:0 10px 34px rgba(0,0,0,.14);display:flex;overflow:hidden;}
  .presets{width:196px;border-right:1px solid #eee;padding:8px;}
  .presets a{display:block;padding:9px 12px;border-radius:8px;font-size:13px;color:#444;text-decoration:none;cursor:pointer;}
  .presets a:hover{background:#f1f4f2;}
  .presets a.active{background:#e9f3ee;color:#2f6b4f;font-weight:600;}
  .presets .div{height:1px;background:#eee;margin:6px 8px;}
  .calarea{padding:14px 16px;width:278px;}
  .calnav{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;font-weight:600;font-size:14px;color:#333;}
  .calnav button{border:none;background:none;font-size:20px;line-height:1;cursor:pointer;color:#666;padding:2px 10px;border-radius:6px;}
  .calnav button:hover{background:#f0f0f0;}
  .calgrid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}
  .caldow{font-size:10.5px;color:#aaa;text-align:center;padding:4px 0;text-transform:uppercase;}
  .calday{text-align:center;font-size:12.5px;padding:8px 0;border-radius:7px;cursor:pointer;color:#333;}
  .calday:hover{background:#eef2f0;}
  .calday.muted{visibility:hidden;}
  .calday.inrange{background:#e9f3ee;border-radius:0;}
  .calday.edge{background:#1a7a4f;color:#fff;border-radius:7px;}
  .calfoot{display:flex;align-items:center;gap:10px;margin-top:14px;}
  .calsel{flex:1;font-size:12px;color:#666;}
  .cbtn{border:1px solid #dcdcdc;background:#fff;border-radius:8px;padding:8px 16px;font-size:12.5px;cursor:pointer;}
  .cbtn.primary{background:#1a7a4f;color:#fff;border-color:#1a7a4f;font-weight:600;}
  @media(max-width:560px){.rangePanel{flex-direction:column;}.presets{width:100%;border-right:none;border-bottom:1px solid #eee;}.calarea{width:100%;}}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:18px;}
  .kpi{background:#fff;border-radius:12px;padding:16px 18px;}
  .kpi .lbl{font-size:12px;color:#8a8a8a;text-transform:uppercase;letter-spacing:.04em;}
  .kpi .val{font-size:30px;font-weight:600;margin-top:4px;}
  .card{background:#fff;border-radius:14px;padding:20px;margin-bottom:18px;overflow-x:auto;}
  .sec-title{font-size:15px;font-weight:600;margin-bottom:10px;}
  .tb{width:100%;border-collapse:collapse;font-size:12.5px;}
  .tb th{text-align:left;color:#8a8a8a;font-size:11px;text-transform:uppercase;letter-spacing:.03em;padding:8px 10px;border-bottom:1px solid #eee;}
  .tb td{padding:9px 10px;border-bottom:1px solid #f2f2f2;}
  .tb .nm{font-weight:600;white-space:nowrap;}
  .tb .hk{max-width:220px;}
  .tb .num,.tb th.num{text-align:right;white-space:nowrap;}
  .tb .vend{font-weight:600;color:#1a7a4f;}
  .note{font-size:12px;color:#8a8a8a;margin-top:6px;}
  .charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px;margin-bottom:18px;}
  .chartCard{background:#fff;border-radius:14px;padding:16px;}
  .chartCard .ttl{font-size:13.5px;font-weight:600;margin-bottom:2px;}
  .chartCard .hint{font-size:11px;color:#9a9a9a;margin-bottom:8px;}
  .chartCard .cv{position:relative;height:170px;}
  .pieCtrls{display:flex;gap:16px;margin-bottom:12px;flex-wrap:wrap;}
  .pieCtrls label{font-size:12px;color:#8a8a8a;display:flex;flex-direction:column;gap:4px;font-weight:600;}
  .pieCtrls select{border:1px solid #d5d5d5;background:#fff;border-radius:8px;padding:8px 12px;font-size:13px;color:#333;min-width:200px;}
  .pieWrap{position:relative;height:260px;max-width:520px;margin:0 auto;}
</style></head><body>
<div class="wrap">
  <h1>Criativos — Ciclo Fértil</h1>
  <div class="sub">Etiquetas editáveis na aba 'Criativos' da planilha · métricas do Meta</div>
  <div class="tabs">
    <a href="${URL}" target="_top">Funil</a>
    <a class="active" href="${URL}?aba=criativos" target="_top">Criativos</a>
  </div>
  <div class="filterbar">
    <button class="rangeBtn" onclick="var p=document.getElementById('rangePanel');p.style.display=(p.style.display==='none'||!p.style.display)?'flex':'none';">
      <span>📅</span><span>${range.label}</span><span class="car">▼</span>
    </button>
    <div id="rangePanel" class="rangePanel" style="display:none;">
      <div class="presets">
        <a class="${on('hoje')}"  href="${URL}?aba=criativos&preset=hoje"  target="_top">Hoje</a>
        <a class="${on('ontem')}" href="${URL}?aba=criativos&preset=ontem" target="_top">Ontem</a>
        <a class="${on('7d')}"    href="${URL}?aba=criativos&preset=7d"    target="_top">Últimos 7 dias</a>
        <a class="${on('14d')}"   href="${URL}?aba=criativos&preset=14d"   target="_top">Últimos 14 dias</a>
        <a class="${on('28d')}"   href="${URL}?aba=criativos&preset=28d"   target="_top">Últimos 28 dias</a>
        <a class="${on('30d')}"   href="${URL}?aba=criativos&preset=30d"   target="_top">Últimos 30 dias</a>
        <div class="div"></div>
        <a class="${on('semana')}"         href="${URL}?aba=criativos&preset=semana"         target="_top">Esta semana</a>
        <a class="${on('semana_passada')}" href="${URL}?aba=criativos&preset=semana_passada" target="_top">Semana passada</a>
        <a class="${on('mes')}"            href="${URL}?aba=criativos&preset=mes"            target="_top">Este mês</a>
        <a class="${on('mes_passado')}"    href="${URL}?aba=criativos&preset=mes_passado"    target="_top">Mês passado</a>
        <div class="div"></div>
        <a class="${on('todos')}" href="${URL}?aba=criativos&preset=todos" target="_top">Máximo (todos)</a>
      </div>
      <div class="calarea">
        <div class="calnav"><button onclick="calMove(-1)">‹</button><span id="calTitle"></span><button onclick="calMove(1)">›</button></div>
        <div id="calGrid" class="calgrid"></div>
        <div class="calfoot">
          <span id="calSel" class="calsel">Selecione as datas</span>
          <button class="cbtn primary" onclick="calApply()">Aplicar</button>
        </div>
      </div>
    </div>
  </div>
  <div class="kpis">
    <div class="kpi"><div class="lbl">Criativos</div><div class="val">${testados}/25</div></div>
    <div class="kpi"><div class="lbl">Gasto total</div><div class="val">${moeda_(totGasto)}</div></div>
    <div class="kpi"><div class="lbl">Checkouts</div><div class="val">${totCheckouts}</div></div>
    <div class="kpi"><div class="lbl">Vendas</div><div class="val">${totVendas}</div></div>
    <div class="kpi"><div class="lbl">Faturamento</div><div class="val">${moeda_(totFaturamento)}</div></div>
    <div class="kpi"><div class="lbl">ROAS</div><div class="val" style="color:${!totGasto ? '#8a8a8a' : (roasGeral >= 1 ? '#2e7d52' : '#c0392b')}">${totGasto ? roasGeral.toFixed(2) + 'x' : '—'}</div></div>
    <div class="kpi"><div class="lbl">ROI</div><div class="val" style="color:${!totGasto ? '#8a8a8a' : (roiGeral >= 0 ? '#2e7d52' : '#c0392b')}">${totGasto ? roiGeral.toFixed(0) + '%' : '—'}</div></div>
  </div>
  <div class="charts">
    <div class="chartCard"><div class="ttl">Custo por venda</div><div class="hint">menor = melhor · só criativos com venda</div><div class="cv"><canvas id="chVenda"></canvas></div></div>
    <div class="chartCard"><div class="ttl">Custo por checkout</div><div class="hint">menor = melhor</div><div class="cv"><canvas id="chCheckout"></canvas></div></div>
    <div class="chartCard"><div class="ttl">Hook rate</div><div class="hint">maior = melhor · % que passa dos 3s</div><div class="cv"><canvas id="chHook"></canvas></div></div>
    <div class="chartCard"><div class="ttl">Hold rate</div><div class="hint">maior = melhor · % dos que passaram dos 3s e chegam ao ThruPlay</div><div class="cv"><canvas id="chHold"></canvas></div></div>
  </div>
  <div class="card">
    <div class="sec-title">Análise por grupo</div>
    <div class="pieCtrls">
      <label>Métrica
        <select id="selMetrica">
          <option value="cpa">CPA (custo por venda)</option>
          <option value="roas">ROAS</option>
          <option value="roi">ROI</option>
          <option value="faturamento">Faturamento</option>
          <option value="ctr">CTR</option>
          <option value="cpc">CPC</option>
          <option value="gasto">Valor gasto</option>
        </select>
      </label>
      <label>Agrupar por
        <select id="selGrupo">
          <option value="formato">Formato</option>
          <option value="estrutura">Estrutura</option>
          <option value="angulo">Ângulo</option>
          <option value="emocao">Emoção</option>
          <option value="amplificador">Amplificador</option>
          <option value="arquetipo">Arquétipo</option>
          <option value="segmentacao">Segmentação</option>
          <option value="prova">Prova</option>
        </select>
      </label>
    </div>
    <div class="pieWrap"><canvas id="chPie"></canvas></div>
    <div class="note" id="pieNota"></div>
  </div>
  <div class="card">
    <div class="sec-title">Todos os criativos</div>
    <table class="tb">
      <tr><th>Anúncio</th><th>Formato</th><th>Estrutura</th><th>Ângulo</th><th>Hook</th><th>Emoção</th><th>Amplificador</th><th>Arquétipo</th><th>Segmentação</th><th>Prova</th><th class="num">Gasto</th><th class="num">CTR</th><th class="num">Hook rate</th><th class="num">Hold rate</th><th class="num">Views 25%</th><th class="num">Checkouts</th><th class="num">R$/Checkout</th><th class="num">Vendas</th><th class="num">R$/Venda</th><th class="num">Faturamento</th><th class="num">ROAS</th><th class="num">ROI</th></tr>
      ${rows.map(linha).join('')}
    </table>
    <div class="note">Preencha as etiquetas (formato, estrutura, angulo, hook, emocao, amplificador, arquetipo, segmentacao, prova) direto na planilha (aba Criativos). Os agrupamentos e a pizza usam essas etiquetas.</div>
  </div>
  ${grupoHtml('Por formato', 'formato')}
  ${grupoHtml('Por ângulo', 'angulo')}
  ${grupoHtml('Por emoção', 'emocao')}
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>
  var URL = "${URL}", TODAY = "${T}";
  var selFrom = "${curFrom}", selTo = "${curTo}";
  var MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  var vd; (function(){ var base = selFrom || TODAY; var p = base.split('-'); vd = new Date(+p[0], +p[1]-1, 1); })();
  function pad2(n){ return (n<10?'0':'')+n; }
  function keyOf(y,m,d){ return y+'-'+pad2(m+1)+'-'+pad2(d); }
  function brdate(k){ var p=k.split('-'); return p[2]+'/'+p[1]+'/'+p[0]; }
  function renderCal(){
    var y=vd.getFullYear(), m=vd.getMonth();
    document.getElementById('calTitle').textContent = MESES[m]+' '+y;
    var startDow=(new Date(y,m,1).getDay()+6)%7, dias=new Date(y,m+1,0).getDate();
    var html='';
    ['seg','ter','qua','qui','sex','sáb','dom'].forEach(function(d){ html+='<div class="caldow">'+d+'</div>'; });
    for(var i=0;i<startDow;i++) html+='<div class="calday muted"></div>';
    for(var d=1;d<=dias;d++){
      var k=keyOf(y,m,d), cls='calday';
      if(selFrom && selTo && k>=selFrom && k<=selTo) cls+=' inrange';
      if(k===selFrom || k===selTo) cls+=' edge';
      html+='<div class="'+cls+'" data-k="'+k+'">'+d+'</div>';
    }
    document.getElementById('calGrid').innerHTML=html;
    var sel=document.getElementById('calSel');
    if(selFrom && selTo) sel.textContent=brdate(selFrom)+' – '+brdate(selTo);
    else if(selFrom) sel.textContent=brdate(selFrom)+' – ...';
    else sel.textContent='Selecione as datas';
  }
  function pickDay(k){
    if(!selFrom || (selFrom && selTo)){ selFrom=k; selTo=''; }
    else if(k>=selFrom){ selTo=k; }
    else { selTo=selFrom; selFrom=k; }
    renderCal();
  }
  function calMove(n){ vd=new Date(vd.getFullYear(), vd.getMonth()+n, 1); renderCal(); }
  function calApply(){ if(!selFrom){ alert('Selecione uma data no calendário'); return; } window.top.location.href = URL+'?aba=criativos&from='+selFrom+'&to='+(selTo||selFrom); }
  document.getElementById('calGrid').addEventListener('click', function(e){ var k=e.target.getAttribute('data-k'); if(k) pickDay(k); });
  renderCal();

  var CR = ${rowsJson};
  var PRECO_VENDA = ${PRECO_VENDA}, PRECO_BUMP = ${PRECO_BUMP};
  var LIQUIDO_VENDA = ${LIQUIDO_VENDA}, LIQUIDO_BUMP = ${LIQUIDO_BUMP};
  function liquidoJs(fat, vendas){
    var v = Number(vendas) || 0; if(!v) return 0;
    var bumps = Math.max(0, Math.round(((Number(fat)||0) - v * PRECO_VENDA) / PRECO_BUMP));
    return v * LIQUIDO_VENDA + bumps * LIQUIDO_BUMP;
  }
  function nomeCurto(s){ s = String(s||''); return s.length > 22 ? s.slice(0,22) + '…' : s; }
  function ehDiamante(r){
    var ctr = Number(r.ctr_pct) || 0, impr = Number(r.impressoes) || 0, gasto = Number(r.gasto) || 0;
    var cliques = impr * ctr / 100, cpc = cliques ? gasto / cliques : null;
    return gasto >= ${DIAMANTE_GASTO_MIN} && ctr >= ${DIAMANTE_CTR_MIN} && cpc !== null && cpc <= ${DIAMANTE_CPC_MAX};
  }
  function barra(id, itens, campo, cor, moeda, asc){
    itens = itens.filter(function(r){ var v = Number(r[campo]); return r[campo] !== '' && !isNaN(v) && v > 0; });
    itens.sort(function(a,b){ return asc ? (a[campo]-b[campo]) : (b[campo]-a[campo]); });
    itens = itens.slice(0,8);
    if(!itens.length){ document.getElementById(id).parentNode.innerHTML = '<div style="color:#aaa;font-size:12px;padding:20px 0;">Sem dados ainda</div>'; return; }
    new Chart(document.getElementById(id), {
      type: 'bar',
      data: { labels: itens.map(function(r){ return (ehDiamante(r) ? '💎 ' : '') + nomeCurto(r.anuncio); }),
        datasets: [{ data: itens.map(function(r){ return Number(r[campo]); }), backgroundColor: cor, borderRadius: 6 }] },
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c){ return moeda ? 'R$' + c.parsed.x.toFixed(2) : c.parsed.x.toFixed(1) + '%'; } } } },
        scales: { x: { ticks: { font: { size: 10 }, callback: function(v){ return moeda ? 'R$' + v : v + '%'; } }, grid: { color: '#f0f0f0' } },
                  y: { ticks: { font: { size: 10.5 } }, grid: { display: false } } } }
    });
  }
  barra('chVenda',   CR, 'custo_venda',    '#1a7a4f', true,  true);
  barra('chCheckout',CR, 'custo_checkout', '#5d8aa8', true,  true);
  barra('chHook',    CR, 'hook_rate_pct',  '#b0764a', false, false);
  barra('chHold',    CR, 'hold_rate_pct',  '#7a5da8', false, false);

  // ── Pizza com seletores: métrica × agrupador ──
  var CORES = ['#1a7a4f','#b0764a','#5d8aa8','#7a5da8','#c05b4d','#c9a97a','#4a6b6b','#8a8a5d'];
  var pieChart = null;

  function agregaPizza(campoGrupo, metrica){
    var g = {};
    CR.forEach(function(r){
      var k = String(r[campoGrupo] || '(sem etiqueta)');
      if(!g[k]) g[k] = { gasto:0, vendas:0, impr:0, cliques:0, fat:0 };
      var gasto = Number(r.gasto) || 0;
      var impr = Number(r.impressoes) || 0;
      var ctr = Number(r.ctr_pct) || 0;
      g[k].gasto += gasto;
      g[k].vendas += Number(r.vendas) || 0;
      g[k].fat += Number(r.faturamento) || 0;
      g[k].impr += impr;
      g[k].cliques += impr * ctr / 100;
    });
    var labels = [], valores = [], fora = [];
    Object.keys(g).forEach(function(k){
      var x = g[k], v = null;
      var fat = x.fat;
      var liq = liquidoJs(fat, x.vendas);
      if(metrica === 'cpa') v = x.vendas ? x.gasto / x.vendas : null;
      if(metrica === 'roas') v = x.gasto ? fat / x.gasto : null;
      if(metrica === 'faturamento') v = fat;
      if(metrica === 'roi'){
        var roi = x.gasto ? 100 * (liq - x.gasto) / x.gasto : null;
        if(roi !== null && roi <= 0) fora.push(k + ' (' + roi.toFixed(0) + '%)');
        v = roi;
      }
      if(metrica === 'ctr') v = x.impr ? 100 * x.cliques / x.impr : null;
      if(metrica === 'cpc') v = x.cliques ? x.gasto / x.cliques : null;
      if(metrica === 'gasto') v = x.gasto;
      if(v !== null && isFinite(v) && v > 0){ labels.push(k); valores.push(Math.round(v*100)/100); }
    });
    return { labels: labels, valores: valores, fora: fora };
  }

  function desenhaPizza(){
    var metrica = document.getElementById('selMetrica').value;
    var grupo = document.getElementById('selGrupo').value;
    var d = agregaPizza(grupo, metrica);
    var nota = document.getElementById('pieNota');
    var ehMoeda = (metrica === 'cpa' || metrica === 'cpc' || metrica === 'gasto' || metrica === 'faturamento');
    var ehX = (metrica === 'roas');
    nota.textContent = metrica === 'cpa' ? 'CPA = gasto ÷ vendas do grupo. Grupos sem venda ficam de fora.'
      : metrica === 'roas' ? 'ROAS = faturamento ÷ gasto do grupo. Faturamento é o valor de compra reportado pelo Meta, já com order bump. Abaixo de 1x o grupo dá prejuízo.'
      : metrica === 'roi' ? 'ROI = (líquido − gasto) ÷ gasto. Líquido = R$33,50 por venda + R$14,48 por order bump. Só grupos com ROI positivo cabem na pizza.'
        + (d.fora && d.fora.length ? ' No vermelho: ' + d.fora.join(', ') + '.' : '')
      : metrica === 'faturamento' ? 'Faturamento do grupo, valor de compra reportado pelo Meta.'
      : metrica === 'cpc' ? 'CPC = gasto ÷ cliques do grupo.'
      : metrica === 'gasto' ? 'Soma do gasto de cada grupo.'
      : 'CTR médio ponderado por impressões do grupo.';
    if(pieChart) pieChart.destroy();
    if(!d.labels.length){
      nota.textContent = metrica === 'roi' && d.fora && d.fora.length
        ? 'Nenhum grupo com ROI positivo ainda. No vermelho: ' + d.fora.join(', ') + '.'
        : 'Sem dados para essa combinação ainda.';
      return;
    }
    pieChart = new Chart(document.getElementById('chPie'), {
      type: 'pie',
      data: { labels: d.labels, datasets: [{ data: d.valores, backgroundColor: CORES.slice(0, d.labels.length), borderColor: '#fff', borderWidth: 2 }] },
      options: { responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { font: { size: 11.5 }, boxWidth: 14 } },
          tooltip: { callbacks: { label: function(c){
            var v = c.parsed;
            return c.label + ': ' + (ehMoeda ? 'R$' + v.toFixed(2) : ehX ? v.toFixed(2) + 'x' : v.toFixed(2) + '%');
          } } }
        } }
    });
  }
  document.getElementById('selMetrica').addEventListener('change', desenhaPizza);
  document.getElementById('selGrupo').addEventListener('change', desenhaPizza);
  desenhaPizza();
</script>
</body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ════════════════════════════════════════════════════
   ATUALIZAÇÃO AUTOMÁTICA VIA API DO META
   Requisitos:
   1. Propriedade do script META_TOKEN (Configurações do projeto → Propriedades do script)
   2. Rodar atualizarCriativos() 1x no editor para autorizar
   3. Rodar criarGatilhoDiario() 1x para agendar (todo dia às 6h)
   ════════════════════════════════════════════════════ */
const META_AD_ACCOUNT = 'act_700396725669598';   // Info produtos - Brasil (Bals digital)
const META_CAMP_FILTRO = 'Quiz/v2';              // só campanhas com isso no nome
const META_EXCLUIR_CAMPANHAS = ['120251947450640652','120251870932890652']; // testes antigos descartados
const META_EXCLUIR_ADS = ['120251807360980652']; // ad2 = copia acidental do ad3
const META_API_VER = 'v21.0';

function atualizarCriativos() {
  const token = PropertiesService.getScriptProperties().getProperty('META_TOKEN');
  if (!token) throw new Error('Defina a propriedade META_TOKEN nas Configurações do projeto.');

  const fields = ['ad_id','ad_name','campaign_id','campaign_name','spend','impressions','clicks',
    'actions','action_values','video_thruplay_watched_actions','video_continuous_2_sec_watched_actions',
    'video_p25_watched_actions','video_play_actions'].join(',');
  const filtering = encodeURIComponent(JSON.stringify([{ field: 'campaign.name', operator: 'CONTAIN', value: META_CAMP_FILTRO }]));
  let url = 'https://graph.facebook.com/' + META_API_VER + '/' + META_AD_ACCOUNT + '/insights'
    + '?level=ad&date_preset=maximum&time_increment=1&limit=500'
    + '&fields=' + fields
    + '&filtering=' + filtering
    + '&access_token=' + encodeURIComponent(token);

  // tenta incluir o campo de 3s (hook rate oficial); se a API rejeitar, refaz sem ele
  let data;
  try {
    data = metaFetch_(url.replace('&fields=' + fields, '&fields=' + fields + ',video_3_sec_watched_actions'));
  } catch (err) {
    data = metaFetch_(url);
  }

  // cada linha da API = um anúncio num dia. Monta a aba diária e o total por anúncio.
  const ag = {};
  const diario = [];
  (data || []).forEach(function (r) {
    if (META_EXCLUIR_CAMPANHAS.indexOf(String(r.campaign_id)) !== -1) return;
    if (META_EXCLUIR_ADS.indexOf(String(r.ad_id)) !== -1) return;
    const nome = r.ad_name || r.ad_id;

    // valores desse dia
    const d = {
      spend: Number(r.spend) || 0,
      impr: Number(r.impressions) || 0,
      clicks: Number(r.clicks) || 0,
      thruplay: somaAcao_(r.video_thruplay_watched_actions),
      v3s: somaAcao_(r.video_3_sec_watched_actions),
      v2s: somaAcao_(r.video_continuous_2_sec_watched_actions),
      p25: somaAcao_(r.video_p25_watched_actions),
      lpv: 0, ic: 0, compras: 0, valor: 0
    };
    (r.actions || []).forEach(function (ac) {
      const t = ac.action_type || '';
      const v = Number(ac.value) || 0;
      if (t === 'video_view') d.v3s += v;
      if (t === 'landing_page_view') d.lpv += v;
      if (t === 'initiate_checkout' || t === 'omni_initiated_checkout' || t === 'offsite_conversion.fb_pixel_initiate_checkout') d.ic = Math.max(d.ic, v);
      if (t === 'purchase' || t === 'omni_purchase' || t === 'offsite_conversion.fb_pixel_purchase') d.compras = Math.max(d.compras, v);
    });
    // valor de compra reportado pelo pixel: já vem com o order bump somado
    (r.action_values || []).forEach(function (ac) {
      const t = ac.action_type || '';
      const v = Number(ac.value) || 0;
      if (t === 'purchase' || t === 'omni_purchase' || t === 'offsite_conversion.fb_pixel_purchase') d.valor = Math.max(d.valor, v);
    });

    diario.push([String(r.date_start || ''), String(r.ad_id || ''), nome, r.campaign_name || '',
      round2_(d.spend), d.impr, d.clicks, d.v3s, d.v2s, d.thruplay, d.p25, d.lpv, d.ic, d.compras, round2_(d.valor)]);

    if (!ag[nome]) ag[nome] = { ad_id: r.ad_id, campanha: r.campaign_name, spend: 0, impr: 0, clicks: 0, thruplay: 0, v3s: 0, v2s: 0, p25: 0, lpv: 0, ic: 0, compras: 0, valor: 0 };
    const a = ag[nome];
    a.spend += d.spend; a.impr += d.impr; a.clicks += d.clicks;
    a.thruplay += d.thruplay; a.v3s += d.v3s; a.v2s += d.v2s; a.p25 += d.p25;
    a.lpv += d.lpv; a.ic += d.ic; a.compras += d.compras; a.valor += d.valor;
    a.campanha = r.campaign_name || a.campanha;
  });

  gravarDiario_(diario);

  // grava na planilha preservando as etiquetas
  const sh = getCriativosSheet_();
  const vals = sh.getDataRange().getValues();
  const idxPorNome = {};
  for (let i = 1; i < vals.length; i++) idxPorNome[String(vals[i][1])] = i + 1; // linha real
  const hoje = fmt_(new Date());
  const col = function (h) { return CRIATIVOS_HEADERS.indexOf(h) + 1; };

  Object.keys(ag).forEach(function (nome) {
    const a = ag[nome];
    if (!a.impr) return; // ignora ads sem entrega
    const hookPct = a.impr ? round1_(100 * (a.v3s || a.v2s) / a.impr) : '';
    const base3s = a.v3s || a.v2s;
    const holdPct = base3s ? round1_(100 * a.thruplay / base3s) : '';
    const v25Pct = a.impr ? round1_(100 * a.p25 / a.impr) : '';
    const ctrPct = a.impr ? Math.round(10000 * a.clicks / a.impr) / 100 : '';
    let linha = idxPorNome[nome];
    if (!linha) {
      sh.appendRow([a.ad_id, nome, a.campanha || '', 'rodando', '', '', '', '', '', '', '', '', '', 0, 0, '', '', '', '', '', '', 0, '', 0, '', 0, '', '', hoje]);
      linha = sh.getLastRow();
      idxPorNome[nome] = linha;
    }
    sh.getRange(linha, col('gasto')).setValue(round2_(a.spend));
    sh.getRange(linha, col('impressoes')).setValue(a.impr);
    if (ctrPct !== '') sh.getRange(linha, col('ctr_pct')).setValue(ctrPct);
    if (hookPct !== '' && (a.v3s || a.v2s)) sh.getRange(linha, col('hook_rate_pct')).setValue(hookPct);
    sh.getRange(linha, col('hold_rate_pct')).setValue(holdPct);
    sh.getRange(linha, col('views25_pct')).setValue(v25Pct);
    sh.getRange(linha, col('lpv')).setValue(a.lpv || '');
    sh.getRange(linha, col('custo_lpv')).setValue(a.lpv ? round2_(a.spend / a.lpv) : '');
    sh.getRange(linha, col('checkouts')).setValue(a.ic || 0);
    sh.getRange(linha, col('custo_checkout')).setValue(a.ic ? round2_(a.spend / a.ic) : '');
    sh.getRange(linha, col('vendas')).setValue(a.compras || 0);
    sh.getRange(linha, col('custo_venda')).setValue(a.compras ? round2_(a.spend / a.compras) : '');
    const fat = a.valor > 0 ? a.valor : (a.compras || 0) * PRECO_VENDA;
    const liq = liquido_(fat, a.compras);
    sh.getRange(linha, col('faturamento')).setValue(round2_(fat));
    sh.getRange(linha, col('roas')).setValue(a.spend ? Math.round(100 * fat / a.spend) / 100 : '');
    sh.getRange(linha, col('roi_pct')).setValue(a.spend ? round1_(100 * (liq - a.spend) / a.spend) : '');
    sh.getRange(linha, col('atualizado')).setValue(hoje);
  });
  return 'Atualizado: ' + Object.keys(ag).length + ' criativos';
}

/* Reescreve a aba diária inteira. Como a busca é sempre date_preset=maximum,
   o conjunto é o histórico completo e não precisa de merge nem dedupe. */
function gravarDiario_(linhas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CRIATIVOS_DIARIO);
  if (!sh) { sh = ss.insertSheet(CRIATIVOS_DIARIO); sh.hideSheet(); }
  sh.clear();
  sh.getRange(1, 1, 1, DIARIO_HEADERS.length).setValues([DIARIO_HEADERS]);
  if (linhas.length) sh.getRange(2, 1, linhas.length, DIARIO_HEADERS.length).setValues(linhas);
  return linhas.length;
}

/* Soma a aba diária dentro do período e devolve um mapa por nome do anúncio.
   from/to vazios = tudo. Retorna null se a aba ainda não existir. */
function lerDiarioAgregado_(from, to) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CRIATIVOS_DIARIO);
  if (!sh || sh.getLastRow() < 2) return null;
  const larg = Math.min(DIARIO_HEADERS.length, sh.getLastColumn());
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, larg).getValues();
  const out = {};
  vals.forEach(function (r) {
    // o Sheets converte "2026-08-05" em Date; normaliza os dois casos
    const dia = (r[0] instanceof Date) ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0] || '').slice(0, 10);
    if (!dia) return;
    if (from && dia < from) return;
    if (to && dia > to) return;
    const nome = String(r[2] || '');
    if (!nome) return;
    if (!out[nome]) out[nome] = { ad_id: r[1], campanha: r[3], spend: 0, impr: 0, clicks: 0, v3s: 0, v2s: 0, thruplay: 0, p25: 0, lpv: 0, ic: 0, compras: 0, valor: 0 };
    const a = out[nome];
    a.spend += Number(r[4]) || 0;
    a.impr  += Number(r[5]) || 0;
    a.clicks += Number(r[6]) || 0;
    a.v3s += Number(r[7]) || 0;
    a.v2s += Number(r[8]) || 0;
    a.thruplay += Number(r[9]) || 0;
    a.p25 += Number(r[10]) || 0;
    a.lpv += Number(r[11]) || 0;
    a.ic += Number(r[12]) || 0;
    a.compras += Number(r[13]) || 0;
    a.valor += Number(r[14]) || 0;
  });
  return out;
}

function metaFetch_(url) {
  const out = [];
  let next = url;
  let guard = 0;
  while (next && guard < 10) {
    const resp = UrlFetchApp.fetch(next, { muteHttpExceptions: true });
    const json = JSON.parse(resp.getContentText());
    if (json.error) throw new Error('Meta API: ' + json.error.message);
    (json.data || []).forEach(function (d) { out.push(d); });
    next = json.paging && json.paging.next ? json.paging.next : null;
    guard++;
  }
  return out;
}

function somaAcao_(arr) {
  if (!arr || !arr.length) return 0;
  return arr.reduce(function (s, x) { return s + (Number(x.value) || 0); }, 0);
}
function round1_(v) { return Math.round(v * 10) / 10; }
function round2_(v) { return Math.round(v * 100) / 100; }

function criarGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'atualizarCriativos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('atualizarCriativos').timeBased().everyHours(2).create();
  return 'Gatilho criado (a cada 2 horas).';
}

/* Teste: a API devolve o campo de 3 segundos (hook rate)?
   Rodar no editor e olhar o log (Ctrl+Enter mostra o registro de execução). */
function testarHook3s() {
  const token = PropertiesService.getScriptProperties().getProperty('META_TOKEN');
  if (!token) throw new Error('Defina a propriedade META_TOKEN primeiro.');
  const url = 'https://graph.facebook.com/' + META_API_VER + '/' + META_AD_ACCOUNT + '/insights'
    + '?level=ad&date_preset=maximum&limit=10'
    + '&fields=ad_name,impressions,video_3_sec_watched_actions'
    + '&filtering=' + encodeURIComponent(JSON.stringify([{ field: 'campaign.name', operator: 'CONTAIN', value: META_CAMP_FILTRO }]))
    + '&access_token=' + encodeURIComponent(token);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(resp.getContentText());
  return resp.getContentText();
}
