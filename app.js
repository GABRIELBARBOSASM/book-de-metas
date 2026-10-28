/* Book de Metas – app */
const C = BookCalc, MESES = C.MESES, EMPRESAS = ['CSC','Ferguminas','MIB','MIG','MML','SM Metais','SM Mineração'];
let ROWS = [], S = { ano:null, mes:null, empresa:'CSC', cargos:new Set(), resp:'Todos', nivel:'Todos', indicador:'Todos', selInd:null, atgView:'ind' };
const charts = {};
let LAST_SRC = null; // última fonte carregada (para o botão Atualizar)
if (window.Chart) { Chart.defaults.font.family = "'Poppins', Arial, sans-serif"; Chart.defaults.color = '#5f6c77'; Chart.defaults.borderColor = '#eef0f1'; }
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const nf = (d) => new Intl.NumberFormat('pt-BR', { minimumFractionDigits:d, maximumFractionDigits:d });
const pct = (v, d = 0) => v === null || v === undefined ? '—' : nf(d).format((Math.abs(v) < 5e-7 ? 0 : v) * 100).replace(/^-(0(,0+)?)$/, '$1') + '%';
// casas decimais "inteligentes": valores pequenos (ex.: 0,0001) não viram 0,00
const fmtV = (v, r) => { if (v === null || v === undefined) return ''; const base = r && r.formato === 'TRES CASAS' ? 3 : 2, a = Math.abs(v);
  const d = a > 0 && a < 1 ? Math.min(6, Math.max(base, Math.ceil(-Math.log10(a)) + 1)) : base; return nf(d).format(v); };
const LIM = C.limiteMes(new Date());

// ---------- carga ----------
function loadWorkbook(buf, isDemo, stamp) {
  const wb = XLSX.read(buf, { type:'array', cellDates:true });
  const sh = n => { const k = wb.SheetNames.find(x => x.trim().toLowerCase() === n.toLowerCase()); if (!k) throw new Error(`Planilha "${n}" não encontrada.`); return XLSX.utils.sheet_to_json(wb.Sheets[k], { header:1, raw:true, defval:null }); };
  loadAoa(sh('Meta'), sh('Realizado'), isDemo, stamp);
}
function loadAoa(metaAoa, realAoa, isDemo, stamp) {
  const sh = n => n === 'Meta' ? metaAoa : realAoa;
  const meta = C.transformSheet(sh('Meta'), { filtraIndicadorNulo:true });
  const real = C.transformSheet(sh('Realizado'), { filtraIndicadorNulo:false });
  ROWS = C.buildModel(meta, real);
  if (!ROWS.length) throw new Error('Nenhuma linha de meta encontrada.');
  const anos = [...new Set(ROWS.map(r => r.ano))].sort();
  S.ano = anos.includes(String(new Date().getFullYear())) ? String(new Date().getFullYear()) : anos.at(-1);
  S.mes = S.ano === String(new Date().getFullYear()) ? MESES[LIM - 1] : 'Acumulado';
  if (!ROWS.some(r => r.empresa === S.empresa)) S.empresa = ROWS[0].empresa;
  S.cargos = new Set(); S.resp = 'Todos';
  const first = ROWS.find(r => r.empresa === S.empresa && r.ano === S.ano); if (first) S.cargos.add(first.cargo);
  $('#atualizado').textContent = stamp || new Date().toLocaleString('pt-BR');
  $('#demoBanner').classList.toggle('hidden', !isDemo); $('#drop').classList.add('hidden'); $('#main').classList.remove('hidden');
  render();
}
function readFile(f) { const fr = new FileReader(); fr.onload = e => { try { LAST_SRC = { kind:'file', buf:new Uint8Array(e.target.result) }; loadWorkbook(new Uint8Array(e.target.result), false, new Date(f.lastModified).toLocaleString('pt-BR', { timeZone:'America/Sao_Paulo' })); } catch (err) { alert('Erro ao ler o arquivo: ' + err.message); } }; fr.readAsArrayBuffer(f); }
$('#file').onchange = e => e.target.files[0] && readFile(e.target.files[0]);
$('#demo').onclick = () => { LAST_SRC = { kind:'demo' }; const b = atob(window.DEMO_XLSX_B64), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); loadWorkbook(u, true); };
['dragover','dragenter'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); $('#drop').classList.add('over'); }));
document.addEventListener('dragleave', () => $('#drop').classList.remove('over'));
document.addEventListener('drop', e => { e.preventDefault(); $('#drop').classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) readFile(f); });
$$('.tabs button[data-tab]').forEach(b => b.onclick = () => { $$('.tabs button[data-tab]').forEach(x => x.classList.toggle('active', x === b)); $$('.tab').forEach(t => t.classList.add('hidden')); $('#tab-' + b.dataset.tab).classList.remove('hidden'); render(); });
$('#btnPrint').onclick = () => window.print();
// exportar: redimensiona os gráficos para o layout de impressão (1 página A4 paisagem) e restaura depois
// A4 paisagem, margem 8mm → área útil ≈ 281mm de largura; o gráfico do Consolidado é redesenhado no tamanho de impressão (sem distorção)
const MM = 96 / 25.4;
window.addEventListener('beforeprint', () => { const c = charts.chMes3; if (!c || curTab() !== 'print') return;
  const w = Math.round((281 - 10) * MM), h = Math.round(62 * MM); c.canvas.parentElement.style.width = w + 'px'; c.canvas.parentElement.style.height = h + 'px'; c.resize(w, h); });
window.addEventListener('afterprint', () => { const c = charts.chMes3; if (!c) return; c.canvas.parentElement.style.width = ''; c.canvas.parentElement.style.height = ''; c.resize(); });

// ---------- filtros ----------
const curTab = () => $$('.tabs button.active')[0].dataset.tab;
const fAno = r => r.ano === S.ano, fEmp = r => !S.empresa || r.empresa === S.empresa;
const fCargo = r => !S.cargos.size || S.cargos.has(r.cargo), fResp = r => S.resp === 'Todos' || r.responsavel === S.resp;
const fNivel = r => S.nivel === 'Todos' || r.nivel === S.nivel, fIndic = r => S.indicador === 'Todos' || r.indicador === S.indicador;
function base() { const t = curTab(); return ROWS.filter(r => fAno(r) && fEmp(r) && fCargo(r) && (t === 'ating' ? fNivel(r) && fIndic(r) : fResp(r))); }
const uniq = a => [...new Set(a)].sort((x, y) => x.localeCompare(y, 'pt-BR'));

function btnSlicer(el, items, isOn, onClick) { el.innerHTML = ''; items.forEach(it => { const b = document.createElement('button'); b.textContent = it; b.className = isOn(it) ? 'on' : ''; b.onclick = () => { onClick(it); render(); }; el.appendChild(b); }); }
function selSlicer(el, items, val, onCh, multiple) {
  el.innerHTML = ''; const s = document.createElement('select'); if (multiple) { s.multiple = true; s.size = Math.min(5, Math.max(3, items.length)); }
  if (!multiple) items = ['Todos', ...items];
  items.forEach(it => { const o = document.createElement('option'); o.value = o.textContent = it; o.selected = multiple ? val.has(it) : val === it; s.appendChild(o); });
  s.onchange = () => { onCh(multiple ? new Set([...s.selectedOptions].map(o => o.value)) : s.value); render(); }; el.appendChild(s);
  if (multiple) { const c = document.createElement('div'); c.innerHTML = '<small style="color:#888">Ctrl+clique p/ múltiplos · <a href="#">limpar</a></small>'; c.querySelector('a').onclick = e => { e.preventDefault(); S.cargos = new Set(); render(); }; el.appendChild(c); }
}
function renderSlicers(sec) {
  const anos = uniq(ROWS.map(r => r.ano));
  const scope = ROWS.filter(r => fAno(r) && fEmp(r));
  sec.querySelectorAll('[data-sl]').forEach(el => {
    const k = el.dataset.sl;
    if (k === 'ano') btnSlicer(el, anos, a => a === S.ano, a => S.ano = a);
    if (k === 'mes') btnSlicer(el, MESES, m => m === S.mes, m => S.mes = m);
    if (k === 'empresa') btnSlicer(el, uniq([...EMPRESAS, ...ROWS.map(r => r.empresa)]), e => e === S.empresa, e => { S.empresa = (S.empresa === e && curTab() === 'ating') ? null : e; S.cargos = new Set(); S.resp = 'Todos'; });
    if (k === 'cargo') selSlicer(el, uniq(scope.filter(fNivel).map(r => r.cargo)), S.cargos.size === 1 ? [...S.cargos][0] : 'Todos', v => { S.cargos = v === 'Todos' ? new Set() : new Set([v]); S.resp = 'Todos'; });
    if (k === 'resp') selSlicer(el, uniq(scope.filter(fCargo).map(r => r.responsavel)), S.resp, v => S.resp = v);
    if (k === 'nivel') selSlicer(el, Object.keys(C.NIVEL_ORD), S.nivel, v => S.nivel = v);
    if (k === 'indicador') selSlicer(el, uniq(scope.filter(fCargo).map(r => r.indicador)), S.indicador, v => S.indicador = v);
  });
}

// ---------- tabela ----------
const FAROL = { ok:'✓', falha:'✗', parcial:'!', sem:'—' };
function table(el, rows) {
  rows = [...rows].sort((a, b) => b.peso - a.peso || a.indicador.localeCompare(b.indicador));
  const multi = new Set(rows.map(r => r.cargo + r.responsavel)).size > 1;
  let h = `<tr><th>Indicador</th>${multi ? '<th>Responsável</th>' : ''}<th>Peso</th><th>Und.</th><th>Sinal</th><th>Limite</th><th>Desf.</th><th>Meta</th><th>Realizado</th><th>∆</th><th>∆(%)</th><th>Ind.(%)</th><th>Farol</th><th>Book(%)</th><th>Desf.(%)</th></tr>`;
  for (const r of rows) {
    const f = C.farol(r);
    h += `<tr class="${r.semDado ? 'semdado' : ''}"><td class="l" title="${r.indicador}">${r.indicador}</td>${multi ? `<td class="l">${r.cargo} – ${r.responsavel}</td>` : ''}<td>${r.peso}</td><td class="c">${r.unidade}</td><td class="c">${r.sinal}</td>
      <td>${pct(r.limite)}</td><td>${pct(r.desafio || 0, 1)}</td><td>${fmtV(r.meta, r)}</td><td>${r.semDado ? '<i>sem dados</i>' : fmtV(r.realizado, r)}</td>
      <td class="fc ${f}">${r.semDado ? '—' : fmtV(r.delta, r)}</td><td class="fc ${f}">${r.semDado ? '—' : pct(r.deltaPct)}</td>
      <td>${pct(r.atgInd)}</td><td class="c"><span class="f ${f}" title="${f === 'sem' ? 'sem dados' : pct(r.atgInd)}">${FAROL[f]}</span></td><td>${pct(r.atgBook3)}</td><td>${r.atgDesafio === null ? '' : pct(r.atgDesafio)}</td></tr>`;
  }
  if (!rows.length) h += `<tr><td class="l" colspan="14">Nenhum indicador para os filtros selecionados.</td></tr>`;
  el.innerHTML = h;
}
// grupos = um book (Empresa-Cargo-Responsável-Ano); com vários selecionados, média dos books
function bookSeries(rows) {
  const out = MESES.map((m, i) => {
    const mr = rows.filter(r => r.mesOrd === i + 1); const g = new Map();
    mr.forEach(r => { const k = r.c5 + r.responsavel; const o = g.get(k) || { b:0, d:0, p:0, real:false }; o.b += r.atgBook3; o.d += r.atgBook4; o.p += r.peso; o.real ||= !r.semDado; g.set(k, o); });
    const gs = [...g.values()]; if (!gs.length || !gs.some(x => x.real)) return null;
    const avg = f => gs.reduce((a, x) => a + f(x), 0) / gs.length;
    return { book: avg(x => x.b), desf: avg(x => x.d), farol: avg(x => x.p ? x.b / (x.p / 100) : 0) };
  });
  // oculta meses futuros (após o mês anterior ao atual) no ano corrente, como "Atg. Book (Gráfico)"
  if (S.ano === String(new Date().getFullYear())) for (let i = LIM; i < 12; i++) out[i] = null;
  return out;
}
const corAtg = v => v >= 1 ? '#007024' : v >= 0.7 ? '#f2b400' : '#a30000';
function chart(id, cfg) { if (charts[id]) charts[id].destroy();
  if (cfg.type !== 'doughnut') { const sc = cfg.options.scales = cfg.options.scales || {}; sc.x = { ...(sc.x || {}), grid:{ display:false } };
    sc.y = { ...(sc.y || {}), border:{ display:false }, grid:{ color:'rgba(39,46,50,.07)', drawTicks:false } }; } charts[id] = new Chart(document.getElementById(id), cfg); }
const pctTick = { callback: v => nf(0).format(v * 100) + '%' };
function atgChart(id, rows) {
  // mesmo modelo do Painel: barras na cor do farol, linha Meta 100% grafite tracejada, rótulos, "sem dados"
  const s = bookSeries(rows), st = s.map(x => !x || x.book === null || x.book === undefined ? 'sem' : x.farol >= 1 ? 'ok' : x.farol >= 0.7 ? 'parcial' : 'falha');
  chart(id, { type:'bar', data:{ labels:MESES, datasets:[
    { type:'line', label:'Meta 100%', data:MESES.map(() => 1), borderColor:'#272e32', pointRadius:0, borderWidth:2.5 },
    { label:'Atingimento Book', data:s.map(x => x && x.book), backgroundColor:st.map(k => FAROL_COR[k]), borderRadius:6, borderSkipped:false, maxBarThickness:46 } ] },
    options:{ maintainAspectRatio:false, animation:false, layout:{ padding:{ top:18, right:8, left:4 } },
      scales:{ x:{ ...AX, ticks:{ padding:6, maxRotation:0, autoSkip:false } }, y:{ min:0, suggestedMax:1.1, border:{ display:false }, grid:GRID, ticks:{ ...pctTick, padding:8 } } },
      plugins:{ legend:legendaFarol([{ text:'Meta 100%', fillStyle:'#272e32', strokeStyle:'#272e32', pointStyle:'line', lineWidth:3 }]),
        semdado:s.map((x, i) => st[i] === 'sem' && (i === 12 || mesEsperado(i))),
        tooltip:{ padding:10, callbacks:{ label:c => c.dataset.label === 'Meta 100%' ? ' Meta: 100%' : ` ${c.dataset.label}: ${pct(c.raw, 1)} · ${FAROL_TXT[st[c.dataIndex]]}` } } } },
    plugins:[labelPlugin, semDadoPlugin] });
}
const labelPlugin = { id:'lbl', afterDatasetsDraw(ch) { const ctx = ch.ctx; ctx.save(); ctx.font = "600 11px 'Poppins', Arial, sans-serif"; ctx.fillStyle = '#3e3e3e'; ctx.textAlign = 'center';
  ch.data.datasets.forEach((ds, di) => { if (ds.type === 'line' || ch.options.indexAxis === 'y') return; ch.getDatasetMeta(di).data.forEach((el, i) => { const v = ds.data[i]; if (v === null || v === undefined || (el.width && el.width < 30)) return; { const t = ds.fmt ? ds.fmt(v) : pct(v, 1); ctx.strokeStyle = '#fff'; ctx.lineWidth = 3; ctx.lineJoin = 'round'; ctx.strokeText(t, el.x, el.y - 5); ctx.fillText(t, el.x, el.y - 5); } }); });
  if (ch.options.indexAxis === 'y') { ctx.textAlign = 'left'; ch.data.datasets.forEach((ds, di) => ch.getDatasetMeta(di).data.forEach((el, i) => ctx.fillText(pct(ds.data[i], 1), el.x + 4, el.y + 4))); }
  ctx.restore(); } };

// ---------- gráficos do indicador (cores = farol do calc.js, considerando ↑/↓ e Limite) ----------
const FAROL_COR = { ok:'#1f9d55', parcial:'#f0a202', falha:'#a30000', sem:'#cfd4d8' };
const FAROL_TXT = { ok:'Atingido', parcial:'Parcial', falha:'Não atingido', sem:'Sem dados' };
// atingimento de um ponto (mês ou acumulado) agregando as linhas filtradas, com as mesmas regras do modelo
function atgPonto(x, ref) { return x.real === null || x.meta === null || !ref ? null : C.atgIndicador(ref.sinal, ref.limite, x.meta, x.real); }
function statusPonto(x, ref) { const a = atgPonto(x, ref); return a === null ? 'sem' : C.farol({ semDado:false, atgInd:a }); }
// meses com dado esperado: anos passados = todos; ano corrente = até o mês anterior
const mesEsperado = i => S.ano !== String(new Date().getFullYear()) || i < LIM;
function desafioAlvo(x, ref) { const d = ref && C.num(ref.desafio); if (!d || x.meta === null) return null; return ref.sinal === '↓' ? x.meta * (1 - d) : x.meta * (1 + d); }
const GRID = { color:'rgba(39,46,50,.07)', drawTicks:false }, AX = { grid:{ display:false }, ticks:{ padding:6 } };
const legendaFarol = (extra) => ({ position:'bottom', labels:{ usePointStyle:true, boxWidth:10, padding:14, generateLabels: () => [
  // espaço maior entre o grupo de linhas (Meta/Desafio) e o grupo de colunas (farol)
  ...extra.map((e, i) => i === extra.length - 1 ? { ...e, text:e.text + '\u2003\u2003\u2003' } : e), ...['ok', 'parcial', 'falha', 'sem'].map(k => ({ text:FAROL_TXT[k], fillStyle:FAROL_COR[k], strokeStyle:FAROL_COR[k], pointStyle:'rectRounded', lineWidth:0 })) ] } });
const semDadoPlugin = { id:'semdado', afterDatasetsDraw(ch) { const o = ch.options.plugins.semdado; if (!o) return; const ctx = ch.ctx, y = ch.scales.y.getPixelForValue(ch.scales.y.min);
  ctx.save(); ctx.font = "italic 10px 'Poppins', Arial, sans-serif"; ctx.fillStyle = '#8a959e'; ctx.textAlign = 'center';
  o.forEach((on, i) => { if (on) ctx.fillText('sem dados', ch.scales.x.getPixelForValue(i), y - 6); }); ctx.restore(); } };
function indChart(id, M12, ref) {
  const st = M12.map(x => statusPonto(x, ref)), des = M12.map(x => desafioAlvo(x, ref)), temDes = des.some(v => v !== null);
  const ds = [
    { type:'line', label:'Meta', data:M12.map(x => x.meta), borderColor:'#272e32', backgroundColor:'#272e32', pointRadius:3.5, pointHoverRadius:5, borderWidth:2.5, tension:0, order:0 },
    { label:'Realizado', data:M12.map(x => x.real), backgroundColor:st.map(k => FAROL_COR[k]), borderRadius:6, borderSkipped:false, maxBarThickness:46, order:2, fmt:v => fmtV(v, ref) } ];
  if (temDes) ds.splice(1, 0, { type:'line', label:'Desafio', data:des, borderColor:'#5f6c77', borderDash:[6, 5], borderWidth:1.8, pointRadius:0, tension:0, order:1 });
  const extra = [{ text:'Meta', fillStyle:'#272e32', strokeStyle:'#272e32', pointStyle:'line', lineWidth:3 }];
  if (temDes) extra.push({ text:'Desafio', fillStyle:'#5f6c77', strokeStyle:'#5f6c77', pointStyle:'dash', lineWidth:2, lineDash:[6, 5] });
  chart(id, { type:'bar', data:{ labels:MESES.slice(0, 12), datasets:ds },
    options:{ maintainAspectRatio:false, animation:false, layout:{ padding:{ top:18, right:8, left:4 } },
      scales:{ x:AX, y:{ beginAtZero:true, grace:'12%', border:{ display:false }, grid:GRID, ticks:{ padding:8, callback:v => fmtV(v, ref) } } },
      plugins:{ legend:legendaFarol(extra), semdado:M12.map((x, i) => x.real === null && x.meta !== null && mesEsperado(i)),
        tooltip:{ mode:'index', intersect:false, padding:10, usePointStyle:true,
          itemSort:(a, b) => ['Desafio', 'Meta', 'Realizado'].indexOf(a.dataset.label) - ['Desafio', 'Meta', 'Realizado'].indexOf(b.dataset.label),
          callbacks:{ label:c => ` ${c.dataset.label}: ${c.raw === null || c.raw === undefined ? 'sem dados' : fmtV(c.raw, ref)}`,
            labelColor:c => { const k = c.dataset.label === 'Realizado' ? FAROL_COR[st[c.dataIndex]] : c.dataset.borderColor; return { borderColor:k, backgroundColor:k }; },
            labelPointStyle:c => ({ pointStyle:c.dataset.label === 'Realizado' ? 'rectRounded' : c.dataset.label === 'Desafio' ? 'dash' : 'line', rotation:0 }),
            footer:items => { const i = items[0]?.dataIndex; if (i === undefined) return ''; const x = M12[i];
              return x.real === null ? '' : `${FAROL_TXT[st[i]]} · Atingimento ${pct(atgPonto(x, ref), 0)}`; } } } } },
    plugins:[labelPlugin, semDadoPlugin] });
}
// Atingimento do INDICADOR selecionado (como no Power BI), mês a mês + Acumulado
function indAtgChart(id, P13, ref) {
  const at = P13.map(x => atgPonto(x, ref)), st = P13.map(x => statusPonto(x, ref));
  chart(id, { type:'bar', data:{ labels:MESES, datasets:[
    { type:'line', label:'Meta 100%', data:MESES.map(() => 1), borderColor:'#272e32', pointRadius:0, borderWidth:2.5 },
    { label:'Atingimento', data:at, backgroundColor:st.map(k => FAROL_COR[k]), borderRadius:6, borderSkipped:false, maxBarThickness:46 } ] },
    options:{ maintainAspectRatio:false, animation:false, layout:{ padding:{ top:18, right:8, left:4 } },
      scales:{ x:{ ...AX, ticks:{ padding:6, maxRotation:0, autoSkip:false } }, y:{ min:0, suggestedMax:1.1, border:{ display:false }, grid:GRID, ticks:{ ...pctTick, padding:8 } } },
      plugins:{ legend:legendaFarol([{ text:'Meta 100%', fillStyle:'#272e32', strokeStyle:'#272e32', pointStyle:'line', lineWidth:3 }]),
        semdado:P13.map((x, i) => x.real === null && x.meta !== null && (i === 12 || mesEsperado(i))),
        tooltip:{ padding:10, callbacks:{ label:c => c.dataset.label === 'Atingimento' ? ` Atingimento: ${pct(c.raw, 1)} · ${FAROL_TXT[st[c.dataIndex]]}` : ' Meta: 100%' } } } },
    plugins:[labelPlugin, semDadoPlugin] });
}

function desChart(id, rows) { const s = bookSeries(rows); desBars(id, s.map(x => x ? x.desf : null)); }
// Desafio de UM indicador (selInd): mesmas regras do calc.js (C.atgDesafio, respeita ↑/↓; Realizado em branco = sem dados)
function desIndChart(id, P13, ref) {
  const d = ref && C.num(ref.desafio);
  if (!d) return desBars(id, P13.map(() => null), 'Indicador sem desafio');
  desBars(id, P13.map(x => x.real === null || x.meta === null ? null : C.atgDesafio(ref.sinal, d, x.meta, x.real)), null, P13.map((x, i) => x.real === null && x.meta !== null && (i === 12 || mesEsperado(i))));
}
const vazioPlugin = { id:'vazio', afterDraw(ch) { const t = ch.options.plugins.vazio; if (!t) return; const a = ch.chartArea, ctx = ch.ctx; ctx.save();
  ctx.font = "600 14px 'Poppins', Arial, sans-serif"; ctx.fillStyle = '#8a959e'; ctx.textAlign = 'center'; ctx.fillText(t, (a.left + a.right) / 2, (a.top + a.bottom) / 2); ctx.restore(); } };
function desBars(id, v, vazio, sd) {
  // Atingimento do DESAFIO: barras na cor do farol, apenas a linha tracejada do Desafio (100%)
  const st = v.map(x => x === null || x === undefined ? 'sem' : x >= 1 ? 'ok' : x >= 0.7 ? 'parcial' : 'falha');
  chart(id, { type:'bar', data:{ labels:MESES, datasets:[
    { type:'line', label:'Desafio 100%', data:MESES.map(() => 1), borderColor:'#5f6c77', borderDash:[6, 5], pointRadius:0, borderWidth:1.8 },
    { label:'Atingimento Desafio', data:v, backgroundColor:st.map(k => FAROL_COR[k]), borderRadius:6, borderSkipped:false, maxBarThickness:46 } ] },
    options:{ maintainAspectRatio:false, animation:false, layout:{ padding:{ top:18, right:8, left:4 } },
      scales:{ x:{ ...AX, ticks:{ padding:6, maxRotation:0, autoSkip:false } }, y:{ min:0, suggestedMax:1.1, border:{ display:false }, grid:GRID, ticks:{ ...pctTick, padding:8 } } },
      plugins:{ legend:legendaFarol([{ text:'Desafio 100%', fillStyle:'#5f6c77', strokeStyle:'#5f6c77', pointStyle:'dash', lineWidth:2, lineDash:[6, 5] }]), vazio,
        semdado:vazio ? null : sd || v.map((x, i) => st[i] === 'sem' && (i === 12 || mesEsperado(i))),
        tooltip:{ padding:10, callbacks:{ label:c => c.dataset.label === 'Desafio 100%' ? ' Desafio: 100%' : ` Atingimento Desafio: ${pct(c.raw, 1)} · ${FAROL_TXT[st[c.dataIndex]]}` } } } },
    plugins:[labelPlugin, semDadoPlugin, vazioPlugin] });
}

// ---------- páginas ----------
function renderPainel() {
  const sec = $('#tab-painel'); renderSlicers(sec);
  const b = base(), cur = b.filter(r => r.mes === S.mes);
  { const NM = { Jan:'Janeiro', Fev:'Fevereiro', Mar:'Março', Abr:'Abril', Mai:'Maio', Jun:'Junho', Jul:'Julho', Ago:'Agosto', Set:'Setembro', Out:'Outubro', Nov:'Novembro', Dez:'Dezembro' };
    $('#bookTitle').textContent = S.mes === 'Acumulado' ? 'Book Acumulado' : 'Book de ' + (NM[S.mes] || S.mes); }
  table($('#tblPainel'), cur);
  const g = new Set(cur.map(r => r.c5 + r.responsavel)).size || 1;
  $('#totBook').textContent = pct(cur.reduce((a, r) => a + r.atgBook3, 0) / g, 1);
  $('#totDesf').textContent = pct(cur.reduce((a, r) => a + r.atgBook4, 0) / g, 1);
  const inds = uniq(b.map(r => r.indicador)); if (!inds.includes(S.selInd)) S.selInd = [...b].sort((x, y) => y.peso - x.peso)[0]?.indicador || null;
  const si = $('#selInd'); si.innerHTML = inds.map(i => `<option ${i === S.selInd ? 'selected' : ''}>${i}</option>`).join(''); si.onchange = () => { S.selInd = si.value; render(); };
  const ir = b.filter(r => r.indicador === S.selInd);
  const byM = m => { const x = ir.filter(r => r.mes === m); return { meta:x.length ? x.at(-1).meta : null, real:x.some(r => !r.semDado) ? x.reduce((a, r) => a + (r.realizado ?? 0), 0) : null, r:x[0] }; };
  const M12 = MESES.slice(0, 12).map(byM), ref = ir[0];
  indChart('chInd', M12, ref);
  const ac = byM('Acumulado'), mx = Math.max(ac.real || 0, ac.meta || 0) * 1.1 || 1;
  gauge($('#gauge'), ac, ref);
  // cabeçalho de KPIs (estilo ACT / PLAN / Δ)
  const tb = cur.reduce((s, r) => s + r.atgBook3, 0) / g, td = cur.reduce((s, r) => s + r.atgBook4, 0) / g;
  kpiDelta($('#kDeltaBook'), cur.length ? td - tb : null, v => (v > 0 ? '+' : '') + nf(1).format(v * 100) + ' p.p.', null);
  $('#kReal').textContent = ac.real === null ? 'sem dados' : fmtV(ac.real, ref); $('#kMeta').textContent = ac.meta === null ? '—' : fmtV(ac.meta, ref);
  const dlt = ac.real === null || ac.meta === null ? null : ac.real - ac.meta, aP = atgPonto(ac, ref);
  kpiDelta($('#kDelta'), dlt, v => `${v > 0 ? '+' : ''}${fmtV(v, ref)} | ${pct(aP, 1)}`, statusPonto(ac, ref), ref && ref.sinal === '↓');
  $$('#atgToggle button').forEach(bt => { bt.classList.toggle('on', bt.dataset.v === S.atgView); bt.onclick = () => { S.atgView = bt.dataset.v; render(); }; });
  $('#atgSub').textContent = S.atgView === 'book' ? '· Book (todos os indicadores)' : S.atgView === 'des' ? '· Desafio do Book (todos os indicadores)' : S.atgView === 'desind' ? (S.selInd ? '· Desafio · ' + S.selInd : '') : (S.selInd ? '· ' + S.selInd : '');
  if (S.atgView === 'book') atgChart('chAtg', b); else if (S.atgView === 'des') desChart('chAtg', b); else if (S.atgView === 'desind') desIndChart('chAtg', [...M12, ac], ref); else indAtgChart('chAtg', [...M12, ac], ref);
}
// velocímetro (SVG semicircular): preenchimento na cor do farol, marcador da meta
function gauge(el, ac, ref) {
  const st = statusPonto(ac, ref), cor = st === 'sem' ? '#cfd4d8' : FAROL_COR[st], at = atgPonto(ac, ref);
  const dv = desafioAlvo(ac, ref), mx = Math.max(Math.abs(ac.real || 0), Math.abs(ac.meta || 0), Math.abs(dv || 0)) * 1.2 || 1;
  const f = ac.real === null ? 0 : Math.max(0, Math.min(1, ac.real / mx)), fm = ac.meta === null ? null : Math.max(0, Math.min(1, ac.meta / mx));
  const P = t => [100 - 80 * Math.cos(Math.PI * t), 100 - 80 * Math.sin(Math.PI * t)], arc = 'M20 100 A80 80 0 0 1 180 100';
  let mk = ''; if (fm !== null) { const [x1, y1] = [100 - 64 * Math.cos(Math.PI * fm), 100 - 64 * Math.sin(Math.PI * fm)], [x2, y2] = [100 - 96 * Math.cos(Math.PI * fm), 100 - 96 * Math.sin(Math.PI * fm)];
    const [tx, ty] = [100 - 106 * Math.cos(Math.PI * fm), 100 - 106 * Math.sin(Math.PI * fm)];
    mk = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#272e32" stroke-width="3" stroke-linecap="round"/><text x="${tx}" y="${ty}" class="gmk" text-anchor="middle" dominant-baseline="middle">Meta</text>`; }
  // marcador tracejado do Desafio (mesmo estilo da linha Desafio do gráfico ao lado)
  if (dv !== null) { const fd = Math.max(0, Math.min(1, dv / mx)), c = Math.cos(Math.PI * fd), s = Math.sin(Math.PI * fd);
    const near = fm !== null && Math.abs(fd - fm) < 0.09, rr = near ? 118 : 106, dy = near ? (fd < fm ? 6 : -6) : 0;
    mk += `<line x1="${100 - 62 * c}" y1="${100 - 62 * s}" x2="${100 - 98 * c}" y2="${100 - 98 * s}" stroke="#5f6c77" stroke-width="2.2" stroke-dasharray="4 3"/><text x="${100 - rr * c}" y="${100 - rr * s + dy}" class="gmk gdes" text-anchor="middle" dominant-baseline="middle">Desafio</text>`; }
  // atingimento do desafio no acumulado (mesma regra do gráfico "Desafio Indicador")
  const dd = ref && C.num(ref.desafio), ad = dd && ac.real !== null && ac.meta !== null ? C.atgDesafio(ref.sinal, dd, ac.meta, ac.real) : null;
  const adCor = ad === null ? '#8a959e' : FAROL_COR[ad >= 1 ? 'ok' : ad >= 0.7 ? 'parcial' : 'falha'];
  const sem = ac.real === null, gv = sem ? 'sem dados' : fmtV(ac.real, ref), gfs = Math.max(14, Math.min(25, 190 / Math.max(gv.length, 1)));
  el.innerHTML = `<svg viewBox="-12 -14 224 124" preserveAspectRatio="xMidYMid meet" role="img" aria-label="Acumulado">
    <path d="${arc}" fill="none" stroke="#eef0f1" stroke-width="18" stroke-linecap="round"/>
    ${f > 0 ? `<path d="${arc}" fill="none" stroke="${cor}" stroke-width="18" stroke-linecap="round" pathLength="100" stroke-dasharray="${(f * 100).toFixed(2)} 200"/>` : ''}
    ${mk}<text x="100" y="86" text-anchor="middle" class="gval" font-size="${gfs.toFixed(1)}" fill="${sem ? '#8a959e' : cor}">${gv}</text>
    <text x="100" y="102" text-anchor="middle" class="gst">${FAROL_TXT[st]}</text></svg>
    <div class="glbl${dv !== null ? ' g4' : ''}">${dv !== null ? `<div><small>Desafio</small><b>${fmtV(dv, ref)}</b></div><div><small>Atingimento Desafio</small><b style="color:${adCor}">${ad === null ? '—' : pct(ad, 1)}</b></div>` : ''}<div><small>Meta</small><b>${ac.meta === null ? '—' : fmtV(ac.meta, ref)}</b></div><div><small>Atingimento</small><b style="color:${at === null ? '#8a959e' : cor}">${at === null ? '—' : pct(at, 1)}</b></div></div>`;
}
// KPIs Atingimento Book / Desafio / Δ (mesma regra do Painel) para um conjunto de linhas
function kpiBook(p, rows) { const g = new Set(rows.map(r => r.c5 + r.responsavel)).size || 1;
  const tb = rows.reduce((s, r) => s + r.atgBook3, 0) / g, td = rows.reduce((s, r) => s + r.atgBook4, 0) / g;
  $('#' + p + 'Book').textContent = pct(tb, 1); $('#' + p + 'Desf').textContent = pct(td, 1);
  kpiDelta($('#' + p + 'Delta'), rows.length ? td - tb : null, v => (v > 0 ? '+' : '') + nf(1).format(v * 100) + ' p.p.', null); }
// Δ com seta: cor pelo farol (ou sinal, quando não há farol)
function kpiDelta(el, v, fmt, st, menorMelhor) {
  if (v === null || v === undefined || isNaN(v)) { el.textContent = '—'; el.style.color = '#8a959e'; return; }
  const up = v >= 0, arrow = Math.abs(v) < 1e-12 ? '' : (up ? ' ▲' : ' ▼');
  el.textContent = fmt(v) + arrow;
  el.style.color = st ? FAROL_COR[st] === '#cfd4d8' ? '#8a959e' : FAROL_COR[st] : ((up !== !!menorMelhor) ? FAROL_COR.ok : FAROL_COR.falha);
}
function renderAting() {
  const sec = $('#tab-ating'); renderSlicers(sec);
  const cur = base().filter(r => r.mes === S.mes), g = new Map();
  cur.forEach(r => { const k = r.empresa + ' - ' + r.responsavel + (cur.some(x => x.empresa === r.empresa && x.responsavel === r.responsavel && x.cargo !== r.cargo) ? ' (' + r.cargo + ')' : ''); const o = g.get(k) || { b:0, d:0 }; o.b += r.atgBook3; o.d += r.atgBook4; g.set(k, o); });
  const ent = [...g.entries()].sort((a, b) => b[1].b - a[1].b);
  $('#colabWrap').style.height = Math.max(220, ent.length * 44 + 60) + 'px';
  chart('chColab', { type:'bar', data:{ labels:ent.map(e => e[0]), datasets:[
    { label:'Atingimento Book', data:ent.map(e => e[1].b), backgroundColor:'#8a959e' }, { label:'Atingimento Desafio', data:ent.map(e => e[1].d), backgroundColor:'#007024' } ] },
    options:{ indexAxis:'y', maintainAspectRatio:false, animation:false, scales:{ x:{ min:0, suggestedMax:1.15, ticks:pctTick } }, plugins:{ legend:{ position:'top' } } }, plugins:[labelPlugin] });
  atgChart('chMes2', base());
}
function renderPrint() {
  const sec = $('#tab-print'); renderSlicers(sec); const b = base();
  $('#printHead').textContent = `· Ano ${S.ano} · ${S.empresa || 'Todas'} · ${S.cargos.size ? [...S.cargos].join(', ') : 'Todos os cargos'}`;
  $('#pt1').textContent = 'Mês: ' + S.mes;
  table($('#tblP1'), b.filter(r => r.mes === S.mes)); table($('#tblP2'), b.filter(r => r.mes === 'Acumulado'));
  kpiBook('kP1', b.filter(r => r.mes === S.mes)); kpiBook('kP2', b.filter(r => r.mes === 'Acumulado'));
  atgChart('chMes3', b);
}
function render() { if (!ROWS.length) return; ({ painel:renderPainel, ating:renderAting, print:renderPrint })[curTab()](); }
if (location.hash === '#demo' || !window.REAL_DATA) $('#demo').click();
else if (window.REAL_DATA) (LAST_SRC = { kind:'real' }, loadAoa(REAL_DATA.meta, REAL_DATA.realizado, false, REAL_DATA.mtime));

// ---------- botão Atualizar: recarrega a mesma fonte (demo, arquivo ou dados embutidos) mantendo os filtros ----------
// Produção: refazer o fetch de /api/book?refresh=1 (a API ignora o cache de ~10 min, com limite de 1x/min por usuário).
$('#btnRefresh').onclick = () => {
  const btn = $('#btnRefresh'); if (btn.disabled || !LAST_SRC) return; btn.disabled = true; btn.classList.add('spin');
  const keep = { ...S, cargos:new Set(S.cargos) }, tab = curTab();
  setTimeout(() => {
    try {
      const now = new Date().toLocaleString('pt-BR', { timeZone:'America/Sao_Paulo' });
      if (LAST_SRC.kind === 'demo') { const b = atob(window.DEMO_XLSX_B64), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); loadWorkbook(u, true, now); }
      else if (LAST_SRC.kind === 'file') loadWorkbook(LAST_SRC.buf, false, now);
      else loadAoa(REAL_DATA.meta, REAL_DATA.realizado, false, now);
      // restaura filtros que ainda existem nos dados
      if (ROWS.some(r => r.ano === keep.ano)) S.ano = keep.ano;
      if (keep.empresa === null || ROWS.some(r => r.empresa === keep.empresa)) S.empresa = keep.empresa;
      S.mes = keep.mes; S.cargos = new Set([...keep.cargos].filter(c => ROWS.some(r => r.cargo === c)));
      S.resp = keep.resp; S.nivel = keep.nivel; S.indicador = keep.indicador; S.selInd = keep.selInd; S.atgView = keep.atgView;
      render();
    } catch (err) { alert('Erro ao atualizar: ' + err.message); }
    setTimeout(() => { btn.disabled = false; btn.classList.remove('spin'); }, 450);
  }, 60);
};

// ---------- topo congelado: mede a barra de título e decide se os filtros também ficam fixos (mobile: até ~40% da tela) ----------
function stickyLayout() {
  const top = document.querySelector('.top'), th = top ? top.offsetHeight : 0; document.documentElement.style.setProperty('--toph', th + 'px');
  const f = document.querySelector('.tab:not(.hidden)>.panel.filters'); if (!f) return;
  document.body.classList.toggle('nostick', th + f.offsetHeight > innerHeight * 0.4 && innerWidth < 900);
  f.classList.toggle('stuck', !document.body.classList.contains('nostick') && f.getBoundingClientRect().top <= th + 1 && scrollY > 0);
}
addEventListener('scroll', stickyLayout, { passive:true }); addEventListener('resize', stickyLayout);
{ const r0 = render; render = function () { r0.apply(this, arguments); stickyLayout(); }; }
setTimeout(stickyLayout, 0);

// sincroniza a altura da base da sidebar com o rodapé (inclui quebra em 2 linhas)
(() => { const f = document.querySelector('footer'); if (!f) return; const set = () => document.documentElement.style.setProperty('--foot-h', f.getBoundingClientRect().height + 'px'); set(); if (window.ResizeObserver) new ResizeObserver(set).observe(f); window.addEventListener('resize', set); })();
