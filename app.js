/* Book de Metas – app */
const C = BookCalc, MESES = C.MESES, EMPRESAS = ['CSC','Ferguminas','MIB','MIG','MML','SM Metais','SM Mineração'];
let ROWS = [], S = { ano:null, mes:null, empresa:'CSC', cargos:new Set(), resp:'Todos', nivel:'Todos', indicador:'Todos', selInd:null };
const charts = {};
if (window.Chart) { Chart.defaults.font.family = "'Poppins', Arial, sans-serif"; Chart.defaults.color = '#5f6c77'; Chart.defaults.borderColor = '#eef0f1'; }
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const nf = (d) => new Intl.NumberFormat('pt-BR', { minimumFractionDigits:d, maximumFractionDigits:d });
const pct = (v, d = 0) => v === null || v === undefined ? '—' : nf(d).format((Math.abs(v) < 5e-7 ? 0 : v) * 100).replace(/^-(0(,0+)?)$/, '$1') + '%';
const fmtV = (v, r) => v === null || v === undefined ? '' : nf(r && r.formato === 'TRES CASAS' ? 3 : 2).format(v);
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
function readFile(f) { const fr = new FileReader(); fr.onload = e => { try { loadWorkbook(new Uint8Array(e.target.result), false, new Date(f.lastModified).toLocaleString('pt-BR', { timeZone:'America/Sao_Paulo' })); } catch (err) { alert('Erro ao ler o arquivo: ' + err.message); } }; fr.readAsArrayBuffer(f); }
$('#file').onchange = e => e.target.files[0] && readFile(e.target.files[0]);
$('#demo').onclick = () => { const b = atob(window.DEMO_XLSX_B64), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); loadWorkbook(u, true); };
['dragover','dragenter'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); $('#drop').classList.add('over'); }));
document.addEventListener('dragleave', () => $('#drop').classList.remove('over'));
document.addEventListener('drop', e => { e.preventDefault(); $('#drop').classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) readFile(f); });
$$('.tabs button[data-tab]').forEach(b => b.onclick = () => { $$('.tabs button[data-tab]').forEach(x => x.classList.toggle('active', x === b)); $$('.tab').forEach(t => t.classList.add('hidden')); $('#tab-' + b.dataset.tab).classList.remove('hidden'); render(); });
$('#btnPrint').onclick = () => window.print();

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
    if (k === 'cargo') selSlicer(el, uniq(scope.filter(fNivel).map(r => r.cargo)), S.cargos, v => { S.cargos = v; S.resp = 'Todos'; }, true);
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
      <td>${pct(r.limite)}</td><td>${pct(r.desafio || 0, 1)}</td><td>${fmtV(r.meta, r)}</td><td>${r.semDado ? '<i>sem dado</i>' : fmtV(r.realizado, r)}</td>
      <td class="${r.delta < 0 ? 'neg' : ''}">${r.semDado ? '—' : fmtV(r.delta, r)}</td><td class="${r.deltaPct < 0 ? 'neg' : ''}">${r.semDado ? '—' : pct(r.deltaPct)}</td>
      <td>${pct(r.atgInd)}</td><td class="c"><span class="f ${f}" title="${f === 'sem' ? 'sem dado' : pct(r.atgInd)}">${FAROL[f]}</span></td><td>${pct(r.atgBook3)}</td><td>${r.atgDesafio === null ? '' : pct(r.atgDesafio)}</td></tr>`;
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
function chart(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart(document.getElementById(id), cfg); }
const pctTick = { callback: v => nf(0).format(v * 100) + '%' };
function atgChart(id, rows) {
  const s = bookSeries(rows);
  chart(id, { type:'bar', data:{ labels:MESES, datasets:[
    { type:'line', label:'Meta 100%', data:MESES.map(() => 1), borderColor:'#272e32', borderDash:[5, 4], pointRadius:0, borderWidth:1.5 },
    { label:'Atingimento Book', data:s.map(x => x && x.book), backgroundColor:s.map(x => x ? corAtg(x.farol) : '#cfd4d8') } ] },
    options:{ maintainAspectRatio:false, animation:false, scales:{ y:{ min:0, suggestedMax:1.1, ticks:pctTick } },
      plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label:c => c.dataset.label + ': ' + pct(c.raw, 1) } } } },
    plugins:[labelPlugin] });
}
const labelPlugin = { id:'lbl', afterDatasetsDraw(ch) { const ctx = ch.ctx; ctx.save(); ctx.font = '600 11px Segoe UI,Arial'; ctx.fillStyle = '#3e3e3e'; ctx.textAlign = 'center';
  ch.data.datasets.forEach((ds, di) => { if (ds.type === 'line' || ch.options.indexAxis === 'y') return; ch.getDatasetMeta(di).data.forEach((el, i) => { const v = ds.data[i]; if (v === null || v === undefined) return; ctx.fillText(ds.fmt ? ds.fmt(v) : pct(v, 1), el.x, el.y - 4); }); });
  if (ch.options.indexAxis === 'y') { ctx.textAlign = 'left'; ch.data.datasets.forEach((ds, di) => ch.getDatasetMeta(di).data.forEach((el, i) => ctx.fillText(pct(ds.data[i], 1), el.x + 4, el.y + 4))); }
  ctx.restore(); } };

// ---------- páginas ----------
function renderPainel() {
  const sec = $('#tab-painel'); renderSlicers(sec);
  const b = base(), cur = b.filter(r => r.mes === S.mes);
  table($('#tblPainel'), cur);
  const g = new Set(cur.map(r => r.c5 + r.responsavel)).size || 1;
  $('#totBook').textContent = pct(cur.reduce((a, r) => a + r.atgBook3, 0) / g, 1);
  $('#totDesf').textContent = pct(cur.reduce((a, r) => a + r.atgBook4, 0) / g, 1);
  const inds = uniq(b.map(r => r.indicador)); if (!inds.includes(S.selInd)) S.selInd = [...b].sort((x, y) => y.peso - x.peso)[0]?.indicador || null;
  const si = $('#selInd'); si.innerHTML = inds.map(i => `<option ${i === S.selInd ? 'selected' : ''}>${i}</option>`).join(''); si.onchange = () => { S.selInd = si.value; render(); };
  const ir = b.filter(r => r.indicador === S.selInd);
  const byM = m => { const x = ir.filter(r => r.mes === m); return { meta:x.length ? x.at(-1).meta : null, real:x.some(r => !r.semDado) ? x.reduce((a, r) => a + (r.realizado ?? 0), 0) : null, r:x[0] }; };
  const M12 = MESES.slice(0, 12).map(byM), ref = ir[0];
  chart('chInd', { type:'bar', data:{ labels:MESES.slice(0, 12), datasets:[
    { type:'line', label:'Meta', data:M12.map(x => x.meta), borderColor:'#272e32', backgroundColor:'#272e32', pointRadius:3, borderWidth:2 },
    { label:'Realizado', data:M12.map(x => x.real), backgroundColor:'#a30000', fmt:v => fmtV(v, ref) } ] },
    options:{ maintainAspectRatio:false, animation:false, plugins:{ legend:{ position:'bottom' } } }, plugins:[labelPlugin] });
  const ac = byM('Acumulado'), mx = Math.max(ac.real || 0, ac.meta || 0) * 1.1 || 1;
  chart('chGauge', { type:'doughnut', data:{ labels:['Realizado', ''], datasets:[{ data:[ac.real || 0, mx - (ac.real || 0)], backgroundColor:[ac.r && C.farol(ac.r) === 'ok' ? '#007024' : '#a30000', '#e6e6e6'], borderWidth:0 }] },
    options:{ maintainAspectRatio:false, animation:false, rotation:-90, circumference:180, cutout:'70%', plugins:{ legend:{ display:false }, tooltip:{ enabled:false } } } });
  $('#gaugeTxt').innerHTML = ac.real === null ? '<span style="color:#999">(Em branco)</span>' : `${fmtV(ac.real, ref)}<br><small style="font-size:12px;color:#666">Meta ${fmtV(ac.meta, ref)}</small>`;
  atgChart('chAtg', b);
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
  $('#printHead').textContent = `Ano ${S.ano} · ${S.empresa || 'Todas'} · ${S.cargos.size ? [...S.cargos].join(', ') : 'Todos os cargos'} · Responsável: ${S.resp}`;
  $('#pt1').textContent = 'Mês: ' + S.mes;
  table($('#tblP1'), b.filter(r => r.mes === S.mes)); table($('#tblP2'), b.filter(r => r.mes === 'Acumulado'));
  atgChart('chMes3', b);
}
function render() { if (!ROWS.length) return; ({ painel:renderPainel, ating:renderAting, print:renderPrint })[curTab()](); }
if (location.hash === '#demo' || !window.REAL_DATA) $('#demo').click();
else if (window.REAL_DATA) loadAoa(REAL_DATA.meta, REAL_DATA.realizado, false, REAL_DATA.mtime);
