/* Book de Metas – fórmulas (porta do DAX, spec.md §2) + transformações Power Query (§1). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.BookCalc = factory();
})(this, function () {
  const MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez','Acumulado'];
  const MES_ORD = Object.fromEntries(MESES.map((m, i) => [m, i + 1]));
  const FORMATO = { '%':'PERCENTUAL', 'Dias':'NUMERO', 'R$':'NUMERO', '#':'NUMERO', 'R$/t':'NUMERO DECIMAL',
    'Horas':'HORAS', 'd':'TRES CASAS', 't':'NUMERO', 'm³':'NUMERO', 'm³/t':'NUMERO DECIMAL' };
  const NIVEL_ORD = { 'Empresa':0,'Diretoria':1,'Superintendência':2,'Gerência':3,'Staff':4,'Supervisão':5,'Outros':6 };
  // Nível hierárquico por "Cargo-Responsável". NÃO contém nomes no código: a tabela real é carregada
  // em tempo de execução (API: aba "Nivel" do BD BOOK ou app setting NIVEL_JSON). Sem entrada → regra pelo cargo.
  const NIVEL_DEMO = { // somente nomes FICTÍCIOS (dados demo)
    'Gerente de Área Alfa-Carla Mendes': ['Gerência', 3], 'Gerente de Área Beta-Paulo Teixeira': ['Gerência', 3],
    'Diretor de Área Gama-Ricardo Nogueira': ['Diretoria', 1], 'Gerente de Área Delta-Beatriz Lima': ['Gerência', 3],
    'Superintendente de Área Épsilon-Sérgio Prado': ['Superintendência', 2], 'Gerente de Área Zeta-Marina Duarte': ['Gerência', 3],
    'Book Geral CSC-Empresa': ['Empresa', 0], 'Book Geral MIG-Empresa': ['Empresa', 0], 'Book Geral MML-Empresa': ['Empresa', 0] };
  const NIVEL = { ...NIVEL_DEMO }; // demo público: só a tabela fictícia
  function nivelPorCargo(cargo, resp) {
    const c = String(cargo || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(), r = String(resp || '').trim().toLowerCase();
    if (/^book geral/.test(c) && r === 'empresa') return 'Empresa';
    if (/^diretor/.test(c)) return 'Diretoria';
    if (/^superintendente|^gerente geral/.test(c)) return 'Superintendência';
    if (/^gerente/.test(c)) return 'Gerência';
    if (/^supervis/.test(c)) return 'Supervisão';
    if (/^(coordenador|engenheiro|comprador)/.test(c)) return 'Staff';
    return 'Outros';
  }
  // Aceita {"Cargo-Responsável": "Gerência"} ou {"Cargo-Responsável": ["Gerência", 3]}
  function normalizeNivel(map) {
    const out = {}; for (const [k, v] of Object.entries(map || {})) { const n = Array.isArray(v) ? v[0] : v; if (n in NIVEL_ORD) out[k] = [n, NIVEL_ORD[n]]; } return out;
  }
  function setNivel(map) { for (const k of Object.keys(NIVEL)) delete NIVEL[k]; Object.assign(NIVEL, normalizeNivel(map)); }
  function nivelOf(cargo, resp, map) { const nv = (map || NIVEL)[`${cargo}-${resp}`]; if (nv) return nv; const n = nivelPorCargo(cargo, resp); return [n, NIVEL_ORD[n]]; }

  const isBlank = v => v === null || v === undefined || v === '' || (typeof v === 'number' && isNaN(v));
  const num = v => { if (isBlank(v)) return null; if (typeof v === 'number') return v;
    let s = String(v).trim(); if (s === '-' || s === '') return null;
    if (/^-?[\d.]+,\d+$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    const pct = s.endsWith('%'); if (pct) s = s.slice(0, -1);
    const n = Number(s); return isNaN(n) ? null : (pct ? n / 100 : n); };

  // Δ% com IFERROR -> 0. opts.pbiBlankAsZero = comportamento original do PBI (blank conta como 0).
  function deltaPct(meta, real) {
    const r = real ?? 0;
    if (!meta) return 0;
    return (r - meta) / meta;
  }
  function atgIndicador(sinal, limite, meta, real, opts = {}) {
    if (isBlank(real) && !opts.pbiBlankAsZero) return 0;            // correção pedida: sem dado = 0
    const d = deltaPct(meta, real);
    if (limite === 0.9999) {
      if (sinal === '↑') return d < -1 ? 0 : d > 0 ? 1 : 1 + d;
      if (sinal === '↓') return d > 1 ? 0 : d < 0 ? 1 : 1 - d;
    }
    if (sinal === '↑') { if (d >= 0) return 1; return -d <= limite ? 1 - (-d / limite) * 0.3 : 0; }
    if (sinal === '↓') { if (d <= 0) return 1; return d <= limite ? 1 - (d / limite) * 0.3 : 0; }
    return null;
  }
  function atgDesafio(sinal, desafio, meta, real, opts = {}) {
    if (!desafio) return null;
    if (isBlank(real) && !opts.pbiBlankAsZero) return 0;
    const d = deltaPct(meta, real);
    if (sinal === '↑' && d >= 0) return Math.min(1, d / desafio);
    if (sinal === '↓' && d <= 0) return Math.min(1, -d / desafio);
    return 0;
  }
  function limiteMes(today) { const m = (today || new Date()).getMonth() + 1; return m === 1 ? 12 : m - 1; }

  // ---------- Power Query: planilha -> linhas longas ----------
  function headerToMes(h) {
    if (h instanceof Date) return MESES[h.getMonth()];
    if (typeof h === 'number' && h > 20000) { const d = new Date(Math.round((h - 25569) * 86400000)); return MESES[d.getUTCMonth()]; }
    const s = String(h ?? '').trim();
    if (/^acumulado$/i.test(s)) return 'Acumulado';
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (m) return MESES[+m[2] - 1];
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return MESES[+m[2] - 1];
    if (MES_ORD[s] && s !== 'Acumulado') return s;
    return null;
  }
  const FIXAS = ['Empresa','Cargo','Responsável','Indicador','Unidade','Peso','Sinal','Limite','Desafio','Ano'];
  function transformSheet(aoa, { filtraIndicadorNulo }) {
    const head = aoa[0] || []; const out = [];
    const idx = {}; FIXAS.forEach(f => { idx[f] = head.findIndex(h => String(h ?? '').trim().normalize('NFC') === f.normalize('NFC')); });
    if (idx.Indicador < 0 || idx.Ano < 0) throw new Error('Cabeçalho inválido: colunas obrigatórias (Indicador, Ano…) não encontradas.');
    const meses = []; head.forEach((h, i) => { if (Object.values(idx).includes(i)) return; const m = headerToMes(h); if (m) meses.push([i, m]); });
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r] || []; const g = f => idx[f] >= 0 ? row[idx[f]] : null;
      const ind = g('Indicador');
      if (filtraIndicadorNulo && isBlank(ind)) continue;
      let resp = isBlank(g('Responsável')) ? '' : String(g('Responsável'));
      resp = resp.split('0').join('Empresa').split('-').join('Empresa');   // Replacer.ReplaceText (substring)
      let sinal = String(g('Sinal') ?? '').trim(); sinal = sinal.replace(/i/g, '↓').replace(/h/g, '↑');
      let desafio = g('Desafio'); desafio = (String(desafio).trim() === '-') ? 0 : num(desafio);
      const base = { empresa: String(g('Empresa') ?? '').trim(), cargo: String(g('Cargo') ?? '').trim(), responsavel: resp.trim(),
        indicador: String(ind ?? '').trim(), unidade: String(g('Unidade') ?? '').trim(), peso: num(g('Peso')) ?? 0,
        sinal, limite: num(g('Limite')), desafio, ano: String(num(g('Ano')) ?? g('Ano') ?? '').trim() };
      for (const [i, mes] of meses) {
        const v = row[i]; if (isBlank(v)) continue;        // unpivot descarta nulos
        const val = num(v); if (val === null) continue;
        out.push({ ...base, mes, valor: val, key: `${base.empresa}-${base.cargo}-${base.responsavel}-${base.indicador}/${mes}-${base.ano}` });
      }
    }
    return out;
  }
  // Tabela "Concatenado": uma linha por chave de Meta, com Realizado por LOOKUPVALUE.
  function buildModel(metaRows, realRows, opts = {}) {
    const real = new Map(); realRows.forEach(r => real.set(r.key, r.valor));
    const seen = new Set(); const rows = [];
    for (const m of metaRows) {
      if (seen.has(m.key)) continue; seen.add(m.key);
      const r = real.has(m.key) ? real.get(m.key) : null;
      const semDado = r === null;
      const d = semDado ? null : r - m.valor;
      const dp = semDado ? null : deltaPct(m.valor, r);
      const atg = atgIndicador(m.sinal, m.limite, m.valor, r, opts);
      const nv = nivelOf(m.cargo, m.responsavel, opts.nivel ? normalizeNivel(opts.nivel) : null);
      rows.push({ ...m, meta: m.valor, realizado: r, semDado, delta: d, deltaPct: dp, mesOrd: MES_ORD[m.mes],
        formato: FORMATO[m.unidade] || null, atgInd: atg, atgBook3: (atg ?? 0) * m.peso / 100,
        atgDesafio: atgDesafio(m.sinal, m.desafio, m.valor, r, opts), pesoDesafio: m.desafio ? m.peso : 0,
        nivel: nv ? nv[0] : 'Outros', nivelOrd: nv ? nv[1] : 6, c5: `${m.empresa}-${m.cargo}-${m.ano}` });
    }
    // Atg. Book4: renormaliza peso entre indicadores com desafio da mesma Empresa-Cargo-Ano (soma/13 = soma por linha-mês)
    const pd = new Map(); rows.forEach(x => pd.set(x.c5, (pd.get(x.c5) || 0) + x.pesoDesafio));
    const lim = limiteMes(opts.today);
    rows.forEach(x => { const den = (pd.get(x.c5) || 0) / 13; const v = (x.atgDesafio === null || !den) ? null : x.atgDesafio * x.peso / den;
      x.atgBook4 = (x.mesOrd > lim && x.mesOrd < 13) ? 0 : (v ?? 0); });
    return rows;
  }
  // Variante "por linha-mês" do denominador (mais robusta quando nem todos os meses têm meta): usada no teste p/ comparação
  function farol(r) { if (r.semDado) return 'sem'; if (r.atgInd >= 1) return 'ok'; if (r.atgInd > 0) return 'parcial'; return 'falha'; }
  return { MESES, MES_ORD, FORMATO, NIVEL, NIVEL_DEMO, NIVEL_ORD, setNivel, normalizeNivel, nivelOf, nivelPorCargo, deltaPct, atgIndicador, atgDesafio, limiteMes, transformSheet, buildModel, farol, num, headerToMes };
});
