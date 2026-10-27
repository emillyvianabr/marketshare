/* PAX Concorrência — painel estático. Dados: data/pax.csv · Regras: data/config.json */
(function () {
  'use strict';
  const MESES = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  const MESES_LONGO = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  const RAW = ['Bondinho','AquaRio','Trem Corcovado','Paineiras','BioParque','YupStar'];
  const CRISTO = 'Cristo (total)', MERC = 'Mercado';
  const ORDEM = ['Bondinho','AquaRio',CRISTO,'Trem Corcovado','Paineiras','BioParque','YupStar'];
  const COR = {'Bondinho':'#4B4F9E','AquaRio':'#62BAB6',[CRISTO]:'#233744','Trem Corcovado':'#1D6DB6','Paineiras':'#2F9E78','BioParque':'#E8A13A','YupStar':'#EB4A75'};
  const nf1 = new Intl.NumberFormat('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1});
  const nf0 = new Intl.NumberFormat('pt-BR',{maximumFractionDigits:0});
  const $ = (id) => document.getElementById(id);

  let DB = {}, CFG = {}, ZERO = '9999-99', charts = [];
  const key = (y, m) => y + '-' + String(m).padStart(2, '0');
  const lab = (ym) => MESES[+ym.slice(5) - 1] + '/' + ym.slice(2, 4);
  const addM = (y, m, k) => { const t = y * 12 + (m - 1) + k; return [Math.floor(t / 12), t % 12 + 1]; };

  /* ---------- dados ---------- */
  function parseCSV(text) {
    const out = {}; const lines = text.trim().split(/\r?\n/); lines.shift();
    for (const ln of lines) {
      if (!ln.trim()) continue;
      const [d, a, v] = ln.split(',');
      const ym = d.slice(0, 7); (out[ym] = out[ym] || {})[a.trim()] = parseFloat(v);
    }
    return out;
  }
  function applyRules() {
    for (const ym of Object.keys(DB)) if (ym >= ZERO) DB[ym]['BioParque'] = 0;
  }
  function val(ym, att) {
    const r = DB[ym]; if (!r) return 0;
    if (att === CRISTO) return (r['Trem Corcovado'] || 0) + (r['Paineiras'] || 0);
    if (att === MERC) return ['Bondinho','AquaRio',CRISTO,'BioParque','YupStar'].reduce((s, a) => s + val(ym, a), 0);
    return r[att] || 0;
  }
  const hasData = (ym) => !!DB[ym];
  function ytd(y, m, att) { let s = 0; for (let i = 1; i <= m; i++) s += val(key(y, i), att); return s; }

  /* ---------- formatação ---------- */
  const pct = (x) => nf1.format(x * 100) + '%';
  const sgn = (x, f) => (x > 0 ? '+' : x < 0 ? '−' : '') + f(Math.abs(x));
  const varPct = (a, b) => (a > 0 && b > 0) ? a / b - 1 : null;
  const fVar = (v) => v === null ? '—' : sgn(v, (z) => nf1.format(z * 100) + '%');
  const fPP = (v) => v === null ? '—' : sgn(v, (z) => nf1.format(z * 100)) + ' p.p.';
  const cls = (v) => v === null ? 'v-nd' : v > 0.00049 ? 'v-pos' : v < -0.00049 ? 'v-neg' : '';

  /* ---------- estado ---------- */
  function state() {
    return { A: +$('f-ano').value, B: +$('f-comp').value, M: +$('f-mes').value };
  }
  function marcado(s) { return key(s.A, s.M) >= ZERO || key(s.B, s.M) >= ZERO; }

  function metrics(att, s) {
    const kA = key(s.A, s.M), kB = key(s.B, s.M);
    const mA = val(kA, att), mB = val(kB, att), yA = ytd(s.A, s.M, att), yB = ytd(s.B, s.M, att);
    const kmA = val(kA, MERC), kmB = val(kB, MERC), ymA = ytd(s.A, s.M, MERC), ymB = ytd(s.B, s.M, MERC);
    const shA = kmA ? mA / kmA : null, shB = kmB ? mB / kmB : null;
    const ysA = ymA ? yA / ymA : null, ysB = ymB ? yB / ymB : null;
    return { mA, mB, vM: varPct(mA, mB), yA, yB, vY: varPct(yA, yB), shA, shB,
      dM: (shA !== null && shB !== null) ? shA - shB : null, ysA, ysB,
      dY: (ysA !== null && ysB !== null) ? ysA - ysB : null };
  }

  /* ---------- filtros ---------- */
  function initFilters() {
    const yms = Object.keys(DB).sort();
    const anos = [...new Set(yms.map((k) => +k.slice(0, 4)))].sort((a, b) => b - a);
    $('f-ano').innerHTML = anos.map((a) => `<option>${a}</option>`).join('');
    $('f-ano').value = anos[0];
    fillComp(); fillMes();
    $('f-ano').onchange = () => { fillComp(); fillMes(); render(); };
    $('f-comp').onchange = render; $('f-mes').onchange = render;
    const last = yms[yms.length - 1];
    $('dados-ate').textContent = 'Dados até ' + MESES_LONGO[+last.slice(5) - 1] + ' de ' + last.slice(0, 4) + '.';
  }
  function fillComp() {
    const A = +$('f-ano').value;
    const anos = [...new Set(Object.keys(DB).map((k) => +k.slice(0, 4)))].filter((a) => a !== A).sort((a, b) => b - a);
    $('f-comp').innerHTML = anos.map((a) => `<option>${a}</option>`).join('');
    const prev = anos.find((a) => a < A) || anos[0]; $('f-comp').value = prev;
  }
  function fillMes() {
    const A = +$('f-ano').value; const ms = [];
    for (let m = 1; m <= 12; m++) if (hasData(key(A, m))) ms.push(m);
    $('f-mes').innerHTML = ms.map((m) => `<option value="${m}">${MESES_LONGO[m - 1]}</option>`).join('');
    $('f-mes').value = ms[ms.length - 1];
  }

  /* ---------- blocos ---------- */
  function renderAviso(s) {
    const el = $('aviso');
    const [zy, zm] = ZERO.split('-').map(Number);
    const p = [];
    if (marcado(s)) p.push('<b>* Atenção:</b> ' + CFG.nota_bioparque.replace('{desde}', MESES_LONGO[zm - 1] + ' de ' + zy));
    if (s.A === 2019 || s.B === 2019) p.push('<b>Comparação com 2019:</b> o mercado de 2019 considera só Bondinho, AquaRio e Cristo (a YupStar abriu em dez/19 e o BioParque entrou em mar/21), então variações do mercado e shares não são totalmente comparáveis.');
    el.hidden = !p.length; el.innerHTML = p.join('<br>');
  }

  function renderInsight(s, star) {
    const M = metrics(CRISTO, s), B = metrics('Bondinho', s), K = metrics(MERC, s);
    const pl = `${MESES[s.M - 1]}/${String(s.A).slice(2)}`, cl = `${MESES[s.M - 1]}/${String(s.B).slice(2)}`;
    const col = (v, t) => `<span class="${v > 0 ? 'pos' : v < 0 ? 'neg' : ''}">${t}</span>`;
    let t = `Em ${pl}, o Bondinho recebeu <b>${nf1.format(B.mA)} mil</b> visitantes`;
    if (B.vM !== null) t += ` (${col(B.vM, fVar(B.vM))} vs ${cl})`;
    t += ` e ficou com <b>${B.shA !== null ? pct(B.shA) : '—'}${star}</b> do mercado`;
    if (B.dM !== null) t += ` (${col(B.dM, fPP(B.dM))})`;
    t += '.';
    if (K.vM !== null) t += ` O mercado variou ${col(K.vM, fVar(K.vM))}${star}.`;
    const cand = ['AquaRio', CRISTO, 'YupStar'].map((a) => ({ a, v: metrics(a, s).vM })).filter((x) => x.v !== null);
    if (cand.length > 1) {
      cand.sort((x, y) => y.v - x.v);
      const up = cand[0], dn = cand[cand.length - 1];
      const nm = (a) => a === CRISTO ? 'Cristo' : a;
      t += ` Entre os concorrentes, a maior alta foi do ${nm(up.a)} (${col(up.v, fVar(up.v))}) e a maior queda do ${nm(dn.a)} (${col(dn.v, fVar(dn.v))}).`;
    }
    $('insight').innerHTML = t;
  }

  function chip(v, f) {
    if (v === null) return '<span class="chip chip--neu">sem base de comparação</span>';
    const c = v > 0.00049 ? 'pos' : v < -0.00049 ? 'neg' : 'neu';
    return `<span class="chip chip--${c}">${f(v)}</span>`;
  }
  function renderKpis(s, star) {
    const B = metrics('Bondinho', s);
    const mes = `${MESES[s.M - 1]}/${String(s.A).slice(2)}`;
    const items = [
      ['PAX no mês', nf1.format(B.mA), 'mil', chip(B.vM, fVar), `vs ${nf1.format(B.mB)} mil em ${MESES[s.M - 1]}/${String(s.B).slice(2)}`],
      ['Market share no mês', (B.shA !== null ? pct(B.shA) : '—') + star, '', chip(B.dM, fPP), `vs ${B.shB !== null ? pct(B.shB) : '—'} em ${MESES[s.M - 1]}/${String(s.B).slice(2)}`],
      ['PAX acumulado (YTD)', nf1.format(B.yA), 'mil', chip(B.vY, fVar), `jan–${MESES[s.M - 1]}/${s.A} vs ${nf1.format(B.yB)} mil em ${s.B}`],
      ['Market share YTD', (B.ysA !== null ? pct(B.ysA) : '—') + star, '', chip(B.dY, fPP), `vs ${B.ysB !== null ? pct(B.ysB) : '—'} em ${s.B}`]
    ];
    $('kpis').innerHTML = items.map((i) => `<article class="kpi"><p class="kpi__label">Bondinho · ${i[0]}</p>
      <p class="kpi__value">${i[1]}${i[2] ? `<small>${i[2]}</small>` : ''}</p>${i[3]}<p class="kpi__ref">${i[4]}</p></article>`).join('');
  }

  function renderTabela(s, star) {
    const mm = MESES[s.M - 1];
    $('tab-sub').textContent = `Mês de fechamento: ${MESES_LONGO[s.M - 1]} · YTD = janeiro a ${mm}`;
    const h2 = `<tr><th>${s.A}</th><th>${s.B}</th><th>Var.</th><th>${s.A}</th><th>${s.B}</th><th>Var.</th>
      <th>${s.A}</th><th>${s.B}</th><th>Δ p.p.</th><th>${s.A}</th><th>${s.B}</th><th>Δ p.p.</th></tr>`;
    const h1 = `<tr><th rowspan="2" style="text-align:left">Atração</th><th colspan="3">PAX no mês (mil)</th><th colspan="3">PAX YTD (mil)</th>
      <th colspan="3">Market share no mês${star}</th><th colspan="3">Market share YTD${star}</th></tr>`;
    const rows = ORDEM.concat([MERC]).map((att) => {
      const m = metrics(att, s);
      const nome = att === 'BioParque' ? 'BioParque' + star : att === MERC ? 'Mercado' : att;
      const c = att === MERC ? 'tot' : (att === 'Trem Corcovado' || att === 'Paineiras') ? 'sub' : att === 'Bondinho' ? 'dest' : '';
      const sh = (x) => x === null ? '—' : pct(x);
      const num = (x) => x ? nf1.format(x) : '–';
      return `<tr class="${c}"><td>${nome}</td>
        <td>${num(m.mA)}</td><td>${num(m.mB)}</td><td class="${cls(m.vM)}">${fVar(m.vM)}</td>
        <td class="sep">${num(m.yA)}</td><td>${num(m.yB)}</td><td class="${cls(m.vY)}">${fVar(m.vY)}</td>
        <td class="sep">${sh(m.shA)}</td><td>${sh(m.shB)}</td><td class="${cls(m.dM)}">${fPP(m.dM)}</td>
        <td class="sep">${sh(m.ysA)}</td><td>${sh(m.ysB)}</td><td class="${cls(m.dY)}">${fPP(m.dY)}</td></tr>`;
    }).join('');
    $('tabela').innerHTML = `<thead>${h1}${h2}</thead><tbody>${rows}</tbody>`;
  }

  /* ---------- gráficos ---------- */
  Chart.defaults.font.family = "'Stolzl', system-ui, sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.color = '#53778B';
  const barLabels = {
    id: 'barLabels',
    afterDatasetsDraw(chart, _a, opts) {
      const { ctx } = chart; ctx.save();
      ctx.font = "500 12px 'Stolzl', sans-serif"; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      chart.data.datasets.forEach((ds, i) => {
        chart.getDatasetMeta(i).data.forEach((bar, j) => {
          const v = ds.data[j]; if (!v || bar.height < 17) return;
          ctx.fillStyle = ds.labelColor || '#fff';
          ctx.fillText(opts.fmt(v), bar.x, bar.y + bar.height / 2);
        });
      });
      ctx.restore();
    }
  };
  function mk(id, cfg) { charts.push(new Chart($(id), cfg)); }

  function renderCharts(s, star) {
    charts.forEach((c) => c.destroy()); charts = [];
    // janela móvel de 13 meses
    const wins = []; for (let k = 12; k >= 0; k--) { const [y, m] = addM(s.A, s.M, -k); wins.push(key(y, m)); }
    const labels = wins.map(lab);
    const linhas = ['Bondinho','AquaRio',CRISTO,'Trem Corcovado','Paineiras','YupStar'];
    const ds = (fn) => linhas.map((a) => ({ label: a, borderColor: COR[a], backgroundColor: COR[a],
      borderWidth: a === 'Bondinho' ? 3.5 : 2.2, pointRadius: a === 'Bondinho' ? 3.5 : 2.2, tension: .25, spanGaps: false,
      data: wins.map((w) => hasData(w) ? fn(w, a) : null) }));
    const line = (fmt, tipFmt, data) => ({ type: 'line', data: { labels, datasets: data },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
        plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8 } },
          tooltip: { callbacks: { label: (c) => ' ' + c.dataset.label + ': ' + tipFmt(c.parsed.y) } } },
        scales: { y: { beginAtZero: true, ticks: { callback: fmt }, grid: { color: '#E6ECEF' } }, x: { grid: { display: false } } } } });
    mk('c-pax', line((v) => nf0.format(v), (v) => nf1.format(v) + ' mil', ds((w, a) => val(w, a))));
    mk('c-share', line((v) => nf0.format(v * 100) + '%', (v) => pct(v), ds((w, a) => { const t = val(w, MERC); return t ? val(w, a) / t : null; })));

    // YTD empilhado
    const ordem = ['BioParque','YupStar','AquaRio','Bondinho',CRISTO];
    const lblY = ['YTD ' + String(s.B).slice(2), 'YTD ' + String(s.A).slice(2)];
    const stack = (fn, fmt, tip, max) => ({ type: 'bar', plugins: [barLabels],
      data: { labels: lblY, datasets: ordem.map((a) => ({ label: a === 'BioParque' ? 'BioParque' + star : a, backgroundColor: COR[a], borderWidth: 1, borderColor: '#fff', borderRadius: 0,
        data: [fn(s.B, a), fn(s.A, a)], maxBarThickness: 120 })) },
      options: { responsive: true, maintainAspectRatio: false,
        plugins: { barLabels: { fmt }, legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8 } },
          tooltip: { callbacks: { label: (c) => ' ' + c.dataset.label + ': ' + tip(c.parsed.y) } } },
        scales: { x: { stacked: true, grid: { display: false } },
          y: { stacked: true, beginAtZero: true, max, ticks: { callback: (v) => max ? nf0.format(v * 100) + '%' : nf0.format(v) }, grid: { color: '#E6ECEF' } } } } });
    mk('c-ytd-pax', stack((y, a) => ytd(y, s.M, a), (v) => nf0.format(v), (v) => nf1.format(v) + ' mil'));
    mk('c-ytd-share', stack((y, a) => { const t = ytd(y, s.M, MERC); return t ? ytd(y, s.M, a) / t : 0; }, (v) => nf0.format(v * 100) + '%', (v) => pct(v), 1));
  }

  /* ---------- heatmap ---------- */
  function renderHeat(s, star) {
    $('heat-sub').textContent = `${s.A} vs ${s.B}, mês a mês`;
    let h = '<div class="h"></div>' + MESES.map((m) => `<div class="h">${m}</div>`).join('');
    ORDEM.concat([MERC]).forEach((att) => {
      const nome = att === 'BioParque' ? 'BioParque' + star : att;
      h += `<div class="r">${nome}</div>`;
      for (let m = 1; m <= 12; m++) {
        const a = val(key(s.A, m), att), b = val(key(s.B, m), att);
        const v = hasData(key(s.A, m)) ? varPct(a, b) : null;
        if (v === null) { h += '<div class="e">–</div>'; continue; }
        const k = Math.min(Math.abs(v) / 0.4, 1);
        const bg = v >= 0 ? `rgba(98,186,182,${.15 + .65 * k})` : `rgba(235,74,117,${.12 + .5 * k})`;
        const rr = Math.round(Math.abs(v) * 100);
        h += `<div style="background:${bg}">${(rr === 0 ? '' : v > 0 ? '+' : '−') + rr}%</div>`;
      }
    });
    $('heat').innerHTML = h;
  }

  function render() {
    const s = state(); const star = marcado(s) ? '*' : '';
    renderAviso(s); renderInsight(s, star); renderKpis(s, star);
    renderTabela(s, star); renderCharts(s, star); renderHeat(s, star);
    const [zy, zm] = ZERO.split('-').map(Number);
    $('nota-rodape').textContent = '* ' + CFG.nota_bioparque.replace('{desde}', MESES_LONGO[zm - 1] + ' de ' + zy);
    $('fonte').textContent = 'Fonte: ' + CFG.fonte;
  }

  /* ---------- carga ---------- */
  async function load() {
    try {
      const [c, j] = await Promise.all([fetch('data/pax.csv', { cache: 'no-cache' }), fetch('data/config.json', { cache: 'no-cache' })]);
      if (!c.ok || !j.ok) throw new Error('fetch');
      return [await c.text(), await j.json()];
    } catch (e) {
      if (window.PAX_CSV && window.PAX_CONFIG) return [window.PAX_CSV, window.PAX_CONFIG]; // abertura local sem servidor
      throw e;
    }
  }
  load().then(([csv, cfg]) => {
    CFG = cfg; ZERO = cfg.bioparque_zero_desde || ZERO;
    DB = parseCSV(csv); applyRules(); initFilters(); render();
  }).catch(() => {
    document.querySelector('main').innerHTML = '<section class="card"><h2>Não foi possível carregar os dados</h2><p>Abra o painel pelo endereço publicado (GitHub Pages) ou rode <code>python -m http.server</code> na pasta do projeto.</p></section>';
  });
})();
