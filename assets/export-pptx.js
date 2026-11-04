/* Exportação para PowerPoint — reproduz o layout do relatório "PAX Atrativos Rio de Janeiro".
   Recebe o modelo numérico montado em app.js (buildReport) e gera um .pptx com gráficos nativos (editáveis).
   Depende de assets/vendor/pptxgen.bundle.js (PptxGenJS + JSZip), carregado só quando o botão é clicado. */
(function () {
  'use strict';
  const FONT = 'Ofelia Display', FONT_LT = 'Ofelia Display Light';
  const ORANGE = 'FF6600', CRISTO = 'Cristo (total)';
  const LOGO = 'image/png;base64,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';

  /* cores do relatório (fill) e versão um pouco mais escura para texto sobre fundo branco */
  const COR = { 'Bondinho': 'FF6600', 'AquaRio': 'E93CAC', [CRISTO]: '753BBD', 'Trem Corcovado': 'FFC000', 'Paineiras': 'A6A6A6', 'YupStar': '00B0F0', 'BioParque': '2E9E78' };
  const TXT = Object.assign({}, COR, { 'Trem Corcovado': 'E6A800', 'Paineiras': '7F7F7F' });

  /* ---------- formatação (pt-BR, como no relatório) ---------- */
  const nf0 = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
  const r100 = (v) => Math.round(v * 100);
  const fPct = (v) => v == null ? '–' : (r100(v) < 0 ? '-' : '+') + Math.abs(r100(v)) + '%';          // +0%, -3%
  const fPP = (v) => { if (v == null) return '–'; const r = r100(v); return (r > 0 ? '+' : r < 0 ? '-' : '') + Math.abs(r) + ' p.p'; };
  const pInt = (v) => v == null ? '–' : r100(v) + '%';
  const cap = (t) => t.charAt(0).toUpperCase() + t.slice(1);
  const yy = (a) => String(a).slice(2);
  const niceCeil = (v, step) => Math.ceil(v / step) * step;

  /* ---------- textos de abertura (mesmo estilo do relatório) ---------- */
  function textoPax(R) {
    const m = R.m, B = m['Bondinho'], C = m[CRISTO], T = m['Trem Corcovado'], P = m['Paineiras'], Q = m['AquaRio'], Y = m['YupStar'], K = m['Mercado'];
    let t = `Em ${R.mes} de ${R.A}, considerando os principais atrativos do Rio de Janeiro,`;
    t += K.vM != null ? ` o mercado variou ${fPct(K.vM)} em relação a ${R.B}.` : ' não há base de comparação para o mercado.';
    const cand = [['Bondinho', B], ['Cristo (Total)', C], ['AquaRio', Q], ['YupStar', Y]].filter((x) => x[1].vM != null);
    const best = cand.filter((x) => r100(x[1].vM) > 0).sort((a, b) => b[1].vM - a[1].vM)[0];
    const itens = cand.map(([nome, x]) => {
      let s = best && best[0] === nome ? `${nome} (${fPct(x.vM)}) apresenta o maior crescimento percentual do período` : `${nome} ${fPct(x.vM)}`;
      if (nome === 'Cristo (Total)' && T.vM != null && P.vM != null) {
        if (r100(C.vM) < 0 && (r100(T.vM) < 0 || r100(P.vM) < 0)) {
          const q = []; if (r100(T.vM) < 0) q.push(`${fPct(T.vM)} do Trem`); if (r100(P.vM) < 0) q.push(`${fPct(P.vM)} de Paineiras`);
          s += `, pressionado pela queda de ${q.join(' e de ')}`;
        } else if (r100(C.vM) > 0 && (r100(T.vM) > 0 || r100(P.vM) > 0)) {
          const q = []; if (r100(T.vM) > 0) q.push(`${fPct(T.vM)} do Trem`); if (r100(P.vM) > 0) q.push(`${fPct(P.vM)} de Paineiras`);
          s += `, impulsionado pela alta de ${q.join(' e de ')}`;
        }
      }
      return s;
    });
    if (itens.length) {
      const ult = itens[itens.length - 1], ant = itens[itens.length - 2] || '';
      t += ' ' + (itens.length > 1 ? itens.slice(0, -1).join('; ') + (ant.indexOf(',') >= 0 ? '; e ' : ' e ') + ult : ult) + '.';
    }
    if (K.vY != null) {
      const y = [['Bondinho', B], ['Cristo (Total)', C], ['AquaRio', Q]].filter((x) => x[1].vY != null).map(([n, x]) => `${n} ${fPct(x.vY)}`);
      t += ` No acumulado do ano (YTD), o mercado varia ${fPct(K.vY)}` + (y.length ? ': ' + (y.length > 1 ? y.slice(0, -1).join(', ') + ' e ' + y[y.length - 1] : y[0]) : '') + '.';
    }
    return t;
  }
  function textoShare(R) {
    const m = R.m, B = m['Bondinho'], C = m[CRISTO], T = m['Trem Corcovado'], P = m['Paineiras'], Q = m['AquaRio'];
    const pp = (v) => fPP(v) + '.';
    let t = `No Market Share de ${R.mes} de ${R.A}, o Bondinho apresentou ${B.dM != null ? pp(B.dM) + ' ' : ''}(${pInt(B.shA)}), enquanto o Cristo (Total) ${C.dM != null ? pp(C.dM) + ' ' : ''}(${pInt(C.shA)}).`;
    if (C.dM != null && T.dM != null && P.dM != null) {
      if (r100(C.dM) < 0 && r100(T.dM) < 0) t += ` Esse movimento foi influenciado pela retração do Trem do Corcovado (${pp(T.dM)}), ${r100(P.dM) >= 0 ? 'parcialmente compensada por' : 'somada à queda de'} Paineiras (${pp(P.dM)}).`;
      else if (r100(C.dM) > 0 && r100(T.dM) > 0) t += ` Esse movimento foi influenciado pelo avanço do Trem do Corcovado (${pp(T.dM)}) e de Paineiras (${pp(P.dM)}).`;
      else if (r100(C.dM) !== 0) t += ` Trem do Corcovado (${pp(T.dM)}) e Paineiras (${pp(P.dM)}).`;
    }
    if (Q.dM != null) {
      const r = r100(Q.dM);
      t += r > 0 ? ` AquaRio ganhou ${pp(Q.dM)} (${pInt(Q.shA)}).` : r < 0 ? ` AquaRio perdeu ${Math.abs(r)} p.p. (${pInt(Q.shA)}).` : ` AquaRio manteve ${pInt(Q.shA)}.`;
    }
    return t;
  }

  /* ---------- helpers de layout ---------- */
  // afasta rótulos que ficariam sobrepostos (y = centro, em polegadas)
  function afastar(itens, gap, lo, hi) {
    itens.sort((a, b) => a.y - b.y);
    itens.forEach((it) => { it.y = Math.max(lo, it.y); });
    for (let i = 1; i < itens.length; i++) if (itens[i].y - itens[i - 1].y < gap) itens[i].y = itens[i - 1].y + gap;
    for (let i = itens.length - 1; i >= 0; i--) {
      const lim = hi - (itens.length - 1 - i) * gap;
      if (itens[i].y > lim) itens[i].y = lim;
      if (i < itens.length - 1 && itens[i + 1].y - itens[i].y < gap) itens[i].y = itens[i + 1].y - gap;
    }
    return itens;
  }
  // 't' (acima) ou 'b' (abaixo) para cada rótulo de ponto: o lado com mais espaço livre
  function posicoes(nomes, vals, max) {
    const plano = {}; nomes.forEach((n) => { plano[n] = []; });
    const n = vals[nomes[0]].length;
    for (let i = 0; i < n; i++) {
      const pts = nomes.filter((nm) => vals[nm][i] != null).map((nm) => ({ nm, v: vals[nm][i] })).sort((a, b) => b.v - a.v);
      pts.forEach((p, k) => {
        const up = pts[k - 1], dn = pts[k + 1];
        const gu = up ? up.v - p.v : Infinity, gd = dn ? p.v - dn.v : Infinity;
        let pos = gu > gd ? 't' : 'b';
        if (!dn) pos = (p.v / max < 0.12) ? 't' : 'b';
        if (!up) pos = 't';
        plano[p.nm][i] = pos;
      });
    }
    return plano;
  }
  const txPr = (cor) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill><a:latin typeface="${FONT}"/></a:defRPr></a:pPr><a:endParaRPr lang="pt-BR"/></a:p></c:txPr>`;
  const FLAGS = '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>';
  // rótulos por série: cor da série + posição ponto a ponto (PptxGenJS só aceita um padrão para o gráfico todo)
  function ajustarRotulos(xml, plano, fmt, nulos) {
    return xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (ser) => {
      const nm = (ser.match(/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/) || [])[1];
      const p = plano[nm]; if (!p) return ser;
      if (nulos[nm] && nulos[nm].length) {   // meses sem dado: ponto removido = lacuna na linha (não vira zero)
        const i = ser.indexOf('<c:val>'), cab = ser.slice(0, i); let val = ser.slice(i);
        nulos[nm].forEach((k) => { val = val.replace(new RegExp(`<c:pt idx="${k}"><c:v>[^<]*</c:v></c:pt>`), ''); });
        ser = cab + val;
      }
      const cor = TXT[nm] || '404040';
      const nfmt = `<c:numFmt formatCode="${fmt}" sourceLinked="0"/>`, sp = '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>';
      const pts = p.map((pos, i) => pos ? `<c:dLbl><c:idx val="${i}"/>${nfmt}${sp}${txPr(cor)}<c:dLblPos val="${pos}"/>${FLAGS}</c:dLbl>` : '').join('');
      const d = `<c:dLbls>${pts}${nfmt}${sp}${txPr(cor)}<c:dLblPos val="t"/>${FLAGS}</c:dLbls>`;
      return ser.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, '').replace('<c:cat>', d + '<c:cat>');
    });
  }

  /* ---------- moldura comum aos slides ---------- */
  function moldura(pptx, slide, R, intro, caixa) {
    slide.background = { color: ORANGE };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.14, w: 13.333, h: 6.05, fill: { color: 'FFFFFF' }, line: { type: 'none' } });
    slide.addShape(pptx.ShapeType.line, { x: 0.971, y: 0.805, w: 11.8, h: 0, line: { color: 'FFFFFF', width: 2 } });
    slide.addImage({ data: LOGO, x: 0.41, y: 0.589, w: 0.433, h: 0.362, altText: 'Grupo Iter' });
    slide.addText(R.titulo, { x: 0.32, y: 0.19, w: 12.39, h: 0.6, align: 'right', valign: 'middle', fontFace: FONT_LT, fontSize: 32, color: 'FFFFFF', margin: 0, isTextBox: true });
    slide.addText(`MENSAL | ${R.mes.toUpperCase()} ${R.A}`, { x: 9.34, y: 0.84, w: 3.5, h: 0.29, align: 'right', valign: 'middle', fontFace: FONT, fontSize: 12, color: 'FFFFFF', margin: 0, isTextBox: true });
    const tam = intro.length > 560 ? 10 : intro.length > 470 ? 11 : 12;
    slide.addText(intro, Object.assign({ align: 'justify', valign: 'top', fontFace: FONT, fontSize: tam, color: '000000', margin: 0, isTextBox: true }, caixa));
    const nota = 'Nota: ' + R.nota.replace(/^(Nota:\s*|\*\s*)/, '');
    slide.addText(nota, { x: 0.13, y: 7.2, w: 9.2, h: 0.28, align: 'left', valign: 'middle', fontFace: FONT, fontSize: nota.length > 150 ? 7 : 8, color: 'FFFFFF', margin: 0, isTextBox: true });
    slide.addText('Fonte: ' + R.fonte.replace(/^Fonte:\s*/, ''), { x: 9.4, y: 7.23, w: 3.78, h: 0.25, align: 'right', valign: 'middle', fontFace: FONT, fontSize: 8, color: 'FFFFFF', margin: 0, isTextBox: true });
  }
  function rotulo(slide, txt, x, y, w, h, o) {
    slide.addText(txt, Object.assign({ x, y, w, h, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 9, margin: 0, isTextBox: true }, o));
  }
  function cartao(pptx, slide, x, y, w, h) {
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'FFFFFF' }, line: { color: 'D9D9D9', width: 0.75 } });
  }

  /* ---------- gráfico de linhas (13 meses) + rótulos de variação à direita ---------- */
  function linhas(pptx, slide, R, o) {
    const nomes = R.linhas, vals = o.vals;
    const todos = [].concat(...nomes.map((n) => vals[n].filter((v) => v != null)));
    const max = niceCeil(Math.max(...todos) * 1.06, o.passo);
    const L = o.layout, F = o.frame;
    const dados = nomes.map((n) => ({ name: n, labels: R.janela.map((j) => j.label), values: vals[n].map((v) => v == null ? 0 : v) }));
    slide.addChart(pptx.charts.LINE, dados, {
      x: F.x, y: F.y, w: F.w, h: F.h, layout: L,
      chartColors: nomes.map((n) => COR[n]), lineSize: 2.25, lineDataSymbol: 'none', lineSmooth: false,
      showTitle: true, title: o.titulo, titleFontFace: FONT, titleFontSize: 12, titleColor: '595959', titleBold: true,
      showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 9, legendColor: '262626',
      catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10, catAxisLabelColor: '262626', catAxisLineShow: true, catGridLine: { style: 'none' },
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: max, valGridLine: { style: 'none' },
      showValue: true, dataLabelPosition: 't', dataLabelFormatCode: o.fmt, dataLabelFontFace: FONT, dataLabelFontSize: 9, dataLabelFontBold: true
    });
    const nulos = {}; nomes.forEach((n) => { nulos[n] = vals[n].map((v, i) => v == null ? i : -1).filter((i) => i >= 0); });
    LINHAS.push({ plano: posicoes(nomes, vals, max), fmt: o.fmt, nulos });
    // rótulos de variação alinhados à ponta de cada linha
    const n = R.janela.length, plotX = F.x + F.w * L.x, plotW = F.w * L.w;
    const yTop = F.y + F.h * L.y, yH = F.h * L.h;
    const fim = nomes.map((nm) => {
      const ult = vals[nm][n - 1];
      return { nm, y: yTop + yH * (1 - (ult || 0) / max) };
    });
    return { fim, xFimLinha: plotX + plotW * ((n - 0.5) / n), max, yTop, yH };
  }
  function varLado(slide, itens, cx, fmt, valorDe, ymax) {
    afastar(itens, 0.205, 2.5, ymax).forEach((it) => {
      rotulo(slide, fmt(valorDe(it.nm)), cx - 0.4, it.y - 0.11, 0.8, 0.22, { bold: true, color: TXT[it.nm] || '000000' });
    });
  }

  /* ---------- barras YTD empilhadas + abertura do Cristo (Trem/Paineiras) ---------- */
  function ytd(pptx, slide, R, o) {
    const F = o.frame, L = o.layout, A = R.A, B = R.B;
    const pilha = (R.mostraBio ? ['BioParque'] : []).concat(['YupStar', 'AquaRio', 'Bondinho', CRISTO]);
    const v = (ano, a) => o.valor(R.m[a], ano === A);
    const tot = [v(B, 'Mercado'), v(A, 'Mercado')];
    const max = o.max || niceCeil(Math.max(...tot) / 0.86, 500);
    const dados = pilha.map((a) => ({ name: a, labels: ['YTD ' + yy(B), 'YTD ' + yy(A)], values: [v(B, a) || 0, v(A, a) || 0] }));
    slide.addChart(pptx.charts.BAR, dados, {
      x: F.x, y: F.y, w: F.w, h: F.h, layout: L, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 150,
      chartColors: pilha.map((a) => COR[a]), showLegend: false, showTitle: false,
      catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10, catAxisLabelColor: '595959', catAxisLineShow: true, catGridLine: { style: 'none' },
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: max, valGridLine: { style: 'none' },
      showValue: true, dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF', dataLabelFontBold: true, dataLabelFontSize: 9, dataLabelFontFace: FONT, dataLabelFormatCode: o.fmt
    });
    const plotX = F.x + F.w * L.x, plotW = F.w * L.w, yTop = F.y + F.h * L.y, yH = F.h * L.h;
    const slot = plotW / 2, barW = slot / 2.5, yOf = (x) => yTop + yH * (1 - x / max);
    const centros = [plotX + slot * 0.5, plotX + slot * 1.5];
    const col = { 0: B, 1: A };
    const alvos = []; // y (centro) de cada segmento do ano atual, para alinhar os rótulos de variação
    [0, 1].forEach((i) => {
      const ano = col[i];
      let base = 0; const seg = {};
      pilha.forEach((a) => { const x = v(ano, a) || 0; seg[a] = [base, base + x]; base += x; });
      if (o.totais) rotulo(slide, o.totais(base), centros[i] - 0.5, yOf(base) - 0.3, 1.0, 0.26, { bold: true, fontSize: 10.5, color: '000000' });
      // abertura do Cristo ao lado da barra: Trem (embaixo) e Paineiras (em cima), somando o segmento do Cristo
      const [c0, c1] = seg[CRISTO], t = v(ano, 'Trem Corcovado') || 0, p = v(ano, 'Paineiras') || 0, st = t + p;
      if (st > 0) {
        const hTot = yOf(c0) - yOf(c1), hT = hTot * t / st, hP = hTot - hT, x0 = centros[i] - barW / 2 - 0.55;
        slide.addShape(pptx.ShapeType.rect, { x: x0, y: yOf(c1), w: 0.55, h: hP, fill: { color: COR['Paineiras'] }, line: { type: 'none' } });
        slide.addShape(pptx.ShapeType.rect, { x: x0, y: yOf(c1) + hP, w: 0.55, h: hT, fill: { color: COR['Trem Corcovado'] }, line: { type: 'none' } });
        rotulo(slide, o.fmtTxt(p), x0, yOf(c1), 0.55, hP, { color: '000000' });
        rotulo(slide, o.fmtTxt(t), x0, yOf(c1) + hP, 0.55, hT, { color: '000000' });
        if (i === 1) { alvos.push({ nm: 'Paineiras', y: yOf(c1) + hP / 2 }); alvos.push({ nm: 'Trem Corcovado', y: yOf(c1) + hP + hT / 2 }); }
      }
      if (i === 1) pilha.forEach((a) => { const [s0, s1] = seg[a]; if (s1 > s0) alvos.push({ nm: a, y: (yOf(s0) + yOf(s1)) / 2 }); });
    });
    return { alvos, yTop, yH };
  }

  /* ---------- montagem ---------- */
  const LINHAS = [];
  async function montar(R) {
    LINHAS.length = 0;
    const pptx = new window.PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title = `${R.titulo} — ${cap(R.mes)}/${R.A}`; pptx.company = 'Grupo Iter';
    pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
    const jan = R.janela, pri = jan[0].key, ult = jan[jan.length - 1].key;
    const mesAno = (k) => cap(['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'][+k.slice(5) - 1]) + '/' + yy(k.slice(0, 4));
    const tag = `${cap(R.mes3)}`;

    /* ===== slide 1: PAX ===== */
    const s1 = pptx.addSlide();
    moldura(pptx, s1, R, textoPax(R), { x: 0.2, y: 1.21, w: 12.93, h: 0.7 });
    const l1 = linhas(pptx, s1, R, {
      vals: R.pax, passo: 50, fmt: '#,##0', titulo: `PAX Atrativos ${mesAno(pri)} a ${mesAno(ult)} ('000)`,
      frame: { x: 0.31, y: 2.11, w: 8.15, h: 3.87 }, layout: { x: 0.03, y: 0.14, w: 0.94, h: 0.62 }
    });
    const col1 = 8.77;
    rotulo(s1, `VAR. ${R.mes3.toUpperCase()}\n${yy(R.A)}x${yy(R.B)}`, col1 - 0.55, 2.08, 1.1, 0.37, { fontSize: 8, color: '404040' });
    rotulo(s1, `Total\n${fPct(R.m['Mercado'].vM)}`, col1 - 0.4, 2.45, 0.8, 0.36, { bold: true, color: '000000' });
    varLado(s1, l1.fim.slice(), col1, fPct, (n) => R.m[n].vM, 5.95);
    // painel YTD
    cartao(pptx, s1, 9.23, 2.2, 4.02, 4.88);
    s1.addText('PAX YTD\n(‘000)', { x: 10.2, y: 2.25, w: 1.5, h: 0.5, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 12, bold: true, color: '000000', margin: 0, isTextBox: true });
    const cy1 = 12.68;
    rotulo(s1, `VAR. YTD\n${yy(R.A)}x${yy(R.B)}`, cy1 - 0.65, 2.26, 1.3, 0.37, { fontSize: 8, color: '404040' });
    rotulo(s1, fPct(R.m['Mercado'].vY), cy1 - 0.4, 2.66, 0.8, 0.25, { bold: true, color: '000000' });
    const y1 = ytd(pptx, s1, R, {
      frame: { x: 9.5, y: 3.0, w: 2.9, h: 4.0 }, layout: { x: 0.03, y: 0.0, w: 0.94, h: 0.88 }, fmt: '#,##0;-#,##0;;',
      valor: (m, atual) => atual ? m.yA : m.yB, fmtTxt: (x) => nf0.format(x), totais: (b) => nf0.format(b)
    });
    varLado(s1, y1.alvos, cy1, fPct, (n) => R.m[n].vY, 6.85);
    // tabela de variações mês a mês
    const H = { fontFace: FONT, fontSize: 8, align: 'center', valign: 'middle', margin: 0, color: '7F7F7F', fill: { color: 'FFFFFF' } };
    const ordemTab = ['Bondinho', CRISTO, 'Trem Corcovado', 'Paineiras'];
    const brd = { type: 'solid', pt: 1, color: 'FFFFFF' };
    const rows = [[{ text: '', options: H }].concat(jan.map((j) => ({ text: j.label, options: H })))];
    ordemTab.forEach((a) => {
      rows.push([{ text: a, options: { fontFace: FONT, fontSize: 8, bold: true, align: 'center', valign: 'middle', margin: 0, color: TXT[a], fill: { color: 'FFFFFF' }, border: brd } }]
        .concat(R.yoy[a].map((x) => ({ text: x == null ? '–' : (r100(x) === 0 ? 0 : r100(x)) + '%', options: { fontFace: FONT, fontSize: 8, bold: true, align: 'center', valign: 'middle', margin: 0, color: 'FFFFFF', fill: { color: COR[a] }, border: brd } }))));
    });
    s1.addTable(rows, { x: 0, y: 5.98, w: 8.37, colW: [1.0].concat(jan.map(() => (8.37 - 1.0) / jan.length)), rowH: 0.224, autoPage: false });

    /* ===== slide 2: Market share ===== */
    const s2 = pptx.addSlide();
    moldura(pptx, s2, R, textoShare(R), { x: 0.44, y: 1.33, w: 12.47, h: 0.5 });
    cartao(pptx, s2, 0.36, 2.02, 8.3, 4.69);
    const l2 = linhas(pptx, s2, R, {
      vals: R.share, passo: 0.1, fmt: '0%', titulo: 'MKT SHARE',
      frame: { x: 0.46, y: 2.16, w: 7.42, h: 4.39 }, layout: { x: 0.03, y: 0.14, w: 0.92, h: 0.62 }
    });
    const col2 = 8.26;
    rotulo(s2, `VAR. ${R.mes3.toUpperCase()}\n${yy(R.A)}x${yy(R.B)}`, col2 - 0.55, 2.16, 1.1, 0.37, { fontSize: 8, color: '404040' });
    varLado(s2, l2.fim.slice(), col2, fPP, (n) => R.m[n].dM, 6.4);
    cartao(pptx, s2, 8.95, 2.02, 4.02, 4.69);
    s2.addText('MKT SHARE  YTD', { x: 9.4, y: 2.14, w: 2.6, h: 0.3, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 12, bold: true, color: '595959', margin: 0, isTextBox: true });
    const cy2 = 12.4;
    rotulo(s2, `VAR. YTD\n${yy(R.A)}x${yy(R.B)}`, cy2 - 0.65, 2.22, 1.3, 0.37, { fontSize: 8, color: '404040' });
    const y2 = ytd(pptx, s2, R, {
      frame: { x: 9.2, y: 2.62, w: 2.9, h: 3.95 }, layout: { x: 0.03, y: 0.0, w: 0.94, h: 0.86 }, fmt: '0%;-0%;;', max: 1,
      valor: (m, atual) => atual ? m.ysA : m.ysB, fmtTxt: (x) => pInt(x)
    });
    varLado(s2, y2.alvos, cy2, fPP, (n) => R.m[n].dY, 6.55);

    // ---- gera e ajusta o XML dos gráficos de linha ----
    const buf = await pptx.write({ outputType: 'arraybuffer', compression: true });
    const zip = await window.JSZip.loadAsync(buf);
    const arqs = Object.keys(zip.files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f)).sort((a, b) => parseInt(a.match(/(\d+)\.xml/)[1], 10) - parseInt(b.match(/(\d+)\.xml/)[1], 10));
    let k = 0;
    for (const f of arqs) {
      const xml = await zip.file(f).async('string');
      if (xml.indexOf('<c:lineChart>') < 0) continue;
      const pl = LINHAS[k++]; if (!pl) continue;
      zip.file(f, ajustarRotulos(xml, pl.plano, pl.fmt, pl.nulos).replace('<c:dispBlanksAs val="span"/>', '<c:dispBlanksAs val="gap"/>'));
    }
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }

  async function baixar(R) {
    const blob = await montar(R);
    const nome = `Report_Pax_Atrativos_Rio_de_Janeiro_${cap(R.mes3)}${yy(R.A)}.pptx`;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = nome; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    return nome;
  }
  window.PaxExport = { baixar, montar, textoPax, textoShare };
})();
