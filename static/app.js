// ── Estado ───────────────────────────────────────────────
let DATA = [];

const estado = {
  unidade:     null,
  unidadeRows: [],
  departamento: null,
  depRows:     [],
  abaAtiva:    'Em Atraso',
};

const filtrosRel = { coordenador: '', responsavel: '', grupo: '', busca: '' };
const filtrosTab = { coordenador: '', responsavel: '', grupo: '', busca: '' };
let buscaTimerRel = null;
let buscaTimerTab = null;

const NOME_UNIDADE = {
  'SP':     'São Paulo',
  'GOIAS':  'Goiás',
  'Santos': 'Santos',
  'RJ':     'Rio de Janeiro',
};

// ── Utilitários ───────────────────────────────────────────
function esc(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fmtData(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return v.toLocaleDateString('pt-BR');
  }
  return String(v);
}

function contarStatus(rows) {
  let emAtraso = 0, emAberto = 0, baixadas = 0;
  for (const r of rows) {
    if (r.Status === 'Em Atraso') emAtraso++;
    else if (r.Status === 'Em Aberto') emAberto++;
    else if (r.Status === 'Baixado')   baixadas++;
  }
  return { total: rows.length, emAtraso, emAberto, baixadas };
}

// ── Carregamento SheetJS ──────────────────────────────────
async function carregarDados() {
  try {
    const resp = await fetch('dados/base/base.xlsx');
    const buf  = await resp.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws   = wb.Sheets['Pendencias'];
    DATA = XLSX.utils.sheet_to_json(ws, { raw: true });

    document.getElementById('loading').style.display = 'none';
    renderCards();
  } catch (e) {
    document.getElementById('loading').innerHTML =
      '<div class="loading-box"><p style="color:#C00000">Erro ao carregar base de dados.<br>Verifique o arquivo e recarregue.</p></div>';
  }
}

// ── Breadcrumb ────────────────────────────────────────────
function setBreadcrumb(partes) {
  const bar = document.getElementById('unidade-bar');
  if (!partes.length) { bar.classList.add('hidden'); return; }

  bar.classList.remove('hidden');
  document.getElementById('bar-crumbs').innerHTML = partes
    .map((p, i) => (i === 0
      ? `<span class="unidade-crumb" onclick="mostrarTela1()">${esc(p)}</span>`
      : `<span class="unidade-sep">›</span><span class="unidade-nome">${esc(p)}</span>`
    ))
    .join('');
}

// ── Placares ─────────────────────────────────────────────
function renderPlacares(containerId, rows) {
  const c = contarStatus(rows);
  const defs = [
    { classe: 'total',   valor: c.total,    label: 'Total de Tarefas', desc: 'da competência' },
    { classe: 'atraso',  valor: c.emAtraso, label: 'Em Atraso',        desc: 'tarefas' },
    { classe: 'aberto',  valor: c.emAberto, label: 'Em Aberto',        desc: 'tarefas' },
    { classe: 'baixado', valor: c.baixadas, label: 'Baixadas',         desc: 'tarefas' },
  ];

  document.getElementById(containerId).innerHTML = defs.map(p => `
    <div class="placar ${p.classe}">
      <div class="placar-label">${p.label}</div>
      <div class="placar-valor">${p.valor.toLocaleString('pt-BR')}</div>
      <div class="placar-desc">${p.desc}</div>
    </div>
  `).join('');
}

// ── Tela 1 ───────────────────────────────────────────────
function renderCards() {
  const grid = document.getElementById('cards-grid');
  const unidadeCods = [...new Set(DATA.map(r => r.Unidade).filter(Boolean))].sort();

  grid.innerHTML = unidadeCods.map(cod => {
    const nome = NOME_UNIDADE[cod] || cod;
    const rows = DATA.filter(r => r.Unidade === cod);
    const c    = contarStatus(rows);
    return `
      <div class="card" onclick="mostrarTela2('${cod}', '${esc(nome)}')">
        <div class="card-unit">${esc(nome)}</div>
        <div class="card-info">
          <span class="card-count">${c.total.toLocaleString('pt-BR')}</span>
          <span class="card-label">- Tarefas</span>
        </div>
        <div class="card-footer">Clique para ver os detalhes</div>
      </div>`;
  }).join('');
}

// ── Tela 2 ───────────────────────────────────────────────
function mostrarTela2(cod, nome) {
  estado.unidade     = { cod, nome };
  estado.unidadeRows = DATA.filter(r => r.Unidade === cod);

  setBreadcrumb(['Painel de controle', nome]);
  renderPlacares('placares-tela2', estado.unidadeRows);
  renderDepartamentos(estado.unidadeRows);
  trocarTela('tela-2');
}

function renderDepartamentos(rows) {
  const deps = [...new Set(rows.map(r => r.Departamento).filter(Boolean))].sort();

  document.getElementById('deps-grid').innerHTML = deps.map(dep => {
    const dr = rows.filter(r => r.Departamento === dep);
    const c  = contarStatus(dr);
    return `
      <div class="dep-card" onclick="mostrarTela3('${esc(dep)}')">
        <div class="dep-nome">${esc(dep)}</div>
        <div class="dep-linhas">
          <div class="dep-linha">
            <span class="dep-dot atraso"></span>
            <span class="dep-linha-texto">Em Atraso</span>
            <span class="dep-linha-valor">${c.emAtraso.toLocaleString('pt-BR')}</span>
          </div>
          <div class="dep-linha">
            <span class="dep-dot aberto"></span>
            <span class="dep-linha-texto">Em Aberto</span>
            <span class="dep-linha-valor">${c.emAberto.toLocaleString('pt-BR')}</span>
          </div>
          <div class="dep-linha">
            <span class="dep-dot baixado"></span>
            <span class="dep-linha-texto">Baixadas</span>
            <span class="dep-linha-valor">${c.baixadas.toLocaleString('pt-BR')}</span>
          </div>
        </div>
        <div class="dep-footer">Clique aqui para ver os detalhes</div>
      </div>`;
  }).join('');
}

// ── Relatórios ────────────────────────────────────────────
function inicioSemana() {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const diasDesdeSegunda = hoje.getDay() === 0 ? 6 : hoje.getDay() - 1;
  const seg = new Date(hoje);
  seg.setDate(hoje.getDate() - diasDesdeSegunda);
  return seg;
}

function normData(d) {
  const n = new Date(d);
  n.setHours(0, 0, 0, 0);
  return n;
}

function rankColaboradores(rows) {
  const contagem = {};
  for (const r of rows) {
    const nome = String(r.UsuarioResponsavel || '').trim();
    if (!nome) continue;
    contagem[nome] = (contagem[nome] || 0) + 1;
  }
  return Object.entries(contagem)
    .sort((a, b) => b[1] - a[1])
    .map(([nome, qtd]) => ({ nome, qtd }));
}

function renderRelatorios(rows) {
  const agora  = new Date();
  const hoje   = normData(agora);
  const semana = inicioSemana();
  const mes    = new Date(agora.getFullYear(), agora.getMonth(), 1);
  const nomeMes = agora.toLocaleDateString('pt-BR', { month: 'long' });

  const rowsHoje   = rows.filter(r => r.DataBaixa instanceof Date && normData(r.DataBaixa) >= hoje);
  const rowsSemana = rows.filter(r => r.DataBaixa instanceof Date && normData(r.DataBaixa) >= semana);
  const rowsMes    = rows.filter(r => r.DataBaixa instanceof Date && normData(r.DataBaixa) >= mes);

  const defs = [
    { classe: 'hoje',   label: 'Baixadas hoje',          rows: rowsHoje,   desc: agora.toLocaleDateString('pt-BR') },
    { classe: 'semana', label: 'Baixadas na semana',     rows: rowsSemana, desc: 'segunda-feira até hoje' },
    { classe: 'mes',    label: `Baixadas em ${nomeMes}`, rows: rowsMes,    desc: String(agora.getFullYear()) },
  ];

  document.getElementById('baixadas-grid').innerHTML = defs.map(d => {
    const ranking = rankColaboradores(d.rows);
    const max = ranking[0]?.qtd || 1;

    const rankHTML = ranking.length
      ? ranking.map(item => `
          <div class="rank-item">
            <span class="rank-nome" title="${esc(item.nome)}">${esc(item.nome)}</span>
            <div class="rank-barra-wrapper">
              <div class="rank-barra" style="width:${(item.qtd / max * 100).toFixed(1)}%"></div>
            </div>
            <span class="rank-qtd">${item.qtd}</span>
          </div>`).join('')
      : '<span class="rank-vazio">Nenhum registro no período</span>';

    return `
      <div class="baixada-card ${d.classe}">
        <div class="baixada-card-topo">
          <div class="baixada-card-label">${d.label}</div>
          <div class="baixada-card-valor">${d.rows.length.toLocaleString('pt-BR')}</div>
          <div class="baixada-card-desc">${d.desc}</div>
        </div>
        <div class="baixada-ranking">${rankHTML}</div>
      </div>`;
  }).join('');
}

// ── Geração de PDF ───────────────────────────────────────
async function gerarPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const PW = doc.internal.pageSize.getWidth();
  const PH = doc.internal.pageSize.getHeight();

  // ── Semanas do mês atual (blocos de 7 dias a partir do dia 1) ──
  const agora  = new Date();
  const ano    = agora.getFullYear();
  const mes    = agora.getMonth();
  const diasMes = new Date(ano, mes + 1, 0).getDate();
  const hoje   = agora.getDate();

  const semanas = [];
  for (let ini = 1; ini <= diasMes && ini <= hoje; ini += 7) {
    const fim = Math.min(ini + 6, diasMes, hoje);
    const label = `Semana ${semanas.length + 1}\n${String(ini).padStart(2,'0')}–${String(fim).padStart(2,'0')}/${String(mes+1).padStart(2,'0')}`;
    semanas.push({ ini, fim, label });
  }

  // ── Dados filtrados (baixadas do mês corrente) ──
  const base = filtrarRows(
    estado.depRows.filter(r =>
      r.Status === 'Baixado' &&
      r.DataBaixa instanceof Date &&
      r.DataBaixa.getFullYear() === ano &&
      r.DataBaixa.getMonth() === mes
    ),
    filtrosRel
  );

  // ── Agrupamento por funcionário ──
  const empMap = {};
  for (const r of base) {
    const nome  = String(r.UsuarioResponsavel || '').trim();
    const coord = String(r.Coordenador || '').trim();
    if (!nome) continue;
    if (!empMap[nome]) empMap[nome] = { nome, coord, qtds: new Array(semanas.length).fill(0), total: 0 };
    const dia = r.DataBaixa.getDate();
    const idx = semanas.findIndex(s => dia >= s.ini && dia <= s.fim);
    if (idx >= 0) empMap[nome].qtds[idx]++;
    empMap[nome].total++;
  }

  const empRows = Object.values(empMap).sort((a, b) => b.total - a.total || a.nome.localeCompare(b.nome, 'pt-BR'));
  const totSem  = semanas.map((_, i) => empRows.reduce((s, e) => s + e.qtds[i], 0));
  const totGeral = empRows.reduce((s, e) => s + e.total, 0);

  // ── Logo (tentativa; silencioso se falhar) ──
  let logoB64 = null;
  try {
    const r = await fetch('dados/imagens/logo.png');
    const blob = await r.blob();
    logoB64 = await new Promise(res => {
      const fr = new FileReader();
      fr.onloadend = () => res(fr.result);
      fr.readAsDataURL(blob);
    });
  } catch (_) {}

  // ── Cabeçalho (todas as páginas via didDrawPage) ──
  const nomeMes = agora.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
  const dataGer = agora.toLocaleDateString('pt-BR');
  const subHead = `${estado.unidade.nome}  —  ${estado.departamento}`;
  const filtrosAtivos = [
    filtrosRel.coordenador && `Coordenação: ${filtrosRel.coordenador}`,
    filtrosRel.responsavel && `Responsável: ${filtrosRel.responsavel}`,
    filtrosRel.grupo       && `Grupo: ${filtrosRel.grupo}`,
    filtrosRel.busca       && `Busca: "${filtrosRel.busca}"`,
  ].filter(Boolean);
  const filtrosStr = filtrosAtivos.length ? filtrosAtivos.join('  |  ') : 'Sem filtros ativos';

  function desenharCabecalho(pageNum) {
    // Faixa vermelha
    doc.setFillColor(192, 0, 0);
    doc.rect(0, 0, PW, 18, 'F');

    // Logo
    if (logoB64) doc.addImage(logoB64, 'PNG', 6, 3.5, 14, 10);

    // Título
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('MG Contécnica — Relatório de Baixas', logoB64 ? 24 : 10, 11.5);

    // Sub-cabeçalho: unidade — departamento
    doc.setTextColor(18, 18, 18);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(subHead, 10, 23);

    // Linha decorativa fina vermelho-suave
    doc.setDrawColor(192, 0, 0);
    doc.setLineWidth(0.35);
    doc.line(10, 25.8, PW - 10, 25.8);

    // Período + filtros numa linha só
    doc.setTextColor(110, 110, 110);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'normal');
    const periodoInfo = `${nomeMes}  ·  Gerado em ${dataGer}  ·  Filtros: ${filtrosStr}`;
    doc.text(periodoInfo, 10, 30.5);
  }

  function desenharRodape(pageNum, totalPages) {
    doc.setFontSize(7);
    doc.setTextColor(160, 160, 160);
    doc.setFont('helvetica', 'normal');
    doc.text('MG Contécnica — Relatório Confidencial', 10, PH - 5);
    doc.text(`Página ${pageNum} de ${totalPages}`, PW - 10, PH - 5, { align: 'right' });
  }

  // ── Colunas da tabela ──
  const cols = [
    { header: 'Funcionário', dataKey: 'nome' },
    { header: 'Coordenador', dataKey: 'coord' },
    ...semanas.map((s, i) => ({ header: s.label, dataKey: `s${i}` })),
    { header: 'Total', dataKey: 'total' },
  ];

  const tableBody = empRows.map(e => {
    const row = { nome: e.nome, coord: e.coord, total: e.total };
    e.qtds.forEach((v, i) => { row[`s${i}`] = v > 0 ? v : '—'; });
    return row;
  });
  // Linha de totais
  const totalRow = { nome: 'TOTAL', coord: '', total: totGeral };
  totSem.forEach((v, i) => { totalRow[`s${i}`] = v > 0 ? v : '—'; });
  tableBody.push(totalRow);

  // ── AutoTable ──
  doc.autoTable({
    startY: 34,
    columns: cols,
    body: tableBody,
    styles: { fontSize: 8, textColor: [18, 18, 18], cellPadding: { top: 3, right: 4, bottom: 3, left: 4 } },
    headStyles: {
      fillColor: [26, 26, 26],
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      halign: 'center',
      valign: 'middle',
    },
    columnStyles: {
      nome:  { cellWidth: 52 },
      coord: { cellWidth: 48 },
      total: { cellWidth: 18, halign: 'center', fontStyle: 'bold' },
      ...Object.fromEntries(semanas.map((_, i) => [`s${i}`, { halign: 'center', cellWidth: 'auto' }])),
    },
    didParseCell(data) {
      // Linha de totais em negrito com fundo cinza
      if (data.row.index === tableBody.length - 1) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.fillColor = [235, 235, 235];
        data.cell.styles.textColor = [30, 30, 30];
      }
      // Linhas em tiras: par = cinza suave, ímpar = branco
      if (data.section === 'body' && data.row.index < tableBody.length - 1) {
        data.cell.styles.fillColor = data.row.index % 2 === 0 ? [244, 244, 248] : [255, 255, 255];
      }
    },
    didDrawPage(data) {
      desenharCabecalho(data.pageNumber);
    },
    margin: { top: 34, left: 10, right: 10, bottom: 12 },
  });

  // ── Página do gráfico de colunas ──
  const chartDados = [...empRows].filter(e => e.total > 0).sort((a, b) => b.total - a.total);
  if (chartDados.length > 0) {
    doc.addPage();
    desenharCabecalho(doc.internal.getNumberOfPages());

    const mL = 22;
    const mR = 12;
    const cTop = 40;
    const cLabH = 34;
    const cH = PH - cTop - cLabH - 14;
    const cW = PW - mL - mR;
    const n = chartDados.length;
    const slotW = cW / n;
    const barW = Math.min(18, Math.max(4, slotW * 0.58));
    const maxVal = Math.max(...chartDados.map(d => d.total));
    const yMax = Math.ceil(maxVal / 5) * 5 || 5;
    const yScale = cH / yMax;

    // Título do gráfico
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(30, 30, 30);
    doc.text(`Baixas por Colaborador — ${nomeMes}`, mL + cW / 2, cTop - 5, { align: 'center' });

    // Linhas de grade e rótulos do eixo Y
    const gridSteps = 5;
    for (let i = 0; i <= gridSteps; i++) {
      const val = Math.round(yMax * i / gridSteps);
      const gy = cTop + cH - val * yScale;
      doc.setDrawColor(210, 210, 210);
      doc.setLineWidth(0.2);
      doc.line(mL, gy, mL + cW, gy);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6);
      doc.setTextColor(120, 120, 120);
      doc.text(String(val), mL - 2, gy + 1.5, { align: 'right' });
    }

    // Linhas dos eixos
    doc.setDrawColor(130, 130, 130);
    doc.setLineWidth(0.5);
    doc.line(mL, cTop, mL, cTop + cH);
    doc.line(mL, cTop + cH, mL + cW, cTop + cH);

    // Barras + valores + rótulos
    chartDados.forEach((d, i) => {
      const barH = Math.max(0.5, d.total * yScale);
      const bx = mL + i * slotW + (slotW - barW) / 2;
      const by = cTop + cH - barH;

      doc.setFillColor(192, 0, 0);
      doc.rect(bx, by, barW, barH, 'F');

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6);
      doc.setTextColor(40, 40, 40);
      doc.text(String(d.total), bx + barW / 2, by - 1.5, { align: 'center' });

      const partes = d.nome.trim().split(' ');
      const label = partes.length > 1 ? `${partes[0]} ${partes[partes.length - 1]}` : partes[0];
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(50, 50, 50);
      doc.text(label, bx + barW / 2, cTop + cH + 2, { angle: -45 });
    });
  }

  // ── Rodapés (passa por todas as páginas após geração) ──
  const totalPages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    desenharRodape(p, totalPages);
  }

  // ── Download ──
  const depSlug = estado.departamento.replace(/[^a-z0-9]/gi, '_');
  doc.save(`relatorio_baixas_${depSlug}_${ano}-${String(mes+1).padStart(2,'0')}.pdf`);
}

// ── Tela 3 ───────────────────────────────────────────────
function mostrarTela3(dep) {
  estado.departamento = dep;
  estado.depRows      = estado.unidadeRows.filter(r => r.Departamento === dep);
  estado.abaAtiva     = 'Em Atraso';

  Object.assign(filtrosRel, { coordenador: '', responsavel: '', grupo: '', busca: '' });
  Object.assign(filtrosTab, { coordenador: '', responsavel: '', grupo: '', busca: '' });
  popularFiltrosPrefixo('rel', estado.depRows);
  popularFiltrosPrefixo('tab', estado.depRows);

  setBreadcrumb(['Painel de controle', estado.unidade.nome, dep]);
  renderPlacares('placares-tela3', estado.depRows);
  renderRelatorios(estado.depRows);

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('ativo'));
  document.querySelector('.tab[data-status="Em Atraso"]').classList.add('ativo');

  renderTabela('Em Atraso');
  trocarTela('tela-3');
}

function voltarTela2() {
  mostrarTela2(estado.unidade.cod, estado.unidade.nome);
}

function mudarAba(btn) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('ativo'));
  btn.classList.add('ativo');
  estado.abaAtiva = btn.dataset.status;
  renderTabela(estado.abaAtiva);
}

// ── Filtros ───────────────────────────────────────────────
function popularSelect(id, valores) {
  const el = document.getElementById(id);
  const atual = el.value;
  el.innerHTML = '<option value="">Todos</option>' +
    valores.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if (valores.includes(atual)) el.value = atual;
}

function popularFiltrosPrefixo(prefixo, rows) {
  const sort = arr => arr.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
  popularSelect(`${prefixo}-coordenador`, sort([...new Set(rows.map(r => r.Coordenador).filter(Boolean))]));
  popularSelect(`${prefixo}-responsavel`, sort([...new Set(rows.map(r => r.UsuarioResponsavel).filter(Boolean))]));
  popularSelect(`${prefixo}-grupo`,       sort([...new Set(rows.map(r => r.Grupo).filter(Boolean))]));
}

function lerFiltros(prefixo, obj) {
  obj.coordenador = document.getElementById(`${prefixo}-coordenador`).value;
  obj.responsavel = document.getElementById(`${prefixo}-responsavel`).value;
  obj.grupo       = document.getElementById(`${prefixo}-grupo`).value;
  obj.busca       = document.getElementById(`${prefixo}-busca`).value.trim().toLowerCase();
}

function filtrarRows(rows, f) {
  return rows.filter(r => {
    if (f.coordenador && String(r.Coordenador || '')        !== f.coordenador) return false;
    if (f.responsavel && String(r.UsuarioResponsavel || '') !== f.responsavel) return false;
    if (f.grupo       && String(r.Grupo || '')              !== f.grupo)       return false;
    if (f.busca) {
      const hay = [r.CodCliente, r.RazaoSocial, r.Grupo, r.Titulo, r.UsuarioResponsavel, r.Coordenador, r.Comentario]
        .map(v => String(v || '').toLowerCase()).join('\0');
      if (!hay.includes(f.busca)) return false;
    }
    return true;
  });
}

function aplicarFiltroRel() {
  lerFiltros('rel', filtrosRel);
  renderRelatorios(filtrarRows(estado.depRows, filtrosRel));
}

function aplicarFiltroTab() {
  lerFiltros('tab', filtrosTab);
  renderTabela(estado.abaAtiva);
}

function agendarBuscaRel() {
  clearTimeout(buscaTimerRel);
  buscaTimerRel = setTimeout(aplicarFiltroRel, 200);
}

function agendarBuscaTab() {
  clearTimeout(buscaTimerTab);
  buscaTimerTab = setTimeout(aplicarFiltroTab, 200);
}

function limparFiltros(prefixo) {
  const obj = prefixo === 'rel' ? filtrosRel : filtrosTab;
  [`${prefixo}-coordenador`, `${prefixo}-responsavel`, `${prefixo}-grupo`]
    .forEach(id => document.getElementById(id).value = '');
  document.getElementById(`${prefixo}-busca`).value = '';
  obj.coordenador = ''; obj.responsavel = ''; obj.grupo = ''; obj.busca = '';
  if (prefixo === 'rel') renderRelatorios(estado.depRows);
  else renderTabela(estado.abaAtiva);
}

function sortRows(rows) {
  return [...rows].sort((a, b) => {
    const tA = a.DataVencimento instanceof Date ? a.DataVencimento.getTime() : 0;
    const tB = b.DataVencimento instanceof Date ? b.DataVencimento.getTime() : 0;
    if (tA !== tB) return tA - tB;
    const cA = String(a.RazaoSocial || a.CodCliente || '').toLowerCase();
    const cB = String(b.RazaoSocial || b.CodCliente || '').toLowerCase();
    return cA.localeCompare(cB, 'pt-BR');
  });
}

function renderTabela(status) {
  const rows = sortRows(filtrarRows(estado.depRows.filter(r => r.Status === status), filtrosTab));
  const corpo = document.getElementById('tabela-corpo');

  if (!rows.length) {
    corpo.innerHTML = `<tr><td colspan="7" class="tabela-vazia">Nenhuma tarefa encontrada.</td></tr>`;
    return;
  }

  corpo.innerHTML = rows.map((r, idx) => {
    const codCliente = esc(r.CodCliente || '');
    const razao      = esc(r.RazaoSocial || '');
    const cliente    = codCliente && razao ? `${codCliente} - ${razao}` : codCliente || razao;

    const cmt = r.Comentario ? String(r.Comentario).trim() : '';
    const colCmt = cmt
      ? `<button class="btn-comentario" onclick="toggleComentario(this,'cmt-${idx}')">Ver comentário</button>
         <span class="comentario-texto hidden" id="cmt-${idx}">${esc(cmt)}</span>`
      : '-';

    return `<tr>
      <td>${cliente}</td>
      <td>${esc(r.Grupo)}</td>
      <td>${esc(r.Titulo)}</td>
      <td>${esc(r.UsuarioResponsavel)}</td>
      <td>${fmtData(r.DataVencimento)}</td>
      <td>${fmtData(r.DataPrevisaoConclusao)}</td>
      <td>${colCmt}</td>
    </tr>`;
  }).join('');
}

function toggleComentario(btn, id) {
  const span = document.getElementById(id);
  const aberto = !span.classList.contains('hidden');
  span.classList.toggle('hidden');
  btn.textContent = aberto ? 'Ver comentário' : 'Ocultar';
}

// ── Navegação ─────────────────────────────────────────────
function trocarTela(id) {
  ['tela-1', 'tela-2', 'tela-3'].forEach(t =>
    document.getElementById(t).classList.toggle('hidden', t !== id)
  );
  if (id === 'tela-1') setBreadcrumb([]);
  window.scrollTo(0, 0);
}

function mostrarTela1() {
  trocarTela('tela-1');
}

// ── Init ─────────────────────────────────────────────────
carregarDados();
