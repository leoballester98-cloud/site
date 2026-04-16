// ============================================================
// COLE ESSE CÓDIGO NO GOOGLE APPS SCRIPT
// Instruções de como fazer isso estão no README abaixo
// ============================================================

const NOMES_ETAPAS = {
  1:  'Etapa 1 — Início do Quiz',
  2:  'Etapa 2 — Pergunta 1',
  3:  'Etapa 3 — Pergunta 2',
  4:  'Etapa 4 — Pergunta 3',
  5:  'Etapa 5 — Pergunta 4',
  6:  'Etapa 6 — Pergunta 5',
  7:  'Etapa 7 — Pergunta 6',
  8:  'Etapa 8 — Pergunta 7',
  9:  'Etapa 9 — Pergunta 8',
  10: 'Etapa 10 — Pergunta 9',
  11: 'Etapa 11 — Carregando Análise',
  12: 'Etapa 12 — Diagnóstico',
  13: 'Etapa 13 — Carregando Resultado',
  14: 'Etapa 14 — Resultado',
  15: 'Etapa 15 — Resultado Final',
  16: 'Etapa 16 — Carregando Oferta',
  17: 'Etapa 17 — Página de Vendas'
};

// Recebe os dados do quiz e salva na planilha
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Aba de dados brutos
    let abaEventos = ss.getSheetByName('Eventos');
    if (!abaEventos) {
      abaEventos = ss.insertSheet('Eventos');
      abaEventos.appendRow(['Data', 'Hora', 'Sessão', 'Nº Etapa', 'Nome da Etapa']);
      abaEventos.setFrozenRows(1);
      abaEventos.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
    }

    const agora = new Date();
    const data = Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy');
    const hora = Utilities.formatDate(agora, 'America/Sao_Paulo', 'HH:mm:ss');

    abaEventos.appendRow([
      data,
      hora,
      dados.sessao || 'desconhecido',
      dados.etapa,
      NOMES_ETAPAS[dados.etapa] || ('Etapa ' + dados.etapa)
    ]);

    // Atualiza o dashboard após cada novo evento
    atualizarDashboard(ss);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Atualiza a aba Dashboard com contagens e % de conversão
function atualizarDashboard(ss) {
  let abaDash = ss.getSheetByName('Dashboard');
  if (!abaDash) {
    abaDash = ss.insertSheet('Dashboard');
  }
  abaDash.clearContents();

  const abaEventos = ss.getSheetByName('Eventos');
  const totalLinhas = abaEventos.getLastRow() - 1; // descontar cabeçalho
  if (totalLinhas <= 0) return;

  // Contar sessões únicas por etapa
  const dados = abaEventos.getRange(2, 1, totalLinhas, 5).getValues();

  const contagemPorEtapa = {};
  const sessoesUnicasPorEtapa = {};

  for (let i = 1; i <= 17; i++) {
    sessoesUnicasPorEtapa[i] = new Set();
  }

  dados.forEach(linha => {
    const sessao = linha[2];
    const etapa  = linha[3];
    if (etapa && sessoesUnicasPorEtapa[etapa]) {
      sessoesUnicasPorEtapa[etapa].add(sessao);
    }
  });

  // Montar tabela no Dashboard
  abaDash.appendRow(['Nº', 'Etapa', 'Usuários que chegaram', '% em relação ao início', '% em relação à etapa anterior']);
  abaDash.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  abaDash.setFrozenRows(1);

  const totalInicio = sessoesUnicasPorEtapa[1].size || 1;
  let anteriorCount = totalInicio;

  for (let i = 1; i <= 17; i++) {
    const count = sessoesUnicasPorEtapa[i].size;
    const pctInicio   = count > 0 ? Math.round((count / totalInicio) * 100) + '%' : '0%';
    const pctAnterior = count > 0 && anteriorCount > 0 ? Math.round((count / anteriorCount) * 100) + '%' : '0%';

    abaDash.appendRow([
      i,
      NOMES_ETAPAS[i],
      count,
      pctInicio,
      pctAnterior
    ]);

    anteriorCount = count > 0 ? count : anteriorCount;
  }

  // Formatar colunas
  abaDash.autoResizeColumns(1, 5);
}

// Roda quando você abre a planilha — atualiza o dashboard manualmente também
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Quiz Funil')
    .addItem('Atualizar Dashboard', 'atualizarDashboardManual')
    .addToUi();
}

function atualizarDashboardManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  atualizarDashboard(ss);
  SpreadsheetApp.getUi().alert('Dashboard atualizado!');
}
