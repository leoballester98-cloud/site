// ============================================================
// COLE ESSE CÓDIGO NO GOOGLE APPS SCRIPT
// Após colar: Implantar → Gerenciar implantações → Editar → Nova versão → Implantar
// ============================================================

const NOMES_ETAPAS = {
  1:  'Início do Quiz',
  2:  'Pergunta 1',
  3:  'Pergunta 2',
  4:  'Pergunta 3',
  5:  'Pergunta 4',
  6:  'Pergunta 5',
  7:  'Pergunta 6',
  8:  'Pergunta 7',
  9:  'Pergunta 8',
  10: 'Pergunta 9',
  11: 'Carregando Análise',
  12: 'Diagnóstico',
  13: 'Carregando Resultado',
  14: 'Resultado',
  15: 'Resultado Final',
  16: 'Carregando Oferta',
  17: 'Página de Vendas'
};

function doGet(e) {
  try {
    const etapa = parseInt(e.parameter.etapa);
    if (!etapa || etapa < 1 || etapa > 17) {
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const aba  = obterOuCriarAba(ss);
    const linha = etapa + 1; // linha 1 = cabeçalho, linha 2 = etapa 1, etc.

    // Incrementa o contador da etapa
    const contadorAtual = aba.getRange(linha, 3).getValue() || 0;
    aba.getRange(linha, 3).setValue(contadorAtual + 1);

    // Recalcula as porcentagens
    recalcularPorcentagens(aba);

    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput('erro: ' + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function obterOuCriarAba(ss) {
  let aba = ss.getSheetByName('Funil do Quiz');

  if (!aba) {
    aba = ss.insertSheet('Funil do Quiz');

    // Cabeçalho
    aba.getRange(1, 1, 1, 5).setValues([['Nº', 'Etapa', 'Acessos', '% do Início', '% da Etapa Anterior']]);
    aba.getRange(1, 1, 1, 5)
      .setFontWeight('bold')
      .setBackground('#4a90d9')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    // Uma linha por etapa com contador zerado
    for (let i = 1; i <= 17; i++) {
      aba.getRange(i + 1, 1, 1, 5).setValues([[i, NOMES_ETAPAS[i], 0, '0%', '0%']]);
    }

    aba.setColumnWidth(1, 40);
    aba.setColumnWidth(2, 220);
    aba.setColumnWidth(3, 100);
    aba.setColumnWidth(4, 130);
    aba.setColumnWidth(5, 180);
    aba.setFrozenRows(1);
  }

  return aba;
}

function recalcularPorcentagens(aba) {
  const totalInicio = aba.getRange(2, 3).getValue() || 1; // acessos da etapa 1
  let anteriorCount = totalInicio;

  for (let i = 1; i <= 17; i++) {
    const linha = i + 1;
    const count = aba.getRange(linha, 3).getValue() || 0;

    const pctInicio   = Math.round((count / totalInicio) * 100) + '%';
    const pctAnterior = anteriorCount > 0 ? Math.round((count / anteriorCount) * 100) + '%' : '0%';

    aba.getRange(linha, 4).setValue(pctInicio);
    aba.getRange(linha, 5).setValue(pctAnterior);

    if (count > 0) anteriorCount = count;
  }

  // Colorir a coluna de % do início: vermelho → amarelo → verde
  for (let i = 1; i <= 17; i++) {
    const linha = i + 1;
    const pct   = aba.getRange(linha, 3).getValue() / (totalInicio || 1);
    const cor   = pct >= 0.7 ? '#b6d7a8' : pct >= 0.4 ? '#ffe599' : '#ea9999';
    aba.getRange(linha, 4).setBackground(cor);
  }
}
