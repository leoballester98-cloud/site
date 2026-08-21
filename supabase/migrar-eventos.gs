/* ═══════════════════════════════════════════════════════════════════════════
   Migra o histórico da aba Eventos pro Supabase.

   Roda em lotes e guarda onde parou nas Propriedades do script, então:
   - se estourar o limite de 6 minutos do Apps Script, é só rodar de novo que
     continua de onde ficou
   - rodar duas vezes por engano NÃO duplica: ele só manda o que ainda não foi

   Usa a chave publishable, a mesma do quiz. O RLS só permite INSERT em eventos
   com ela — é exatamente o que esta migração precisa.
   ═══════════════════════════════════════════════════════════════════════════ */

const SUPA_URL = 'https://lfmjvgkbvkkexwytkfeg.supabase.co';
const SUPA_KEY = 'sb_publishable_YMHtH3C0VuB5X8BFXADwBw_E7FcBFgE';
const LOTE = 500;

function migrarEventos() {
  const props = PropertiesService.getScriptProperties();
  const sh = getSheet_();
  const ultima = sh.getLastRow();
  let linha = Number(props.getProperty('MIGRACAO_LINHA')) || 2;

  if (linha > ultima) {
    return 'Nada a fazer: já migrou até a linha ' + (linha - 1) + ' de ' + ultima + '.';
  }

  const larg = Math.max(5, Math.min(6, sh.getLastColumn()));
  let enviados = 0, pulados = 0, lotes = 0;
  const inicio = Date.now();

  while (linha <= ultima) {
    // para antes dos 6 min pra não morrer no meio de um lote
    if (Date.now() - inicio > 4.5 * 60 * 1000) {
      props.setProperty('MIGRACAO_LINHA', String(linha));
      return 'Parcial: ' + enviados + ' enviados, parou na linha ' + linha +
             ' de ' + ultima + '. Rode de novo pra continuar.';
    }

    const qtd = Math.min(LOTE, ultima - linha + 1);
    const vals = sh.getRange(linha, 1, qtd, larg).getValues();
    const corpo = [];

    vals.forEach(function (r) {
      const sessao = String(r[1] || '').trim();
      const tipo = String(r[2] || 'etapa').trim();
      // as constraints do banco recusariam o lote inteiro por causa de uma linha ruim
      if (sessao.length < 4 || sessao.length > 64) { pulados++; return; }
      if (tipo !== 'etapa' && tipo !== 'respostas') { pulados++; return; }

      const etapa = parseInt(r[3], 10);
      if (tipo === 'etapa' && (isNaN(etapa) || etapa < 1 || etapa > 40)) { pulados++; return; }

      const ts = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
      if (isNaN(ts.getTime())) { pulados++; return; }

      const extra = String(r[4] || '').trim();
      const linhaJson = {
        ts: ts.toISOString(),
        sessao: sessao,
        tipo: tipo,
        etapa: isNaN(etapa) ? null : etapa,
        pagina: String(r[5] || '').trim() || null,
        variante: null,
        dados: null
      };
      if (tipo === 'respostas') {
        try { linhaJson.dados = JSON.parse(extra); } catch (e) { linhaJson.dados = null; }
      } else if (extra) {
        linhaJson.variante = extra.slice(0, 24);
      }
      corpo.push(linhaJson);
    });

    if (corpo.length) {
      const resp = UrlFetchApp.fetch(SUPA_URL + '/rest/v1/eventos', {
        method: 'post',
        contentType: 'application/json',
        headers: { apikey: SUPA_KEY, Prefer: 'return=minimal' },
        payload: JSON.stringify(corpo),
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code < 200 || code >= 300) {
        props.setProperty('MIGRACAO_LINHA', String(linha));
        throw new Error('Falhou na linha ' + linha + ' (HTTP ' + code + '): ' + resp.getContentText().slice(0, 300));
      }
      enviados += corpo.length;
    }

    linha += qtd;
    lotes++;
    props.setProperty('MIGRACAO_LINHA', String(linha));
  }

  const msg = 'Concluído: ' + enviados + ' eventos enviados em ' + lotes + ' lotes. ' +
              'Linhas puladas por dado inválido: ' + pulados + '. Total na planilha: ' + (ultima - 1) + '.';
  Logger.log(msg);
  return msg;
}

/* Se precisar recomeçar do zero: apague os eventos migrados no Supabase e rode
   isto antes de migrar de novo. */
function resetarMigracao() {
  PropertiesService.getScriptProperties().deleteProperty('MIGRACAO_LINHA');
  return 'Contador zerado. A próxima migração começa da primeira linha.';
}
