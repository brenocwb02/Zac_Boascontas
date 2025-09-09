/**
 * @file Budget.gs
 * @description REFATORADO: Contém a lógica para gerir e atualizar a aba de Orçamento a partir do Hub.
 */

/**
 * **REFATORADO E CORRIGIDO:** Atualiza os valores gastos na aba 'Orcamento'.
 * A chamada 'toast' foi removida pois não há interface de utilizador no Hub.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} userSpreadsheet O objeto da planilha do cliente.
 */
function updateBudgetSpentValues(userSpreadsheet) {
  const ss = userSpreadsheet;
  const budgetSheet = ss.getSheetByName("Orcamento");
  const transacoesSheet = ss.getSheetByName("Transacoes");

  if (!budgetSheet || !transacoesSheet) {
    logToSheet(userSpreadsheet, "Aba 'Orcamento' ou 'Transacoes' não encontrada para atualização.", "WARN");
    return;
  }


  const budgetData = budgetSheet.getRange("A2:F" + budgetSheet.getLastRow()).getValues();
  const transacoesData = transacoesSheet.getDataRange().getValues();
  const transacoesHeaders = transacoesData[0];
  const transacoesColMap = getColumnMap(transacoesHeaders);

  const spentByCategoryMonth = {};

  for (let i = 1; i < transacoesData.length; i++) {
    const row = transacoesData[i];
    const tipo = row[transacoesColMap["Tipo"]];
    const categoria = row[transacoesColMap["Categoria"]];
    const dataRelevante = parseData(row[transacoesColMap["Data de Vencimento"]]);
    const valor = parseBrazilianFloat(String(row[transacoesColMap["Valor"]]));

    if (tipo === "Despesa" && dataRelevante && categoria) {
      const monthYearKey = Utilities.formatDate(dataRelevante, Session.getScriptTimeZone(), "MMMM/yyyy").toLowerCase();
      const categoryKey = normalizarTexto(categoria);
      const compositeKey = `${monthYearKey}|${categoryKey}`;

      if (!spentByCategoryMonth[compositeKey]) {
        spentByCategoryMonth[compositeKey] = 0;
      }
      spentByCategoryMonth[compositeKey] += valor;
    }
  }

  const newSpentValues = [];

  for (let i = 0; i < budgetData.length; i++) {
    const row = budgetData[i];
    const mesReferencia = (row[1] || "").toString().toLowerCase();
    const categoriaOrcamento = normalizarTexto(row[2]);
    
    const compositeKey = `${mesReferencia}|${categoriaOrcamento}`;
    const spentValue = spentByCategoryMonth[compositeKey] || 0;
    
    newSpentValues.push([spentValue]);
  }

  if (newSpentValues.length > 0) {
    budgetSheet.getRange(2, 5, newSpentValues.length, 1).setValues(newSpentValues);
    logToSheet(userSpreadsheet, "Valores gastos na aba 'Orcamento' atualizados com sucesso via script.", "INFO");
  }
}


/**
 * **REFATORADO:** Busca os dados de progresso do orçamento para o Telegram.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} userSpreadsheet O objeto da planilha do cliente.
 * @param {number} mes O mês para o resumo (1-12).
 * @param {number} ano O ano para o resumo.
 * @returns {Array<Object>} Um array de objetos de progresso do orçamento.
 */
function getBudgetProgressForTelegram(userSpreadsheet, mes, ano) {
    const ss = userSpreadsheet;
    const orcamentoSheet = ss.getSheetByName(SHEET_ORCAMENTO);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);

    if (!orcamentoSheet || !transacoesSheet) return [];

    const dadosOrcamento = orcamentoSheet.getDataRange().getValues();
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    
    // Esta função _getBudgetProgress está definida noutro ficheiro (Code_Dashboard.gs)
    // mas estará disponível no mesmo projeto Hub.
    return _getBudgetProgress(dadosOrcamento, dadosTransacoes, mes - 1, ano, {});
}


/**
 * **REFATORADO:** Busca os dados de progresso das metas de poupança.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} userSpreadsheet O objeto da planilha do cliente.
 * @returns {Array<Object>} Um array de objetos de progresso de metas.
 */
function getGoalsStatusForTelegram(userSpreadsheet) {
    const ss = userSpreadsheet;
    const metasSheet = ss.getSheetByName(SHEET_METAS);

    if (!metasSheet || metasSheet.getLastRow() < 2) {
        return [];
    }

    const dadosMetas = metasSheet.getRange("A2:C" + metasSheet.getLastRow()).getValues();
    
    const status = [];
    dadosMetas.forEach(row => {
        const nome = row[0];
        const objetivo = parseBrazilianFloat(String(row[1] || '0'));
        const salvo = parseBrazilianFloat(String(row[2] || '0'));

        if (nome && objetivo > 0) {
            const percentage = (salvo / objetivo) * 100;
            status.push({
                nome: nome,
                objetivo: objetivo,
                salvo: salvo,
                percentage: percentage
            });
        }
    });

    return status;
}

