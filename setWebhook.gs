function setWebhook() {
  const token = "7355401077:AAHMCnFBu9lPTcIcQXfNvEdTbWhTCB2BFeA";
  const webhookUrl = "https://script.google.com/macros/s/AKfycbw6wCghyEoO7b7gKYe42l1897_BJYaMt0T1_OqjW8nmE7s3U3pdCyNS8sVQ3U0MH-tX/exec";
  const url = `https://api.telegram.org/bot${token}/setWebhook?url=${webhookUrl}`;
  
  const response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

/**
 * Opcional: Função para remover o webhook.
 * Útil para depuração ou se você quiser desativar o bot temporariamente.
 */
function deleteWebhook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfiguracoes = ss.getSheetByName("Configuracoes");

  if (!abaConfiguracoes) {
    Logger.log("ERRO: Aba 'Configuracoes' não encontrada para obter o token do Telegram.");
    return;
  }

  let telegramToken = "";
  const dadosConfig = abaConfiguracoes.getDataRange().getValues();
  for (let i = 0; i < dadosConfig.length; i++) {
    if (dadosConfig[i][0] === "TELEGRAM_TOKEN") {
      telegramToken = dadosConfig[i][1];
      break;
    }
  }

  if (!telegramToken) {
    Logger.log("ERRO: Token do Telegram não encontrado na aba 'Configuracoes'.");
    return;
  }

  const url = `https://api.telegram.org/bot${telegramToken}/deleteWebhook`;

  try {
    const response = UrlFetchApp.fetch(url);
    const responseText = response.getContentText();
    Logger.log(`Resposta do Telegram ao deletar o webhook: ${responseText}`);

    const result = JSON.parse(responseText);
    if (result.ok) {
      Logger.log("Webhook deletado com sucesso!");
      SpreadsheetApp.getUi().alert("Sucesso!", "Webhook deletado com sucesso!", SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log(`Falha ao deletar o webhook: ${result.description}`);
      SpreadsheetApp.getUi().alert("Erro!", `Falha ao deletar o webhook: ${result.description}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    Logger.log(`Erro ao fazer a requisição para o Telegram: ${e.message}`);
    SpreadsheetApp.getUi().alert("Erro!", `Erro ao fazer a requisição para o Telegram: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
