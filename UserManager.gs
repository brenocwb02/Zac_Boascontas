/**
 * @file UserManager.gs
 * @description Funções para gerir a sua base de dados de utilizadores (`Master_Users_DB`).
 * Crie este ficheiro no seu novo projeto Apps Script "Hub".
 */

// !!! AÇÃO NECESSÁRIA: Substitua os IDs abaixo pelos IDs REAIS das suas planilhas !!!
const MASTER_USERS_DB_ID = "1G2_Y-pcRxl3271CYM8Y2FXSLnujr2yISEUasFRohgqA"; // ID da planilha Master_Users_DB
const USER_DB_SHEET_NAME = "Utilizadores";
const MODEL_SPREADSHEET_ID = "1b1YL0U1awnNonTptOxPdTATmcneJ0PRRmM8OB5VLATI"; // ID da sua planilha modelo (a que não terá código)

/**
 * Procura um utilizador na base de dados pelo Chat ID.
 * @param {string|number} chatId O ID do chat do Telegram.
 * @returns {Object|null} Dados do utilizador ou null se não for encontrado.
 */
function findUserByChatId(chatId) {
  try {
    const dbSheet = SpreadsheetApp.openById(MASTER_USERS_DB_ID).getSheetByName(USER_DB_SHEET_NAME);
    const data = dbSheet.getDataRange().getValues();
    // Colunas: A=UserID, B=UserName, C=UserEmail, D=SpreadsheetID, E=SubscriptionStatus
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === chatId.toString()) {
        return {
          chatId: data[i][0],
          name: data[i][1],
          email: data[i][2],
          spreadsheetId: data[i][3],
          status: data[i][4]
        };
      }
    }
    return null;
  } catch(e) {
    console.error(`Erro ao procurar utilizador ${chatId}: ${e.message}`);
    return null;
  }
}

/**
 * Lida com o primeiro contacto de um novo utilizador.
 * @param {string|number} chatId O ID do chat do novo utilizador.
 * @param {string} name O nome do novo utilizador.
 */
function handleNewUserOnboarding(chatId, name) {
  const message = `👋 Olá, ${name}! Bem-vindo ao Boas Contas.\n\n` +
                  `Para ativar a sua conta, por favor, use o comando abaixo, substituindo com a sua chave e o seu email:\n\n` +
                  `\`/ativar SUA-CHAVE-AQUI seu.email@exemplo.com\``;
  // Substituir esta linha pela sua função `enviarMensagemTelegram` quando a mover para o Hub.
  console.log(`Mensagem de Onboarding para ${chatId}: ${message}`);
}

/**
 * Cria uma nova conta de utilizador, copia a planilha modelo e regista na base de dados.
 * @param {string|number} chatId O ID do chat do Telegram.
 * @param {string} userName O nome do utilizador.
 * @param {string} userEmail O email do utilizador.
 * @param {string} licenseKey A chave de licença usada.
 * @returns {string|null} O ID da nova planilha criada, ou null em caso de erro.
 */
function createNewUser(chatId, userName, userEmail, licenseKey) {
  try {
    const modelFile = DriveApp.getFileById(MODEL_SPREADSHEET_ID);
    const newFileName = `Boas Contas - ${userName}`;
    const newSpreadsheetFile = modelFile.makeCopy(newFileName);
    const newSpreadsheetId = newSpreadsheetFile.getId();

    newSpreadsheetFile.addEditor(userEmail);

    const dbSheet = SpreadsheetApp.openById(MASTER_USERS_DB_ID).getSheetByName(USER_DB_SHEET_NAME);
    dbSheet.appendRow([
      chatId,
      userName,
      userEmail,
      newSpreadsheetId,
      'active',
      licenseKey
    ]);
    
    console.log(`Novo utilizador criado: ${userName} (${userEmail}) com a planilha ID: ${newSpreadsheetId}`);
    
    return newSpreadsheetId;
  } catch(e) {
    console.error(`Erro ao criar novo utilizador ${userName}: ${e.message}`);
    return null;
  }
}

