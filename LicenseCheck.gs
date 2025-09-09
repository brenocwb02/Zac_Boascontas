/**
 * @file LicenseCheck.gs
 * @description Lógica de verificação de licença refatorada para o modelo centralizado.
 */

// A URL do servidor de licenças e a função activateProduct podem ser mantidas para
// um fluxo de ativação inicial, mas a verificação principal muda.
const LICENSE_SERVER_URL = "https://script.google.com/macros/s/AKfycbzxbwGNWISM_fhByxxMSUMjYW2fil83p42VRpHP9poFC06VgKGh0WMqtz2kaVGV_xKpbw/exec";

/**
 * **FUNÇÃO ATUALIZADA**
 * Verifica se a licença/subscrição do utilizador é válida consultando a base de dados central.
 * Esta função é chamada a cada interação com o bot ou dashboard.
 * * @param {string|number} chatId O ID do chat do Telegram do utilizador a ser verificado.
 * @returns {boolean} True se o utilizador for válido e ativo, false caso contrário.
 */
function isLicenseValid(chatId) {
  try {
    // Procura o utilizador na base de dados central
    const user = findUserByChatId(chatId); // Função do UserManager.gs

    // A licença é válida se o utilizador existir e o seu status for 'active'
    if (user && user.status === 'active') {
      return true;
    }
    
    logToSheet(`Verificação de licença falhou para o chatId ${chatId}. Utilizador encontrado: ${!!user}, Status: ${user ? user.status : 'N/A'}.`, "WARN");
    return false;

  } catch (e) {
    handleError(e, `isLicenseValid para ${chatId}`);
    return false; // Por segurança, falha a verificação em caso de erro.
  }
}
