/**
 * @file Constants.gs
 * @description REFATORADO: Este arquivo centraliza todas as constantes e configurações globais do projeto Hub.
 * A maioria das constantes permanece a mesma, pois são essenciais para o funcionamento do sistema.
 */

// Chaves para o PropertiesService do PROJETO HUB.
const TELEGRAM_TOKEN_PROPERTY_KEY = 'TELEGRAM_TOKEN';
const ADMIN_CHAT_ID_PROPERTY_KEY = 'ADMIN_CHAT_ID';
const WEB_APP_URL_PROPERTY_KEY = 'WEB_APP_URL';
const SPEECH_API_KEY_PROPERTY_KEY = 'SPEECH_API_KEY';

/**
 * Obtém o Chat ID do administrador das Propriedades do Script do Hub.
 * Não precisa de `userSpreadsheet` pois é uma propriedade do seu projeto central.
 * @returns {string} O Chat ID do administrador.
 */
function getAdminChatIdFromProperties() {
    return PropertiesService.getScriptProperties().getProperty(ADMIN_CHAT_ID_PROPERTY_KEY);
}

const URL_BASE_TELEGRAM = "https://api.telegram.org/bot";

// --- Nomes das Abas da Planilha (usados para aceder às planilhas dos clientes) ---
const SHEET_TRANSACOES = "Transacoes";
const SHEET_CONTAS = "Contas";
const SHEET_PALAVRAS_CHAVE = "PalavrasChave";
const SHEET_CONFIGURACOES = "Configuracoes";
const SHEET_LOGS_SISTEMA = "Logs_Sistema";
const SHEET_CATEGORIAS = "Categorias";
const SHEET_METAS = "Metas";
const SHEET_ORCAMENTO = "Orcamento";
const SHEET_ALERTAS_ENVIADOS = "AlertasEnviados";
const SHEET_CONTAS_A_PAGAR = "Contas_a_Pagar";
const SHEET_NOTIFICACOES_CONFIG = 'Notificacoes_Config';
const SHEET_LEARNED_CATEGORIES = "LearnedCategories";
const SHEET_INVESTIMENTOS = "Investimentos";
const SHEET_ATIVOS_MANUAIS = "Ativos";
const SHEET_PASSIVOS_MANUAIS = "Passivos";

// --- Constantes de Cache (usadas pelo Hub para otimizar o desempenho) ---
const CACHE_KEY_PALAVRAS = 'palavras_chave_cache';
const CACHE_KEY_CONTAS = 'contas_cache';
const CACHE_KEY_CONFIG = 'config_cache';
const CACHE_KEY_TUTORIAL_STATE = 'tutorial_state';
const CACHE_KEY_PENDING_TRANSACTIONS = 'pending_transaction';
const CACHE_KEY_EDIT_STATE = 'edit_state';
const CACHE_KEY_ASSISTANT_STATE = 'assistant_state';
const CACHE_KEY_CATEGORIAS = 'categorias_cache';
const CACHE_KEY_TRANSACOES = 'transacoes_cache';
const CACHE_KEY_CONTAS_A_PAGAR = 'contas_a_pagar_cache';
const CACHE_KEY_DASHBOARD_TOKEN = 'dashboard_access_token';

const CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS = 300; // 5 minutos
const CACHE_EXPIRATION_SECONDS = 21600; // 6 horas
const CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS = 900; // 15 minutos
const CACHE_EXPIRATION_TUTORIAL_STATE_SECONDS = 1800; // 30 minutos
const CACHE_EXPIRATION_EDIT_STATE_SECONDS = 900; // 15 minutos

// --- Constantes de Lógica Financeira ---
const SIMILARITY_THRESHOLD = 0.75;
const BUDGET_ALERT_THRESHOLD_PERCENT = 80;
const BILL_REMINDER_DAYS_BEFORE = 3;
const MIN_CONFIDENCE_TO_APPLY = 2; // Constante do Learning.gs movida para cá

// --- Constantes de Lógica de Tutorial ---
const TUTORIAL_STATE_WAITING_DESPESA = "waiting_for_despesa";
const TUTORIAL_STATE_WAITING_RECEITA = "waiting_for_receita";
const TUTORIAL_STATE_WAITING_SALDO = "waiting_for_saldo";
const TUTORIAL_STATE_WAITING_CONTAS_A_PAGAR = "waiting_for_contas_a_pagar";

// --- Níveis de Log ---
const LOG_LEVEL_MAP = {
  "DEBUG": 1,
  "INFO": 2,
  "WARN": 3,
  "ERROR": 4,
  "NONE": 5
};
let currentLogLevel = "INFO";

// --- Regex para Pagamento de Fatura ---
const regexPagamentoFatura = /paguei\s+(?:r\$)?\s*([\d.,]+)\s+da\s+fatura\s+(?:do|da|de)?\s*(.+?)\s+com\s+(.+)/i;
