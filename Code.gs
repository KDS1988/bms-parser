/**
 * BMS Event Watcher
 * ------------------
 * Раз в N минут опрашивает bms-api.vsporte.ru на предмет новых мероприятий,
 * которые внесли менеджеры, шлёт уведомление в группу "отдел назначений"
 * и добавляет строку на отдельный лист "Новые мероприятия".
 * (Перенос на "доску" — отдельная задача на потом.)
 *
 * УСТАНОВКА:
 * 1. Extensions -> Apps Script в нужной таблице, вставить этот файл.
 * 2. Сохранить, перезагрузить страницу таблицы — появится меню "BMS".
 * 3. Меню BMS -> "🔑 Войти в BMS" — введи телефон и код из Telegram (один раз,
 *    дальше сессия обновляется сама; повторный вход нужен только если
 *    refresh_token протухнет, скрипт сам подскажет).
 * 4. Поставить триггер: doPoll, time-driven, every 5-10 minutes.
 */

const CONFIG = {
  BMS_API: 'https://bms-api.vsporte.ru/api/v1/bms',

  NEW_EVENTS_SHEET_NAME: 'Новые мероприятия', // сюда пишем строки по новым событиям
  TECH_BLOCK_SHEET_NAME: 'Тех блок',           // амплуа из тех.блока услуг "Прямая трансляция" / "Студия"
  LEDGER_SHEET_NAME: '_bms_seen',              // реестр обработанных event_id (создастся сам)
  BOARD_STATE_SHEET_NAME: '_bms_board_state',  // положение блоков на доске + слепок назначений (создастся сам)

  // услуги, чей тех.блок нас интересует (сравнение по service.name)
  TARGET_SERVICES: ['Прямая трансляция', 'Студия'],

  // горизонт, на который смотрим вперёд при опросе (дней)
  LOOKAHEAD_DAYS: 60,
};

// ============================== СЕКРЕТЫ (Script Properties) ==============================
// Токен бота, номер телефона и chat_id — НЕ в коде (чтобы не улетели в git), а в
// Project Settings -> Script Properties этого Apps Script проекта. Один раз
// выполнить setupSecrets() ниже, подставив свои значения, либо вписать руками
// через интерфейс Apps Script: значок шестерёнки слева -> Script Properties.

function setupSecrets() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    TELEGRAM_BOT_TOKEN: 'ВСТАВЬ_ТОКЕН_ОТ_BOTFATHER',
    TELEGRAM_CHAT_ID: 'ВСТАВЬ_CHAT_ID',   // отрицательное число для группы
    BMS_PHONE: 'ВСТАВЬ_НОМЕР_ТЕЛЕФОНА',   // без +, как 79268777745
    WEBHOOK_SECRET: 'ПРИДУМАЙ_ДЛИННУЮ_СЛУЧАЙНУЮ_СТРОКУ', // для проверки запросов от GitHub Actions
  });
  Logger.log('Секреты сохранены в Script Properties.');
}

function getSecret_(key) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  if (!val) throw new Error(`Не задан ${key} — выполни setupSecrets() или впиши вручную в Project Settings -> Script Properties.`);
  return val;
}

// ============================== МЕНЮ ==============================

/** Добавляет пункт меню "BMS" при открытии таблицы. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BMS')
    .addItem('🔑 Войти в BMS', 'menuLogin')
    .addItem('▶ Запустить проверку новых мероприятий', 'doPoll')
    .addItem('🧪 Тест: заполнить "Тех блок" по текущим данным', 'test_FillTechBlockForAllCurrentEvents')
    .addItem('🧪 Тест: записать на доску (event 14175)', 'test_WriteBoardForEvent')
    .addItem('▶ Записать на доску дату...', 'test_WriteBoardForDate')
    .addItem('📋 Обновить график персонала сейчас', 'refreshStaffScheduleNow')
    .addItem('🔬 Дамп структуры выделенных ячеек (для отладки доски)', 'debugDumpSelection')
    .addItem('🧹 Сбросить реестр положений на доске', 'resetBoardState')
    .addItem('🔍 Найти замену по ID мероприятия...', 'findReplacementsManually')
    .addItem('📝 Перенести черновики в BMS...', 'scanDraftAssignments')
    .addItem('✏️ Добавить мероприятие на доску вручную', 'addManualEntry')
    .addToUi();
}

// ============================== АВТОРИЗАЦИЯ ==============================

/**
 * Интерактивный логин через диалоговые окна таблицы.
 * Запускать из меню BMS -> "Войти в BMS", когда нужен первый вход
 * или когда refresh_token протух.
 */
function menuLogin() {
  const ui = SpreadsheetApp.getUi();

  const phoneResp = ui.prompt(
    'Вход в BMS',
    'Номер телефона (как в BMS, без +):',
    ui.ButtonSet.OK_CANCEL
  );
  if (phoneResp.getSelectedButton() !== ui.Button.OK) return;
  const phone = phoneResp.getResponseText().trim() || getSecret_('BMS_PHONE');

  const otpResp = requestOtp_(phone);
  if (!otpResp.ok) {
    ui.alert('Не удалось запросить код: ' + otpResp.message);
    return;
  }

  const codeResp = ui.prompt(
    'Вход в BMS',
    'Код из Telegram-бота @vsporte_bms_bot:',
    ui.ButtonSet.OK_CANCEL
  );
  if (codeResp.getSelectedButton() !== ui.Button.OK) return;
  const code = codeResp.getResponseText().trim();

  const loginResp = loginWithOtp_(phone, code);
  if (loginResp.ok) {
    ui.alert('Готово, вход выполнен.');
  } else {
    ui.alert('Логин не удался: ' + loginResp.message);
  }
}

/** Запрашивает отправку OTP-кода в Telegram-бот, привязанный к аккаунту. */
function requestOtp_(phone) {
  const resp = UrlFetchApp.fetch(
    `${CONFIG.BMS_API}/notifications/user/send_otp?phone_number=${phone}`,
    {
      method: 'post',
      headers: { origin: 'https://bms.vsporte.ru', referer: 'https://bms.vsporte.ru/' },
      muteHttpExceptions: true,
    }
  );
  const ok = resp.getResponseCode() === 200;
  return { ok, message: resp.getContentText() };
}

/** Логинится по телефону+коду, сохраняет refresh_token. */
function loginWithOtp_(phone, code) {
  const resp = UrlFetchApp.fetch(`${CONFIG.BMS_API}/users/token`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ phone_number: phone, otp_code: code }),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    return { ok: false, message: resp.getContentText() };
  }
  const data = JSON.parse(resp.getContentText());
  PropertiesService.getScriptProperties().setProperty('BMS_REFRESH_TOKEN', data.refresh_token);
  try {
    sendTelegramMessage_('✅ Выполнен вход в BMS.');
  } catch (e) {
    // не критично, если уведомление не ушло — сам вход уже сохранён
  }
  return { ok: true };
}

/**
 * Возвращает свежий access_token, обновляя refresh_token по мере ротации.
 * Если refresh_token отсутствует или протух — бросает понятную ошибку
 * с указанием зайти через меню (сам диалог отсюда не откроется:
 * триггеры по расписанию не могут показывать UI).
 */
function getAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const refreshToken = props.getProperty('BMS_REFRESH_TOKEN');
  if (!refreshToken) {
    throw new Error('Нет сохранённого входа в BMS. В таблице: меню BMS -> "Войти в BMS".');
  }

  const resp = UrlFetchApp.fetch(`${CONFIG.BMS_API}/users/token/refresh`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ refresh_token: refreshToken }),
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Не удалось обновить токен: ' + resp.getContentText() +
      '\nПохоже, вход протух — в таблице: меню BMS -> "Войти в BMS".');
  }

  const data = JSON.parse(resp.getContentText());
  props.setProperty('BMS_REFRESH_TOKEN', data.refresh_token); // токен ротируется каждый раз
  return data.access_token;
}

// ============================== ПОЛУЧЕНИЕ ДАННЫХ ==============================

function bmsGet_(path, params) {
  const accessToken = getAccessToken_();
  const qs = Object.keys(params || {})
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');
  const url = `${CONFIG.BMS_API}/${path}${qs ? '?' + qs : ''}`;

  const MAX_ATTEMPTS = 3;
  let lastError;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          origin: 'https://bms.vsporte.ru',
          referer: 'https://bms.vsporte.ru/',
        },
        muteHttpExceptions: true,
      });

      if (resp.getResponseCode() !== 200) {
        throw new Error(`BMS API ${path} -> ${resp.getResponseCode()}: ${resp.getContentText()}`);
      }
      return JSON.parse(resp.getContentText());
    } catch (e) {
      lastError = e;
      // транзиентные сетевые сбои (DNS error и т.п.) на стороне UrlFetchApp — просто повторяем
      if (attempt < MAX_ATTEMPTS) {
        Utilities.sleep(1000 * attempt);
      }
    }
  }
  throw lastError;
}

/**
 * POST в BMS API. Без ретраев на сетевые сбои (в отличие от bmsGet_) —
 * это запись реальных данных, при неясном исходе (таймаут и т.п.) лучше явно
 * упасть с ошибкой, чем рискнуть создать назначение дважды повторным запросом.
 */
function bmsPost_(path, body) {
  const accessToken = getAccessToken_();
  const url = `${CONFIG.BMS_API}/${path}`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      origin: 'https://bms.vsporte.ru',
      referer: 'https://bms.vsporte.ru/',
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`BMS API POST ${path} -> ${code}: ${resp.getContentText()}`);
  }
  return JSON.parse(resp.getContentText());
}

/**
 * Назначает сотрудника на конкретную строку услуги в BMS.
 * eventServiceId — item.id (event_service_id), eventServiceLineId — line.id
 * из event_service_lines (НЕ line.line.id — это id каталожного амплуа, а
 * нужен id конкретной строки-назначения внутри этой услуги).
 */
function assignEmployeeInBms_(eventServiceId, eventServiceLineId, employeeId) {
  return bmsPost_('projects/assignment/technical', {
    comment: '',
    dismantling: false,
    employee_id: employeeId,
    event_service_additional_service_list: [],
    event_service_id: eventServiceId,
    event_service_line_id: eventServiceLineId,
    installation: false,
    logistics: null,
    event_service_video_task_list: [],
  });
}

/** Собирает все мероприятия на ближайшие LOOKAHEAD_DAYS дней (все страницы). */
function fetchUpcomingAssignments_() {
  const today = new Date();
  const dateFrom = Utilities.formatDate(today, 'Europe/Moscow', 'yyyy-MM-dd');
  const until = new Date(today.getTime() + CONFIG.LOOKAHEAD_DAYS * 86400000);
  const dateTo = Utilities.formatDate(until, 'Europe/Moscow', 'yyyy-MM-dd');

  let items = [];
  let page = 1;
  while (true) {
    const data = bmsGet_('projects/assignment_table/technical', {
      page, size: 100,
      order_by: '+date,start_time,id',
      date__gte: dateFrom, date__lte: dateTo,
    });
    items = items.concat(data.items);
    if (page >= (data.pages || 1)) break;
    page++;
    Utilities.sleep(300); // небольшая пауза между страницами, снижает шанс DNS/сетевых сбоев
  }
  return items;
}

// ============================== РЕЕСТР ОБРАБОТАННЫХ ==============================

function getLedgerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.LEDGER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LEDGER_SHEET_NAME);
    sheet.hideSheet();
    sheet.appendRow(['event_id', 'first_seen_at']);
  }
  return sheet;
}

function getSeenEventIds_() {
  const sheet = getLedgerSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return new Set(values.map(r => String(r[0])));
}

function markEventSeen_(eventId) {
  getLedgerSheet_().appendRow([eventId, new Date()]);
}

// ============================== TELEGRAM ==============================

function sendTelegramMessage_(text) {
  UrlFetchApp.fetch(`https://api.telegram.org/bot${getSecret_('TELEGRAM_BOT_TOKEN')}/sendMessage`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: getSecret_('TELEGRAM_CHAT_ID'),
      text,
      parse_mode: 'HTML',
    }),
    muteHttpExceptions: true,
  });
}

/**
 * Короткое отображаемое название мероприятия для доски/уведомлений:
 * - сингл (event_type: "single") — команда хозяев (или "хозяева — гости", где применимо)
 * - гибрид (event_type: "hybrid") — короткое additional_info ("Информация -> Название"
 *   в самой BMS), т.к. у гибридных нет команд, а event_description обычно длинный и
 *   многословный, для доски не годится
 * Фолбэк на event_description, если additional_info пуст.
 */
function eventDisplayName_(event) {
  if (event.home_team && event.away_team) {
    return `${event.home_team.name} — ${event.away_team.name}`;
  }
  if (event.home_team) return event.home_team.name;
  return event.additional_info || event.event_description || '';
}

function formatEventMessage_(event) {
  const sport = event.sport_type ? event.sport_type.name : '';
  const league = event.league ? event.league.short_name : '';
  const teams = eventDisplayName_(event);
  const location = event.location ? event.location.name : (event.venue || '');

  return [
    `🆕 <b>Новое мероприятие в BMS</b>`,
    `${event.date} ${event.start_time || ''}`,
    [sport, league].filter(Boolean).join(' / '),
    teams,
    location ? `📍 ${location}` : '',
  ].filter(Boolean).join('\n');
}

// ============================== ЛИСТ "НОВЫЕ МЕРОПРИЯТИЯ" ==============================

function getNewEventsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.NEW_EVENTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NEW_EVENTS_SHEET_NAME);
    sheet.appendRow([
      'event_id', 'Дата', 'Время', 'Вид спорта', 'Лига',
      'Команды / описание', 'Локация', 'Город',
      'Проект', 'Заказчик', 'Менеджер', 'Добавлено в BMS', 'Обнаружено ботом',
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * item — представитель строки assignment_table для этого event_id
 * (первая попавшаяся услуга; project/customer/manager берутся с неё,
 * т.к. в самом event их нет).
 */
function appendNewEventRow_(event, item) {
  const sheet = getNewEventsSheet_();
  const sport = event.sport_type ? event.sport_type.name : '';
  const league = event.league ? event.league.short_name : '';
  const teams = eventDisplayName_(event);
  const location = event.location ? event.location.name : (event.venue || '');
  const city = (event.location && event.location.city)
    || (event.home_team && event.home_team.city)
    || '';

  sheet.appendRow([
    event.id,
    event.date,
    event.start_time || '',
    sport,
    league,
    teams,
    location,
    city,
    item.project ? item.project.name : '',
    item.customer ? item.customer.name : '',
    item.manager ? `${item.manager.first_name} ${item.manager.last_name}` : '',
    event.date_first_filled_at || '',
    new Date(),
  ]);
}

// ============================== ЛИСТ "ТЕХ БЛОК" (амплуа) ==============================

function getTechBlockSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TECH_BLOCK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TECH_BLOCK_SHEET_NAME);
    sheet.appendRow([
      'event_id', 'event_service_id', 'Дата', 'Команды / описание',
      'Услуга', '№', 'Амплуа', 'Статус', 'Исполнитель', 'Обнаружено ботом',
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * item — строка из assignment_table/technical с полями service и event_service_lines.
 * Пишет по одной строке на каждое амплуа (line) из тех.блока этой услуги.
 */
function appendTechBlockRows_(event, item) {
  if (!item.event_service_lines || item.event_service_lines.length === 0) return;

  const sheet = getTechBlockSheet_();
  const teams = eventDisplayName_(event);
  const serviceName = item.service ? item.service.name : '';
  const now = new Date();

  const rows = item.event_service_lines.map(line => {
    const at = line.assignment_technical;
    // Предположение по подрядчикам (assignment_contractor) не проверено на реальном
    // примере с заполненным полем — если формат окажется другим, поправим.
    const executor = (at && at.employee)
      ? `${at.employee.last_name} ${at.employee.first_name}`
      : (line.assignment_contractor && line.assignment_contractor.contractor
          ? line.assignment_contractor.contractor.name
          : '');

    return [
      event.id,
      item.event_id,
      event.date,
      teams,
      serviceName,
      line.line_serial_number,
      line.line ? line.line.name : '',
      line.status,
      executor,
      now,
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/** Группирует плоский список item-ов assignment_table по event_id. */
function groupByEvent_(items) {
  const byEventId = {};
  for (const item of items) {
    const ev = item.event;
    if (!byEventId[ev.id]) byEventId[ev.id] = { event: ev, items: [] };
    byEventId[ev.id].items.push(item);
  }
  return Object.values(byEventId).sort((a, b) => a.event.id - b.event.id);
}

/**
 * ТЕСТ: прогоняет извлечение амплуа по ВСЕМ текущим мероприятиям в горизонте
 * LOOKAHEAD_DAYS, независимо от реестра "просмотрено". НЕ шлёт в Telegram,
 * НЕ трогает _bms_seen и "Новые мероприятия" — только заполняет "Тех блок",
 * чтобы можно было визуально сверить результат с реальными карточками в BMS.
 * Запускать вручную из редактора.
 */
function test_FillTechBlockForAllCurrentEvents() {
  const entries = groupByEvent_(fetchUpcomingAssignments_());
  let written = 0;

  for (const { event, items } of entries) {
    const targetItems = items.filter(it =>
      it.service && CONFIG.TARGET_SERVICES.includes(it.service.name)
    );
    for (const item of targetItems) {
      try {
        appendTechBlockRows_(event, item);
        written++;
      } catch (e) {
        Logger.log(`event_id=${event.id}, service=${item.service.name}: ${e}`);
      }
    }
  }
  Logger.log(`Готово. Обработано услуг: ${written}`);
}

// ============================== ДОСКА (по месяцам) ==============================

const MONTHS_RU = ['ЯНВАРЬ', 'ФЕВРАЛЬ', 'МАРТ', 'АПРЕЛЬ', 'МАЙ', 'ИЮНЬ',
  'ИЮЛЬ', 'АВГУСТ', 'СЕНТЯБРЬ', 'ОКТЯБРЬ', 'НОЯБРЬ', 'ДЕКАБРЬ'];
const WEEKDAYS_RU = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];

const AMPLUA_ABBR = {
  'Режиссер многокамерного эфира': 'реж',
  'Режиссер однокамерного эфира': 'реж',
  'Режиссер повторов': 'повт',
  'Режиссер графики': 'титры',
  'Выпускающий режиссер': 'ВР',
  'Звукорежиссер': 'звук',
  'Инженер': 'инж', // ловит и просто "Инженер", и "Инженер видеоповторов"
  'Видеооператор': 'опер',
  'Продюсер': 'прод',
};

/**
 * Порядок вывода ролей на доске — принудительный, не зависит от того, в каком
 * порядке их вернула BMS (там это просто порядок ввода менеджером, часто
 * хаотичный). Первое совпадение по подстроке побеждает — специфичные виды
 * "Режиссера" идут раньше общего "Режиссер", иначе "Режиссер повторов"
 * ошибочно попал бы в общую категорию.
 */
/**
 * Правила сортировки ролей: проверка совпадения идёт по паттерну (специфичные
 * виды режиссёра — раньше общего "Режиссер", иначе "Режиссер повторов" сам
 * попал бы под общий), а порядок вывода на доске — по rank, отдельно от
 * порядка проверки. Прод -> Реж -> Повт -> Титры -> Звук -> Инж -> (операторы,
 * они всегда последними и обрабатываются отдельно).
 */
const ROLE_RANK_RULES = [
  { pattern: 'Продюсер', rank: 0 },
  { pattern: 'Режиссер повторов', rank: 2 },
  { pattern: 'Режиссер графики', rank: 3 },
  { pattern: 'Звукорежиссер', rank: 4 },
  { pattern: 'Инженер', rank: 5 },
  { pattern: 'Режиссер', rank: 1 }, // общий (много-/однокамерного эфира)
];

function roleSortRank_(lineName) {
  const rule = ROLE_RANK_RULES.find(r => lineName && lineName.includes(r.pattern));
  return rule ? rule.rank : 99;
}

const BOARD_PINK = '#f4cccc';
const BOARD_BLUE = '#cfe2f3';
const BOARD_GRAY = '#999999';
const BOARD_GROUP_WIDTH = 3;    // ширина одного блока-матча в колонках
const DEFAULT_DAY_ROWS = 15;    // стартовое резервирование на дату (шапка + 14 строк), растёт при нехватке
const BOARD_ROW_BUFFER = 2;     // запас строк сверх необходимого при досоздании места
const BOARD_HEADER_WIDTH = 40;  // на сколько колонок красим серую шапку дня

function daysInMonth_(year, month0) {
  return new Date(year, month0 + 1, 0).getDate();
}

function dayRowMapKey_(sheetName) {
  return `DAY_ROWS::${sheetName}`;
}

/** Карта "день месяца -> строка серой шапки" для конкретного листа. Растёт при вставке строк. */
function loadDayRowMap_(sheetName) {
  const raw = PropertiesService.getScriptProperties().getProperty(dayRowMapKey_(sheetName));
  return raw ? JSON.parse(raw) : null;
}

function saveDayRowMap_(sheetName, map) {
  PropertiesService.getScriptProperties().setProperty(dayRowMapKey_(sheetName), JSON.stringify(map));
}

/**
 * Возвращает лист доски для месяца события ("СЕНТЯБРЬ 2026"). При первом
 * создании сразу генерирует шапки ВСЕХ дат месяца по порядку (стартово по
 * DEFAULT_DAY_ROWS строк на день) — так дни идут строго по календарю. Если
 * позже блоку не хватит места, строки досоздаются точечно (см. ensureRoomForBlock_).
 */
function getOrCreateMonthSheet_(dateStr) {
  const d = new Date(dateStr);
  const year = d.getFullYear();
  const month0 = d.getMonth();
  const name = `${MONTHS_RU[month0]} ${year}`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const existing = ss.getSheetByName(name);
  if (existing) return existing;

  const sheet = ss.insertSheet(name);
  const total = daysInMonth_(year, month0);
  const map = {};
  for (let day = 1; day <= total; day++) {
    const headerRow = 1 + (day - 1) * DEFAULT_DAY_ROWS;
    map[day] = headerRow;
    const dayDate = new Date(year, month0, day);
    const shortDate = Utilities.formatDate(dayDate, 'Europe/Moscow', 'dd.MM');
    const weekday = WEEKDAYS_RU[dayDate.getDay()];

    sheet.getRange(headerRow, 1, 1, BOARD_HEADER_WIDTH).setBackground(BOARD_GRAY);
    sheet.getRange(headerRow, 1).setNumberFormat('@STRING@')
      .setValue(shortDate).setFontWeight('bold').setFontColor('#ffffff');
    sheet.getRange(headerRow, 2).setValue(weekday).setFontColor('#ffffff');
  }
  saveDayRowMap_(name, map);
  return sheet;
}

/** Строка серой шапки для даты — из динамической карты (не арифметика, т.к. дни могут вырасти). */
function findOrCreateDateHeaderRow_(sheet, dateStr) {
  const d = new Date(dateStr);
  const day = d.getDate();
  let map = loadDayRowMap_(sheet.getName());
  if (!map || !map[day]) {
    // Старый лист без карты (создан до этого обновления) — восстанавливаем арифметикой.
    // История возможных прошлых расширений для такого листа теряется — если увидишь
    // наезжающие блоки на старом листе, удали его и пересоздай через любую запись.
    const total = daysInMonth_(d.getFullYear(), d.getMonth());
    map = {};
    for (let dd = 1; dd <= total; dd++) map[dd] = 1 + (dd - 1) * DEFAULT_DAY_ROWS;
    saveDayRowMap_(sheet.getName(), map);
  }
  return map[day];
}

/**
 * Гарантирует, что под блок высотой neededRows (считая от headerRow+1) хватает
 * места до шапки следующего дня. Если нет — вставляет недостающие строки
 * (+BOARD_ROW_BUFFER с запасом) прямо перед шапкой следующего дня, сдвигая
 * вниз карту дней и уже записанные положения блоков в _bms_board_state для
 * всех более поздних дат этого листа (иначе следующее обновление затрёт не те строки).
 */
function ensureRoomForBlock_(sheet, day, headerRow, neededRows, boardStateMap) {
  const sheetName = sheet.getName();
  const map = loadDayRowMap_(sheetName) || {};
  const nextHeaderRow = map[day + 1]; // undefined для последнего дня месяца — расти некуда упираться

  if (nextHeaderRow == null) return;

  const available = nextHeaderRow - headerRow - 1;
  if (neededRows <= available) return;

  const insertCount = (neededRows - available) + BOARD_ROW_BUFFER;
  sheet.insertRowsBefore(nextHeaderRow, insertCount);

  Object.keys(map).forEach(k => {
    if (Number(k) > day) map[k] = map[k] + insertCount;
  });
  saveDayRowMap_(sheetName, map);

  bumpBoardStateHeaderRows_(sheetName, headerRow, insertCount, boardStateMap);
}

/** Сдвигает header_row и start_row в реестре доски (и в загруженной в память карте) для блоков ниже afterHeaderRow. */
function bumpBoardStateHeaderRows_(sheetName, afterHeaderRow, delta, boardStateMap) {
  const sheet = getBoardStateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, 9);
    const values = range.getValues();
    let changed = false;
    values.forEach(r => {
      if (r[2] === sheetName && r[3] > afterHeaderRow) {
        r[3] = r[3] + delta;
        if (r[5] > afterHeaderRow) r[5] = r[5] + delta; // start_row
        changed = true;
      }
    });
    if (changed) range.setValues(values);
  }

  Object.keys(boardStateMap).forEach(key => {
    const st = boardStateMap[key];
    if (st.sheetName === sheetName && st.headerRow > afterHeaderRow) {
      st.headerRow += delta;
      if (st.startRow > afterHeaderRow) st.startRow += delta;
    }
  });
}

/**
 * Ищет первую свободную группу колонок (ширина BOARD_GROUP_WIDTH) в строке,
 * начиная с B. Колонка считается занятой, если непуста ячейка с командой/
 * описанием (col+1) ИЛИ если её для этого дня уже застолбил столбик
 * подрядчиков (CSTACK). Проверяем именно col+1, а не col (лига) — у части
 * мероприятий (например, соревнования без структурированной лиги в BMS)
 * ячейка с лигой законно пустая, и раньше это заставляло код считать занятую
 * колонку свободной — из-за чего новые блоки садились поверх старых.
 */
function findFreeColumnGroup_(sheet, row, day) {
  const reservedByStack = day != null
    ? getStackReservedCol_(sheet.getName(), day)
    : null;

  let col = 2; // B
  while (sheet.getRange(row, col + 1).getValue() !== '' || col === reservedByStack) {
    col += BOARD_GROUP_WIDTH;
  }
  return col;
}

/** Колонка, которую для этого дня держит столбик подрядчиков (если есть). */
function getStackReservedCol_(sheetName, day) {
  const raw = PropertiesService.getScriptProperties().getProperty(`CSTACK::${sheetName}::day${day}`);
  return raw ? JSON.parse(raw).col : null;
}

/**
 * Резервирует место для очередного КОМПАКТНОГО (подрядчик) блока: если в этот
 * день уже есть "столбик" подрядчиков с местом до конца дня — ставим следующим
 * под предыдущим (2 строки на запись). Если места нет или столбика ещё нет —
 * начинаем новую группу колонок. Позиция столбика хранится в Script Properties,
 * ключ — по номеру ДНЯ (не строки!): строка-шапка дня может сдвинуться вниз,
 * если раньше в этом же месяце другому мероприятию понадобилось больше места —
 * а номер дня месяца от этого не меняется.
 */
function reserveContractorSlot_(sheet, day, headerRow) {
  const sheetName = sheet.getName();
  const stackKey = `CSTACK::${sheetName}::day${day}`;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(stackKey);
  const stack = raw ? JSON.parse(raw) : null;

  const map = loadDayRowMap_(sheetName) || {};
  const nextHeaderRow = map[day + 1];
  const dayBottom = nextHeaderRow ? nextHeaderRow - 1 : headerRow + DEFAULT_DAY_ROWS - 1;

  let col, startRow;
  if (stack && (stack.nextRow + 1) <= dayBottom) {
    col = stack.col;
    startRow = stack.nextRow;
  } else {
    col = findFreeColumnGroup_(sheet, headerRow + 1, day);
    startRow = headerRow + 1;
  }

  props.setProperty(stackKey, JSON.stringify({ col, nextRow: startRow + 2 }));
  return { col, startRow };
}

/**
 * Пишет заголовочную ячейку блока (команда + ID мероприятия — по нему ищут
 * в фильтре BMS, ID услуги для этого не годится). Если мероприятие отменено
 * в BMS (service_status начинается с "Отмен") — добавляет " ОТМЕНА" красным
 * жирным, остальной текст — обычным жирным, как всегда.
 */
/**
 * Пишет заголовочную ячейку блока (команда + ID мероприятия — по нему ищут
 * в фильтре BMS, ID услуги для этого не годится). Если мероприятие отменено
 * в BMS (service_status начинается с "Отмен") — добавляет " ОТМЕНА" красным
 * жирным, остальной текст — обычным жирным, как всегда.
 *
 * itemId (event_service_id) пишется в ЗАМЕТКУ ячейки (невидимую, только по
 * наведению) — по ней код узнаёт блок, даже если его перенесли вручную в
 * пределах дня (Sheets переносит заметку вместе с содержимым при вырезании/
 * вставке или перетаскивании диапазона).
 */
function setHeaderTeamCell_(cell, homeTeam, eventId, isCancelled, itemId) {
  const base = `${homeTeam} (ID ${eventId})`;
  cell.setNote(String(itemId));
  if (!isCancelled) {
    cell.setValue(base).setFontWeight('bold').setFontColor(null);
    return;
  }
  const suffix = ' ОТМЕНА';
  const full = base + suffix;
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(full)
    .setTextStyle(0, base.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#000000').build())
    .setTextStyle(base.length, full.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#cc0000').build())
    .build();
  cell.setRichTextValue(rich);
}

function amplua_(lineName) {
  const key = Object.keys(AMPLUA_ABBR).find(k => lineName && lineName.includes(k));
  return key ? AMPLUA_ABBR[key] : lineName;
}

/**
 * Доп.услуги с нужными short_name, заявленные на уровне ВСЕЙ услуги (item.
 * additional_services — то, что менеджер отметил как нужное для услуги в
 * целом). Не путать с assignment_technical.event_service_additional_service_list
 * у конкретной строки — та привязка появляется только после назначения
 * человека и относится к его конкретной оплате, а не к декларации на услуге.
 */
function additionalServiceShortNames_(item, wantedShortNames) {
  const list = item.additional_services || [];
  return list
    .map(x => x.additional_service && x.additional_service.short_name)
    .filter(sn => sn && wantedShortNames.includes(sn));
}

/**
 * Подпись роли на доске: сокращение амплуа, и для "Режиссера графики" —
 * с добавлением GTZp/БК, если такая доп.услуга заявлена на самой услуге.
 */
function roleLabel_(line, item) {
  const name = line.line ? line.line.name : '';
  let label = amplua_(name);
  if (name.includes('Режиссер графики')) {
    const badges = additionalServiceShortNames_(item, ['GTZp', 'БК']);
    if (badges.length > 0) label += ' ' + badges.join('/');
  }
  return label;
}

function executorName_(line) {
  const at = line.assignment_technical;
  if (at && at.employee) return `${at.employee.last_name} ${at.employee.first_name}`.trim();
  // Подрядчики (ИП/СЗ/АНО/ООО/ОГАУ), похоже, тоже приходят через assignment_technical.employee
  // (у некоторых last_name пустой, а всё название — в first_name, отсюда случайные
  // ведущие/замыкающие пробелы) — .trim() выше это лечит.
  // Ниже — запасной путь на случай, если для КАКИХ-ТО строк формат окажется другим.
  if (line.assignment_contractor && line.assignment_contractor.contractor) {
    return String(line.assignment_contractor.contractor.name || '').trim();
  }
  return '';
}

/** Статус принятия конкретного назначения: 'appointed' (зелёный в BMS, принял) /
 * 'pending' (жёлтый, отправлено — ждём подтверждения) / 'declined' (красный, отказался) / ''. */
function executorStatus_(line) {
  const at = line.assignment_technical;
  return at ? (at.status || '') : '';
}

/** Сводный статус по нескольким строкам сразу (для сокращённого блока подрядчика,
 * где на доске одна строка представляет сразу несколько исходных назначений):
 * если хоть один отказался — red; если все приняли — appointed; иначе pending. */
function aggregateStatus_(lines) {
  const statuses = lines.map(executorStatus_).filter(Boolean);
  if (statuses.includes('declined')) return 'declined';
  if (statuses.length > 0 && statuses.every(s => s === 'appointed')) return 'appointed';
  if (statuses.includes('pending')) return 'pending';
  return '';
}

/** Жирный шрифт для принятых, красный текст для отказавшихся, обычный — для ожидающих/пустых. */
function applyStatusStyle_(cell, status) {
  if (status === 'declined') {
    cell.setFontColor('#cc0000');
  } else if (status === 'appointed') {
    cell.setFontWeight('bold');
  }
}

/** Раскладывает строки услуги на роли аппаратной и видеооператоров, определяет "форму" блока. */
function classifyLines_(item) {
  const lines = (item.event_service_lines || []).filter(l =>
    !(l.line && l.line.name && l.line.name.includes('Выпускающий режиссер'))
  );
  const cameraLines = lines.filter(l => l.line && l.line.name && l.line.name.includes('Видеооператор'));
  const roleLines = lines
    .filter(l => !cameraLines.includes(l))
    .sort((a, b) => roleSortRank_(a.line ? a.line.name : '') - roleSortRank_(b.line ? b.line.name : ''));

  const allLines = roleLines.concat(cameraLines);
  const executors = allLines.map(executorName_).filter(Boolean);
  const uniqueExecutors = [...new Set(executors)];
  const isContractorPackage = uniqueExecutors.length === 1 && /^(ИП|СЗ|АНО|ООО|ОГАУ)(\s|$)/i.test(uniqueExecutors[0]);

  return { roleLines, cameraLines, uniqueExecutors, kind: isContractorPackage ? 'contractor' : 'full' };
}

/**
 * Рисует блок одного матча по заданным координатам (headerRow/col/startRow/kind).
 * kind='full' — Лига | Команда | N кам, затем амплуа (кроме Выпускающего режиссера):
 * сначала роли аппаратной (розовые), потом видеооператоры (голубые, по числу камер);
 * сама досоздаёт строки через ensureRoomForBlock_, если не хватает места.
 * kind='contractor' — вся техническая часть отдана одному подрядчику (ИП/СЗ/АНО/ООО/ОГАУ):
 * ровно 2 строки, заголовок + имя подрядчика, без досоздания места (столбик
 * подрядчиков резервируется заранее через reserveContractorSlot_).
 *
 * oldExecutorById — исполнители по line.id с прошлого обновления (из слепка в
 * реестре доски). Нужно, чтобы не затирать вручную вписанные черновые фамилии:
 * если роль как была, так и осталась не назначена в BMS — колонку исполнителя
 * для этой строки НЕ трогаем вообще (вдруг там черновик). Если раньше человек
 * был назначен, а теперь снят — это не черновик, а реальное снятие, тогда чистим.
 */
function renderMatchBlock_(sheet, day, headerRow, col, startRow, kind, event, item, boardStateMap, oldExecutorById) {
  const league = event.league ? event.league.short_name : '';
  const homeTeam = eventDisplayName_(event);
  const { roleLines, cameraLines, uniqueExecutors } = classifyLines_(item);
  const isCancelled = !!(item.service_status && item.service_status.indexOf('Отмен') === 0);

  if (kind === 'contractor') {
    // Блок подрядчика формируется только когда исполнитель у ВСЕХ строк один и
    // тот же реальный (не пустой) — черновиков тут по определению не бывает,
    // можно чистить и писать смело.
    sheet.getRange(startRow, col, 2, BOARD_GROUP_WIDTH).clearContent().clearFormat();
    sheet.getRange(startRow, col).setValue(league).setFontWeight('bold');
    setHeaderTeamCell_(sheet.getRange(startRow, col + 1), homeTeam, event.id, isCancelled, item.id);
    sheet.getRange(startRow, col + 2).setValue(`${cameraLines.length} кам`);
    const contractorCell = sheet.getRange(startRow + 1, col + 1).setValue(uniqueExecutors[0]).setBackground(BOARD_PINK);
    applyStatusStyle_(contractorCell, aggregateStatus_(roleLines.concat(cameraLines)));
    styleMatchBlock_(sheet, startRow, startRow + 1, col);
    return;
  }

  const neededRows = 1 + roleLines.length + cameraLines.length;
  ensureRoomForBlock_(sheet, day, headerRow, neededRows, boardStateMap);

  // после возможной вставки строк узнаём актуальную границу дня для очистки старого содержимого
  const map = loadDayRowMap_(sheet.getName()) || {};
  const nextHeaderRow = map[day + 1];
  const clearRows = nextHeaderRow ? (nextHeaderRow - startRow) : Math.max(neededRows, DEFAULT_DAY_ROWS - 1);

  // Колонки "роль" (label) и "N кам" — целиком выводятся из BMS, чистим смело.
  sheet.getRange(startRow, col, clearRows, 1).clearContent().clearFormat();
  sheet.getRange(startRow, col + 2, clearRows, 1).clearContent().clearFormat();
  // Колонка исполнителя — только формат (цвет/жирность сбрасываем), СОДЕРЖИМОЕ
  // не трогаем оптом: в пустых на сегодня строках может лежать черновая фамилия.
  sheet.getRange(startRow, col + 1, clearRows, 1).clearFormat();

  sheet.getRange(startRow, col).setValue(league).setFontWeight('bold');
  setHeaderTeamCell_(sheet.getRange(startRow, col + 1), homeTeam, event.id, isCancelled, item.id); // шапка — не черновая зона, пишем всегда
  sheet.getRange(startRow, col + 2).setValue(`${cameraLines.length} кам`);

  let row = startRow + 1;
  for (const line of roleLines) {
    sheet.getRange(row, col).setValue(roleLabel_(line, item));
    const cell = sheet.getRange(row, col + 1).setBackground(BOARD_PINK);
    const executor = executorName_(line);
    if (executor) {
      cell.setValue(executor);
      applyStatusStyle_(cell, executorStatus_(line));
    } else if (oldExecutorById && oldExecutorById[line.id]) {
      // раньше тут реально кто-то был назначен в BMS, теперь снят — это не
      // черновик, а настоящее снятие с назначения, чистим
      cell.setValue('');
    }
    // иначе — роль как была не назначена, так и осталась: не трогаем ячейку,
    // вдруг там вписанная вручную черновая фамилия
    row++;
  }
  for (const line of cameraLines) {
    const cell = sheet.getRange(row, col + 1).setBackground(BOARD_BLUE);
    const executor = executorName_(line);
    if (executor) {
      cell.setValue(executor);
      applyStatusStyle_(cell, executorStatus_(line));
    } else if (oldExecutorById && oldExecutorById[line.id]) {
      cell.setValue('');
    }
    row++;
  }
  styleMatchBlock_(sheet, startRow, row - 1, col);
}

const BOARD_NAME_COL_WIDTH_PX = 189; // ~5 см при 96 DPI — фиксированная, без авторасчёта

/**
 * Оформление блока матча: внешняя рамка по периметру, центрирование по
 * горизонтали и вертикали, перенос строк во всех колонках. Колонка с
 * командой/ФИО (col+1) — фиксированная ширина (не пересчитывается по
 * содержимому — колонка общая для разных матчей в разные дни, и авторасчёт
 * по содержимому "прыгал" туда-сюда в зависимости от того, какой матч
 * дописывался последним).
 */
function styleMatchBlock_(sheet, startRow, endRow, col) {
  const numRows = endRow - startRow + 1;
  const fullRange = sheet.getRange(startRow, col, numRows, BOARD_GROUP_WIDTH);

  fullRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  fullRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidth(col + 1, BOARD_NAME_COL_WIDTH_PX);
}

/** Чистит старое место блока (при переезде в новую позицию). Для подрядчика —
 * строго 2 строки (не задеть соседей по столбику), для обычного блока — с
 * запасом (у него всегда своя выделенная колонка, соседей там нет). */
function clearStaleBlock_(sheetName, startRow, col, kind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName(sheetName);
  if (!oldSheet) return;
  const rows = kind === 'contractor' ? 2 : 20;
  oldSheet.getRange(startRow, col, rows, BOARD_GROUP_WIDTH)
    .clearContent().clearFormat()
    .setBorder(false, false, false, false, false, false);
}

/** Слепок состояния услуги: по одному объекту на амплуа (кроме ВР), для сравнения между опросами. */
function lineSignature_(item) {
  const lines = (item.event_service_lines || []).filter(l =>
    !(l.line && l.line.name && l.line.name.includes('Выпускающий режиссер'))
  );
  return lines.map(l => ({
    id: l.id,
    name: l.line ? l.line.name : '',
    status: l.status,
    executor: executorName_(l),
  }));
}

/** Сравнивает старый и новый слепки, возвращает список человекочитаемых изменений. */
function diffSignatures_(oldSig, newSig) {
  const oldMap = {};
  (oldSig || []).forEach(l => { oldMap[l.id] = l; });

  const changes = [];
  for (const l of newSig) {
    const prev = oldMap[l.id];
    if (!prev) {
      if (l.executor) changes.push(`${amplua_(l.name)}: назначен ${l.executor}`);
      continue;
    }
    if (prev.executor !== l.executor || prev.status !== l.status) {
      changes.push(`${amplua_(l.name)}: ${prev.executor || '—'} → ${l.executor || '—'}`);
    }
  }
  return changes;
}

// ---- реестр положений блоков на доске (для перезаписи при изменениях) ----

/**
 * Удаляет реестр положений на доске (_bms_board_state) с нуля — нужно один раз
 * после смены схемы реестра (например, при добавлении новых полей), чтобы
 * старые неполные записи не путали логику размещения. Сами листы месяцев
 * не трогает — их пересоздать/перезаписать отдельно как обычно.
 */
function resetBoardState() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Сбросить реестр доски?',
    'Это удалит служебный лист _bms_board_state (положения уже отрисованных блоков) ' +
    'и все сохранённые "столбики" подрядчиков. Сами листы месяцев не изменятся, но при ' +
    'следующей перерисовке блоки могут переместиться на новые места. Продолжить?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.BOARD_STATE_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);

  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(key => {
    if (key.startsWith('CSTACK::')) props.deleteProperty(key);
  });

  ui.alert('Готово, реестр и столбики подрядчиков сброшены.');
}

/**
 * Ручной поиск замены: запрашивает ID мероприятия (event_id, не event_service_id),
 * находит все услуги этого мероприятия, для КАЖДОЙ строки со статусом "declined"
 * (независимо от того, впервые это или нет) показывает подбор замен — минуя
 * логику "уведомляем только на переходе", удобно для проверки/повторного запроса.
 */
function findReplacementsManually() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Найти замену', 'ID мероприятия (event_id из BMS):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const eventId = parseInt(resp.getResponseText().trim(), 10);
  if (!eventId) { ui.alert('Не понял ID: ' + resp.getResponseText()); return; }

  const items = fetchUpcomingAssignments_();
  const matches = items.filter(it => it.event.id === eventId);
  if (matches.length === 0) {
    ui.alert(`event_id=${eventId} не найден в горизонте LOOKAHEAD_DAYS (${CONFIG.LOOKAHEAD_DAYS} дней вперёд).`);
    return;
  }

  let sent = 0;
  for (const item of matches) {
    if (!item.service || !CONFIG.TARGET_SERVICES.includes(item.service.name)) continue;
    const declinedLines = lineSignature_(item).filter(l => l.status === 'declined');
    for (const line of declinedLines) {
      notifyReplacementCandidates_(matches[0].event, item, line);
      sent++;
    }
  }

  ui.alert(sent > 0
    ? `Готово, отправлено сообщений: ${sent}.`
    : 'Отказавшихся строк по нужным услугам не найдено — проверь ID или статус в BMS.');
}

const BOARD_MANUAL_BORDER = '#34a853'; // зелёный — ручная запись, ещё не подтверждена в BMS

/**
 * Вариант A для мероприятий, о которых известно заранее, но менеджер ещё не
 * внёс их в BMS: добавляет на доску простой блок (только заголовок — лига/
 * команда/камеры, без разбивки по ролям, т.к. это ещё не BMS-данные) с зелёной
 * рамкой — чтобы визуально отличать от автоматических записей. У блока нет
 * event_service_id, поэтому автоматика никогда не сможет сама его найти и
 * обновить/удалить — когда менеджер реально внесёт это мероприятие в BMS,
 * появится ОТДЕЛЬНЫЙ автоматический блок рядом, и эту ручную запись нужно
 * будет удалить самому.
 */
function addManualEntry() {
  const ui = SpreadsheetApp.getUi();

  const dateResp = ui.prompt('Ручная запись на доску', 'Дата (дд.мм.гггг)?', ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  const dateIso = parseDateInput_(dateResp.getResponseText().trim());
  if (!dateIso) { ui.alert('Не понял дату: ' + dateResp.getResponseText()); return; }

  const leagueResp = ui.prompt('Ручная запись на доску', 'Лига (можно оставить пустым)?', ui.ButtonSet.OK_CANCEL);
  if (leagueResp.getSelectedButton() !== ui.Button.OK) return;

  const teamResp = ui.prompt('Ручная запись на доску', 'Команда / название мероприятия?', ui.ButtonSet.OK_CANCEL);
  if (teamResp.getSelectedButton() !== ui.Button.OK) return;
  const teamName = teamResp.getResponseText().trim();
  if (!teamName) { ui.alert('Название не может быть пустым.'); return; }

  const camResp = ui.prompt('Ручная запись на доску', 'Число камер (можно оставить пустым)?', ui.ButtonSet.OK_CANCEL);
  if (camResp.getSelectedButton() !== ui.Button.OK) return;
  const cam = camResp.getResponseText().trim();

  const sheet = getOrCreateMonthSheet_(dateIso);
  const headerRow = findOrCreateDateHeaderRow_(sheet, dateIso);
  const day = new Date(dateIso).getDate();
  const col = findFreeColumnGroup_(sheet, headerRow + 1, day);
  const startRow = headerRow + 1;

  sheet.getRange(startRow, col).setValue(leagueResp.getResponseText().trim()).setFontWeight('bold');
  sheet.getRange(startRow, col + 1).setValue(`${teamName} (вручную)`).setFontWeight('bold');
  if (cam) sheet.getRange(startRow, col + 2).setValue(`${cam} кам`);

  const range = sheet.getRange(startRow, col, 1, BOARD_GROUP_WIDTH);
  range.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  range.setBorder(true, true, true, true, false, false, BOARD_MANUAL_BORDER, SpreadsheetApp.BorderStyle.SOLID_THICK);

  ui.alert(
    'Добавлено с зелёной рамкой. Когда менеджер внесёт то же самое в BMS, ' +
    'автоматика создаст рядом отдельный блок (связать их напрямую код не умеет) — ' +
    'не забудь тогда удалить эту ручную запись.'
  );
}

/**
 * Ищет положение блока по невидимой метке (event_service_id в заметке
 * заголовочной ячейки) в пределах строк одного дня — на случай, если блок
 * перенесли вручную и сохранённые координаты устарели. Проверяет только
 * "заголовочные" позиции колонок (col+1 у каждой группы из BOARD_GROUP_WIDTH).
 */
function findBlockPositionByNote_(sheet, headerRow, dayBottom, itemIdStr) {
  const numRows = dayBottom - headerRow;
  if (numRows <= 0) return null;
  const maxCol = 60; // с запасом на реальную ширину доски
  for (let col = 2; col <= maxCol; col += BOARD_GROUP_WIDTH) {
    const notes = sheet.getRange(headerRow + 1, col + 1, numRows, 1).getNotes();
    for (let i = 0; i < notes.length; i++) {
      if (notes[i][0] === itemIdStr) {
        return { col, startRow: headerRow + 1 + i };
      }
    }
  }
  return null;
}

function getBoardStateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.BOARD_STATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.BOARD_STATE_SHEET_NAME);
    sheet.hideSheet();
    sheet.appendRow(['event_service_id', 'event_id', 'sheet_name', 'header_row', 'col', 'start_row', 'kind', 'signature_json', 'updated_at']);
  }
  return sheet;
}

/** Загружает весь реестр в память: event_service_id -> {row, sheetName, headerRow, col, startRow, kind, signature}. */
function getBoardStateMap_() {
  const sheet = getBoardStateSheet_();
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  values.forEach((r, i) => {
    map[String(r[0])] = {
      row: i + 2, eventId: r[1], sheetName: r[2],
      headerRow: r[3], col: r[4], startRow: r[5], kind: r[6], signature: r[7],
    };
  });
  return map;
}

function upsertBoardState_(map, eventServiceId, eventId, sheetName, headerRow, col, startRow, kind, signatureJson) {
  const sheet = getBoardStateSheet_();
  const key = String(eventServiceId);
  const existing = map[key];
  const rowData = [eventServiceId, eventId, sheetName, headerRow, col, startRow, kind, signatureJson, new Date()];
  if (existing) {
    sheet.getRange(existing.row, 1, 1, 9).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  map[key] = {
    row: existing ? existing.row : sheet.getLastRow(),
    eventId, sheetName, headerRow, col, startRow, kind, signature: signatureJson,
  };
}

/**
 * Главная точка входа для доски: определяет форму блока (обычный/подрядчик),
 * находит или переиспользует его положение, перерисовывает и возвращает
 * список изменений по сравнению с прошлым опросом (пустой массив, если это
 * первая запись или ничего не изменилось).
 *
 * Позиция переиспользуется только если форма блока не изменилась с прошлого
 * раза — если, например, матч был "весь на подрядчике", а теперь на роли
 * назначили конкретных людей (или наоборот), блок переезжает на новое место
 * (старое место подчищается), т.к. компактный столбик подрядчиков не
 * рассчитан на то, что один из его блоков вдруг вырастет в полный.
 */
function upsertMatchBlock_(event, item, boardStateMap) {
  const key = String(item.id);
  let existing = boardStateMap[key];

  const sheet = getOrCreateMonthSheet_(event.date);
  const headerRow = findOrCreateDateHeaderRow_(sheet, event.date);
  const day = new Date(event.date).getDate();

  // Если помним прошлую позицию — проверим, что там всё ещё лежит именно этот
  // блок (по невидимой метке в заметке ячейки). Если нет — вероятно, блок
  // перенесли вручную в пределах того же дня; ищем его заново по всему дню и
  // подхватываем новое место, вместо того чтобы писать поверх старых координат.
  if (existing && existing.sheetName === sheet.getName()) {
    const storedNote = sheet.getRange(existing.startRow, existing.col + 1).getNote();
    if (storedNote !== key) {
      const dayMap = loadDayRowMap_(sheet.getName()) || {};
      const dayBottom = (dayMap[day + 1] || (headerRow + DEFAULT_DAY_ROWS)) - 1;
      const found = findBlockPositionByNote_(sheet, headerRow, dayBottom, key);
      existing = found ? Object.assign({}, existing, { col: found.col, startRow: found.startRow }) : null;
    }
  }

  const { kind } = classifyLines_(item);

  const samePlace = existing
    && existing.sheetName === sheet.getName()
    && existing.headerRow === headerRow
    && existing.kind === kind;

  let col, startRow;
  if (samePlace) {
    col = existing.col;
    startRow = existing.startRow;
  } else {
    if (existing) clearStaleBlock_(existing.sheetName, existing.startRow, existing.col, existing.kind);
    if (kind === 'contractor') {
      ({ col, startRow } = reserveContractorSlot_(sheet, day, headerRow));
    } else {
      col = findFreeColumnGroup_(sheet, headerRow + 1, day);
      startRow = headerRow + 1;
    }
  }

  const oldSig = existing ? JSON.parse(existing.signature || '[]') : [];
  const oldExecutorById = {};
  oldSig.forEach(l => { oldExecutorById[l.id] = l.executor; });

  renderMatchBlock_(sheet, day, headerRow, col, startRow, kind, event, item, boardStateMap, oldExecutorById);

  const newSig = lineSignature_(item);
  upsertBoardState_(boardStateMap, item.id, event.id, sheet.getName(), headerRow, col, startRow, kind, JSON.stringify(newSig));

  return {
    isNew: !existing,
    changes: diffSignatures_(oldSig, newSig),
    newlyDeclined: detectNewlyDeclined_(oldSig, newSig),
  };
}

/** Строки, которые ИМЕННО СЕЙЧАС стали "declined" (раньше были другим статусом). */
function detectNewlyDeclined_(oldSig, newSig) {
  const oldMap = {};
  (oldSig || []).forEach(l => { oldMap[l.id] = l; });
  return newSig.filter(l => l.status === 'declined' && (!oldMap[l.id] || oldMap[l.id].status !== 'declined'));
}

/**
 * Отказ от назначения: смотрит "График персонала" (кэш _staff_schedule_raw) —
 * все сотрудники с этим же амплуа, свободные в эту дату — и шлёт в Telegram
 * список кандидатов на замену, отсортированный по категории (звёздам) по убыванию.
 */
function notifyReplacementCandidates_(event, item, declinedLine) {
  const amplua = declinedLine.name;
  const ampluaShort = amplua_(amplua);
  const teams = eventDisplayName_(event);
  const serviceName = item.service ? item.service.name : '';

  const header = [
    `⚠️ <b>Отказ от назначения</b>`,
    `${event.date} ${event.start_time || ''} ${teams}`.trim(),
    serviceName,
    `${ampluaShort}: ${declinedLine.executor || '—'} — отказался`,
    '',
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(STAFF_RAW_SHEET_NAME);
  if (!raw || raw.getLastRow() < 2) {
    sendTelegramMessage_(header.concat(
      '⚠️ График персонала пуст — не могу предложить замену. Меню BMS → "Обновить график персонала сейчас".'
    ).join('\n'));
    return;
  }

  const data = raw.getRange(2, 1, raw.getLastRow() - 1, 8).getValues();
  const candidates = [];
  for (const r of data) {
    const [line, category, fullName, phone, , , date, cellText] = r;
    if (line !== amplua) continue;
    const dStr = date instanceof Date ? Utilities.formatDate(date, 'Europe/Moscow', 'yyyy-MM-dd') : String(date);
    if (dStr !== event.date) continue;
    if (cellText !== 'свободен') continue;
    if (fullName === declinedLine.executor) continue; // сам отказавшийся не в списке замен
    candidates.push({ fullName, category: String(category || ''), phone });
  }

  // по убыванию числа звёзд в категории ("****" впереди "**")
  candidates.sort((a, b) => b.category.length - a.category.length);

  const body = candidates.length === 0
    ? ['Свободных замен на эту дату по графику персонала не найдено.']
    : ['Возможные замены (по убыванию категории):',
      ...candidates.slice(0, 10).map(c =>
        `${c.category || '—'} ${c.fullName}${c.phone ? ' — ' + c.phone : ''}`
      )];

  sendTelegramMessage_(header.concat(body).join('\n'));
}

// ============================== ЧЕРНОВИКИ ДЛЯ ПЕРЕНОСА В BMS ==============================

/**
 * Ищет всех сотрудников из графика персонала (_staff_schedule_raw) с этим
 * амплуа, чьё full_name начинается с заданного префикса (без учёта отчества).
 */
/**
 * Ищет всех сотрудников из графика персонала (_staff_schedule_raw) с этим
 * амплуа, чьё full_name начинается с заданного префикса. Дедуп по паре
 * "имя+телефон", а не только по имени — если совпадение имени, но телефоны
 * разные, это настоящие тёзки (два разных человека), а не один и тот же
 * сотрудник, встреченный дважды — их нужно вернуть ОБОИХ, чтобы дальше это
 * ушло в разряд "неоднозначно", а не тихо схлопнулось в одного наугад.
 */
function findStaffByNamePrefix_(lineName, prefix) {
  const raw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_RAW_SHEET_NAME);
  if (!raw || raw.getLastRow() < 2) return [];
  const data = raw.getRange(2, 1, raw.getLastRow() - 1, 9).getValues();
  const seen = {};
  const results = [];
  for (const r of data) {
    const [line, , fullName, phone, , , , , employeeId] = r;
    if (line !== lineName) continue;
    if (!fullName || !fullName.startsWith(prefix)) continue;
    const dedupKey = `${fullName}::${phone}`;
    if (seen[dedupKey]) continue;
    seen[dedupKey] = true;
    results.push({ fullName, phone, employeeId });
  }
  return results;
}

/**
 * Смотрит на текст в ячейке исполнителя (для роли, которая в BMS всё ещё
 * not_appointed). Если текст — два слова в порядке "Имя Фамилия" (обратном
 * тому, что везде использует наш код и сама BMS — "Фамилия Имя") и совпадает
 * с реальным сотрудником из графика персонала по этому же амплуа — считаем
 * это черновиком для переноса в BMS. Если текст уже в порядке "Фамилия Имя" —
 * не трогаем (либо уже синхронизировано, либо инертный черновик в привычном виде).
 *
 * Возвращает: null (не наш случай), {fullName, phone} (однозначное совпадение),
 * или {ambiguous: true, candidates: [...]} (несколько совпадений сразу).
 */
function detectDraftAssignment_(cellText, lineName) {
  const parts = String(cellText || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length !== 2) return null; // работаем только с простым "Имя Фамилия" в два слова

  const [w1, w2] = parts;
  const asIs = `${w1} ${w2}`;
  const swapped = `${w2} ${w1}`;

  if (findStaffByNamePrefix_(lineName, asIs).length > 0) {
    return null; // уже в привычном порядке "Фамилия Имя" — не трогаем
  }

  const swappedMatches = findStaffByNamePrefix_(lineName, swapped);
  if (swappedMatches.length === 1) {
    return swappedMatches[0];
  }
  if (swappedMatches.length > 1) {
    return { ambiguous: true, candidates: swappedMatches };
  }
  return null; // не нашли ни разу — не похоже на реального сотрудника
}

/**
 * ПРОВЕРОЧНАЯ функция (не пишет в BMS — этого шага пока нет, ждём HAR с
 * реальным запросом назначения). Проходит по всем 'full'-блокам на доске за
 * LOOKAHEAD_DAYS, для каждой ещё не назначенной в BMS роли смотрит текущий
 * текст в ячейке и проверяет detectDraftAssignment_. Результат — в Telegram.
 */
/**
 * Проходит по всем 'full'-блокам на доске за LOOKAHEAD_DAYS, для каждой ещё
 * не назначенной в BMS роли смотрит текущий текст в ячейке. Однозначные
 * черновики ("Имя Фамилия", совпадающие с реальным сотрудником по этому же
 * амплуа) — СРАЗУ ПРИМЕНЯЕТ в BMS (реальная запись, см. assignEmployeeInBms_).
 * Неоднозначные (несколько тёзок) — не трогает, только сообщает.
 */
function scanDraftAssignments() {
  const ui = SpreadsheetApp.getUi();
  const items = fetchUpcomingAssignments_();
  const boardStateMap = getBoardStateMap_();
  const found = [];
  const ambiguous = [];

  for (const item of items) {
    if (!item.service || !CONFIG.TARGET_SERVICES.includes(item.service.name)) continue;
    const boardState = boardStateMap[String(item.id)];
    if (!boardState || boardState.kind !== 'full') continue; // блока подрядчика это не касается

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(boardState.sheetName);
    if (!sheet) continue;

    const { roleLines, cameraLines } = classifyLines_(item);
    let row = boardState.startRow + 1;
    for (const line of roleLines.concat(cameraLines)) {
      if (!executorName_(line)) { // роль ещё не назначена в BMS — тут может быть черновик
        const cellText = sheet.getRange(row, boardState.col + 1).getValue();
        const lineName = line.line ? line.line.name : '';
        const result = detectDraftAssignment_(cellText, lineName);
        if (result && result.ambiguous) {
          ambiguous.push({ event: item.event, line: lineName, cellText, candidates: result.candidates });
        } else if (result) {
          found.push({
            event: item.event, line: lineName, cellText, match: result,
            eventServiceId: item.id, eventServiceLineId: line.id,
            cell: sheet.getRange(row, boardState.col + 1),
          });
        }
      }
      row++;
    }
  }

  if (found.length === 0 && ambiguous.length === 0) {
    ui.alert('Черновиков для переноса не найдено.');
    return;
  }

  const applied = [];
  const failed = [];
  for (const f of found) {
    try {
      assignEmployeeInBms_(f.eventServiceId, f.eventServiceLineId, f.match.employeeId);
      f.cell.setValue(f.match.fullName); // приводим к обычному виду "Фамилия Имя", как из BMS
      applied.push(f);
    } catch (e) {
      failed.push({ ...f, error: String(e) });
    }
  }

  const lines = ['📝 <b>Черновики перенесены в BMS</b>', ''];
  applied.forEach(f => {
    lines.push(`✅ ${f.event.date} ${eventDisplayName_(f.event)}\n${amplua_(f.line)}: "${f.cellText}" → ${f.match.fullName}${f.match.phone ? ' (' + f.match.phone + ')' : ''}`);
  });
  failed.forEach(f => {
    lines.push(`❌ ${f.event.date} ${eventDisplayName_(f.event)}\n${amplua_(f.line)}: "${f.cellText}" → ошибка записи: ${f.error}`);
  });
  ambiguous.forEach(a => {
    lines.push(`⚠️ ${a.event.date} ${eventDisplayName_(a.event)}\n${amplua_(a.line)}: "${a.cellText}" → неоднозначно (${a.candidates.map(c => c.fullName).join(', ')})`);
  });

  sendTelegramMessage_(lines.join('\n\n'));
  ui.alert(`Готово. Применено: ${applied.length}, ошибок: ${failed.length}, неоднозначных: ${ambiguous.length}. Подробности в Telegram.`);
}

function formatChangeMessage_(event, item, changes) {
  const teams = eventDisplayName_(event);
  return [
    `🔄 <b>Изменение в мероприятии</b>`,
    `${event.date} ${event.start_time || ''} ${teams}`.trim(),
    item.service ? item.service.name : '',
    ...changes,
  ].filter(Boolean).join('\n');
}

/**
 * ТЕСТ: пишет/обновляет блок на доску для конкретного event_id, беря данные
 * из текущего assignment_table. Запускать из меню, например для event_id=14175.
 */
function test_WriteBoardForEvent() {
  const EVENT_ID = 14175; // поменяй на нужный

  const items = fetchUpcomingAssignments_();
  const matches = items.filter(it => it.event.id === EVENT_ID);
  if (matches.length === 0) {
    Logger.log(`event_id=${EVENT_ID} не найден в горизонте LOOKAHEAD_DAYS`);
    return;
  }
  const boardStateMap = getBoardStateMap_();
  for (const item of matches) {
    if (item.service && CONFIG.TARGET_SERVICES.includes(item.service.name)) {
      const { isNew, changes } = upsertMatchBlock_(item.event, item, boardStateMap);
      Logger.log(`event_id=${EVENT_ID}, service=${item.service.name}, isNew=${isNew}, changes=${JSON.stringify(changes)}`);
    }
  }
}

/**
 * Пишет/обновляет на доске ВСЕ мероприятия конкретной даты. Дата запрашивается
 * через диалоговое окно (формат ДД.ММ.ГГГГ или ГГГГ-ММ-ДД).
 */
function test_WriteBoardForDate() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Записать на доску',
    'Какую дату спарсить? (например 26.09.2026 или 2026-09-26)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const raw = resp.getResponseText().trim();
  const DATE = parseDateInput_(raw);
  if (!DATE) {
    ui.alert('Не понял дату: ' + raw);
    return;
  }

  const items = fetchUpcomingAssignments_();
  const dayItems = items.filter(it => it.event.date === DATE);
  const entries = groupByEvent_(dayItems);

  if (entries.length === 0) {
    ui.alert(`На ${DATE} мероприятий не найдено в горизонте LOOKAHEAD_DAYS (${CONFIG.LOOKAHEAD_DAYS} дней вперёд).`);
    return;
  }

  const boardStateMap = getBoardStateMap_();
  let written = 0;
  for (const { event, items: evItems } of entries) {
    const targetItems = evItems.filter(it =>
      it.service && CONFIG.TARGET_SERVICES.includes(it.service.name)
    );
    for (const item of targetItems) {
      upsertMatchBlock_(event, item, boardStateMap);
      written++;
    }
  }
  ui.alert(`Готово. Записано/обновлено блоков: ${written}`);
}

/** Принимает "26.09.2026" или "2026-09-26", возвращает "2026-09-26" (формат BMS) или null. */
function parseDateInput_(raw) {
  let m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return raw;
  m = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) {
    const [, dd, mm, yyyy] = m;
    return `${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`;
  }
  return null;
}

const AUTH_ALERT_COOLDOWN_MS = 3 * 60 * 60 * 1000; // не чаще раза в 3 часа, чтобы не спамить каждые 5 минут

/** Шлёт в Telegram алерт о протухшем входе, но не чаще AUTH_ALERT_COOLDOWN_MS. */
function notifyAuthFailure_(errorMessage) {
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('LAST_AUTH_ALERT_AT') || 0);
  const now = Date.now();
  if (now - last < AUTH_ALERT_COOLDOWN_MS) return;
  props.setProperty('LAST_AUTH_ALERT_AT', String(now));
  sendTelegramMessage_(
    `⚠️ <b>BMS: нужен повторный вход</b>\nАвтоматическая проверка мероприятий остановлена.\n` +
    `В таблице: меню BMS → "🔑 Войти в BMS".\n\n${String(errorMessage).substring(0, 300)}`
  );
}

// ============================== ГРАФИК ПЕРСОНАЛА ==============================

const STAFF_SHEET_NAME = 'График персонала';

function getEmployeeDeptInfo_(employee) {
  const jt = employee.user_job_titles && employee.user_job_titles[0];
  const title = jt && jt.job_title ? jt.job_title.name : '';
  const dept = jt && jt.job_title && jt.job_title.job_title_block && jt.job_title.job_title_block.unit
    ? jt.job_title.job_title_block.unit.name : '';
  return { title, dept };
}

/** Список ISO-дат [dateFrom..dateTo] включительно. */
function dateRangeList_(dateFromIso, dateToIso) {
  const from = new Date(dateFromIso);
  const to = new Date(dateToIso);
  const list = [];
  for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
    list.push(Utilities.formatDate(d, 'Europe/Moscow', 'yyyy-MM-dd'));
  }
  return list;
}

/** Короткое описание одного назначения на день, для ячейки графика. */
function assignmentLabel_(a) {
  const es = a.event_service;
  const ev = es ? es.event : null;
  if (!ev) return es && es.service ? es.service.name : 'занят';

  const league = ev.league ? ev.league.short_name : '';
  const teams = ev.home_team && ev.away_team
    ? `${ev.home_team.name}-${ev.away_team.name}`
    : (ev.event_description || (ev.sport_type ? ev.sport_type.name : ''));
  const time = ev.start_time ? ev.start_time.substring(0, 5) : '';

  return [league, teams, time].filter(Boolean).join(' ');
}

/** Что писать в ячейку графика для одного дня сотрудника. */
function dayCellValue_(day) {
  if (!day) return '';
  if (day.is_dayoff && (!day.assignments || day.assignments.length === 0)) return 'не может';
  if (!day.assignments || day.assignments.length === 0) return 'свободен';
  return day.assignments.map(assignmentLabel_).join('; ');
}

/**
 * Строит лист "График персонала": одна строка = один сотрудник + одно его
 * амплуа, дальше по одной колонке на КАЖДУЮ дату выбранного периода — в
 * ячейке либо чем занят (мероприятие), либо "свободен", либо "не может"
 * (соответствует датам, обведённым красным в самой BMS). Сортировка/фильтр
 * по колонке "Амплуа" сразу группирует людей по ролям.
 *
 * Данные из BMS тянутся не при каждом просмотре, а РАЗ В СУТКИ (по триггеру
 * dailyStaffScheduleFetch, ставится вручную на 9:00 МСК) на скользящий месяц
 * вперёд от текущей даты — и складываются в скрытый кэш (STAFF_RAW_SHEET_NAME).
 * Сам видимый лист просто фильтрует уже готовый кэш под период, который можно
 * поменять прямо в ячейках B1/D1 листа — правка этих ячеек сама пересобирает
 * таблицу (см. onEdit), новых обращений к BMS API при этом не происходит.
 */

const STAFF_RAW_SHEET_NAME = '_staff_schedule_raw'; // скрытый кэш "сырых" данных из BMS
const STAFF_FETCH_MONTHS_AHEAD = 1; // на сколько вперёд от сегодня тянуть при ежедневном обновлении

/**
 * Тянет данные из BMS по всем амплуа за период и сохраняет плоским списком
 * (по строке на сотрудника+амплуа+дату) в скрытый лист-кэш — независимо от
 * того, какой период потом захотят посмотреть на видимом листе.
 */
function fetchAndCacheStaffSchedule_(dateFrom, dateTo) {
  const lines = bmsGet_('admin/catalog/line/list', { page: 1, size: 100 }).items;
  const dates = dateRangeList_(dateFrom, dateTo);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let raw = ss.getSheetByName(STAFF_RAW_SHEET_NAME);
  if (raw) {
    raw.clear();
  } else {
    raw = ss.insertSheet(STAFF_RAW_SHEET_NAME);
    raw.hideSheet();
  }
  raw.appendRow(['line', 'category', 'full_name', 'phone', 'title', 'dept', 'date', 'cell_text', 'employee_id']);

  const rows = [];
  for (const line of lines) {
    let entries;
    try {
      entries = bmsGet_('admin/employee/schedule', {
        date_from: dateFrom, date_to: dateTo, line_id__in: line.id,
      });
    } catch (e) {
      Logger.log(`Пропущено амплуа "${line.name}": ${e}`);
      continue;
    }
    for (const entry of entries) {
      const emp = entry.employee;
      const { title, dept } = getEmployeeDeptInfo_(emp);
      const category = (emp.user_line_categories || []).find(c => c.line && c.line.id === line.id);
      const fullName = `${emp.last_name} ${emp.first_name} ${emp.middle_name || ''}`.trim();
      const phone = emp.phone_number || '';
      const daysByDate = {};
      (entry.days || []).forEach(d => { daysByDate[d.date] = d; });

      for (const d of dates) {
        rows.push([line.name, category ? category.category : '', fullName, phone, title, dept, d, dayCellValue_(daysByDate[d]), emp.id]);
      }
    }
    Utilities.sleep(150); // не долбим API слишком часто — 31 амплуа подряд
  }

  if (rows.length > 0) {
    raw.getRange(2, 1, rows.length, 9).setValues(rows);
  }

  PropertiesService.getScriptProperties().setProperty(
    'STAFF_CACHE_RANGE', JSON.stringify({ dateFrom, dateTo, fetchedAt: new Date().toISOString() })
  );
}

/** Принимает значение ячейки (Date или текст дд.мм.гггг/гггг-мм-дд), возвращает ISO 'yyyy-MM-dd' или null. */
function parseSheetDate_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Europe/Moscow', 'yyyy-MM-dd');
  if (!val) return null;
  return parseDateInput_(String(val).trim());
}

/**
 * Перерисовывает видимый лист "График персонала" из уже накопленного кэша —
 * без обращений к BMS API. Период берёт из ячеек B1 (с) / D1 (по) листа; если
 * они пустые — использует весь диапазон, что сейчас в кэше.
 */
function renderStaffScheduleView_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(STAFF_RAW_SHEET_NAME);
  if (!raw || raw.getLastRow() < 2) return; // кэша ещё нет — нечего показывать

  let sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(STAFF_SHEET_NAME);

  const cacheMeta = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAFF_CACHE_RANGE') || '{}');
  let dateFrom = parseSheetDate_(sheet.getRange(1, 2).getValue()) || cacheMeta.dateFrom;
  let dateTo = parseSheetDate_(sheet.getRange(1, 4).getValue()) || cacheMeta.dateTo;
  if (!dateFrom || !dateTo) return;

  const dates = dateRangeList_(dateFrom, dateTo);
  const dateSet = {};
  dates.forEach(d => { dateSet[d] = true; });

  const rawData = raw.getRange(2, 1, raw.getLastRow() - 1, 8).getValues();
  const groups = {};
  const order = [];
  for (const r of rawData) {
    const [line, category, fullName, phone, title, dept, date, cellText] = r;
    const dStr = date instanceof Date ? Utilities.formatDate(date, 'Europe/Moscow', 'yyyy-MM-dd') : String(date);
    if (!dateSet[dStr]) continue; // строка кэша вне выбранного периода — пропускаем
    const key = `${line}|||${fullName}`;
    if (!groups[key]) {
      groups[key] = { line, category, fullName, phone, title, dept, byDate: {} };
      order.push(key);
    }
    groups[key].byDate[dStr] = cellText;
  }

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();

  sheet.getRange(1, 1).setValue('С:');
  sheet.getRange(1, 2).setValue(dateFrom).setNumberFormat('@STRING@');
  sheet.getRange(1, 3).setValue('По:');
  sheet.getRange(1, 4).setValue(dateTo).setNumberFormat('@STRING@');
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(1, 5).setValue('← период можно поменять здесь, таблица перестроится сама');

  const dateHeaders = dates.map(d => Utilities.formatDate(new Date(d), 'Europe/Moscow', 'dd.MM'));
  const header = ['Амплуа', 'Категория', 'ФИО', 'Телефон', 'Должность', 'Департамент', ...dateHeaders];
  sheet.getRange(2, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(6);

  const outRows = order.map(key => {
    const g = groups[key];
    return [g.line, g.category, g.fullName, g.phone, g.title, g.dept, ...dates.map(d => g.byDate[d] || '')];
  });

  if (outRows.length > 0) {
    sheet.getRange(3, 1, outRows.length, header.length).setValues(outRows);
    sheet.getRange(2, 1, outRows.length + 1, header.length).createFilter();

    const dataRange = sheet.getRange(3, 7, outRows.length, dates.length);
    const rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('не может').setBackground('#f4cccc').setRanges([dataRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('свободен').setBackground('#d9ead3').setRanges([dataRange]).build());
    sheet.setConditionalFormatRules(rules);
  }
  sheet.autoResizeColumns(1, 6);
}

/**
 * Ежедневное обновление (ставится вручную триггером: Time-driven, Day timer,
 * 9:00 МСК) — тянет из BMS скользящий месяц вперёд от сегодняшней даты и
 * перерисовывает видимый лист текущим выбранным (или дефолтным) периодом.
 */
function dailyStaffScheduleFetch() {
  const today = new Date();
  const dateFrom = Utilities.formatDate(today, 'Europe/Moscow', 'yyyy-MM-dd');
  const until = new Date(today);
  until.setMonth(until.getMonth() + STAFF_FETCH_MONTHS_AHEAD);
  const dateTo = Utilities.formatDate(until, 'Europe/Moscow', 'yyyy-MM-dd');

  fetchAndCacheStaffSchedule_(dateFrom, dateTo);
  renderStaffScheduleView_();
}

/** Ручной пункт меню: немедленно обновить кэш (на скользящий месяц вперёд) и перерисовать лист. */
function refreshStaffScheduleNow() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const dateFrom = Utilities.formatDate(today, 'Europe/Moscow', 'yyyy-MM-dd');
  const until = new Date(today);
  until.setMonth(until.getMonth() + STAFF_FETCH_MONTHS_AHEAD);
  const dateTo = Utilities.formatDate(until, 'Europe/Moscow', 'yyyy-MM-dd');

  fetchAndCacheStaffSchedule_(dateFrom, dateTo);
  renderStaffScheduleView_();
  ui.alert(`Готово. В кэше данные ${dateFrom} — ${dateTo}. Период показа можно менять в ячейках B1/D1 листа "${STAFF_SHEET_NAME}".`);
}

/**
 * Простой триггер Apps Script — срабатывает автоматически при любой ручной
 * правке ячейки в таблице, без отдельной установки. Если поменяли период (B1
 * или D1) на листе "График персонала" — перестраиваем таблицу из кэша.
 */
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== STAFF_SHEET_NAME) return;
    if (e.range.getRow() !== 1) return;
    if (e.range.getColumn() !== 2 && e.range.getColumn() !== 4) return;
    renderStaffScheduleView_();
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/**
 * Обрабатывает одно мероприятие (событие + его услуги): решает, новое ли оно,
 * пишет "Новые мероприятия"/"Тех блок"/доску, шлёт Telegram. Общая точка входа
 * и для doPoll() (когда опрос BMS делает сам Apps Script), и для doPost()
 * (когда опрос BMS вынесен в GitHub Actions, а сюда прилетают уже готовые
 * данные по HTTP).
 */
function processEventEntry_(event, items, seenEvents, boardStateMap) {
  const isNewEvent = !seenEvents.has(String(event.id));

  if (isNewEvent) {
    sendTelegramMessage_(formatEventMessage_(event));
    try {
      appendNewEventRow_(event, items[0]);
    } catch (e) {
      Logger.log(`Не удалось записать строку "Новые мероприятия" event_id=${event.id}: ${e}`);
    }
    markEventSeen_(event.id);
  }

  const targetItems = items.filter(it =>
    it.service && CONFIG.TARGET_SERVICES.includes(it.service.name)
  );

  for (const item of targetItems) {
    let result;
    try {
      result = upsertMatchBlock_(event, item, boardStateMap);
    } catch (e) {
      Logger.log(`Не удалось записать на доску event_id=${event.id}, service=${item.service.name}: ${e}`);
      continue;
    }

    const somethingChanged = result.isNew || result.changes.length > 0;
    if (!somethingChanged) continue;

    try {
      appendTechBlockRows_(event, item);
    } catch (e) {
      Logger.log(`Не удалось записать тех.блок event_id=${event.id}, service=${item.service.name}: ${e}`);
    }

    if (!isNewEvent && result.changes.length > 0) {
      sendTelegramMessage_(formatChangeMessage_(event, item, result.changes));
    } else if (!isNewEvent && result.isNew) {
      sendTelegramMessage_(
        `➕ В мероприятии ${event.date} ${eventDisplayName_(event)} добавлена услуга "${item.service.name}"`.trim()
      );
    }

    for (const declinedLine of (result.isNew ? [] : result.newlyDeclined)) {
      try {
        notifyReplacementCandidates_(event, item, declinedLine);
      } catch (e) {
        Logger.log(`Не удалось подобрать замену event_id=${event.id}, line=${declinedLine.name}: ${e}`);
      }
    }
  }
}

function doPoll() {
  const seenEvents = getSeenEventIds_();
  const boardStateMap = getBoardStateMap_();

  let entries;
  try {
    entries = groupByEvent_(fetchUpcomingAssignments_());
  } catch (e) {
    notifyAuthFailure_(e);
    return;
  }

  for (const { event, items } of entries) {
    processEventEntry_(event, items, seenEvents, boardStateMap);
  }
}

/**
 * Веб-эндпоинт (HTTP POST) — принимает уже готовые данные от внешнего опроса
 * (GitHub Actions). Тело запроса: {"secret": "...", "events": [{"event":...,
 * "items":[...]}]}. Проверяет секрет (WEBHOOK_SECRET в Script Properties, см.
 * setupSecrets), дальше работает точно так же, как doPoll — только сам не
 * ходит в BMS API, данные уже пришли готовые.
 *
 * УСТАНОВКА: Deploy -> New deployment -> Web app, Execute as: Me, Who has
 * access: Anyone. Полученный URL — это и есть APPS_SCRIPT_WEBHOOK_URL для
 * GitHub Actions.
 */
function doPost(e) {
  // Блокировка: если два запроса (например, ручной запуск GitHub Actions
  // наложился на плановый по крону) прилетят почти одновременно, второй
  // дождётся, пока первый допишет данные — без этого оба могли бы одновременно
  // не увидеть изменений друг друга и задвоить мероприятия на доске.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'busy, retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== getSecret_('WEBHOOK_SECRET')) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const seenEvents = getSeenEventIds_();
    const boardStateMap = getBoardStateMap_();

    for (const entry of body.events || []) {
      processEventEntry_(entry.event, entry.items, seenEvents, boardStateMap);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}



// ============================== ОТЛАДКА СТРУКТУРЫ ДОСКИ ==============================

/**
 * Выгружает точную структуру выделенного диапазона: значения, HEX-цвета
 * заливки/шрифта, жирность, объединённые ячейки. Пустые белые ячейки
 * пропускает, чтобы не раздувать вывод.
 *
 * КАК ПОЛЬЗОВАТЬСЯ:
 * 1. На листе с доской выдели мышкой один блок дня (например, весь 15.08 —
 *    от серой строки-шапки до последней использованной строки/колонки).
 * 2. Перейди в Apps Script (не кликай обратно в таблицу — собьётся выделение).
 * 3. Выбери в списке функций debugDumpSelection -> Выполнить.
 * 4. View -> Logs (или Ctrl+Enter), скопируй весь JSON, пришли мне.
 */
function debugDumpSelection() {
  const range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  const sheet = range.getSheet();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontWeights = range.getFontWeights();

  const cells = [];
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const v = values[r][c];
      const bg = backgrounds[r][c];
      if (v === '' && (bg === '#ffffff' || bg === null)) continue; // пропускаем пустые белые
      cells.push({
        row: startRow + r,
        col: startCol + c,
        value: v,
        bg: bg,
        fontColor: fontColors[r][c],
        bold: fontWeights[r][c] === 'bold',
      });
    }
  }

  const merges = sheet.getRange(startRow, startCol, numRows, numCols)
    .getMergedRanges()
    .map(m => ({
      row: m.getRow(), col: m.getColumn(),
      numRows: m.getNumRows(), numCols: m.getNumColumns(),
    }));

  Logger.log(JSON.stringify({ sheet: sheet.getName(), cells, merges }));
}
