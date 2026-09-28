/**
 * Приём заявок с сайта «Лаванда» -> Google Таблица.
 * Файл: Google Таблица -> Расширения -> Apps Script -> Code.gs
 *
 * ---------------------------------------------------------------------------
 * ВАЖНО ПРИ УСТАНОВКЕ
 * ---------------------------------------------------------------------------
 * Вставьте этот код ВМЕСТО всего, что было в файле.
 * Откройте Code.gs, нажмите Ctrl+A, потом Delete, и только затем вставьте.
 *
 * Если старый код останется в файле, редактор покрасит его в красный:
 * функция doPost будет объявлена дважды. Apps Script на такое ругается
 * сразу на всех строках, поэтому красным кажется весь файл.
 *
 * Проверка, что всё вставилось правильно: в выпадающем списке функций
 * сверху должно быть ровно 21 имя, перечисленное в конце этого файла.
 *
 * ---------------------------------------------------------------------------
 * ФОРМАТ ДАННЫХ, КОТОРЫЙ ПРИХОДИТ С САЙТА (index.html)
 * ---------------------------------------------------------------------------
 * Форма отправляет один POST на адрес из настройки BOOKING_ENDPOINT:
 *
 *     POST https://script.google.com/macros/s/AKfycb.../exec
 *     Content-Type:  text/plain
 *     Тело:          одна строка JSON, примерно 200-250 байт
 *
 *     {
 *       "source":      "Забронировать 90 минут тишины за 2 900 руб.",
 *       "firstName":   "Анна",
 *       "lastName":    "Петрова",
 *       "phone":       "+7 (999) 123-45-67",
 *       "consent":     "yes",
 *       "page":        "https://...",
 *       "submittedAt": "2026-09-27T10:58:55.948Z"
 *     }
 *
 * Имена полей скрипт читает именно такие, с этими буквами.
 * В таблицу они попадают в колонки «Имя», «Фамилия», «Телефон»,
 * «Согласие», «Откуда», «Страница», «Отправлено с сайта».
 *
 * ---------------------------------------------------------------------------
 * ПОЧЕМУ ПЕРЕД ТЕЛЕФОМ СТАВИТСЯ АПОСТРОФ
 * ---------------------------------------------------------------------------
 * Google Таблицы считают строку, начинающуюся с +, началом формулы.
 * Номер +7 (999) 123-45-67 превращался в #ERROR! «синтаксическая ошибка
 * в формуле». Апостроф в начале значения заставляет Таблицы считать
 * ячейку текстом. В самой ячейке апострофа не видно, там остаётся
 * +7 (999) 123-45-67. Поэтому в appendToSheet телефон пишется так:
 *
 *     data.phone ? "'" + data.phone : ''
 *
 * Такими же апострофами в самой таблице ничего дописывать не нужно.
 *
 * ---------------------------------------------------------------------------
 * ПОЧЕМУ ТЕЛО JSON, НО contentType text/plain
 * ---------------------------------------------------------------------------
 * Apps Script не отдаёт браузеру CORS-заголовки. Если отправить
 * application/json, браузер спросит разрешения, не получит его и
 * вообще не отправит заявку. С text/plain запрос уходит без проверки.
 * Скрипт принимает и такой запрос, и form-encoded, и пустой вызов -
 * и никогда не падает с ошибкой 500, иначе форма на сайте напишет
 * «Не удалось отправить», а заявка потеряется.
 *
 * ---------------------------------------------------------------------------
 * ПРАВИЛА ЗАПИСИ В ЛИСТ «ЗАЯВКИ»
 * ---------------------------------------------------------------------------
 *   данных в запросе нет вовсе      -> строка НЕ пишется, пишется «Журнал»
 *   тело есть, разобралось         -> строка со статусом «ок»
 *   тело есть, не разобралось      -> строка со статусом
 *                                      «НЕ РАСПОЗНАНО: причина»
 *                                      плюс подробности в «Журнале»
 *
 * Такой порядок нужен, потому что раньше вызов doPost без аргументов
 * (кнопка «Выполнить» при выбранной doPost) создавал в таблице строку
 * с одной датой и жалобой. Теперь такой вызов не создаёт ничего,
 * кроме записи в «Журнале».
 *
 * ---------------------------------------------------------------------------
 * СПИСОК ФУНКЦИЙ ЭТОГО ФАЙЛА (21)
 * ---------------------------------------------------------------------------
 *   doPost doGet readData hasName normalize fieldName parseQuery
 *   describeEvent appendToSheet getBook getSheet currentHeader
 *   writeHeader sendMail sendTelegram writeLog
 *   testFromSite testEmpty testAllFormats setup rebuildSheets
 */


var SETTINGS = {

  // Куда складывать.
  // Пусто - скрипт работает с той таблицей, из которой создан.
  // Отдельный скрипт: вставьте кусок между /d/ и /edit в адресе таблицы.
  SHEET_ID: '',

  SHEET_NAME: 'Заявки',
  LOG_SHEET_NAME: 'Журнал',

  // Порядок колонок в листе «Заявки» менять нельзя: по нему идёт запись.
  // «Согласие» - доказательство согласия на обработку персональных
  // данных (152-ФЗ). Форма передаёт «yes», когда человек поставил
  // галочку под текстом согласия.
  HEADERS: ['Дата', 'Имя', 'Фамилия', 'Телефон', 'Согласие', 'Откуда', 'Страница', 'Отправлено с сайта', 'Статус'],

  // Почта.
  MAIL_TO: 'insigt159@gmail.com',
  MAIL_SUBJECT: 'Новая заявка — Лаванда',

  // Телеграм. Если не заполнено, сообщения просто не отправляются.
  TELEGRAM_TOKEN: '',
  TELEGRAM_CHAT: '',

  // Пока true, в лист «Журнал» пишется ТЕЛО каждого запроса целиком,
  // то есть имя, фамилия и телефон. Это персональные данные, поэтому
  // по умолчанию выключено. Включайте на минуту, только когда нужно
  // разобраться, почему заявка не дошла, и сразу ставьте обратно false.
  DEBUG: false
};


// Тело ровно такое, какое присылает форма в index.html при
// text/plain + no-cors. Снято с живого сайта.
// Если после testFromSite() в таблице появится строка с именем
// и телефоном - скрипт и форма совместимы.
var SITE_BODY = '{"source":"Забронировать 90 минут тишины за 2 900 руб.",'
  + '"firstName":"Анна","lastName":"Петрова","phone":"+7 (999) 123-45-67",'
  + '"consent":"yes","page":"https://lavanda.ru/",'
  + '"submittedAt":"2026-09-27T10:58:55.948Z"}';


/**
 * Точка входа. Сюда прилетает POST из формы.
 * Всегда заканчивается ответом 200, иначе сайт покажет ошибку.
 */
function doPost(e) {
  var parsed = readData(e);
  var data = parsed.data;
  var log = [];

  // Данных в запросе нет вообще. Ни с сайта, ни откуда.
  // Строку в «Заявках» не создаём, только пишем в «Журнал».
  if (parsed.empty) {
    writeLog('ЗАПРОС БЕЗ ДАННЫХ - строка в «' + SETTINGS.SHEET_NAME + '» не создана.\n'
      + describeEvent(e));
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  if (SETTINGS.DEBUG) log.push('DEBUG:\n' + describeEvent(e));
  if (parsed.problem) log.push('РАЗБОР: ' + parsed.problem);

  try {
    appendToSheet(data, parsed.problem ? ('НЕ РАСПОЗНАНО: ' + parsed.problem) : 'ок');
  } catch (error) {
    log.push('ТАБЛИЦА: ' + error);
  }

  try {
    sendMail(data, parsed.problem);
  } catch (error) {
    log.push('ПОЧТА: ' + error);
  }

  try {
    sendTelegram(data);
  } catch (error) {
    log.push('ТЕЛЕГРАМ: ' + error);
  }

  if (log.length) writeLog(log.join('\n\n'));

  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Открыть адрес развёртывания в браузере - проверка, что всё живо.
 * Приходит на GET, строку в «Заявках» не пишет.
 */
function doGet() {
  return ContentService.createTextOutput('Работает. Ждём заявку на ' + SETTINGS.SHEET_NAME
    + ' и ' + SETTINGS.LOG_SHEET_NAME + '.')
    .setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Разбор запроса. Никогда не бросает исключений.
 *
 * Порядок попыток:
 *   1) e.parameter - заполняется Apps Script сам, если запрос
 *      form-encoded. Это самый надёжный путь.
 *   2) JSON.parse тела - так шлёт сайт (text/plain, no-cors).
 *   3) Разбор тела как a=1&b=2 - запасной вариант.
 *
 * Возвращает объект:
 *   data        - поля заявки с приведёнными именами
 *   problem     - '' если всё хорошо, иначе описание проблемы
 *   empty       - true, если данных в запросе не было вовсе
 *   contentType - Content-Type запроса
 */
function readData(e) {
  var result = { data: {}, problem: '', empty: true, contentType: '—' };

  try {
    if (!e) {
      result.problem = 'doPost вызван без аргумента (запуск из редактора)';
      return result;
    }

    var body = '';
    if (e.postData) {
      result.contentType = String(e.postData.type || '—');
      body = String(e.postData.contents || '');
    }

    var hasParameter = !!(e.parameter && Object.keys(e.parameter).length);

    // 1) нативные параметры Apps Script
    if (hasParameter) {
      var fromParameter = normalize(e.parameter);
      if (hasName(fromParameter)) {
        result.data = fromParameter;
        result.empty = false;
        return result;
      }
    }

    if (!body) {
      result.problem = 'тело запроса пустое (браузер отправил POST без данных)';
      return result;
    }

    result.empty = false;

    // 2) JSON - именно так отправляет форма в index.html
    try {
      var parsed = JSON.parse(body);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        result.data = normalize(parsed);
        if (!hasName(result.data)) {
          // Само тело в жалобу не пишем: там имя и телефон. Достаточно
          // перечислить, какие ключи пришли.
          result.problem = 'JSON разобран, но полей firstName/lastName/phone в нём нет.'
            + ' Пришли ключи: ' + Object.keys(parsed).join(', ');
        }
        return result;
      }
    } catch (jsonError) {
      // не JSON - пробуем ниже
    }

    // 3) form-запрос. Принимаем, только если нашлись хоть какие-то поля.
    var fromQuery = normalize(parseQuery(body));
    if (hasName(fromQuery)) {
      result.data = fromQuery;
      return result;
    }

    // Тело сюда не пишем целиком: это персональные данные, а в журнал
    // им нельзя. Хватает типа и длины. Посмотреть тело можно, включив
    // SETTINGS.DEBUG на минуту.
    result.problem = 'не удалось распознать тело, ' + result.contentType
      + ', ' + body.length + ' символов. Чтобы увидеть тело, включите SETTINGS.DEBUG.';
  } catch (error) {
    result.problem = 'ошибка при разборе: ' + error;
    result.empty = false;
  }

  return result;
}


/**
 * Есть ли в данных хоть одно имя или телефон.
 * Нужно, чтобы мусор в теле запроса не проходил как успешный разбор.
 */
function hasName(data) {
  return !!(data.firstName || data.lastName || data.phone);
}


/**
 * Приводит ключи к тем именам, которые читает таблица.
 * Регистр не важен: firstName, firstname, FirstName -> firstName.
 */
function normalize(obj) {
  var out = {};
  Object.keys(obj).forEach(function (key) {
    out[fieldName(key)] = obj[key];
  });
  return out;
}


/**
 * Имя поля, которое читает таблица, по любому варианту написания.
 * Список намеренно разбросан по switch, а не собран в объект:
 * так в коде не остаётся ни одной кириллической клавиши.
 */
function fieldName(key) {
  var k = String(key).toLowerCase().replace(/[-\s]/g, '_');
  switch (k) {
    case 'firstname':
    case 'first_name':
    case 'name':
    case 'имя':
      return 'firstName';
    case 'lastname':
    case 'last_name':
    case 'surname':
    case 'фамилия':
      return 'lastName';
    case 'phone':
    case 'tel':
    case 'mobile':
    case 'telephone':
    case 'телефон':
      return 'phone';
    case 'source':
      return 'source';
    case 'page':
    case 'url':
      return 'page';
    case 'submittedat':
      return 'submittedAt';
    case 'consent':
      return 'consent';
    default:
      return k;
  }
}


/** Разбирает строку a=1&b=2 в объект. */
function parseQuery(raw) {
  var out = {};
  String(raw).split('&').forEach(function (pair) {
    if (!pair) return;

    // Режем по ПЕРВОМУ знаку =, чтобы значение с = не потерялось.
    var cut = pair.indexOf('=');
    var key = cut < 0 ? pair : pair.slice(0, cut);
    var value = cut < 0 ? '' : pair.slice(cut + 1);
    if (!key) return;

    try {
      out[decodeURIComponent(key)] = decodeURIComponent(value.replace(/\+/g, ' '));
    } catch (error) {
      out[key] = value;
    }
  });
  return out;
}


/**
 * Описание запроса словами - попадает в «Журнал».
 * По этой записи видно, что именно пришло на сервер.
 */
function describeEvent(e) {
  if (!e) {
    return 'события нет: doPost вызван без аргументов.\n'
      + 'Это запуск из редактора, а не заявка с сайта. Такие вызовы\n'
      + 'в лист «' + SETTINGS.SHEET_NAME + '» больше не попадают.\n'
      + 'Для проверки запустите testFromSite - он имитирует запрос сайта.';
  }

  var lines = [];

  lines.push('действие:     ' + (e.action || '—'));
  lines.push('длина запроса: ' + (e.contentLength === undefined ? '—' : e.contentLength));
  lines.push('параметры:     ' + (e.parameter && Object.keys(e.parameter).length
    ? Object.keys(e.parameter).join(', ')
    : 'нет'));

  if (e.postData) {
    lines.push('content-type:  ' + (e.postData.type || '—'));
    lines.push('длина тела:    ' + String(e.postData.contents || '').length);
    lines.push('тело:          ' + String(e.postData.contents || ''));
  } else {
    lines.push('postData:      нет - запрос пришёл вообще без тела');
  }

  return lines.join('\n');
}


/**
 * Дописывает строку в лист «Заявки». Лист и шапка создаются сами.
 *
 * Порядок значений совпадает с SETTINGS.HEADERS, иначе поля
 * разъедутся по колонкам.
 *
 * Телефон пишется с апострофом: строка, начинающаяся с +, для Google
 * Таблиц выглядит как формула и даёт #ERROR! «синтаксическая ошибка
 * в формуле». Апостроф заставляет Таблицы считать ячейку текстом,
 * при этом в самой ячейке его не видно.
 *
 * В последнюю колонку пишется статус: «ок» либо причина, по которой
 * поля оказались пустыми. Так таблица сама показывает, где проблема.
 */
function appendToSheet(data, status) {
  var sheet = getSheet(SETTINGS.SHEET_NAME, SETTINGS.HEADERS);

  var phone = data.phone ? "'" + data.phone : '';

  sheet.appendRow([
    new Date(),
    data.firstName || '',
    data.lastName || '',
    phone,
    data.consent || '',
    data.source || '',
    data.page || '',
    data.submittedAt || '',
    status || 'ок'
  ]);
}


/** Открывает нужную таблицу. */
function getBook() {
  var book = SETTINGS.SHEET_ID
    ? SpreadsheetApp.openById(SETTINGS.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!book) throw new Error('не удалось открыть таблицу, проверьте SHEET_ID');
  return book;
}


/**
 * Возвращает лист, создавая его при необходимости.
 *
 * Если лист уже есть и его первая строка не совпадает с headers,
 * шапка переписывается. Иначе данные легли бы не в те колонки,
 * и вы этого даже не заметите: номер попал бы в столбец «E-mail».
 */
function getSheet(name, headers) {
  var book = getBook();

  var sheet = book.getSheetByName(name);
  if (!sheet) {
    sheet = book.insertSheet(name);
    if (headers) writeHeader(sheet, headers);
    return sheet;
  }

  if (!headers) return sheet;

  if (sheet.getLastRow() === 0 || currentHeader(sheet, headers) !== headers.join('|')) {
    writeHeader(sheet, headers);
  }

  return sheet;
}


/** Первая строка листа, склеенная через |, для сравнения с HEADERS. */
function currentHeader(sheet, headers) {
  return sheet.getRange(1, 1, 1, headers.length).getValues()[0].join('|');
}


/** Пишет и оформляет шапку листа. */
function writeHeader(sheet, headers) {
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  range.setBackground('#F2EBE2');
  range.setFontColor('#5C4033');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}


/** Письмо на почту салона. */
function sendMail(data, problem) {
  if (!SETTINGS.MAIL_TO) return;

  var body = [
    'Новая заявка с сайта «Лаванда»',
    '',
    'Имя:      ' + (data.firstName || '—'),
    'Фамилия:  ' + (data.lastName || '—'),
    'Телефон:  ' + (data.phone || '—'),
    'Согласие: ' + (data.consent === 'yes' ? 'да, отмечено' : 'НЕТ - проверьте'),
    '',
    'Откуда:   ' + (data.source || '—'),
    'Страница: ' + (data.page || '—'),
    'Время:    ' + new Date().toLocaleString('ru-RU')
  ];

  if (problem) body.push('', 'ВНИМАНИЕ: ' + problem);

  MailApp.sendEmail({
    to: SETTINGS.MAIL_TO,
    subject: SETTINGS.MAIL_SUBJECT + (problem ? ' (не распознано)' : ''),
    body: body.join('\n')
  });
}


/** Сообщение в Телеграм. Молча выходит, если токен или чат не заданы. */
function sendTelegram(data) {
  if (!SETTINGS.TELEGRAM_TOKEN || !SETTINGS.TELEGRAM_CHAT) return;

  var text = [
    'Новая заявка — Лаванда',
    (data.firstName || '') + ' ' + (data.lastName || ''),
    data.phone || '',
    'Откуда: ' + (data.source || '—')
  ].join('\n');

  var response = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + SETTINGS.TELEGRAM_TOKEN + '/sendMessage',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: SETTINGS.TELEGRAM_CHAT, text: text }),
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) throw new Error(response.getContentText());
}


/** Подробности - в отдельный лист. Если и он недоступен - в журнал выполнения. */
function writeLog(message) {
  try {
    getSheet(SETTINGS.LOG_SHEET_NAME, ['Время', 'Что произошло']).appendRow([
      new Date(),
      String(message)
    ]);
  } catch (error) {
    Logger.log('не удалось записать в «' + SETTINGS.LOG_SHEET_NAME + '»: ' + error);
    Logger.log(String(message));
  }
}


// ---------------------------------------------------------------------------
// ПРОВЕРКА. Выберите функцию в списке сверху и нажмите «Выполнить».
// ВАЖНО: не нажимайте «Выполнить» при выбранной doPost - Apps Script
// вызовет её без аргумента. Это не заявка с сайта.
// ---------------------------------------------------------------------------

/** Имитирует запрос сайта один в один. Главная проверка. */
function testFromSite() {
  doPost({
    action: 'doPost',
    contentLength: SITE_BODY.length,
    parameter: {},
    postData: { type: 'text/plain', contents: SITE_BODY }
  });
  Logger.log('Готово. Откройте лист «' + SETTINGS.SHEET_NAME + '» - там должна быть строка '
    + 'Анна / Петрова / +7 (999) 123-45-67, в колонке «Согласие» - yes, '
    + 'в последней колонке - «ок». Телефон должен быть виден как +7 (999) 123-45-67, '
    + 'а не как #ERROR!.');
}


/** Проверяет doPost без аргументов - так его вызывали случайно. */
function testEmpty() {
  doPost();
  Logger.log('Готово. Строка в «' + SETTINGS.SHEET_NAME + '» НЕ создана - так и должно быть. '
    + 'Подробности в «' + SETTINGS.LOG_SHEET_NAME + '».');
}


/**
 * Прогоняет скрипт по всем форматам, которые он обязан понимать.
 * Ничего не пишет, только печатает отчёт в журнал выполнения.
 * Первые четыре строки обязаны содержать «Имя: Анна».
 */
function testAllFormats() {
  var cases = [
    ['JSON (шлёт сайт)', { parameter: {}, postData: { type: 'text/plain', contents: SITE_BODY } }],
    ['form-запрос', { parameter: {}, postData: { type: 'application/x-www-form-urlencoded', contents: 'firstName=%D0%90%D0%BD%D0%BD%D0%B0&phone=%2B7%20%28999%29%20123-45-67' } }],
    ['e.parameter', { parameter: { firstName: 'Анна', phone: '+7 (999) 123-45-67' }, postData: null }],
    ['другое имя поля (tel)', { parameter: {}, postData: { type: 'text/plain', contents: '{"name":"Анна","tel":"9161234567"}' } }],
    ['тело пустое', { parameter: {}, postData: { type: 'text/plain', contents: '' } }],
    ['мусор вместо JSON', { parameter: {}, postData: { type: 'text/plain', contents: '<garbage>' } }],
    ['doPost без аргументов', null]
  ];

  cases.forEach(function (item) {
    var r = readData(item[1]);
    Logger.log(item[0] + ' -> Имя: ' + (r.data.firstName || '—')
      + ' | Телефон: ' + (r.data.phone || '—')
      + ' | строку в таблицу: ' + (r.empty ? 'НЕТ' : 'да')
      + ' | Проблема: ' + (r.problem || 'нет'));
  });
}


/**
 * Создаёт оба листа и пишет тестовую заявку.
 * Запустите один раз после заполнения SETTINGS и после разрешений.
 * В таблицу пишет строку Тест / Тестова напрямую, минуя doPost,
 * поэтому в ней не может появиться пометка НЕ РАСПОЗНАНО.
 */
function setup() {
  getSheet(SETTINGS.SHEET_NAME, SETTINGS.HEADERS);
  getSheet(SETTINGS.LOG_SHEET_NAME, ['Время', 'Что произошло']);

  var test = {
    firstName: 'Тест',
    lastName: 'Тестова',
    phone: '+7 (999) 123-45-67',
    consent: 'yes',
    source: 'Проверка из редактора',
    page: 'https://example.com',
    submittedAt: new Date().toISOString()
  };

  appendToSheet(test, 'ок');
  sendMail(test, '');
  sendTelegram(test);

  Logger.log('Готово. Лист «' + SETTINGS.SHEET_NAME + '» создан, строка Тест / Тестова '
    + 'записана со статусом «ок», письмо отправлено.');
  Logger.log('Теперь запустите testFromSite - он имитирует запрос с сайта.');
}


/**
 * Удаляет листы «Заявки» и «Журнал» и создаёт их заново.
 * Нужно, если в таблице уже есть старая шапка.
 * Данные будут потеряны - сохраните их, если нужны.
 */
function rebuildSheets() {
  var book = getBook();

  var names = [SETTINGS.SHEET_NAME, SETTINGS.LOG_SHEET_NAME];
  names.forEach(function (name) {
    var old = book.getSheetByName(name);
    if (old) book.deleteSheet(old);
  });

  getSheet(SETTINGS.SHEET_NAME, SETTINGS.HEADERS);
  getSheet(SETTINGS.LOG_SHEET_NAME, ['Время', 'Что произошло']);

  Logger.log('Листы «' + SETTINGS.SHEET_NAME + '» и «' + SETTINGS.LOG_SHEET_NAME
    + '» пересозданы с правильной шапкой. Теперь запустите testFromSite.');
}
