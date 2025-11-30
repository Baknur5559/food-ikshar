// ===============================================================
//        КОНСТАНТЫ И ГЛОБАЛЬНЫЕ НАСТРОЙКИ
// ===============================================================

const ORDERS_SHEET_NAME = "Заказы";
const BASE_SHEET_NAME = "База";
const CLIENTS_SHEET_NAME = "Клиенты";
const SETTINGS_SHEET_NAME = "Настройки";
const SCHEDULE_SHEET_NAME = "График";


// Индексы колонок в листе "Заказы" (0-based)
const ORDER_NUMBER_COL = 1; // Колонка B
const ORDER_STATUS_COL = 2; // Колонка C
const ORDER_PHONE_COL = 4; // Колонка E
const ORDER_DETAILS_COL = 6; // Колонка G
const ORDER_TOTAL_COL = 7; // Колонка H
const ORDER_LOCATION_COL = 8; // Колонка I
// Новая константа для хранения всех ID сообщений
const ORDER_TELEGRAM_MESSAGES_COL = 12; // Колонка M



// Индексы колонок в листе "База" (0-based)
const BASE_ITEM_NAME_COL = 0;
const BASE_PRICE_COL = 1;
const BASE_IMAGE_URL_COL = 2;
const BASE_PROMO_PRICE_COL = 3;
const BASE_DESCRIPTION_COL = 4;
const BASE_HAS_ADDONS_COL = 5;
const BASE_GROUP_COL = 7;
const BASE_LOCATIONS_START_COL = 8;


// ===============================================================
//         СИСТЕМА КЭШИРОВАНИЯ ДЛЯ УСКОРЕНИЯ ЗАГРУЗКИ
// ===============================================================

/**
 * ВРЕМЕННО ОТКЛЮЧЕН КЭШ ДЛЯ ОТЛАДКИ.
 * Получает данные из кэша или, если их там нет, выполняет функцию и кэширует результат.
 */
function getCachedOrFetch(key, fetchFunction, expirationInSeconds) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(key);
  if (cachedData !== null) {
    Logger.log(`Данные для "${key}" взяты из КЭША.`);
    return JSON.parse(cachedData);
  }

  Logger.log(`КЭШ для "${key}" пуст. Выполняю функцию для получения свежих данных.`);
  const freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), expirationInSeconds);
  return freshData;
}

function testSheetAccess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Заказы");

  if (!sheet) {
    Logger.log("ОШИБКА: Лист 'Заказы' не найден.");
    return;
  }

  const range = sheet.getRange("A1");
  Logger.log("УСПЕХ: Лист найден. Значение в ячейке A1: " + range.getValue());
}


/**
 * Единая функция для получения всех общих данных, которые можно кэшировать.
 * Читает листы "База", "Настройки", "График" только ОДИН раз.
 */
// ЗАМЕНИТЕ СТАРУЮ ФУНКЦИЮ getConsolidatedData НА ЭТУ
function getConsolidatedData() {
  Logger.log("==========================================================");
  Logger.log("--- [НАЧАЛО] Запуск getConsolidatedData: Чтение всех данных ---");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);

  if (!baseSheet || !settingsSheet || !scheduleSheet) {
    Logger.log("[КРИТИЧЕСКАЯ ОШИБКА] Один из листов не найден!");
    return {};
  }

  const baseData = baseSheet.getDataRange().getValues();
  const settingsData = settingsSheet.getDataRange().getValues();
  const scheduleData = scheduleSheet.getDataRange().getValues();

  const workingHoursMap = new Map();
  scheduleData.slice(1).forEach(row => {
      const locationName = row[0];
      if (locationName && row[1] && row[2]) {
          const formatTime = (time) => (time instanceof Date) ? Utilities.formatDate(time, "GMT+6", "HH:mm") : String(time).trim();
          workingHoursMap.set(locationName.trim(), { open: formatTime(row[1]), close: formatTime(row[2]) });
      }
  });

  const locationsMap = new Map();
  baseData.slice(1).forEach((row, index) => {
    for (let j = BASE_LOCATIONS_START_COL; j < row.length; j += 2) {
        const locName = row[j];
        const locAddress = row[j+1];
        if (locName && locAddress && !locationsMap.has(locName.trim())) {
            locationsMap.set(locName.trim(), { name: locName.trim(), address: String(locAddress).trim() });
        }
    }
  });

  const groupedMenuItems = {}, allPromoItemsForSlider = [], addonItems = [], allItemsMapForParsing = new Map();
  baseData.slice(1).forEach(row => {
    const name = row[BASE_ITEM_NAME_COL];
    if (!name) return;
    let price = parseFloat(String(row[BASE_PRICE_COL]).replace(/[^\d.,]/g, '').replace(',', '.'));
    if (isNaN(price)) return;
    let promoPrice = null;
    if (row[BASE_PROMO_PRICE_COL]) {
        const parsedPromo = parseFloat(String(row[BASE_PROMO_PRICE_COL]).replace(/[^\d.,]/g, '').replace(',', '.'));
        if (!isNaN(parsedPromo) && parsedPromo > 0) promoPrice = parsedPromo;
    }
    allItemsMapForParsing.set(String(name).trim().toLowerCase(), { price: (promoPrice || price) });
    const group = String(row[BASE_GROUP_COL] || "Без категории").trim();
    if (group === 'Дополнительно') addonItems.push({ name: name, price: price });
    if (promoPrice) allPromoItemsForSlider.push({ name, price, promoPrice, imageUrl: row[BASE_IMAGE_URL_COL] || "", description: row[BASE_DESCRIPTION_COL] || "" });
    const fullItemData = { name, price, promoPrice, imageUrl: row[BASE_IMAGE_URL_COL] || "", description: row[BASE_DESCRIPTION_COL] || "", group, hasAddons: row[BASE_HAS_ADDONS_COL] === true };
    if (!groupedMenuItems[group]) groupedMenuItems[group] = [];
    groupedMenuItems[group].push(fullItemData);
  });

  const deliveryTimes = {}, appSettings = { paymentMethods: [], deliveryTypes: {} };
  const knownDeliveryTypes = ["Зал", "Доставка", "На вынос"];
  settingsData.slice(1).forEach(row => {
    if (row[4]) deliveryTimes[String(row[4]).trim()] = { delivery: parseFloat(String(row[5] || '0').replace(',', '.')) || 0, pickup: parseFloat(String(row[6] || '0').replace(',', '.')) || 0 };
    const name = row[12];
    if (name && row[13] === true) {
        if (knownDeliveryTypes.includes(name)) appSettings.deliveryTypes[name] = true;
        else appSettings.paymentMethods.push({ name: name.trim(), locations: String(row[16] || '').split(',').map(s => s.trim()).filter(Boolean) });
    }
  });

  // *** НАЧАЛО ИСПРАВЛЕНИЙ В ЛОГИКЕ ВРЕМЕНИ ***
  const nowString = Utilities.formatDate(new Date(), "GMT+6", "HH:mm");
  const finalLocations = Array.from(locationsMap.values()).map(loc => {
      const schedule = workingHoursMap.get(loc.name);
      let status = "Неизвестно", statusText = "Нет данных о графике", workingHoursText = "";
      if (schedule) {
          workingHoursText = `с ${schedule.open} до ${schedule.close}`;
          let isOpen = false;
          if (schedule.open === '00:00' && schedule.close === '00:00') {
              isOpen = true;
              statusText = "Круглосуточно";
          } else if (schedule.close < schedule.open) { // Работа через ночь
              if (nowString >= schedule.open || nowString < schedule.close) isOpen = true;
          } else { // Обычный день
              if (nowString >= schedule.open && nowString < schedule.close) isOpen = true;
          }
          if (isOpen) {
              status = "Открыто";
              if (statusText !== "Круглосуточно") statusText = `Закроется в ${schedule.close}`;
          } else {
              status = "Закрыто";
              statusText = (nowString < schedule.open) ? `Откроется в ${schedule.open}` : `Открыто до ${schedule.close}`;
          }
      }
      return { ...loc, status, statusText, workingHoursText };
  });
  // *** КОНЕЦ ИСПРАВЛЕНИЙ ***

  const globalPromoItems = allPromoItemsForSlider.sort(() => 0.5 - Math.random()).slice(0, 5);

  Logger.log("[РЕЗУЛЬТАТ] Итоговый массив finalLocations: " + JSON.stringify(finalLocations));
  Logger.log("--- [КОНЕЦ] Завершение getConsolidatedData ---");
  Logger.log("==========================================================");

  return { locations: finalLocations, deliveryTimes, addonItems, settings: appSettings, globalPromoItems, groupedMenuItems, allPromoItems: allPromoItemsForSlider, allItemsMapForParsing: Object.fromEntries(allItemsMapForParsing) };
}


// ===============================================================
//        ИНТЕРФЕЙС В GOOGLE SHEETS (САЙДБАР)
// ===============================================================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Управление заказами')
    .addItem('Изменить состав заказа', 'showOrderEditorSidebar')
    .addToUi();
}


function showOrderEditorSidebar() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
    const range = sheet.getActiveRange();
    
    if (range.getRow() === 1 || range.getNumRows() > 1) {
      SpreadsheetApp.getUi().alert('Пожалуйста, выберите одну ячейку в строке того заказа, который хотите изменить.');
      return;
    }
    
    const row = range.getRow();
    const orderDataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const orderInfo = {
      row: row,
      number: orderDataRow[ORDER_NUMBER_COL],
      itemsText: orderDataRow[ORDER_DETAILS_COL]
    };
    
    const template = HtmlService.createTemplateFromFile('EditorSidebar');
    template.orderInfo = orderInfo;
    
    const html = template.evaluate()
        .setTitle('Редактор заказа #' + orderInfo.number)
        .setWidth(850);  
        
    SpreadsheetApp.getUi().showSidebar(html);


  } catch (e) {
    Logger.log("КРИТИЧЕСКАЯ ОШИБКА в showOrderEditorSidebar: " + e.toString());
    SpreadsheetApp.getUi().alert("Произошла критическая ошибка при открытии панели: " + e.message);
  }
}




// ===============================================================
//        ОСНОВНЫЕ ФУНКЦИИ WEB APP (doGet, doPost)
// ===============================================================


function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index') // Убедись, что здесь НЕ .evaluate()
    .setTitle('SushiSan47: Система Заказов')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  // Логируем весь входящий JSON-объект, чтобы увидеть его структуру
  Logger.log(JSON.stringify(e));

  // Получаем JSON-строку из объекта `e`
  const postData = e.postData.contents;
  const data = JSON.parse(postData);

  // Проверяем, есть ли в данных информация о сообщении
  if (data.message) {
    const chatId = data.message.chat.id;
    const messageText = data.message.text;
    Logger.log("Найдено сообщение. Chat ID: " + chatId + ", Текст: " + messageText);
  }

  // Возвращаем "OK", чтобы Telegram знал, что мы получили запрос
  return ContentService.createTextOutput("OK");
}


function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}




// ===============================================================
//        ФУНКЦИИ ДЛЯ ПОЛУЧЕНИЯ ДАННЫХ ФРОНТЕНДОМ
// ===============================================================

/**
 * Загружает все данные для сессии пользователя. ИСПОЛЬЗУЕТ КЭШ.
 */

function getUserSessionData(phoneNumber) {
  try {
    const consolidatedData = getCachedOrFetch('consolidatedData', getConsolidatedData, 1); // Кэш на 1 секунду для тестов
    const contactInfo = getContactInfo();
    const allClientOrders = getClientOrders(phoneNumber, consolidatedData.allItemsMapForParsing);
    const activeOrders = allClientOrders.filter(o => o.status === 'Новый' || o.status === 'Подтвержден');

    let clientData = null;
    const clientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET_NAME);
    const requestedPhone = normalizePhoneRU_GS(phoneNumber);

    if (clientsSheet && requestedPhone) {
      // Быстрый поиск клиента через TextFinder
      const phoneColumn = clientsSheet.getRange("B:B");
      const textFinder = phoneColumn.createTextFinder(requestedPhone).matchEntireCell(true);
      const foundCell = textFinder.findNext();

      if (foundCell) {
        const row = foundCell.getRow();
        const clientRowData = clientsSheet.getRange(row, 1, 1, 3).getValues()[0];
        clientData = { name: clientRowData[0], phone: clientRowData[1], address: clientRowData[2] };
      }
    }

    return {
      activeOrders: activeOrders,
      locations: consolidatedData.locations,
      settings: consolidatedData.settings,
      deliveryTimes: consolidatedData.deliveryTimes,
      clientData: clientData,
      contactInfo: contactInfo
    };
  } catch (e) {
    Logger.log("Критическая ошибка в getUserSessionData: " + e.stack);
    return { activeOrders: [], locations: [], settings: {}, deliveryTimes: {}, clientData: null, contactInfo: [] };
  }
}

/**
 * Получает меню для выбранной точки продаж. ИСПОЛЬЗУЕТ КЭШ.
 */
function getMenuItems(selectedLocationName) {
  // Получаем все данные из кэша
  const consolidatedData = getCachedOrFetch('consolidatedData', getConsolidatedData, 1);
  const { groupedMenuItems, allPromoItems } = consolidatedData;

  const locationSpecificMenu = {};
  
  // Фильтруем меню, оставляя только доступные для данной точки
  for (const group in groupedMenuItems) {
      const availableItems = groupedMenuItems[group].filter(item => {
         // Предполагаем, что если у товара нет привязки к точке, он доступен везде
         // Это упрощение, логику доступности нужно будет адаптировать под вашу структуру в "Базе"
         return true; // Здесь нужна ваша логика проверки доступности по колонке "Точка продаж"
      });
      if(availableItems.length > 0) {
        locationSpecificMenu[group] = availableItems;
      }
  }

  const locationSpecificPromos = allPromoItems.filter(item => {
      // Та же логика доступности для акций
      return true;
  });

  return { menuItems: locationSpecificMenu, promoItems: locationSpecificPromos };
}

/**
 * Получает заказы клиента. Теперь принимает карту цен, чтобы не читать лист "База".
 * ОБНОВЛЕННАЯ ВЕРСИЯ: Если карта цен не передана, берет ее из кэша.
 */
function getClientOrders(phoneNumber, allItemsMapObject) {
    try {
        let allItemsMap;

        if (allItemsMapObject) {
            allItemsMap = new Map(Object.entries(allItemsMapObject));
        } else {
            Logger.log("Карта товаров не была предоставлена в getClientOrders. Запрашиваю данные из кэша.");
            const consolidatedData = getCachedOrFetch('consolidatedData', getConsolidatedData, 1);
            allItemsMap = new Map(Object.entries(consolidatedData.allItemsMapForParsing));
        }
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const ordersSheet = ss.getSheetByName(ORDERS_SHEET_NAME);
        if (!ordersSheet) return [];
        
        const allOrdersData = ordersSheet.getDataRange().getValues();
        const clientOrders = [];
        const requestedPhoneNumber = normalizePhoneRU_GS(phoneNumber);

        for (let i = allOrdersData.length - 1; i > 0; i--) { 
            const row = allOrdersData[i];
            const sheetPhoneNumber = normalizePhoneRU_GS(row[ORDER_PHONE_COL]);

            if (sheetPhoneNumber === requestedPhoneNumber) {
                const orderDate = row[0];
                const orderDetailsText = row[ORDER_DETAILS_COL];
                const parsedItems = parseOrderDetailsString(orderDetailsText, allItemsMap);

                clientOrders.push({
                    number: row[ORDER_NUMBER_COL],
                    status: row[ORDER_STATUS_COL],
                    selectedLocation: row[ORDER_LOCATION_COL],
                    total: Number(row[ORDER_TOTAL_COL]) || 0,
                    deliveryFee: Number(row[18]) || 0,
                    date: orderDate instanceof Date ? Utilities.formatDate(orderDate, "GMT+6", "dd.MM.yyyy в HH:mm") : String(orderDate),
                    items: parsedItems
                });
            }
        }
        return clientOrders;
    } catch (e) {
        Logger.log("Критическая ошибка в getClientOrders: " + e.stack);
        return [];  
    }
}


/**
 * Получает меню для выбранной точки продаж.
 */
function getMenuItems(selectedLocationName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) return { menuItems: {}, promoItems: [] };
  const data = baseSheet.getDataRange().getValues();
  const lastCol = baseSheet.getLastColumn();
  const groupedMenuItems = {};
  const promoItems = [];


  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemName = row[BASE_ITEM_NAME_COL];
    const itemPriceRaw = row[BASE_PRICE_COL];
    const itemImageUrl = row[BASE_IMAGE_URL_COL];
    const itemHasAddons = row[BASE_HAS_ADDONS_COL] === true;
    const itemGroup = row[BASE_GROUP_COL] ? String(row[BASE_GROUP_COL]).trim() : "Без категории";
    const itemPromoPriceRaw = row[BASE_PROMO_PRICE_COL];
    const itemDescription = row[BASE_DESCRIPTION_COL] || "";
    let itemPrice = parseFloat(String(itemPriceRaw).replace(/[^\d.,]/g, '').replace(',', '.'));
    if (isNaN(itemPrice)) continue;
    let itemPromoPrice = null;
    if (itemPromoPriceRaw) {
      const parsedPromo = parseFloat(String(itemPromoPriceRaw).replace(/[^\d.,]/g, '').replace(',', '.'));
      if (!isNaN(parsedPromo) && parsedPromo > 0) { itemPromoPrice = parsedPromo; }
    }
    let isAvailable = false;
    for (let j = BASE_LOCATIONS_START_COL; j < lastCol; j += 2) {
      const salesPoints = row[j];
      if (salesPoints && String(salesPoints).split(',').map(s => s.trim()).includes(selectedLocationName)) {
        isAvailable = true;
        break;
      }
    }
    if (itemName && itemImageUrl && isAvailable) {
      const fullItemData = {
        name: itemName,
        price: itemPrice,
        promoPrice: itemPromoPrice,
        imageUrl: itemImageUrl,
        description: itemDescription,
        group: itemGroup,
        hasAddons: itemHasAddons
      };


      if (!groupedMenuItems[itemGroup]) { groupedMenuItems[itemGroup] = []; }
      groupedMenuItems[itemGroup].push(fullItemData);


      if (itemPromoPrice) {
        promoItems.push(fullItemData);
      }
    }
  }
  return { menuItems: groupedMenuItems, promoItems: promoItems }; 
}


// ===============================================================
//        ЛОГИКА ОБРАБОТКИ И СОХРАНЕНИЯ ЗАКАЗА
// ===============================================================




/**
 * Создает совершенно новый заказ.
 * ОБНОВЛЕНО: Отправляет уведомления всем получателям из настроек + на E-MAIL.
 */
function createNewOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName(ORDERS_SHEET_NAME);
  const newOrderNumber = "ORD-" + new Date().getTime().toString().slice(-6) + Math.floor(Math.random() * 100);

  let deliveryFee = 0;
  let finalTotalAmount = orderData.totalAmount;
  let yandexMapsLink = "";

  if (orderData.deliveryType === 'Доставка' && orderData.deliveryAddress) {
    const apiKey = getYandexApiKey();
    const startCoords = getCoordinatesForAddress(orderData.startPointAddress, apiKey);
    const endCoords = getCoordinatesForAddress(orderData.deliveryAddress, apiKey);
    if (startCoords && endCoords) {
      const routeDetails = getRouteDetails(startCoords, endCoords, apiKey);
      if (routeDetails && routeDetails.distance) {
        const distanceKm = routeDetails.distance;
        const tiers = getDeliveryFeeTiers();
        if (tiers.length > 0 && distanceKm > 0) {
          let lastTier = tiers[tiers.length - 1];
          deliveryFee = lastTier.fee;
          for (const tier of tiers) {
            if (distanceKm <= tier.km) {
              deliveryFee = tier.fee;
              break;
            }
          }
        }
      }
    }
    finalTotalAmount += deliveryFee;
    const startAddr = orderData.startPointAddress || "ул. Исы Ахунбаева, 283, Бишкек";
    const endAddr = orderData.deliveryAddress;
    yandexMapsLink = `https://yandex.ru/maps/?rtext=${encodeURIComponent(startAddr)}~${encodeURIComponent(endAddr)}&rtt=auto`;
  }

  const orderDetailsForSheet = formatOrderDetailsForSheet(orderData.cartItems);

  // 1. Создаем заказ в таблице
  ordersSheet.appendRow([
    new Date(), newOrderNumber, "Новый", orderData.clientName, orderData.clientPhone,
    orderData.deliveryAddress, orderDetailsForSheet, finalTotalAmount,
    orderData.selectedLocation, orderData.comments, "", yandexMapsLink, "", // Оставляем M и N пустыми, они заполнятся позже
    orderData.paymentMethod, orderData.deliveryType, orderData.selectedTime,
    orderData.changeFrom, deliveryFee
  ]);

  updateClientData(orderData.clientName, orderData.clientPhone, orderData.deliveryAddress, new Date(), newOrderNumber);

  // 2. Собираем ПОЛНЫЙ объект данных для уведомлений
  const dataForNotifications = {
    orderNumber: newOrderNumber,
    status: "Новый",
    clientName: orderData.clientName,
    clientPhone: orderData.clientPhone,
    deliveryAddress: orderData.deliveryAddress || "Самовывоз",
    cartItems: orderData.cartItems,
    orderDetailsText: orderDetailsForSheet,
    totalAmount: finalTotalAmount,
    subtotalAmount: orderData.totalAmount,
    deliveryFee: deliveryFee,
    selectedLocation: orderData.selectedLocation,
    comments: orderData.comments || "Нет",
    yandexMapsLink: yandexMapsLink,
    paymentMethod: orderData.paymentMethod,
    deliveryType: orderData.deliveryType,
    selectedTime: orderData.selectedTime,
    changeFrom: orderData.changeFrom || ""
  };

  // 3. Вызываем наши новые функции для отправки уведомлений
  sendNewOrderNotification(dataForNotifications); // Отправка в Telegram

  const emailTitle = "Новый заказ";
  const emailBody = generateHtmlEmailBody(dataForNotifications, emailTitle);
  sendEmailNotification(`${emailTitle} #${newOrderNumber}`, emailBody); // Отправка на E-mail

  return { status: "success", orderNumber: newOrderNumber };
}


// ===============================================================
//      ТРИГГЕРЫ И ОБРАБОТЧИКИ СОБЫТИЙ
// ===============================================================

/**
 * Срабатывает при РУЧНОМ редактировании таблицы.
 * Выполняет проверки безопасности и вызывает функцию для обновления сообщения в Telegram.
 * @param {object} e Объект события.
 */
function handleEdit(e) {
  // 1. Проверяем, что событие редактирования корректно
  if (!e || !e.range) {
    Logger.log("Выполнение handleEdit() прервано: отсутствует объект события. Возможно, скрипт был запущен вручную.");
    return;
  }

  const sheet = e.range.getSheet();
  // 2. Убеждаемся, что редактирование происходит на листе "Заказы"
  if (sheet.getName() !== ORDERS_SHEET_NAME) {
    return;
  }

  const editedColumn = e.range.getColumn();
  const editedRow = e.range.getRow();

  // 3. Определяем, за какими колонками мы следим (индексы колонок для getColumn(), 1-based)
  const columnsToWatch = {
    [ORDER_STATUS_COL + 1]: "статус",      // C
    [ORDER_DETAILS_COL + 1]: "состав",     // G
    [ORDER_TOTAL_COL + 1]: "сумма",        // H
    [ORDER_PHONE_COL + 1]: "телефон",      // E
    [ORDER_LOCATION_COL + 1]: "точка продаж", // I
    10: "комментарий",                     // J
    14: "оплата",                         // N
    15: "тип заказа",                     // O
    16: "время",                          // P
    17: "сдача",                          // Q
    18: "доставка"                        // R
  };

  // 4. Если отредактирована неважная колонка или заголовок - выходим
  if (!columnsToWatch[editedColumn] || editedRow === 1) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const orderDataRow = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const orderNumber = orderDataRow[ORDER_NUMBER_COL];
  const currentStatus = orderDataRow[ORDER_STATUS_COL];
  const allRoles = getRolesAndPins();
  let validatedRole = null;

  try {
    // БЛОК ПРОВЕРОК БЕЗОПАСНОСТИ (с запросом ПИН-кода)
    if (editedColumn === ORDER_DETAILS_COL + 1 && currentStatus !== "Новый") {
      validatedRole = validatePinForRoles(['Менеджер', 'Руководитель'], allRoles);
      if (!validatedRole) { e.range.setValue(e.oldValue); return; }
      logChange(validatedRole, orderNumber, "Изменение состава", e.oldValue, e.value);
    } else if (editedColumn === 18) { // Проверка для колонки "Доставка" (R)
      validatedRole = validatePinForRoles(['Кассир', 'Менеджер', 'Руководитель'], allRoles);
      if (!validatedRole) { e.range.setValue(e.oldValue); return; }
      logChange(validatedRole, orderNumber, "Изменение суммы доставки", e.oldValue, e.value);
    } else if (editedColumn === ORDER_STATUS_COL + 1) {
      const newStatus = e.value;
      const statusFlow = { "Новый": 1, "Подтвержден": 2, "Отправлен": 3, "Доставлен": 4, "Отказ": 0 };
      const isBackwardMove = (statusFlow[newStatus] < statusFlow[currentStatus]) && newStatus !== 'Отказ';
      const isDangerousChange = (currentStatus === 'Отказ' && newStatus === 'Подтвержден') || (currentStatus === 'Доставлен' && newStatus === 'Отказ');
      if (isBackwardMove || isDangerousChange) {
        validatedRole = validatePinForRoles(['Менеджер', 'Руководитель'], allRoles);
        if (!validatedRole) { e.range.setValue(e.oldValue); return; }
        logChange(validatedRole, orderNumber, "Критическое изменение статуса", currentStatus, newStatus);
      }
    }
    
    // После всех проверок вызываем функцию для обновления Telegram
    const updatedOrderDataRow = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cartItems = parseOrderDetailsString(updatedOrderDataRow[ORDER_DETAILS_COL]);
    
    // =======================================================
    //     ИСПРАВЛЕННЫЙ БЛОК С ПРАВИЛЬНЫМИ ИНДЕКСАМИ
    // =======================================================
    const orderData = {
        orderNumber: updatedOrderDataRow[ORDER_NUMBER_COL],    // Индекс 1 (Колонка B)
        status: updatedOrderDataRow[ORDER_STATUS_COL],         // Индекс 2 (Колонка C)
        clientName: updatedOrderDataRow[3],                    // Индекс 3 (Колонка D)
        clientPhone: updatedOrderDataRow[4],                   // Индекс 4 (Колонка E)
        deliveryAddress: updatedOrderDataRow[5] || "Самовывоз",// Индекс 5 (Колонка F)
        orderDetailsText: updatedOrderDataRow[ORDER_DETAILS_COL], // Индекс 6 (Колонка G)
        cartItems: cartItems,
        totalAmount: Number(updatedOrderDataRow[ORDER_TOTAL_COL]), // Индекс 7 (Колонка H)
        selectedLocation: updatedOrderDataRow[ORDER_LOCATION_COL], // Индекс 8 (Колонка I)
        comments: updatedOrderDataRow[9] || "Нет",             // Индекс 9 (Колонка J)
        yandexMapsLink: updatedOrderDataRow[11],                 // Индекс 11 (Колонка L)
        paymentMethod: updatedOrderDataRow[13],                  // Индекс 13 (Колонка N)
        deliveryType: updatedOrderDataRow[14],                   // Индекс 14 (Колонка O)
        selectedTime: updatedOrderDataRow[15],                   // Индекс 15 (Колонка P)
        changeFrom: updatedOrderDataRow[16] || "",               // Индекс 16 (Колонка Q)
        deliveryFee: Number(updatedOrderDataRow[17]) || 0      // Индекс 17 (Колонка R)
    };
    // =======================================================

    updateTelegramMessageForOrderFromData(orderData, columnsToWatch[editedColumn]);
    
    // Безопасный вызов ui.toast
    if (ui && ui.toast) {
      ui.toast(`Заказ #${orderNumber} в Telegram обновлен!`, '✅ Готово', 5);
    }

  } catch (err) {
    Logger.log("Критическая ошибка в handleEdit: " + err.message + " | Строка: " + err.lineNumber);
    // Если UI доступен, покажем alert
    const ui = SpreadsheetApp.getUi();
    if (ui && ui.alert) {
      ui.alert("Произошла критическая ошибка: " + err.message);
    }
  }
}
// ===============================================================
//        ВСПОМОГАТЕЛЬНЫЕ И УТИЛИТАРНЫЕ ФУНКЦИИ
// ===============================================================


// --- Функции для работы с данными ---


function updateClientData(clientName, clientPhone, deliveryAddress, orderDate, orderNumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientsSheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
    if (!clientsSheet) return;
    const clientData = clientsSheet.getDataRange().getValues();
    let clientFound = false;
    for (let i = 1; i < clientData.length; i++) {
        if (clientData[i][1] === clientPhone) {
            clientsSheet.getRange(i + 1, 5).setValue(Number(clientData[i][4]) + 1);
            clientsSheet.getRange(i + 1, 6).setValue(orderNumber);
            clientFound = true;
            break;
        }
    }
    if (!clientFound) {
        clientsSheet.appendRow([clientName, clientPhone, deliveryAddress, orderDate, 1, orderNumber]);
    }
}

/**
 * Нормализует российский номер телефона на сервере.
 * @param {string} phone - Номер телефона.
 * @returns {string} Нормализованный номер в формате 7XXXXXXXXXX.
 */
function normalizePhoneRU_GS(phone) {
  if (!phone) return '';
  // Удаляем все символы, кроме цифр
  var cleaned = String(phone).replace(/\D/g, '');
  
  // Если номер начинается с 8, заменяем на 7
  if (cleaned.startsWith('8')) {
    cleaned = '7' + cleaned.substring(1);
  } 
  // Если номер начинается с 9 (без кода страны) и его длина 10 цифр
  else if (cleaned.length === 10 && cleaned.startsWith('9')) {
    cleaned = '7' + cleaned;
  }
  return cleaned;
}


function getSalesLocations() {
  // --- НАЧАЛО БЛОКА ЛОГИРОВАНИЯ ---
  Logger.log("--- Запуск функции getSalesLocations ---"); 
  // --- КОНЕЦ БЛОКА ЛОГИРОВАНИЯ ---

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) {
    // --- НАЧАЛО БЛОКА ЛОГИРОВАНИЯ ---
    Logger.log("ОШИБКА: Лист 'База' не найден. Возвращаем пустой массив.");
    // --- КОНЕЦ БЛОКА ЛОГИРОВАНИЯ ---
    return [];
  }

  const workingHoursMap = getWorkingHours();
  // --- НАЧАЛО БЛОКА ЛОГИРОВАНИЯ ---
  // Преобразуем Map в объект для красивого вывода в лог
  Logger.log("Загружены рабочие часы для " + workingHoursMap.size + " точек.");
  Logger.log("Данные по часам: " + JSON.stringify(Array.from(workingHoursMap.entries())));
  // --- КОНЕЦ БЛОКА ЛОГИРОВАНИЯ ---

  const locations = new Map();
  const data = baseSheet.getDataRange().getValues();
  const lastCol = baseSheet.getLastColumn();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    for (let j = BASE_LOCATIONS_START_COL; j < lastCol; j += 2) {
      const name = row[j];
      const address = row[j + 1];

      if (name && address && !locations.has(name.trim())) {
        const locationName = name.trim();
        const schedule = workingHoursMap.get(locationName);
        let status = "Неизвестно", statusText = "Нет данных о графике", workingHoursText = "";

        if (schedule) {
          if (!(schedule.open === '00:00' && schedule.close === '00:00')) { workingHoursText = `с ${schedule.open} до ${schedule.close}`; }
          const [openHour, openMin] = schedule.open.split(':').map(Number);
          const [closeHour, closeMin] = schedule.close.split(':').map(Number);
          if (schedule.open === '00:00' && schedule.close === '00:00') { status = "Открыто"; statusText = "Круглосуточно"; }
          else {
            const openTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), openHour, openMin);
            const closeTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), closeHour, closeMin);
            if (closeTime < openTime) { if (now < closeTime) { openTime.setDate(openTime.getDate() - 1); } else { closeTime.setDate(closeTime.getDate() + 1); } }
            if (now >= openTime && now < closeTime) { status = "Открыто"; statusText = `Закроется в ${schedule.close}`; }
            else { status = "Закрыто"; if (now < openTime) { statusText = `Откроется в ${schedule.open}`; } else { statusText = `Открыто до ${schedule.close}`; } }
          }
        }
        
        // --- НАЧАЛО БЛОКА ЛОГИРОВАНИЯ ---
        Logger.log("Найдена и обработана точка: '" + locationName + "' со статусом: '" + status + "'");
        // --- КОНЕЦ БЛОКА ЛОГИРОВАНИЯ ---

        locations.set(locationName, { name: locationName, address: String(address).trim(), status: status, statusText: statusText, workingHoursText: workingHoursText });
      }
    }
  }

  const finalLocations = Array.from(locations.values());
  // --- НАЧАЛО БЛОКА ЛОГИРОВАНИЯ ---
  Logger.log("--- Итоговый результат для отправки на фронтенд ---");
  // JSON.stringify с форматированием для удобного чтения
  Logger.log(JSON.stringify(finalLocations, null, 2)); 
  // --- КОНЕЦ БЛОКА ЛОГИРОВАНИЯ ---

  return finalLocations;
}


function getWorkingHours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const workingHours = new Map();
  if (!scheduleSheet) return workingHours;
  const data = scheduleSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const locationName = row[0];
    const openTime = row[1];
    const closeTime = row[2];
    if (locationName && openTime && closeTime) {
      const formatTime = (time) => { if (time instanceof Date) { return Utilities.formatDate(time, "GMT+6", "HH:mm"); } return String(time).trim(); };
      workingHours.set(locationName.trim(), { open: formatTime(openTime), close: formatTime(closeTime) });
    }
  }
  return workingHours;
}


// --- Функции для работы с Telegram ---


function sendTelegramMessage(chatId, text, inlineKeyboard, botToken) {
  // --- НАЧАЛО ИСПРАВЛЕНИЯ ---
  // Проверяем, что Chat ID существует и не пустой
  if (!chatId || String(chatId).trim() === '') {
    Logger.log("ПРЕРВАНО: Попытка отправки сообщения без Chat ID. Текст: " + text);
    return null; // Прерываем выполнение функции
  }
  // --- КОНЕЦ ИСПРАВЛЕНИЯ ---

  if (!botToken) { const config = getTelegramConfig("По умолчанию"); botToken = config.token; }
  if (!botToken) { Logger.log("Ошибка: Telegram Bot Token не найден."); return null; }
  const TELEGRAM_API_URL = `https://api.telegram.org/bot${botToken}/sendMessage`;
  const payload = { chat_id: String(chatId), text: text, parse_mode: "MarkdownV2" };
  if (inlineKeyboard) { payload.reply_markup = JSON.stringify(inlineKeyboard); }
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(TELEGRAM_API_URL, options);
    const responseJson = JSON.parse(response.getContentText());
    if (response.getResponseCode() === 200 && responseJson.ok) { return responseJson.result.message_id; }
    else { 
        // Добавим более подробный лог
        Logger.log(`Ошибка отправки сообщения в Telegram для чата ${chatId}: ${response.getContentText()}`); 
        return null; 
    }
  } catch (e) { Logger.log(`Критическая ошибка при вызове Telegram API для чата ${chatId}: ${e.message}`); return null; }
}

/**
 * Отправляет отдельное короткое уведомление об изменении в заказе.
 * @param {string} orderNumber Номер заказа.
 * @param {string} locationName Название точки продаж.
 * @param {string} reason Причина изменения.
 * @param {string} clientName Имя клиента.
 * @param {string} clientPhone Телефон клиента.
 */
function sendUpdateNotification(orderNumber, locationName, reason, clientName, clientPhone) {
  try {
    const config = getTelegramConfig(locationName);
    const notificationList = getNotificationChatIds();
    const allRecipients = new Set(notificationList);
    if (config.chatId) {
      allRecipients.add(String(config.chatId));
    }

    if (!config.token || allRecipients.size === 0) {
      Logger.log(`Не удалось отправить уведомление для заказа ${orderNumber}: не найдены адреса чатов или токен.`);
      return;
    }

    const message = `🔔 *Изменение в заказе* \`\\#${escapeMarkdown(orderNumber)}\`` +  
                     `\n*Клиент:* ${escapeMarkdown(clientName)} \\(${escapeMarkdown(clientPhone)}\\)` +
                     `\n*Причина:* ${escapeMarkdown(reason)}`;

    // Отправляем уведомление всем получателям
    allRecipients.forEach(chatId => {
      sendTelegramMessage(chatId, message, null, config.token);
    });

  } catch (e) {
    Logger.log(`Ошибка при отправке уведомления для заказа ${orderNumber}: ${e.message}`);
  }
}

function editTelegramMessage(chatId, messageId, newText, botToken) {
  if (!botToken || !chatId || !messageId || !newText) { return; }
  const TELEGRAM_API_URL = `https://api.telegram.org/bot${botToken}/editMessageText`;
  const payload = { chat_id: String(chatId), message_id: Number(messageId), text: newText, parse_mode: "MarkdownV2" };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  try { UrlFetchApp.fetch(TELEGRAM_API_URL, options); }
  catch (e) { Logger.log("Критическая ошибка при вызове Telegram API (editMessageText): " + e.message); }
}


// --- Функции для работы с API Яндекса ---


function getYandexApiKey() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    return settingsSheet.getRange("D3").getValue(); // Предполагаем, что ключ в D2
  } catch (e) {
    Logger.log("Ошибка при чтении ключа API: " + e.message);
    return null;
  }
}


function getRouteDetails(startCoords, endCoords, apiKey) {
  if (!apiKey || !startCoords || !endCoords) return null;
  const [startLon, startLat] = startCoords.split(',');
  const [endLon, endLat] = endCoords.split(',');
  const url = `https://api.routing.yandex.net/v2/route?apikey=${apiKey}&waypoints=${startLat},${startLon}|${endLat},${endLon}&mode=driving`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.routes && json.routes.length > 0) {
        return { distance: json.routes[0].summary.distance / 1000 };
      }
    }
  } catch (e) { Logger.log(`Критическая ошибка при вызове API Маршрутизатора: ${e.message}`); }
  return null;
}


function getCoordinatesForAddress(address, apiKey) {
  if (!apiKey || !address) return null;
  
  // Убрали строку, которая добавляла "Бишкек". Теперь используется адрес как есть.
  const fullAddress = address;
  
  const url = `https://geocode-maps.yandex.ru/1.x/?apikey=${apiKey}&format=json&geocode=${encodeURIComponent(fullAddress)}&lang=ru_RU`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const geoObjects = json.response.GeoObjectCollection.featureMember;
      if (geoObjects.length > 0) {
        const point = geoObjects[0].GeoObject.Point.pos;
        const [lon, lat] = point.split(' ');
        return `${lon},${lat}`;
      }
    }
  } catch (e) { Logger.log(`Критическая ошибка при вызове API Геокодера для адреса "${address}": ${e.message}`); }
  return null;
}

// --- Прочие вспомогательные функции ---


const escapeMarkdown = (text) => {
    if (text === null || text === undefined) return '';
    return String(text).replace(/[_*[\]()~`>#+\-=|{}.!]/g, '\\$&');
};


function findRowByOrderNumber(sheet, orderNumber) {
    const data = sheet.getRange("B:B").getValues();
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] == orderNumber) {
            return i + 1;
        }
    }
    return null;
}


/**
 * Форматирует детали заказа для записи в Google Sheet.
 * Функция группирует одинаковые товары и их добавки,
 * чтобы запись в таблице была чистой и понятной.
 * @param {Array} cartItems - Массив объектов товаров в корзине.
 * @returns {string} Строка с деталями заказа, сгруппированными по позициям.
 */
function formatOrderDetailsForSheet(cartItems) {
    // Используем Map для группировки товаров по уникальному ключу
    const groupedItems = new Map();

    cartItems.forEach(item => {
        // Создаем уникальный ключ для каждого товара с учётом его добавок
        const addonsKey = item.addons && item.addons.length > 0 ? 
            JSON.stringify(item.addons.map(a => `${a.name}x${a.quantity}`)) : 
            '';
        const key = item.name + addonsKey;

        if (groupedItems.has(key)) {
            // Если товар уже есть, просто увеличиваем его количество
            const existingItem = groupedItems.get(key);
            existingItem.quantity += item.quantity;
        } else {
            // Если товара нет, добавляем его в Map
            groupedItems.set(key, {
                name: item.name,
                quantity: item.quantity,
                addons: item.addons || []
            });
        }
    });

    // Формируем финальную строку для таблицы
    return Array.from(groupedItems.values()).map(item => {
        let details = `${item.name} (*${item.quantity}*)`;
        
        // Добавляем информацию о добавках, если они есть
        if (item.addons && item.addons.length > 0) {
            const addonsText = item.addons.map(addon => 
                `${addon.name} x${addon.quantity}`
            ).join(', ');
            details += ` (Допы: ${addonsText})`;
        }
        
        return details;
    }).join('; ');
}


/**
 * Разбирает строку деталей заказа, теперь использует карту цен.
 */
function parseOrderDetailsString(orderDetailsText, allItemsMap) {
    if (!orderDetailsText) return [];
    const itemsStrings = String(orderDetailsText).split(';').map(s => s.trim()).filter(Boolean);
    const parsedItems = [];
    itemsStrings.forEach(itemString => {
        const addonMatch = itemString.match(/(.+?) \((\d+)\) \(Допы: (.*)\)/);
        const simpleMatch = itemString.match(/(.+?) \((\d+)\)$/);
        let name, quantity, itemFound = false;
        const addons = [];

        if (addonMatch) {
            name = addonMatch[1].trim();
            quantity = parseInt(addonMatch[2], 10);
            const addonsText = addonMatch[3].trim();
            addonsText.split(',').forEach(addonStr => {
                const parts = addonStr.trim().split(' x');
                const addonName = parts[0].trim();
                const addonQty = parseInt(parts[1], 10) || 1;
                const addonData = allItemsMap ? allItemsMap.get(addonName.toLowerCase()) : null;
                addons.push({ name: addonName, quantity: addonQty, price: addonData ? addonData.price : 0 });
            });
            itemFound = true;
        } else if (simpleMatch) {
            name = simpleMatch[1].trim();
            quantity = parseInt(simpleMatch[2], 10);
            itemFound = true;
        }

        if (itemFound) {
            const itemData = allItemsMap ? allItemsMap.get(name.toLowerCase()) : null;
            parsedItems.push({ name, quantity, price: itemData ? itemData.price : 0, addons });
        }
    });
    return parsedItems;
}




// ===============================================================
//        ФУНКЦИИ ДЛЯ РАБОТЫ С НАСТРОЙКАМИ
// ===============================================================


// ===============================================================
//         ФУНКЦИИ ДЛЯ РАБОТЫ С НАСТРОЙКАМИ
// ===============================================================

function getAppSettings() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) return { paymentMethods: [], deliveryTypes: {} };

    const settings = {
      paymentMethods: [],  // Это будет массив объектов для способов оплаты
      deliveryTypes: {}    // Это будет объект для типов заказа
    };

    // Читаем весь диапазон настроек M:Q
    const data = settingsSheet.getRange("M2:Q" + settingsSheet.getLastRow()).getValues();
    
    // Список известных типов заказов (НЕ способы оплаты)
    const knownDeliveryTypes = ["Зал", "Доставка", "На вынос"];

    data.forEach(row => {
      const name = row[0]; // Название из колонки M
      const isEnabled = row[1] === true; // Галочка из колонки N
      // Колонка Q - пятая по счету в диапазоне M:Q, поэтому ее индекс 4
      const locationsRaw = row[4] || ''; 

      // Пропускаем пустые или отключенные строки
      if (!name || !isEnabled) {
        return;
      }

      // Проверяем, является ли запись типом заказа
      if (knownDeliveryTypes.includes(name)) {
        // Если это тип заказа...
        settings.deliveryTypes[name] = true;
      } else {
        // Иначе, это способ оплаты...
        settings.paymentMethods.push({
          name: name.trim(),
          locations: locationsRaw === '' ? [] : String(locationsRaw).split(',').map(s => s.trim())
        });
      }
    });

    return settings;

  } catch (e) {
    Logger.log("Ошибка в getAppSettings: " + e.message);
    return { paymentMethods: [], deliveryTypes: {} };
  }
}


function getTelegramConfig(selectedLocationName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) return { token: null, chatId: null };
  const settingsData = settingsSheet.getDataRange().getValues();
  let config = { token: null, chatId: null };
  let defaultChatId = null;
  for (let i = 1; i < settingsData.length; i++) {
    const row = settingsData[i];
    const locationParamName = row[0];
    const locationParamValue = row[1];
    if (locationParamName === `Telegram_Chat_ID_${selectedLocationName}`) { config.chatId = locationParamValue; }
    else if (locationParamName === "Telegram_Chat_ID_По умолчанию") { defaultChatId = locationParamValue; }
    const globalParamName = row[2];
    const globalParamValue = row[3];
    if (!config.token && globalParamName === "Telegram_Bot_Token") { config.token = globalParamValue; }
  }
  if (!config.chatId) { config.chatId = defaultChatId; }
  return config;
}


function getDeliveryFeeTiers() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  const tiers = [];
  if (settingsSheet) {
    const data = settingsSheet.getRange("I3:J" + settingsSheet.getLastRow()).getValues();
    data.forEach(row => {
      const km = parseFloat(String(row[0]).replace(',', '.'));
      const fee = parseFloat(String(row[1]).replace(',', '.'));
      if (!isNaN(km) && !isNaN(fee) && km > 0) {
        tiers.push({ km: km, fee: fee });
      }
    });
  }
  return tiers.sort((a, b) => a.km - b.km);
}




// ===============================================================
//        ЛОГИКА WHATSAPP И РЕДАКТОРА
// ===============================================================

/**
 * Генерирует и форматирует сообщение о заказе для разных платформ.
 * ИСПРАВЛЕННАЯ ВЕРСИЯ (с рублями): Все пользовательские данные обернуты в escapeMarkdown.
 * @param {object} data - Объект с данными заказа.
 * @returns {object} Объект, содержащий отформатированные строки для Telegram и WhatsApp.
 */
function generateOrderMessageParts(data) {
    const rawClientPhone = (data.clientPhone || '').toString().replace(/\D/g, '');
    const clientPhoneFormatted = `+${rawClientPhone.substring(0, 1)} (${rawClientPhone.substring(1, 4)}) ${rawClientPhone.substring(4, 7)}-${rawClientPhone.substring(7, 9)}-${rawClientPhone.substring(9, 11)}`;
    const separator = '\n`--------------------------------------`\n';

    const formatOrderItems = (items) => {
        if (!items || items.length === 0) {
            return [];
        }
        
        const groupedItems = {};
        items.forEach(item => {
            const key = item.name + (item.addons ? JSON.stringify(item.addons.map(a => a.name)) : '');
            if (!groupedItems[key]) {
                groupedItems[key] = { ...item, quantity: 0 };
            }
            groupedItems[key].quantity += item.quantity;
        });

        return Object.values(groupedItems).map(item => {
            const itemPrice = item.price || 0;
            const itemQuantity = item.quantity || 0;
            const itemSum = itemPrice * itemQuantity;
            
            let itemTextTelegram = `* ${escapeMarkdown(item.name)} ${escapeMarkdown(itemQuantity)} шт\\. x ${escapeMarkdown(itemPrice.toFixed(0))} руб\\. \\= ${escapeMarkdown(itemSum.toFixed(0))} руб\\.`;
            let itemTextWhatsapp = `* ${item.name} ${itemQuantity} шт. x ${itemPrice.toFixed(0)} руб. = ${itemSum.toFixed(0)} руб.`;
            
            if (item.addons && item.addons.length > 0) {
                const addonsTextTelegram = item.addons.map(addon => {
                    const addonPrice = addon.price || 0;
                    const addonQuantity = addon.quantity || 0;
                    const addonSum = addonPrice * addonQuantity;
                    return `\n    \\+ ${escapeMarkdown(addon.name)} ${escapeMarkdown(addonQuantity)} шт\\. x ${escapeMarkdown(addonPrice.toFixed(0))} руб\\. \\= ${escapeMarkdown(addonSum.toFixed(0))} руб\\.`;
                }).join('');
                itemTextTelegram += addonsTextTelegram;
                
                const addonsTextWhatsapp = item.addons.map(addon => {
                    const addonPrice = addon.price || 0;
                    const addonQuantity = addon.quantity || 0;
                    const addonSum = addonPrice * addonQuantity;
                    return `\n    + ${addon.name} ${addonQuantity} шт. x ${addonPrice.toFixed(0)} руб. = ${addonSum.toFixed(0)} руб.`;
                }).join('');
                itemTextWhatsapp += addonsTextWhatsapp;
            }
            return { telegram: itemTextTelegram, whatsapp: itemTextWhatsapp };
        });
    };
    
    const telegramItemsArray = formatOrderItems(data.cartItems);
    const telegramItems = telegramItemsArray.length > 0 ? telegramItemsArray.map(i => i.telegram).join('\n') : '_Состав заказа пуст_';

    const telegramSummaryInfo = `*Сумма заказа:* *${escapeMarkdown(Number(data.subtotalAmount || 0).toFixed(0))} руб*` +
        (Number(data.deliveryFee || 0) > 0 ? `\n*Доставка:* *${escapeMarkdown(Number(data.deliveryFee || 0).toFixed(0))} руб*` : '') +
        `\n*ИТОГО:* *${escapeMarkdown(Number(data.totalAmount || 0).toFixed(0))} руб*${separator}`;

    const telegramBody =
        `*Тип заказа:* ${escapeMarkdown(data.deliveryType)}\n` +
        `*Оплата:* ${escapeMarkdown(data.paymentMethod)}\n` +
        (data.paymentMethod === 'Наличными' && data.changeFrom ? `*Сдача с:* ${escapeMarkdown(data.changeFrom)}\n` : '') +
        `*Время:* ${escapeMarkdown(String(data.selectedTime))}${separator}` +
        `*Клиент:* ${escapeMarkdown(data.clientName)} \\(${escapeMarkdown(clientPhoneFormatted)}\\)\n` +
        `📞 [Позвонить](${escapeMarkdown('tel:+' + rawClientPhone)}) 💬 [Написать в WhatsApp](${escapeMarkdown(`https://wa.me/${rawClientPhone}`)})\n` +
        `*Адрес:* ${escapeMarkdown(data.deliveryAddress)}${separator}` +
        `*Состав:*\n${telegramItems}\n${separator}` +
        telegramSummaryInfo +
        `*Комментарий:* ${escapeMarkdown(data.comments || 'Нет')}${separator}` +
        (data.yandexMapsLink ? `[🗺️ Маршрут на Яндекс\\.Картах](${escapeMarkdown(data.yandexMapsLink)})\n` : '');


    // --- Сообщение для WhatsApp (остается без изменений) ---
    let whatsappText = '';
    const whatsappItemsArray = formatOrderItems(data.cartItems);
    const whatsappItems = whatsappItemsArray.length > 0 ? whatsappItemsArray.map(i => i.whatsapp).join('\n') : 'Состав заказа пуст.';
    
    if (data.status === 'Новый' || data.status === 'Подтвержден') {
      whatsappText += `👋 Здравствуйте, ${data.clientName}! Информация по вашему заказу №${data.orderNumber} в «${data.selectedLocation}»:\n\n*Способ получения:* ${data.deliveryType}\n`;
      if (data.deliveryType === 'Доставка') { whatsappText += `*Адрес доставки:* ${data.deliveryAddress}\n`; }
      whatsappText += `*Оплата:* ${data.paymentMethod}\n*Время:* ${String(data.selectedTime)}\n\n*Чек по вашему заказу:*\n${whatsappItems}\n\n*Сумма заказа:* ${Number(data.subtotalAmount || 0).toFixed(0)} руб\n`;
      if (Number(data.deliveryFee || 0) > 0) { whatsappText += `*Доставка:* ${Number(data.deliveryFee || 0).toFixed(0)} руб\n`; }
      whatsappText += `*Итого к оплате:* ${Number(data.totalAmount || 0).toFixed(0)} руб\n\n✅ Ваш заказ принят. Для передачи на кухню, просим подтвердить ваш заказ.`;
    } else {
      whatsappText += `Статус вашего заказа №${data.orderNumber} обновлен: ${data.status}.`;
    }

    return { 
        telegramBody: telegramBody,
        whatsappLink: `https://wa.me/${rawClientPhone}?text=${encodeURIComponent(whatsappText)}`
    };
}



function getPaymentDetailsForLocation(locationName) {
    try {
        const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
        if (!settingsSheet) return [];
        const data = settingsSheet.getRange("O2:Q" + settingsSheet.getLastRow()).getValues();
        const relevantDetails = [];
        data.forEach(row => {
            const name = row[0], number = row[1], locations = String(row[2] || '').trim();
            if (name && number && (locations === '' || locations.split(',').map(s => s.trim()).includes(locationName))) {
                relevantDetails.push({ name: name, number: number });
            }
        });
        return relevantDetails;
    } catch(e) {
        Logger.log("Ошибка в getPaymentDetailsForLocation: " + e.message);
        return [];
    }
}


function getEditorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) return { allMenuItems: [], allAddonItems: [] };
  const data = baseSheet.getDataRange().getValues();
  const menuItems = [], addonItems = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i], itemName = row[BASE_ITEM_NAME_COL], itemGroup = row[BASE_GROUP_COL];
    const regularPrice = parseFloat(String(row[BASE_PRICE_COL]).replace(/[^\d.,]/g, '').replace(',', '.'));
    if (itemName && !isNaN(regularPrice)) {
      let finalPrice = regularPrice;
      if (row[BASE_PROMO_PRICE_COL]) {
        const promoPrice = parseFloat(String(row[BASE_PROMO_PRICE_COL]).replace(/[^\d.,]/g, '').replace(',', '.'));
        if (!isNaN(promoPrice) && promoPrice > 0) finalPrice = promoPrice;
      }
      const itemData = { name: String(itemName).trim(), price: finalPrice, hasAddons: row[BASE_HAS_ADDONS_COL] === true };
      if (itemGroup === 'Дополнительно') { addonItems.push(itemData); }
      else { menuItems.push(itemData); }
    }
  }
  return { allMenuItems: menuItems, allAddonItems: addonItems };
}


function updateOrderFromSidebar(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
    sheet.getRange(data.row, 7).setValue(data.newItemsText);
    sheet.getRange(data.row, 8).setValue(data.newTotal);
    return "Заказ успешно обновлен!";
  } catch (e) {
    return "Ошибка при обновлении: " + e.message;
  }
}




// ===============================================================
//        БЕЗОПАСНОСТЬ, ОТЧЕТЫ И НАСТРОЙКА
// ===============================================================


function getRolesAndPins() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    const data = settingsSheet.getRange("K2:L" + settingsSheet.getLastRow()).getValues();
    const roles = {};
    data.forEach(row => {
      if (row[0] && row[1]) {
        roles[row[0].toString().trim()] = row[1].toString().trim();
      }
    });
    return roles;
  } catch (e) {
    Logger.log("Ошибка получения ролей и пин-кодов: " + e.message);
    return {};
  }
}


function validatePinForRoles(requiredRoles, allRolesAndPins) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Требуется подтверждение', 'Для выполнения этого действия введите ваш ПИН-код:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const enteredPin = response.getResponseText().trim();
  for (const role of requiredRoles) {
    if (allRolesAndPins[role] === enteredPin) return role;
  }
  ui.alert('Неверный ПИН-код', 'У вас нет прав для выполнения этого действия.', ui.ButtonSet.OK);
  return null;
}


function logChange(user, orderNumber, action, oldValue, newValue, reason = '') {
  try {
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Логи");
    if (logsSheet) {
      logsSheet.appendRow([new Date(), user, orderNumber, action, oldValue, newValue, reason]);
    }
  } catch(e) { Logger.log("Ошибка записи в лог: " + e.message); }
}


function generateAndSendDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName(ORDERS_SHEET_NAME);
  const logsSheet = ss.getSheetByName("Логи");
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const botToken = getTelegramConfig("По умолчанию").token;


  if (!botToken) { Logger.log("Отчет не отправлен: не найден токен бота."); return; }


  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const reportDate = Utilities.formatDate(yesterday, "GMT+6", "dd.MM.yyyy");


  const allOrders = ordersSheet.getDataRange().getValues();
  const allLogs = logsSheet ? logsSheet.getDataRange().getValues() : [];


  let totalRevenue = 0, deliveryOrdersCount = 0, pickupOrdersCount = 0, hallOrdersCount = 0;
  let deliveryRevenue = 0, pickupRevenue = 0, hallRevenue = 0, cashRevenue = 0, cardRevenue = 0;


  const relevantOrders = allOrders.filter(row => {
    if (!row[0] || !row[2]) return false;
    return new Date(row[0]).toDateString() === yesterday.toDateString() && row[2] === 'Доставлен';
  });


  relevantOrders.forEach(row => {
    const amount = Number(row[7]) || 0;
    const orderType = row[15];
    const paymentType = row[14];
    totalRevenue += amount;
    if (orderType === 'Доставка') { deliveryOrdersCount++; deliveryRevenue += amount; }
    else if (orderType === 'На вынос') { pickupOrdersCount++; pickupRevenue += amount; }
    else if (orderType === 'Зал') { hallOrdersCount++; hallRevenue += amount; }
    if (paymentType === 'Наличными') { cashRevenue += amount; }
    else { cardRevenue += amount; }
  });


  let reportText = `*📊 Z\\-Отчет за ${escapeMarkdown(reportDate)}*\n\n` +
                   `*ОБЩИЕ ПОКАЗАТЕЛИ:*\n` +
                   `_Общая выручка:_ *${totalRevenue.toFixed(0)} руб*\n` +
                   `_Всего заказов:_ *${relevantOrders.length} шт\\.*\n\n` +
                   `*ПО ТИПУ ПОЛУЧЕНИЯ:*\n` +
                   `_Доставка:_ ${deliveryOrdersCount} шт\\. на *${deliveryRevenue.toFixed(0)} руб*\n` +
                   `_На вынос:_ ${pickupOrdersCount} шт\\. на *${pickupRevenue.toFixed(0)} руб*\n` +
                   `_В зале:_ ${hallOrdersCount} шт\\. на *${hallRevenue.toFixed(0)} руб*\n\n` +
                   `*ПО ТИПУ ОПЛАТЫ:*\n` +
                   `_Наличными:_ *${cashRevenue.toFixed(0)} руб*\n` +
                   `_Переводом/Картой:_ *${cardRevenue.toFixed(0)} руб*\n\n` +
                   `\`--------------------------------------\`\n` +
                   `*🔏 Журнал действий за день:*\n`;


  const relevantLogs = allLogs.filter(row => row[0] && new Date(row[0]).toDateString() === yesterday.toDateString());
  if (relevantLogs.length > 0) {
    relevantLogs.forEach(log => {
      const time = Utilities.formatDate(new Date(log[0]), "GMT+6", "HH:mm");
      reportText += `\`[${time}]\` *${escapeMarkdown(log[1])}*: ${escapeMarkdown(log[3])} в заказе *${escapeMarkdown(log[2])}* с \`${escapeMarkdown(log[4])}\` на \`${escapeMarkdown(log[5])}\`\n`;
    });
  } else {
    reportText += `_Действий, требующих логирования, за день не было\\._\n`;
  }


  const directorChatId = settingsSheet.getRange("D4").getValue();
  const managerChatId = settingsSheet.getRange("D5").getValue();


  if (directorChatId) sendTelegramMessage(directorChatId, reportText, null, botToken);
  if (managerChatId) sendTelegramMessage(managerChatId, reportText, null, botToken);
}


function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsConfig = {
    "База": [ "Наименование товара", "Цена", "Ссылка на Фото блюда", "Цена по акции", "Описание", "Есть допы", "", "Группа", "Точка продаж 1", "Адрес 1", "Точка продаж 2", "Адрес 2" ],
    "Заказы": [ "Дата и время заказа", "Номер заказа", "Статус заказа", "Имя клиента", "Телефон клиента", "Адрес доставки", "Детали заказа", "Общая сумма заказа", "Точка продаж", "Комментарии клиента", "Курьер", "Ссылка на Яндекс.Карты", "Telegram Chat ID", "Telegram Message ID", "Способ оплаты", "Тип заказа", "Время получения", "Сдача с", "Сумма доставки" ],
    "Клиенты": [ "Имя клиента", "Телефон клиента", "Основной адрес доставки", "Дата первого заказа", "Количество заказов", "Последний заказ" ],
    "Настройки": [ "Параметр для точки продаж", "ID чата для точки", "Общий параметр", "Значение общего параметра", "Точка (время)", "Время доставки (в часах)", "Время на вынос (в часах)", "", "Растояние км.", "Сумма доставки", "", "Роль", "Пин-код" ]
  };
  for (const sheetName in sheetsConfig) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) { sheet = ss.insertSheet(sheetName); }
    const headers = sheetsConfig[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  }
  const ordersSheet = ss.getSheetByName(ORDERS_SHEET_NAME);
  if (ordersSheet) {
    const statusColumn = ordersSheet.getRange("C2:C");
    const rules = SpreadsheetApp.newDataValidation().requireValueInList(["Новый", "Подтвержден", "Отправлен", "Доставлен", "Отказ"]).setAllowInvalid(false).build();
    statusColumn.setDataValidation(rules);
  }
}

/**
 * НОВАЯ ФУНКЦИЯ! Собирает все товары из базы в один объект для быстрого поиска.
 */
function getAllItemsMap() {
    try {
        const baseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BASE_SHEET_NAME);
        if (!baseSheet) return {};

        const baseData = baseSheet.getDataRange().getValues();
        const allItemsMap = {};

        baseData.slice(1).forEach(row => {
            const name = String(row[BASE_ITEM_NAME_COL] || '').trim();
            if (!name) return;

            let price = parseFloat(String(row[BASE_PRICE_COL]).replace(/[^\d.,]/g, '').replace(',', '.'));
            if (row[BASE_PROMO_PRICE_COL]) {
                const promoPrice = parseFloat(String(row[BASE_PROMO_PRICE_COL]).replace(/[^\d.,]/g, '').replace(',', '.'));
                if (!isNaN(promoPrice) && promoPrice > 0) price = promoPrice;
            }

            if (!isNaN(price)) {
                allItemsMap[name.toLowerCase()] = {
                    name: name,
                    price: price,
                    promoPrice: row[BASE_PROMO_PRICE_COL] ? price : null,
                    imageUrl: row[BASE_IMAGE_URL_COL] || "",
                    description: row[BASE_DESCRIPTION_COL] || "",
                    group: row[BASE_GROUP_COL] || "Без категории",
                    hasAddons: row[BASE_HAS_ADDONS_COL] === true
                };
            }
        });
        return allItemsMap;
    } catch (e) {
        Logger.log("Критическая ошибка в getAllItemsMap: " + e.message);
        return {};
    }
}

/**
 * НОВАЯ ФУНКЦИЯ
 * Собирает список Chat ID для отправки общих уведомлений о новых заказах.
 * Данные берутся из листа "Настройки", колонка D, начиная с 3-й строки.
 */
function getNotificationChatIds() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Настройки");
    if (!settingsSheet) return [];
    
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 3) return []; // Если данных нет, возвращаем пустой массив
    
    // Читаем все значения из колонки D, начиная с D3
    const chatIdsRange = settingsSheet.getRange("D4:D" + lastRow).getValues();
    const chatIds = [];

    chatIdsRange.forEach(row => {
      const chatId = row[0];
      // Добавляем в список только если ячейка не пустая
      if (chatId && String(chatId).trim() !== '') {
        chatIds.push(String(chatId).trim());
      }
    });
    
    return chatIds;
  } catch (e) {
    Logger.log("Ошибка в getNotificationChatIds: " + e.message);
    return [];
  }
}

// ===============================================================
//     ЛОГИКА ОБРАБОТКИ И СОХРАНЕНИЯ ЗАКАЗА
// ===============================================================


/**
 * Главная функция-диспетчер. Получает заказ от клиента и решает,
 * что с ним делать: создать, обновить или дополнить.
 */
function processOrderSubmission(orderData) {
  try {
    const mode = orderData.editingState ? orderData.editingState.mode : null;
    const orderNumber = orderData.editingState ? orderData.editingState.number : null;

    if (mode === 'update' && orderNumber) {
      return updateExistingOrder(orderNumber, orderData);
    } else if (mode === 'add' && orderNumber) {
      return addToExistingOrder(orderNumber, orderData);
    } else {
      return createNewOrder(orderData);
    }
  } catch (e) {
    Logger.log("Критическая ошибка в processOrderSubmission: " + e.stack);
    throw new Error("Не удалось обработать заказ на сервере: " + e.message);
  }
}

/**
 * Полностью обновляет существующий заказ (для статуса "Новый").
 * Теперь обновляет не только состав, но и данные клиента.
 */
function updateExistingOrder(orderNumber, orderData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  const orderRow = findRowByOrderNumber(sheet, orderNumber);

  if (!orderRow) {
    throw new Error("Не удалось найти заказ " + orderNumber + " для обновления.");
  }

  // Создаем единый объект для обновления
  const dataToUpdate = {
    clientName: orderData.clientName,
    clientPhone: orderData.clientPhone,
    deliveryAddress: orderData.deliveryAddress,
    orderDetailsText: formatOrderDetailsForSheet(orderData.cartItems),
    totalAmount: orderData.totalAmount
  };

  sheet.getRange(orderRow, 4).setValue(dataToUpdate.clientName);
  sheet.getRange(orderRow, 5).setValue(dataToUpdate.clientPhone);
  sheet.getRange(orderRow, 6).setValue(dataToUpdate.deliveryAddress);
  sheet.getRange(orderRow, 7).setValue(dataToUpdate.orderDetailsText);
  sheet.getRange(orderRow, 8).setValue(dataToUpdate.totalAmount);

  // Обновляем сообщение в Telegram
  const updatedData = sheet.getRange(orderRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const orderForUpdate = {
    orderNumber: updatedData[ORDER_NUMBER_COL],
    status: updatedData[ORDER_STATUS_COL],
    clientName: updatedData[3],
    clientPhone: updatedData[4],
    deliveryAddress: updatedData[5],
    orderDetailsText: updatedData[ORDER_DETAILS_COL],
    totalAmount: updatedData[ORDER_TOTAL_COL],
    selectedLocation: updatedData[ORDER_LOCATION_COL],
    comments: updatedData[9],
    yandexMapsLink: updatedData[11],
    paymentMethod: updatedData[14],
    deliveryType: updatedData[15],
    selectedTime: updatedData[16],
    changeFrom: updatedData[17],
    deliveryFee: updatedData[18],
    cartItems: parseOrderDetailsString(updatedData[ORDER_DETAILS_COL])
  };

  // Теперь вызываем функцию без передачи листа и строки
  updateTelegramMessageForOrderFromData(orderForUpdate, "состав заказа");

  return { status: "success", orderNumber: orderNumber };
}

/**
 * Дополняет существующий заказ (для статуса "Подтвержден").
 */
function addToExistingOrder(orderNumber, orderData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  const orderRow = findRowByOrderNumber(sheet, orderNumber);

  if (!orderRow) {
    throw new Error("Не удалось найти заказ " + orderNumber + " для дополнения.");
  }

  const oldDetails = sheet.getRange(orderRow, ORDER_DETAILS_COL + 1).getValue();
  const oldTotal = Number(sheet.getRange(orderRow, ORDER_TOTAL_COL + 1).getValue() || 0);

  const newDetails = formatOrderDetailsForSheet(orderData.cartItems);
  const combinedDetails = oldDetails + "; " + newDetails;
  const newTotal = oldTotal + orderData.totalAmount;

  sheet.getRange(orderRow, ORDER_DETAILS_COL + 1).setValue(combinedDetails);
  sheet.getRange(orderRow, ORDER_TOTAL_COL + 1).setValue(newTotal);

  // Обновляем сообщение в Telegram
  const updatedData = sheet.getRange(orderRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const orderForUpdate = {
    orderNumber: updatedData[ORDER_NUMBER_COL],
    status: updatedData[ORDER_STATUS_COL],
    clientName: updatedData[3],
    clientPhone: updatedData[4],
    deliveryAddress: updatedData[5],
    orderDetailsText: combinedDetails,
    totalAmount: newTotal,
    selectedLocation: updatedData[ORDER_LOCATION_COL],
    comments: updatedData[9],
    yandexMapsLink: updatedData[11],
    paymentMethod: updatedData[14],
    deliveryType: updatedData[15],
    selectedTime: updatedData[16],
    changeFrom: updatedData[17],
    deliveryFee: updatedData[18],
    cartItems: parseOrderDetailsString(combinedDetails)
  };

  // Теперь вызываем функцию без передачи листа и строки
  updateTelegramMessageForOrderFromData(orderForUpdate, "дополнение к заказу");

  return { status: "success", orderNumber: orderNumber };
}

/**
 * Обновляет сообщение в Telegram при ручном редактировании заказа в таблице.
 * @param {number} editedRow Номер измененной строки.
 * @param {string} updatedField Название измененного поля.
 */
function updateTelegramMessageForOrder(editedRow, updatedField = "состав") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
    const updatedOrderDataRow = sheet.getRange(editedRow, 1, 1, 19).getValues()[0];
    
    // Создаем полный объект данных для обновления
    const cartItems = parseOrderDetailsString(updatedOrderDataRow[ORDER_DETAILS_COL]);
    const orderData = {
        orderNumber: updatedOrderDataRow[ORDER_NUMBER_COL],
        status: updatedOrderDataRow[ORDER_STATUS_COL],
        clientName: updatedOrderDataRow[3],
        clientPhone: updatedOrderDataRow[4],
        deliveryAddress: updatedOrderDataRow[5] || "Самовывоз",
        orderDetailsText: updatedOrderDataRow[ORDER_DETAILS_COL],
        cartItems: cartItems,
        totalAmount: Number(updatedOrderDataRow[ORDER_TOTAL_COL]),
        subtotalAmount: Number(updatedOrderDataRow[ORDER_TOTAL_COL]) - (Number(updatedOrderDataRow[18]) || 0),
        deliveryFee: Number(updatedOrderDataRow[18]) || 0,
        selectedLocation: updatedOrderDataRow[ORDER_LOCATION_COL],
        comments: updatedOrderDataRow[9] || "Нет",
        yandexMapsLink: updatedOrderDataRow[11],
        paymentMethod: updatedOrderDataRow[14],
        deliveryType: updatedOrderDataRow[15],
        selectedTime: updatedOrderDataRow[16],
        changeFrom: updatedOrderDataRow[17] || ""
    };

    updateTelegramMessageForOrderFromData(orderData, updatedField, true);

    SpreadsheetApp.getActiveSpreadsheet().toast(`Заказ #${orderData.orderNumber} в Telegram обновлен!`, '✅ Готово', 5);
}

/**
 * Обновляет сообщение в Telegram, принимая полный объект данных заказа.
 * @param {object} orderData Полный объект с данными заказа.
 * @param {string} updatedField Название измененного поля.
 */
function updateTelegramMessageForOrderFromData(orderData, updatedField) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ORDERS_SHEET_NAME);
    if (!sheet) {
      Logger.log("Ошибка: Лист 'Заказы' не найден.");
      return;
    }

    // Находим строку заказа по номеру заказа
    const data = sheet.getDataRange().getValues();
    let orderRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][ORDER_NUMBER_COL] === orderData.orderNumber) {
        orderRow = i + 1;
        break;
      }
    }

    if (orderRow === -1) {
      Logger.log("Заказ не найден в таблице: " + orderData.orderNumber);
      return;
    }

    const telegramConfig = getTelegramConfig("По умолчанию");
    if (!telegramConfig.token) {
      Logger.log("Ошибка: Telegram Bot Token не найден. Уведомление не отправлено.");
      return;
    }

    const finalOrderData = {
        ...orderData,
        subtotalAmount: orderData.totalAmount - (orderData.deliveryFee || 0),
    };

    const messageParts = generateOrderMessageParts(finalOrderData);
    const updateReason = escapeMarkdown(updatedField);
    const separator = '\n`--------------------------------------`\n';

    let finalMessageText = `*❗️ ЗАКАЗ ${escapeMarkdown(orderData.orderNumber)} ОБНОВЛЕН \\(${updateReason}\\)*\n` +
                           `_Текущий статус: ${escapeMarkdown(finalOrderData.status)}_${separator}` +
                           messageParts.telegramBody +
                           `*Текущий статус:* *${escapeMarkdown(finalOrderData.status)}*`;

    // Получаем JSON-строку с данными сообщений
    const messagesString = sheet.getRange(orderRow, ORDER_TELEGRAM_MESSAGES_COL + 1).getValue();
    let messagesData = [];
    if (messagesString) {
      try {
        messagesData = JSON.parse(messagesString);
      } catch (e) {
        Logger.log("Ошибка парсинга JSON-строки сообщений: " + e.message);
      }
    }

    // Проходим по каждому сохраненному сообщению и обновляем его
    messagesData.forEach(msg => {
      editTelegramMessage(msg.chatId, msg.messageId, finalMessageText, telegramConfig.token);
    });

    const emailTitle = `Заказ ОБНОВЛЕН (${updatedField})`;
    const emailBody = generateHtmlEmailBody(finalOrderData, emailTitle);
    sendEmailNotification(`${emailTitle} #${finalOrderData.orderNumber}`, emailBody);

  } catch (e) {
    Logger.log("Критическая ошибка в updateTelegramMessageForOrderFromData: " + e.stack);
    throw new Error("Ошибка обновления сообщения в Telegram: " + e.message);
  }
}

/**
 * Вызывается клиентом для отмены заказа со статусом "Новый".
 */
function cancelOrderByClient(orderNumber, clientPhone) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
    const orderRow = findRowByOrderNumber(sheet, orderNumber);
    if (!orderRow) {
      throw new Error("Заказ №" + orderNumber + " не найден.");
    }

    const rowData = sheet.getRange(orderRow, 1, 1, 19).getValues()[0];
    const currentStatus = rowData[ORDER_STATUS_COL];
    const orderPhone = normalizePhoneRU_GS(rowData[ORDER_PHONE_COL]);
    const requestPhone = normalizePhoneRU_GS(clientPhone);

    if (orderPhone !== requestPhone) {
      throw new Error("Ошибка безопасности: Попытка отменить чужой заказ.");
    }

    if (currentStatus !== 'Новый') {
      throw new Error("Нельзя отменить заказ. Он уже в работе. Статус: " + currentStatus);
    }

    sheet.getRange(orderRow, ORDER_STATUS_COL + 1).setValue("Отказ");

    const orderDataForUpdate = {
      orderNumber: orderNumber,
      status: "Отказ (отменен клиентом)",
      clientName: rowData[3],
      clientPhone: rowData[4],
      deliveryAddress: rowData[5],
      orderDetailsText: rowData[6],
      totalAmount: rowData[7],
      selectedLocation: rowData[8],
      comments: rowData[9],
      yandexMapsLink: rowData[11],
      paymentMethod: rowData[14],
      deliveryType: rowData[15],
      selectedTime: rowData[16],
      changeFrom: rowData[17],
      deliveryFee: rowData[18],
      cartItems: parseOrderDetailsString(rowData[6])
    };

    updateTelegramMessageForOrderFromData(orderDataForUpdate, "отменен клиентом");

    const emailTitle = "Заказ ОТМЕНЕН КЛИЕНТОМ";
    const emailBody = generateHtmlEmailBody(orderDataForUpdate, emailTitle);
    sendEmailNotification(`${emailTitle} #${orderNumber}`, emailBody);

    return { status: "success", message: "Заказ " + orderNumber + " успешно отменен." };

  } catch (e) {
    Logger.log("Ошибка в cancelOrderByClient: " + e.message);
    throw new Error("Ошибка на сервере: " + e.message);
  }
}

// ===============================================================
//         НОВЫЙ БЛОК: ФУНКЦИИ ДЛЯ E-MAIL УВЕДОМЛЕНИЙ
// ===============================================================

/**
 * Собирает все email-адреса для рассылки из листа "Настройки".
 * @returns {string} Строка с email-адресами через запятую, или null если адресов нет.
 */
function getEmailRecipients() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) return null;

    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) return null;

    // Читаем колонку R (18-я по счету) со второй строки
    const emailRange = settingsSheet.getRange("R2:R" + lastRow).getValues();
    const emailList = emailRange
      .map(row => row[0]) // Получаем значение из каждой ячейки
      .filter(email => typeof email === 'string' && email.includes('@')); // Оставляем только валидные email

    if (emailList.length > 0) {
      return emailList.join(','); // Возвращаем адреса в виде строки "a@b.com,c@d.com"
    } else {
      return null;
    }
  } catch (e) {
    Logger.log("Ошибка при получении списка email-адресов: " + e.message);
    return null;
  }
}

/**
 * Создает красивое HTML-тело письма с деталями заказа.
 * @param {object} orderData - Объект с данными заказа.
 * @param {string} title - Главный заголовок письма (напр. "Новый заказ").
 * @returns {string} Готовый HTML-код для вставки в письмо.
 */
function generateHtmlEmailBody(orderData, title) {
    let itemsHtml = '';
    const cartItems = orderData.cartItems || []; // Проверка на null или undefined
    cartItems.forEach(item => {
        if (!item || !item.name) return;
        const itemQuantity = item.quantity || 0;
        let addonsText = '';
        if (item.addons && item.addons.length > 0) {
            addonsText = item.addons.map(addon => {
                if (!addon || !addon.name) return '';
                const addonQuantity = addon.quantity || 0;
                return `&nbsp;&nbsp;&nbsp;+ ${addon.name} (${addonQuantity} шт.)`;
            }).join('<br>');
        }
        itemsHtml += `<b>${item.name}</b> (${itemQuantity} шт.)<br>${addonsText}`;
    });

    const subtotal = orderData.subtotalAmount || 0;
    const deliveryFee = orderData.deliveryFee || 0;
    const totalAmount = orderData.totalAmount || 0;

    const styles = `
        <style>
            body { font-family: Arial, sans-serif; color: #333; }
            .container { border: 1px solid #ddd; padding: 20px; max-width: 600px; margin: auto; border-radius: 8px; }
            h1 { color: #1a73e8; }
            table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            td { padding: 8px; border-bottom: 1px solid #eee; }
            td.label { font-weight: bold; width: 150px; }
        </style>
    `;

    return `
        <html>
        <head>${styles}</head>
        <body>
            <div class="container">
                <h1>${title} #${orderData.orderNumber}</h1>
                <p>Статус: <b>${orderData.status || 'Неизвестно'}</b></p>
                <table>
                    <tr><td class="label">Клиент:</td><td>${orderData.clientName || 'Не указано'}</td></tr>
                    <tr><td class="label">Телефон:</td><td>${orderData.clientPhone || 'Не указано'}</td></tr>
                    <tr><td class="label">Тип заказа:</td><td>${orderData.deliveryType || 'Не указано'}</td></tr>
                    <tr><td class="label">Адрес:</td><td>${orderData.deliveryAddress || 'Самовывоз'}</td></tr>
                    <tr><td class="label">Точка продаж:</td><td>${orderData.selectedLocation || 'Не указано'}</td></tr>
                    <tr><td class="label">Состав заказа:</td><td>${itemsHtml || 'Состав не указан'}</td></tr>
                    <tr><td class="label">Сумма:</td><td>${subtotal.toFixed(0)} руб</td></tr>
                    <tr><td class="label">Доставка:</td><td>${deliveryFee.toFixed(0)} руб</td></tr>
                    <tr><td class="label"><b>Итого:</b></td><td><b>${totalAmount.toFixed(0)} руб</b></td></tr>
                    <tr><td class="label">Оплата:</td><td>${orderData.paymentMethod || 'Не указано'}${orderData.changeFrom ? ` (Сдача с: ${orderData.changeFrom})` : ''}</td></tr>
                    <tr><td class="label">Комментарий:</td><td>${orderData.comments || 'Нет'}</td></tr>
                </table>
            </div>
        </body>
        </html>
    `;
}

/**
 * Главная функция для отправки email-уведомлений.
 * @param {string} subject - Тема письма.
 * @param {string} htmlBody - HTML-содержимое письма.
 */
function sendEmailNotification(subject, htmlBody) {
    const recipients = getEmailRecipients();
    if (recipients) {
        try {
            MailApp.sendEmail({
                to: recipients, // ИЗМЕНЕНИЕ: Убран .join()
                subject: subject,
                htmlBody: htmlBody
            });
            Logger.log("Email-уведомление успешно отправлено на: " + recipients);
        } catch (e) {
            Logger.log("Не удалось отправить email: " + e.message);
        }
    } else {
        Logger.log("Email-адреса для отправки не найдены в Настройках.");
    }
}

function grantMailPermission() {
  // Эта функция нужна только для того, чтобы вызвать окно разрешений
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Тест разрешений", "Это тестовое письмо для подтверждения разрешений.");
}

// ===============================================================
//         НОВАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ ПРОФИЛЯ КЛИЕНТА
// ===============================================================

// ===============================================================
//         НОВАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ ПРОФИЛЯ КЛИЕНТА
// ===============================================================

/**
 * Находит клиента по номеру телефона и обновляет его данные.
 * @param {object} profileData Объект с данными {phone, newName, newAddress}.
 * @returns {object} Объект со статусом операции.
 */
function updateClientProfile(profileData) {
  try {
    const { phone, newName, newAddress } = profileData;
    
    // 1. Проверяем, что все необходимые данные пришли с фронтенда
    if (!phone || !newName || !newAddress) {
      throw new Error("Не все данные для обновления профиля были предоставлены.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientsSheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
    
    if (!clientsSheet) {
      throw new Error(`Лист "${CLIENTS_SHEET_NAME}" не найден.`);
    }
    
    // 2. Используем TextFinder для эффективного поиска по номеру телефона в колонке B
    const phoneColumn = clientsSheet.getRange("B:B");
    // Ищем точное совпадение по всей ячейке
    const textFinder = phoneColumn.createTextFinder(phone).matchEntireCell(true);
    const foundCell = textFinder.findNext();
    
    // 3. Если ячейка найдена, обновляем данные в этой строке
    if (foundCell) {
      const row = foundCell.getRow();
      // Обновляем Имя (колонка A, индекс 1) и Адрес (колонка C, индекс 3)
      clientsSheet.getRange(row, 1).setValue(newName);
      clientsSheet.getRange(row, 3).setValue(newAddress);
      
      Logger.log(`Профиль для номера ${phone} обновлен. Новое имя: ${newName}, новый адрес: ${newAddress}`);
      return { status: "success", message: "Профиль успешно обновлен." };
    } else {
      // 4. Если клиент не найден, возвращаем ошибку
      Logger.log(`Не удалось найти клиента с номером ${phone} для обновления профиля.`);
      throw new Error("Не удалось найти ваш профиль для обновления.");
    }
  } catch (e) {
    Logger.log("Критическая ошибка в updateClientProfile: " + e.message);
    // "Пробрасываем" ошибку дальше, чтобы фронтенд мог ее поймать и показать пользователю
    throw new Error("Ошибка на сервере: " + e.message);
  }
}

/**
 * ОБНОВЛЕННАЯ ФУНКЦИЯ
 * Получает контактную информацию из колонок S (тип) и T (значение).
 */
function getContactInfo() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) return [];

    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) return [];

    // Читаем диапазон сразу из двух колонок S и T
    const contactsRange = settingsSheet.getRange("S2:T" + lastRow).getValues();
    const contacts = [];

    contactsRange.forEach(row => {
      const type = row[0];  // Данные из колонки S
      const value = row[1]; // Данные из колонки T

      // Добавляем контакт, только если указан и тип, и значение
      if (typeof type === 'string' && type.trim() !== '' && value) {
        contacts.push({ 
          type: type.trim().toLowerCase(), 
          value: value.toString().trim() 
        });
      }
    });

    Logger.log("Загружены контакты (новая структура): " + JSON.stringify(contacts));
    return contacts;

  } catch (e) {
    Logger.log("Ошибка в getContactInfo: " + e.message);
    return [];
  }
}

/**
 * ФОРМИРУЕТ И ОТПРАВЛЯЕТ УВЕДОМЛЕНИЕ О НОВОМ ЗАКАЗЕ В TELEGRAM
 * @param {object} orderData - Полный объект с данными заказа.
 */
function sendNewOrderNotification(orderData) {
  try {
    const ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
    const orderRow = findRowByOrderNumber(ordersSheet, orderData.orderNumber);
    if (!orderRow) {
      Logger.log(`[ОШИБКА] Не найдена строка для заказа ${orderData.orderNumber} для сохранения ID сообщений.`);
      return;
    }

    // Получаем настройки для Telegram
    const telegramConfig = getTelegramConfig(orderData.selectedLocation);
    if (!telegramConfig.token) {
      Logger.log("Отправка в Telegram пропущена: не найден токен бота.");
      return;
    }

    // Собираем всех получателей уведомления
    const notificationList = getNotificationChatIds();
    const allRecipients = new Set(notificationList); // Используем Set, чтобы избежать дубликатов
    if (telegramConfig.chatId) {
      allRecipients.add(String(telegramConfig.chatId));
    }

    if (allRecipients.size === 0) {
      Logger.log("Отправка в Telegram пропущена: не найдены ID чатов для уведомлений.");
      return;
    }

    // Генерируем текст сообщения и кнопку WhatsApp
    const messageParts = generateOrderMessageParts(orderData);
    const separator = '\n`--------------------------------------`\n';
    const finalMessageText = `*НОВЫЙ ЗАКАЗ \\#${escapeMarkdown(orderData.orderNumber)}*${separator}` +
                             messageParts.telegramBody +
                             `*Текущий статус:* Новый`;

    const inlineKeyboard = {
      inline_keyboard: [
        [{
          text: "💬 Написать клиенту в WhatsApp",
          url: messageParts.whatsappLink
        }]
      ]
    };

    // Отправляем сообщение каждому получателю и собираем ID
    const messagesData = [];
    allRecipients.forEach(chatId => {
      const messageId = sendTelegramMessage(chatId, finalMessageText, inlineKeyboard, telegramConfig.token);
      if (messageId) {
        messagesData.push({ chatId: String(chatId), messageId: messageId });
      }
    });

    // Если хотя бы одно сообщение было успешно отправлено, записываем ID в таблицу
    if (messagesData.length > 0) {
      ordersSheet.getRange(orderRow, ORDER_TELEGRAM_MESSAGES_COL + 1).setValue(JSON.stringify(messagesData));
      Logger.log(`ID сообщений для заказа ${orderData.orderNumber} успешно сохранены.`);
    }

  } catch (e) {
    Logger.log(`[КРИТИЧЕСКАЯ ОШИБКА] в функции sendNewOrderNotification: ${e.stack}`);
  }
}
