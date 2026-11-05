/**
 * Модуль интеграции с REST API сервиса «МойСклад» (v1.2).
 *
 * Отвечает за:
 * 1. Получение информации об ассортименте (товарах и модификациях) по переданным штрихкодам.
 * 2. Порционную разбивку штрихкодов на чанки строго по 10 штук для предотвращения превышения длины URL и лимитов API.
 * 3. Валидацию наличия каждого запрошенного штрихкода в ответе МойСклад.
 * 4. Формирование специализированных ошибок с текстами для отправки клиенту в случае проблем.
 */

/**
 * Ошибка, возникающая когда один или несколько штрихкодов не найдены в сервисе МойСклад.
 */
export class BarcodeNotFoundError extends Error {
  /**
   * @param {string} barcode - Штрихкод, который не удалось найти в базе МойСклад.
   */
  constructor(barcode) {
    super(`Товар с ш/к ${barcode} не найден в МойСклад. Пожалуйста, проверьте правильность ш/к в таблице и в МойСклад и повторите попытку`);
    this.name = 'BarcodeNotFoundError';
    this.barcode = barcode;
  }
}

/**
 * Ошибка, возникающая при сетевых сбоях, недоступности сервиса МойСклад (5xx, таймаут)
 * или получении некорректного ответа от сервера API.
 */
export class MoySkladUnavailableError extends Error {
  /**
   * @param {string} [details] - Технические подробности ошибки для логов.
   */
  constructor(details) {
    super('API МойСклад временно недоступен. Пожалуйста, повторите попытку позже');
    this.name = 'MoySkladUnavailableError';
    this.details = details;
  }
}

/**
 * Вспомогательная функция для разделения массива элементов на чанки (порции) заданного размера.
 *
 * @template T
 * @param {T[]} array - Исходный массив элементов.
 * @param {number} size - Максимальный размер одной порции (чанка).
 * @returns {T[][]} Массив порций.
 *
 * Пример: chunkArray([1, 2, 3, 4, 5], 2) => [[1, 2], [3, 4], [5]]
 */
export function chunkArray(array, size) {
  if (!Array.isArray(array) || size <= 0) return [];
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

/**
 * Выполняет один GET-запрос к API МойСклад для одной пачки штрихкодов (до 10 штук).
 *
 * Формирует фильтр вида:
 * filter=barcode=B1;barcode=B2;barcode=B3...
 * и параметр expand=product для автоматической подгрузки родительского товара у модификаций.
 *
 * @param {string[]} barcodeChunk - Массив штрихкодов (максимум 10 элементов).
 * @param {object} config - Конфигурация API МойСклад.
 * @param {string} config.apiUrl - Базовый URL API (например, https://api.moysklad.ru/api/remap/1.2).
 * @param {string} config.token - Секретный Bearer-токен авторизации.
 * @param {typeof fetch} [customFetch=fetch] - Функция для отправки HTTP-запроса (для внедрения моков в тестах).
 * @returns {Promise<Array<object>>} Список сущностей ассортимента из МойСклад для данного чанка.
 * @throws {MoySkladUnavailableError} При сетевых ошибках, HTTP статусах 4xx/5xx или недоступности API.
 */
export async function fetchAssortmentChunk(barcodeChunk, config, customFetch = fetch) {
  if (!barcodeChunk.length) return [];

  // Формируем строковый фильтр МойСклад: фильтры через точку с запятой работают как логическое «ИЛИ»
  // для одинаковых ключей в API МойСклад (filter=barcode=111;barcode=222)
  const filterParam = barcodeChunk.map((bc) => `barcode=${encodeURIComponent(bc)}`).join(';');
  const url = `${config.apiUrl}/entity/assortment?limit=${barcodeChunk.length}&expand=product&filter=${filterParam}`;

  let response;
  try {
    response = await customFetch(url, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${config.token}`,
        'Content-Type': 'application/json;charset=utf-8',
        'Accept-Encoding': 'gzip',
      },
    });
  } catch (networkError) {
    throw new MoySkladUnavailableError(networkError.message);
  }

  if (!response.ok) {
    throw new MoySkladUnavailableError(`HTTP ${response.status}: ${response.statusText}`);
  }

  let data;
  try {
    data = await response.json();
  } catch (parseError) {
    throw new MoySkladUnavailableError(`Ошибка парсинга JSON: ${parseError.message}`);
  }

  if (!data || !Array.isArray(data.rows)) {
    throw new MoySkladUnavailableError('Ответ API МойСклад не содержит массив rows');
  }

  return data.rows;
}

/**
 * Проверяет, привязан ли указанный штрихкод к конкретной сущности МойСклад (товару или модификации).
 *
 * В МойСклад штрихкоды сущности могут храниться:
 * 1. В массиве barcodes: [{ ean13: "...", ... }, ...] или строках/объектах с кодами;
 * 2. В поле barcode (в старых или упрощенных схемах).
 *
 * @param {object} entity - Сущность из API МойСклад (товар или модификация).
 * @param {string} barcode - Искомый строковый штрихкод.
 * @returns {boolean} true, если штрихкод совпал.
 */
function entityMatchesBarcode(entity, barcode) {
  const target = String(barcode).trim();

  // Проверка прямого совпадения строкового поля barcode
  if (entity.barcode && String(entity.barcode).trim() === target) {
    return true;
  }

  // Проверка вложенного массива barcodes
  if (Array.isArray(entity.barcodes)) {
    for (const item of entity.barcodes) {
      if (typeof item === 'string' && item.trim() === target) {
        return true;
      }
      if (item && typeof item === 'object') {
        // Объект штрихкода может иметь вид { ean13: "460..." } или { code128: "..." }
        for (const val of Object.values(item)) {
          if (String(val).trim() === target) {
            return true;
          }
        }
      }
    }
  }

  return false;
}

/**
 * Основная функция получения ассортимента из МойСклад по списку штрихкодов.
 *
 * Алгоритм работы:
 * 1. Получает на вход массив штрихкодов из входного файла.
 * 2. Выделяет уникальные штрихкоды и делит их на чанки строго по 10 штук.
 * 3. Отправляет GET-запросы по каждому чанку в МойСклад.
 * 4. Собирает все найденные сущности в единый массив.
 * 5. Для каждого исходного штрихкода ищет соответствующую запись в полученном массиве:
 *    - Если для какого-либо штрихкода товар не найден -> выбрасывает ошибку BarcodeNotFoundError.
 * 6. Возвращает структуру данных со сведениями по каждому запрошенному штрихкоду.
 *
 * @param {string[]} barcodes - Массив штрихкодов из входной таблицы.
 * @param {object} config - Конфигурация подключения к МойСклад.
 * @param {string} config.apiUrl - URL API.
 * @param {string} config.token - Bearer-токен.
 * @param {typeof fetch} [customFetch=fetch] - Опциональный клиент fetch (для тестов).
 * @returns {Promise<Map<string, { type: 'product'|'variant', id: string, productId?: string }>>}
 * Карта соответствия: barcode -> нормализованные данные сущности в МойСклад.
 * @throws {BarcodeNotFoundError} Если хотя бы один штрихкод не найден в МойСклад.
 * @throws {MoySkladUnavailableError} Если сервис МойСклад вернул ошибку или недоступен.
 */
export async function fetchAssortmentByBarcodes(barcodes, config, customFetch = fetch) {
  if (!barcodes || barcodes.length === 0) {
    return new Map();
  }

  // Получаем уникальный список штрихкодов для оптимизации числа сетевых запросов
  const uniqueBarcodes = [...new Set(barcodes.map((b) => String(b).trim()).filter(Boolean))];

  // Разбиваем на чанки строго по 10 штрихкодов в соответствии с техническим заданием
  const chunks = chunkArray(uniqueBarcodes, 10);

  // Выполняем запросы для всех чанков и объединяем результаты в единый массив
  const allRows = [];
  for (const chunk of chunks) {
    const rows = await fetchAssortmentChunk(chunk, config, customFetch);
    allRows.push(...rows);
  }

  // Создаем карту соответствия: штрихкод -> нормализованная запись
  const barcodeMap = new Map();

  for (const barcode of uniqueBarcodes) {
    // Ищем сущность, которой принадлежит данный штрихкод
    const entity = allRows.find((row) => entityMatchesBarcode(row, barcode));

    // Если товар по штрихкоду не найден — немедленно прерываем с указанием точного штрихкода
    if (!entity) {
      throw new BarcodeNotFoundError(barcode);
    }

    // Определяем тип сущности по метаданным (type = 'product' или 'variant')
    const type = entity.meta?.type || (entity.product ? 'variant' : 'product');
    const id = entity.id;

    let productId = null;
    if (type === 'variant') {
      // Для модификаций id родительского товара берется из раскрытого объекта product
      productId = entity.product?.id || entity.product?.meta?.href?.split('/')?.pop() || null;
    }

    barcodeMap.set(barcode, {
      type,
      id,
      productId,
    });
  }

  return barcodeMap;
}
