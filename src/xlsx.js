/**
 * Модуль трансформации таблиц Excel (XLSX).
 *
 * Отвечает за:
 * 1. Чтение входного XLSX-файла и валидацию структуры.
 * 2. Фильтрацию строк: отбор только тех, где в столбце J (Количество) указано число > 0.
 * 3. Извлечение необходимых колонок из исходной таблицы:
 *    - B (Штрихкод EAN13);
 *    - I (Цена);
 *    - J (Количество).
 * 4. Запрос идентификаторов и типов сущностей в сервисе «МойСклад» (через модуль src/moysklad.js).
 * 5. Формирование выходной книги Excel строго по шаблону «Образец»:
 *    - Строка 1: Заголовки ('Тип', 'UUID', 'Штрихкод EAN13', 'UUID товара модификации', 'Цена', 'Количество');
 *    - Для товаров (type = 'product'): строка «Товар» с заполнением колонок A, B, C, E, F (D оставляется пустой);
 *    - Для модификаций (type = 'variant'): группировка модификаций с одинаковым product.id под единой
 *      строкой «Товар» (в строке «Товар» заполняются только A и B, остальные колонки пустые),
 *      после чего следуют строки «Модификация» с заполнением A, B, C, D, E, F.
 */

import ExcelJS from 'exceljs';
import { fetchAssortmentByBarcodes } from './moysklad.js';

/**
 * Ошибка валидации структуры исходного файла Excel.
 * Выбрасывается, если файл поврежден, не содержит листов или не является ожидаемой таблицей.
 */
export class InvalidTableError extends Error {
  constructor() {
    super('Таблица некорректна');
    this.name = 'InvalidTableError';
  }
}

/**
 * Извлекает текстовое представление штрихкода из ячейки Excel.
 *
 * Обрабатывает случаи, когда штрихкод записан числом с форматом ведущих нулей (например, '0000000000000'),
 * предотвращая потерю ведущих нулей при конвертации в строку.
 *
 * @param {import('exceljs').Cell} cell - Ячейка Excel со штрихкодом.
 * @returns {string} Нормализованная строка штрихкода.
 */
function barcodeText(cell) {
  const value = cell.value;
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    const format = cell.numFmt?.split(';')[0];
    if (Number.isInteger(value) && value >= 0 && /^0+$/.test(format)) {
      return String(value).padStart(format.length, '0');
    }
  }
  return String(cell.text || value).trim();
}

/**
 * Извлекает числовое или скалярное значение из ячейки Excel.
 * Корректно разворачивает составные объекты ячеек ExcelJS (формулы, форматированные тексты).
 *
 * @param {import('exceljs').Cell} cell - Ячейка Excel.
 * @returns {any} Конечное значение ячейки.
 */
function extractCellValue(cell) {
  const val = cell.value;
  if (val === null || val === undefined) return null;
  // Если ячейка содержит вычисленную формулу: { formula: '...', result: 123 }
  if (typeof val === 'object' && 'result' in val) {
    return val.result;
  }
  return val;
}

/**
 * Главная функция трансформации входящего XLSX-файла в итоговую структуру МойСклад.
 *
 * @param {Buffer|Uint8Array} content - Бинарное содержимое входящего файла XLSX.
 * @param {object} moyskladConfig - Конфигурация доступа к API МойСклад ({ apiUrl, token }).
 * @param {typeof fetch} [customFetch=fetch] - Опциональный клиент fetch (для мокирования в тестах).
 * @returns {Promise<Buffer>} Буфер сформированного результирующего XLSX-документа.
 * @throws {InvalidTableError} Если структура входящего файла неверна или файл поврежден.
 * @throws {import('./moysklad.js').BarcodeNotFoundError} Если штрихкод не найден в МойСклад.
 * @throws {import('./moysklad.js').MoySkladUnavailableError} При сбоях API МойСклад.
 */
export async function transformXlsx(content, moyskladConfig, customFetch = fetch) {
  // 1. Загрузка исходной книги Excel
  const source = new ExcelJS.Workbook();
  try {
    await source.xlsx.load(content);
  } catch {
    throw new InvalidTableError();
  }

  const sheet = source.worksheets[0];
  if (!sheet) {
    throw new InvalidTableError();
  }

  // 2. Сканирование строк и сбор позиций с условием: J > 0 (число)
  // Столбец 2 = B (Штрихкод), Столбец 9 = I (Цена), Столбец 10 = J (Количество)
  const items = [];

  sheet.eachRow((row, rowNumber) => {
    // Пропускаем шапку таблицы (строка 1)
    if (rowNumber === 1) return;

    const rawQuantity = extractCellValue(row.getCell(10));
    const quantity = Number(rawQuantity);

    // Условие п. 2: значение в столбце J есть, числовое и больше 0
    if (rawQuantity !== null && rawQuantity !== '' && !Number.isNaN(quantity) && quantity > 0) {
      const barcode = barcodeText(row.getCell(2));
      const price = extractCellValue(row.getCell(9));

      items.push({
        barcode,
        price,
        quantity,
      });
    }
  });

  // 3. Запрос в API МойСклад по всем найденным штрихкодам
  const barcodes = items.map((it) => it.barcode).filter(Boolean);
  const uniqueCount = new Set(barcodes).size;
  const chunkCount = Math.ceil(uniqueCount / 10);
  console.log(`[INFO] [${new Date().toISOString()}] Отобрано позиций с количеством > 0: ${items.length}, уникальных штрихкодов: ${uniqueCount}, чанков МойСклад: ${chunkCount}`);

  const assortmentMap = await fetchAssortmentByBarcodes(barcodes, moyskladConfig, customFetch);
  // 4. Формирование итогового документа Excel
  const result = new ExcelJS.Workbook();
  const output = result.addWorksheet('1');

  // Строка 1: Заголовки строго в соответствии с файлом «Образец»
  output.addRow([
    'Тип',
    'UUID',
    'Штрихкод EAN13',
    'UUID товара модификации',
    'Цена',
    'Количество',
  ]);

  // Настройка типов и ширины колонок для удобства и сохранения строковых штрихкодов/UUID
  output.getColumn(1).width = 18; // Тип
  output.getColumn(2).width = 40; // UUID
  output.getColumn(3).width = 24; // Штрихкод EAN13
  output.getColumn(3).numFmt = '@';
  output.getColumn(4).width = 40; // UUID товара модификации
  output.getColumn(5).width = 16; // Цена
  output.getColumn(6).width = 16; // Количество

  // 5. Заполнение строк по правилам п. 4.1 и 4.2
  // Хранилище для отслеживания уже выведенных родительских строк "Товар" для модификаций (variant)
  // Позволяет группировать варианты с одинаковым product.id под одной строкой товара
  const emittedParentProductIds = new Set();

  for (const item of items) {
    const msEntity = assortmentMap.get(item.barcode);
    if (!msEntity) continue;

    if (msEntity.type === 'product') {
      // П. 4.1: type = product
      // а. A = "Товар"
      // б. B = id
      // в. C = B[n] (штрихкод)
      // г. D = пустая ячейка
      // д. E = I[n] (цена)
      // е. F = J[n] (количество)
      output.addRow([
        'Товар',
        msEntity.id,
        item.barcode,
        null,
        item.price,
        item.quantity,
      ]);
    } else if (msEntity.type === 'variant') {
      // П. 4.2: type = variant
      const parentProductId = msEntity.productId;

      // Если родительский товар для данной группы модификаций еще не был добавлен в таблицу:
      // выводим строку "Товар" с product.id (пункты 4.2.а - 4.2.б, а также п. 4.2.к)
      if (!emittedParentProductIds.has(parentProductId)) {
        output.addRow([
          'Товар',
          parentProductId,
          null,
          null,
          null,
          null,
        ]);
        emittedParentProductIds.add(parentProductId);
      }

      // Выводим строку "Модификация" (пункты 4.2.в - 4.2.з):
      // в. A = "Модификация"
      // г. B = id модификации
      // д. C = B[n] (штрихкод)
      // е. D = id модификации
      // ж. E = I[n] (цена)
      // з. F = J[n] (количество)
      output.addRow([
        'Модификация',
        msEntity.id,
        item.barcode,
        msEntity.id,
        item.price,
        item.quantity,
      ]);
    }
  }

  return Buffer.from(await result.xlsx.writeBuffer());
}
