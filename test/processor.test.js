/**
 * Набор модульных и интеграционных тестов для проверки нового пайплайна:
 * 1. Интеграция с API «МойСклад» (чанки по 10 элементов, обработка ошибок).
 * 2. Парсинг входящих XLSX, фильтрация строк по J > 0, извлечение B, I, J.
 * 3. Формирование результирующего XLSX по образцу (разметка типов "Товар", "Модификация", группировка variant по product.id).
 * 4. Формирование текстов ответных писем при успехе, отсутствии товара в МойСклад и сетевых сбоях.
 */

import assert from 'node:assert/strict';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import { InvalidTableError, transformXlsx } from '../src/xlsx.js';
import { prepareReply, processMessage, shouldSkip } from '../src/mail.js';
import { drainMailbox } from '../src/service.js';
import { StateError, StateStore } from '../src/state.js';
import {
  BarcodeNotFoundError,
  MoySkladUnavailableError,
  chunkArray,
  fetchAssortmentByBarcodes,
} from '../src/moysklad.js';

/**
 * Вспомогательная функция для генерации входящей книги Excel для тестов.
 *
 * @param {Array<{ barcode: any, price?: any, quantity: any }>} rows - Тестовые строки.
 * @returns {Promise<Buffer>} Буфер файла XLSX.
 */
async function createTestInputXlsx(rows = []) {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('Лист1');

  // Шапка таблицы (строка 1)
  sheet.getCell('B1').value = 'Штрихкод';
  sheet.getCell('I1').value = 'Цена';
  sheet.getCell('J1').value = 'Количество';

  // Заполнение строк данных
  rows.forEach((row, index) => {
    const rowNum = index + 2;
    sheet.getCell(`B${rowNum}`).value = row.barcode;
    if (row.price !== undefined) sheet.getCell(`I${rowNum}`).value = row.price;
    sheet.getCell(`J${rowNum}`).value = row.quantity;
  });

  return Buffer.from(await book.xlsx.writeBuffer());
}

/**
 * Читает буфер XLSX и возвращает объект Workbook для проверки утверждений.
 */
async function readWorkbook(content) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(content);
  return book;
}

/**
 * Вспомогательный конструктор распарсенного письма.
 */
function createMockMail(attachments = []) {
  return {
    from: { value: [{ address: 'sender@example.org' }] },
    headers: new Map(),
    attachments,
    messageId: '<original@example.org>',
  };
}

function memoryStore() {
  const store = new StateStore('unused', 'account', 'INBOX');
  store.save = async (value) => { store.value = structuredClone(value); };
  return store;
}

// =========================================================================
// ТЕСТЫ МОДУЛЯ MOYSKLAD
// =========================================================================

test('chunkArray корректно делит массив на чанки указанного размера', () => {
  assert.deepEqual(chunkArray([1, 2, 3, 4, 5], 2), [[1, 2], [3, 4], [5]]);
  assert.deepEqual(chunkArray([], 10), []);
  assert.deepEqual(chunkArray(['a', 'b'], 10), [['a', 'b']]);
});

test('fetchAssortmentByBarcodes разбивает штрихкоды на чанки по 10 и объединяет результат', async () => {
  const requestedUrls = [];
  const fakeBarcodes = Array.from({ length: 25 }, (_, i) => `20000000000${i.toString().padStart(2, '0')}`);

  const mockFetch = async (url) => {
    requestedUrls.push(url);
    const match = url.match(/filter=(.*?)$/);
    const filterParts = match ? match[1].split(';') : [];
    const barcodesInFilter = filterParts.map((f) => decodeURIComponent(f.replace('barcode=', '')));

    return {
      ok: true,
      json: async () => ({
        rows: barcodesInFilter.map((bc, idx) => ({
          id: `uuid-${bc}`,
          barcode: bc,
          meta: { type: idx % 2 === 0 ? 'product' : 'variant' },
          product: idx % 2 !== 0 ? { id: `parent-${bc}` } : undefined,
        })),
      }),
    };
  };

  const map = await fetchAssortmentByBarcodes(fakeBarcodes, { apiUrl: 'https://api.test', token: 'mock-token' }, mockFetch);

  // 25 штрихкодов должны дать 3 чанка: 10 + 10 + 5
  assert.equal(requestedUrls.length, 3);
  assert.equal(map.size, 25);
  assert.equal(map.get('2000000000000').type, 'product');
  assert.equal(map.get('2000000000001').type, 'variant');
  assert.equal(map.get('2000000000001').productId, 'parent-2000000000001');
});

test('fetchAssortmentByBarcodes выбрасывает BarcodeNotFoundError при отсутствии товара', async () => {
  const mockFetch = async () => ({
    ok: true,
    json: async () => ({ rows: [] }),
  });

  await assert.rejects(
    fetchAssortmentByBarcodes(['4600000000001'], { apiUrl: 'https://api.test', token: 'tok' }, mockFetch),
    (err) => err instanceof BarcodeNotFoundError && err.barcode === '4600000000001',
  );
});

test('fetchAssortmentByBarcodes выбрасывает MoySkladUnavailableError при 5xx ответе или сбое сети', async () => {
  const mockFetchFail = async () => ({
    ok: false,
    status: 502,
    statusText: 'Bad Gateway',
  });

  await assert.rejects(
    fetchAssortmentByBarcodes(['4600000000001'], { apiUrl: 'https://api.test', token: 'tok' }, mockFetchFail),
    MoySkladUnavailableError,
  );
});

// =========================================================================
// ТЕСТЫ ТРАНСФОРМАЦИИ XLSX И СТРУКТУРЫ "ОБРАЗЕЦ"
// =========================================================================

test('генерация выходного файла по образцу (п. 4.1 Товар и п. 4.2 Модификация с группировкой)', async () => {
  const inputBuffer = await createTestInputXlsx([
    // Строка 2: Товар (product)
    { barcode: '4600000000001', price: 1500, quantity: 2 },
    // Строка 3: Модификация 1 первого родителя
    { barcode: '4600000000002', price: 2000, quantity: 5 },
    // Строка 4: Модификация 2 того же первого родителя
    { barcode: '4600000000003', price: 2100, quantity: 3 },
    // Строка 5: Строка с количеством 0 (должна быть пропущена)
    { barcode: '4600000000004', price: 500, quantity: 0 },
    // Строка 6: Модификация второго родителя
    { barcode: '4600000000005', price: 3000, quantity: 1 },
  ]);

  const mockEntities = [
    { id: 'prod-uuid-1', barcode: '4600000000001', meta: { type: 'product' } },
    {
      id: 'var-uuid-2',
      barcode: '4600000000002',
      meta: { type: 'variant' },
      product: { id: 'parent-uuid-A' },
    },
    {
      id: 'var-uuid-3',
      barcode: '4600000000003',
      meta: { type: 'variant' },
      product: { id: 'parent-uuid-A' },
    },
    {
      id: 'var-uuid-5',
      barcode: '4600000000005',
      meta: { type: 'variant' },
      product: { id: 'parent-uuid-B' },
    },
  ];

  const mockFetch = async () => ({
    ok: true,
    json: async () => ({ rows: mockEntities }),
  });

  const outputBuffer = await transformXlsx(inputBuffer, { apiUrl: 'https://api.test', token: 'mock' }, mockFetch);
  const outBook = await readWorkbook(outputBuffer);
  const sheet = outBook.worksheets[0];

  // Вспомогательная функция для получения значений ячеек колонок A-F (1-6) с нормализацией пустых ячеек в null
  const getRowValues = (rowNum) => [1, 2, 3, 4, 5, 6].map((col) => sheet.getRow(rowNum).getCell(col).value ?? null);

  // Проверка строки 1 (Заголовки из файла "Образец")
  assert.deepEqual(getRowValues(1), [
    'Тип',
    'UUID',
    'Штрихкод EAN13',
    'UUID товара модификации',
    'Цена',
    'Количество',
  ]);

  // Строка 2: Товар (п. 4.1)
  assert.deepEqual(getRowValues(2), [
    'Товар',
    'prod-uuid-1',
    '4600000000001',
    null,
    1500,
    2,
  ]);

  // Строка 3: Родитель для модификаций parent-uuid-A (п. 4.2.а-б)
  assert.deepEqual(getRowValues(3), [
    'Товар',
    'parent-uuid-A',
    null,
    null,
    null,
    null,
  ]);

  // Строка 4: Модификация 1 parent-uuid-A (п. 4.2.в-з)
  assert.deepEqual(getRowValues(4), [
    'Модификация',
    'var-uuid-2',
    '4600000000002',
    'parent-uuid-A',
    2000,
    5,
  ]);

  // Строка 5: Модификация 2 parent-uuid-A (п. 4.2.к - сгруппирована под тем же родителем!)
  assert.deepEqual(getRowValues(5), [
    'Модификация',
    'var-uuid-3',
    '4600000000003',
    'parent-uuid-A',
    2100,
    3,
  ]);

  // Строка 6: Родитель для модификации parent-uuid-B
  assert.deepEqual(getRowValues(6), [
    'Товар',
    'parent-uuid-B',
    null,
    null,
    null,
    null,
  ]);

  // Строка 7: Модификация parent-uuid-B
  assert.deepEqual(getRowValues(7), [
    'Модификация',
    'var-uuid-5',
    '4600000000005',
    'parent-uuid-B',
    3000,
    1,
  ]);
  // Итого должно быть ровно 7 строк (1 заголовок + 6 строк данных)
  assert.equal(sheet.rowCount, 7);});

// =========================================================================
// ТЕСТЫ ПОЧТОВЫХ ОТВЕТОВ ПРИ ОШИБКАХ
// =========================================================================

test('ответ клиенту текстом при отсутствии товара в МойСклад', async () => {
  const inputBuffer = await createTestInputXlsx([
    { barcode: '9999999999999', price: 100, quantity: 1 },
  ]);

  const mockFetchEmpty = async () => ({
    ok: true,
    json: async () => ({ rows: [] }),
  });

  const mailObj = createMockMail([{ filename: 'order.xlsx', content: inputBuffer }]);
  const reply = await prepareReply(mailObj, { address: 'bot@example.com', moysklad: { apiUrl: 'https://test', token: 'tok' } }, mockFetchEmpty);

  assert.equal(reply.attachments.length, 0);
  assert.match(
    reply.text,
    /Товар с ш\/к 9999999999999 не найден в МойСклад\. Пожалуйста, проверьте правильность ш\/к в таблице и в МойСклад и повторите попытку/,
  );
});

test('ответ клиенту текстом при недоступности API МойСклад', async () => {
  const inputBuffer = await createTestInputXlsx([
    { barcode: '4600000000001', price: 100, quantity: 1 },
  ]);

  const mockFetchError = async () => ({
    ok: false,
    status: 503,
    statusText: 'Service Unavailable',
  });

  const mailObj = createMockMail([{ filename: 'order.xlsx', content: inputBuffer }]);
  const reply = await prepareReply(mailObj, { address: 'bot@example.com', moysklad: { apiUrl: 'https://test', token: 'tok' } }, mockFetchError);

  assert.equal(reply.attachments.length, 0);
  assert.match(reply.text, /API МойСклад временно недоступен\. Пожалуйста, повторите попытку позже/);
});

// =========================================================================
// ПРОВЕРКА ПОЧТОВЫХ МЕХАНИЗМОВ И IMAP/SMTP
// =========================================================================

test('защита от циклов автоматических ответов', () => {
  assert.equal(shouldSkip(createMockMail(), 'sender@example.org'), true);
  const automatic = createMockMail();
  automatic.headers.set('auto-submitted', 'auto-replied');
  assert.equal(shouldSkip(automatic, 'own@example.org'), true);
  const bounce = createMockMail();
  bounce.headers.set('content-type', { value: 'multipart/report' });
  assert.equal(shouldSkip(bounce, 'own@example.org'), true);
});

test('фильтрация по белому списку разрешенных отправителей allowedSenders', () => {
  const allowed = ['vip@client.ru', 'partner@firm.org'];

  const mailFromAllowed = {
    from: { value: [{ address: 'VIP@Client.RU' }] }, // проверка toLowerCase
    headers: new Map(),
  };
  assert.equal(shouldSkip(mailFromAllowed, 'own@example.org', allowed), false);

  const mailFromStranger = {
    from: { value: [{ address: 'stranger@spam.com' }] },
    headers: new Map(),
  };
  assert.equal(shouldSkip(mailFromStranger, 'own@example.org', allowed), true);

  // Если список не задан - письмо не фильтруется
  assert.equal(shouldSkip(mailFromStranger, 'own@example.org', undefined), false);
});
test('первый запуск пропускает старые UID, перезапуск сохраняет границу', async () => {
  const store = memoryStore();
  await store.initialize(123n, 51);
  assert.equal(store.value.lastUid, 50);
  await store.checkpoint(52);
  await store.initialize(123n, 60);
  assert.equal(store.value.lastUid, 52);
  await assert.rejects(store.initialize(124n, 60), StateError);
});

test('новые письма обрабатываются по порядку, старые UID исключаются', async () => {
  const store = memoryStore();
  await store.initialize(123n, 51);
  const fetched = [];
  let released = false;
  const client = {
    mailbox: { uidValidity: 123n, uidNext: 54 },
    getMailboxLock: async () => ({ release: () => { released = true; } }),
    search: async () => [53, 50, 51],
    fetchOne: async (uid) => {
      fetched.push(uid);
      return { source: Buffer.from('From: sender@example.org\r\nSubject: No attachment\r\n\r\nHello') };
    },
  };
  await drainMailbox(client, store, { mailbox: 'INBOX', address: 'own@example.org' }, {}, new AbortController().signal);
  assert.deepEqual(fetched, [51, 53]);
  assert.equal(store.value.lastUid, 53);
  assert.equal(released, true);
});

test('ошибка получения не повторяется и не продвигает состояние', async () => {
  const store = memoryStore();
  await store.initialize(123n, 51);
  let attempts = 0;
  let released = false;
  const client = {
    mailbox: { uidValidity: 123n, uidNext: 52 },
    getMailboxLock: async () => ({ release: () => { released = true; } }),
    search: async () => [51],
    fetchOne: async () => { attempts++; throw new Error('offline'); },
  };
  await assert.rejects(drainMailbox(client, store, { mailbox: 'INBOX' }, {}, new AbortController().signal), /offline/);
  assert.equal(attempts, 1);
  assert.equal(store.value.lastUid, 50);
  assert.equal(released, true);
});

test('ошибка SMTP не повторяется', async () => {
  const source = Buffer.from([
    'From: sender@example.org',
    'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="test-boundary"',
    '',
    '--test-boundary',
    'Content-Type: application/octet-stream',
    'Content-Disposition: attachment; filename="bad.xlsx"',
    '',
    'broken',
    '--test-boundary--',
  ].join('\r\n'));
  let attempts = 0;
  await assert.rejects(processMessage(source, {
    config: { address: 'own@example.org', moysklad: { apiUrl: 'https://test', token: 'tok' } },
    identity: 'test',
    transport: { sendMail: async () => { attempts++; throw new Error('offline'); } },
  }), /offline/);
  assert.equal(attempts, 1);
});
