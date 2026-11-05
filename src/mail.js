/**
 * Модуль обработки входящих писем и формирования ответов.
 *
 * Отвечает за:
 * 1. Фильтрацию служебных, циклических и спам-сообщений (shouldSkip).
 * 2. Парсинг вложений XLSX и запуск процесса их трансформации с интеграцией МойСклад.
 * 3. Формирование текстов ответов:
 *    - Успешная обработка (файлы во вложении);
 *    - Ошибка поиска штрихкода (BarcodeNotFoundError);
 *    - Недоступность API МойСклад (MoySkladUnavailableError);
 *    - Некорректный формат таблицы (InvalidTableError).
 * 4. Отправку сформированного ответа через SMTP.
 */

import { createHash } from 'node:crypto';
import { simpleParser } from 'mailparser';
import { InvalidTableError, transformXlsx } from './xlsx.js';
import { BarcodeNotFoundError, MoySkladUnavailableError } from './moysklad.js';

/**
 * Проверяет, следует ли пропустить обработку входящего письма во избежание почтовых петель,
 * спама и ответов на служебные рассылки.
 *
 * @param {object} mail - Распарсенное входящее письмо от mailparser.
 * @param {string} ownAddress - Собственный email-адрес бота.
 * @returns {boolean} true, если письмо требуется пропустить без ответа.
 */
export function shouldSkip(mail, ownAddress) {
  const from = mail.from?.value ?? [];
  if (from.length !== 1 || !from[0].address) return true;
  const sender = from[0].address.toLowerCase();
  if (sender === ownAddress.toLowerCase()) return true;
  if (/^(mailer-daemon|postmaster)@/i.test(sender)) return true;
  const automatic = mail.headers.get('auto-submitted');
  if (automatic && String(automatic).toLowerCase() !== 'no') return true;
  if (mail.headers.has('x-auto-response-suppress')) return true;
  if (mail.headers.has('list-id')) return true;
  if (/^(bulk|junk|list)$/i.test(String(mail.headers.get('precedence') ?? ''))) return true;
  const contentType = mail.headers.get('content-type');
  if (contentType?.value === 'multipart/report') return true;
  if (mail.headers.has('return-path')) {
    const returnPath = mail.headers.get('return-path');
    if (returnPath === '<>' || returnPath?.value?.length === 0) return true;
  }
  return false;
}

/**
 * Формирует структуру ответного письма на основе прикрепленных XLSX-файлов.
 *
 * Обрабатывает результаты трансформации файлов и возможные ошибки:
 * - Если товар не найден по ш/к -> отправляет точный текст с штрихкодом;
 * - Если МойСклад недоступен -> отправляет уведомление о временной недоступности;
 * - Если таблица некорректна -> отправляет сообщение об ошибке формата.
 *
 * @param {object} mail - Распарсенное письмо.
 * @param {object} config - Общая конфигурация приложения ({ address, moysklad, ... }).
 * @param {typeof fetch} [customFetch=fetch] - Функция fetch (для тестирования).
 * @returns {Promise<object|null>} Объект параметров письма для Nodemailer или null при пропуске.
 */
export async function prepareReply(mail, config, customFetch = fetch) {
  const ownAddress = typeof config === 'string' ? config : config.address;
  const msConfig = typeof config === 'object' ? config.moysklad : undefined;

  if (shouldSkip(mail, ownAddress)) return null;
  const attachments = [];
  const errors = [];

  for (const attachment of mail.attachments ?? []) {
    if (!/\.xlsx$/i.test(attachment.filename ?? '')) continue;
    try {
      const processedBuffer = await transformXlsx(attachment.content, msConfig, customFetch);
      attachments.push({
        filename: attachment.filename,
        content: processedBuffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    } catch (error) {
      if (error instanceof BarcodeNotFoundError) {
        // Требование п. 2 Рисков:
        // "Товар с ш/к ... не найден в МойСклад. Пожалуйста, проверьте правильность ш/к в таблице и в МойСклад и повторите попытку"
        errors.push(error.message);
      } else if (error instanceof MoySkladUnavailableError) {
        // Требование п. 3 Рисков:
        // "API МойСклад временно недоступен. Пожалуйста, повторите попытку позже"
        errors.push(error.message);
      } else if (error instanceof InvalidTableError) {
        const safeName = attachment.filename.replace(/[\r\n]/g, ' ');
        errors.push(`Таблица некорректна: ${safeName}`);
      } else {
        throw error;
      }
    }
  }

  // Если во входящем письме вообще не было XLSX файлов и ошибок
  if (!attachments.length && !errors.length) return null;

  const textLines = [];
  if (attachments.length) {
    textLines.push('Готовые файлы во вложении.');
  }
  if (errors.length) {
    textLines.push(...errors);
  }

  return {
    from: { address: ownAddress },
    to: { address: mail.from.value[0].address },
    subject: `Re: ${(mail.subject || 'Обработка XLSX').replace(/[\r\n]/g, ' ')}`,
    text: textLines.join('\n\n'),
    attachments,
    inReplyTo: mail.messageId,
    references: mail.messageId,
    headers: {
      'Auto-Submitted': 'auto-replied',
      'X-Auto-Response-Suppress': 'All',
    },
  };
}

/**
 * Ошибка постоянного сбоя SMTP-доставки (5xx), при которой повторные попытки не требуются.
 */
export class PermanentDeliveryError extends Error {
  constructor(code) {
    super('SMTP отклонил отправку ответа');
    this.name = 'PermanentDeliveryError';
    this.code = code;
  }
}

/**
 * Главная точка входа для обработки отдельного сырого MIME-сообщения, полученного из IMAP.
 *
 * @param {Buffer|string} source - Сырое тело письма из IMAP.
 * @param {object} context - Контекст выполнения.
 * @param {object} context.config - Конфигурация сервиса.
 * @param {object} context.transport - Nodemailer транспорт.
 * @param {string} context.identity - Уникальный идентификатор письма (UID/Message-ID).
 * @param {typeof fetch} [context.customFetch=fetch] - Опциональный fetch для моков.
 * @returns {Promise<'skipped'|'sent'>} Статус обработки письма.
 */
export async function processMessage(source, { config, transport, identity, customFetch = fetch }) {
  const mail = await simpleParser(source, {
    skipHtmlToText: true,
    skipTextToHtml: true,
    skipImageLinks: true,
  });
  const reply = await prepareReply(mail, config, customFetch);
  if (!reply) return 'skipped';
  const digest = createHash('sha256').update(identity).digest('hex');
  reply.messageId = `<xlsx-${digest}@${config.address.split('@')[1]}>`;
  try {
    await transport.sendMail(reply);
  } catch (error) {
    if (error.responseCode >= 500 && error.responseCode < 600) {
      throw new PermanentDeliveryError(error.responseCode);
    }
    throw error;
  }
  return 'sent';
}
