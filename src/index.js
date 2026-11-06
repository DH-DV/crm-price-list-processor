/**
 * Главная точка входа приложения. *
 * Отвечает за:
 * 1. Загрузку конфигурации и хранилища состояния (state.json).
 * 2. Инициализацию SMTP-транспорта Nodemailer.
 * 3. Организацию надежного цикла работы сервиса (Reconnect Loop):
 *    - При штатных разрывах TCP/TLS сессии почтовым сервером (close, ETIMEDOUT, ECONNRESET)
 *      сервис не завершается, а выдерживает экспоненциальную паузу (1с -> 2с -> 5с -> ... до 60с)
 *      и подключается заново со свежим экземпляром ImapFlow.
 *    - При фатальных ошибках (неверный логин/пароль, ошибки в state.json, синтаксис конфигурации)
 *      сервис немедленно останавливается с уведомлением администратора.
 * 4. Корректную обработку сигналов остановки ОС (SIGINT, SIGTERM).
 */

import nodemailer from 'nodemailer';
import { loadConfig } from './config.js';
import { StateError, StateStore } from './state.js';
import { runService } from './service.js';

const controller = new AbortController();
process.once('SIGINT', () => controller.abort());
process.once('SIGTERM', () => controller.abort());

/**
 * Проверяет, является ли ошибка фатальной (требующей полной остановки процесса).
 *
 * @param {Error} error - Объект ошибки.
 * @returns {boolean} true, если ошибку невозможно решить простым переподключением.
 */
function isFatalError(error) {
  if (!error) return false;
  // Ошибка состояния или контрольной суммы ящика
  if (error instanceof StateError) return true;

  const msg = String(error.message || '').toUpperCase();
  const code = String(error.code || '').toUpperCase();
  const response = String(error.responseText || '').toUpperCase();

  // Ошибки аутентификации (неверный пароль или логин)
  if (
    msg.includes('AUTHENTICATIONFAILED') ||
    msg.includes('INVALID CREDENTIALS') ||
    response.includes('AUTHENTICATIONFAILED') ||
    response.includes('NO [AUTHENTICATIONFAILED]') ||
    code === 'EAUTH'
  ) {
    return true;
  }

  // Ошибки обязательной конфигурации
  if (msg.includes('НЕ ЗАДАНА ПЕРЕМЕННАЯ') || msg.includes('НЕКОРРЕКТНЫЙ')) {
    return true;
  }

  return false;
}

let transport;
try {
  const config = loadConfig();
  const store = new StateStore(
    config.stateFile,
    `${config.imap.host}/${config.imap.auth.user}`,
    config.mailbox,
  );
  await store.load();
  transport = nodemailer.createTransport(config.smtp);

  console.log(`[INFO] [${new Date().toISOString()}] Сервис запущен. Автоматический reconnect включен.`);

  // Паузы переподключения в секундах: 1 -> 2 -> 5 -> 10 -> 15 (максимум 15 секунд)
  const backoffDelays = [1, 2, 5, 10, 15];  let failureCount = 0;

  // Основной цикл поддержания постоянного соединения
  while (!controller.signal.aborted) {
    try {
      await runService(config, store, transport, controller.signal);
      // Если runService завершился успешно без ошибок (например, при abort)
      failureCount = 0;
    } catch (error) {
      if (controller.signal.aborted) {
        break;
      }

      // Если ошибка фатальная (пароль, настройки) — останавливаем сервис
      if (isFatalError(error)) {
        console.error(`[FATAL] [${new Date().toISOString()}] Фатальная ошибка: ${error.message}`);
        throw error;
      }

      // Если сетевой сбой / обрыв со стороны Яндекса — выполняем паузу и переподключаемся
      const delay = backoffDelays[Math.min(failureCount, backoffDelays.length - 1)];
      failureCount++;

      console.warn(
        `[WARN] [${new Date().toISOString()}] Соединение с IMAP разорвано (${error.message || 'сетевой сбой'}). Повторное подключение через ${delay} сек... (попытка ${failureCount})`,
      );

      // Ждем паузу с возможностью мгновенной отмены по сигналу остановки
      await new Promise((resolve) => {
        const timer = setTimeout(resolve, delay * 1000);
        controller.signal.addEventListener(
          'abort',
          () => {
            clearTimeout(timer);
            resolve();
          },
          { once: true },
        );
      });
    }
  }
} catch (error) {
  if (transport && !controller.signal.aborted) {
    try {
      const config = loadConfig();
      if (config.adminAddress) {
        await transport.sendMail({
          from: { address: config.address },
          to: { address: config.adminAddress },
          subject: 'Сервис обработки XLSX остановлен',
          text: `При обработке почты произошла фатальная ошибка: ${error.message}. `
            + 'Проверьте журнал сервиса и перезапустите после устранения причины.',
          headers: { 'Auto-Submitted': 'auto-generated', 'X-Auto-Response-Suppress': 'All' },
        });
      }
    } catch {
      console.error('Не удалось уведомить администратора.');
    }
  }

  const code = /^[A-Z0-9_]{1,40}$/.test(String(error.code)) ? String(error.code) : undefined;
  const smtpCode = Number.isInteger(error.responseCode) ? error.responseCode : undefined;
  console.error(JSON.stringify({
    event: 'service_stopped',
    code,
    smtpCode,
    reason: error instanceof StateError ? error.message : (error.message || 'Ошибка конфигурации, соединения или обработки'),
  }));
  process.exitCode = 1;
} finally {
  transport?.close();
}
