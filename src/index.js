import nodemailer from 'nodemailer';
import { loadConfig } from './config.js';
import { StateError, StateStore } from './state.js';
import { runService } from './service.js';

const controller = new AbortController();
process.once('SIGINT', () => controller.abort());
process.once('SIGTERM', () => controller.abort());

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
  console.log('Сервис запущен. Автоматические повторы отключены.');
  try {
    await runService(config, store, transport, controller.signal);
  } catch (error) {
    if (config.adminAddress && !controller.signal.aborted) {
      try {
        await transport.sendMail({
          from: { address: config.address },
          to: { address: config.adminAddress },
          subject: 'Сервис обработки XLSX остановлен',
          text: 'При обработке почты произошла ошибка. Проверьте журнал сервиса. '
            + 'После устранения причины запустите сервис вручную.',
          headers: { 'Auto-Submitted': 'auto-generated', 'X-Auto-Response-Suppress': 'All' },
        });
      } catch {
        console.error('Не удалось уведомить администратора. Повтор не выполняется.');
      }
    }
    throw error;
  }
} catch (error) {
  // Не выводим сообщения сторонних библиотек: они могут содержать данные письма или авторизации.
  const code = /^[A-Z0-9_]{1,40}$/.test(String(error.code)) ? String(error.code) : undefined;
  const smtpCode = Number.isInteger(error.responseCode) ? error.responseCode : undefined;
  console.error(JSON.stringify({
    event: 'service_stopped',
    code,
    smtpCode,
    reason: error instanceof StateError ? error.message : 'Ошибка конфигурации, соединения или обработки',
  }));
  process.exitCode = 1;
} finally {
  transport?.close();
}
