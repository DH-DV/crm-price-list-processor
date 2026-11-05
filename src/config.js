function required(env, name) {
  const value = env[name];
  if (!value) throw new Error(`Не задана переменная ${name}`);
  return value;
}

function port(value, fallback) {
  const result = Number(value ?? fallback);
  if (!Number.isInteger(result) || result < 1 || result > 65535) {
    throw new Error('Некорректный порт');
  }
  return result;
}

export function loadConfig(env = process.env) {
  const address = required(env, 'MAIL_ADDRESS');
  if (!/^[^\s<>@]+@[^\s<>@]+$/.test(address)) {
    throw new Error('Некорректный MAIL_ADDRESS');
  }
  const user = env.MAIL_USER || address;
  const pass = required(env, 'MAIL_PASSWORD');
  return {
    address,
    mailbox: env.IMAP_MAILBOX || 'INBOX',
    stateFile: env.STATE_FILE || './data/state.json',
    adminAddress: env.ADMIN_ADDRESS || undefined,
    moysklad: {
      apiUrl: env.MOYSKLAD_API_URL || 'https://api.moysklad.ru/api/remap/1.2',
      token: required(env, 'MOYSKLAD_TOKEN'),
    },
    imap: {
      host: env.IMAP_HOST || 'imap.yandex.ru',
      port: port(env.IMAP_PORT, 993),
      secure: true,
      auth: { user, pass },
      logger: false,
      maxIdleTime: 240_000,
    },
    smtp: {
      host: env.SMTP_HOST || 'smtp.yandex.ru',
      port: port(env.SMTP_PORT, 465),
      secure: true,
      auth: { user, pass },
      logger: false,
      debug: false,
      disableFileAccess: true,
      disableUrlAccess: true,
    },
  };
}
