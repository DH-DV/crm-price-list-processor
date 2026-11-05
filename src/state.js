import { mkdir, readFile, writeFile, rename } from 'node:fs/promises';
import { dirname } from 'node:path';

export class StateError extends Error {}

export class StateStore {
  constructor(path, account, mailbox) {
    this.path = path;
    this.account = account;
    this.mailbox = mailbox;
    this.value = null;
  }

  async load() {
    try {
      this.value = JSON.parse(await readFile(this.path, 'utf8'));
    } catch (error) {
      if (error.code === 'ENOENT') return;
      throw new StateError('Не удалось прочитать состояние; автоматический сброс запрещён');
    }
    const value = this.value;
    if (value.version !== 1 || value.account !== this.account
      || value.mailbox !== this.mailbox || !/^\d+$/.test(value.uidValidity)
      || !Number.isSafeInteger(value.lastUid) || value.lastUid < 0) {
      throw new StateError('Состояние повреждено или относится к другому ящику');
    }
  }

  async initialize(uidValidity, uidNext) {
    const validity = String(uidValidity);
    if (this.value) {
      if (this.value.uidValidity !== validity) {
        throw new StateError('UIDVALIDITY изменился; требуется проверка состояния оператором');
      }
      return;
    }
    await this.save({
      version: 1,
      account: this.account,
      mailbox: this.mailbox,
      uidValidity: validity,
      lastUid: Math.max(0, uidNext - 1),
    });
  }

  async checkpoint(uid) {
    if (uid <= this.value.lastUid) return;
    await this.save({ ...this.value, lastUid: uid });
  }

  async save(value) {
    try {
      await mkdir(dirname(this.path), { recursive: true, mode: 0o700 });
      await writeFile(`${this.path}.tmp`, `${JSON.stringify(value)}\n`, {
        mode: 0o600,
        flush: true,
      });
      await rename(`${this.path}.tmp`, this.path);
      this.value = value;
    } catch {
      throw new StateError('Не удалось сохранить состояние');
    }
  }
}
