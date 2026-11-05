import { ImapFlow } from 'imapflow';
import { processMessage } from './mail.js';

export async function drainMailbox(client, store, config, transport, signal) {
  const lock = await client.getMailboxLock(config.mailbox);
  try {
    await store.initialize(client.mailbox.uidValidity, client.mailbox.uidNext);
    const lastUid = store.value.lastUid;
    const found = await client.search({ uid: `${lastUid + 1}:*` }, { uid: true });
    const uids = (found || []).filter((uid) => uid > lastUid).sort((a, b) => a - b);
    for (const uid of uids) {
      if (signal.aborted) return;
      const message = await client.fetchOne(uid, { source: true }, { uid: true });
      if (message) {
        const identity = JSON.stringify([
          store.account, config.mailbox, store.value.uidValidity, uid,
        ]);
        await processMessage(message.source, { config, transport, identity });
      }
      await store.checkpoint(uid);
    }
  } finally {
    lock.release();
  }
}

export async function runService(config, store, transport, signal) {
  const client = new ImapFlow(config.imap);
  let pending = true;
  let failure;
  let wake;
  const notify = () => {
    pending = true;
    wake?.();
  };
  const fail = (error) => {
    failure ??= error;
    notify();
  };
  const stop = () => {
    notify();
    client.close();
  };
  client.on('exists', notify);
  client.on('close', () => fail(new Error('Соединение IMAP закрыто')));
  client.on('error', fail);
  signal.addEventListener('abort', stop, { once: true });
  try {
    if (signal.aborted) return;
    await client.connect();
    await client.mailboxOpen(config.mailbox);
    await store.initialize(client.mailbox.uidValidity, client.mailbox.uidNext);
    while (!signal.aborted) {
      if (failure) throw failure;
      pending = false;
      await drainMailbox(client, store, config, transport, signal);
      if (pending || signal.aborted) continue;
      await new Promise((resolve) => {
        wake = resolve;
        if (pending || signal.aborted) resolve();
      });
      wake = undefined;
    }
  } catch (error) {
    if (!signal.aborted) throw error;
  } finally {
    signal.removeEventListener('abort', stop);
    client.close();
  }
}
