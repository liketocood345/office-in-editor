import { realpath } from 'node:fs/promises';
import path from 'node:path';

// The listener itself is the reservation. Never probe a candidate port and
// release it between discovery and use.
export class DocumentPorts {
  #sessions = new Map();
  #capacity;

  constructor(capacity) {
    if (!Number.isSafeInteger(capacity) || capacity < 1 || capacity > 8) throw new Error('Invalid document capacity');
    this.#capacity = capacity;
  }

  async acquire(document, open) {
    // The lab prepares its fixture inside an existing, validated directory.
    // Resolve the directory so an alternate spelling maps to the same lease.
    const directory = await realpath(path.dirname(document));
    const resolved = path.join(directory, path.basename(document));
    const key = process.platform === 'win32' ? resolved.toLowerCase() : resolved;
    const existing = this.#sessions.get(key);
    if (existing) return existing;
    if (this.#sessions.size >= this.#capacity) throw new Error('Document session limit reached');

    const ready = Promise.resolve().then(() => open(resolved));
    this.#sessions.set(key, ready);
    const release = () => {
      if (this.#sessions.get(key) === ready) this.#sessions.delete(key);
    };
    void ready.then(session => { void session.closed.then(release); }, release);
    return ready;
  }
}
