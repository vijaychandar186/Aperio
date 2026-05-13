"use strict";

/**
 * Generic IndexedDB wrapper.
 * Each viewer instantiates its own named database.
 */
class FileDatabase {
  constructor(dbName, storeName = "files") {
    this._dbName    = dbName;
    this._storeName = storeName;
    this._db        = null;
  }

  _open() {
    if (this._db) return Promise.resolve(this._db);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(this._dbName, 1);
      req.onupgradeneeded = () =>
        req.result.createObjectStore(this._storeName, { keyPath: "id", autoIncrement: true });
      req.onsuccess = () => { this._db = req.result; resolve(this._db); };
      req.onerror  = () => reject(req.error);
    });
  }

  _req(r) {
    return new Promise((res, rej) => {
      r.onsuccess = () => res(r.result);
      r.onerror   = () => rej(r.error);
    });
  }

  async _store(mode) {
    const db = await this._open();
    return db.transaction(this._storeName, mode).objectStore(this._storeName);
  }

  /** Save a new record. Returns the auto-generated id. */
  async save(record) {
    const s = await this._store("readwrite");
    return this._req(s.add({ ...record, savedAt: Date.now() }));
  }

  /** List all records without the raw data blob (for display). */
  async list() {
    const s   = await this._store("readonly");
    const all = await this._req(s.getAll());
    return all
      .map(({ data, ...meta }) => meta)
      .sort((a, b) => b.savedAt - a.savedAt);
  }

  /** Get full record including data blob. */
  async get(id) {
    const s = await this._store("readonly");
    return this._req(s.get(id));
  }

  /** Patch an existing record in-place (used for auto-save). */
  async update(id, patch) {
    const db = await this._open();
    return new Promise((resolve, reject) => {
      const tx    = db.transaction(this._storeName, "readwrite");
      const store = tx.objectStore(this._storeName);
      const get   = store.get(id);
      get.onsuccess = () => {
        const rec = get.result;
        if (!rec) { resolve(); return; }
        const put = store.put({ ...rec, ...patch, savedAt: Date.now() });
        put.onsuccess = () => resolve();
        put.onerror   = () => reject(put.error);
      };
      get.onerror = () => reject(get.error);
    });
  }

  /** Remove a record. */
  async delete(id) {
    const s = await this._store("readwrite");
    return this._req(s.delete(id));
  }
}
