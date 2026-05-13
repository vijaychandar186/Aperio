"use strict";

class AutoSave {
  /**
   * @param {object} opts
   * @param {FileDatabase}       opts.db
   * @param {() => number|null}  opts.getFileId
   * @param {() => object}       opts.getWorkbook
   * @param {() => object}       opts.getHighlights
   * @param {HTMLElement}        opts.indicator   - element to show save status
   */
  constructor({ db, getFileId, getWorkbook, getHighlights, indicator }) {
    this._db           = db;
    this._getFileId    = getFileId;
    this._getWb        = getWorkbook;
    this._getHl        = getHighlights;
    this._indicator    = indicator;
    this._timer        = null;
    this._fadeTimer    = null;
    this._DELAY        = 1500; // ms debounce
    this._dirty        = false;
  }

  /** Schedule a save (debounced). Call after every mutation. */
  schedule() {
    if (!this._getFileId()) return;
    this._dirty = true;
    this._setStatus("unsaved");
    clearTimeout(this._timer);
    this._timer = setTimeout(() => this._flush(), this._DELAY);
  }

  /** True when there are unsaved changes. */
  isDirty() { return this._dirty; }

  /** Cancel debounce and immediately save. Returns the save promise. */
  flushNow() {
    clearTimeout(this._timer);
    this._timer = null;
    return this._flush();
  }

  /** Cancel any pending save. */
  cancel() { clearTimeout(this._timer); this._dirty = false; }

  // ── private ───────────────────────────────────────────────────────────────
  async _flush() {
    const id = this._getFileId();
    const wb = this._getWb();
    if (!id || !wb) return;

    this._setStatus("saving…");
    try {
      const raw = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
      const buf = new Uint8Array(raw).buffer;
      await this._db.update(id, { data: buf, size: buf.byteLength, highlights: JSON.stringify(this._getHl()) });
      this._dirty = false;
      this._setStatus("saved");
      clearTimeout(this._fadeTimer);
      this._fadeTimer = setTimeout(() => this._setStatus(""), 2000);
    } catch (err) {
      console.warn("auto-save failed:", err);
      this._setStatus("save failed");
    }
  }

  _setStatus(text) {
    if (this._indicator) this._indicator.textContent = text;
  }
}
