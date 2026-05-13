"use strict";

class HighlightManager {
  constructor() {
    this._store = {}; // { sheetName: { 'r_c': '#hex' } }
  }

  /** Apply (or clear) a highlight color to all cells in range. */
  apply(sheetName, range, color, selectionManager, getWorkbook) {
    if (!range || !sheetName) return;
    const sheet = this._store[sheetName] || (this._store[sheetName] = {});
    for (let r = range.minR; r <= range.maxR; r++) {
      for (let c = range.minC; c <= range.maxC; c++) {
        const key = `${r}_${c}`;
        const td  = selectionManager.el(r, c);
        if (!td) continue;
        if (color === null) {
          delete sheet[key];
          td.style.background = HighlightManager._baseBg(r, c, sheetName, getWorkbook);
        } else {
          sheet[key]          = color;
          td.style.background = color;
        }
      }
    }
  }

  /** Restore saved highlights onto DOM cells (call after rendering a sheet). */
  restoreToDOM(sheetName, selectionManager) {
    const sheet = this._store[sheetName];
    if (!sheet) return;
    for (const [key, color] of Object.entries(sheet)) {
      const [r, c] = key.split("_").map(Number);
      const td = selectionManager.el(r, c);
      if (td) td.style.background = color;
    }
  }

  /** Return a copy of the full store (for serialisation). */
  getAll() { return this._store; }

  /** Replace the store from a parsed JSON object (on load). */
  restoreFrom(saved) {
    this._store = {};
    if (saved && typeof saved === "object") Object.assign(this._store, saved);
  }

  /** Clear all highlights. */
  clear() { this._store = {}; }

  // ── private ───────────────────────────────────────────────────────────────
  static _baseBg(r, c, sheetName, getWorkbook) {
    const ws   = getWorkbook()?.Sheets[sheetName];
    const cell = ws?.[XLSX.utils.encode_cell({ r, c })];
    if (cell?.s?.fill?.fgColor?.rgb) {
      const rgb = cell.s.fill.fgColor.rgb;
      if (rgb.slice(-6) !== "FFFFFF" && rgb !== "00000000") return "#" + rgb.slice(-6);
    }
    return r % 2 === 0 ? "#fff" : "#fafafa";
  }
}
