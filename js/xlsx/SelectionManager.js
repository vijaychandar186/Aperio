"use strict";

class SelectionManager {
  constructor() {
    this.anchor   = null;  // { r, c }
    this.focus    = null;  // { r, c }
    this.dragging = false;
    this._cells   = {};    // 'r_c' → <td>
    this._active  = new Set();
    this._onChange = null;
  }

  /** Called whenever anchor/focus changes — use to update cell-ref display etc. */
  onChange(cb) { this._onChange = cb; }

  /** Register a rendered td element so it can be selected. */
  register(r, c, el) { this._cells[`${r}_${c}`] = el; }

  /** Drop all registrations (call before re-rendering a sheet). */
  reset() {
    this._cells  = {};
    this._active = new Set();
    this.anchor  = null;
    this.focus   = null;
  }

  /** Single-cell click select. */
  select(r, c) {
    this.anchor = { r, c };
    this.focus  = { r, c };
    this._refresh();
  }

  /** Extend selection rectangle toward (r, c). */
  extendTo(r, c) {
    this.focus = { r, c };
    this._refresh();
  }

  /**
   * Move anchor (or extend focus) by (dr, dc).
   * Returns the new target cell {r, c} — caller may clamp to sheet bounds first.
   */
  moveBy(dr, dc, extend = false) {
    if (!this.anchor) return null;
    const base = extend ? { ...this.focus } : { ...this.anchor };
    const next = { r: base.r + dr, c: base.c + dc };
    if (extend) { this.focus = next; }
    else        { this.anchor = next; this.focus = next; }
    this._refresh();
    return next;
  }

  /** Clear the whole selection. */
  clear() {
    for (const k of this._active) this._cells[k]?.classList.remove("sel");
    this._active.clear();
    this.anchor = null;
    this.focus  = null;
    this._onChange?.(null);
  }

  /** Current bounding box, or null. */
  getRange() {
    if (!this.anchor) return null;
    const f = this.focus || this.anchor;
    return {
      minR: Math.min(this.anchor.r, f.r), maxR: Math.max(this.anchor.r, f.r),
      minC: Math.min(this.anchor.c, f.c), maxC: Math.max(this.anchor.c, f.c),
    };
  }

  /** Returns the registered td for a cell key, or undefined. */
  el(r, c) { return this._cells[`${r}_${c}`]; }

  // ── private ───────────────────────────────────────────────────────────────
  _refresh() {
    for (const k of this._active) this._cells[k]?.classList.remove("sel");
    this._active.clear();
    const range = this.getRange();
    if (range) {
      for (let r = range.minR; r <= range.maxR; r++) {
        for (let c = range.minC; c <= range.maxC; c++) {
          const k = `${r}_${c}`;
          if (this._cells[k]) { this._cells[k].classList.add("sel"); this._active.add(k); }
        }
      }
    }
    this._onChange?.(this.anchor);
  }
}
