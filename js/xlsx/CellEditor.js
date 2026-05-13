"use strict";

class CellEditor {
  /**
   * @param {object} opts
   * @param {() => object}  opts.getWorkbook   - returns current XLSX workbook
   * @param {() => string}  opts.getSheet      - returns current sheet name
   * @param {SelectionManager} opts.selection
   * @param {(r:number, c:number, val:string) => void} opts.onCommit
   * @param {(dir: 'down'|'right'|'left') => void}    opts.onMove
   */
  constructor({ getWorkbook, getSheet, selection, onCommit, onMove }) {
    this._getWb  = getWorkbook;
    this._getSh  = getSheet;
    this._sel    = selection;
    this._onCommit = onCommit;
    this._onMove   = onMove;
    this._editing  = null;  // { r, c, originalValue }
    this._input    = null;  // <input> element
    this._scroller = null;  // scrollable container reference
  }

  setScroller(el) { this._scroller = el; }

  isEditing() { return this._editing !== null; }

  /** Start editing cell (r, c). Optionally pre-fill with initialChar. */
  start(r, c, initialChar = null) {
    this.commit();
    const td = this._sel.el(r, c);
    if (!td || !this._scroller) return;

    this._sel.select(r, c);
    td.scrollIntoView({ block: "nearest", inline: "nearest" });

    const ws      = this._getWb()?.Sheets[this._getSh()];
    const cell    = ws?.[XLSX.utils.encode_cell({ r, c })];
    const curVal  = cell ? (cell.v != null ? String(cell.v) : "") : "";

    this._editing = { r, c, originalValue: curVal };

    const tdRect = td.getBoundingClientRect();
    const sRect  = this._scroller.getBoundingClientRect();

    const inp = document.createElement("input");
    inp.id    = "cell-editor";
    inp.value = initialChar !== null ? initialChar : curVal;
    inp.style.top       = (tdRect.top  - sRect.top  + this._scroller.scrollTop)  + "px";
    inp.style.left      = (tdRect.left - sRect.left + this._scroller.scrollLeft) + "px";
    inp.style.width     = Math.max(td.offsetWidth, 120) + "px";
    inp.style.height    = td.offsetHeight + "px";
    inp.style.textAlign = td.style.textAlign || "left";

    this._scroller.appendChild(inp);
    this._input = inp;
    inp.focus();
    if (initialChar !== null) inp.setSelectionRange(inp.value.length, inp.value.length);
    else inp.select();

    inp.addEventListener("keydown", e => {
      if      (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); this.commit(); this._onMove?.("down"); }
      else if (e.key === "Tab")                  { e.preventDefault(); this.commit(); this._onMove?.(e.shiftKey ? "left" : "right"); }
      else if (e.key === "Escape")               { e.preventDefault(); this.cancel(); }
      e.stopPropagation();
    });

    inp.addEventListener("blur", () => {
      setTimeout(() => { if (this.isEditing()) this.commit(); }, 80);
    });
  }

  /** Commit the current edit value to the workbook. */
  commit() {
    if (!this._editing) return;
    const { r, c }  = this._editing;
    const inp        = this._input;
    if (!inp) { this._editing = null; return; }

    const newVal = inp.value;
    inp.remove();
    this._input   = null;
    this._editing = null;

    const ws   = this._getWb()?.Sheets[this._getSh()];
    const td   = this._sel.el(r, c);
    if (!ws) return;

    const addr = XLSX.utils.encode_cell({ r, c });

    if (newVal === "") {
      if (ws[addr]) { ws[addr].v = undefined; ws[addr].w = ""; ws[addr].t = "z"; }
      if (td) td.textContent = "";
    } else {
      const num   = Number(newVal);
      const isNum = !isNaN(num) && newVal.trim() !== "";
      ws[addr]    = ws[addr] || {};
      if (isNum) {
        ws[addr].t = "n"; ws[addr].v = num; ws[addr].w = newVal;
        if (td) { td.textContent = newVal; td.style.textAlign = "right"; }
      } else {
        ws[addr].t = "s"; ws[addr].v = newVal; ws[addr].w = newVal;
        if (td) td.textContent = newVal;
      }
      this._expandRef(ws, r, c);
    }

    this._onCommit?.(r, c, newVal);
  }

  /** Cancel without saving. */
  cancel() {
    if (!this._editing) return;
    this._input?.remove();
    this._input   = null;
    this._editing = null;
  }

  // ── private ───────────────────────────────────────────────────────────────
  _expandRef(ws, r, c) {
    if (!ws["!ref"]) {
      ws["!ref"] = XLSX.utils.encode_cell({ r, c }) + ":" + XLSX.utils.encode_cell({ r, c });
      return;
    }
    const range = XLSX.utils.decode_range(ws["!ref"]);
    range.s.r = Math.min(range.s.r, r); range.s.c = Math.min(range.s.c, c);
    range.e.r = Math.max(range.e.r, r); range.e.c = Math.max(range.e.c, c);
    ws["!ref"] = XLSX.utils.encode_range(range);
  }
}
