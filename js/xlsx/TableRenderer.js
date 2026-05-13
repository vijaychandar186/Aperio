"use strict";

const MAX_ROWS = 5000;
const MAX_COLS = 500;

class TableRenderer {
  /**
   * @param {SelectionManager} selection
   * @param {HighlightManager} highlights
   */
  constructor(selection, highlights) {
    this._sel = selection;
    this._hl  = highlights;
  }

  /**
   * Build and return the scrollable table container for a worksheet.
   * Also registers every data cell with the SelectionManager.
   *
   * @param {object} ws          - SheetJS worksheet
   * @param {string} sheetName
   * @param {object} opts
   * @param {(scroller: HTMLElement) => void} opts.onMousedown
   * @param {(scroller: HTMLElement) => void} opts.onMousemove
   * @param {(scroller: HTMLElement) => void} opts.onDblclick
   * @returns {HTMLElement} scroller div
   */
  build(ws, sheetName, { onMousedown, onMousemove, onDblclick } = {}) {
    this._sel.reset();

    if (!ws["!ref"]) {
      return Object.assign(document.createElement("div"), {
        className: "p-10 text-gray-400 text-sm text-center",
        textContent: "empty sheet",
      });
    }

    const range   = XLSX.utils.decode_range(ws["!ref"]);
    const endRow  = Math.min(range.e.r, MAX_ROWS - 1);
    const endCol  = Math.min(range.e.c, MAX_COLS - 1);
    const sheetHl = this._hl.getAll()[sheetName] || {};

    // Build merge lookup
    const skipSet  = new Set();
    const mergeMap = {};
    for (const m of (ws["!merges"] || [])) {
      const er = Math.min(m.e.r, endRow), ec = Math.min(m.e.c, endCol);
      if (m.s.r > endRow || m.s.c > endCol) continue;
      mergeMap[`${m.s.r}_${m.s.c}`] = { rs: er - m.s.r + 1, cs: ec - m.s.c + 1 };
      for (let r = m.s.r; r <= er; r++)
        for (let c = m.s.c; c <= ec; c++)
          if (r !== m.s.r || c !== m.s.c) skipSet.add(`${r}_${c}`);
    }

    const colWidths = ws["!cols"] || [];

    // ── DOM ────────────────────────────────────────────────────────────────
    const scroller = document.createElement("div");
    scroller.style.cssText = "width:100%;height:100%;overflow:auto;background:#f3f4f6;position:relative;";

    const tableWrap = document.createElement("div");
    tableWrap.style.cssText = "display:inline-block;min-width:100%;background:#fff;border-right:1px solid #e0e0e0;border-bottom:1px solid #e0e0e0;";

    const table = document.createElement("table");
    table.style.cssText = "border-collapse:separate;border-spacing:0;font-size:12px;font-family:ui-monospace,SFMono-Regular,monospace;white-space:nowrap;";

    // Header row
    const thead = document.createElement("thead");
    const hRow  = document.createElement("tr");
    const corner = document.createElement("th");
    corner.style.cssText = "position:sticky;left:0;top:0;z-index:3;background:#f0f0f0;padding:3px 8px;min-width:44px;text-align:center;font-weight:normal;color:#aaa;font-size:11px;border-right:2px solid #ccc;border-bottom:2px solid #ccc;border-top:1px solid #ddd;";
    hRow.appendChild(corner);

    for (let c = range.s.c; c <= endCol; c++) {
      const th  = document.createElement("th");
      const wch = colWidths[c]?.wch;
      th.textContent   = XLSX.utils.encode_col(c);
      th.style.cssText = `position:sticky;top:0;z-index:2;background:#f0f0f0;padding:3px 8px;text-align:center;font-weight:normal;color:#555;font-size:11px;min-width:${wch ? Math.max(60, wch * 7) : 80}px;border-right:1px solid #ddd;border-bottom:2px solid #ccc;border-top:1px solid #ddd;`;
      hRow.appendChild(th);
    }
    thead.appendChild(hRow);
    table.appendChild(thead);

    // Body
    const tbody = document.createElement("tbody");
    for (let r = range.s.r; r <= endRow; r++) {
      const tr  = document.createElement("tr");
      const odd = r % 2 === 1;

      const rn = document.createElement("td");
      rn.textContent   = r + 1;
      rn.style.cssText = `position:sticky;left:0;z-index:1;background:${odd ? "#efefef" : "#f5f5f5"};padding:2px 8px;text-align:right;color:#aaa;font-size:11px;border-right:2px solid #ccc;border-bottom:1px solid #e8e8e8;user-select:none;`;
      tr.appendChild(rn);

      for (let c = range.s.c; c <= endCol; c++) {
        const key = `${r}_${c}`;
        if (skipSet.has(key)) continue;

        const td = document.createElement("td");
        td.dataset.r = String(r);
        td.dataset.c = String(c);

        const minfo = mergeMap[key];
        if (minfo) { if (minfo.rs > 1) td.rowSpan = minfo.rs; if (minfo.cs > 1) td.colSpan = minfo.cs; }

        td.style.background   = sheetHl[key] || (odd ? "#fafafa" : "#fff");
        td.style.padding      = "2px 8px";
        td.style.borderRight  = "1px solid #e8e8e8";
        td.style.borderBottom = "1px solid #e8e8e8";
        td.style.maxWidth     = "300px";
        td.style.overflow     = "hidden";
        td.style.textOverflow = "ellipsis";

        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell) {
          td.textContent = cell.w !== undefined ? cell.w : (cell.v != null ? String(cell.v) : "");
          if (cell.t === "n") td.style.textAlign = "right";
          if (cell.s && !sheetHl[key]) {
            try {
              if (cell.s.font?.bold)      td.style.fontWeight     = "bold";
              if (cell.s.font?.italic)    td.style.fontStyle      = "italic";
              if (cell.s.font?.underline) td.style.textDecoration = "underline";
              const fc = cell.s.font?.color?.rgb;
              if (fc) td.style.color = "#" + fc.slice(-6);
              const bg = cell.s.fill?.fgColor?.rgb;
              if (bg && bg.slice(-6) !== "FFFFFF" && bg !== "00000000") td.style.background = "#" + bg.slice(-6);
              const ha = cell.s.alignment?.horizontal;
              if (ha === "center") td.style.textAlign = "center";
              else if (ha === "right") td.style.textAlign = "right";
              else if (ha === "left")  td.style.textAlign = "left";
            } catch { /* ignore style errors */ }
          }
        }

        this._sel.register(r, c, td);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    tableWrap.appendChild(table);

    if (endRow < range.e.r || endCol < range.e.c) {
      const note = document.createElement("div");
      note.style.cssText = "padding:6px 12px;font-size:11px;color:#92400e;background:#fffbeb;border-top:1px solid #fde68a;position:sticky;left:0;";
      note.textContent = `Showing ${endRow - range.s.r + 1} of ${range.e.r - range.s.r + 1} rows and ${endCol - range.s.c + 1} of ${range.e.c - range.s.c + 1} cols (large sheet truncated).`;
      tableWrap.appendChild(note);
    }

    scroller.appendChild(tableWrap);

    // Attach mouse event delegation
    if (onMousedown) scroller.addEventListener("mousedown", onMousedown);
    if (onMousemove) scroller.addEventListener("mousemove", onMousemove);
    if (onDblclick)  scroller.addEventListener("dblclick",  onDblclick);

    return scroller;
  }
}
