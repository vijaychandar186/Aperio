"use strict";

// Depends (global): XLSX, FileDatabase, SelectionManager, HighlightManager,
//                   TableRenderer, CellEditor, AutoSave

const HIGHLIGHT_COLORS = [
  { hex: "#fef08a", name: "yellow" },
  { hex: "#fed7aa", name: "orange" },
  { hex: "#fecaca", name: "red"    },
  { hex: "#fbcfe8", name: "pink"   },
  { hex: "#bbf7d0", name: "green"  },
  { hex: "#99f6e4", name: "teal"   },
  { hex: "#bfdbfe", name: "blue"   },
  { hex: "#ddd6fe", name: "purple" },
];

class XlsxApp {
  constructor() {
    this._db     = new FileDatabase("xlsx-viewer", "files");
    this._sel    = new SelectionManager();
    this._hl     = new HighlightManager();
    this._tr     = new TableRenderer(this._sel, this._hl);
    this._wb     = null;
    this._sheet  = null;
    this._fileId = null;

    this._els = {
      fileInput:      document.getElementById("fileInput"),
      status:         document.getElementById("status"),
      stage:          document.getElementById("stage"),
      sidebar:        document.getElementById("sidebar"),
      library:        document.getElementById("library"),
      sheetList:      document.getElementById("sheetList"),
      menu:           document.getElementById("menuBtn"),
      swatches:       document.getElementById("swatches"),
      clearHighlight: document.getElementById("clearHighlight"),
      cellRef:        document.getElementById("cellRef"),
      saveIndicator:  document.getElementById("saveIndicator"),
      downloadBtn:    document.getElementById("downloadBtn"),
    };

    this._editor = new CellEditor({
      getWorkbook: () => this._wb,
      getSheet:    () => this._sheet,
      selection:   this._sel,
      onCommit:    () => this._save.schedule(),
      onMove:      dir => {
        if      (dir === "down")  this._moveBy(1,  0);
        else if (dir === "right") this._moveBy(0,  1);
        else if (dir === "left")  this._moveBy(0, -1);
      },
    });

    this._save = new AutoSave({
      db:            this._db,
      getFileId:     () => this._fileId,
      getWorkbook:   () => this._wb,
      getHighlights: () => this._hl.getAll(),
      indicator:     this._els.saveIndicator,
    });

    this._sel.onChange(anchor => {
      this._els.cellRef.textContent = anchor
        ? XLSX.utils.encode_cell({ r: anchor.r, c: anchor.c })
        : "";
    });

    this._init();
  }

  _init() {
    if (window.matchMedia("(max-width: 767px)").matches) {
      this._els.sidebar.dataset.collapsed = "true";
    }
    this._els.menu.addEventListener("click", () => {
      const s = this._els.sidebar;
      s.dataset.collapsed = s.dataset.collapsed === "true" ? "false" : "true";
    });

    HIGHLIGHT_COLORS.forEach(({ hex, name }) => {
      const btn = document.createElement("button");
      btn.title = name;
      btn.style.cssText = `width:18px;height:18px;background:${hex};border:1px solid rgba(0,0,0,0.18);cursor:pointer;flex-shrink:0;`;
      btn.addEventListener("mousedown", e => e.preventDefault());
      btn.addEventListener("click", () => this._applyHighlight(hex));
      this._els.swatches.appendChild(btn);
    });
    this._els.clearHighlight.addEventListener("mousedown", e => e.preventDefault());
    this._els.clearHighlight.addEventListener("click", () => this._applyHighlight(null));
    this._els.downloadBtn.addEventListener("click", () => this._exportWorkbook());

    this._els.fileInput.addEventListener("change", async e => {
      const file = e.target.files[0];
      e.target.value = "";
      if (!file) return;
      const buf = await file.arrayBuffer();
      let fileId = null;
      try {
        fileId = await this._db.save({ name: file.name, size: buf.byteLength, data: buf.slice(0), highlights: "{}" });
      } catch (err) { console.warn("IndexedDB save failed:", err); }
      await this._openWorkbook(buf, file.name, fileId);
    });

    document.addEventListener("mouseup", () => { this._sel.dragging = false; });

    window.addEventListener("beforeunload", e => {
      if (!this._save.isDirty()) return;
      this._save.flushNow();   // kick off the save immediately
      e.preventDefault();      // triggers browser "Leave site?" dialog
    });

    document.addEventListener("keydown", e => {
      if (e.target.closest("#sidebar, header, #toolbar")) return;
      if (this._editor.isEditing()) return;
      if (!this._wb || !this._sel.anchor) return;
      switch (e.key) {
        case "ArrowUp":    e.preventDefault(); this._moveBy(-1,  0, e.shiftKey); break;
        case "ArrowDown":  e.preventDefault(); this._moveBy( 1,  0, e.shiftKey); break;
        case "ArrowLeft":  e.preventDefault(); this._moveBy( 0, -1, e.shiftKey); break;
        case "ArrowRight": e.preventDefault(); this._moveBy( 0,  1, e.shiftKey); break;
        case "Enter":
          if (e.shiftKey) { this._moveBy(-1, 0); }
          else            { this._editor.start(this._sel.anchor.r, this._sel.anchor.c); }
          e.preventDefault(); break;
        case "F2":
          this._editor.start(this._sel.anchor.r, this._sel.anchor.c);
          e.preventDefault(); break;
        case "Delete":
        case "Backspace":
          this._deleteSelected(); e.preventDefault(); break;
        case "Escape":
          this._sel.clear(); break;
        default:
          if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            this._editor.start(this._sel.anchor.r, this._sel.anchor.c, e.key);
            e.preventDefault();
          }
      }
    });

    this._refreshLibrary().catch(err => console.warn("library init failed:", err));
  }

  // ── workbook ──────────────────────────────────────────────────────────────

  _detectAndRead(buf) {
    const b = new Uint8Array(buf, 0, 4);
    if (b[0] === 0x50 && b[1] === 0x4B) {
      return XLSX.read(new Uint8Array(buf), { type: "array", cellStyles: true });
    }
    return XLSX.read(new TextDecoder().decode(buf), { type: "string" });
  }

  async _openWorkbook(buf, name, fileId, savedHighlights = null) {
    this._save.cancel();
    this._els.saveIndicator.textContent = "";
    this._els.status.textContent = `parsing ${name} …`;
    try {
      const wb = this._detectAndRead(buf);
      this._wb     = wb;
      this._fileId = fileId ?? null;
      this._hl.restoreFrom(savedHighlights || {});
      this._buildSheetList();
      const first = wb.SheetNames[0];
      if (first) this._showSheet(first);
      else this._els.status.textContent = name + " — no sheets";
      await this._refreshLibrary();
    } catch (err) {
      this._showError(err);
    }
  }

  _showError(err) {
    console.error(err);
    this._els.status.textContent = "error";
    this._els.stage.replaceChildren(Object.assign(document.createElement("div"), {
      className: "text-red-600 font-mono whitespace-pre-wrap p-4",
      textContent: err.message || String(err),
    }));
  }

  // ── sheet rendering ───────────────────────────────────────────────────────

  _showSheet(name) {
    this._editor.commit();
    this._sheet = name;
    const ws = this._wb.Sheets[name];

    if (ws?.["!ref"]) {
      const r = XLSX.utils.decode_range(ws["!ref"]);
      this._els.status.textContent = `${name} — ${r.e.r - r.s.r + 1} rows × ${r.e.c - r.s.c + 1} cols`;
    } else {
      this._els.status.textContent = name;
    }
    this._updateSheetActive();

    const scroller = this._tr.build(ws, name, {
      onMousedown: e => {
        const td = e.target.closest("td[data-r]");
        if (!td) return;
        this._editor.commit();
        const r = parseInt(td.dataset.r, 10);
        const c = parseInt(td.dataset.c, 10);
        if (e.shiftKey && this._sel.anchor) { this._sel.extendTo(r, c); }
        else                                { this._sel.select(r, c); }
        this._sel.dragging = true;
        e.preventDefault();
      },
      onMousemove: e => {
        if (!this._sel.dragging) return;
        const td = e.target.closest("td[data-r]");
        if (!td) return;
        const r = parseInt(td.dataset.r, 10);
        const c = parseInt(td.dataset.c, 10);
        if (r !== this._sel.focus?.r || c !== this._sel.focus?.c) this._sel.extendTo(r, c);
      },
      onDblclick: e => {
        const td = e.target.closest("td[data-r]");
        if (!td) return;
        this._editor.start(parseInt(td.dataset.r, 10), parseInt(td.dataset.c, 10));
      },
    });

    this._editor.setScroller(scroller);
    this._els.stage.replaceChildren(scroller);
  }

  // ── sheet list ────────────────────────────────────────────────────────────

  _sheetItemClass(active) {
    return "px-2 py-1.5 cursor-pointer text-xs border mb-1 truncate " +
      (active ? "text-black border-black" : "text-gray-600 border-gray-200 hover:border-gray-500");
  }

  _buildSheetList() {
    const wb = this._wb;
    if (!wb?.SheetNames.length) {
      this._els.sheetList.replaceChildren(Object.assign(document.createElement("div"), {
        className: "text-[11px] text-gray-300 italic px-0.5 py-1", textContent: "no sheets",
      }));
      return;
    }
    this._els.sheetList.replaceChildren(...wb.SheetNames.map(name => {
      const d = document.createElement("div");
      d.className = this._sheetItemClass(name === this._sheet);
      d.textContent = name;
      d.title = name;
      d.addEventListener("click", () => this._showSheet(name));
      return d;
    }));
  }

  _updateSheetActive() {
    for (const d of this._els.sheetList.children) {
      d.className = this._sheetItemClass(d.textContent === this._sheet);
    }
  }

  // ── navigation ────────────────────────────────────────────────────────────

  _moveBy(dr, dc, extend = false) {
    if (!this._sel.anchor || !this._wb) return;
    const ws = this._wb.Sheets[this._sheet];
    const sr = ws?.["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
    if (!sr) return;
    const base = extend ? { ...this._sel.focus } : { ...this._sel.anchor };
    const nr = Math.max(sr.s.r, Math.min(sr.e.r, base.r + dr));
    const nc = Math.max(sr.s.c, Math.min(sr.e.c, base.c + dc));
    if (extend) { this._sel.extendTo(nr, nc); }
    else        { this._sel.select(nr, nc); }
    this._sel.el(nr, nc)?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  // ── cell operations ───────────────────────────────────────────────────────

  _deleteSelected() {
    const range = this._sel.getRange();
    if (!range || !this._wb) return;
    const ws = this._wb.Sheets[this._sheet];
    for (let r = range.minR; r <= range.maxR; r++) {
      for (let c = range.minC; c <= range.maxC; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const td = this._sel.el(r, c);
        if (td) td.textContent = "";
        if (ws[addr]) { ws[addr].v = undefined; ws[addr].w = ""; ws[addr].t = "z"; }
      }
    }
    this._save.schedule();
  }

  _applyHighlight(color) {
    const range = this._sel.getRange();
    if (!range || !this._sheet) return;
    this._hl.apply(this._sheet, range, color, this._sel, () => this._wb);
    this._save.schedule();
  }

  _exportWorkbook() {
    if (!this._wb) return;
    this._editor.commit();
    const wb = this._wb;
    for (const [sheetName, sheetHl] of Object.entries(this._hl.getAll())) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;
      for (const [key, color] of Object.entries(sheetHl)) {
        const [r, c] = key.split("_").map(Number);
        const addr = XLSX.utils.encode_cell({ r, c });
        ws[addr] = ws[addr] || { t: "z", v: "" };
        const hex = color.replace("#", "").toUpperCase();
        ws[addr].s = { ...(ws[addr].s || {}), fill: { patternType: "solid", fgColor: { rgb: "FF" + hex } } };
      }
    }
    try {
      const data = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
      const url = URL.createObjectURL(new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }));
      Object.assign(document.createElement("a"), { href: url, download: "edited.xlsx" }).click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error("export failed:", err);
      alert("Export failed: " + (err.message || err));
    }
  }

  // ── library ───────────────────────────────────────────────────────────────

  _emptyLibrary(text) {
    this._els.library.replaceChildren(Object.assign(document.createElement("div"), {
      className: "text-[11px] text-gray-300 italic px-0.5 py-1", textContent: text,
    }));
  }

  _makeLibraryRow(file) {
    const active = file.id === this._fileId;
    const row = document.createElement("div");
    row.className = "flex items-center gap-1.5 px-2 py-1.5 border mb-1 text-xs hover:border-gray-500 " +
      (active ? "text-black border-black" : "text-gray-600 border-gray-200");

    const label = document.createElement("span");
    label.className = "flex-1 cursor-pointer overflow-hidden text-ellipsis whitespace-nowrap";
    label.textContent = file.name;
    label.title = `${file.name} — ${(file.size / 1024).toFixed(0)} KB`;
    label.addEventListener("click", async () => {
      const rec = await this._db.get(file.id);
      if (!rec) return;
      let savedHighlights = null;
      try { savedHighlights = rec.highlights ? JSON.parse(rec.highlights) : null; } catch { /* ignore */ }
      await this._openWorkbook(rec.data, rec.name, file.id, savedHighlights);
    });

    const del = document.createElement("span");
    del.className = "cursor-pointer text-gray-400 hover:text-red-600 px-1 text-sm leading-none";
    del.textContent = "×";
    del.title = "remove from library";
    del.addEventListener("click", async e => {
      e.stopPropagation();
      await this._db.delete(file.id);
      if (this._fileId === file.id) this._fileId = null;
      await this._refreshLibrary();
    });

    row.append(label, del);
    return row;
  }

  async _refreshLibrary() {
    let files;
    try { files = await this._db.list(); } catch { this._emptyLibrary("storage unavailable"); return; }
    if (!files.length) { this._emptyLibrary("no saved files"); return; }
    this._els.library.replaceChildren(...files.map(f => this._makeLibraryRow(f)));
  }
}

new XlsxApp();
