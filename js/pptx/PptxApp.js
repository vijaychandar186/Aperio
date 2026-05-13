"use strict";

// Depends on: FileDatabase, PptxLoader, SlideRenderer

class PptxApp {
  constructor() {
    this.db           = new FileDatabase("pptx-viewer", "decks");
    this.presentation = null;
    this.current      = 0;
    this.currentId    = null;

    this._els = {
      fileInput: document.getElementById("fileInput"),
      status:    document.getElementById("status"),
      thumbs:    document.getElementById("thumbs"),
      library:   document.getElementById("library"),
      stage:     document.getElementById("stage"),
      sidebar:   document.getElementById("sidebar"),
      prev:      document.getElementById("prevBtn"),
      next:      document.getElementById("nextBtn"),
      counter:   document.getElementById("counter"),
      present:   document.getElementById("presentBtn"),
      menu:      document.getElementById("menuBtn"),
    };

    this._resizeQueued = false;
    this._init();
  }

  _init() {
    const els = this._els;
    // Sidebar
    if (window.matchMedia("(max-width: 767px)").matches) els.sidebar.dataset.collapsed = "true";
    els.menu.addEventListener("click", () => {
      els.sidebar.dataset.collapsed = els.sidebar.dataset.collapsed === "true" ? "false" : "true";
    });

    // File input
    els.fileInput.addEventListener("change", async e => {
      const file = e.target.files[0]; e.target.value = "";
      if (!file) return;
      const buf = await file.arrayBuffer();
      const b   = new Uint8Array(buf, 0, 4);
      if (!(b[0] === 0x50 && b[1] === 0x4B && file.name.toLowerCase().endsWith(".pptx"))) {
        this._error(new Error("upload only .pptx files")); return;
      }
      let id = null;
      try { id = await this.db.save({ name: file.name, size: buf.byteLength, data: buf.slice(0) }); } catch { /* ignore */ }
      await this._open(buf, file.name, id);
    });

    // Controls
    els.prev.addEventListener("click", () => this._show(this.current - 1));
    els.next.addEventListener("click", () => this._show(this.current + 1));
    els.present.addEventListener("click", () => this._isPresenting() ? this._exitPresent() : this._enterPresent());

    // Keyboard
    document.addEventListener("keydown", e => {
      if (!this.presentation) return;
      const last = this.presentation.slides.length - 1;
      switch (e.key) {
        case "ArrowRight": case " ": case "PageDown": e.preventDefault(); this._show(this.current + 1); break;
        case "ArrowLeft":  case "PageUp":             e.preventDefault(); this._show(this.current - 1); break;
        case "Home":                                  e.preventDefault(); this._show(0);    break;
        case "End":                                   e.preventDefault(); this._show(last); break;
        case "Escape": if (this._isPresenting()) this._exitPresent(); break;
        case "f": case "F": if (!this._isPresenting()) this._enterPresent(); break;
      }
    });

    // Fullscreen
    document.addEventListener("fullscreenchange", () => {
      if (!document.fullscreenElement && this._isPresenting()) this._exitPresent();
    });

    // Resize
    window.addEventListener("resize", () => {
      if (this._resizeQueued || !this.presentation) return;
      this._resizeQueued = true;
      requestAnimationFrame(() => { this._resizeQueued = false; this._show(this.current); });
    });

    this.db.list().then(decks => this._renderLibrary(decks)).catch(() => this._emptyLibrary("storage unavailable"));
  }

  async _open(buf, name, deckId) {
    this._els.status.textContent = `parsing ${name} …`;
    try {
      this.presentation = await PptxLoader.load(buf);
      this.current      = 0;
      this.currentId    = deckId ?? null;
      this._els.status.textContent = `${name} — ${this.presentation.slides.length} slides`;
      this._buildThumbs();
      this._show(0);
      const decks = await this.db.list();
      this._renderLibrary(decks);
    } catch (err) { this._error(err); }
  }

  _error(err) {
    console.error(err);
    this._els.status.textContent = "error";
    this._els.stage.replaceChildren(Object.assign(document.createElement("div"), {
      className: "text-red-600 font-mono whitespace-pre-wrap p-4",
      textContent: err.message || String(err),
    }));
  }

  // ── slide display ─────────────────────────────────────────────────────────
  _scale() {
    const pres = this.presentation;
    if (this._isPresenting()) return Math.min(window.innerWidth / pres.slideW, window.innerHeight / pres.slideH);
    const r = this._els.stage.getBoundingClientRect();
    return Math.min((r.width - 32) / pres.slideW, (r.height - 32) / pres.slideH, 1);
  }

  _show(i) {
    const pres = this.presentation;
    if (!pres) return;
    this.current = Math.max(0, Math.min(i, pres.slides.length - 1));
    const presenting = this._isPresenting();
    const scale      = this._scale();
    const wrap       = document.createElement("div");
    wrap.className   = "bg-white text-black relative shadow-[0_0_0_1px_#ccc]";
    wrap.style.cssText = `width:${pres.slideW}px;height:${pres.slideH}px;transform-origin:${presenting ? "center center" : "top left"};transform:scale(${scale});`;
    this._els.stage.replaceChildren(wrap);
    this._els.stage.style.minHeight = presenting ? "" : `${pres.slideH * scale + 32}px`;
    SlideRenderer.render(pres.slides[this.current], wrap, pres);
    this._els.counter.textContent = `${this.current + 1} / ${pres.slides.length}`;
    this._updateThumbActive();
  }

  // ── present mode ──────────────────────────────────────────────────────────
  _isPresenting() { return this._els.stage.dataset.presenting === "true"; }

  async _enterPresent() {
    if (!this.presentation) return;
    this._els.stage.dataset.presenting = "true";
    try { await this._els.stage.requestFullscreen(); } catch { /* ignore */ }
    this._show(this.current);
  }

  _exitPresent() {
    this._els.stage.dataset.presenting = "false";
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    if (this.presentation) this._show(this.current);
  }

  // ── thumbnails ────────────────────────────────────────────────────────────
  _thumbClass(active) {
    return "px-2 py-1.5 cursor-pointer text-xs border mb-1 " +
           (active ? "text-black border-black" : "text-gray-600 border-gray-200");
  }

  _buildThumbs() {
    const rows = this.presentation.slides.map((_, i) => {
      const d = document.createElement("div");
      d.dataset.index = String(i);
      d.className = this._thumbClass(i === this.current);
      d.textContent = `Slide ${i + 1}`;
      d.addEventListener("click", () => this._show(i));
      return d;
    });
    this._els.thumbs.replaceChildren(...rows);
  }

  _updateThumbActive() {
    for (const t of this._els.thumbs.children)
      t.className = this._thumbClass(Number(t.dataset.index) === this.current);
  }

  // ── library ───────────────────────────────────────────────────────────────
  _emptyLibrary(text) {
    this._els.library.replaceChildren(Object.assign(document.createElement("div"), {
      className: "text-[11px] text-gray-300 italic px-0.5 py-1", textContent: text,
    }));
  }

  _renderLibrary(decks) {
    if (!decks.length) { this._emptyLibrary("no saved decks"); return; }
    this._els.library.replaceChildren(...decks.map(d => this._libraryRow(d)));
  }

  _libraryRow(deck) {
    const active = deck.id === this.currentId;
    const row    = document.createElement("div");
    row.className = "flex items-center gap-1.5 px-2 py-1.5 border mb-1 text-xs hover:border-gray-500 " +
                    (active ? "text-black border-black" : "text-gray-600 border-gray-200");

    const name = document.createElement("span");
    name.className = "flex-1 cursor-pointer overflow-hidden text-ellipsis whitespace-nowrap";
    name.textContent = deck.name;
    name.title = `${deck.name} — ${(deck.size / 1024).toFixed(0)} KB`;
    name.addEventListener("click", async () => {
      const rec = await this.db.get(deck.id);
      if (rec) await this._open(rec.data, rec.name, deck.id);
    });

    const del = document.createElement("span");
    del.className = "cursor-pointer text-gray-400 hover:text-red-600 px-1 text-sm leading-none";
    del.textContent = "×"; del.title = "remove from library";
    del.addEventListener("click", async e => {
      e.stopPropagation();
      await this.db.delete(deck.id);
      if (this.currentId === deck.id) this.currentId = null;
      this._renderLibrary(await this.db.list());
    });

    row.append(name, del);
    return row;
  }
}

// Boot
new PptxApp();
