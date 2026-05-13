"use strict";

// Depends (global): mammoth, FileDatabase

class DocxApp {
  constructor() {
    this._db      = new FileDatabase("docx-viewer", "files");
    this._fileId  = null;

    this._els = {
      fileInput: document.getElementById("fileInput"),
      status:    document.getElementById("status"),
      stage:     document.getElementById("stage"),
      sidebar:   document.getElementById("sidebar"),
      library:   document.getElementById("library"),
      outline:   document.getElementById("outline"),
      menu:      document.getElementById("menuBtn"),
    };

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

    this._els.fileInput.addEventListener("change", async e => {
      const file = e.target.files[0];
      e.target.value = "";
      if (!file) return;
      const buf = await file.arrayBuffer();
      let fileId = null;
      try {
        fileId = await this._db.save({ name: file.name, size: buf.byteLength, data: buf.slice(0) });
      } catch (err) { console.warn("IndexedDB save failed:", err); }
      await this._openDoc(buf, file.name, fileId);
    });

    this._refreshLibrary().catch(err => console.warn("library init failed:", err));
  }

  // ── document rendering ────────────────────────────────────────────────────

  async _openDoc(buf, name, fileId) {
    this._els.status.textContent = `converting ${name} …`;

    const magic = new Uint8Array(buf, 0, 4);
    const isZip = magic[0] === 0x50 && magic[1] === 0x4B;
    const isOle = magic[0] === 0xD0 && magic[1] === 0xCF;

    if (isOle) {
      this._showError(new Error(
        ".doc (old binary format) is not supported.\n" +
        "Please open the file in Word or LibreOffice and save as .docx, then upload again."
      ));
      return;
    }
    if (!isZip) {
      this._showError(new Error("File does not appear to be a valid .docx (Office Open XML) file."));
      return;
    }

    try {
      const result = await mammoth.convertToHtml({ arrayBuffer: buf });

      const page = document.createElement("div");
      page.id = "doc-content";
      page.style.cssText = "max-width:800px;margin:0 auto;background:#fff;padding:48px 64px;min-height:600px;box-shadow:0 1px 4px rgba(0,0,0,0.1);font-size:15px;line-height:1.6;color:#111;";
      page.innerHTML = result.value;

      const headings = [];
      let hIdx = 0;
      page.querySelectorAll("h1,h2,h3,h4,h5,h6").forEach(h => {
        const id = `h-${hIdx++}`;
        h.id = id;
        headings.push({ id, level: parseInt(h.tagName[1], 10), text: h.textContent.trim() });
      });

      this._els.stage.replaceChildren(page);
      this._buildOutline(headings);

      const wordCount = (page.textContent.match(/\S+/g) || []).length;
      this._els.status.textContent = `${name} — ~${wordCount.toLocaleString()} words`;
      this._fileId = fileId ?? null;

      if (result.messages.length) {
        const notice = document.createElement("div");
        notice.style.cssText = "max-width:800px;margin:8px auto 0;padding:8px 12px;font-size:11px;color:#92400e;background:#fffbeb;border:1px solid #fde68a;border-radius:2px;";
        notice.textContent = `${result.messages.length} conversion note(s): ` +
          result.messages.slice(0, 3).map(m => m.message).join("; ");
        this._els.stage.appendChild(notice);
      }

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

  // ── outline ───────────────────────────────────────────────────────────────

  _buildOutline(headings) {
    if (!headings.length) {
      this._els.outline.replaceChildren(Object.assign(document.createElement("div"), {
        className: "text-[11px] text-gray-300 italic px-0.5 py-1", textContent: "no headings",
      }));
      return;
    }
    const items = headings.map(h => {
      const d = document.createElement("div");
      d.style.paddingLeft = (h.level - 1) * 10 + "px";
      d.className = "py-1 px-1 text-[11px] cursor-pointer text-gray-600 hover:text-black hover:bg-gray-50 truncate leading-snug";
      d.textContent = h.text;
      d.title = h.text;
      d.addEventListener("click", () => {
        document.getElementById(h.id)?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
      return d;
    });
    this._els.outline.replaceChildren(...items);
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
      if (rec) await this._openDoc(rec.data, rec.name, file.id);
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

new DocxApp();
