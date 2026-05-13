# Aperio

Local, browser-based viewers for Office file formats. *Aperio* — Latin for "to open, to reveal." No uploads, no server-side processing — everything runs in the browser using IndexedDB for persistence.

## Viewers

| File | Formats | Features |
|------|---------|----------|
| `templates/pptx.html` | `.pptx` | Slide deck viewer, thumbnail strip, fullscreen present mode |
| `templates/xlsx.html` | `.xlsx`, `.xls`, `.csv` | Editable cells, cell highlighting, multi-sheet, auto-save, download |
| `templates/docx.html` | `.docx` | Document renderer, heading outline sidebar |

## Getting started

```bash
npm install
npm run serve
```

Then open `http://localhost:3000` in your browser.

## Scripts

| Command | Description |
|---------|-------------|
| `npm run serve` | Start a local static file server on port 3000 |
| `npm run css` | Compile Tailwind CSS to `css/app.css` (minified) |
| `npm run css:watch` | Watch and recompile CSS on changes |

> Run `npm run css` after editing any HTML or JS files that introduce new Tailwind classes.

## Project structure

```
├── index.html              # Landing page with links to all viewers
├── templates/
│   ├── pptx.html           # PPTX viewer
│   ├── xlsx.html           # XLSX / CSV viewer
│   └── docx.html           # DOCX viewer
├── css/
│   ├── input.css           # Tailwind source + custom styles
│   └── app.css             # Compiled output (committed)
├── js/
│   ├── lib/                # Vendored third-party libraries (no CDN)
│   │   ├── jszip.min.js
│   │   ├── xlsx.full.min.js
│   │   └── mammoth.browser.min.js
│   ├── shared/
│   │   └── FileDatabase.js # Generic IndexedDB wrapper
│   ├── pptx/
│   │   ├── ColorUtils.js   # EMU units, color math, theme resolution
│   │   ├── XmlUtils.js     # XML parsing helpers
│   │   ├── ShapeParser.js  # Shape / text / table parsing
│   │   ├── SlideRenderer.js# DOM rendering
│   │   ├── PptxLoader.js   # Async zip + parse orchestration
│   │   └── PptxApp.js      # UI and event wiring
│   ├── xlsx/
│   │   ├── SelectionManager.js  # Cell selection (click, drag, keyboard)
│   │   ├── HighlightManager.js  # Per-cell color highlights
│   │   ├── TableRenderer.js     # Builds the sticky-header spreadsheet table
│   │   ├── CellEditor.js        # Floating inline cell editor
│   │   ├── AutoSave.js          # Debounced IndexedDB auto-save
│   │   └── XlsxApp.js           # Main app class
│   └── docx/
│       └── DocxApp.js      # Main app class
├── tailwind.config.js
└── package.json
```

## Libraries used

| Library | Version | Purpose |
|---------|---------|---------|
| [SheetJS (xlsx)](https://sheetjs.com) | 0.18.5 | XLSX / XLS / CSV parsing and writing |
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | ZIP extraction for PPTX |
| [mammoth.js](https://github.com/mwilliamson/mammoth.js) | 1.6.0 | DOCX to HTML conversion |
| [Tailwind CSS](https://tailwindcss.com) | 3.4.x | Utility-first CSS (compiled locally) |

## XLSX features

- **Edit cells** — double-click or start typing on any selected cell; Enter / Tab to confirm and move
- **Keyboard navigation** — arrow keys, Shift+arrow to extend selection, F2 to edit, Delete to clear
- **Highlight colors** — select a range then click a swatch; persisted across sessions
- **Auto-save** — edits and highlights are saved to IndexedDB automatically (debounced 1.5 s)
- **Unsaved changes guard** — refreshing or closing with pending changes triggers a "Leave site?" prompt and immediately flushes the save
- **Download** — exports the current workbook as `.xlsx` with highlights baked into cell styles
- **Library** — previously opened files are stored locally and can be reopened from the sidebar

## Notes

- All file storage uses the browser's **IndexedDB** — data stays on your machine
- The `css` script must be re-run whenever you add new Tailwind classes to HTML or JS files
- `.doc` (old binary format) is not supported in the DOCX viewer; resave as `.docx` first
