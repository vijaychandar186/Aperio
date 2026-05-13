"use strict";

// Depends on: ShapeParser, XmlUtils (for resolveMediaUrl)

const SlideRenderer = {
  render(slide, host, presentation) {
    const { slideW, slideH } = presentation;
    host.style.width  = slideW + "px";
    host.style.height = slideH + "px";
    host.innerHTML    = "";

    // Background: slide → layout → master → white
    host.style.background =
      slide.bg                        ? slide.bg :
      slide.layout?.bg                ? slide.layout.bg :
      slide.master?.bg                ? slide.master.bg : "#fff";

    const renderNodes = (nodes, pt) => nodes.forEach(n => SlideRenderer._node(n, host, presentation, slide, pt));
    if (slide.master?.nodes) renderNodes(slide.master.nodes, null);
    if (slide.layout?.nodes) renderNodes(slide.layout.nodes, null);
    renderNodes(slide.nodes, null);
  },

  _node(node, host, presentation, slide, parentTransform) {
    try {
      switch (node.kind) {
        case "shape": SlideRenderer._shape(node, host, presentation, slide, parentTransform); break;
        case "pic":   SlideRenderer._pic(node, host, presentation, slide, parentTransform);   break;
        case "table": SlideRenderer._table(node, host, presentation, slide, parentTransform); break;
        case "group": SlideRenderer._group(node, host, presentation, slide, parentTransform); break;
      }
    } catch (e) { console.warn("render fail:", e, node); }
  },

  _applyTransform(node, pt) {
    if (!node.xfrm) return null;
    let { x, y, w, h, rot, flipH, flipV } = node.xfrm;
    if (pt) {
      const sx = pt.chW ? pt.w / pt.chW : 1;
      const sy = pt.chH ? pt.h / pt.chH : 1;
      x = pt.x + (x - pt.chX) * sx;
      y = pt.y + (y - pt.chY) * sy;
      w = w * sx; h = h * sy;
    }
    return { x, y, w, h, rot, flipH, flipV };
  },

  _wrapStyle(t) {
    const tr = [];
    if (t.rot)   tr.push(`rotate(${t.rot}deg)`);
    if (t.flipH) tr.push("scaleX(-1)");
    if (t.flipV) tr.push("scaleY(-1)");
    return `position:absolute;left:${t.x}px;top:${t.y}px;width:${t.w}px;height:${t.h}px;` +
           (tr.length ? `transform:${tr.join(" ")};` : "");
  },

  _fillAttr(fill, defs, idGen) {
    if (!fill || fill.kind === "none") return "none";
    if (fill.kind === "solid")    return fill.value;
    if (fill.kind === "gradient") {
      const id  = "g" + idGen();
      const ang = (fill.angle || 0) * Math.PI / 180;
      const x1  = 0.5 - Math.cos(ang) * 0.5, y1 = 0.5 - Math.sin(ang) * 0.5;
      const x2  = 0.5 + Math.cos(ang) * 0.5, y2 = 0.5 + Math.sin(ang) * 0.5;
      const stops = fill.stops.map(s => `<stop offset="${s.pos}%" stop-color="${s.col}"/>`).join("");
      defs.push(`<linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient>`);
      return `url(#${id})`;
    }
    if (fill.kind === "image") return { image: true, rId: fill.rId };
    return "none";
  },

  _shape(node, host, presentation, slide, pt) {
    const t = SlideRenderer._applyTransform(node, pt);
    if (!t) return;
    const wrap = Object.assign(document.createElement("div"), { style: { cssText: SlideRenderer._wrapStyle(t) } });
    wrap.style.cssText = SlideRenderer._wrapStyle(t);

    const w = Math.max(1, t.w), h = Math.max(1, t.h);
    let path = node.prst ? ShapeParser.presetPath(node.prst, w, h, node.adj) : null;
    if (!path && node.prst && (node.prst.includes("line") || node.prst.includes("Connector"))) path = `M0,0 L${w},${h}`;

    if (path || node.fill || node.line) {
      const ns  = "http://www.w3.org/2000/svg";
      const svg = document.createElementNS(ns, "svg");
      svg.setAttribute("width", w); svg.setAttribute("height", h);
      svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
      svg.style.cssText = "position:absolute;left:0;top:0;overflow:visible;";

      const defs = []; let nextId = 1;
      const idGen = () => (nextId++) + "-" + Math.random().toString(36).slice(2, 6);
      const fillVal  = SlideRenderer._fillAttr(node.fill, defs, idGen);
      let fillAttr   = fillVal || "none";
      let bgImageRId = null;
      if (fillVal && typeof fillVal === "object" && fillVal.image) { bgImageRId = fillVal.rId; fillAttr = "none"; }
      if (defs.length) { const d = document.createElementNS(ns, "defs"); d.innerHTML = defs.join(""); svg.appendChild(d); }

      if (path) {
        const p = document.createElementNS(ns, "path");
        p.setAttribute("d", path); p.setAttribute("fill", fillAttr || "none");
        if (node.line?.kind !== "none" && node.line) { p.setAttribute("stroke", node.line.color); p.setAttribute("stroke-width", node.line.width); }
        else p.setAttribute("stroke", "none");
        svg.appendChild(p);
      }
      wrap.appendChild(svg);

      if (bgImageRId) {
        const url = SlideRenderer._mediaUrl(bgImageRId, slide, presentation);
        if (url) { wrap.style.backgroundImage = `url(${url})`; wrap.style.backgroundSize = "cover"; wrap.style.backgroundRepeat = "no-repeat"; }
      }
    }

    if (node.text?.paragraphs.length) {
      const txt = SlideRenderer._textBody(node.text);
      txt.style.cssText = `position:absolute;left:0;top:0;width:${w}px;height:${h}px;`;
      wrap.appendChild(txt);
    }
    host.appendChild(wrap);
  },

  _textBody(tb) {
    const el = document.createElement("div");
    Object.assign(el.style, { boxSizing: "border-box", padding: "4px 8px", overflow: "hidden", display: "flex", flexDirection: "column", justifyContent: tb.anchor === "ctr" ? "center" : tb.anchor === "b" ? "flex-end" : "flex-start" });
    for (const p of tb.paragraphs) {
      const para  = document.createElement("div");
      para.style.textAlign  = p.align === "ctr" ? "center" : p.align === "r" ? "right" : p.align === "just" ? "justify" : "left";
      para.style.margin     = "0"; para.style.lineHeight = "1.2";
      if (p.lvl) para.style.marginLeft = (p.lvl * 16) + "px";
      if (!p.runs.length) { para.appendChild(document.createElement("br")); }
      else {
        for (const r of p.runs) {
          if (r.br) { para.appendChild(document.createElement("br")); continue; }
          const span = document.createElement("span");
          span.textContent = r.text || "";
          const rp = r.rPr || {};
          if (rp.bold)      span.style.fontWeight     = "bold";
          if (rp.italic)    span.style.fontStyle      = "italic";
          if (rp.underline) span.style.textDecoration = "underline";
          if (rp.size)      span.style.fontSize       = rp.size + "pt";
          if (rp.color)     span.style.color          = rp.color;
          if (rp.font)      span.style.fontFamily     = rp.font + ", sans-serif";
          para.appendChild(span);
        }
      }
      el.appendChild(para);
    }
    return el;
  },

  _pic(node, host, presentation, slide, pt) {
    const t = SlideRenderer._applyTransform(node, pt);
    if (!t) return;
    const wrap = document.createElement("div");
    wrap.style.cssText = SlideRenderer._wrapStyle(t);
    const url = SlideRenderer._mediaUrl(node.rId, slide, presentation);
    if (url) {
      const img = document.createElement("img");
      img.src = url; img.style.width = "100%"; img.style.height = "100%"; img.style.objectFit = "fill";
      wrap.appendChild(img);
    } else { wrap.style.background = "#eee"; wrap.style.border = "1px dashed #999"; }
    host.appendChild(wrap);
  },

  _table(node, host, presentation, slide, pt) {
    const t = SlideRenderer._applyTransform(node, pt);
    if (!t) return;
    const wrap = document.createElement("div");
    wrap.style.cssText = SlideRenderer._wrapStyle(t);
    const tbl = document.createElement("table");
    tbl.style.cssText = "border-collapse:collapse;width:100%;height:100%;table-layout:fixed;";
    const cg = document.createElement("colgroup");
    for (const cw of node.cols) { const col = document.createElement("col"); col.style.width = cw + "px"; cg.appendChild(col); }
    tbl.appendChild(cg);
    for (const row of node.rows) {
      const tr = document.createElement("tr");
      if (row.h) tr.style.height = row.h + "px";
      for (const cell of row.cells) {
        if (cell.hMerge || cell.vMerge) continue;
        const td = document.createElement("td");
        td.style.cssText = "border:1px solid #888;padding:4px;vertical-align:top;";
        if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
        if (cell.rowSpan  > 1) td.rowSpan = cell.rowSpan;
        if (cell.fill?.kind === "solid") td.style.background = cell.fill.value;
        if (cell.text) { const txt = SlideRenderer._textBody(cell.text); txt.style.padding = "0"; td.appendChild(txt); }
        tr.appendChild(td);
      }
      tbl.appendChild(tr);
    }
    wrap.appendChild(tbl);
    host.appendChild(wrap);
  },

  _group(node, host, presentation, slide, pt) {
    if (!node.xfrm) return;
    let g = node.xfrm;
    if (pt) {
      const sx = pt.chW ? pt.w / pt.chW : 1, sy = pt.chH ? pt.h / pt.chH : 1;
      g = { x: pt.x + (g.x - pt.chX) * sx, y: pt.y + (g.y - pt.chY) * sy, w: g.w * sx, h: g.h * sy, chX: g.chX, chY: g.chY, chW: g.chW, chH: g.chH };
    }
    for (const child of node.children) SlideRenderer._node(child, host, presentation, slide, g);
  },

  _mediaUrl(rId, slide, presentation) {
    if (!rId) return null;
    const rel  = slide.rels?.[rId];
    if (!rel) return null;
    const path = XmlUtils.resolvePath(slide.baseDir, rel.target);
    return presentation.mediaUrls[path] || null;
  },
};
