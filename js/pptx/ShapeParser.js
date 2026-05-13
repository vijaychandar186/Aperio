"use strict";

// Depends on: XmlUtils, ColorUtils (emu, deg, pct, readColor)

const ShapeParser = {
  // ── geometry ────────────────────────────────────────────────────────────
  presetPath(prst, w, h, adj) {
    const r = (k, d) => adj && adj[k] != null ? adj[k] : d;
    switch (prst) {
      case "rect": case "snip1Rect": case "snip2SameRect":
        return `M0,0 L${w},0 L${w},${h} L0,${h} Z`;
      case "roundRect": {
        const rad = Math.min(w, h) * pct(r("adj", 16667));
        return `M${rad},0 L${w - rad},0 Q${w},0 ${w},${rad} L${w},${h - rad} Q${w},${h} ${w - rad},${h} L${rad},${h} Q0,${h} 0,${h - rad} L0,${rad} Q0,0 ${rad},0 Z`;
      }
      case "ellipse": case "circle": {
        const rx = w / 2, ry = h / 2;
        return `M0,${ry} A${rx},${ry} 0 1,0 ${w},${ry} A${rx},${ry} 0 1,0 0,${ry} Z`;
      }
      case "triangle":      return `M${w / 2},0 L${w},${h} L0,${h} Z`;
      case "rtTriangle":    return `M0,0 L0,${h} L${w},${h} Z`;
      case "diamond":       return `M${w / 2},0 L${w},${h / 2} L${w / 2},${h} L0,${h / 2} Z`;
      case "parallelogram": { const o = w * 0.25; return `M${o},0 L${w},0 L${w - o},${h} L0,${h} Z`; }
      case "trapezoid":     { const o = w * 0.25; return `M${o},0 L${w - o},0 L${w},${h} L0,${h} Z`; }
      case "pentagon": {
        const cx = w / 2, cy = h / 2, pts = [];
        for (let i = 0; i < 5; i++) {
          const a = -Math.PI / 2 + i * 2 * Math.PI / 5;
          pts.push(`${cx + (w / 2) * Math.cos(a)},${cy + (h / 2) * Math.sin(a)}`);
        }
        return "M" + pts.join(" L") + " Z";
      }
      case "hexagon":  { const o = w * 0.25; return `M${o},0 L${w - o},0 L${w},${h / 2} L${w - o},${h} L${o},${h} L0,${h / 2} Z`; }
      case "octagon":  { const o = Math.min(w, h) * 0.29289; return `M${o},0 L${w - o},0 L${w},${o} L${w},${h - o} L${w - o},${h} L${o},${h} L0,${h - o} L0,${o} Z`; }
      case "rightArrow": { const aw = w * 0.3; return `M0,${h * 0.25} L${w - aw},${h * 0.25} L${w - aw},0 L${w},${h / 2} L${w - aw},${h} L${w - aw},${h * 0.75} L0,${h * 0.75} Z`; }
      case "leftArrow":  return `M${w},${h * 0.25} L${w * 0.3},${h * 0.25} L${w * 0.3},0 L0,${h / 2} L${w * 0.3},${h} L${w * 0.3},${h * 0.75} L${w},${h * 0.75} Z`;
      case "upArrow":    return `M${w * 0.25},${h} L${w * 0.25},${h * 0.3} L0,${h * 0.3} L${w / 2},0 L${w},${h * 0.3} L${w * 0.75},${h * 0.3} L${w * 0.75},${h} Z`;
      case "downArrow":  return `M${w * 0.25},0 L${w * 0.25},${h * 0.7} L0,${h * 0.7} L${w / 2},${h} L${w},${h * 0.7} L${w * 0.75},${h * 0.7} L${w * 0.75},0 Z`;
      case "star5": {
        const cx = w / 2, cy = h / 2, ro = Math.min(w, h) / 2, ri = ro * 0.38;
        let d = "";
        for (let i = 0; i < 10; i++) {
          const a = -Math.PI / 2 + i * Math.PI / 5;
          const rr = i % 2 === 0 ? ro : ri;
          d += (i === 0 ? "M" : "L") + (cx + rr * Math.cos(a)) + "," + (cy + rr * Math.sin(a)) + " ";
        }
        return d + "Z";
      }
      case "line": case "straightConnector1": return `M0,0 L${w},${h}`;
      default: return null;
    }
  },

  // ── transform ────────────────────────────────────────────────────────────
  parseTransform(spPr) {
    const xfrm = XmlUtils.childByLocal(spPr, "xfrm");
    if (!xfrm) return null;
    const off = XmlUtils.childByLocal(xfrm, "off");
    const ext = XmlUtils.childByLocal(xfrm, "ext");
    return {
      x: off ? emu(off.getAttribute("x")) : 0,
      y: off ? emu(off.getAttribute("y")) : 0,
      w: ext ? emu(ext.getAttribute("cx")) : 0,
      h: ext ? emu(ext.getAttribute("cy")) : 0,
      rot:   xfrm.getAttribute("rot")   ? deg(xfrm.getAttribute("rot")) : 0,
      flipH: xfrm.getAttribute("flipH") === "1",
      flipV: xfrm.getAttribute("flipV") === "1",
    };
  },

  parseGroupTransform(grpSp) {
    const grpSpPr = XmlUtils.childByLocal(grpSp, "grpSpPr");
    const xfrm    = XmlUtils.childByLocal(grpSpPr, "xfrm");
    if (!xfrm) return null;
    const off   = XmlUtils.childByLocal(xfrm, "off");
    const ext   = XmlUtils.childByLocal(xfrm, "ext");
    const chOff = XmlUtils.childByLocal(xfrm, "chOff");
    const chExt = XmlUtils.childByLocal(xfrm, "chExt");
    return {
      x: off ? emu(off.getAttribute("x")) : 0, y: off ? emu(off.getAttribute("y")) : 0,
      w: ext ? emu(ext.getAttribute("cx")) : 0, h: ext ? emu(ext.getAttribute("cy")) : 0,
      chX: chOff ? emu(chOff.getAttribute("x")) : 0, chY: chOff ? emu(chOff.getAttribute("y")) : 0,
      chW: chExt ? emu(chExt.getAttribute("cx")) : 0, chH: chExt ? emu(chExt.getAttribute("cy")) : 0,
    };
  },

  // ── fill / line ──────────────────────────────────────────────────────────
  parseFill(spPr, ctx) {
    if (!spPr) return null;
    if (XmlUtils.childByLocal(spPr, "noFill")) return { kind: "none" };
    const sf = XmlUtils.childByLocal(spPr, "solidFill");
    if (sf) { const c = readColor(sf, ctx); if (c) return { kind: "solid", value: c }; }
    const gf = XmlUtils.childByLocal(spPr, "gradFill");
    if (gf) {
      const lst   = XmlUtils.childByLocal(gf, "gsLst");
      const stops = [];
      if (lst) {
        for (const gs of XmlUtils.childrenByLocal(lst, "gs")) {
          const col = readColor(gs, ctx);
          if (col) stops.push({ pos: pct(gs.getAttribute("pos")) * 100, col });
        }
      }
      const lin   = XmlUtils.childByLocal(gf, "lin");
      const angle = lin ? deg(lin.getAttribute("ang")) : 90;
      if (stops.length) return { kind: "gradient", angle, stops };
    }
    const bf = XmlUtils.childByLocal(spPr, "blipFill");
    if (bf) {
      const blip  = XmlUtils.childByLocal(bf, "blip");
      const embed = blip ? XmlUtils.attrAny(blip, "embed") : null;
      if (embed) return { kind: "image", rId: embed };
    }
    return null;
  },

  parseLine(spPr, ctx) {
    if (!spPr) return null;
    const ln = XmlUtils.childByLocal(spPr, "ln");
    if (!ln) return null;
    if (XmlUtils.childByLocal(ln, "noFill")) return { kind: "none" };
    const sf    = XmlUtils.childByLocal(ln, "solidFill");
    const color = sf ? readColor(sf, ctx) : null;
    return { width: ln.getAttribute("w") ? emu(ln.getAttribute("w")) : 1, color: color || "#000" };
  },

  // ── text ─────────────────────────────────────────────────────────────────
  parseTextBody(txBody, ctx) {
    if (!txBody) return null;
    const bodyPr     = XmlUtils.childByLocal(txBody, "bodyPr");
    const paragraphs = [];
    for (const p of XmlUtils.childrenByLocal(txBody, "p")) {
      const pPr  = XmlUtils.childByLocal(p, "pPr");
      const runs = [];
      for (const child of p.children) {
        if (child.localName === "r") {
          const tEl = XmlUtils.childByLocal(child, "t");
          runs.push({ text: tEl ? tEl.textContent : "", rPr: ShapeParser.parseRunPr(XmlUtils.childByLocal(child, "rPr"), ctx) });
        } else if (child.localName === "br") {
          runs.push({ br: true });
        } else if (child.localName === "fld") {
          const tEl = XmlUtils.childByLocal(child, "t");
          runs.push({ text: tEl ? tEl.textContent : "", rPr: ShapeParser.parseRunPr(XmlUtils.childByLocal(child, "rPr"), ctx) });
        }
      }
      paragraphs.push({ runs, align: pPr?.getAttribute("algn") || null, lvl: Number(pPr?.getAttribute("lvl") || 0) });
    }
    return { paragraphs, anchor: bodyPr ? (bodyPr.getAttribute("anchor") || "t") : "t" };
  },

  parseRunPr(rPr, ctx) {
    if (!rPr) return {};
    const out = {
      bold:      rPr.getAttribute("b") === "1",
      italic:    rPr.getAttribute("i") === "1",
      underline: rPr.getAttribute("u") && rPr.getAttribute("u") !== "none",
    };
    const sz = rPr.getAttribute("sz");
    if (sz) out.size = Number(sz) / 100;
    const sf = XmlUtils.childByLocal(rPr, "solidFill");
    if (sf) out.color = readColor(sf, ctx);
    const latin = XmlUtils.childByLocal(rPr, "latin");
    if (latin) out.font = latin.getAttribute("typeface");
    return out;
  },

  // ── shape tree ───────────────────────────────────────────────────────────
  parseShapeTree(spTree, ctx) {
    const nodes = [];
    if (!spTree) return nodes;
    for (const child of spTree.children) {
      const ln = child.localName;
      if      (ln === "sp")           nodes.push(ShapeParser.parseShape(child, ctx));
      else if (ln === "pic")          nodes.push(ShapeParser.parsePic(child, ctx));
      else if (ln === "grpSp")        nodes.push({ kind: "group", xfrm: ShapeParser.parseGroupTransform(child), children: ShapeParser.parseShapeTree(child, ctx) });
      else if (ln === "graphicFrame") nodes.push(ShapeParser.parseGraphicFrame(child, ctx));
      else if (ln === "cxnSp")        nodes.push(ShapeParser.parseShape(child, ctx));
    }
    return nodes.filter(Boolean);
  },

  parseShape(sp, ctx) {
    const spPr     = XmlUtils.childByLocal(sp, "spPr");
    const txBody   = XmlUtils.childByLocal(sp, "txBody");
    const prstGeom = XmlUtils.childByLocal(spPr, "prstGeom");
    let prst = null, adj = {};
    if (prstGeom) {
      prst = prstGeom.getAttribute("prst");
      const avLst = XmlUtils.childByLocal(prstGeom, "avLst");
      if (avLst) {
        for (const gd of XmlUtils.childrenByLocal(avLst, "gd")) {
          const m = (gd.getAttribute("fmla") || "").match(/^val\s+(-?\d+)/);
          if (m) adj[gd.getAttribute("name")] = Number(m[1]);
        }
      }
    }
    return { kind: "shape", xfrm: ShapeParser.parseTransform(spPr), prst, adj, fill: ShapeParser.parseFill(spPr, ctx), line: ShapeParser.parseLine(spPr, ctx), text: ShapeParser.parseTextBody(txBody, ctx) };
  },

  parsePic(pic, ctx) {
    const spPr     = XmlUtils.childByLocal(pic, "spPr");
    const blipFill = XmlUtils.childByLocal(pic, "blipFill");
    const blip     = XmlUtils.childByLocal(blipFill, "blip");
    const rId      = blip ? XmlUtils.attrAny(blip, "embed") : null;
    return { kind: "pic", xfrm: ShapeParser.parseTransform(spPr), rId, line: ShapeParser.parseLine(spPr, ctx) };
  },

  parseGraphicFrame(gf, ctx) {
    const xfrm = XmlUtils.childByLocal(gf, "xfrm");
    let tx = null;
    if (xfrm) {
      const off = XmlUtils.childByLocal(xfrm, "off");
      const ext = XmlUtils.childByLocal(xfrm, "ext");
      tx = { x: off ? emu(off.getAttribute("x")) : 0, y: off ? emu(off.getAttribute("y")) : 0, w: ext ? emu(ext.getAttribute("cx")) : 0, h: ext ? emu(ext.getAttribute("cy")) : 0, rot: 0, flipH: false, flipV: false };
    }
    const tbl = XmlUtils.descendByPath(gf, ["graphic", "graphicData", "tbl"]);
    return tbl ? ShapeParser.parseTable(tbl, tx, ctx) : null;
  },

  parseTable(tbl, xfrm, ctx) {
    const tblGrid = XmlUtils.childByLocal(tbl, "tblGrid");
    const cols    = [];
    if (tblGrid) for (const gc of XmlUtils.childrenByLocal(tblGrid, "gridCol")) cols.push(emu(gc.getAttribute("w")));
    const rows = [];
    for (const tr of XmlUtils.childrenByLocal(tbl, "tr")) {
      const cells = [];
      for (const tc of XmlUtils.childrenByLocal(tr, "tc")) {
        cells.push({
          text:     ShapeParser.parseTextBody(XmlUtils.childByLocal(tc, "txBody"), ctx),
          fill:     ShapeParser.parseFill(XmlUtils.childByLocal(tc, "tcPr"), ctx),
          gridSpan: Number(tc.getAttribute("gridSpan") || 1),
          rowSpan:  Number(tc.getAttribute("rowSpan")  || 1),
          hMerge:   tc.getAttribute("hMerge") === "1",
          vMerge:   tc.getAttribute("vMerge") === "1",
        });
      }
      rows.push({ h: emu(tr.getAttribute("h")), cells });
    }
    return { kind: "table", xfrm, cols, rows };
  },
};
