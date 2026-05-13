"use strict";

// Depends on: XmlUtils, ColorUtils (parseTheme, parseClrMap, readColor, emu),
//             ShapeParser

const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const PptxLoader = {
  async load(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const readText = async path => { const f = zip.file(path); return f ? f.async("string") : null; };

    // 1. Presentation
    const presRels = XmlUtils.parseRels(await readText("ppt/_rels/presentation.xml.rels"));
    const presXml  = await readText("ppt/presentation.xml");
    if (!presXml) throw new Error("no ppt/presentation.xml — not a valid pptx");
    const presRoot = XmlUtils.parse(presXml);

    const sldSz  = XmlUtils.childByLocal(presRoot, "sldSz");
    const slideW = sldSz ? emu(sldSz.getAttribute("cx")) : 960;
    const slideH = sldSz ? emu(sldSz.getAttribute("cy")) : 540;

    const slideRIds = [];
    for (const sld of XmlUtils.childrenByLocal(XmlUtils.childByLocal(presRoot, "sldIdLst"), "sldId")) {
      const rid = sld.getAttributeNS(NS_R, "id");
      if (rid) slideRIds.push(rid);
    }

    // 2. Media → blob URLs
    const MIME = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp" };
    const mediaUrls   = {};
    const mediaFolder = zip.folder("ppt/media");
    const mediaJobs   = [];
    mediaFolder?.forEach((rel, file) => {
      if (file.dir) return;
      const ext  = rel.split(".").pop().toLowerCase();
      const mime = MIME[ext] || "application/octet-stream";
      mediaJobs.push(file.async("blob").then(b => { mediaUrls["ppt/media/" + rel] = URL.createObjectURL(new Blob([b], { type: mime })); }));
    });
    await Promise.all(mediaJobs);

    // 3. Theme
    const theme1Xml = await readText("ppt/theme/theme1.xml");
    const theme     = theme1Xml ? parseTheme(theme1Xml) : { scheme: {} };

    // 4. Master / layout cache
    const layoutCache  = new Map();
    const masterCache  = new Map();
    const readRels     = (dir, name) => readText(`${dir}/_rels/${name}.rels`).then(XmlUtils.parseRels);
    const findRel      = (rels, suffix) => { for (const k in rels) if (rels[k].type.endsWith(suffix)) return rels[k]; return null; };
    const extractBg    = (root, ctx) => {
      const bgPr = XmlUtils.descendByPath(root, ["cSld", "bg", "bgPr"]);
      const f    = bgPr ? ShapeParser.parseFill(bgPr, ctx) : null;
      return f?.kind === "solid" ? f.value : null;
    };

    const loadPart = async (path, cache, withMaster) => {
      if (cache.has(path)) return cache.get(path);
      const xml  = await readText(path);
      if (!xml) return null;
      const dir      = path.substring(0, path.lastIndexOf("/"));
      const fileName = path.split("/").pop();
      const rels     = await readRels(dir, fileName);
      const root     = XmlUtils.parse(xml);
      let master     = null;
      if (withMaster) {
        const rel = findRel(rels, "/slideMaster");
        if (rel) master = await loadPart(XmlUtils.resolvePath(dir, rel.target), masterCache, false);
      }
      const clrMap = parseClrMap(XmlUtils.childByLocal(root, "clrMap"));
      const ctx    = { theme, colorMap: clrMap || master?.ctx.colorMap || null };
      const part   = { dir, rels, ctx, master, bg: extractBg(root, ctx), nodes: ShapeParser.parseShapeTree(XmlUtils.descendByPath(root, ["cSld", "spTree"]), ctx) };
      cache.set(path, part);
      return part;
    };

    // 5. Slides
    const slides = [];
    for (const rId of slideRIds) {
      const rel = presRels[rId];
      if (!rel) continue;
      const slidePath = XmlUtils.resolvePath("ppt", rel.target);
      const xml       = await readText(slidePath);
      if (!xml) continue;
      const dir       = slidePath.substring(0, slidePath.lastIndexOf("/"));
      const fileName  = slidePath.split("/").pop();
      const slideRels = await readRels(dir, fileName);
      const layoutRel = findRel(slideRels, "/slideLayout");
      const layout    = layoutRel ? await loadPart(XmlUtils.resolvePath(dir, layoutRel.target), layoutCache, true) : null;
      const master    = layout?.master || null;
      const root      = XmlUtils.parse(xml);
      const slideClrMap = parseClrMap(XmlUtils.childByLocal(root, "clrMapOvr"));
      const colorMap    = slideClrMap || layout?.ctx.colorMap || master?.ctx.colorMap || null;
      const ctx         = { theme, colorMap };
      slides.push({ baseDir: dir, rels: slideRels, ctx, bg: extractBg(root, ctx), layout, master, nodes: ShapeParser.parseShapeTree(XmlUtils.descendByPath(root, ["cSld", "spTree"]), ctx) });
    }

    return { slideW, slideH, slides, mediaUrls, theme };
  },
};
