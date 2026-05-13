"use strict";

// OOXML uses English Metric Units: 914400 EMU per inch, 96 px per inch.
const EMU_PER_PX = 9525;
const emu = v => (v ? Math.round(Number(v) / EMU_PER_PX * 100) / 100 : 0);
const deg = v => (v ? Number(v) / 60000 : 0);
const pct = v => (v ? Number(v) / 100000 : 0);

const ALIAS = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" };

const PRESET_COLORS = {
  black: "#000000", white: "#FFFFFF", red: "#FF0000", green: "#008000",
  blue: "#0000FF", yellow: "#FFFF00", cyan: "#00FFFF", magenta: "#FF00FF",
  gray: "#808080",
};

const COLOR_ELEMENTS = ["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr"];
const COLOR_MODS     = ["lumMod", "lumOff", "tint", "shade", "alpha", "satMod"];

function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16) };
}

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h, s, l = (max + min) / 2;
  if (max === min) { h = 0; s = 0; }
  else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r)      h = ((g - b) / d + (g < b ? 6 : 0));
    else if (max === g) h = (b - r) / d + 2;
    else                h = (r - g) / d + 4;
    h /= 6;
  }
  return { h, s, l };
}

function hslToRgb(h, s, l) {
  let r, g, b;
  if (s === 0) { r = g = b = l; }
  else {
    const hue2rgb = (p, q, t) => {
      if (t < 0) t += 1; if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }
  return { r: r * 255, g: g * 255, b: b * 255 };
}

function applyColorMods(hex, mods) {
  let { r, g, b } = hexToRgb(hex);
  let a = 1;
  for (const m of mods) {
    if (m.op === "alpha") { a = m.val; continue; }
    let hsl = rgbToHsl(r, g, b);
    if      (m.op === "lumMod") hsl.l *= m.val;
    else if (m.op === "lumOff") hsl.l = Math.min(1, hsl.l + m.val);
    else if (m.op === "tint")   hsl.l = hsl.l + (1 - hsl.l) * m.val;
    else if (m.op === "shade")  hsl.l = hsl.l * (1 - m.val);
    else if (m.op === "satMod") hsl.s = Math.min(1, hsl.s * m.val);
    const c = hslToRgb(hsl.h, hsl.s, hsl.l);
    r = c.r; g = c.g; b = c.b;
  }
  return a < 1
    ? `rgba(${r | 0},${g | 0},${b | 0},${a})`
    : `rgb(${r | 0},${g | 0},${b | 0})`;
}

function readColor(node, ctx) {
  if (!node) return null;
  let inner = node;
  if (!COLOR_ELEMENTS.includes(node.localName)) {
    for (const c of node.children) if (COLOR_ELEMENTS.includes(c.localName)) { inner = c; break; }
    if (inner === node) return null;
  }
  let hex = null;
  if (inner.localName === "srgbClr") {
    hex = "#" + inner.getAttribute("val");
  } else if (inner.localName === "schemeClr") {
    let key = inner.getAttribute("val");
    if (ctx?.colorMap?.[key]) key = ctx.colorMap[key];
    if (ALIAS[key]) key = ALIAS[key];
    hex = ctx?.theme?.scheme?.[key] || "#000000";
  } else if (inner.localName === "sysClr") {
    hex = "#" + (inner.getAttribute("lastClr") || "000000");
  } else if (inner.localName === "prstClr") {
    hex = PRESET_COLORS[inner.getAttribute("val")] || "#000000";
  } else if (inner.localName === "scrgbClr") {
    const r = pct(inner.getAttribute("r")) * 255;
    const g = pct(inner.getAttribute("g")) * 255;
    const b = pct(inner.getAttribute("b")) * 255;
    hex = "#" + [r, g, b].map(x => Math.round(x).toString(16).padStart(2, "0")).join("");
  }
  if (!hex) return null;
  const mods = [];
  for (const m of inner.children) {
    const v = m.getAttribute("val");
    if (v != null && COLOR_MODS.includes(m.localName)) mods.push({ op: m.localName, val: pct(v) });
  }
  return applyColorMods(hex, mods);
}

function parseTheme(xmlText) {
  const root       = XmlUtils.parse(xmlText);
  const elements   = XmlUtils.childByLocal(root, "themeElements");
  const clrScheme  = XmlUtils.childByLocal(elements, "clrScheme");
  const scheme     = {};
  if (clrScheme) {
    for (const c of clrScheme.children) {
      const srgb = XmlUtils.childByLocal(c, "srgbClr");
      const sys  = XmlUtils.childByLocal(c, "sysClr");
      if (srgb) scheme[c.localName] = "#" + srgb.getAttribute("val");
      else if (sys) scheme[c.localName] = "#" + (sys.getAttribute("lastClr") || "000000");
    }
  }
  return { scheme };
}

function parseClrMap(node) {
  if (!node) return null;
  const map = {};
  for (const a of node.attributes) map[a.name] = a.value;
  return map;
}
