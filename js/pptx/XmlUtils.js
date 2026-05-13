"use strict";

const XmlUtils = {
  parse(text) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    const err = doc.querySelector("parsererror");
    if (err) throw new Error("xml parse: " + err.textContent);
    return doc.documentElement;
  },

  // Walk by local-name to dodge namespace prefix variance.
  childByLocal(node, name) {
    if (!node) return null;
    for (const c of node.children) if (c.localName === name) return c;
    return null;
  },

  childrenByLocal(node, name) {
    const out = [];
    if (!node) return out;
    for (const c of node.children) if (c.localName === name) out.push(c);
    return out;
  },

  descendByPath(node, path) {
    let cur = node;
    for (const p of path) {
      cur = XmlUtils.childByLocal(cur, p);
      if (!cur) return null;
    }
    return cur;
  },

  attrAny(node, name) {
    if (!node) return null;
    for (const a of node.attributes) if (a.localName === name) return a.value;
    return node.getAttribute(name);
  },

  parseRels(xmlText) {
    if (!xmlText) return {};
    const root = XmlUtils.parse(xmlText);
    const out  = {};
    for (const r of root.children) {
      if (r.localName !== "Relationship") continue;
      out[r.getAttribute("Id")] = {
        type:   r.getAttribute("Type")   || "",
        target: r.getAttribute("Target") || "",
      };
    }
    return out;
  },

  resolvePath(baseDir, target) {
    if (!target) return "";
    if (target.startsWith("/")) return target.slice(1);
    const parts = (baseDir + "/" + target).split("/");
    const stack = [];
    for (const p of parts) {
      if (p === "" || p === ".") continue;
      if (p === "..") stack.pop();
      else stack.push(p);
    }
    return stack.join("/");
  },
};
