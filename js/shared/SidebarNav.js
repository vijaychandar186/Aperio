"use strict";

(function () {
  const PAGES = [
    { key: "pptx", label: "Presentation", ext: ".pptx", href: "pptx.html" },
    { key: "xlsx", label: "Spreadsheet",  ext: ".xlsx", href: "xlsx.html" },
    { key: "docx", label: "Document",     ext: ".docx", href: "docx.html" },
  ];

  const current = location.pathname.split("/").pop().replace(".html", "");

  const nav = document.createElement("nav");
  nav.className = "mt-auto border-t border-gray-200 p-2";

  const heading = document.createElement("div");
  heading.className = "text-[10px] tracking-widest text-gray-400 uppercase mb-1.5 px-0.5";
  heading.textContent = "Open";
  nav.appendChild(heading);

  function link(href, label, ext) {
    const a = document.createElement("a");
    a.href = href;
    a.className = "flex items-center gap-2 px-2 py-1.5 text-xs text-gray-500 hover:text-black hover:bg-gray-50 border border-transparent hover:border-gray-200 mb-1";
    a.appendChild(document.createTextNode(label));
    if (ext) {
      const badge = document.createElement("span");
      badge.className = "ml-auto text-gray-300 text-[10px]";
      badge.textContent = ext;
      a.appendChild(badge);
    }
    return a;
  }

  nav.appendChild(link("../index.html", "Home", null));

  for (const p of PAGES) {
    if (p.key === current) continue;
    nav.appendChild(link(p.href, p.label, p.ext));
  }

  document.getElementById("sidebar").appendChild(nav);
})();
