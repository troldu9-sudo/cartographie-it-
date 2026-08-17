/* ============================================================================
   extract.js — relève la géométrie réelle d'une planche HTML (index.html)
   pour la rejouer en formes natives PowerPoint (voir build.py).

   Injecté dans la page par Playwright ; renvoie un objet JSON décrivant :
     · shapes  : rectangles / ovales (fond, bordure, rayon)
     · texts   : une entrée par ligne visuelle (position mesurée, style)
     · wires   : courbes de Bézier des connecteurs + pastilles d'origine
   Toutes les coordonnées sont exprimées en pixels dans le repère 1600 × 900
   de la planche, converties ensuite en pouces par build.py.
   ========================================================================== */
window.__extractSlide = function (slideSelector) {

  const slide = document.querySelector(slideSelector);
  const S     = slide.getBoundingClientRect();
  const scale = S.width / slide.offsetWidth;          // la planche est mise à l'échelle par CSS

  const rel = r => ({
    x: (r.left - S.left) / scale,
    y: (r.top  - S.top ) / scale,
    w: r.width  / scale,
    h: r.height / scale
  });

  const out = { shapes: [], texts: [], wires: [], ports: [], w: 1600, h: 900 };

  /* ---------- utilitaires couleur -------------------------------------- */
  function parseColor(v) {
    if (!v || v === "transparent" || v === "none") return null;

    /* color-mix() est calculé par Chromium en « color(srgb r g b / a) », r/g/b sur 0–1 */
    const c = v.match(/color\(srgb\s+([^)]+)\)/);
    if (c) {
      const parts = c[1].replace("/", " ").split(/\s+/).filter(Boolean).map(parseFloat);
      const a = parts.length > 3 ? parts[3] : 1;
      if (a <= 0.004) return null;
      return { rgb: parts.slice(0, 3).map(n => Math.round(Math.max(0, Math.min(1, n)) * 255)), a };
    }

    const m = v.match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const p = m[1].replace(/[/,]/g, " ").split(/\s+/).filter(Boolean).map(parseFloat);
    const a = p.length > 3 ? p[3] : 1;
    if (a <= 0.004) return null;
    return { rgb: [Math.round(p[0]), Math.round(p[1]), Math.round(p[2])], a };
  }

  function radiusOf(cs, box) {
    const v = cs.borderTopLeftRadius || "0px";
    if (v.indexOf("%") > -1) return Math.min(box.w, box.h) / 2;
    return Math.min(parseFloat(v) || 0, Math.min(box.w, box.h) / 2);
  }

  /* ---------- découpage d'un nœud texte en lignes visuelles ------------ */
  function visualLines(node) {
    const raw = node.nodeValue;
    const groups = [];
    const re = /\S+/g;
    let m;
    while ((m = re.exec(raw)) !== null) {
      const range = document.createRange();
      range.setStart(node, m.index);
      range.setEnd(node, m.index + m[0].length);
      const rc = range.getBoundingClientRect();
      if (!rc.width && !rc.height) continue;
      let g = groups.find(g => Math.abs(g.top - rc.top) < 3);
      if (!g) { g = { top: rc.top, bottom: rc.bottom, left: rc.left, right: rc.right, words: [] }; groups.push(g); }
      g.top    = Math.min(g.top,    rc.top);
      g.bottom = Math.max(g.bottom, rc.bottom);
      g.left   = Math.min(g.left,   rc.left);
      g.right  = Math.max(g.right,  rc.right);
      g.words.push(m[0]);
    }
    return groups;
  }

  /* ---------- style typographique d'un élément -------------------------- */
  function typo(el) {
    const cs = getComputedStyle(el);
    return {
      size: parseFloat(cs.fontSize),
      bold: (parseInt(cs.fontWeight, 10) || 400) >= 650,
      color: parseColor(cs.color) || { rgb: [0, 0, 0], a: 1 },
      spc: cs.letterSpacing === "normal" ? 0 : (parseFloat(cs.letterSpacing) || 0)
    };
  }

  /* ---------- texte enrichi : un seul paragraphe, plusieurs runs --------
     Un élément qui mêle du texte et des enfants strictement inline (<b>, <em>…)
     doit rester UN paragraphe : le découper en boîtes séparées ferait perdre
     les espaces entre fragments, les métriques PowerPoint n'étant pas celles
     du navigateur. ------------------------------------------------------- */
  function inlineRuns(el) {
    const kids = Array.from(el.children);
    if (!kids.length) return null;
    /* un <br> impose des lignes distinctes : on repasse au relevé ligne par ligne */
    if (el.querySelector("br")) return null;
    const painted = e => {
      const cs = getComputedStyle(e);
      return parseColor(cs.backgroundColor) || parseFloat(cs.borderTopWidth) > 0;
    };
    for (const k of kids) {
      const kcs = getComputedStyle(k);
      if (!kcs.display.startsWith("inline")) return null;
      if (k.tagName.toLowerCase() === "svg") return null;
      /* un descendant peint (pastille, puce…) est exporté comme forme à sa
         position propre : le texte ne peut donc pas être fusionné ici */
      if (painted(k) || Array.from(k.querySelectorAll("*")).some(painted)) return null;
    }
    if (!el.textContent.trim()) return null;
    /* soit l'élément porte lui-même du texte, soit il enchaîne au moins deux
       fragments inline qui doivent rester dans le même paragraphe */
    const hasOwnText = [...el.childNodes].some(n => n.nodeType === 3 && n.nodeValue.trim());
    if (!hasOwnText && kids.length < 2) return null;

    const runs = [];
    for (const n of el.childNodes) {
      if (n.nodeType === 3) {
        const t = n.nodeValue.replace(/\s+/g, " ");
        if (t) runs.push(Object.assign({ t: t }, typo(el)));
      } else if (n.nodeType === 1) {
        const kcs = getComputedStyle(n);
        let t = n.textContent.replace(/\s+/g, " ");
        if (kcs.textTransform === "uppercase") t = t.toUpperCase();
        if (parseFloat(kcs.marginLeft) > 2) t = " " + t;
        if (parseFloat(kcs.marginRight) > 2) t = t + " ";
        if (t) runs.push(Object.assign({ t: t }, typo(n)));
      }
    }
    if (!runs.length) return null;
    runs[0].t = runs[0].t.replace(/^\s+/, "");
    runs[runs.length - 1].t = runs[runs.length - 1].t.replace(/\s+$/, "");
    return runs.filter(r => r.t.length);
  }

  /* ---------- parcours du DOM ------------------------------------------ */
  const walker = slide.querySelectorAll("*");
  const consumed = new Set();          // éléments dont le texte est déjà pris en charge

  for (const el of walker) {
    const tag = el.tagName.toLowerCase();
    if (tag === "svg" || el.closest("svg")) continue;      // traités séparément
    const cs = getComputedStyle(el);
    if (cs.display === "none" || cs.visibility === "hidden" || parseFloat(cs.opacity) === 0) continue;

    const box = rel(el.getBoundingClientRect());
    if (box.w < 0.5 || box.h < 0.5) continue;

    /* --- fond / bordure --------------------------------------------- */
    const fill    = parseColor(cs.backgroundColor);
    const bw      = parseFloat(cs.borderTopWidth) || 0;
    const stroke  = bw > 0 ? parseColor(cs.borderTopColor) : null;
    const isGrad  = cs.backgroundImage && cs.backgroundImage.indexOf("linear-gradient") === 0;

    if (fill || stroke || isGrad) {
      const rad = radiusOf(cs, box);
      out.shapes.push({
        kind: (cs.borderTopLeftRadius.indexOf("%") > -1 && Math.abs(box.w - box.h) < 1.5) ? "oval" : "rect",
        x: box.x, y: box.y, w: box.w, h: box.h,
        r: rad,
        fill: fill,
        gradient: isGrad ? el.getAttribute("data-grad") || "hub" : null,
        line: stroke ? { c: stroke, w: bw } : null,
        shadow: el.classList.contains("card") || el.classList.contains("hub"),
        halo: el.classList.contains("dot") ? parseColor(cs.backgroundColor) : null,
        id: el.id || null,
        cls: el.className && el.className.baseVal === undefined ? el.className : ""
      });
    }

    /* --- texte propre à l'élément ------------------------------------ */
    const align = cs.textAlign === "center" ? "c" : (cs.textAlign === "right" ? "r" : "l");

    let skip = false;
    for (let a = el.parentElement; a && a !== slide; a = a.parentElement) {
      if (consumed.has(a)) { skip = true; break; }
    }
    if (skip) continue;

    const runs = inlineRuns(el);
    if (runs) {
      consumed.add(el);
      /* boîte de contenu : le texte commence après le remplissage interne */
      const pl = parseFloat(cs.paddingLeft) || 0, pr = parseFloat(cs.paddingRight) || 0;
      const pt = parseFloat(cs.paddingTop) || 0, pb = parseFloat(cs.paddingBottom) || 0;
      out.texts.push({
        runs: runs,
        x: box.x + pl, y: box.y + pt,
        w: Math.max(box.w - pl - pr, 1), h: Math.max(box.h - pt - pb, 1),
        align: align,
        lh: parseFloat(cs.lineHeight) || 0,
        nowrap: cs.whiteSpace.indexOf("nowrap") > -1
      });
      continue;
    }

    const upper = cs.textTransform === "uppercase";
    const size = parseFloat(cs.fontSize);
    const weight = parseInt(cs.fontWeight, 10) || 400;
    const color = parseColor(cs.color) || { rgb: [0, 0, 0], a: 1 };
    const spc = cs.letterSpacing === "normal" ? 0 : (parseFloat(cs.letterSpacing) || 0);

    for (const node of el.childNodes) {
      if (node.nodeType !== 3 || !node.nodeValue.trim()) continue;
      for (const g of visualLines(node)) {
        const r = rel({ left: g.left, top: g.top, width: g.right - g.left, height: g.bottom - g.top });
        let t = g.words.join(" ");
        if (upper) t = t.toUpperCase();
        out.texts.push({
          t: t, x: r.x, y: r.y, w: r.w, h: r.h,
          size: size, bold: weight >= 650, color: color, spc: spc, align: align
        });
      }
    }
  }

  /* ---------- connecteurs (SVG) ---------------------------------------- */
  const board = slide.querySelector(".board");
  if (board) {
    const B = rel(board.getBoundingClientRect());
    slide.querySelectorAll("svg.wires path.wire").forEach(p => {
      const d = p.getAttribute("d").trim();
      const n = d.match(/-?\d+(\.\d+)?/g).map(Number);      // M x y C x y x y x y
      out.wires.push({
        pts: [[n[0] + B.x, n[1] + B.y], [n[2] + B.x, n[3] + B.y],
              [n[4] + B.x, n[5] + B.y], [n[6] + B.x, n[7] + B.y]],
        color: parseColor(getComputedStyle(p).stroke),
        width: parseFloat(getComputedStyle(p).strokeWidth),
        /* le pointillé se lit sur le style calculé, jamais sur un nom de classe :
           le CSS reste la source unique, rien à tenir synchronisé avec index.html */
        dashed: getComputedStyle(p).strokeDasharray !== "none"
      });
    });
    slide.querySelectorAll("svg.wires circle.port").forEach(c => {
      const r = parseFloat(c.getAttribute("r"));
      out.ports.push({
        x: parseFloat(c.getAttribute("cx")) + B.x,
        y: parseFloat(c.getAttribute("cy")) + B.y,
        r: r,
        fill: parseColor(getComputedStyle(c).fill),
        line: parseColor(getComputedStyle(c).stroke)
      });
    });
  }

  /* ---------- flèches de légende --------------------------------------- */
  slide.querySelectorAll('svg[data-pptx="arrow"]').forEach(sv => {
    const b = rel(sv.getBoundingClientRect());
    /* couleur relevée sur le SVG lui-même : « color » étant héritée, une flèche sans
       couleur propre rend comme avant, et une flèche teintée peut rester enfant
       direct de .item — l'en sortir casserait le garde-fou « svg » d'inlineRuns() */
    const col = parseColor(getComputedStyle(sv).color);
    out.wires.push({
      pts: [[b.x + 1, b.y + b.h / 2], [b.x + 10, b.y + b.h / 2],
            [b.x + 24, b.y + b.h / 2], [b.x + b.w, b.y + b.h / 2]],
      color: col, width: 2, dashed: sv.hasAttribute("data-dash")
    });
  });

  return out;
};
