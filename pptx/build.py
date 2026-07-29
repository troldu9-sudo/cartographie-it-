#!/usr/bin/env python3
"""
build.py — génère « Cartographie-flux-de-donnees.pptx » à partir de index.html.

Le HTML est la source unique de vérité : la page est ouverte dans Chromium
(Playwright), la géométrie réelle de chaque planche est relevée par extract.js,
puis rejouée en **formes natives PowerPoint** (rectangles, ovales, courbes de
Bézier en géométrie personnalisée, zones de texte). Aucune image n'est produite :
tout reste éditable dans PowerPoint.

    python3 pptx/build.py [--html index.html] [--out dist/Cartographie-flux-de-donnees.pptx]
"""

import argparse
import json
import pathlib
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

ROOT = pathlib.Path(__file__).resolve().parent.parent

# La planche HTML fait 1600 × 900 px et représente une diapositive 13,333 × 7,5 po.
EMU_PER_PX = int(round(914400 * 13.3333333 / 1600))     # 7620 EMU / px
PT_PER_PX = 72 * 13.3333333 / 1600                      # 0,6 pt / px
FONT = "Calibri"                                        # présente partout, métriques stables

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def E(px):
    return Emu(int(round(px * EMU_PER_PX)))


def P(px):
    return Pt(round(px * PT_PER_PX, 1))


# --------------------------------------------------------------------------- #
#  Aides XML                                                                   #
# --------------------------------------------------------------------------- #
def frag(tag_xml):
    from pptx.oxml import parse_xml
    return parse_xml(f'<a:root xmlns:a="{A}">{tag_xml}</a:root>')[0]


def hexof(rgb):
    return "%02X%02X%02X" % tuple(rgb)


def _apply_alpha(clr_el, alpha):
    if alpha is not None and alpha < 0.999:
        clr_el.append(frag('<a:alpha val="%d"/>' % int(alpha * 100000)))


def solid_fill(shape, color):
    """color = {'rgb': [r,g,b], 'a': float} ou None."""
    if color is None:
        shape.fill.background()
        return
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color["rgb"])
    sf = shape._element.spPr.find(qn("a:solidFill"))
    _apply_alpha(sf.find(qn("a:srgbClr")), color.get("a", 1))


def line_style(shape, color, width_px):
    if color is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = RGBColor(*color["rgb"])
    shape.line.width = P(max(width_px, 0.75))
    ln = shape._element.spPr.find(qn("a:ln"))
    _apply_alpha(ln.find(qn("a:solidFill")).find(qn("a:srgbClr")), color.get("a", 1))


def no_shadow(shape):
    shape.shadow.inherit = False


def soft_shadow(shape, blur_px=26, dist_px=9, alpha=0.13, rgb=(0x0D, 0x15, 0x26)):
    shape.shadow.inherit = False
    eff = shape._element.spPr.find(qn("a:effectLst"))
    eff.append(frag(
        '<a:outerShdw blurRad="%d" dist="%d" dir="5400000" rotWithShape="0">'
        '<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw>'
        % (int(blur_px * EMU_PER_PX), int(dist_px * EMU_PER_PX), hexof(rgb), int(alpha * 100000))
    ))


def rounded_adj(shape, radius_px, w_px, h_px):
    """L'ajustement d'un rectangle arrondi vaut rayon / petit côté."""
    small = max(min(w_px, h_px), 0.01)
    try:
        shape.adjustments[0] = max(0.0, min(0.5, radius_px / small))
    except (IndexError, ValueError):
        pass


# --------------------------------------------------------------------------- #
#  Formes                                                                      #
# --------------------------------------------------------------------------- #
HUB_GRADIENT = ((0x18, 0x23, 0x3C), (0x0F, 0x17, 0x29), 0.62, 78)


def add_shape(shapes, s):
    if s["kind"] == "oval":
        sh = shapes.add_shape(MSO_SHAPE.OVAL, E(s["x"]), E(s["y"]), E(s["w"]), E(s["h"]))
    elif s["r"] > 0.6:
        sh = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(s["x"]), E(s["y"]), E(s["w"]), E(s["h"]))
        rounded_adj(sh, s["r"], s["w"], s["h"])
    else:
        sh = shapes.add_shape(MSO_SHAPE.RECTANGLE, E(s["x"]), E(s["y"]), E(s["w"]), E(s["h"]))

    if s.get("gradient"):
        c0, c1, pos, ang = HUB_GRADIENT
        sh.fill.gradient()
        stops = sh.fill.gradient_stops
        stops[0].color.rgb = RGBColor(*c0)
        stops[0].position = 0.0
        stops[1].color.rgb = RGBColor(*c1)
        stops[1].position = pos
        sh.fill.gradient_angle = ang
    else:
        solid_fill(sh, s["fill"])

    line_style(sh, s["line"]["c"] if s["line"] else None, s["line"]["w"] if s["line"] else 0)

    if s.get("shadow"):
        soft_shadow(sh, blur_px=30, dist_px=10, alpha=0.16)
    else:
        no_shadow(sh)

    sh.text_frame.word_wrap = False
    return sh


def _run(p, spec):
    run = p.add_run()
    run.text = spec["t"]
    f = run.font
    f.name = FONT
    f.size = P(spec["size"])
    f.bold = spec["bold"]
    f.color.rgb = RGBColor(*spec["color"]["rgb"])
    if abs(spec.get("spc", 0)) > 0.05:
        run.font._rPr.set("spc", str(int(round(spec["spc"] * PT_PER_PX * 100))))
    return run


def add_rich_text(shapes, t):
    """Paragraphe mêlant plusieurs styles : une seule zone de texte, plusieurs runs,
    pour que PowerPoint gère lui-même les espaces et le retour à la ligne."""
    box = shapes.add_textbox(E(t["x"] - 2), E(t["y"] - 2), E(t["w"] + 6), E(t["h"] + 8))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP

    p = tf.paragraphs[0]
    p.alignment = {"c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}.get(t["align"], PP_ALIGN.LEFT)
    if t.get("lh"):
        p.line_spacing = P(t["lh"])
    for spec in t["runs"]:
        _run(p, spec)
    return box


def add_text(shapes, t):
    if "runs" in t:
        return add_rich_text(shapes, t)

    slack = 90                                   # marge : les métriques PowerPoint diffèrent
    w = t["w"] + slack
    if t["align"] == "c":
        left = t["x"] + t["w"] / 2 - w / 2
    elif t["align"] == "r":
        left = t["x"] + t["w"] - w
    else:
        left = t["x"]

    box = shapes.add_textbox(E(left), E(t["y"] - 2), E(w), E(t["h"] + 4))
    tf = box.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.alignment = {"c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}.get(t["align"], PP_ALIGN.LEFT)
    _run(p, t)
    return box


def add_bezier(shapes, w):
    """Connecteur : courbe de Bézier cubique en géométrie personnalisée + flèche."""
    pts = w["pts"]
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    pad = 8
    x0, y0 = min(xs) - pad, min(ys) - pad
    W = max(max(xs) + pad - x0, 1.0)
    H = max(max(ys) + pad - y0, 1.0)

    sh = shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x0), E(y0), E(W), E(H))
    sh.fill.background()
    line_style(sh, w["color"], w["width"])
    no_shadow(sh)

    spPr = sh._element.spPr
    prst = spPr.find(qn("a:prstGeom"))
    idx = list(spPr).index(prst)
    spPr.remove(prst)

    def pt(i):
        return '<a:pt x="%d" y="%d"/>' % (
            int(round((pts[i][0] - x0) * EMU_PER_PX)),
            int(round((pts[i][1] - y0) * EMU_PER_PX)),
        )

    spPr.insert(idx, frag(
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="0" t="0" r="r" b="b"/><a:pathLst>'
        '<a:path w="%d" h="%d">'
        '<a:moveTo>%s</a:moveTo>'
        '<a:cubicBezTo>%s%s%s</a:cubicBezTo>'
        '</a:path></a:pathLst></a:custGeom>'
        % (int(round(W * EMU_PER_PX)), int(round(H * EMU_PER_PX)), pt(0), pt(1), pt(2), pt(3))
    ))

    ln = spPr.find(qn("a:ln"))
    if w.get("dashed"):
        ln.append(frag('<a:prstDash val="dash"/>'))
    ln.append(frag("<a:round/>"))
    ln.append(frag('<a:tailEnd type="triangle" w="med" len="med"/>'))
    return sh


def add_port(shapes, p):
    d = p["r"] * 2
    sh = shapes.add_shape(MSO_SHAPE.OVAL, E(p["x"] - p["r"]), E(p["y"] - p["r"]), E(d), E(d))
    solid_fill(sh, p["fill"])
    line_style(sh, p["line"], 2)
    no_shadow(sh)
    return sh


# --------------------------------------------------------------------------- #
#  Relevé de la page                                                           #
# --------------------------------------------------------------------------- #
def scrape(html_path, chromium=None):
    from playwright.sync_api import sync_playwright

    js = (ROOT / "pptx" / "extract.js").read_text(encoding="utf-8")
    url = pathlib.Path(html_path).resolve().as_uri()
    launch = {"executable_path": chromium} if chromium else {}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(**launch)
        page = browser.new_page(viewport={"width": 1700, "height": 1000})
        page.emulate_media(color_scheme="light")
        page.goto(url)
        page.wait_for_timeout(900)
        page.add_script_tag(content=js)
        planches = [page.evaluate("window.__extractSlide('#slide1')"),
                    page.evaluate("window.__extractSlide('#slide2')")]
        browser.close()
    return planches


# --------------------------------------------------------------------------- #
#  Assemblage                                                                  #
# --------------------------------------------------------------------------- #
def build(planches, out_path):
    prs = Presentation()
    prs.slide_width = Emu(int(round(13.3333333 * 914400)))
    prs.slide_height = Emu(int(round(7.5 * 914400)))
    blank = prs.slide_layouts[6]

    for data in planches:
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shapes = slide.shapes

        for s in data["shapes"]:
            if s.get("halo"):                       # pastille de direction : anneau translucide
                halo = dict(s["halo"])
                halo["a"] = 0.18
                ring = shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    E(s["x"] - 3.5), E(s["y"] - 3.5), E(s["w"] + 7), E(s["h"] + 7))
                solid_fill(ring, halo)
                ring.line.fill.background()
                no_shadow(ring)
            add_shape(shapes, s)

        for t in data["texts"]:
            add_text(shapes, t)

        for w in data["wires"]:
            add_bezier(shapes, w)

        for p in data["ports"]:
            add_port(shapes, p)

    out_path = pathlib.Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--html", default=str(ROOT / "index.html"))
    ap.add_argument("--out", default=str(ROOT / "dist" / "Cartographie-flux-de-donnees.pptx"))
    ap.add_argument("--chromium", default=None,
                    help="chemin d'un binaire Chromium (sinon celui de Playwright)")
    args = ap.parse_args()

    planches = scrape(args.html, args.chromium)
    n_sh = sum(len(p["shapes"]) + len(p["texts"]) + len(p["wires"]) + len(p["ports"]) for p in planches)
    out = build(planches, args.out)
    print(f"{out}  —  {len(planches)} planches, {n_sh} formes natives")


if __name__ == "__main__":
    sys.exit(main())
