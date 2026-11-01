"""Convert an SVG document into native, editable PowerPoint shapes (python-pptx).

Supported: ``g``/``svg``/``use``/``a``/``switch`` (as group shapes), ``rect`` (also with rounded
corners), ``circle``, ``ellipse``, ``line``, ``polyline``, ``polygon``, ``path`` (as freeforms with
bezier curves), ``text``/``tspan`` (as text boxes), transforms, presentation attributes, inline
styles and simple ``<style>`` sheets, opacity, dashes, linear/radial gradients and markers (as
native arrow heads). Images, filters, masks and clip paths are ignored.
"""
import math
import re

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu

NUM = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"
NUM_RE = re.compile(NUM)
IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

CONTAINERS = {"g", "svg", "a", "switch", "symbol"}
SKIP = {"defs", "marker", "clipPath", "mask", "pattern", "linearGradient", "radialGradient",
        "filter", "style", "title", "desc", "metadata", "script", "image", "foreignObject"}
PROPS = ["fill", "fill-opacity", "stroke", "stroke-width", "stroke-opacity", "stroke-dasharray",
         "stroke-linecap", "stroke-linejoin", "opacity", "display", "visibility", "font-family",
         "font-size", "font-weight", "font-style", "text-anchor", "dominant-baseline",
         "alignment-baseline", "text-decoration", "marker", "marker-start", "marker-end",
         "stop-color", "stop-opacity", "color"]
NOT_INHERITED = {"opacity", "display", "stop-color", "stop-opacity"}
GENERIC_FONTS = {"serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui",
                 "ui-sans-serif", "ui-serif", "ui-monospace", "-apple-system", "blinkmacsystemfont"}
UNITS = {"px": 1, "pt": 4 / 3, "pc": 16, "mm": 96 / 25.4, "cm": 96 / 2.54, "in": 96}
BACKGROUND_COVERAGE = 0.95  # rects covering this much of the figure are dropped as background


def _local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else None


# ---------------------------------------------------------------- geometry

def mat_mul(m, n):
    """The affine transform applying ``n`` first, then ``m``."""
    a1, b1, c1, d1, e1, f1 = m
    a2, b2, c2, d2, e2, f2 = n
    return (a1 * a2 + c1 * b2, b1 * a2 + d1 * b2, a1 * c2 + c1 * d2, b1 * c2 + d1 * d2,
            a1 * e2 + c1 * f2 + e1, b1 * e2 + d1 * f2 + f1)


def apply(m, x, y):
    return m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]


def mean_scale(m):
    return math.sqrt(abs(m[0] * m[3] - m[1] * m[2]))


def decompose(m):
    """(scale_x, scale_y, rotation in degrees) of ``m``, or None if it skews."""
    a, b, c, d = m[:4]
    sx = math.hypot(a, b)
    if sx == 0:
        return None
    sy = (a * d - b * c) / sx
    if abs(a * c + b * d) > 1e-6 * sx * max(abs(sy), 1e-9):
        return None
    return sx, abs(sy), math.degrees(math.atan2(b, a))


def parse_transform(s):
    m = IDENTITY
    for name, args in re.findall(r"([a-zA-Z]+)\s*\(([^)]*)\)", s or ""):
        v = [float(x) for x in NUM_RE.findall(args)]
        if name == "matrix" and len(v) == 6:
            t = tuple(v)
        elif name == "translate" and v:
            t = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0)
        elif name == "scale" and v:
            t = (v[0], 0, 0, v[1] if len(v) > 1 else v[0], 0, 0)
        elif name == "rotate" and v:
            a = math.radians(v[0])
            t = (math.cos(a), math.sin(a), -math.sin(a), math.cos(a), 0, 0)
            if len(v) == 3:
                t = mat_mul(mat_mul((1, 0, 0, 1, v[1], v[2]), t), (1, 0, 0, 1, -v[1], -v[2]))
        elif name == "skewX" and v:
            t = (1, 0, math.tan(math.radians(v[0])), 1, 0, 0)
        elif name == "skewY" and v:
            t = (1, math.tan(math.radians(v[0])), 0, 1, 0, 0)
        else:
            continue
        m = mat_mul(m, t)
    return m


def viewbox_transform(viewbox, x, y, width, height):
    """Map ``viewbox`` (min_x, min_y, w, h) into the given viewport, fitted and centered
    (the default ``preserveAspectRatio="xMidYMid meet"``)."""
    vx, vy, vw, vh = viewbox
    s = min(width / vw, height / vh)
    return (s, 0, 0, s, x + (width - vw * s) / 2 - vx * s, y + (height - vh * s) / 2 - vy * s)


def length(value, ref=0.0, font_size=16.0, default=0.0):
    if value is None:
        return default
    m = re.match(r"\s*(" + NUM + r")\s*([a-z%]*)", str(value))
    if not m:
        return default
    v, unit = float(m.group(1)), m.group(2)
    if unit == "%":
        return v * ref / 100
    if unit in ("em", "rem"):
        return v * font_size
    if unit == "ex":
        return v * font_size / 2
    return v * UNITS.get(unit, 1)


class _Scanner:
    def __init__(self, s):
        self.s, self.i = s, 0

    def _skip(self):
        while self.i < len(self.s) and self.s[self.i] in " \t\r\n,":
            self.i += 1

    def has_number(self):
        self._skip()
        return self.i < len(self.s) and (self.s[self.i].isdigit() or self.s[self.i] in "+-.")

    def command(self):
        self._skip()
        if self.i < len(self.s):
            self.i += 1
            return self.s[self.i - 1]
        return None

    def number(self):
        self._skip()
        m = NUM_RE.match(self.s, self.i)
        if not m:
            raise ValueError(f"invalid path data at {self.s[self.i:self.i + 10]!r}")
        self.i = m.end()
        return float(m.group(0))

    def flag(self):
        self._skip()
        self.i += 1
        return self.s[self.i - 1] == "1"


def arc_to_cubics(p0, rx, ry, phi, large, sweep, p1):
    """Bezier approximation of an SVG elliptical arc, as list of (c1, c2, end) tuples."""
    (x0, y0), (x1, y1) = p0, p1
    if p0 == p1:
        return []
    rx, ry = abs(rx), abs(ry)
    if rx == 0 or ry == 0:
        return [(p0, p1, p1)]
    phi = math.radians(phi)
    cos, sin = math.cos(phi), math.sin(phi)
    dx, dy = (x0 - x1) / 2, (y0 - y1) / 2
    x1p, y1p = cos * dx + sin * dy, -sin * dx + cos * dy
    lam = x1p ** 2 / rx ** 2 + y1p ** 2 / ry ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
    den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
    coef = math.sqrt(max(0.0, num / den)) * (-1 if large == sweep else 1)
    cxp, cyp = coef * rx * y1p / ry, -coef * ry * x1p / rx
    cx = cos * cxp - sin * cyp + (x0 + x1) / 2
    cy = sin * cxp + cos * cyp + (y0 + y1) / 2

    def angle(ux, uy, vx, vy):
        return math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)

    t1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    n = max(1, math.ceil(abs(dt) / (math.pi / 2) - 1e-9))
    step = dt / n
    k = 4 / 3 * math.tan(step / 4)

    def point(t):
        ex, ey = rx * math.cos(t), ry * math.sin(t)
        return cos * ex - sin * ey + cx, sin * ex + cos * ey + cy

    def tangent(t):
        ex, ey = -rx * math.sin(t), ry * math.cos(t)
        return cos * ex - sin * ey, sin * ex + cos * ey

    out = []
    for i in range(n):
        ta, tb = t1 + i * step, t1 + (i + 1) * step
        a, b, da, db = point(ta), point(tb), tangent(ta), tangent(tb)
        out.append(((a[0] + k * da[0], a[1] + k * da[1]), (b[0] - k * db[0], b[1] - k * db[1]), b))
    return out


def parse_path(d):
    """SVG path data -> list of ops with absolute coordinates: ("M", p), ("L", p),
    ("C", c1, c2, p) and ("Z",)."""
    sc = _Scanner(d or "")
    ops = []
    cur = start = (0.0, 0.0)
    cmd = prev = None
    ctrl = None  # last control point, for reflection by S / T
    while True:
        if not sc.has_number():
            cmd = sc.command()
            if cmd is None:
                break
        elif cmd is None:
            raise ValueError("path data does not start with a command")
        rel, c = cmd.islower(), cmd.upper()
        ox, oy = cur if rel else (0.0, 0.0)

        def pt():
            x = sc.number() + ox
            return x, sc.number() + oy

        if c == "M":
            cur = start = pt()
            ops.append(("M", cur))
            cmd = "l" if rel else "L"
        elif c == "L":
            cur = pt()
            ops.append(("L", cur))
        elif c == "H":
            cur = (sc.number() + ox, cur[1])
            ops.append(("L", cur))
        elif c == "V":
            cur = (cur[0], sc.number() + oy)
            ops.append(("L", cur))
        elif c in "CS":
            if c == "C":
                c1 = pt()
            else:
                c1 = (2 * cur[0] - ctrl[0], 2 * cur[1] - ctrl[1]) if prev in "CS" else cur
            c2, end = pt(), pt()
            ops.append(("C", c1, c2, end))
            cur, ctrl = end, c2
        elif c in "QT":
            if c == "Q":
                q = pt()
            else:
                q = (2 * cur[0] - ctrl[0], 2 * cur[1] - ctrl[1]) if prev in "QT" else cur
            end = pt()
            ops.append(("C", (cur[0] + 2 / 3 * (q[0] - cur[0]), cur[1] + 2 / 3 * (q[1] - cur[1])),
                        (end[0] + 2 / 3 * (q[0] - end[0]), end[1] + 2 / 3 * (q[1] - end[1])), end))
            cur, ctrl = end, q
        elif c == "A":
            rx, ry, phi = sc.number(), sc.number(), sc.number()
            large, sweep = sc.flag(), sc.flag()
            end = pt()
            ops.extend(("C",) + seg for seg in arc_to_cubics(cur, rx, ry, phi, large, sweep, end))
            cur = end
        elif c == "Z":
            ops.append(("Z",))
            cur, cmd = start, None
        else:
            raise ValueError(f"unknown path command {cmd!r}")
        prev = c
    return ops


def transform_ops(ops, m):
    return [(op[0],) + tuple(apply(m, *p) for p in op[1:]) for op in ops]


def ellipse_ops(cx, cy, rx, ry):
    ops = [("M", (cx + rx, cy))]
    for a, b in [((cx + rx, cy), (cx, cy + ry)), ((cx, cy + ry), (cx - rx, cy)),
                 ((cx - rx, cy), (cx, cy - ry)), ((cx, cy - ry), (cx + rx, cy))]:
        ops.extend(("C",) + seg for seg in arc_to_cubics(a, rx, ry, 0, False, True, b))
    return ops + [("Z",)]


def rect_ops(x, y, w, h, rx, ry):
    if not rx or not ry:
        return [("M", (x, y)), ("L", (x + w, y)), ("L", (x + w, y + h)), ("L", (x, y + h)), ("Z",)]
    d = (f"M{x + rx},{y} H{x + w - rx} A{rx},{ry} 0 0 1 {x + w},{y + ry} V{y + h - ry} "
         f"A{rx},{ry} 0 0 1 {x + w - rx},{y + h} H{x + rx} A{rx},{ry} 0 0 1 {x},{y + h - ry} "
         f"V{y + ry} A{rx},{ry} 0 0 1 {x + rx},{y} Z")
    return parse_path(d)


# ---------------------------------------------------------------- styles

def parse_declarations(s):
    out = {}
    for decl in (s or "").split(";"):
        if ":" in decl:
            k, v = decl.split(":", 1)
            out[k.strip().lower()] = v.replace("!important", "").strip()
    return out


def parse_css(root):
    """Rules of all ``<style>`` elements as (tag, id, classes, declarations), sorted by
    specificity. Of compound selectors, only the last part is considered."""
    rules = []
    for el in root.iter():
        if _local(el.tag) != "style":
            continue
        text = re.sub(r"/\*.*?\*/", "", "".join(el.itertext()), flags=re.DOTALL)
        for selectors, body in re.findall(r"([^{}]+)\{([^{}]*)\}", text):
            decls = parse_declarations(body)
            for sel in selectors.split(","):
                parts = re.split(r"[\s>+~]+", sel.strip())
                m = re.fullmatch(r"(\*|[\w-]+)?(?:#([\w-]+))?((?:\.[\w-]+)*)", parts[-1])
                if not parts[-1] or not m:
                    continue
                tag, id_ = m.group(1), m.group(2)
                classes = {c for c in m.group(3).split(".") if c}
                specificity = (bool(id_), len(classes), bool(tag and tag != "*"))
                rules.append((specificity, len(rules), tag, id_, classes, decls))
    rules.sort(key=lambda r: r[:2])
    return [r[2:] for r in rules]


def parse_color(value):
    """(r, g, b), alpha of a CSS color, or None for none/transparent/unknown."""
    v = (value or "").strip().lower()
    if v in ("", "none", "transparent"):
        return None
    m = re.fullmatch(r"#([0-9a-f]{3,8})", v)
    if m:
        h = m.group(1)
        if len(h) in (3, 4):
            h = "".join(ch * 2 for ch in h)
        if len(h) not in (6, 8):
            return None
        alpha = int(h[6:8], 16) / 255 if len(h) == 8 else 1.0
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)), alpha
    m = re.fullmatch(r"rgba?\(([^)]*)\)", v)
    if m:
        parts = [p for p in re.split(r"[\s,/]+", m.group(1)) if p]
        if len(parts) < 3:
            return None
        rgb = tuple(max(0, min(255, round(length(p, 255)))) for p in parts[:3])
        alpha = length(parts[3], 1) if len(parts) > 3 else 1.0
        return rgb, max(0.0, min(1.0, alpha))
    try:
        from PIL import ImageColor

        rgba = ImageColor.getrgb(v)
    except (ImportError, ValueError):
        return None
    return tuple(rgba[:3]), (rgba[3] / 255 if len(rgba) > 3 else 1.0)


def fraction(value, default=1.0):
    """An opacity like "0.5" or "50%", clipped to 0..1."""
    return max(0.0, min(1.0, length(value, 1, default=default)))


def _remove_style(shape):
    """Remove the theme style reference (theme fill, line and shadow effects) so that only the
    explicit formatting applies."""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def _set_alpha(solid_fill, alpha):
    if solid_fill is not None and alpha < 1:
        etree.SubElement(solid_fill[0], qn("a:alpha"), val=str(int(round(alpha * 100000))))


# ---------------------------------------------------------------- conversion

class SvgConverter:
    def __init__(self, root):
        self.root = root
        self.ids = {el.get("id"): el for el in root.iter() if _local(el.tag) and el.get("id")}
        self.css = parse_css(root)
        self.viewport = (100.0, 100.0)  # size of the root viewBox, for percentages
        self.figure_extent = None  # (width, height) of the root viewport in EMU

    def style(self, el, parent):
        st = {k: v for k, v in parent.items() if k not in NOT_INHERITED}
        own = {k: el.get(k) for k in PROPS if el.get(k) is not None}
        classes = set((el.get("class") or "").split())
        for tag, id_, cls, decls in self.css:
            if ((not tag or tag in ("*", _local(el.tag))) and (not id_ or id_ == el.get("id"))
                    and cls <= classes):
                own.update(decls)
        own.update(parse_declarations(el.get("style")))
        if "marker" in own:
            for k in ("marker-start", "marker-end"):
                own.setdefault(k, own["marker"])
        for k, v in own.items():
            v = v.strip()
            if v == "inherit" or k not in PROPS:
                continue
            if k == "font-size":
                parent_size = float(parent.get("font-size", 16))
                sizes = {"small": 13, "medium": 16, "large": 18, "x-large": 24, "xx-large": 32}
                v = str(sizes.get(v) or length(v, parent_size, parent_size, parent_size))
            st[k] = v
        return st

    def href(self, el):
        ref = el.get("href") or el.get("{http://www.w3.org/1999/xlink}href") or ""
        return self.ids.get(ref[1:]) if ref.startswith("#") else None

    def url_target(self, value):
        m = re.match(r"\s*url\(\s*['\"]?#([^)'\"]+)['\"]?\s*\)", value or "")
        return self.ids.get(m.group(1)) if m else None

    # ---- paint

    def gradient_stops(self, grad, depth=0):
        stops = [s for s in grad if _local(s.tag) == "stop"]
        if not stops and depth < 10 and self.href(grad) is not None:
            return self.gradient_stops(self.href(grad), depth + 1)
        out = []
        for s in stops:
            st = self.style(s, {})
            color = parse_color(st.get("stop-color", "black"))
            if color is None:
                continue
            offset = max(0.0, min(1.0, length(s.get("offset", "0"), 1)))
            out.append((offset, color[0], color[1] * fraction(st.get("stop-opacity"))))
        return out

    def paint(self, value):
        """('solid', rgb, alpha), ('gradient', element, stops) or None."""
        grad = self.url_target(value)
        if grad is not None and _local(grad.tag) in ("linearGradient", "radialGradient"):
            stops = self.gradient_stops(grad)
            if len(stops) == 1:
                return "solid", stops[0][1], stops[0][2]
            if stops:
                return "gradient", grad, stops
        if value and value.strip().startswith("url("):
            value = value.split(")", 1)[1]  # fallback color after the url
        color = parse_color(value)
        return ("solid",) + color if color else None

    @staticmethod
    def solid_of(paint):
        """A single color for a paint: the mean color of a gradient."""
        if paint is None or paint[0] == "solid":
            return paint
        stops = paint[2]
        rgb = tuple(round(sum(s[1][i] for s in stops) / len(stops)) for i in range(3))
        return "solid", rgb, sum(s[2] for s in stops) / len(stops)

    def apply_fill(self, shape, st, opacity):
        paint = self.paint(st.get("fill", "black"))
        alpha = opacity * fraction(st.get("fill-opacity"))
        if paint is None or alpha <= 0:
            shape.fill.background()
        elif paint[0] == "solid":
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(*paint[1])
            _set_alpha(shape._element.spPr.find(qn("a:solidFill")), paint[2] * alpha)
        else:
            self.apply_gradient(shape, paint[1], paint[2], alpha)

    def apply_gradient(self, shape, grad, stops, alpha):
        shape.fill.gradient()
        grad_fill = shape._element.spPr.find(qn("a:gradFill"))
        for child in list(grad_fill):
            grad_fill.remove(child)
        grad_fill.set("rotWithShape", "1")
        gs_lst = etree.SubElement(grad_fill, qn("a:gsLst"))
        for offset, rgb, a in stops:
            gs = etree.SubElement(gs_lst, qn("a:gs"), pos=str(int(round(offset * 100000))))
            clr = etree.SubElement(gs, qn("a:srgbClr"), val="%02X%02X%02X" % rgb)
            if a * alpha < 1:
                etree.SubElement(clr, qn("a:alpha"), val=str(int(round(a * alpha * 100000))))
        if _local(grad.tag) == "radialGradient":
            path = etree.SubElement(grad_fill, qn("a:path"), path="circle")
            etree.SubElement(path, qn("a:fillToRect"), l="50000", t="50000", r="50000", b="50000")
        else:
            x1, y1 = length(grad.get("x1", "0"), 1), length(grad.get("y1", "0"), 1)
            x2, y2 = length(grad.get("x2", "1"), 1), length(grad.get("y2", "0"), 1)
            angle = math.degrees(math.atan2(y2 - y1, x2 - x1)) % 360
            etree.SubElement(grad_fill, qn("a:lin"), ang=str(int(round(angle * 60000))), scaled="0")

    def apply_stroke(self, shape, st, opacity, m, markers=False):
        line = shape.line
        paint = self.solid_of(self.paint(st.get("stroke", "none")))
        alpha = opacity * fraction(st.get("stroke-opacity"))
        width = length(st.get("stroke-width"), default=1.0)
        if paint is None or alpha <= 0 or width <= 0:
            line.fill.background()
            return
        line.width = Emu(max(1, int(round(width * mean_scale(m)))))
        line.color.rgb = RGBColor(*paint[1])
        ln = shape._element.spPr.find(qn("a:ln"))
        _set_alpha(ln.find(qn("a:solidFill")), paint[2] * alpha)
        dashes = [length(v) for v in re.split(r"[\s,]+", st.get("stroke-dasharray", "none").strip())
                  if NUM_RE.match(v)]
        if any(dashes):
            ratio = dashes[0] / width
            line.dash_style = (MSO_LINE_DASH_STYLE.ROUND_DOT if ratio <= 1.5 else
                               MSO_LINE_DASH_STYLE.DASH if ratio <= 4 else MSO_LINE_DASH_STYLE.LONG_DASH)
        cap = {"round": "rnd", "square": "sq"}.get(st.get("stroke-linecap"))
        if cap:
            ln.set("cap", cap)
        join = {"round": "a:round", "bevel": "a:bevel"}.get(st.get("stroke-linejoin"))
        if join:
            etree.SubElement(ln, qn(join))
        if markers:
            for prop, tag in (("marker-start", "a:headEnd"), ("marker-end", "a:tailEnd")):
                arrow = self.arrow_head(st.get(prop), width)
                if arrow:
                    etree.SubElement(ln, qn(tag), type=arrow[0], w=arrow[1], len=arrow[1])

    def arrow_head(self, value, stroke_width):
        """(type, size) of a native arrow head resembling the marker, or None."""
        marker = self.url_target(value)
        if marker is None or _local(marker.tag) != "marker":
            return None
        kinds = {_local(el.tag) for el in marker.iter()}
        kind = "oval" if kinds & {"circle", "ellipse"} else "diamond" if "rect" in kinds else "triangle"
        size = length(marker.get("markerWidth"), default=3.0)
        if marker.get("markerUnits") == "userSpaceOnUse":
            size /= max(stroke_width, 1e-9)
        return kind, "sm" if size <= 2.5 else "med" if size <= 4.5 else "lg"

    # ---- shapes

    def finish(self, shape, st, opacity, m, fill=True, markers=False):
        _remove_style(shape)
        if fill:
            self.apply_fill(shape, st, opacity)
        self.apply_stroke(shape, st, opacity, m, markers)
        return shape

    def add_freeform(self, shapes, ops, st, opacity, m):
        ops = transform_ops(ops, m)
        points = [p for op in ops for p in op[1:]]
        if not points:
            return None
        xs, ys = [p[0] for p in points], [p[1] for p in points]
        x0, y0 = int(math.floor(min(xs))), int(math.floor(min(ys)))
        builder = shapes.build_freeform(x0, y0)
        builder.add_line_segments([(max(int(math.ceil(max(xs))), x0 + 1),
                                    max(int(math.ceil(max(ys))), y0 + 1))], close=False)
        shape = builder.convert_to_shape()
        w, h = shape.width, shape.height
        path_lst = shape._element.spPr.find(qn("a:custGeom")).find(qn("a:pathLst"))
        for child in list(path_lst):
            path_lst.remove(child)
        path = etree.SubElement(path_lst, qn("a:path"), w=str(w), h=str(h))
        tags = {"M": "a:moveTo", "L": "a:lnTo", "C": "a:cubicBezTo"}
        for op in ops:
            if op[0] == "Z":
                etree.SubElement(path, qn("a:close"))
                continue
            el = etree.SubElement(path, qn(tags[op[0]]))
            for x, y in op[1:]:
                etree.SubElement(el, qn("a:pt"), x=str(int(round(x - shape.left))),
                                 y=str(int(round(y - shape.top))))
        closed = any(op[0] == "Z" for op in ops)
        return self.finish(shape, st, opacity, m, markers=not closed)

    def add_line(self, shapes, p0, p1, st, opacity, m):
        (x0, y0), (x1, y1) = apply(m, *p0), apply(m, *p1)
        shape = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, int(round(x0)), int(round(y0)),
                                     int(round(x1)), int(round(y1)))
        return self.finish(shape, st, opacity, m, fill=False, markers=True)

    def add_box(self, shapes, kind, x, y, w, h, st, opacity, m, adjustment=None):
        """An autoshape for an axis-aligned box in user space, or None if ``m`` skews it."""
        dec = decompose(m)
        if dec is None:
            return None
        sx, sy, angle = dec
        cx, cy = apply(m, x + w / 2, y + h / 2)
        width, height = w * sx, h * sy
        shape = shapes.add_shape(kind, int(round(cx - width / 2)), int(round(cy - height / 2)),
                                 max(1, int(round(width))), max(1, int(round(height))))
        if adjustment is not None:
            shape.adjustments[0] = adjustment
        if abs(angle) > 0.01:
            shape.rotation = angle % 360
        return self.finish(shape, st, opacity, m)

    def is_background(self, x, y, w, h, m):
        corners = [apply(m, x, y), apply(m, x + w, y + h)]
        width = abs(corners[1][0] - corners[0][0])
        height = abs(corners[1][1] - corners[0][1])
        fw, fh = self.figure_extent
        return width >= BACKGROUND_COVERAGE * fw and height >= BACKGROUND_COVERAGE * fh

    def rect(self, el, shapes, st, opacity, m):
        vw, vh = self.viewport
        x, y = length(el.get("x"), vw), length(el.get("y"), vh)
        w, h = length(el.get("width"), vw), length(el.get("height"), vh)
        if w <= 0 or h <= 0 or self.is_background(x, y, w, h, m):
            return
        rx, ry = el.get("rx"), el.get("ry")
        rx, ry = length(rx if rx is not None else ry, vw), length(ry if ry is not None else rx, vh)
        rx, ry = min(rx, w / 2), min(ry, h / 2)
        if rx > 0 and ry > 0:
            dec = decompose(m)
            radius = rx * dec[0] if dec else 0
            adjustment = min(0.5, radius / max(1e-9, min(w * dec[0], h * dec[1]))) if dec else 0
            shape = self.add_box(shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, st, opacity, m,
                                 adjustment)
        else:
            shape = self.add_box(shapes, MSO_SHAPE.RECTANGLE, x, y, w, h, st, opacity, m)
        if shape is None:
            self.add_freeform(shapes, rect_ops(x, y, w, h, rx, ry), st, opacity, m)

    def ellipse(self, el, shapes, st, opacity, m):
        vw, vh = self.viewport
        cx, cy = length(el.get("cx"), vw), length(el.get("cy"), vh)
        if _local(el.tag) == "circle":
            rx = ry = length(el.get("r"), math.hypot(vw, vh) / math.sqrt(2))
        else:
            rx, ry = length(el.get("rx"), vw), length(el.get("ry"), vh)
        if rx <= 0 or ry <= 0:
            return
        if self.add_box(shapes, MSO_SHAPE.OVAL, cx - rx, cy - ry, 2 * rx, 2 * ry, st, opacity, m) is None:
            self.add_freeform(shapes, ellipse_ops(cx, cy, rx, ry), st, opacity, m)

    def poly(self, el, shapes, st, opacity, m):
        v = [float(n) for n in NUM_RE.findall(el.get("points", ""))]
        points = list(zip(v[0::2], v[1::2]))
        if len(points) < 2:
            return
        closed = _local(el.tag) == "polygon"
        if len(points) == 2 and not closed:
            self.add_line(shapes, points[0], points[1], st, opacity, m)
            return
        ops = [("M", points[0])] + [("L", p) for p in points[1:]] + ([("Z",)] if closed else [])
        self.add_freeform(shapes, ops, st, opacity, m)

    def path(self, el, shapes, st, opacity, m):
        ops = parse_path(el.get("d"))
        if [op[0] for op in ops] == ["M", "L"]:
            self.add_line(shapes, ops[0][1], ops[1][1], st, opacity, m)
        elif any(op[0] != "M" for op in ops):
            self.add_freeform(shapes, ops, st, opacity, m)

    # ---- text

    def text(self, el, shapes, st, opacity, m):
        """Each line (the text and each tspan having own x/y/dy) becomes a text box."""
        fs = float(st.get("font-size", 16))
        x = length((el.get("x") or "0").split()[0] if el.get("x") else None, font_size=fs)
        y = length((el.get("y") or "0").split()[0] if el.get("y") else None, font_size=fs)
        x += length(el.get("dx"), font_size=fs)
        y += length(el.get("dy"), font_size=fs)
        lines = [{"x": x, "y": y, "style": st, "runs": []}]

        def visit(node, node_style):
            if node.text:
                lines[-1]["runs"].append((node.text, node_style))
            for child in node:
                if _local(child.tag) == "tspan" and self.visible(self.style(child, node_style)):
                    cst = self.style(child, node_style)
                    cfs = float(cst.get("font-size", 16))
                    if any(child.get(k) is not None for k in ("x", "y", "dy")):
                        last = lines[-1]
                        nx, ny = child.get("x"), child.get("y")
                        lines.append({
                            "x": length(nx.split()[0], font_size=cfs) if nx else last["x"],
                            "y": (length(ny.split()[0], font_size=cfs) if ny else last["y"])
                            + length(child.get("dy"), font_size=cfs),
                            "style": cst, "runs": []})
                    visit(child, cst)
                if child.tail:
                    lines[-1]["runs"].append((child.tail, node_style))

        visit(el, st)
        for line in lines:
            runs = [(re.sub(r"\s+", " ", t), s) for t, s in line["runs"]]
            if runs:
                runs[0] = (runs[0][0].lstrip(), runs[0][1])
                runs[-1] = (runs[-1][0].rstrip(), runs[-1][1])
            runs = [(t, s) for t, s in runs if t]
            if runs:
                self.add_text_line(shapes, line["x"], line["y"], line["style"], runs, opacity, m)

    def add_text_line(self, shapes, x, y, st, runs, opacity, m):
        dec = decompose(m) or (mean_scale(m), mean_scale(m), 0.0)
        sx, sy, angle = dec
        fs = max(float(s.get("font-size", 16)) for _, s in runs)
        width = sum(len(t) * float(s.get("font-size", 16)) * (0.62 if self.bold(s) else 0.56)
                    for t, s in runs)
        anchor = st.get("text-anchor", "start")
        cx = x + {"middle": 0, "end": -width / 2}.get(anchor, width / 2)
        baseline = st.get("dominant-baseline") or st.get("alignment-baseline") or ""
        if baseline in ("middle", "central"):
            cy = y
        elif baseline in ("hanging", "text-before-edge", "text-top"):
            cy = y + 0.4 * fs
        else:  # alphabetic: the center of the line lies about 0.35 em above the baseline
            cy = y - 0.35 * fs
        cx, cy = apply(m, cx, cy)
        w, h = max(1, width * sx), max(1, 1.3 * fs * sy)
        box = shapes.add_textbox(int(round(cx - w / 2)), int(round(cy - h / 2)),
                                 int(round(w)), int(round(h)))
        if abs(angle) > 0.01:
            box.rotation = angle % 360
        tf = box.text_frame
        tf.word_wrap = False
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = {"middle": PP_ALIGN.CENTER, "end": PP_ALIGN.RIGHT}.get(anchor, PP_ALIGN.LEFT)
        for t, s in runs:
            run = p.add_run()
            run.text = t
            font = run.font
            font.size = Emu(max(12700, int(round(float(s.get("font-size", 16)) * sy))))
            font.bold = self.bold(s)
            font.italic = s.get("font-style") in ("italic", "oblique")
            if "underline" in s.get("text-decoration", ""):
                font.underline = True
            family = self.font_family(s)
            if family:
                font.name = family
            paint = self.solid_of(self.paint(s.get("fill", "black")))
            if paint is not None:
                font.color.rgb = RGBColor(*paint[1])
                alpha = paint[2] * opacity * fraction(s.get("fill-opacity"))
                _set_alpha(run._r.get_or_add_rPr().find(qn("a:solidFill")), alpha)
        return box

    @staticmethod
    def bold(st):
        w = st.get("font-weight", "normal")
        return w in ("bold", "bolder") or (w.isdigit() and int(w) >= 600)

    @staticmethod
    def font_family(st):
        for name in st.get("font-family", "").split(","):
            name = name.strip().strip("'\"")
            if name and name.lower() not in GENERIC_FONTS:
                return name
        return None

    # ---- tree traversal

    @staticmethod
    def visible(st):
        return st.get("display") != "none" and st.get("visibility") not in ("hidden", "collapse")

    def drawable_children(self, el):
        """Child elements that may render; symbols only render when referenced by <use>."""
        return [c for c in el if _local(c.tag) and _local(c.tag) not in SKIP | {"symbol"}]

    def convert(self, el, shapes, m, parent_style, opacity=1.0, depth=0):
        tag = _local(el.tag)
        if tag is None or tag in SKIP or depth > 50:
            return
        st = self.style(el, parent_style)
        if not self.visible(st):
            return
        opacity *= fraction(st.get("opacity"))
        m = mat_mul(m, parse_transform(el.get("transform")))
        if tag in CONTAINERS or tag == "use":
            if tag == "use":
                target = self.href(el)
                if target is None:
                    return
                m = mat_mul(m, (1, 0, 0, 1, length(el.get("x")), length(el.get("y"))))
                if _local(target.tag) in ("svg", "symbol"):
                    # the referenced viewport gets the size given by <use>
                    st = self.style(target, st)
                    m = self.nested_viewport(target, m, el.get("width"), el.get("height"))
                    children = self.drawable_children(target)
                else:
                    children = [target]
            else:
                if tag in ("svg", "symbol") and depth > 0:
                    m = self.nested_viewport(el, m)
                children = self.drawable_children(el)
            group = shapes.add_group_shape() if tag != "use" and len(children) > 1 else None
            for child in children:
                self.convert(child, group.shapes if group else shapes, m, st, opacity, depth + 1)
            if group is not None:
                if len(group.shapes):
                    group._element.recalculate_extents()
                else:
                    group._element.getparent().remove(group._element)
            shapes._recalculate_extents()
        else:
            try:
                self.convert_shape(tag, el, shapes, st, opacity, m)
            except (ValueError, ZeroDivisionError, IndexError):
                pass  # skip a malformed element, keep the rest of the figure

    def convert_shape(self, tag, el, shapes, st, opacity, m):
        if tag == "rect":
            self.rect(el, shapes, st, opacity, m)
        elif tag in ("circle", "ellipse"):
            self.ellipse(el, shapes, st, opacity, m)
        elif tag == "line":
            p0 = (length(el.get("x1")), length(el.get("y1")))
            p1 = (length(el.get("x2")), length(el.get("y2")))
            self.add_line(shapes, p0, p1, st, opacity, m)
        elif tag in ("polyline", "polygon"):
            self.poly(el, shapes, st, opacity, m)
        elif tag == "path":
            self.path(el, shapes, st, opacity, m)
        elif tag == "text":
            self.text(el, shapes, st, opacity, m)

    def nested_viewport(self, el, m, width=None, height=None):
        vw, vh = self.viewport
        box = [float(v) for v in NUM_RE.findall(el.get("viewBox") or "")]
        w = length(width or el.get("width"), vw, default=box[2] if len(box) == 4 else vw)
        h = length(height or el.get("height"), vh, default=box[3] if len(box) == 4 else vh)
        if _local(el.tag) == "svg":
            m = mat_mul(m, (1, 0, 0, 1, length(el.get("x"), vw), length(el.get("y"), vh)))
        if len(box) == 4 and box[2] > 0 and box[3] > 0:
            m = mat_mul(m, viewbox_transform(box, 0, 0, w, h))
        return m


def root_viewbox(root, default):
    box = [float(v) for v in NUM_RE.findall(root.get("viewBox") or "")]
    if len(box) == 4 and box[2] > 0 and box[3] > 0:
        return tuple(box)
    w, h = length(root.get("width")), length(root.get("height"))
    return (0, 0, w, h) if w > 0 and h > 0 else default


def add_svg_to_slide(prs, slide, svg, name="Polished figure"):
    """Add the SVG as a group of native shapes to the slide, with the SVG's viewBox fitted and
    centered within the slide. Returns the group shape."""
    root = etree.fromstring(svg.encode("utf-8"), etree.XMLParser(recover=True, resolve_entities=False))
    if root is None or _local(root.tag) != "svg":
        raise ValueError("Not an SVG document")
    converter = SvgConverter(root)
    viewbox = root_viewbox(root, (0, 0, prs.slide_width / 9525, prs.slide_height / 9525))
    converter.viewport = viewbox[2:]
    m = viewbox_transform(viewbox, 0, 0, prs.slide_width, prs.slide_height)
    converter.figure_extent = (viewbox[2] * m[0], viewbox[3] * m[3])
    group = slide.shapes.add_group_shape()
    group.name = name
    for child in converter.drawable_children(root):
        converter.convert(child, group.shapes, m, converter.style(root, {}), depth=1)
    if not len(group.shapes):
        group._element.getparent().remove(group._element)
        raise ValueError("SVG did not contain any supported shapes")
    group._element.recalculate_extents()
    return group
