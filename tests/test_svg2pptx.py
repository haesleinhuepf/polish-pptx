import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

from polish_pptx.svg2pptx import (add_svg_to_slide, apply, arc_to_cubics, parse_color, parse_path,
                                  parse_transform)

EMU_PER_UNIT = 100000  # slide of 10,000,000 x 5,000,000 EMU and viewBox of 100 x 50


def convert(body, viewbox="0 0 100 50"):
    prs = Presentation()
    prs.slide_width, prs.slide_height = 100 * EMU_PER_UNIT, 50 * EMU_PER_UNIT
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    svg = f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{viewbox}">{body}</svg>'
    return add_svg_to_slide(prs, slide, svg)


def only(group):
    assert len(group.shapes) == 1
    return group.shapes[0]


def local_tags(el):
    return [c.tag.split("}")[1] for c in el]


def test_rect():
    shape = only(convert('<rect x="10" y="5" width="30" height="20" fill="#ff0000" stroke="blue" stroke-width="2"/>'))
    assert shape.auto_shape_type == MSO_SHAPE.RECTANGLE
    assert (shape.left, shape.top, shape.width, shape.height) == (1000000, 500000, 3000000, 2000000)
    assert str(shape.fill.fore_color.rgb) == "FF0000"
    assert str(shape.line.color.rgb) == "0000FF"
    assert shape.line.width == 2 * EMU_PER_UNIT
    assert shape._element.find(qn("p:style")) is None  # no theme effects like shadows


def test_rounded_rect_and_opacity():
    shape = only(convert('<rect x="10" y="5" width="40" height="20" rx="5" fill="green" fill-opacity="0.5"/>'))
    assert shape.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE
    assert shape.adjustments[0] == pytest.approx(0.25)
    alpha = shape._element.spPr.find(qn("a:solidFill"))[0].find(qn("a:alpha"))
    assert alpha.get("val") == "50000"
    assert shape._element.spPr.find(qn("a:ln")).find(qn("a:noFill")) is not None  # stroke none


def test_background_rect_is_dropped():
    shape = only(convert('<rect width="100" height="50" fill="white"/><circle cx="50" cy="25" r="10"/>'))
    assert shape.auto_shape_type == MSO_SHAPE.OVAL
    assert (shape.left, shape.top, shape.width, shape.height) == (4000000, 1500000, 2000000, 2000000)
    assert str(shape.fill.fore_color.rgb) == "000000"  # SVG default fill


def test_viewbox_is_fitted_and_centered():
    shape = only(convert('<rect x="0" y="0" width="10" height="10"/>', viewbox="0 0 50 50"))
    assert (shape.left, shape.top, shape.width, shape.height) == (2500000, 0, 1000000, 1000000)


def test_css_classes_and_inheritance():
    group = convert('<style>.box { fill: #123456; stroke: none }</style>'
                    '<g fill="red"><rect class="box" x="0" y="0" width="10" height="10"/>'
                    '<rect x="20" y="0" width="10" height="10"/></g>')
    inner = only(group)
    assert inner.shape_type == MSO_SHAPE_TYPE.GROUP
    assert [str(s.fill.fore_color.rgb) for s in inner.shapes] == ["123456", "FF0000"]


def test_transform_rotation():
    shape = only(convert('<rect x="-10" y="-5" width="20" height="10" transform="translate(50 25) rotate(90)"/>'))
    assert shape.rotation == pytest.approx(90)
    assert (shape.left + shape.width / 2, shape.top + shape.height / 2) == (5000000, 2500000)


def test_arrow_with_marker():
    line = only(convert('<defs><marker id="arr" markerWidth="4" markerHeight="4"><path d="M0,0 L4,2 L0,4 z"/></marker></defs>'
                        '<line x1="10" y1="40" x2="90" y2="10" stroke="#333" marker-end="url(#arr)"/>'))
    assert line.shape_type == MSO_SHAPE_TYPE.LINE
    assert (line.begin_x, line.begin_y, line.end_x, line.end_y) == (1000000, 4000000, 9000000, 1000000)
    ln = line._element.spPr.find(qn("a:ln"))
    assert ln.find(qn("a:tailEnd")).get("type") == "triangle"
    assert ln.find(qn("a:headEnd")) is None


def test_curved_path_is_freeform():
    shape = only(convert('<path d="M10 10 C 20 0, 40 0, 50 10 Q 60 30 90 40" fill="none" stroke="black"/>'))
    assert shape.shape_type == MSO_SHAPE_TYPE.FREEFORM
    path = shape._element.spPr.find(qn("a:custGeom")).find(qn("a:pathLst"))[0]
    assert local_tags(path) == ["moveTo", "cubicBezTo", "cubicBezTo"]
    assert (int(path.get("w")), int(path.get("h"))) == (shape.width, shape.height)
    assert (shape.left, shape.top) == (1000000, 0)


def test_text():
    box = only(convert('<text x="50" y="25" font-size="5" text-anchor="middle" font-family="Arial, sans-serif" '
                       'font-weight="bold" fill="#00ff00">Hello  <tspan font-style="italic">world</tspan></text>'))
    assert box.text_frame.text == "Hello world"
    p = box.text_frame.paragraphs[0]
    assert p.alignment == PP_ALIGN.CENTER
    assert [r.font.italic for r in p.runs] == [False, True]
    assert all(r.font.bold and r.font.name == "Arial" and str(r.font.color.rgb) == "00FF00" for r in p.runs)
    assert p.runs[0].font.size == pytest.approx(5 * EMU_PER_UNIT, abs=127)
    assert box.left + box.width / 2 == pytest.approx(5000000, abs=1)


def test_multiline_text():
    group = convert('<text x="10" y="10" font-size="4"><tspan>one</tspan><tspan x="10" dy="5">two</tspan></text>')
    assert [s.text_frame.text for s in group.shapes] == ["one", "two"]
    assert group.shapes[1].top - group.shapes[0].top == 5 * EMU_PER_UNIT


def test_gradient_fill():
    shape = only(convert('<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="red"/>'
                         '<stop offset="100%" stop-color="blue"/></linearGradient></defs>'
                         '<rect x="10" y="10" width="20" height="20" fill="url(#g)"/>'))
    grad = shape._element.spPr.find(qn("a:gradFill"))
    assert [gs.get("pos") for gs in grad.find(qn("a:gsLst"))] == ["0", "100000"]
    assert grad.find(qn("a:lin")).get("ang") == "5400000"


def test_use_and_symbol():
    group = convert('<defs><symbol id="icon" viewBox="0 0 10 10"><circle cx="5" cy="5" r="5"/></symbol></defs>'
                    '<symbol id="unused"><rect width="5" height="5"/></symbol>'
                    '<use href="#icon" x="20" y="10" width="20" height="20"/>')
    shape = only(group)
    assert shape.auto_shape_type == MSO_SHAPE.OVAL
    assert (shape.left, shape.top, shape.width, shape.height) == (2000000, 1000000, 2000000, 2000000)


def test_malformed_element_is_skipped():
    shape = only(convert('<path d="M0 0 L 10 X"/><rect x="0" y="0" width="10" height="10"/>'))
    assert shape.auto_shape_type == MSO_SHAPE.RECTANGLE


def test_unsupported_svg_raises():
    with pytest.raises(ValueError):
        convert('<image href="x.png" width="10" height="10"/>')


def test_parse_path_relative_and_arcs():
    ops = parse_path("m10 10 h10 v10 l-10 0 z")
    assert ops == [("M", (10, 10)), ("L", (20, 10)), ("L", (20, 20)), ("L", (10, 20)), ("Z",)]
    ops = parse_path("M0 0 a10 10 0 016 6")  # compact arc flags
    assert ops[-1][-1] == pytest.approx((6, 6))
    quarter = arc_to_cubics((10, 0), 10, 10, 0, False, True, (0, 10))
    assert len(quarter) == 1 and quarter[0][2] == pytest.approx((0, 10))


def test_parse_transform_and_color():
    m = parse_transform("translate(10,20) scale(2)")
    assert apply(m, 1, 1) == (12, 22)
    assert parse_color("#abc") == ((0xAA, 0xBB, 0xCC), 1.0)
    assert parse_color("rgba(255, 0, 0, 0.5)") == ((255, 0, 0), 0.5)
    assert parse_color("steelblue") == ((70, 130, 180), 1.0)
    assert parse_color("none") is None
