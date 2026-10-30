import io
from types import SimpleNamespace

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from polish_pptx.core import count_diagram_shapes, extract_svg, polish_presentation

SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"><rect width="90" height="40" fill="red"/></svg>'


def make_prs():
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    for i in range(3):
        s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(i), Inches(1), Inches(1), Inches(1))
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_textbox(0, 0, Inches(1), Inches(1))
    return prs


def test_count():
    prs = make_prs()
    assert [count_diagram_shapes(s) for s in prs.slides] == [3, 0]


def test_extract_svg():
    assert extract_svg("```svg\n" + SVG + "\n```") == SVG


def test_polish():
    prs = make_prs()

    class Msgs:
        def create(self, **kw):
            return SimpleNamespace(content=[SimpleNamespace(type="text", text=SVG)])

    client = SimpleNamespace(messages=Msgs())
    imgs = [Image.new("RGB", (100, 50))] * 2
    assert polish_presentation(prs, client, images=imgs, log=lambda m: None) == 1
    assert len(prs.slides[0].shapes) == 4
    assert len(prs.slides[1].shapes) == 1
