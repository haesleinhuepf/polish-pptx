import io
from types import SimpleNamespace

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from polish_pptx.core import (CONSISTENCY_PROMPT, add_png_to_slide, content_bbox,
                              count_diagram_shapes,
                              extract_svg, move_shapes_off_slide, polish_presentation)

SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"><rect width="90" height="40" fill="red"/></svg>'


def visible_shapes(prs, slide):
    return [s for s in slide.shapes if s.top < prs.slide_height]


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
    assert len(visible_shapes(prs, prs.slides[0])) == 1  # diagram shapes replaced by the picture
    assert len(prs.slides[0].shapes) == 4  # diagram shapes kept below the slide
    assert len(prs.slides[1].shapes) == 1


def test_polish_selected_slides():
    prs = make_prs()

    class Msgs:
        def create(self, **kw):
            return SimpleNamespace(content=[SimpleNamespace(type="text", text=SVG)])

    client = SimpleNamespace(messages=Msgs())
    imgs = [Image.new("RGB", (100, 50))] * 2
    assert polish_presentation(prs, client, images=imgs, slides=[2], log=lambda m: None) == 1
    assert len(prs.slides[0].shapes) == 3
    assert len(visible_shapes(prs, prs.slides[1])) == 1  # empty text box moved off the slide
    assert len(prs.slides[1].shapes) == 2


def test_polish_sends_earlier_slides_as_examples():
    prs = make_prs()
    calls = []

    class Msgs:
        def create(self, **kw):
            calls.append(kw["messages"])
            return SimpleNamespace(content=[SimpleNamespace(type="text", text=SVG)])

    client = SimpleNamespace(messages=Msgs())
    imgs = [Image.new("RGB", (100, 50))] * 2
    assert polish_presentation(prs, client, images=imgs, slides=[1, 2], log=lambda m: None) == 2
    assert [m["role"] for m in calls[0]] == ["user"]
    assert [m["role"] for m in calls[1]] == ["user", "assistant", "user"]
    assert calls[1][1]["content"][0]["text"] == SVG
    assert calls[1][2]["content"][-1]["text"].endswith(CONSISTENCY_PROMPT)
    assert CONSISTENCY_PROMPT not in calls[0][0]["content"][-1]["text"]


def test_request_svg_limits_examples():
    from polish_pptx.core import MAX_EXAMPLES, request_svg

    calls = []

    class Msgs:
        def create(self, **kw):
            calls.append(kw["messages"])
            return SimpleNamespace(content=[SimpleNamespace(type="text", text=SVG)])

    client = SimpleNamespace(messages=Msgs())
    history = []
    for _ in range(MAX_EXAMPLES + 2):
        request_svg(client, Image.new("RGB", (100, 50)), history=history)
    assert len(history) == MAX_EXAMPLES + 2
    assert len(calls[-1]) == 2 * MAX_EXAMPLES + 1


def test_move_shapes_off_slide_keeps_title_and_main_text():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # title and content
    slide.shapes.title.text = "Headline"
    slide.placeholders[1].text = "Main text with several words"
    slide.shapes.add_textbox(0, 0, Inches(1), Inches(1)).text = "label"
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(1), Inches(1)).text = "a much longer box label"
    tops = [s.top for s in slide.shapes]
    move_shapes_off_slide(slide, prs.slide_height)
    assert [s.text_frame.text for s in visible_shapes(prs, slide)] == ["Headline", "Main text with several words"]
    assert [s.top for s in slide.shapes] == tops[:2] + [t + prs.slide_height for t in tops[2:]]


def test_content_bbox():
    img = Image.new("RGBA", (100, 50), (0, 0, 0, 0))
    img.paste((255, 0, 0, 255), (10, 5, 30, 25))
    assert content_bbox(img) == (10, 5, 30, 25)
    img = Image.new("RGB", (100, 50), "white")
    img.paste((255, 0, 0), (10, 5, 30, 25))
    assert content_bbox(img) == (10, 5, 30, 25)


def test_add_png_keeps_content_position():
    prs = Presentation()
    prs.slide_width, prs.slide_height = 2000000, 2000000  # PNG of 100x50 is scaled by 20000, top offset 500000
    img = Image.new("RGBA", (100, 50), (0, 0, 0, 0))
    img.paste((255, 0, 0, 255), (10, 5, 30, 25))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    pic = add_png_to_slide(prs, prs.slides.add_slide(prs.slide_layouts[6]), buf.getvalue())
    assert (pic.left, pic.top, pic.width, pic.height) == (200000, 600000, 400000, 400000)
    assert Image.open(io.BytesIO(pic.image.blob)).size == (20, 20)
