import io
import re

from pptx.enum.shapes import MSO_SHAPE_TYPE

from .svg2pptx import add_svg_to_slide

DEFAULT_MODEL = "claude-sonnet-5-5"
MAX_EXAMPLES = 3  # most recent polished slides sent as few-shot examples

PROMPT = (
    "This is a screenshot of a presentation slide. Polish the figure on it: focus on the "
    "arrows, boxes, circles and similar diagram shapes (and their labels) and redraw them "
    "cleanly and consistently aligned. Ignore the headline/title text and layout logos. "
    "Make this figure look nice! Modify the colors, line styles and layout to fit to the slide-layout. "
    "Add icons and make it easily digestible. "
    "Return only the polished figure as a single SVG document (no explanation, no markdown), "
    "with a viewBox of the same aspect ratio as the screenshot and a transparent background. "
    "The SVG will be converted into editable PowerPoint shapes, so only use g, rect, circle, "
    "ellipse, line, polyline, polygon, path and text/tspan elements (styled with presentation "
    "attributes, solid colors or simple linear gradients) and markers for arrow heads. Draw icons "
    "with paths and simple shapes. Do not use images, filters, shadows, masks, clip paths, "
    "patterns, foreignObject or embedded fonts. Write each line of a label as text or tspan."
)

CONSISTENCY_PROMPT = (
    " The earlier slides of this presentation were already polished in this conversation. "
    "Keep the result visually consistent with them: reuse the same color palette, fonts, line "
    "styles, box shapes and arrow heads, and if this slide shows the same or similar concepts "
    "as an earlier slide, use the very same icons and visual representation for them."
)


def is_diagram_shape(shape):
    """True for autoshapes, connectors and freeforms (not text boxes, pictures, placeholders)."""
    if shape.is_placeholder:
        return False
    return shape.shape_type in (
        MSO_SHAPE_TYPE.AUTO_SHAPE,
        MSO_SHAPE_TYPE.FREEFORM,
        MSO_SHAPE_TYPE.LINE,
    ) or shape.element.tag.endswith("}cxnSp")


def count_diagram_shapes(slide):
    return sum(1 for s in slide.shapes if is_diagram_shape(s))


def image_to_base64_png(image):
    import base64

    buf = io.BytesIO()
    image.convert("RGB").save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode()


def extract_svg(text):
    m = re.search(r"<svg\b.*?</svg>", text, re.DOTALL | re.IGNORECASE)
    if not m:
        raise ValueError("Model response did not contain an SVG")
    return m.group(0)


def slide_message(image, text):
    return {
        "role": "user",
        "content": [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": image_to_base64_png(image),
                },
            },
            {"type": "text", "text": text},
        ],
    }


def request_svg(client, image, model=DEFAULT_MODEL, history=None):
    """Request a polished SVG for ``image``. ``history`` is a list of earlier
    (user message, svg) pairs; the last ``MAX_EXAMPLES`` are sent as few-shot examples
    so that similar slides get the same style and icons. On success the new pair is
    appended to it."""
    if history is None:
        history = []
    messages = []
    for user, svg in history[-MAX_EXAMPLES:]:
        messages.append(user)
        messages.append({"role": "assistant", "content": [{"type": "text", "text": svg}]})
    user = slide_message(image, PROMPT + (CONSISTENCY_PROMPT if history else ""))
    messages.append(user)
    response = client.messages.create(model=model, max_tokens=16000, messages=messages)
    text = "".join(b.text for b in response.content if getattr(b, "type", "") == "text")
    svg = extract_svg(text)
    history.append((user, svg))
    return svg


def is_title(shape):
    return shape.is_placeholder and "title" in str(shape.placeholder_format.type).lower()


def main_text_shape(slide):
    """The non-title placeholder or text box holding the most text, or None."""
    candidates = [
        s for s in slide.shapes
        if s.has_text_frame and s.text_frame.text.strip() and not is_title(s)
        and (s.is_placeholder or s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX)
    ]
    return max(candidates, key=lambda s: len(s.text_frame.text), default=None)


def move_shapes_off_slide(slide, slide_height):
    """Move all shapes except the title and the main text shape below the visible slide
    area, so that they are kept but not shown."""
    keep = [s._element for s in (slide.shapes.title, main_text_shape(slide)) if s is not None]
    for shape in list(slide.shapes):
        if not any(shape._element is e for e in keep):
            # set all four, as placeholders may only inherit their position from the layout
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            shape.left, shape.top, shape.width, shape.height = left, top + slide_height, width, height


def polish_presentation(prs, client, images=None, model=DEFAULT_MODEL, min_shapes=3, slides=None,
                        log=print):
    """Polish all slides having at least ``min_shapes`` diagram shapes, or, if ``slides``
    (1-based slide numbers) is given, exactly those slides. Returns polished count."""
    if slides is not None:
        targets = sorted({n - 1 for n in slides if 1 <= n <= len(prs.slides)})
    else:
        targets = [
            i for i, s in enumerate(prs.slides) if count_diagram_shapes(s) >= min_shapes
        ]
    if not targets:
        return 0
    if images is None:
        from render_pptx import render_slides

        images = render_slides(prs)
    done = 0
    history = []  # earlier polished slides, sent as few-shot examples for consistency
    for i in targets:
        slide = prs.slides[i]
        log(f"Polishing slide {i + 1}")
        try:
            svg = request_svg(client, images[i], model, history)
            move_shapes_off_slide(slide, prs.slide_height)
            add_svg_to_slide(prs, slide, svg)
            done += 1
        except Exception as e:  # keep going with other slides
            log(f"  slide {i + 1} skipped: {e}")
    return done
