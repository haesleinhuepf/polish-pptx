# polish-pptx

CLI tool that polishes pptx slides containing multiple arrows, boxes, circles etc.
Each such slide is rendered (`render-pptx`), sent to Claude, which returns a polished
SVG of the figure; the SVG is rasterized to PNG and added to the slide.

```
pip install .
export ANTHROPIC_API_KEY=...
polish-pptx deck.pptx      # writes deck.pptx.polished.pptx
```

Options: `--model` (default `claude-sonnet-5`), `--min-shapes` (default 3).
