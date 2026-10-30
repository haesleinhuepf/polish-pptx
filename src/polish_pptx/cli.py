import argparse
import sys

from .core import DEFAULT_MODEL, polish_presentation


def parse_slides(value):
    try:
        return [int(v) for v in value.split(",") if v.strip()]
    except ValueError:
        raise argparse.ArgumentTypeError(f"expected comma-separated slide numbers, got {value!r}")


def main(argv=None):
    p = argparse.ArgumentParser(prog="polish-pptx", description="Polish diagrams in a pptx file.")
    p.add_argument("filename", help="path of the pptx file")
    p.add_argument("--model", default=DEFAULT_MODEL, help="Anthropic model name")
    p.add_argument("--min-shapes", type=int, default=3,
                   help="minimum number of arrows/boxes/circles for a slide to be polished")
    p.add_argument("--slides", type=parse_slides, default=None,
                   help="comma-separated slide numbers to polish, e.g. 1,3,5 (overrides --min-shapes)")
    args = p.parse_args(argv)

    import anthropic
    from pptx import Presentation

    prs = Presentation(args.filename)
    client = anthropic.Anthropic()
    n = polish_presentation(prs, client, model=args.model, min_shapes=args.min_shapes,
                            slides=args.slides, log=lambda m: print(m, file=sys.stderr))
    out = args.filename + ".polished.pptx"
    prs.save(out)
    print(f"Polished {n} slide(s); saved {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
