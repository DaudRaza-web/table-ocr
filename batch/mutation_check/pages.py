"""Render register pages upright, cropped to the table, as small JPEGs for checking by eye.

    python pages.py <pdf> <out dir> [pages, e.g. 1,3-5]   (default: all pages)

Each page is turned the way the register grid says it is upright (falling back to the detected turn)
and cropped to the table plus its footer, ~1600 px wide, JPEG quality 70 - legible digits at a small size.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(HERE)))
_argv, sys.argv = sys.argv, sys.argv[:1]
import pymupdf  # noqa: E402
import mutation  # noqa: E402
import server  # noqa: E402
sys.argv = _argv


def page_list(spec, count):
    if not spec:
        return list(range(count))
    out = []
    for part in spec.split(","):
        a, _, b = part.partition("-")
        out += list(range(int(a) - 1, int(b or a)))
    return [p for p in out if 0 <= p < count]


def main():
    pdf, out = sys.argv[1], sys.argv[2]
    os.makedirs(out, exist_ok=True)
    doc = pymupdf.open(pdf)
    for n in page_list(sys.argv[3] if len(sys.argv) > 3 else "", doc.page_count):
        rot = server.detect_orientation(doc, n)[0]
        found = mutation.locate_grid(doc[n], rot) or mutation.locate_grid(doc[n], (rot + 180) % 360)
        if found:
            page, r, g, _ = found
            x0, x1 = g["cols"][-1][0] - 0.01, g["cols"][0][1] + 0.01
            clip = (max(0, x0), max(0, g["top"] - 0.02), min(1, x1), 0.99)
        else:
            page, r, clip = doc[n], rot, (0, 0, 1, 1)
        z = 1600 / max(page.rect.width, page.rect.height) / max(0.3, clip[2] - clip[0])
        pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(r), alpha=False)
        w, h = pix.width, pix.height
        tmp = pymupdf.open()
        p = tmp.new_page(width=w, height=h)
        p.insert_image(p.rect, pixmap=pix)
        crop = pymupdf.Rect(clip[0] * w, clip[1] * h, clip[2] * w, clip[3] * h)
        img = p.get_pixmap(clip=crop, dpi=72)
        path = os.path.join(out, f"p{n + 1:03d}.jpg")
        img.save(path, jpg_quality=70)
        print(path, img.width, img.height, "grid" if found else "NO GRID - whole page", flush=True)


if __name__ == "__main__":
    main()
