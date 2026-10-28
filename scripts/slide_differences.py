import io
import os
import re
from pathlib import Path
from urllib.parse import urlparse

import fitz
import imageio.v2 as imageio
import numpy as np
import requests
from PIL import Image, ImageChops
from pptx import Presentation
from pptx.util import Inches

url = os.environ["SLIDES_URL"].strip()
match = re.search(r"/presentation/d/([a-zA-Z0-9_-]+)", url)
if not match or urlparse(url).netloc not in {"docs.google.com", "slides.google.com"}:
    raise SystemExit("Enter a public Google Slides presentation URL.")

presentation_id = match.group(1)
pdf_url = f"https://docs.google.com/presentation/d/{presentation_id}/export/pdf"
response = requests.get(pdf_url, timeout=90, allow_redirects=True)
response.raise_for_status()
if not response.content.startswith(b"%PDF"):
    raise SystemExit("Could not export this presentation as PDF. Check that link sharing allows anyone with the link to view it.")

work = Path("work")
work.mkdir(exist_ok=True)
doc = fitz.open(stream=response.content, filetype="pdf")
if len(doc) < 2:
    raise SystemExit("The presentation must contain at least two slides.")

images = []
for i, page in enumerate(doc):
    # Save each slide as SVG and render it for pixel-wise differencing.
    (work / f"slide-{i + 1:03}.svg").write_text(page.get_svg_image(), encoding="utf-8")
    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    images.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# Keep each adjacent-slide difference for the PPTX and the cumulative video.
differences = []
for i in range(len(images) - 1):
    a, b = images[i], images[i + 1]
    if b.size != a.size:
        b = b.resize(a.size, Image.Resampling.LANCZOS)
    diff = ImageChops.difference(b, a)
    # Boost subtle pixel changes so they remain visible.
    diff = diff.point(lambda p: min(255, p * 3))
    path = work / f"difference-{i + 1:03}-to-{i + 2:03}.png"
    diff.save(path)
    differences.append(np.asarray(diff, dtype=np.uint8))

    slide = prs.slides.add_slide(blank)
    slide.shapes.add_picture(str(path), 0, 0, width=prs.slide_width, height=prs.slide_height)

out = Path("outputs")
out.mkdir(exist_ok=True)
pptx_path = out / "slide-differences.pptx"
prs.save(pptx_path)

# Animate a persistent cumulative canvas: each new difference is wiped in
# from left to right, while all earlier differences remain visible.
video_path = out / "slide-differences.mp4"
fps = 24
seconds_per_difference = 2
reveal_frames = fps * seconds_per_difference
hold_frames = fps * 2
height, width, _ = differences[0].shape
cumulative = np.zeros((height, width, 3), dtype=np.uint8)

with imageio.get_writer(
    str(video_path),
    fps=fps,
    codec="libx264",
    quality=8,
    macro_block_size=2,
) as writer:
    # Start with a short black frame so the build-up is clear.
    for _ in range(fps):
        writer.append_data(cumulative)

    for diff in differences:
        for frame_number in range(1, reveal_frames + 1):
            boundary = min(width, int(width * frame_number / reveal_frames))
            frame = cumulative.copy()
            frame[:, :boundary] = np.maximum(frame[:, :boundary], diff[:, :boundary])
            writer.append_data(frame)

        cumulative = np.maximum(cumulative, diff)
        for _ in range(fps // 2):
            writer.append_data(cumulative)

    # Hold on the fully accumulated final image.
    for _ in range(hold_frames):
        writer.append_data(cumulative)

print(f"Created {pptx_path} with {len(differences)} difference slides.")
print(f"Created {video_path} with cumulative visual build-up.")
