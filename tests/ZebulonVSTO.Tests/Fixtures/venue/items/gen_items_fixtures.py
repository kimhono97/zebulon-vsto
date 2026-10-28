"""Regenerate the payload-planning fixtures from zebulon-exporter's /venue/export.

Exports payload.json on a few VENUE_ templates and records each generated slide: its role,
its speaker notes and the text of every praise/word box the template has. VenuePlannerFixtureTests
checks that VenuePlanner plans the same slides.

    python gen_items_fixtures.py <zebulon-exporter dir> <zebulon-content templates dir>
"""
import json
import os
import subprocess
import sys

exporterDir, templatesDir = sys.argv[1], sys.argv[2]
sys.path.insert(0, exporterDir)
import venue  # noqa: E402  (first: it imports collections.abc, which old python-pptx needs)
from pptx import Presentation  # noqa: E402

# 3 languages; a reading line ($lang1.ruby); first language only; side-by-side $lang2/$lang3.
TEMPLATES = ["VENUE_CornerStone.pptx", "VENUE_CornerStone_JP.pptx", "VENUE_UrbanGrow.pptx", "VENUE_Cheonbosan.pptx"]

here = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(here, "payload.json"), encoding="utf-8") as f:
    payload = json.load(f)
revision = subprocess.check_output(["git", "-C", exporterDir, "log", "-1", "--format=%h"], text=True).strip()

for name in TEMPLATES:
    path = os.path.join(templatesDir, name)
    template = venue.loadTemplate(path)
    out = venue.exportDeck(path, payload["items"], payload.get("options"))
    prs = Presentation(out)
    specs = {s.layout.part.partname: s for s in template.layouts}
    slides = []
    for slide in prs.slides:
        spec = specs.get(slide.slide_layout.part.partname)
        role = spec.role if spec is not None else None
        texts = {}
        if role in ("praise", "word"):
            placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
            for token, box in sorted(spec.boxes.items()):
                texts[str(token)] = placeholders[box.idx].text_frame.text
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        slides.append({"role": role, "notes": notes, "texts": texts})
    fixture = {"template": name, "specVersion": venue.SPEC_VERSION, "exporterRevision": revision, "slides": slides}
    target = os.path.join(here, os.path.splitext(name)[0] + ".slides.json")
    with open(target, "w", encoding="utf-8", newline="\n") as f:
        json.dump(fixture, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(target, len(slides))
