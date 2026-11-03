"""Regenerate the VENUE fixtures from zebulon-exporter's /venue/inspect.

Each fixture holds (1) the layout placeholders the way PowerPoint COM reports them
(input for VenueRules) and (2) what /venue/inspect says about each layout (expected).
VenueFixtureTests checks that VenueRules gives the same answers.

    python gen_fixtures.py <zebulon-exporter dir> <zebulon-content templates dir>

Needs the exporter's requirements (python-pptx, lxml, Pillow). Footer/date/slide-number
prompts are left out: their text is venue data (e.g. an offering account).
"""
import glob
import json
import os
import subprocess
import sys

exporterDir, templatesDir = sys.argv[1], sys.argv[2]
sys.path.insert(0, exporterDir)
import venue  # noqa: E402
from venue.layout import LATENT_TYPES  # noqa: E402

EMU_PER_POINT = 12700


def revision(path):
    try:
        return subprocess.check_output(["git", "-C", path, "log", "-1", "--format=%h"], text=True).strip()
    except Exception:
        return None


here = os.path.dirname(os.path.abspath(__file__))
for path in sorted(glob.glob(os.path.join(templatesDir, "VENUE_*.pptx"))):
    template = venue.loadTemplate(path)
    report = template.toDict()
    layouts = []
    for spec, expected in zip(template.layouts, report["layouts"]):
        placeholders = []
        for ph in spec.layout.placeholders:
            latent = ph.placeholder_format.type in LATENT_TYPES
            text = ph.text_frame.text if ph.has_text_frame and not latent else ""
            placeholders.append({
                # COM separates paragraphs with \r (python-pptx: \n); line breaks stay \v.
                "text": text.replace("\n", "\r"),
                "top": round(ph.top / EMU_PER_POINT, 2) if ph.top is not None else 0,
                "height": round(ph.height / EMU_PER_POINT, 2) if ph.height is not None else 0,
                "left": round(ph.left / EMU_PER_POINT, 2) if ph.left is not None else 0,
                "width": round(ph.width / EMU_PER_POINT, 2) if ph.width is not None else 0,
                "latent": latent,
                "hasTextFrame": bool(ph.has_text_frame),
            })
        layouts.append({
            "name": expected["name"],
            "master": expected["master"],
            "index": expected["index"],
            "placeholders": placeholders,
            "expected": {
                "role": expected["role"],
                "kind": expected["kind"],
                "valid": expected["valid"],
                "fillable": expected["fillable"],
                "tokens": sorted(int(t) for t in expected["tokens"]),
                "fields": [f["token"] for f in expected["fields"]],
                "issueCount": len(expected["issues"]),
            },
        })
    fixture = {
        "template": os.path.basename(path),
        "specVersion": venue.SPEC_VERSION,
        "exporterRevision": revision(exporterDir),
        "templatesRevision": revision(templatesDir),
        "roles": report["roles"],
        "layouts": layouts,
    }
    out = os.path.join(here, os.path.splitext(os.path.basename(path))[0] + ".json")
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(fixture, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(out)
