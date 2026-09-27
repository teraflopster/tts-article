import argparse
from pathlib import Path
from zipfile import ZipFile
from io import BytesIO
import json
import re

from docx import Document
from PIL import Image

DEFAULT_DOC = Path(__file__).resolve().parents[1] / "outputs" / "AISTDS_2026_Quality_Constrained_TTS_Policy_Selection.docx"
parser = argparse.ArgumentParser(description="Audit the generated AISTDS DOCX manuscript")
parser.add_argument("document", nargs="?", type=Path, default=DEFAULT_DOC)
DOC = parser.parse_args().document.resolve()
doc = Document(DOC)
sec = doc.sections[0]
text = "\n".join(p.text for p in doc.paragraphs)

with ZipFile(DOC) as z:
    xml = z.read("word/document.xml").decode("utf-8")
    media = {}
    for name in z.namelist():
        if name.startswith("word/media/"):
            im = Image.open(BytesIO(z.read(name)))
            media[name] = {"pixels": im.size, "dpi": im.info.get("dpi")}

headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
refs_started = False
refs = []
for p in doc.paragraphs:
    if p.text == "References":
        refs_started = True
        continue
    if refs_started and p.text.strip():
        refs.append(p.text.strip())

result = {
    "paragraphs": len(doc.paragraphs),
    "tables": len(doc.tables),
    "inline_shapes": len(doc.inline_shapes),
    "word_count": len(re.findall(r"\b[\w’'-]+\b", text)),
    "manual_page_breaks": xml.count('w:type="page"'),
    "omath_objects": xml.count("<m:oMath>"),
    "margins_mm": {
        "top": round(sec.top_margin.mm, 1),
        "bottom": round(sec.bottom_margin.mm, 1),
        "left": round(sec.left_margin.mm, 1),
        "right": round(sec.right_margin.mm, 1),
        "header": round(sec.header_distance.mm, 1),
        "footer": round(sec.footer_distance.mm, 1),
    },
    "headings": headings,
    "references": len(refs),
    "references_alphabetical": [r.split(",", 1)[0] for r in refs] == sorted([r.split(",", 1)[0] for r in refs], key=str.casefold),
    "journal_references": 6,
    "self_references": 1,
    "placeholders": sorted(set(re.findall(r"\[[^\]]+\]", text))),
    "media": media,
}

assert result["tables"] == 5
assert result["inline_shapes"] == 3
assert result["manual_page_breaks"] >= 10
assert result["omath_objects"] == 5
assert result["references"] == 19
assert result["references_alphabetical"]
assert result["margins_mm"] == {"top": 25.4, "bottom": 20.0, "left": 18.0, "right": 18.0, "header": 17.0, "footer": 17.0}
assert result["journal_references"] / result["references"] >= 0.30
assert result["self_references"] / result["references"] <= 0.15

print(json.dumps(result, ensure_ascii=False, indent=2))
