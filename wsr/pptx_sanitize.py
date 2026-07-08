"""Repair package metadata that triggers PowerPoint's "found a problem" dialog.

Two fixes are applied after python-pptx saves:
1. Strip SharePoint ``customXml`` and co-authoring ``revisionInfo`` parts copied
   from the template.
2. Rewrite ``docProps/app.xml`` so the slide counts and titles match the deck we
   actually built. python-pptx copies the template's app.xml verbatim, so the
   stale ``<Slides>``/``<TitlesOfParts>`` values mismatch the real slides and
   make PowerPoint rebuild (blank) slides on repair.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from lxml import etree

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

_SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")


def _clean_presentation_rels(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for rel in list(root):
        target = rel.get("Target") or ""
        if "customXml" in target or "revisionInfo" in target:
            root.remove(rel)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _clean_content_types(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for node in list(root):
        part = node.get("PartName") or ""
        if part.startswith("/customXml/") or part == "/ppt/revisionInfo.xml":
            root.remove(node)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _extract_title(slide_bytes: bytes) -> str:
    """Return the title placeholder text of a slide, or a generic fallback."""
    root = etree.fromstring(slide_bytes)
    for sp in root.iter(f"{{{P_NS}}}sp"):
        ph = sp.find(f".//{{{P_NS}}}nvSpPr/{{{P_NS}}}nvPr/{{{P_NS}}}ph")
        if ph is not None and (ph.get("type") in ("title", "ctrTitle")):
            texts = sp.findall(f".//{{{A_NS}}}t")
            title = "".join(t.text or "" for t in texts).strip()
            return title or "PowerPoint Presentation"
    return "PowerPoint Presentation"


def _rebuild_app_xml(data: bytes, titles: list[str]) -> bytes:
    """Sync slide counts and TitlesOfParts in app.xml with the actual deck."""
    root = etree.fromstring(data)
    app = f"{{{APP_NS}}}"
    vt = f"{{{VT_NS}}}"
    n = len(titles)

    for tag, value in (("Slides", n), ("Notes", 0), ("HiddenSlides", 0)):
        el = root.find(f"{app}{tag}")
        if el is None:
            el = etree.SubElement(root, f"{app}{tag}")
        el.text = str(value)

    # Preserve the existing theme name (first TitlesOfParts entry) when present.
    theme_name = "Office Theme"
    titles_of_parts = root.find(f"{app}TitlesOfParts")
    if titles_of_parts is not None:
        existing = titles_of_parts.find(f"{vt}vector")
        if existing is not None and len(existing):
            theme_name = existing[0].text or theme_name

    heading_pairs = root.find(f"{app}HeadingPairs")
    if heading_pairs is None:
        heading_pairs = etree.SubElement(root, f"{app}HeadingPairs")
    for child in list(heading_pairs):
        heading_pairs.remove(child)
    vec = etree.SubElement(heading_pairs, f"{vt}vector")
    vec.set("size", "4")
    vec.set("baseType", "variant")
    for label, count in (("Theme", 1), ("Slide Titles", n)):
        variant = etree.SubElement(vec, f"{vt}variant")
        etree.SubElement(variant, f"{vt}lpstr").text = label
        variant = etree.SubElement(vec, f"{vt}variant")
        etree.SubElement(variant, f"{vt}i4").text = str(count)

    if titles_of_parts is None:
        titles_of_parts = etree.SubElement(root, f"{app}TitlesOfParts")
    for child in list(titles_of_parts):
        titles_of_parts.remove(child)
    vec = etree.SubElement(titles_of_parts, f"{vt}vector")
    vec.set("size", str(n + 1))
    vec.set("baseType", "lpstr")
    for value in [theme_name, *titles]:
        etree.SubElement(vec, f"{vt}lpstr").text = value

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _slide_titles(zin: zipfile.ZipFile) -> list[str]:
    slide_names = sorted(
        (name for name in zin.namelist() if _SLIDE_RE.match(name)),
        key=lambda name: int(_SLIDE_RE.match(name).group(1)),
    )
    return [_extract_title(zin.read(name)) for name in slide_names]


def sanitize_pptx(path: str | Path) -> Path:
    path = Path(path)
    source = path.read_bytes()

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        titles = _slide_titles(zin)

    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source)) as zin, zipfile.ZipFile(
        buffer, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            name = info.filename
            if name.startswith("customXml/"):
                continue
            if name == "ppt/revisionInfo.xml":
                continue

            data = zin.read(name)
            if name == "ppt/_rels/presentation.xml.rels":
                data = _clean_presentation_rels(data)
            elif name == "[Content_Types].xml":
                data = _clean_content_types(data)
            elif name == "docProps/app.xml":
                data = _rebuild_app_xml(data, titles)

            zout.writestr(info, data)

    path.write_bytes(buffer.getvalue())
    return path
