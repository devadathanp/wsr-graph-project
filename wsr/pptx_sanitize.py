"""Remove template metadata that triggers PowerPoint's repair dialog."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from lxml import etree

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


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


def sanitize_pptx(path: str | Path) -> Path:
    """
    Strip SharePoint customXml and co-authoring revisionInfo copied from the template.

    These parts are a common cause of “PowerPoint found a problem with content” when
    slides are rebuilt programmatically.
    """
    path = Path(path)
    buffer = io.BytesIO()

    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zout:
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

                zout.writestr(info, data)

    path.write_bytes(buffer.getvalue())
    return path
