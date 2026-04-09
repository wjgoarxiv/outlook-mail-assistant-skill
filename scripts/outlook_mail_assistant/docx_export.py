from __future__ import annotations

from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


_CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def _paragraph_xml(text: str, *, bold: bool = False, size_half_points: int | None = None) -> str:
    escaped = escape(text)
    run_properties = []
    if bold:
        run_properties.append("<w:b/>")
    if size_half_points is not None:
        run_properties.append(f'<w:sz w:val="{size_half_points}"/>')
        run_properties.append(f'<w:szCs w:val="{size_half_points}"/>')
    run_properties_xml = f"<w:rPr>{''.join(run_properties)}</w:rPr>" if run_properties else ""
    return f'<w:p><w:r>{run_properties_xml}<w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>'


def _markdown_to_paragraphs(markdown_text: str) -> list[str]:
    paragraphs: list[str] = []
    for raw_line in markdown_text.splitlines():
        stripped = raw_line.strip()
        if not stripped:
            paragraphs.append("<w:p/>")
            continue
        if stripped.startswith("# "):
            paragraphs.append(_paragraph_xml(stripped[2:].strip(), bold=True, size_half_points=32))
            continue
        if stripped.startswith("## "):
            paragraphs.append(_paragraph_xml(stripped[3:].strip(), bold=True, size_half_points=28))
            continue
        if stripped.startswith("### "):
            paragraphs.append(_paragraph_xml(stripped[4:].strip(), bold=True, size_half_points=24))
            continue
        if stripped.startswith(("- ", "* ")):
            paragraphs.append(_paragraph_xml(f"• {stripped[2:].strip()}"))
            continue
        paragraphs.append(_paragraph_xml(stripped))
    return paragraphs or ["<w:p/>"]


def convert_markdown_to_docx(markdown_path: Path, output_path: Path) -> Path:
    markdown_path = Path(markdown_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    paragraphs_xml = "".join(_markdown_to_paragraphs(markdown_path.read_text(encoding="utf-8")))
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {paragraphs_xml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _CONTENT_TYPES_XML)
        archive.writestr("_rels/.rels", _RELS_XML)
        archive.writestr("word/document.xml", document_xml)

    return output_path
