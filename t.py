from zipfile import ZipFile
from lxml import etree

docx_path = "documents\Local File - SOFARSOLAR Netherlands BV - FYE 31 December 2024 - Final Draft - 10-09-2025 - V2.docx"

with ZipFile(docx_path) as docx:
    # Footnotes are stored in this XML file
    with docx.open("word/footnotes.xml") as f:
        xml = etree.parse(f)

# WordprocessingML uses the w: namespace
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

# Extract all footnote texts
footnotes = []
for fn in xml.findall(".//w:footnote", namespaces=ns):
    texts = [t.text for t in fn.findall(".//w:t", namespaces=ns) if t.text]
    if texts:
        footnotes.append("".join(texts))

print(footnotes)
