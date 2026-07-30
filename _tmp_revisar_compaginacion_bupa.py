import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

from docx import Document
from docx.oxml.ns import qn

p = Path(
    r"reports/Bupa_Antofagasta/ABREGADO/AGREGADO_20260728_1950/"
    r"Reporte_Agregado_BUPA_20260723_20260727.docx"
)
d = Document(str(p))
print("KB", round(p.stat().st_size / 1024))
print("sections", len(d.sections))
for si, sec in enumerate(d.sections):
    print(
        f"  sec{si}: {sec.page_width.inches:.2f}x{sec.page_height.inches:.2f} in "
        f"margins T={sec.top_margin.inches:.2f} B={sec.bottom_margin.inches:.2f} "
        f"L={sec.left_margin.inches:.2f} R={sec.right_margin.inches:.2f}"
    )

print("--- body ---")
for i, child in enumerate(d.element.body):
    tag = child.tag.split("}")[-1]
    if tag == "p":
        t = "".join(x.text or "" for x in child.iter(qn("w:t"))).strip()[:95]
        blips = list(child.iter(qn("a:blip")))
        pPr = child.find(qn("w:pPr"))
        flags = []
        if pPr is not None:
            if pPr.find(qn("w:pageBreakBefore")) is not None:
                flags.append("PAGEBREAK")
            if pPr.find(qn("w:keepNext")) is not None:
                flags.append("keepNext")
            if pPr.find(qn("w:keepLines")) is not None:
                flags.append("keepLines")
        for br in child.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                flags.append("BR_PAGE")
        flag = (" [" + ",".join(flags) + "]") if flags else ""
        img = " [IMG]" if blips else ""
        if t or blips or flags:
            print(f"{i:03d}{flag}{img}: {t}")
    elif tag == "tbl":
        n = len(child.findall(qn("w:tr")))
        print(f"{i:03d} TABLE rows={n}")

# Check image rIds valid
print("--- images ---")
for child in d.element.body:
    for blip in child.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        print(f"  {rid}: {'OK' if rid in d.part.rels else 'BROKEN'}")
