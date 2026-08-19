import zipfile
from pathlib import Path

xlsm = Path(r"C:\Users\marcelo.espinosa\OneDrive - PichinchaCorp\Documentos\Claude Code\sri-xml-manager\excel\XML_Manager.xlsm")

with zipfile.ZipFile(xlsm, 'r') as z:
    # Check _rels/.rels content
    print("=== _rels/.rels ===")
    rels_root = z.read("_rels/.rels").decode("utf-8")
    print(rels_root)
    print()

    # Check [Content_Types].xml
    print("=== [Content_Types].xml ===")
    ct = z.read("[Content_Types].xml").decode("utf-8")
    print(ct)
    print()

    # Check xl/_rels/workbook.xml.rels
    print("=== xl/_rels/workbook.xml.rels ===")
    wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    print(wb_rels)
    print()

    # Check customUI/customUI.xml first 200 chars
    print("=== customUI/customUI.xml (first 200 chars) ===")
    cui = z.read("customUI/customUI.xml").decode("utf-8")
    print(cui[:200])
