import zipfile
from pathlib import Path

xlsm = Path(r"C:\Users\marcelo.espinosa\OneDrive - PichinchaCorp\Documentos\Claude Code\sri-xml-manager\excel\XML_Manager.xlsm")

with zipfile.ZipFile(xlsm, 'r') as z:
    names = z.namelist()
    print("=== Contenido del ZIP ===")
    for n in sorted(names):
        info = z.getinfo(n)
        print(f"  {n:45s} {info.file_size:8d} bytes")

    print()

    # Verificar vbaProject.bin
    if "xl/vbaProject.bin" in names:
        vba_data = z.read("xl/vbaProject.bin")
        magic = vba_data[:8].hex()
        print(f"vbaProject.bin: {len(vba_data)} bytes, magic={magic}")
        print(f"  -> CFB magic OK: {magic == 'd0cf11e0a1b11ae1'}")
    else:
        print("ERROR: xl/vbaProject.bin NO encontrado!")

    print()

    # Verificar [Content_Types].xml tiene entry para vbaProject
    ct = z.read("[Content_Types].xml").decode("utf-8")
    has_vba_ct = "vbaProject" in ct
    has_cui_ct = "customUI" in ct
    print(f"[Content_Types].xml contiene vbaProject: {has_vba_ct}")
    print(f"[Content_Types].xml contiene customUI:   {has_cui_ct}")

    # Verificar xl/_rels/workbook.xml.rels tiene relacion vbaProject
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    has_vba_rel = "vbaProject" in rels
    print(f"workbook.xml.rels contiene vbaProject:   {has_vba_rel}")

    # Verificar customUI/customUI.xml existe
    has_cui = "customUI/customUI.xml" in names
    print(f"customUI/customUI.xml presente:          {has_cui}")
