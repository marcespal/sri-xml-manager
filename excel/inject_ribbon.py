"""
inject_ribbon.py — Inyecta customUI.xml en un archivo .xlsm existente.

El formato .xlsm es un ZIP con estructura OpenXML. Para agregar el ribbon
personalizado se necesita:
  1. Agregar customUI/customUI.xml al ZIP
  2. Agregar la relacion en _rels/.rels
  3. Actualizar [Content_Types].xml

Uso:
  python inject_ribbon.py --xlsm XML_Manager.xlsm --cui customUI/customUI.xml
"""
import argparse
import shutil
import zipfile
import re
from pathlib import Path
from xml.etree import ElementTree as ET


CONTENT_TYPE_CUSTOMUI = (
    'application/vnd.ms-office.activeX+xml'
)
# Tipo correcto para customUI
CT_CUSTOMUI = "application/vnd.ms-office.ui+xml.customUI"

# Relacion del workbook con customUI
REL_TYPE_CUSTOMUI = (
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
)


def inject_ribbon(xlsm_path: Path, customui_path: Path):
    xlsm_path   = Path(xlsm_path)
    customui_path = Path(customui_path)

    if not xlsm_path.exists():
        raise FileNotFoundError(f"XLSM no encontrado: {xlsm_path}")
    if not customui_path.exists():
        raise FileNotFoundError(f"customUI.xml no encontrado: {customui_path}")

    # Backup por si falla
    backup = xlsm_path.with_suffix(".xlsm.bak")
    shutil.copy2(xlsm_path, backup)

    # Leer contenido del customUI
    customui_content = customui_path.read_bytes()

    # Leer el ZIP completo en memoria
    with zipfile.ZipFile(xlsm_path, "r") as zin:
        nombres = zin.namelist()
        archivos = {n: zin.read(n) for n in nombres}

    # ── Actualizar [Content_Types].xml ──────────────────────
    ct_xml = archivos.get("[Content_Types].xml", b"")
    ct_str = ct_xml.decode("utf-8")

    if "customUI" not in ct_str:
        # Insertar Override para customUI
        override_tag = (
            '<Override PartName="/customUI/customUI.xml" '
            f'ContentType="{CT_CUSTOMUI}"/>'
        )
        # Insertar antes del cierre de Types
        ct_str = ct_str.replace("</Types>", override_tag + "\n</Types>")
        archivos["[Content_Types].xml"] = ct_str.encode("utf-8")

    # ── Actualizar _rels/.rels ──────────────────────────────
    rels_xml = archivos.get("_rels/.rels", b"")
    rels_str = rels_xml.decode("utf-8")

    if "customUI" not in rels_str:
        # Agregar relacion al paquete
        rel_tag = (
            '<Relationship Id="rIdCustomUI" '
            f'Type="{REL_TYPE_CUSTOMUI}" '
            'Target="customUI/customUI.xml"/>'
        )
        rels_str = rels_str.replace("</Relationships>",
                                    rel_tag + "\n</Relationships>")
        archivos["_rels/.rels"] = rels_str.encode("utf-8")

    # ── Agregar/reemplazar customUI/customUI.xml ─────────────
    archivos["customUI/customUI.xml"] = customui_content

    # ── Reescribir el ZIP ───────────────────────────────────
    temp_path = xlsm_path.with_suffix(".tmp.xlsm")
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for nombre, contenido in archivos.items():
            zout.writestr(nombre, contenido)

    # Reemplazar original
    xlsm_path.unlink()
    temp_path.rename(xlsm_path)

    # Eliminar backup si todo OK
    backup.unlink(missing_ok=True)

    print(f"Ribbon inyectado correctamente en: {xlsm_path.name}")


def main():
    parser = argparse.ArgumentParser(description="Inyecta customUI ribbon en .xlsm")
    parser.add_argument("--xlsm", required=True, help="Ruta al archivo .xlsm")
    parser.add_argument("--cui",  required=True, help="Ruta al customUI.xml")
    args = parser.parse_args()

    try:
        inject_ribbon(Path(args.xlsm), Path(args.cui))
    except Exception as e:
        print(f"ERROR: {e}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
