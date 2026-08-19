import sys, struct
sys.path.insert(0, r"C:\Users\marcelo.espinosa\OneDrive - PichinchaCorp\Documentos\Claude Code\sri-xml-manager\excel")

from create_vba_project import generar_vba_project_bin

test_vba = 'Sub Hola()\n    MsgBox "Hola SRI!"\nEnd Sub\n'

try:
    data = generar_vba_project_bin({"modMain": test_vba})
    print(f"OK: {len(data)} bytes generados")

    # Verificar magic header CFB
    magic = data[:8]
    expected = bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])
    print(f"Magic CFB: {magic.hex()} -> {'OK' if magic == expected else 'ERROR'}")

    # Verificar que el total es multiplo de 512
    print(f"Tamano multiplo de 512: {len(data) % 512 == 0} (resto={len(data) % 512})")

    # Leer header CFB
    (minor_ver, major_ver, byte_order, sect_shift) = struct.unpack_from('<HHHH', data, 24)
    print(f"Version: {major_ver}.{minor_ver}, ByteOrder: {byte_order:#06x}, SectShift: {sect_shift}")

    # Verificar sector size = 2^sect_shift
    sect_size = 2 ** sect_shift
    print(f"Sector size: {sect_size} bytes -> {'OK' if sect_size == 512 else 'ERROR'}")

    # Leer primera entrada de directorio (Root Entry) en sector 1 (offset 512+512=1024)
    root_off = 512 + 512  # header + sector 0 (FAT)
    root_name_raw = data[root_off:root_off+64]
    root_name = root_name_raw.rstrip(b'\x00').decode('utf-16-le', errors='replace')
    root_type = data[root_off + 66]
    print(f"Root Entry: name='{root_name}', type={root_type} (5=root) -> {'OK' if root_type == 5 else 'ERROR'}")

    # Guardar para inspeccion manual
    out_path = r"C:\Users\marcelo.espinosa\OneDrive - PichinchaCorp\Documentos\Claude Code\sri-xml-manager\excel\test_output.bin"
    with open(out_path, 'wb') as f:
        f.write(data)
    print(f"Guardado en: {out_path}")

except Exception as e:
    import traceback
    print(f"ERROR: {e}")
    traceback.print_exc()
