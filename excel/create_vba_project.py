"""
create_vba_project.py — Genera vbaProject.bin desde archivos .bas

Implementa los formatos necesarios para crear un vbaProject.bin válido
que Excel pueda leer y recompilar al abrir el .xlsm:

  [MS-CFB]  Compound File Binary Format — contenedor OLE2
  [MS-OVBA] Office VBA File Format Structure — streams VBA

Los módulos se guardan CON código fuente pero SIN p-code compilado.
Excel recompila el p-code automáticamente al abrir (igual que al abrir
un .xlsm de una versión diferente de Office). Esto es normal e inofensivo.

Uso:
    from create_vba_project import generar_vba_project_bin
    data = generar_vba_project_bin({
        "modMain": Path("vba/modMain.bas").read_text(encoding="utf-8"),
    })
    # Inyectar data en el ZIP del .xlsm como "xl/vbaProject.bin"
"""

import struct
import math
from pathlib import Path
from typing import Dict, Optional

# ══════════════════════════════════════════════════════════════════════════════
# Constantes CFB (Compound File Binary)
# ══════════════════════════════════════════════════════════════════════════════

SECTOR_SIZE = 512          # bytes por sector (ShiftLeft=9 → 2^9=512)
DIR_ENTRY_SIZE = 128       # bytes por entrada de directorio
DIRS_PER_SECTOR = SECTOR_SIZE // DIR_ENTRY_SIZE  # = 4

FREESECT    = 0xFFFFFFFF
ENDOFCHAIN  = 0xFFFFFFFE
FATSECT     = 0xFFFFFFFD
DIFSECT     = 0xFFFFFFFC
NOSTREAM    = 0xFFFFFFFF

ENTRY_EMPTY   = 0   # tipo directorio vacío
ENTRY_STORAGE = 1   # tipo storage (carpeta)
ENTRY_STREAM  = 2   # tipo stream (archivo)
ENTRY_ROOT    = 5   # tipo root

# CLSID del root en vbaProject.bin (requerido por Excel)
# {000204EF-0000-0000-C000-000000000046}
ROOT_CLSID = bytes([
    0xEF, 0x04, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
])


# ══════════════════════════════════════════════════════════════════════════════
# MS-OVBA Compresión  (sección 2.4 de [MS-OVBA])
# ══════════════════════════════════════════════════════════════════════════════

def _bit_length(n: int) -> int:
    return n.bit_length() if n > 0 else 1


def _copy_token_help(decompressed_size: int):
    """
    Calcula LengthMask, OffsetMask, BitCount y MaximumLength
    para codificar un CopyToken según el tamaño actual del buffer.
    """
    if decompressed_size <= 1:
        bit_count = 4
    else:
        bit_count = max(_bit_length(decompressed_size - 1), 4)
    length_mask   = (1 << (16 - bit_count)) - 1
    offset_mask   = 0xFFFF - length_mask
    max_length    = length_mask + 3
    return length_mask, offset_mask, bit_count, max_length


def _compress_chunk(data: bytes) -> bytes:
    """
    Comprime un bloque de hasta 4096 bytes con el algoritmo LZ77 de MS-OVBA.
    Devuelve los bytes del cuerpo comprimido (SIN el header de 2 bytes).
    """
    out          = bytearray()
    decompressed = bytearray()
    i = 0

    while i < len(data):
        flags_byte = 0
        tokens = bytearray()

        for bit in range(8):
            if i >= len(data):
                break

            ds = len(decompressed)
            _, _, bit_count, max_length = _copy_token_help(ds)
            max_match = min(max_length, len(data) - i)

            # Buscar la mejor coincidencia en la ventana
            best_len = 0
            best_off = 0
            window_start = max(0, ds - 4096)

            for back in range(1, ds - window_start + 1):
                src = ds - back
                l = 0
                while l < max_match:
                    idx = src + (l % back)
                    if 0 <= idx < ds and decompressed[idx] == data[i + l]:
                        l += 1
                    else:
                        break
                if l >= 3 and l > best_len:
                    best_len = l
                    best_off = back

            if best_len >= 3:
                # CopyToken (back-reference)
                flags_byte |= (1 << bit)
                length_mask, _, bit_count2, _ = _copy_token_help(ds)
                lf = best_len - 3
                of = best_off - 1
                ct = (of << (16 - bit_count2)) | lf
                tokens.extend(struct.pack('<H', ct))
                decompressed.extend(data[i:i + best_len])
                i += best_len
            else:
                # Literal byte
                tokens.append(data[i])
                decompressed.append(data[i])
                i += 1

        out.append(flags_byte)
        out.extend(tokens)

    return bytes(out)


def ovba_compress(data: bytes) -> bytes:
    """
    Comprime data con el algoritmo MS-OVBA.
    Retorna los bytes en formato CompressedContainer.
    """
    result = bytearray([0x01])  # SignatureByte

    if len(data) == 0:
        # Chunk vacío: un chunk de cuerpo vacío no es válido;
        # en su lugar emitimos un chunk de un byte nulo.
        data = b'\x00'

    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        body  = _compress_chunk(chunk)
        size  = len(body)
        # Bits: [0:11]=size-3, bit12=IsCompressed(1), bits[13:15]=0b011(sig)
        if size >= 3:
            hdr = (size - 3) | (1 << 12) | (3 << 13)
        else:
            # Chunk muy pequeño: pad a mínimo de 3 bytes
            body = body + b'\x00' * (3 - size)
            hdr  = (0) | (1 << 12) | (3 << 13)
        result.extend(struct.pack('<H', hdr))
        result.extend(body)

    return bytes(result)


# ══════════════════════════════════════════════════════════════════════════════
# Streams VBA  (sección 2.3 de [MS-OVBA])
# ══════════════════════════════════════════════════════════════════════════════

def _u16le(name: str) -> bytes:
    """Codifica un string como UTF-16LE."""
    return name.encode("utf-16-le")


def _pack_record(rec_id: int, data: bytes) -> bytes:
    """Construye un registro de longitud variable para el stream 'dir'."""
    return struct.pack('<HI', rec_id, len(data)) + data


def build_dir_stream(modules: Dict[str, str]) -> bytes:
    """
    Construye el stream 'dir' descomprimido (luego se comprime con ovba_compress).
    Contiene los metadatos de todos los módulos del proyecto VBA.
    """
    rec = bytearray()

    # ── Información del proyecto ──────────────────────────────────────────────
    # PROJECTSYSKIND (0x0001): 0x00000001 = Win32
    rec += _pack_record(0x0001, struct.pack('<I', 0x00000001))
    # PROJECTLCID (0x0002): 0x0409 = English US
    rec += _pack_record(0x0002, struct.pack('<I', 0x0409))
    # PROJECTLCIDINVOKE (0x0014)
    rec += _pack_record(0x0014, struct.pack('<I', 0x0409))
    # PROJECTCODEPAGE (0x0003): 1252 = Windows Latin-1
    rec += _pack_record(0x0003, struct.pack('<H', 1252))
    # PROJECTNAME (0x0004)
    rec += _pack_record(0x0004, b'VBAProject')
    # PROJECTDOCSTRING (0x0005) — vacío
    rec += _pack_record(0x0005, b'')
    # PROJECTDOCSTRINGUNICODE (0x0040) — vacío
    rec += _pack_record(0x0040, b'')
    # PROJECTHELPFILEPATH (0x0006) — vacío
    rec += _pack_record(0x0006, b'')
    # PROJECTHELPFILEPATH2 (0x003D) — vacío
    rec += _pack_record(0x003D, b'')
    # PROJECTHELPCONTEXT (0x0007)
    rec += _pack_record(0x0007, struct.pack('<I', 0))
    # PROJECTLIBFLAGS (0x0008)
    rec += _pack_record(0x0008, struct.pack('<I', 0))
    # PROJECTVERSION (0x0009): major=0x81, minor=0x000D (valores típicos)
    rec += struct.pack('<HI', 0x0009, 4) + struct.pack('<I', 0x81)
    rec += _pack_record(0x000A, struct.pack('<H', 0x000D))
    # PROJECTCONSTANTS (0x000C) — vacío
    rec += _pack_record(0x000C, b'')
    # PROJECTCONSTANTSUNICODE (0x003C) — vacío
    rec += _pack_record(0x003C, b'')

    # ── Referencias (ninguna) ────────────────────────────────────────────────
    # (sin referencias externas, sección 2.3.4.2)

    # ── Módulos ───────────────────────────────────────────────────────────────
    # PROJECTMODULES (0x000F): cuenta de módulos
    mod_count = len(modules)
    rec += _pack_record(0x000F, struct.pack('<H', mod_count))
    # PROJECTCOOKIE (0x0013)
    rec += _pack_record(0x0013, struct.pack('<H', 0xFFFF))

    for mod_name, _source in modules.items():
        name_bytes = mod_name.encode("mbcs", errors="replace")
        name_uni   = _u16le(mod_name)

        # MODULENAME (0x0019)
        rec += _pack_record(0x0019, name_bytes)
        # MODULENAMEUNICODE (0x0031) — nota: spec usa 0x0031 para unicode
        rec += _pack_record(0x0031, name_uni)
        # MODULESTREAMNAME (0x001A): nombre del stream en el storage VBA
        rec += _pack_record(0x001A, name_bytes)
        # MODULESTREAMNAMEUNICODE (0x0032)
        rec += _pack_record(0x0032, name_uni)
        # MODULEDOCSTRING (0x001C) — vacío
        rec += _pack_record(0x001C, b'')
        # MODULEDOCSTRINGUNICODE (0x0048) — vacío
        rec += _pack_record(0x0048, b'')
        # MODULEOFFSET (0x0031 en spec vieja, pero es 0x002C para TextOffset)
        # TEXTOFFSET: offset en el stream del módulo donde empieza el source comprimido
        # Usamos 0 → todo el stream es source comprimido (sin p-code)
        rec += _pack_record(0x0031, struct.pack('<I', 0))
        # MODULEHELPCONTEXT (0x001E)
        rec += _pack_record(0x001E, struct.pack('<I', 0))
        # MODULECOOKIE (0x002C) — ID corto del módulo
        rec += _pack_record(0x002C, struct.pack('<H', 0xFFFF))
        # MODULETYPE (0x0021 = procedural standard module, 0x0022 = class)
        rec += _pack_record(0x0021, b'')
        # MODULEREADONLY (0x0025) — no incluir si el módulo es editable
        # (Omitido — módulo escribible)
        # MODULEPRIVATE (0x0028) — no incluir
        # (Omitido)
        # MODULETERM (0x002B): fin del módulo
        rec += _pack_record(0x002B, b'')

    # PROJECTMODULETERMINATOR (0x0010)
    rec += _pack_record(0x0010, b'')

    return bytes(rec)


def build_module_stream(source_code: str) -> bytes:
    """
    Construye el stream de un módulo VBA.
    TextOffset = 0 → todo el stream es el source comprimido.
    El p-code está ausente (Excel recompilará al abrir).
    """
    src_bytes  = source_code.encode("mbcs", errors="replace")
    compressed = ovba_compress(src_bytes)
    return compressed


def build_vba_project_stream() -> bytes:
    """
    Construye el stream '_VBA_PROJECT'.
    Contiene una cabecera mínima que indica que el p-code no es válido
    (PerformanceCache no disponible), forzando recompilación desde fuente.
    """
    # Bytes mágicos que identifican un proyecto VBA sin p-code cacheado.
    # Este valor es reconocido por Excel como "recompilar desde source".
    return bytes([
        0x61, 0x08, 0x1B, 0xE0,  # VBA project header magic
        0x00, 0x00, 0xFF, 0xFF,  # Reserved + performance cache disabled
        0x00, 0x00, 0x00, 0x00,
        0x06, 0x00,               # Version
        0x00, 0x00,
    ])


def build_project_stream(modules: Dict[str, str]) -> bytes:
    """
    Construye el stream 'PROJECT' (texto INI-style con la config del proyecto).
    """
    lines = [
        'ID="{00000000-0000-0000-0000-000000000000}"',
        'Document=ThisWorkbook/&H00000000',
        'Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}',
        'BaseClass=0',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="CAC866BC3E3A3E3A3E3A3E3A"',
        'DPB="8E8CBE3CC03CC03C"',
        'GC="242632642A432A43"',
        '',
        '[Host Extender Info]',
        'Class0={00020820-0000-0000-C000-000000000046}#0#0; Excel',
        '',
        '[Workspace]',
        f'ThisWorkbook=0, 0, 0, 0, C',
    ]
    for mod_name in modules:
        lines.append(f'{mod_name}=0, 0, 0, 0, C')
    return '\r\n'.join(lines).encode("mbcs", errors="replace")


def build_project_wm_stream(modules: Dict[str, str]) -> bytes:
    """
    Construye el stream 'PROJECTwm' que mapea nombres de módulos.
    Puede estar vacío para módulos estándar.
    """
    out = bytearray()
    for name in modules:
        out += name.encode("mbcs", errors="replace") + b'\x00'
        out += name.encode("utf-16-le") + b'\x00\x00'
    return bytes(out)


# ══════════════════════════════════════════════════════════════════════════════
# Escritor CFB mínimo  (sección 2 de [MS-CFB])
# ══════════════════════════════════════════════════════════════════════════════

class _DirEntry:
    """Una entrada del directorio CFB (128 bytes)."""

    def __init__(self, name: str = "",
                 entry_type: int = ENTRY_EMPTY,
                 color: int = 1,
                 left: int = NOSTREAM,
                 right: int = NOSTREAM,
                 child: int = NOSTREAM,
                 clsid: bytes = b'\x00' * 16,
                 state: int = 0,
                 created: int = 0,
                 modified: int = 0,
                 start: int = ENDOFCHAIN,
                 size: int = 0):
        self.name       = name
        self.entry_type = entry_type
        self.color      = color   # 0=red, 1=black
        self.left       = left
        self.right      = right
        self.child      = child
        self.clsid      = clsid
        self.state      = state
        self.created    = created
        self.modified   = modified
        self.start      = start
        self.size       = size

    def pack(self) -> bytes:
        name_utf16 = self.name.encode("utf-16-le")[:62]
        name_len   = len(name_utf16) + 2  # include null terminator
        name_field = name_utf16 + b'\x00' * (64 - len(name_utf16))
        return struct.pack(
            '<64sHBBIII16sIQQIII',
            name_field,
            name_len,
            self.entry_type,
            self.color,
            self.left,
            self.right,
            self.child,
            self.clsid,
            self.state,
            self.created,
            self.modified,
            self.start & 0xFFFFFFFF,
            self.size & 0xFFFFFFFF,
            0,                    # size high (always 0 for v3)
        )


def _pad_sector(data: bytes) -> bytes:
    """Rellena data hasta ser múltiplo de SECTOR_SIZE con bytes 0x00."""
    rem = len(data) % SECTOR_SIZE
    if rem:
        data = data + b'\x00' * (SECTOR_SIZE - rem)
    return data


class CFBWriter:
    """
    Escritor mínimo de archivos Compound File Binary (OLE2).

    Implementa solo lo necesario para crear un vbaProject.bin válido:
    - Un solo sector FAT
    - Directorio en sectores completos
    - Streams en sectores completos (sin mini-stream)
    - Sin double-indirect FAT (DIFAT)
    """

    def __init__(self):
        self._streams: list[tuple[str, bytes]] = []  # (path, data)

    def add_stream(self, path: str, data: bytes):
        """Agrega un stream. El path usa '/' como separador, ej: 'VBA/_VBA_PROJECT'."""
        self._streams.append((path, data))

    def write(self) -> bytes:
        """Genera el archivo CFB completo y lo devuelve como bytes."""
        # ── Paso 1: asignar sectores a cada stream ────────────────────────
        # Primero calculamos cuántos sectores necesita cada stream.

        stream_sectors = {}  # path → (start_sector, sector_count, data_padded)
        sector_idx = 0

        # Reservar: sector 0 = FAT, sector 1+ = directorio (calcular después)
        # Calcular tamaño del directorio: necesitamos un entry por storage+stream

        # Recopilar estructura de storages necesarios
        storages = set()
        for path, _ in self._streams:
            parts = path.split('/')
            for i in range(1, len(parts)):
                storages.add('/'.join(parts[:i]))

        # Total de entradas directorio: 1 Root + storages + streams
        all_entries = [''] + sorted(storages) + [p for p, _ in self._streams]
        n_entries   = len(all_entries)
        dir_sectors_count = math.ceil(n_entries / DIRS_PER_SECTOR)

        # Sector 0 = FAT, sectores 1..dir_sectors_count = directorio
        first_data_sector = 1 + dir_sectors_count
        sector_idx        = first_data_sector

        for path, data in self._streams:
            padded  = _pad_sector(data)
            n_sects = len(padded) // SECTOR_SIZE
            stream_sectors[path] = (sector_idx, n_sects, padded)
            sector_idx += n_sects

        total_sectors = sector_idx

        # ── Paso 2: construir FAT ─────────────────────────────────────────
        fat = [FREESECT] * SECTOR_SIZE  # máximo: 512/4 = 128 entradas → 512 bytes

        # Sector 0 = el FAT mismo
        fat[0] = FATSECT

        # Sectores de directorio: sectores 1..dir_sectors_count
        for i in range(1, dir_sectors_count + 1):
            fat[i] = i + 1 if i < dir_sectors_count else ENDOFCHAIN

        # Sectores de streams
        for path, (start, n, _) in stream_sectors.items():
            for i in range(n):
                fat[start + i] = start + i + 1 if i < n - 1 else ENDOFCHAIN

        fat_bytes = struct.pack('<' + 'I' * len(fat), *fat)
        fat_sector = fat_bytes[:SECTOR_SIZE]

        # ── Paso 3: construir entradas de directorio ──────────────────────
        # Mapa path → entry index
        entry_map = {e: i for i, e in enumerate(all_entries)}

        def _children_of(parent_path: str) -> list[str]:
            """Retorna entradas directas del parent dado."""
            result = []
            prefix = parent_path + '/' if parent_path else ''
            for e in all_entries:
                if not e:
                    continue
                if e.startswith(prefix):
                    rel = e[len(prefix):]
                    if '/' not in rel:
                        result.append(e)
            return result

        def _make_tree(entries: list[str]) -> Optional[int]:
            """
            Construye un árbol binario simple (degenerate, no balanceado)
            suficiente para CFB. Retorna el índice del nodo raíz del árbol.
            """
            if not entries:
                return NOSTREAM
            mid  = len(entries) // 2
            root = entries[mid]
            left = _make_tree(entries[:mid])
            right= _make_tree(entries[mid+1:])
            idx  = entry_map[root]
            # Actualizamos left/right del entry
            _dir_entries[idx].left  = left
            _dir_entries[idx].right = right
            return idx

        # Crear entradas (iniciales)
        _dir_entries: list[_DirEntry] = []
        for e in all_entries:
            if e == '':
                # Root entry
                _dir_entries.append(_DirEntry(
                    name="Root Entry",
                    entry_type=ENTRY_ROOT,
                    color=1,
                    clsid=ROOT_CLSID,
                    start=ENDOFCHAIN,
                    size=0,
                ))
            elif e in storages:
                name = e.split('/')[-1]
                _dir_entries.append(_DirEntry(
                    name=name,
                    entry_type=ENTRY_STORAGE,
                    color=1,
                ))
            else:
                name = e.split('/')[-1]
                start, n, padded = stream_sectors[e]
                _dir_entries.append(_DirEntry(
                    name=name,
                    entry_type=ENTRY_STREAM,
                    color=1,
                    start=start,
                    size=len(next(d for p, d in self._streams if p == e)),
                ))

        # Establecer child pointers
        for e in all_entries:
            idx      = entry_map[e]
            children = _children_of(e)
            if children:
                children_sorted = sorted(children, key=lambda x: x.split('/')[-1].upper())
                root_child = _make_tree(children_sorted)
                _dir_entries[idx].child = root_child
            else:
                _dir_entries[idx].child = NOSTREAM

        # Serializar directorio
        dir_data = bytearray()
        for de in _dir_entries:
            dir_data.extend(de.pack())
        # Padding hasta dir_sectors_count sectores con bytes 0x00
        # (0x00 = ENTRY_EMPTY válido en CFB — no es necesario agregar entradas extra)
        dir_data = bytearray(_pad_sector(bytes(dir_data)))
        dir_data = dir_data[:dir_sectors_count * SECTOR_SIZE]

        # ── Paso 4: construir header CFB (512 bytes) ─────────────────────
        # 109 FAT sector locations en el header (DIFAT)
        difat = [FREESECT] * 109
        difat[0] = 0  # el FAT está en el sector 0

        header = struct.pack(
            '<8s16sHHHHH6sIIIIIIIII',
            b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1',  # magic
            b'\x00' * 16,                            # CLSID (null)
            0x003E,                                  # minor version
            0x0003,                                  # major version (v3)
            0xFFFE,                                  # byte order LE
            0x0009,                                  # sector shift (512 bytes)
            0x0006,                                  # mini sector shift (64 bytes)
            b'\x00' * 6,                             # reserved (6 bytes)
            dir_sectors_count,                       # dir sectors count
            1,                                       # FAT sectors count
            1,                                       # first dir sector index
            0,                                       # transaction signature (unused)
            0x1000,                                  # mini stream cutoff (4096)
            ENDOFCHAIN & 0xFFFFFFFF,                 # first mini FAT sector
            0,                                       # mini FAT sectors count
            ENDOFCHAIN & 0xFFFFFFFF,                 # first DIFAT sector
            0,                                       # DIFAT sectors count
        )
        # El header tiene 76 bytes hasta aquí; necesita 512 bytes total.
        # Los últimos 436 bytes son las 109 entradas DIFAT (4 bytes c/u).
        header += struct.pack('<109I', *difat)
        assert len(header) == 512, f"Header incorrecto: {len(header)} bytes"

        # ── Paso 5: ensamblar el archivo completo ────────────────────────
        out = bytearray(header)
        out += fat_sector                             # sector 0: FAT
        out += bytes(dir_data)                        # sectores 1..N: directorio

        # Datos de cada stream en orden de sector
        stream_by_start = sorted(stream_sectors.values(), key=lambda x: x[0])
        for start, n, padded in stream_by_start:
            out += padded

        return bytes(out)


# ══════════════════════════════════════════════════════════════════════════════
# API pública
# ══════════════════════════════════════════════════════════════════════════════

def generar_vba_project_bin(modules: Dict[str, str]) -> bytes:
    """
    Genera un vbaProject.bin completo con los módulos VBA indicados.

    Args:
        modules: Diccionario {nombre_modulo: codigo_fuente_vba}

    Returns:
        Bytes del vbaProject.bin listo para inyectar en el ZIP del .xlsm.

    Ejemplo:
        data = generar_vba_project_bin({
            "modMain": Path("vba/modMain.bas").read_text(encoding="utf-8")
        })
    """
    # ── Construir streams internos ────────────────────────────────────────
    dir_stream_raw   = build_dir_stream(modules)
    dir_stream_comp  = ovba_compress(dir_stream_raw)
    vba_proj_stream  = build_vba_project_stream()
    project_stream   = build_project_stream(modules)
    projectwm_stream = build_project_wm_stream(modules)

    # ── Construir CFB ────────────────────────────────────────────────────
    cfb = CFBWriter()

    # Streams en el storage raíz
    cfb.add_stream("PROJECT",   project_stream)
    cfb.add_stream("PROJECTwm", projectwm_stream)

    # Streams dentro del storage "VBA"
    cfb.add_stream("VBA/_VBA_PROJECT", vba_proj_stream)
    cfb.add_stream("VBA/dir",          dir_stream_comp)

    for mod_name, source in modules.items():
        cfb.add_stream(f"VBA/{mod_name}", build_module_stream(source))

    return cfb.write()


# ══════════════════════════════════════════════════════════════════════════════
# CLI  (uso directo: python create_vba_project.py modMain.bas output.bin)
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Genera vbaProject.bin desde archivos .bas"
    )
    parser.add_argument("bas_files", nargs="+",
                        help="Archivos .bas a incluir (uno por módulo)")
    parser.add_argument("--output", "-o",
                        default="vbaProject.bin",
                        help="Ruta de salida (default: vbaProject.bin)")
    args = parser.parse_args()

    mods = {}
    for bas in args.bas_files:
        p = Path(bas)
        code = p.read_text(encoding="utf-8")
        # Quitar línea 'Attribute VB_Name = ...' si existe
        lines = [l for l in code.splitlines()
                 if not l.strip().startswith("Attribute VB_Name")]
        mods[p.stem] = "\n".join(lines)
        print(f"  + módulo: {p.stem} ({len(mods[p.stem])} chars)")

    data = generar_vba_project_bin(mods)
    Path(args.output).write_bytes(data)
    print(f"\nvbaProject.bin generado: {args.output} ({len(data)} bytes)")
