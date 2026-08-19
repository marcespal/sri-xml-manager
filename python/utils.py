"""
utils.py — Utilidades comunes: logging, helpers de texto y archivos
"""
import logging
import re
import time
from datetime import datetime
from pathlib import Path

import config


def get_logger(nombre: str = "sri_manager") -> logging.Logger:
    logger = logging.getLogger(nombre)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)-8s] %(name)s — %(message)s",
                             datefmt="%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    log_file = config.LOG_DIR / f"sri_{datetime.now():%Y%m%d}.log"
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


logger = get_logger()


def sanitizar(texto: str, max_len: int = 80) -> str:
    texto = re.sub(r'[\\/:*?"<>|]', "_", texto)
    return texto.strip()[:max_len]


def normalizar_fecha(fecha_str: str) -> str:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y%m%d"):
        try:
            return datetime.strptime(fecha_str.strip(), fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return fecha_str


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def esperar(segundos: float = 1.0):
    time.sleep(segundos)


def listar_xmls(directorio: Path) -> list:
    if not directorio.exists():
        return []
    return sorted(directorio.glob("*.xml"))


def claves_ya_descargadas(xml_dir: Path) -> set:
    return {f.stem for f in listar_xmls(xml_dir)}


def nombre_archivo_xml(clave_acceso: str, tipo: str, emisor: str, fecha: str) -> str:
    fecha_limpia = fecha.replace("/", "-").replace(" ", "_")[:10]
    emisor_limpio = sanitizar(emisor, 30)
    tipo_limpio = sanitizar(tipo, 10)
    return f"{fecha_limpia}_{tipo_limpio}_{emisor_limpio}_{clave_acceso[:8]}.xml"


class Progreso:
    def __init__(self, total: int, descripcion: str = "Procesando"):
        self.total = total
        self.actual = 0
        self.ok = 0
        self.error = 0
        self.descripcion = descripcion

    def avanzar(self, exitoso: bool = True):
        self.actual += 1
        if exitoso:
            self.ok += 1
        else:
            self.error += 1
        pct = (self.actual / self.total * 100) if self.total else 0
        barra = "█" * int(pct / 5) + "░" * (20 - int(pct / 5))
        print(
            f"\r{self.descripcion}: [{barra}] {pct:5.1f}% "
            f"({self.actual}/{self.total}) +{self.ok} x{self.error}",
            end="", flush=True
        )

    def finalizar(self):
        print()
        logger.info(
            f"{self.descripcion} completado: "
            f"{self.ok} exitosos, {self.error} errores de {self.total} total"
        )
