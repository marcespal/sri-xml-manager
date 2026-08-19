"""
progreso.py — Ventana de progreso (Tkinter) para tareas largas del motor.

Permite ejecutar una tarea en un hilo aparte mientras se muestra una ventana
con barra de avance. Si Tk no está disponible, ejecuta la tarea sin ventana.
La ventana NO depende de Excel — Excel queda libre porque el motor corre como
proceso independiente lanzado en segundo plano.
"""
import queue
import threading

import tkinter as tk
from tkinter import ttk

from utils import get_logger

logger = get_logger("progreso")


class _CanalGUI:
    """Canal de avance que envía mensajes a la ventana vía una cola thread-safe."""
    def __init__(self, q: queue.Queue):
        self._q = q

    def progreso(self, done, total, ok=0, err=0):
        self._q.put(("prog", done, total, ok, err))

    def mensaje(self, texto):
        self._q.put(("msg", texto))


class CanalNulo:
    """Canal sin ventana (no-op), para cuando no hay nada que mostrar."""
    def progreso(self, *a, **k):
        pass

    def mensaje(self, *a, **k):
        pass


def ejecutar_con_ventana(titulo: str, trabajo):
    """
    Ejecuta trabajo(canal) en un hilo y muestra una ventana Tk con barra de
    progreso hasta que la tarea termina. Retorna lo que retorne 'trabajo'
    (o relanza su excepción). Si Tk falla, ejecuta sin ventana.
    """
    q = queue.Queue()
    canal = _CanalGUI(q)
    box = {}

    def runner():
        try:
            box["val"] = trabajo(canal)
        except Exception as e:        # noqa: BLE001
            box["err"] = e
        finally:
            q.put(("fin",))

    try:
        root = tk.Tk()
    except Exception as e:            # sin display u otro problema con Tk
        logger.warning(f"Tk no disponible, ejecuto sin ventana: {e}")
        return trabajo(CanalNulo())

    root.title(titulo)
    root.resizable(False, False)
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    frm = ttk.Frame(root, padding=18)
    frm.pack(fill="both", expand=True)
    lbl = ttk.Label(frm, text="Iniciando...", font=("Segoe UI", 11))
    lbl.pack(anchor="w")
    pb = ttk.Progressbar(frm, length=440, mode="indeterminate")
    pb.pack(pady=12, fill="x")
    pb.start(14)
    sub = ttk.Label(frm, text="Conectando con el SRI...",
                    font=("Segoe UI", 9), foreground="#666")
    sub.pack(anchor="w")

    estado = {"det": False}
    threading.Thread(target=runner, daemon=True).start()

    def poll():
        try:
            while True:
                m = q.get_nowait()
                if m[0] == "prog":
                    _, done, total, ok, err = m
                    if not estado["det"]:
                        pb.stop()
                        pb.config(mode="determinate", maximum=max(total, 1))
                        estado["det"] = True
                    pb["value"] = done
                    lbl.config(text=f"Descargando  {done:,} de {total:,}")
                    sub.config(text=f"Descargadas: {ok:,}      Con error: {err:,}")
                elif m[0] == "msg":
                    lbl.config(text=m[1])
                elif m[0] == "fin":
                    lbl.config(text="Proceso completado")
                    try:
                        pb.stop()
                        pb.config(mode="determinate", maximum=1)
                        pb["value"] = 1
                    except Exception:
                        pass
                    sub.config(text="Esta ventana se cerrará automáticamente.")
                    root.after(1200, root.destroy)
                    return
        except queue.Empty:
            pass
        root.after(120, poll)

    root.after(150, poll)
    try:
        root.mainloop()
    except Exception as e:            # noqa: BLE001
        logger.warning(f"Error en ventana de progreso: {e}")

    if "err" in box:
        raise box["err"]
    return box.get("val")
