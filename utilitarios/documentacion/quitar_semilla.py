"""Quita la mención «semilla 42» del texto de Documentacion_Tecnica_10_Modelos.docx (el código no cambia).
Se corre después de ajustes_reunion_04_11.py.

Uso (desde la raíz del repositorio): python utilitarios/documentacion/quitar_semilla.py
"""
import sys
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parent))
from actualizar_30_epocas import global_  # noqa: E402

DOC = Path("6_documentacion/Documentacion_Tecnica_10_Modelos.docx")

if __name__ == "__main__":
    assert not list(DOC.parent.glob("~$cumentacion_Tecnica_10_Modelos*")), "cierra el documento en Word"
    doc = Document(DOC)
    global_(doc, "StratifiedKFold (semilla 42) reparte", "StratifiedKFold reparte")
    global_(doc, "y desenfoque, con semilla 42; se guardan", "y desenfoque; se guardan")
    doc.save(DOC)
