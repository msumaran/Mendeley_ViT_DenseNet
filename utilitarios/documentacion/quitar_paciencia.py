"""Quita de la documentación de Mendeley las referencias a la paciencia y a corridas que no llegaron a
30 épocas: todas las corridas vigentes completan 30. Se corre después de actualizar_30_epocas.py.

Uso (desde la raíz del repositorio): python utilitarios/documentacion/quitar_paciencia.py
"""
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "herramientas_docx"))
from actualizar_30_epocas import cambiar, fila_de, global_, poner  # noqa: E402
from lectura_amigable import _al_ancho  # noqa: E402

CARPETA = Path("6_documentacion")


def tecnica():
    ruta = CARPETA / "Documentacion_Tecnica_10_Modelos.docx"
    doc = Document(ruta)
    # Sección «Cómo funciona la paciencia» y explicación de las corridas detenidas
    cuerpo = list(doc.element.body.iterchildren())
    texto = lambda el: Paragraph(el, doc).text.strip() if el.tag == qn("w:p") else ""
    ini = next(i for i, el in enumerate(cuerpo) if texto(el) == "Cómo funciona la paciencia")
    fin = next(i for i, el in enumerate(cuerpo) if texto(el) == "Corridas experimentales por modelo")
    assert 0 < fin - ini < 40
    for el in cuerpo[ini:fin]:
        doc.element.body.remove(el)
    # Columna Paciencia de las tablas por época
    n = 0
    for t in doc.tables:
        if t.rows[0].cells[-1].text.strip() == "Paciencia":
            rejilla = t._tbl.tblGrid
            ancho = sum(int(c.get(qn("w:w"))) for c in rejilla.findall(qn("w:gridCol")))
            for tr in t._tbl.tr_lst:
                tr.remove(tr.tc_lst[-1])
            rejilla.remove(rejilla.findall(qn("w:gridCol"))[-1])
            _al_ancho(t._tbl, ancho)
            n += 1
    assert n == 50, n
    # Glosarios
    for t in doc.tables:
        for fila in list(t.rows):
            if fila.cells[0].text.strip() in ("Paciencia", "Paciencia (early stopping)", "Early stopping"):
                t._tbl.remove(fila._tr)
    global_(doc, "Cada modelo entrena hasta 30 épocas.", "Cada modelo entrena 30 épocas.")
    global_(doc, "val_F1 y paciencia por época", "val_F1 por época")
    for f, parte in ((2, "parte2"), (4, "parte2"), (3, "parte1")):
        global_(doc, f"El fold {f} viene de 2_logs_entrenamiento/completar_30ep_{parte}.log, que lo completa hasta 30 épocas.",
                f"El fold {f} viene de 2_logs_entrenamiento/completar_30ep_{parte}.log.")
    entorno = next(t for t in doc.tables if t.cell(1, 0).text.strip() == "Hardware")
    poner(fila_de(entorno, "Repetición de 4 corridas a 30 épocas").cells[0], "4 corridas de completar_30ep_*.log")
    doc.save(ruta)


def resumen():
    ruta = CARPETA / "Sustento_Seleccion_Modelos.docx"
    doc = Document(ruta)
    global_(doc, "30 epochs por corrida (las 4 detenidas por paciencia=7 se repitieron hasta 30)", "30 epochs por corrida")
    doc.save(ruta)
    ruta = CARPETA / "Tabla_Mendeley_10_Modelos_ViT_DenseNet.docx"
    doc = Document(ruta)
    global_(doc, " sin cambios en esta tabla (las 4 corridas detenidas por paciencia se repitieron hasta 30 épocas)", " en esta tabla")
    doc.save(ruta)


if __name__ == "__main__":
    tecnica()
    resumen()
