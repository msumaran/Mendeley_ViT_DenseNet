"""Agrega al Anexo C de Documentacion_Tecnica_10_Modelos.docx el artículo y el repositorio original de cada una
de las 10 arquitecturas (pedido de la reunión del 04/11). Se corre después de quitar_semilla.py.

Uso (desde la raíz del repositorio): python utilitarios/documentacion/repositorios_anexo_c.py
"""
import copy
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, str(Path(__file__).resolve().parent))
from actualizar_30_epocas import poner  # noqa: E402
from ajustes_reunion_04_11 import enlace_externo, poner_p  # noqa: E402

DOC = Path("6_documentacion/Documentacion_Tecnica_10_Modelos.docx")
TF_SLIM = "https://github.com/tensorflow/models/tree/master/research/slim/nets/mobilenet"
REPOS = [
    ("ViT-Base/16", "Dosovitskiy et al., 2021 (ICLR)", "https://github.com/google-research/vision_transformer"),
    ("Swin-Base", "Liu et al., 2021 (ICCV)", "https://github.com/microsoft/Swin-Transformer"),
    ("DINOv2-Small", "Oquab et al., 2023", "https://github.com/facebookresearch/dinov2"),
    ("DenseNet-121", "Huang et al., 2017 (CVPR)", "https://github.com/liuzhuang13/DenseNet"),
    ("ResNet-50", "He et al., 2016 (CVPR)", "https://github.com/KaimingHe/deep-residual-networks"),
    ("MobileNetV1", "Howard et al., 2017",
     "https://github.com/tensorflow/models/blob/master/research/slim/nets/mobilenet_v1.md"),
    ("MobileNetV2", "Sandler et al., 2018 (CVPR)", TF_SLIM),
    ("MobileNetV3-Small", "Howard et al., 2019 (ICCV)", TF_SLIM),
    ("EfficientNet-B0", "Tan y Le, 2019 (ICML)",
     "https://github.com/tensorflow/tpu/tree/master/models/official/efficientnet"),
    ("EfficientNet-Lite0", "Variante de EfficientNet para dispositivos móviles (Google, 2020)",
     "https://github.com/tensorflow/tpu/tree/master/models/official/efficientnet/lite"),
]
INTRO = ("Los pesos de la tabla anterior los publica timm (https://github.com/huggingface/pytorch-image-models). "
         "Cada arquitectura fue propuesta antes por sus autores, con su propio código; la tabla siguiente indica el "
         "artículo y el repositorio original de cada una.")


def main():
    assert not list(DOC.parent.glob("~$cumentacion_Tecnica_10_Modelos*")), "cierra el documento en Word"
    doc = Document(DOC)
    assert not any(p.text == INTRO for p in doc.paragraphs), "ya aplicado"
    pesos = next(t for t in doc.tables if t.cell(0, 1).text.strip() == "Nombre en timm")
    intro_c = pesos._tbl.getprevious()
    intro = copy.deepcopy(intro_c)
    pesos._tbl.addnext(intro)
    poner_p(Paragraph(intro, pesos._parent), INTRO)

    tbl = copy.deepcopy(pesos._tbl)
    intro.addnext(tbl)
    tabla = Table(tbl, pesos._parent)
    # 4 → 3 columnas: se quita la de pesos y se reparte su ancho
    grid = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    extra = int(grid[2].get(qn("w:w")))
    tbl.find(qn("w:tblGrid")).remove(grid[2])
    grid[1].set(qn("w:w"), str(int(grid[1].get(qn("w:w"))) + extra))
    for fila in tabla.rows:
        tcs = fila._tr.findall(qn("w:tc"))
        fila._tr.remove(tcs[2])
        ancho = tcs[1].find(qn("w:tcPr")).find(qn("w:tcW"))
        if ancho is not None:
            ancho.set(qn("w:w"), str(int(ancho.get(qn("w:w"))) + extra))
    for celda, texto in zip(tabla.rows[0].cells, ("Modelo", "Artículo original", "Repositorio original")):
        poner(celda, texto)
    assert [f.cells[0].text.strip() for f in tabla.rows[1:]] == [m for m, _, _ in REPOS]
    for fila, (_, articulo, url) in zip(tabla.rows[1:], REPOS):
        poner(fila.cells[1], articulo)
        for h in list(fila.cells[2]._tc.iter(qn("w:hyperlink"))):
            h.getparent().replace(h, h[0])  # queda un run simple que enlace_externo usa de molde
        enlace_externo(doc, fila.cells[2], url)
    doc.save(DOC)


if __name__ == "__main__":
    main()
