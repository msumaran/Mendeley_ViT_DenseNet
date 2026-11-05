"""Ajustes acordados en la reunión del 04/11/2026 en la documentación de Mendeley: aclarar el código del paso 2,
redactar mejor el paso 9 y agregar el origen de los pesos preentrenados (Anexo C, enlazado desde el paso 7).
Se corre después de notas_autor.py.

Uso (desde la raíz del repositorio): python utilitarios/documentacion/ajustes_reunion_04_11.py
"""
import copy
import sys
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "herramientas_docx"))
from actualizar_30_epocas import global_, poner  # noqa: E402
from formato_formulas_y_enlaces import _hipervinculo, _run_con_texto  # noqa: E402

DOC = Path("6_documentacion/Documentacion_Tecnica_10_Modelos.docx")
MARCA = "anexo_c"
# Etiquetas que timm descargó para cada modelo (caché de Hugging Face del equipo de las corridas de 30 épocas)
PESOS = [
    ("ViT-Base/16", "vit_base_patch16_224", "augreg2_in21k_ft_in1k", "ImageNet-21k, ajustado en ImageNet-1k"),
    ("Swin-Base", "swin_base_patch4_window7_224", "ms_in22k_ft_in1k", "ImageNet-22k, ajustado en ImageNet-1k"),
    ("DINOv2-Small", "vit_small_patch14_dinov2", "lvd142m", "LVD-142M, auto-supervisado sin etiquetas"),
    ("DenseNet-121", "densenet121", "ra_in1k", "ImageNet-1k"),
    ("ResNet-50", "resnet50", "a1_in1k", "ImageNet-1k"),
    ("MobileNetV1", "mobilenetv1_100", "ra4_e3600_r224_in1k", "ImageNet-1k"),
    ("MobileNetV2", "mobilenetv2_100", "ra_in1k", "ImageNet-1k"),
    ("MobileNetV3-Small", "mobilenetv3_small_100", "lamb_in1k", "ImageNet-1k"),
    ("EfficientNet-B0", "efficientnet_b0", "ra_in1k", "ImageNet-1k"),
    ("EfficientNet-Lite0", "efficientnet_lite0", "ra_in1k", "ImageNet-1k"),
]


def parrafo(doc, inicio):
    ps = [p for p in doc.paragraphs if p.text.startswith(inicio)]
    assert len(ps) == 1, inicio
    return ps[0]


def poner_p(p, t):
    runs = [r for r in p.runs if r.text] or p.runs
    runs[0].text = t
    for r in runs[1:]:
        r.text = ""


def enlace_externo(doc, celda, url):
    p = celda.paragraphs[0]
    poner(celda, "")
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), doc.part.relate_to(url, RT.HYPERLINK, is_external=True))
    h.append(_hipervinculo(p.runs[0]._r, url, "x")[0])
    p._p.append(h)


def anexo_c(doc):
    anexo_b = parrafo(doc, "Anexo B.")
    anexo_a = parrafo(doc, "Anexo A.")
    ultimo = next(e for e in anexo_b._p.itersiblings() if e.tag == qn("w:tbl"))  # tabla del glosario
    titulo = copy.deepcopy(anexo_b._p)
    ultimo.addnext(titulo)
    titulo = Paragraph(titulo, anexo_b._parent)
    for b in titulo._p.findall(qn("w:bookmarkStart")) + titulo._p.findall(qn("w:bookmarkEnd")):
        titulo._p.remove(b)
    poner_p(titulo, "Anexo C. Origen de los pesos preentrenados")
    ident = str(max(int(b.get(qn("w:id"))) for b in doc.element.body.iter(qn("w:bookmarkStart"))) + 1)
    ini, fin = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    ini.set(qn("w:id"), ident), ini.set(qn("w:name"), MARCA), fin.set(qn("w:id"), ident)
    titulo._p.find(qn("w:pPr")).addnext(ini)
    titulo._p.append(fin)

    intro = copy.deepcopy(anexo_a._p.getnext())
    titulo._p.addnext(intro)
    poner_p(Paragraph(intro, anexo_b._parent),
          "Los 10 modelos parten de pesos preentrenados públicos. La librería timm los descarga de Hugging Face "
          "la primera vez que se crea cada modelo (build_model en models.py, con pretrained=True). La tabla indica, "
          "para cada modelo, su nombre en timm, la versión de los pesos, con qué imágenes se entrenaron y el "
          "enlace de descarga.")

    molde = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Dato")
    tabla = copy.deepcopy(molde._tbl)
    intro.addnext(tabla)
    tabla = Table(tabla, anexo_b._parent)
    for fila in tabla.rows[2:]:
        tabla._tbl.remove(fila._tr)
    for celda, texto in zip(tabla.rows[0].cells, ("Modelo", "Nombre en timm", "Pesos (entrenados con)", "Enlace")):
        poner(celda, texto)
    plantilla = copy.deepcopy(tabla.rows[1]._tr)
    tabla._tbl.remove(tabla.rows[1]._tr)
    for modelo, nombre, etiqueta, datos in PESOS:
        tabla._tbl.append(copy.deepcopy(plantilla))
        celdas = tabla.rows[-1].cells
        for c in celdas:  # la plantilla puede traer enlaces internos
            for h in c._tc.iter(qn("w:hyperlink")):
                h.getparent().replace(h, h[0])
            for rpr in c._tc.iter(qn("w:rPr")):
                for e in rpr.findall(qn("w:u")) + rpr.findall(qn("w:color")):
                    rpr.remove(e)
        poner(celdas[0], modelo)
        poner(celdas[1], nombre)
        poner(celdas[2], f"{etiqueta} ({datos})")
        enlace_externo(doc, celdas[3], f"https://huggingface.co/timm/{nombre}.{etiqueta}")


def enlazar_paso_7(doc):
    p = parrafo(doc, "7. Preparar modelos")
    normal = next(r for r in reversed(p.runs) if not r.bold and r.text.strip())
    p._p.append(_run_con_texto(normal._r, " De dónde se descargó cada modelo preentrenado: ver "))
    p._p.append(_hipervinculo(normal._r, "Anexo C", MARCA))
    p._p.append(_run_con_texto(normal._r, "."))


def main():
    assert not list(DOC.parent.glob("~$cumentacion_Tecnica_10_Modelos*")), "cierra el documento en Word"
    doc = Document(DOC)
    assert not any(b.get(qn("w:name")) == MARCA for b in doc.element.body.iter(qn("w:bookmarkStart"))), "ya aplicado"
    global_(doc, "Código: _label_from_folder en dataset_freshness.py.",
            "Código: función _label_from_folder del archivo dataset_freshness.py; recibe el nombre de la carpeta y "
            "devuelve la clase: si contiene «Not Fresh», no fresco; si contiene «Fresh», fresco.")
    global_(doc, "Esta versión, no la de la última época, es la que pasa al examen.",
            "Esa mejor versión es la que pasa al examen, y no siempre es la de la última época: si el mejor F1 de "
            "validación se logra en la época 24, se examina la versión de la época 24, aunque el entrenamiento "
            "siga hasta la 30.")
    enlazar_paso_7(doc)
    anexo_c(doc)
    doc.save(DOC)


if __name__ == "__main__":
    main()
