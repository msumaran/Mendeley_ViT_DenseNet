"""Agrega al documento técnico de Mendeley los ajustes acordados en la reunión del 30/10/2026:
explicación propia por etapa, explicación de la arquitectura caja por caja, origen del dataset
(3 → 2 clases), justificación del StratifiedKFold, cómo se obtiene cada columna de la tabla
general y por qué algunos folds terminan antes de 30 epochs.

Los textos nuevos van sombreados (amarillo) para ubicarlos rápido; lo que falta confirmar va
en rojo como [PENDIENTE: ...]. Uso (desde la raíz del repositorio): python utilitarios/documentacion/agregar_explicaciones_reunion.py
"""
import copy
import re
import shutil
import struct
import unicodedata
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Emu, Pt, RGBColor
from docx.table import _Cell
from docx.text.paragraph import Paragraph
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "diagramas"))
from generar_diagramas_seleccion_imagenes import TIPO_CAJA  # noqa: E402  (tipo de cada caja, igual que en el diagrama)

ROOT = Path(__file__).resolve().parents[2]
DOC = ROOT / "6_documentacion/Documentacion_Tecnica_10_Modelos.docx"
MARCA = "Explicación del autor"
SOMBRA = "FFF2CC"


def _shade(p):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), SOMBRA)
    # w:shd va antes de estos hijos de w:pPr (orden del esquema de Word)
    p._p.get_or_add_pPr().insert_element_before(
        shd, "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
        "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
        "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
        "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
        "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
        "w:pPrChange")


class Cursor:
    """Inserta párrafos y tablas en orden, a continuación de un elemento de referencia."""

    def __init__(self, doc, ref):
        self.doc, self.el, self.parent = doc, ref._p if hasattr(ref, "_p") else ref._tbl, ref._parent

    def p(self, *runs, style=None, shade=False):
        el = OxmlElement("w:p")
        self.el.addnext(el)
        self.el = el
        par = Paragraph(el, self.parent)
        if style:
            par.style = self.doc.styles[style]
        for item in runs:
            text, kind = (item, None) if isinstance(item, str) else item
            r = par.add_run(text)
            r.bold = kind == "b"
            if kind == "r":
                r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
                r.bold = True
        if shade and not style:
            _shade(par)
        return par

    def f(self, *formulas):
        """Líneas de fórmula con el formato del documento (Consolas, con sangría)."""
        for texto in formulas:
            par = self.p(texto, shade=False)
            par.runs[0].font.name = "Consolas"
            par.paragraph_format.left_indent = Emu(215900)

    def table(self, rows, bold_last=False):
        """Tabla con el mismo formato que las del documento: encabezado azul, texto blanco."""
        t = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = self.doc.styles["Table Grid"]
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = t.cell(i, j)
                cell.text = str(val)
                par = cell.paragraphs[0]
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = par.runs[0]
                run.font.size = Pt(9.5)
                run.bold = i == 0 or (bold_last and i == len(rows) - 1)
                if i == 0:
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:color"), "auto")
                    shd.set(qn("w:fill"), "2A4D7A")
                    cell._tc.get_or_add_tcPr().append(shd)
        self.el.addnext(t._tbl)
        self.el = t._tbl
        return t


def find(doc, prefix):
    return next(p for p in doc.paragraphs if p.text.strip().startswith(prefix))


def _tamano_png(datos):
    """(ancho, alto) en píxeles, leídos de la cabecera IHDR del PNG."""
    return struct.unpack(">II", datos[16:24])


def reemplazar_imagen(doc, blip, ruta):
    """Cambia la imagen de un dibujo y ajusta su alto a la nueva proporción (mismo ancho)."""
    datos = ruta.read_bytes()
    doc.part.related_parts[blip.get(qn("r:embed"))]._blob = datos
    ancho, alto = _tamano_png(datos)
    inline = next(a for a in blip.iterancestors() if a.tag in (qn("wp:inline"), qn("wp:anchor")))
    extent = inline.find(qn("wp:extent"))
    nuevo_cy = str(round(int(extent.get("cx")) * alto / ancho))
    extent.set("cy", nuevo_cy)
    for ext in inline.iter(qn("a:ext")):
        if ext.get("cy") is not None:
            ext.set("cy", nuevo_cy)


def quitar(p):
    p._p.getparent().remove(p._p)


def reemplazar_texto(doc, viejo, nuevo):
    """Reemplaza viejo por nuevo en los runs del cuerpo y de las tablas; devuelve cuántos cambió.
    Solo asigna r.text si el run contiene el texto (asignarlo borra imágenes del run)."""
    parrafos = list(doc.paragraphs) + [p for t in doc.tables for c in t._cells for p in c.paragraphs]
    n = 0
    for p in parrafos:
        for r in p.runs:
            if viejo in r.text:
                r.text = r.text.replace(viejo, nuevo)
                n += 1
    return n


def reemplazar_texto_parrafo(p, viejo, nuevo):
    """Igual que reemplazar_texto, pero en un solo párrafo."""
    n = 0
    for r in p.runs:
        if viejo in r.text:
            r.text = r.text.replace(viejo, nuevo)
            n += 1
    return n


def quitar_aclaracion_f1(doc):
    """La aclaración «la columna F1 es siempre el F1 macro» ya no aplica: las tablas por fold dicen
    «F1 macro» y en las tablas por clase «F1» es el de esa clase."""
    titulo = find(doc, "Aclaración importante")
    quitar(Paragraph(titulo._p.getnext(), titulo._parent))
    quitar(titulo)


def quitar_del_glosario(doc, termino):
    tabla = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Término")
    fila = next(f for f in tabla.rows if f.cells[0].text.strip() == termino)
    fila._tr.getparent().remove(fila._tr)


def enlace_ampliado(doc, despues_de, archivo):
    """Línea centrada bajo una figura con un enlace al PNG en alta resolución (ruta relativa al .docx)."""
    p = OxmlElement("w:p")
    despues_de.addnext(p)
    par = Paragraph(p, despues_de.getparent())
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), doc.part.relate_to(f"diagramas/{archivo}", RT.HYPERLINK, is_external=True))
    r, rpr, t = OxmlElement("w:r"), OxmlElement("w:rPr"), OxmlElement("w:t")
    for tag, val in (("w:color", "2E75B6"), ("w:sz", "18"), ("w:u", "single")):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        rpr.append(e)
    t.text = "Abrir el diagrama en tamaño completo"
    r.extend([rpr, t])
    h.append(r)
    p.append(h)
    return p


def reemplazar(p, viejo, nuevo):
    for r in p.runs:
        if viejo in r.text:
            r.text = r.text.replace(viejo, nuevo)
            return
    raise ValueError(f"No se encontró {viejo!r} en: {p.text[:80]}")


def mover_antes(desde, hasta_excluido, destino):
    """Mueve los elementos desde `desde` hasta antes de `hasta_excluido`, delante de `destino`."""
    el = desde._p
    while el is not hasta_excluido._p:
        siguiente = el.getnext()
        destino._p.addprevious(el)
        el = siguiente


def _sin_acentos(texto):
    return "".join(c for c in unicodedata.normalize("NFD", texto.lower()) if not unicodedata.combining(c))


def _fila_como(modelo_tr, textos, tabla):
    """Copia una fila (conserva su formato) y le pone los textos dados."""
    tr = copy.deepcopy(modelo_tr)
    for tc, texto in zip(tr.findall(qn("w:tc")), textos):
        par = _Cell(tc, tabla).paragraphs[0]
        par.runs[0].text = texto
        for r in par.runs[1:]:
            r._r.getparent().remove(r._r)
    return tr


def agregar_glosario(doc, entradas):
    """Inserta términos en el Anexo B respetando el orden alfabético."""
    tabla = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Término")
    for termino, definicion in entradas:
        filas = tabla.rows[1:]
        nueva = _fila_como(filas[0]._tr, (termino, definicion), tabla)
        sig = next((f for f in filas if _sin_acentos(f.cells[0].text) > _sin_acentos(termino)), None)
        (sig._tr.addprevious if sig is not None else filas[-1]._tr.addnext)(nueva)


def agregar_formulas(doc, filas_nuevas):
    """Agrega filas al final de la tabla de fórmulas de la sección 2.2."""
    tabla = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Métrica"
                 and t.cell(0, 1).text.strip() == "Fórmula")
    for textos in filas_nuevas:
        tabla.rows[-1]._tr.addnext(_fila_como(tabla.rows[-1]._tr, textos, tabla))


def figura(cur, ruta, pie):
    """Imagen centrada con su pie de figura, en la posición del cursor."""
    p = cur.p("")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(ruta), width=Cm(15))
    c = cur.p(pie)
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.runs[0].italic, c.runs[0].font.size = True, Pt(9)
    c.runs[0].font.color.rgb = RGBColor(0x59, 0x74, 0x82)


def con_tipo(titulo, tipos):
    """«Cargador. » → «Cargador (componente). »; usa el mismo tipo que el ícono del diagrama."""
    base = titulo.split(" (")[0].rstrip(". ")
    tipo = tipos.get(base) or ("registro" if base.startswith("Registros") else None)
    if not tipo:
        return titulo
    return titulo.replace(" (", f" ({tipo}, ", 1) if " (" in titulo else f"{base} ({tipo}). "


def autor(cur, *runs):
    cur.p((f"{MARCA}. ", "b"), *runs, shade=True)


def fin_de_etapa(doc, prefix):
    """Último elemento de la etapa: antes del siguiente título «N.N …» o «N. …» o del salto de página."""
    el = find(doc, prefix)._p
    while (sig := el.getnext()) is not None and sig.tag != qn("w:sectPr"):
        if sig.tag == qn("w:p"):
            p = Paragraph(sig, doc._body)
            if p.style.name.startswith("Heading") and re.match(r"^\d+(\.\d+)?\.?\s", p.text.strip()):
                break
            if any(b.get(qn("w:type")) == "page" for b in sig.iter(qn("w:br"))):
                break
        el = sig
    return el


# Apertura de cada etapa: qué se hace, de qué parte y con qué termina (el documento se explica solo)
QUE_SE_HACE = {
    "1.1 Etapa 1": (
        "Se reúnen las fotos de ojos de pescado y se dejan en dos clases, fresco y no fresco, con la "
        "misma cantidad en cada una. Parte de las 4,390 fotos originales de Mendeley y termina con 8,000 "
        "imágenes etiquetadas, 4,000 por clase."),
    "1.2 Etapa 2": (
        "Se prepara todo lo necesario antes de entrenar. Por un lado, las 8,000 imágenes se reparten en "
        "5 folds (grupos), para que cada imagen se pruebe una vez con un modelo que no la vio al entrenar "
        "(1.2.1). Por otro, se configuran los 10 modelos candidatos con la misma receta de entrenamiento "
        "(1.2.2). Termina con 5 particiones de los datos y 10 modelos listos para entrenar."),
    "1.3 Etapa 3": (
        "Cada uno de los 10 modelos aprende a distinguir fresco de no fresco en cada uno de los 5 folds: "
        "50 entrenamientos en total. Después de cada época se mide cómo le va con las imágenes de "
        "validación y se guarda la mejor versión del modelo (checkpoint). Termina con 50 checkpoints, "
        "uno por modelo y fold."),
    "1.4 Etapa 4": (
        "Se mide qué tan bien funciona cada modelo con imágenes que nunca vio. La mejor versión de cada "
        "fold clasifica una sola vez el 20 % de test que se reservó desde el inicio. Termina con la "
        "predicción y la probabilidad de cada imagen de test."),
    "1.5 Etapa 5": (
        "Se juntan los resultados de los 5 folds para obtener un solo resultado por modelo. Como cada "
        "imagen se probó una sola vez, hay 8,000 predicciones por modelo; con ellas se calculan la matriz "
        "de confusión, el F1 macro y el AUC de cada modelo."),
    "1.6 Etapa 6": (
        "Se comparan los 10 modelos con los mismos criterios: F1 macro, AUC, estabilidad entre folds "
        "(desviación estándar del F1) y tamaño (número de parámetros). Con esa comparación se eligen dos "
        "modelos de referencia, uno de cada familia, que pasan a las pruebas con Trucha y Rohu (1.6.1)."),
    "1.7 Etapa 7": (
        "Se guardan y ordenan los resultados para que cualquier número del documento se pueda rastrear "
        "hasta el archivo de donde salió: tablas CSV, gráficas PNG y registros de ejecución."),
}

# Nota interna del autor (resumen informal; no reemplaza la explicación del documento)
ETAPAS = {
    "1.1 Etapa 1": "El proyecto busca detectar si una trucha está fresca a partir de la foto de su ojo. En "
                   "este primer experimento los programas aprenden con un conjunto más grande: 4,390 fotos de "
                   "ojos de 8 especies, agrupadas en fresco y no fresco. Como había muchas más fotos frescas "
                   "que no frescas, se generaron copias con pequeños cambios (giros, volteos, brillo) hasta "
                   "tener 4,000 de cada clase, para que ambas tengan el mismo peso.",
    "1.2 Etapa 2": "Las 8,000 fotos se reparten en 5 lotes, cada uno con la misma proporción de fresco y no "
                   "fresco. En cada ronda, un lote se reserva para la evaluación final y los otros cuatro se "
                   "usan para aprender; así, cada foto se evalúa una sola vez. También se preparan los 10 "
                   "programas a comparar, que ya vienen entrenados para reconocer imágenes en general.",
    "1.3 Etapa 3": "Cada programa aprende a distinguir fresco de no fresco revisando las fotos de entrenamiento "
                   "varias veces. Después de cada revisión se mide su resultado con un grupo de fotos aparte. "
                   "Si mejora, se guarda esa versión; si pasan 7 revisiones seguidas sin mejora, se detiene. "
                   "El máximo es de 30 revisiones.",
    "1.4 Etapa 4": "La mejor versión guardada de cada programa clasifica el lote reservado, con fotos que "
                   "nunca vio durante el aprendizaje. Esta evaluación se hace una sola vez y muestra si el "
                   "programa reconoce la frescura en fotos nuevas, y no solo en las que ya conocía.",
    "1.5 Etapa 5": "Se juntan los resultados de los 5 lotes. Como cada foto se evaluó una sola vez, se obtiene "
                   "una respuesta del programa para cada una de las 8,000 fotos. Con eso se cuentan los "
                   "aciertos y los errores en fresco y en no fresco, y se calcula el puntaje final de cada "
                   "programa.",
    "1.6 Etapa 6": "Se comparan los 10 programas por su puntaje, por qué tan estables fueron sus resultados "
                   "entre los 5 lotes y por su tamaño. Se eligen los dos mejores, uno de cada tipo: "
                   "ViT-Base/16 (96.96) y DenseNet-121 (95.66). Estos dos se prueban después con Trucha y "
                   "Rohu.",
    "1.7 Etapa 7": "Se ordenan los resultados en tablas y gráficos, junto con los registros generados durante "
                   "el aprendizaje y la evaluación. Cada número del documento puede ubicarse en su archivo de "
                   "origen.",
}

ARQUITECTURA = [
    ("Repositorio Mendeley (antes del experimento). ",
     "4,390 fotos originales del ojo, de 8 especies, en tres estados de frescura."),
    ("Dos clases (antes del experimento). ",
     "Muy fresco y fresco se agrupan en fresco (3,084 fotos); no fresco queda igual (1,306)."),
    ("Aumento de datos (antes del experimento). ",
     "Se generan 916 fotos de fresco y 2,694 de no fresco con volteo, giro de ±20°, brillo, color y "
     "desenfoque, para llegar a 4,000 por clase (ver 1.1)."),
    ("Fotos DNC. ",
     "La carpeta con las 8,000 fotos; la clase de cada foto sale del nombre de su carpeta."),
    ("Cargador. ",
     "Anota la ruta y la clase de cada foto y se las entrega a «Dividir en 5 folds». Durante el "
     "entrenamiento abre las fotos por lotes y las ajusta al tamaño y a la escala de colores de cada modelo."),
    ("Dividir en 5 folds. ",
     "En cada fold separa entrenamiento (68 %), validación (12 %) y test (20 %). Entrenamiento y "
     "validación van a «Entrenar»; el test queda reservado y va directo a «Examen con test» (flecha "
     "superior)."),
    ("Preparar modelos. ",
     "Carga los 10 modelos con su conocimiento previo (preentrenados): 3 Transformers y 7 CNN. No "
     "recibe fotos; solo entrega el modelo a «Entrenar»."),
    ("Entrenar. ",
     "Cada modelo aprende con las fotos de entrenamiento del fold, hasta 30 épocas. Al final de cada "
     "época mide el F1 en validación y para si pasan 7 épocas sin mejora."),
    ("Mejor versión. ",
     "«Entrenar» guarda aquí la versión del modelo con mejor F1 de validación (checkpoint): una por "
     "modelo y fold, 50 en total. Esa versión es la que rinde el examen."),
    ("Examen con test. ",
     "Carga la mejor versión y clasifica una sola vez las fotos de test de ese fold, que el modelo "
     "nunca vio."),
    ("Reunir 5 folds. ",
     "Junta las respuestas de los 5 folds: 8,000 por modelo, una por foto. Con ellas calcula la matriz "
     "de confusión, el F1 y el AUC."),
    ("Tablas. ",
     "Guarda las métricas de cada fold y del total en archivos CSV, que alimentan la comparación y las "
     "gráficas."),
    ("Comparar 10 modelos. ",
     "Recibe las métricas de las tablas y compara F1, AUC, estabilidad entre folds y tamaño de cada "
     "modelo. Elige el mejor de cada familia (ViT-Base/16 y DenseNet-121) y pasa la comparación a "
     "«Documentación»."),
    ("Gráficas. ",
     "Dibuja las matrices de confusión, las curvas ROC y las curvas por época a partir de las tablas."),
    ("Documentación. ",
     "Este documento: reúne la comparación, las gráficas y los registros de ejecución."),
    ("Registros de ejecución. ",
     "Lo que se anota en cada época del entrenamiento y en cada examen (2_logs_entrenamiento/). "
     "Respaldan cada número del documento (flechas punteadas)."),
]

# (modelo, fold, época elegida, epochs corridos) — tablas 3.x.1 de este documento
EARLY_STOP = [
    ("ViT-Base/16", 2, 16, 23),
    ("Swin-Base", 4, 18, 25),
    ("EfficientNet-Lite0", 3, 16, 23),
    ("ResNet-50", 3, 21, 28),
]

# Anexo A — (dato, sección, archivo en este repositorio, script que lo genera)
ORIGEN_DATOS = [
    ("Dato", "Sección", "Archivo", "Script"),
    ("Hiperparámetros", "1.2.2", "—", "run_freshness_v3.py (V3_CFG)"),
    ("val_F1 y paciencia por época", "1.3 y 3.x.1",
     "2_logs_entrenamiento/dnc_5fold_resume.log y dnc_reentreno_curvas.log", "dnc_5fold_full_metrics.py"),
    ("Matriz de confusión por fold", "1.4 y 3.x.2",
     "3_resultados_test_por_modelo/{modelo}/confusion_matrix_fold{N}.csv", "dnc_5fold_confmat_roc.py"),
    ("Curvas ROC por fold y pooled", "3.x.3 y 3.x.4",
     "3_resultados_test_por_modelo/{modelo}/roc_curve_fold{N}.csv y roc_curve_pooled.csv",
     "dnc_5fold_confmat_roc.py"),
    ("Matriz pooled", "1.5 y 3.x.4", "3_resultados_test_por_modelo/{modelo}/confusion_matrix_pooled.csv",
     "dnc_5fold_confmat_roc.py"),
    ("Métricas por fold (50 filas)", "3.x.2", "3_resultados_test_por_modelo/metricas_por_fold.csv",
     "dnc_5fold_confmat_roc.py + metrics.py"),
    ("Métricas pooled (tabla general)", "1.5 y 1.6",
     "3_resultados_test_por_modelo/metricas_pooled_por_modelo.csv", "dnc_5fold_confmat_roc.py + metrics.py"),
    ("Registro de la generación de matrices y ROC", "1.4", "2_logs_entrenamiento/dnc_confmat_roc_all.log",
     "dnc_5fold_confmat_roc.py"),
    ("Gráficas (matrices, ROC, AUC)", "3.x", "4_graficas/, 5_graficas_informe/, 7_graficas_por_modelo/",
     "utilitarios/graficas/dnc_graficas_organizadas.py"),
    ("Figuras de arquitectura, flujo y vista simple", "Introducción y 1",
     "6_documentacion/diagramas/ (SVG y PNG)",
     "utilitarios/diagramas/generar_diagramas_seleccion_imagenes.py"),
    ("Fórmulas de las métricas", "2", "—", "metrics.py (compute_metrics)"),
]

GLOSARIO_NUEVO = [
    ("Aumento de datos (data augmentation)", "Generación de copias modificadas de fotos originales (volteo, giro, "
     "brillo, color, desenfoque) para completar la clase con menos imágenes. En Mendeley lleva cada clase a "
     "4,000 imágenes; en Trucha y Rohu no se aplica."),
    ("Variaciones al azar en entrenamiento", "Cambios aleatorios (recorte, volteo, color, rotación, desenfoque, "
     "borrado) que se aplican a cada foto de entrenamiento en cada época; no crean fotos nuevas y no se "
     "aplican a validación ni a test."),
    ("Dominio fuente", "Conjunto con el que se entrenan los modelos: Mendeley (8,000 imágenes)."),
    ("Dominio objetivo", "Conjunto nuevo en el que se prueba lo aprendido: Trucha o Rohu (documento de "
     "generalización)."),
    ("p-valor", "Probabilidad de observar una diferencia como la medida si en realidad los dos modelos "
     "rindieran igual; si es menor que 0.05, la diferencia se considera significativa."),
    ("Prueba de Wilcoxon", "Prueba estadística que compara dos modelos medidos en los mismos folds e indica si "
     "la diferencia entre ellos podría deberse al azar; se aplica en el documento de generalización."),
    ("TRD (tasa de retención del desempeño)", "Porcentaje del F1 de Mendeley que un modelo conserva en un "
     "dominio objetivo: F1_objetivo / F1_fuente × 100 (documento de generalización)."),
    ("Cargador (DataLoader)", "Componente que toma las rutas y etiquetas en memoria y entrega las imágenes "
     "al modelo por lotes, ya transformadas."),
    ("Color jitter", "Variación al azar que cambia al azar brillo, contraste, saturación y tono de la imagen."),
    ("CSV", "Archivo de texto con datos en columnas separadas por comas; aquí guarda métricas, matrices y "
     "curvas ROC."),
    ("DenseNet", "Familia de CNN en la que cada capa recibe la salida de todas las anteriores; aquí se usa "
     "DenseNet-121."),
    ("EfficientNet", "Familia de CNN que escala de forma equilibrada profundidad, ancho y resolución; aquí "
     "se usan EfficientNet-B0 y EfficientNet-Lite0."),
    ("Época", "Ver Epoch."),
    ("Estados de frescura", "Categorías originales del repositorio: muy fresco, fresco y no fresco. Para "
     "este estudio se homologan a dos clases: fresco y no fresco."),
    ("Estratificado", "Reparto que mantiene en cada parte la misma proporción de clases que el conjunto "
     "completo."),
    ("Función de pérdida", "Medida del error del modelo durante el entrenamiento; el optimizador ajusta "
     "los pesos para reducirla. Aquí se usa Focal Loss."),
    ("Homologación de estados", "Paso de los tres estados de frescura a las dos clases del estudio."),
    ("Logs (registros de ejecución)", "Archivos de texto donde el entrenamiento y la evaluación anotan, "
     "época por época, sus resultados."),
    ("Lote (batch)", "Grupo de imágenes que el modelo procesa a la vez; su tamaño es el batch size."),
    ("MobileNet", "Familia de CNN livianas pensadas para dispositivos móviles; aquí se usan MobileNetV1, "
     "MobileNetV2 y MobileNetV3-Small."),
    ("Normalización", "Ajuste de los valores de los píxeles a la media y desviación estándar de ImageNet "
     "antes de entrar al modelo."),
    ("Optimizador", "Algoritmo que actualiza los pesos del modelo a partir del gradiente; aquí AdamW."),
    ("PNG", "Formato de imagen sin pérdida usado para las gráficas."),
    ("Random crop", "Variación al azar que recorta al azar una parte de la imagen ampliada."),
    ("Random erasing", "Variación al azar que borra al azar un rectángulo de la imagen."),
    ("Redimensionado (resize)", "Cambio del tamaño de la imagen al que espera el modelo (224 o 256 px)."),
    ("ResNet", "Familia de CNN con conexiones residuales que saltan capas; aquí se usa ResNet-50."),
    ("stdev", "Desviación estándar muestral (n − 1)."),
]

# ViT-Base/16: 3_resultados_test_por_modelo/vit_base_patch16_224/confusion_matrix_fold{k}.csv
# y metricas_por_fold.csv
VIT_FOLDS = [
    (1, 777, 23, 20, 780, "97.31", "0.9910"),
    (2, 769, 31, 24, 776, "96.56", "0.9913"),
    (3, 773, 27, 22, 778, "96.94", "0.9878"),
    (4, 766, 34, 16, 784, "96.87", "0.9877"),
    (5, 776, 24, 22, 778, "97.12", "0.9925"),
]


INTRO_ORGANIZACION = [
    ("Sección", "Qué contiene"),
    ("Introducción", "De qué trata el documento, el resultado principal y los conceptos clave"),
    ("1. Metodología", "Las 7 etapas del experimento, desde los datos hasta la documentación"),
    ("2. Glosario de fórmulas", "Cómo se calcula cada métrica, con ejemplos"),
    ("3. Detalle por modelo", "Los números de cada uno de los 10 modelos"),
    ("4. Conclusiones", "Qué se encontró y qué modelos se eligieron"),
    ("Anexo A", "De dónde sale cada dato"),
    ("Anexo B", "Glosario de términos técnicos"),
]

CONCEPTOS = [
    ("Concepto", "En palabras simples", "Más detalle"),
    ("Clasificación fresco / no fresco", "El modelo mira la foto del ojo del pescado y responde una de dos "
     "opciones: fresco o no fresco.", "1.1"),
    ("Modelo (CNN y Transformer)", "Programa que aprende a reconocer patrones en imágenes a partir de "
     "ejemplos. Se probaron dos familias: las CNN (redes convolucionales) y los Transformers.", "1.2.2"),
    ("Entrenamiento, validación y test", "Como estudiar, hacer prácticas y rendir el examen final: el modelo "
     "aprende con el entrenamiento (68 % de las fotos), se mide su avance con la validación (12 %) y al "
     "final se evalúa con el test (20 %), fotos que nunca vio.", "1.2.1"),
    ("Fold y validación cruzada", "Las fotos se dividen en 5 partes (folds) y el experimento se repite 5 "
     "veces, cada vez con una parte distinta como examen. Así todas las fotos se evalúan una vez.", "1.2.1"),
    ("Época (epoch)", "Una vuelta completa del modelo por todas las fotos de entrenamiento. Cada modelo "
     "entrena hasta 30 épocas.", "1.3"),
    ("Checkpoint", "La copia guardada del modelo en su mejor época; es la que se evalúa.", "1.3"),
    ("Paciencia (early stopping)", "Si el modelo no mejora durante 7 épocas seguidas, el entrenamiento se "
     "detiene antes de llegar a 30.", "1.3"),
    ("Matriz de confusión", "Tabla que cuenta aciertos y errores: cuántas fotos frescas se reconocieron "
     "como frescas, cuántas se confundieron con no frescas, y lo mismo para las no frescas.", "2.1"),
    ("F1 macro", "Nota de 0 a 100 que resume los aciertos de las dos clases con el mismo peso. Es la "
     "métrica principal del documento.", "2.2"),
    ("Accuracy", "Porcentaje de fotos bien clasificadas, contando las dos clases juntas.", "2.2"),
    ("Precisión y recall", "Precisión: de las fotos que el modelo llamó fresco, cuántas lo eran. Recall: de las "
     "fotos que eran frescas, cuántas encontró. Lo mismo se calcula para no fresco.", "2.2"),
    ("AUC", "Nota de 0 a 1 que indica qué tan bien el modelo separa fresco de no fresco usando sus "
     "probabilidades; 1 es perfecto.", "2.4"),
    ("Resultado pooled", "Se juntan las respuestas de los 5 exámenes (8,000 fotos) y se calcula una sola "
     "nota general por modelo.", "2.5"),
]


def para_publico_general(doc):
    """Agrega lo que necesita alguien que no es del área para seguir el documento de principio a fin."""
    # Introducción antes de la metodología
    metodologia = find(doc, "1. Metodología")
    cur = Cursor(doc, metodologia)
    cur.el = metodologia._p.getprevious()
    cur.p("Introducción", style="Heading 1")
    cur.p(("¿De qué trata este documento? ", "b"),
          "La tesis busca saber si un modelo de inteligencia artificial puede decir, mirando la foto del "
          "ojo de un pescado, si está fresco o no, incluso con fotos distintas a las que usó para aprender. "
          "Este documento cuenta el primer paso: se entrenaron y compararon 10 modelos con 8,000 fotos de "
          "ojos de pescado del repositorio Mendeley (8 especies, Indonesia) con un solo propósito: elegir "
          "los dos modelos con los que se harán las pruebas con otros peces (trucha y Rohu).")
    cur.p(("Objetivo de la tesis que responde. ", "b"),
          "Este documento responde al objetivo específico 1 del plan de tesis: «Evaluar el desempeño de "
          "diversos modelos de Machine Learning para identificar aquellos con mejores resultados en la "
          "clasificación de la frescura a partir de imágenes oculares». Su producto es la línea base "
          "experimental: los modelos con mejor desempeño (plan de tesis, Tabla 9). Los objetivos "
          "específicos 2, 3 y 4 se desarrollan en los documentos de generalización y de adaptación de "
          "dominio (CORAL).")
    cur.p(("Resultado principal. ", "b"),
          "ViT-Base/16 fue el que más acertó: F1 macro de 96.96 %, es decir, acierta en unas 97 de cada 100 "
          "fotos. DenseNet-121 fue la mejor CNN (F1 macro de 95.66 %) y la que mejor separa las dos clases "
          "(AUC de 0.9906). Se eligieron ambos, uno de cada familia, para los experimentos siguientes.")
    cur.p("Por qué esta etapa se diseñó así", style="Heading 2")
    cur.p("Esta etapa no busca el modelo final de la tesis: su función es elegir los dos modelos de "
          "referencia para las pruebas con otros peces. Las decisiones del diseño se derivan de ese "
          "propósito:", shade=False)
    for titulo, texto in [
        ("Una comparación justa entre los 10 modelos. ",
         "Todos se entrenan y evalúan exactamente igual: las mismas 8,000 imágenes, los mismos 5 folds, la "
         "misma semilla y la misma receta de entrenamiento. Así, las diferencias entre ellos se deben a la "
         "arquitectura y no a las condiciones del experimento."),
        ("5 folds. ",
         "Para ordenar los 10 modelos basta con que cada imagen se evalúe una vez. Con 5 folds se cumple "
         "eso con 50 entrenamientos, la mitad de los que exigirían 10 folds."),
        ("Solo deep learning. ",
         "Las etapas siguientes parten de los modelos elegidos: en generalización se conserva lo que "
         "aprendieron en Mendeley y solo se reentrena su última capa, y en adaptación de dominio se ajustan "
         "también capas internas. Eso requiere redes profundas preentrenadas (CNN y Transformers)."),
        ("Un modelo por familia. ",
         "Se elige el mejor Transformer y la mejor CNN para comparar los dos enfoques en las etapas "
         "siguientes; el modelo propuesto de la tesis es un Vision Transformer."),
        ("La generalización se mide después. ",
         "La pregunta central de la tesis, si el modelo sigue funcionando con fotos distintas, no se "
         "responde aquí, sino en el documento de generalización, con fotos de trucha y Rohu que los modelos "
         "no vieron en ningún momento de su entrenamiento en Mendeley."),
    ]:
        cur.p((titulo, "b"), texto, style="List Bullet")
    cur.p("Cómo está organizado el documento", style="Heading 2")
    cur.table(INTRO_ORGANIZACION)
    cur.p("Conceptos clave", style="Heading 2")
    cur.p("Estos son los términos que más se repiten. La última columna indica la sección donde se "
          "explican en detalle; el resto de términos está en el glosario del Anexo B.", shade=False)
    cur.table(CONCEPTOS)


    # Frases «en palabras simples» antes de los párrafos más técnicos
    for inicio, texto in [
        ("Hiperparámetros comunes",
         "En palabras simples: todos los modelos se entrenan con la misma receta, para que la comparación "
         "sea justa. Los hiperparámetros son los ajustes de esa receta (qué tan rápido aprende el modelo, "
         "cuánto se castigan los errores, etc.). Las variaciones al azar en entrenamiento cambian cada foto "
         "(giros, recortes, cambios de color) para que el modelo aprenda a reconocer el ojo y no memorice las "
         "fotos."),
        ("Uso de memoria durante el entrenamiento",
         "En palabras simples: cuánta memoria de la tarjeta gráfica necesita cada modelo para entrenar. No "
         "cambia los resultados; sirve para saber qué equipo hace falta para usar cada modelo."),
    ]:
        ref = find(doc, inicio)
        cur = Cursor(doc, ref)
        cur.el = ref._p.getprevious()
        cur.p(texto)

    # Justificaciones (borradores para que el autor las revise)
    cur = Cursor(doc, find(doc, "Se evalúan 10 modelos candidatos"))
    cur.p(("Por qué estos 10 modelos. ", "b"),
          "Las 7 CNN son las arquitecturas más usadas en los estudios sobre frescura de pescado revisados "
          "en el plan de tesis (Tabla 2): DenseNet-121 y ResNet-50 (16.67 % de los estudios cada una), la "
          "familia MobileNet (~22 %) y EfficientNet (11.11 %) (Prasetyo et al., 2022b; Hamidy et al., 2025; "
          "Lin et al., 2025; Choompol et al., 2025). Los 3 Transformers (ViT-Base/16, Swin-Base y "
          "DINOv2-Small) se incluyen porque esa misma revisión encontró poca exploración de arquitecturas "
          "basadas en atención, pese a su potencial para generalizar, que es el foco de la tesis. Además, "
          "los 10 modelos cubren un rango amplio de tamaños, entre 1.5 y 86.7 millones de parámetros, lo que "
          "permite comparar desempeño y costo.")
    cur = Cursor(doc, find(doc, "Data augmentation solo en entrenamiento"))
    cur.p(("Por qué estos hiperparámetros. ", "b"),
          "Los 10 modelos usan la misma configuración, para que las diferencias se deban a la arquitectura "
          "y no al ajuste de cada uno. Un learning rate bajo (1e-4) con AdamW es habitual al ajustar modelos "
          "preentrenados: permite aprender la tarea sin borrar lo aprendido en ImageNet. El warmup de 3 "
          "épocas evita cambios bruscos al inicio y el scheduler coseno reduce el paso al final para afinar. "
          "Focal Loss da más peso a las imágenes difíciles de clasificar y el label smoothing evita que el "
          "modelo se vuelva demasiado confiado. El límite de 30 épocas, junto con la paciencia de 7, da a "
          "todos los modelos el mismo presupuesto de entrenamiento. Estos valores se fijaron tras pruebas "
          "preliminares y luego se mantuvieron iguales para los 10 modelos.")

    cur = Cursor(doc, find(doc, "Se seleccionan ViT-Base/16 (Transformer) y DenseNet-121"))
    cur.p("La selección se basa en comparar los 10 modelos entre sí, todos con el mismo protocolo. Las "
          "cifras de los dos elegidos en Mendeley se usan después como punto de referencia (dominio fuente) "
          "para medir cuánto de ese desempeño conservan en Trucha y Rohu.")
    cur.p(("Aumento de datos y comparación entre modelos. ", "b"),
          "Las 3,610 imágenes aumentadas se generaron antes de dividir en folds, así que una imagen aumentada "
          "y la foto de la que salió pueden quedar en folds distintos. Esto se aplica por igual a los 10 "
          "modelos (mismas imágenes, mismos folds y misma semilla), por lo que no cambia el orden entre ellos, "
          "que es lo que decide la selección. Además, el F1 de Mendeley es el denominador de TRD en el "
          "documento de generalización: si estuviera algo inflado, TRD saldría más bajo para los dos modelos, es "
          "decir, la medición de cuánto conservan es conservadora.")
    cur.p(("Por qué aquí no se aplica la prueba de Wilcoxon. ", "b"),
          "El objetivo de esta etapa es elegir los dos modelos de referencia, no demostrar que uno es "
          "superior al otro. Además, con 5 folds el p-valor más bajo posible es 0.0625, que no alcanza el "
          "nivel α = 0.05 del plan. La prueba se aplica en el documento de generalización, donde se compara "
          "ViT-Base/16 con DenseNet-121 en Trucha y Rohu.")

    # Figuras con el paso previo de selección de imágenes (generar_diagramas_seleccion_imagenes.py)
    figuras = ROOT / "6_documentacion" / "diagramas"
    for p in doc.paragraphs:  # los diagramas ya no están en 7_graficas_por_modelo/
        for r in p.runs:
            if "SVG editable en 7_graficas_por_modelo/" in r.text:
                r.text = r.text.replace("7_graficas_por_modelo/", "6_documentacion/diagramas/")
    blips = list(doc.element.body.iter(qn("a:blip")))[:4]  # flujo, arquitectura y sus dos detalles
    parrafo = lambda blip: next(a for a in blip.iterancestors() if a.tag == qn("w:p"))  # noqa: E731
    reemplazar_imagen(doc, blips[0], figuras / "flujo_metodologico_con_seleccion.png")
    reemplazar_imagen(doc, blips[1], figuras / "arquitectura_sistema_con_seleccion.png")
    enlace_ampliado(doc, parrafo(blips[1]), "arquitectura_sistema_con_seleccion.png")
    # Los detalles ampliados (pie arriba de cada imagen) sobran: el enlace abre el diagrama completo
    for blip in blips[2:]:
        p = parrafo(blip)
        p.getparent().remove(p.getprevious())
        p.getparent().remove(p)
    # Vista simple para público general, tras el resultado principal (después del cambio de figuras)
    cur = Cursor(doc, find(doc, "Resultado principal."))
    figura(cur, figuras / "vista_simple_mendeley.png",
           "El experimento Mendeley en 6 pasos. El detalle de cada etapa está en la sección 1.")
    pie = find(doc, "Flujo metodológico del experimento Mendeley")
    pie.runs[0].text = ("Flujo metodológico del experimento Mendeley: la selección y el balanceo de las imágenes, "
                        "antes del experimento, y las siete etapas, desde la preparación de datos hasta la "
                        "documentación de resultados.")
    for r in pie.runs[1:]:
        r._r.getparent().remove(r._r)
    tabla = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Etapa del flujo")
    tabla.rows[0]._tr.addnext(_fila_como(tabla.rows[1]._tr, ("0. Antes del experimento: selección y "
                                                              "balanceo de imágenes", "1.1"), tabla))

    # Guía de lectura de la sección 3
    cur = Cursor(doc, find(doc, "3. Detalle por modelo"))
    cur.p(("Cómo leer esta sección. ", "b"),
          "Cada modelo tiene cuatro partes, siempre en el mismo orden: 3.x.1 muestra el entrenamiento, es "
          "decir, cómo cambió el val_F1 época por época y cuál época se eligió; 3.x.2 da la nota de cada "
          "uno de los 5 exámenes (folds) con su matriz de confusión; 3.x.3 muestra las curvas ROC de cada "
          "fold; y 3.x.4 suma los 5 exámenes para obtener la nota general que aparece en la tabla de 1.6. "
          "Los modelos van de mayor a menor F1. Para ver solo lo esencial, basta con los dos modelos "
          "elegidos: ver 3.1 y 3.2.")

    # Conclusiones antes de los anexos
    anexo_a = find(doc, "Anexo A.")
    cur = Cursor(doc, anexo_a)
    cur.el = anexo_a._p.getprevious()
    cur.p("4. Conclusiones", style="Heading 1")
    for texto in [
        "Los 10 modelos alcanzan al menos 89 % de F1 macro; los cuatro mejores quedan entre 95.4 % y 97.0 %.",
        "ViT-Base/16 obtiene el F1 macro más alto (96.96 %) y es muy estable entre folds (desviación "
        "estándar de 0.28).",
        "DenseNet-121 es la mejor CNN (F1 macro de 95.66 %) y tiene el AUC más alto de los 10 modelos "
        "(0.9906), con solo 7.0 millones de parámetros.",
        "Se eligen ViT-Base/16 y DenseNet-121, un Transformer y una CNN, como punto de partida para "
        "evaluar si lo aprendido con Mendeley sirve en trucha y Rohu (documento de generalización).",
    ]:
        cur.p(texto, style="List Bullet")

    # Equipo y software: dentro de «Arquitectura del sistema», tras la explicación de la arquitectura
    equipo = find(doc, "Especificaciones técnicas del entorno de ejecución")
    equipo.text = "Equipo y software usados"
    equipo.style = doc.styles["Heading 3"]


def actualizar():
    doc = Document(DOC)
    if any(p.text.startswith(MARCA) for p in doc.paragraphs):
        print("Ya actualizado; no se hace nada.")
        return
    shutil.copy(DOC, DOC.with_suffix(".antes_reunion.docx"))

    # Tarea 6 — explicación propia al inicio de cada etapa
    # Tarea 4 — origen del dataset y paso de 3 a 2 clases
    cur = Cursor(doc, find(doc, "Se usa el dataset DNC"))
    cur.p(("Origen del conjunto y paso de tres a dos clases. ", "b"),
          "Las imágenes provienen del repositorio público Mendeley Data (Prasetyo, Adityo, Suciati y "
          "Fatichah, 2022a; Indonesia), con fotos de la región ocular ya recortada, de 8 especies y en tres "
          "estados: muy fresco, fresco y no fresco. Para la clasificación binaria se homologan los estados: "
          "«Highly Fresh» y «Fresh» pasan a fresco y «Not Fresh» a no fresco (_label_from_folder en "
          "dataset_freshness.py, usada por collect_dnc_only en run_freshness_v3.py). En disco hay 4,390 "
          "imágenes originales: 3,084 fresco y 1,306 no fresco. El plan de tesis (Tabla 5) indica 4,392; "
          "la diferencia está en Corvina amarilla (Nibea albiflora), no fresco, con 121 imágenes en la "
          "carpeta frente a 123 en el plan.")
    cur.p(("Balanceo a 8,000 imágenes. ", "b"),
          "Las dos clases se llevaron a 4,000 imágenes cada una para balancear la cantidad de fresco y no "
          "fresco, ya que no fresco tenía menos de la mitad de imágenes que fresco (1,306 frente a 3,084). "
          "Las imágenes que faltaban se generaron con aumento de datos (script balance_dnc_4k.py, del "
          "proyecto original; no incluido en este repositorio):")
    cur.table([("Clase", "Originales", "Aumentadas", "Total"),
               ("Fresco", "3,084", "916", "4,000"),
               ("No fresco", "1,306", "2,694", "4,000"),
               ("Total", "4,390", "3,610", "8,000")], bold_last=True)
    cur.p("Las 3,610 imágenes aumentadas se repartieron en proporción a cada especie. Cada una se obtiene "
          "de una foto original aplicando volteo, rotación de ±20°, cambios de brillo, contraste y color, y "
          "desenfoque, con semilla 42; se guardan en JPEG con calidad 82. El resultado es la carpeta "
          "DataSetNotClean_4k/, que usa todo el experimento.", shade=False)

    # Tarea 5 — «Limitación» → justificación; sin párrafos de limitaciones (pedido del autor)
    lim = find(doc, "Limitación: la partición")
    lim.text = ("Justificación: se usa StratifiedKFold porque mantiene la proporción 50/50 de fresco y "
                "no fresco en cada fold y asegura que cada imagen esté en test exactamente una vez, lo "
                "que da 8,000 predicciones fuera de fold por modelo.")
    Cursor(doc, lim).p(("Por qué 5 folds. ", "b"),
                       "El objetivo de esta etapa es elegir los dos modelos para las pruebas con otros "
                       "peces (ver «Por qué esta etapa se diseñó así», en la Introducción). Para eso basta "
                       "con 5 folds: los 10 modelos se comparan con el mismo protocolo, cada imagen se "
                       "evalúa una vez y se hacen 50 entrenamientos en lugar de los 100 que exigirían 10 "
                       "folds.")
    quitar(find(doc, "Alcance de la verificación"))

    # Tarea 1 — arquitectura caja por caja
    cur = Cursor(doc, find(doc, "Especificaciones técnicas del entorno de ejecución"))
    cur.el = cur.el.getprevious()  # va antes de las especificaciones, tras las figuras
    cur.p("Explicación de la arquitectura, componente por componente", style="Heading 3")
    cur.p("El diagrama se lee de izquierda a derecha: primero la selección y el balanceo de las imágenes, "
          "después las siete etapas. La flecha continua es paso de datos y "
          "la punteada, un registro que se guarda o se consulta. El ícono de cada caja indica su tipo: "
          "componente (programa que hace una tarea), función (paso de cálculo), dataset (conjunto "
          "de fotos), archivo (resultado guardado) o registro (logs de ejecución).")
    for titulo, texto in ARQUITECTURA:
        cur.p((con_tipo(titulo, TIPO_CAJA), "b"), texto)

    # Ejemplo de paciencia: de párrafo corrido a tabla
    pac = find(doc, "La columna Paciencia de las tablas")
    texto = pac.text
    pac.text = texto[:texto.index("Ejemplo,")].strip()
    cur = Cursor(doc, pac)
    cur.p("Ejemplo — ViT-Base/16, fold 3 (2_logs_entrenamiento/dnc_5fold_resume.log; curva completa "
          "en 3.1.1):", shade=False)
    cur.table([("Época", "val_F1", "Paciencia (pat)", "Qué pasa"),
               (2, "84.79", 0, "Mejora: se guarda el checkpoint"),
               (3, "81.84", 1, "No mejora"),
               (4, "78.01", 2, "No mejora"),
               (5, "87.79", 0, "Mejora: se guarda el checkpoint y la paciencia vuelve a 0")])
    cur.p(texto[texto.index("Dos valores impresos"):], shade=False)

    # Tarea 8 — por qué algunos folds terminan antes de 30 epochs
    cur.p(("Por qué algunos folds terminan antes de 30 epochs. ", "b"),
          "Cada fold puede correr como máximo 30 epochs; con 5 folds, el máximo por modelo es 150. "
          "Un total menor indica que en algún fold actuó el early stopping: el val_F1 no mejoró "
          "durante 7 epochs seguidas después de la mejor. De las 50 corridas, 46 llegaron a 30 "
          "epochs y 4 se detuvieron antes. En las 4 se cumple epochs corridos = época elegida + 7:")
    cur.table([("Modelo", "Fold", "Época elegida (mejor val_F1)", "Epochs corridos",
                "Comprobación")] +
              [(m, f, e, t, f"{e} + 7 = {t}") for m, f, e, t in EARLY_STOP])
    cur.p("Por eso ViT-Base/16 suma 143 epochs (150 − 7), Swin-Base 145, EfficientNet-Lite0 143 y "
          "ResNet-50 148; los otros 6 modelos completan 150. Detenerse antes no cambia el resultado "
          "del fold: siempre se usa el checkpoint de la mejor época, no el de la última. Las curvas "
          "de esos folds están en 3.1.1, 3.3.1, 3.6.1 y 3.8.1.")

    # Tarea 7 — cómo se obtiene cada columna de la tabla general
    cur = Cursor(doc, find(doc, "Filas resaltadas: los 2 modelos elegidos"))
    cur.p("Cómo se obtiene cada columna de esta tabla", style="Heading 3")
    cur.p("Cada número se calcula una sola vez sumando los 5 folds (ViT-Base/16, datos de 3.1.2 y 3.1.4):")
    cur.table([("Fold", "TP", "FN", "FP", "TN", "F1 macro del fold", "AUC del fold")] +
              [(f"Fold {k}", tp, fn, fp, tn, f1, auc) for k, tp, fn, fp, tn, f1, auc in VIT_FOLDS] +
              [("Suma (pooled)", "3,861", "139", "104", "3,896", "—", "—")], bold_last=True)
    cur.p("Con esa matriz sumada se obtiene cada columna de la tabla:", shade=False)
    cur.table([
        ("Columna", "Cómo se obtiene", "Cálculo (ViT-Base/16)", "Valor"),
        ("Accuracy", "Aciertos de la matriz sumada / total", "(3,861 + 3,896) / 8,000 × 100", "96.96"),
        ("Precisión macro", "Promedio de la precisión de cada clase (matriz sumada)",
         "fresco 3,861 / (3,861 + 104) = 97.38\nno fresco 3,896 / (3,896 + 139) = 96.56\n"
         "(97.38 + 96.56) / 2", "96.97"),
        ("F1 macro (pooled)", "Promedio del F1 de cada clase (matriz sumada)",
         "(F1_fresco + F1_no_fresco) / 2", "96.96"),
        ("Desv. std F1", "Desviación estándar muestral (n − 1) de los 5 F1 de fold",
         "stdev(97.31, 96.56, 96.94, 96.87, 97.12)", "0.28"),
        ("AUC (pooled)", "Con las 8,000 probabilidades concatenadas (roc_curve_pooled.csv), no con "
                         "la matriz", "Área bajo la ROC pooled", "0.9900"),
    ])
    cur.p("El F1 pooled no es el promedio de los F1 de fold, aunque en ViT-Base/16 coinciden (media de "
          "los 5 = 96.96); el AUC pooled tampoco es el promedio de los AUC de fold (sería 0.9901). El "
          "mismo procedimiento se aplica a los 10 modelos; la suma de las matrices de cada uno está "
          "en su sección 3.x.4.", shade=False)

    # Ejemplos con enlace a la sección de donde salen los datos
    find(doc, "Ejemplo instanciado — DenseNet-121").add_run(" (datos de 3.2.2 y 3.2.4)")
    p = find(doc, "La función confusion_matrix de scikit-learn")
    reemplazar(p, "confusion_matrix_fold1.csv:", "confusion_matrix_fold1.csv (fold 1 de ViT-Base/16, "
                                                 "sección 3.1.2):")
    p = find(doc, "El código metrics.py, función compute_metrics()")
    reemplazar(p, "confusion_matrix_pooled.csv de ViT contiene",
               "confusion_matrix_pooled.csv de ViT (ver 3.1.4) contiene")

    # Glosario 2.2.1: ejemplo de párrafo corrido a líneas de fórmula
    ej = find(doc, "Ejemplo — ViT-Base/16, matriz pooled de 8,000")
    ej.text = "Ejemplo — ViT-Base/16, matriz pooled de 8,000 imágenes (datos de 3.1.4):"
    cur = Cursor(doc, ej)
    cur.f("Precisión_fresco = 3,861 / (3,861 + 104) × 100 = 97.38 %",
          "Precisión_no_fresco = 3,896 / (3,896 + 139) × 100 = 96.56 %",
          "Precisión_macro = (97.38 + 96.56) / 2 = 96.97 %")
    cur.p("Para no_fresco se toma esa clase como positiva: TP = 3,896 y FP = 139. La precisión macro "
          "se promedia con los valores sin redondear.", shade=False)

    # Glosario 2.5: fórmulas del resultado pooled en líneas separadas
    pooled = find(doc, "La fórmula de la matriz es")
    pooled.text = "El resultado pooled de cada modelo se calcula así:"
    cur = Cursor(doc, pooled)
    cur.f("M_pooled = Σₖ₌₁⁵ Mₖ (suma celda a celda de las 5 matrices de fold)",
          "F1_clase = 2·TP / (2·TP + FP + FN) (cada clase tomada como positiva)",
          "F1_macro_pooled = (F1_fresco + F1_no_fresco) / 2")
    cur.p("El F1 macro pooled no tiene por qué coincidir con el promedio de los F1 de fold. El AUC "
          "pooled se calcula con las probabilidades concatenadas, no a partir de la matriz.", shade=False)

    # Orden: arquitectura y entorno de ejecución van antes de las etapas (vista general primero)
    bloque = find(doc, "Arquitectura del sistema y entorno de ejecución")
    mover_antes(bloque, find(doc, "1.3 Etapa 3"), find(doc, "1.1 Etapa 1"))
    bloque.style = doc.styles["Heading 2"]

    bloque.text = "Arquitectura del sistema"

    # Público general: introducción, guía de lectura, conclusiones y frases «en palabras simples»
    para_publico_general(doc)

    # Anexo A: de dónde sale cada dato, con las rutas de este repositorio
    intro = find(doc, "Todos los números de este documento provienen")
    viejo = intro._p.getnext()
    intro.text = ("Todos los números de este documento provienen de archivos reales. La tabla indica, para "
                  "cada dato, la sección donde aparece, el archivo del repositorio de donde sale y el "
                  "script que lo genera.")
    cur = Cursor(doc, intro)
    cur.table(ORIGEN_DATOS)
    cur.p("En el servidor de entrenamiento, los archivos de 3_resultados_test_por_modelo/ "
          "y 2_logs_entrenamiento/ están en results_freshness_v3/reportes/dnc_5fold_confmat_roc/ y "
          "results_freshness_v3/logs_5metricas/.", shade=False)
    viejo.getparent().remove(viejo)

    # Notas de cómo se recuperaron algunas curvas: se explican en la exposición, no en el documento
    for p in doc.paragraphs:
        for r in p.runs:  # asignar r.text borra lo que no es texto (imágenes): solo si cambia
            nuevo = re.sub(r" ?(Los folds 1–2 vienen|Curvas) de un reentreno.*?son las originales\.", "", r.text)
            if nuevo != r.text:
                r.text = nuevo
    glosario = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Término")
    for fila in list(glosario.rows):
        if fila.cells[0].text.strip() in ("Reentreno", "MPS"):
            fila._tr.getparent().remove(fila._tr)

    # Glosario: términos usados en el documento que no estaban definidos
    # Revisión: datos corregidos, términos unificados y aclaraciones que ya no aplican
    for viejo, nuevo in (("(línea 102)", "(línea 106)"), ("(línea ~118)", "(línea 133)"),
                         ("(línea 144)", "(línea 148)"),  # dnc_5fold_full_metrics.py
                         ("22.1M", "21.6M"),  # DINOv2-Small se construye con img_size=224 (models.py)
                         ("Data augmentation solo en entrenamiento:", "Variaciones al azar en entrenamiento:")):
        assert reemplazar_texto(doc, viejo, nuevo), viejo
    for p in doc.paragraphs:  # Swin-Base y DINOv2-Small no son ViT: la familia es Transformer
        if p.style.name == "Heading 2" and p.text.endswith("(ViT)"):
            assert reemplazar_texto_parrafo(p, "(ViT)", "(Transformer)")
    tipo = next(t for t in doc.tables if [c.text.strip() for c in t.rows[0].cells][:2] == ["Modelo", "Tipo"])
    for fila in tipo.rows[1:]:
        if fila.cells[1].text.strip() == "ViT":
            reemplazar_texto_parrafo(fila.cells[1].paragraphs[0], "ViT", "Transformer")
    quitar_aclaracion_f1(doc)
    quitar_del_glosario(doc, "Data augmentation")
    agregar_glosario(doc, GLOSARIO_NUEVO)
    agregar_formulas(doc, [("Desv. std del F1 (n − 1)",
                            "√( Σ (F1ᵢ − media)² / (n − 1) ), con n = 5 folds",
                            "Cuánto varía el F1 entre folds; un valor bajo indica un modelo estable")])

    for prefix, texto in QUE_SE_HACE.items():
        Cursor(doc, find(doc, prefix)).p(("Qué se hace en esta etapa. ", "b"), texto)

    # Un solo comentario del autor por etapa, al final de la etapa
    for prefix, texto in ETAPAS.items():
        cur = Cursor(doc, find(doc, prefix))
        cur.el = fin_de_etapa(doc, prefix)
        autor(cur, texto)

    doc.save(DOC)
    print(f"Actualizado: {DOC}")


if __name__ == "__main__":
    actualizar()
