"""Ajustes acordados en la reunión del 03/11/2026 en la documentación de Mendeley: explicación numerada y
narrativa de la arquitectura, cargador y tipos de caja en el glosario, tabla de equipo y lugar de la prueba
de Wilcoxon. Se corre después de quitar_paciencia.py.

Uso (desde la raíz del repositorio): python utilitarios/documentacion/ajustes_reunion_03_11.py
"""
import copy
import hashlib
import re
import sys
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "herramientas_docx"))
from actualizar_30_epocas import fila_de, global_, poner  # noqa: E402
from formato_formulas_y_enlaces import _hipervinculo  # noqa: E402

DOC = Path("6_documentacion/Documentacion_Tecnica_10_Modelos.docx")
PRIMERO = "Repositorio Mendeley ("
FIN = "Equipo y software usados"
INTRO = ("El diagrama se lee de izquierda a derecha y cada caja lleva el mismo número que aquí, en el orden en que "
         "ocurre: primero la selección y el balanceo de las imágenes, después las siete etapas. ")

ARQUITECTURA = [
    ("Repositorio Mendeley (dataset, antes del experimento)",
     "es el directorio con los datos tal como se descargaron del repositorio público Mendeley Data. Guarda 4,390 "
     "fotos originales de la región del ojo, de 8 especies, separadas en carpetas según su estado de frescura: muy "
     "fresco, fresco y no fresco. Todo el experimento parte de estas fotos (ver {1.1})."),
    ("Dos clases (función, antes del experimento)",
     "las fotos del paso 1 tienen tres estados, pero el experimento clasifica solo dos. La función lee el nombre de "
     "la carpeta de cada foto y le asigna su clase: muy fresco y fresco pasan a «fresco» (3,084 fotos) y no fresco "
     "se queda igual (1,306). Se trabaja con dos clases porque Rohu, usado en las pruebas siguientes, solo tiene dos "
     "estados, y porque tras un experimento previo con tres y con dos clases se optó por dos. Código: "
     "_label_from_folder en dataset_freshness.py."),
    ("Aumento de datos (componente, antes del experimento)",
     "después del paso 2 hay muchas más fotos de fresco que de no fresco, y el modelo tendería a favorecer a la clase "
     "mayor. Para igualarlas se crean fotos nuevas a partir de las originales, con volteo, giro de ±20°, cambios de "
     "brillo, contraste y color, y desenfoque: 916 de fresco y 2,694 de no fresco, repartidas en proporción a cada "
     "especie. Así cada clase llega a 4,000 fotos (ver {1.1})."),
    ("Fotos DNC (dataset, etapa 1)",
     "es el resultado del paso 3: la carpeta DataSetNotClean_4k, con 8,000 fotos balanceadas, 4,000 por clase. La "
     "clase de cada foto sale del nombre de su carpeta. Es el único conjunto de fotos que usan los 10 modelos."),
    ("Cargador (componente, etapa 1)",
     "«cargador» es el nombre que usamos aquí para el código que lleva las fotos hasta el modelo; no es una librería, "
     "sino funciones de Python en dos archivos. Primero, collect_dnc_only (run_freshness_v3.py) recorre la carpeta del "
     "paso 4 y anota la ruta y la clase de cada foto en una lista, sin abrirlas; esa lista pasa al paso 6. Después, "
     "durante el entrenamiento, FreshnessDataset (dataset_freshness.py) y DataLoader de PyTorch abren las fotos por "
     "lotes (8 en cada lote en Transformers, 16 en CNN) y get_transforms_v3 las ajusta a lo que espera cada modelo: "
     "224 píxeles en Transformers y 256 en CNN, con el color normalizado. Solo en entrenamiento agrega variaciones al "
     "azar."),
    ("Dividir en 5 folds (función, etapa 2)",
     "con la lista del paso 5, StratifiedKFold (semilla 42) reparte las 8,000 fotos en 5 partes de 1,600, con la misma "
     "proporción de fresco y no fresco. En cada una de las 5 repeticiones, una parte es el test (20 %) y el resto se "
     "separa en entrenamiento (68 %, 5,440 fotos) y validación (12 %, 960 fotos). Entrenamiento y validación van al "
     "paso 8; el test queda reservado y va directo al paso 10 (flecha superior). Código: dnc_5fold_full_metrics.py."),
    ("Preparar modelos (componente, etapa 2)",
     "en paralelo, se cargan las 10 arquitecturas con sus pesos preentrenados, es decir, con lo que ya aprendieron de "
     "imágenes generales: 3 Transformers y 7 CNN. Aquí también se configuran sus parámetros: salida de dos clases, "
     "velocidad de aprendizaje y demás ajustes, iguales para todos. Este paso no recibe fotos; solo entrega cada "
     "modelo al paso 8. Código: build_model en models.py y V3_CFG en run_freshness_v3.py."),
    ("Entrenar (componente, etapa 3)",
     "aquí se juntan las fotos del paso 6 y el modelo del paso 7. El modelo aprende con las fotos de entrenamiento "
     "durante 30 épocas y, al terminar cada época, se mide su F1 macro con las fotos de validación. Se repite para "
     "cada modelo y cada fold: 50 entrenamientos en total. Código: dnc_5fold_full_metrics.py."),
    ("Mejor versión (archivo, etapa 3)",
     "cada vez que el F1 de validación del paso 8 supera su mejor valor, se guarda esa versión del modelo (checkpoint, "
     "archivo .pt) y reemplaza a la anterior. Al terminar queda una por modelo y fold, 50 en total. Esta versión, no "
     "la de la última época, es la que pasa al examen."),
    ("Examen con test (componente, etapa 4)",
     "carga la mejor versión del paso 9 y clasifica una sola vez las 1,600 fotos de test que se reservaron en el paso "
     "6 y que el modelo nunca vio; no se vuelve a entrenar. Por cada foto deja su predicción y su probabilidad. "
     "Código: dnc_5fold_confmat_roc.py."),
    ("Reunir 5 folds (función, etapa 5)",
     "cuando terminan los 5 exámenes del paso 10, junta sus respuestas: 8,000 por modelo, una por foto, porque cada "
     "foto fue test una sola vez. Con ellas calcula la matriz de confusión, el F1 macro y el AUC. Solo se usan fotos "
     "del examen, no del entrenamiento. Código: dnc_5fold_confmat_roc.py y compute_metrics en metrics.py."),
    ("Tablas (archivo, etapa 5)",
     "guarda los resultados del paso 11 en archivos CSV, por fold y del total, en 3_resultados_test_por_modelo/. Estas "
     "tablas alimentan los pasos 13 y 14."),
    ("Gráficas (archivo, etapa 5)",
     "con las tablas del paso 12 se dibujan las matrices de confusión y las curvas ROC; las curvas por época se dibujan "
     "con los registros de entrenamiento. Estas figuras pasan a la documentación (paso 15)."),
    ("Comparar 10 modelos (función, etapa 6)",
     "con las tablas del paso 12 compara los 10 modelos por F1, AUC, estabilidad entre folds y tamaño. Elige el mejor "
     "de cada familia, ViT-Base/16 y DenseNet-121, que pasan a las pruebas con otros peces. La comparación va a la "
     "documentación (paso 15)."),
    ("Documentación (archivo, etapa 7)",
     "es este documento: reúne la comparación del paso 14, las gráficas del paso 13 y los registros del paso 16."),
    ("Registros de ejecución (registro)",
     "es lo que se anota durante el entrenamiento (paso 8) y el examen (paso 10), guardado en 2_logs_entrenamiento/. "
     "Respaldan los números del documento (flechas punteadas)."),
]
GLOSARIO = [
    ("Cargador", "Nombre usado en este documento para el código que lleva las fotos hasta el modelo: anota ruta y "
                 "clase, y abre las fotos por lotes. Se apoya en Dataset y DataLoader de PyTorch "
                 "(torch.utils.data), que son los términos para buscarlo."),
    ("Componente", "Caja del diagrama que corresponde a un archivo o script de Python con varias funciones, que juntas "
                   "hacen una tarea."),
    ("Función (en el diagrama)", "Caja del diagrama que corresponde a un paso de cálculo concreto dentro de un script."),
]


def narrar(par, numero, titulo, texto):
    """Deja el párrafo como «N. Título: texto»; {1.1} se convierte en enlace a esa sección."""
    negrita = copy.deepcopy(next(r for r in par.runs if r.bold)._r)
    normal = copy.deepcopy(next(r for r in par.runs if not r.bold and r.text.strip())._r)
    for hijo in list(par._p):
        if not hijo.tag.endswith("}pPr"):
            par._p.remove(hijo)
    par._p.append(negrita)
    par.runs[-1].text = f"{numero}. {titulo}: "
    for trozo in re.split(r"(\{[\d.]+\})", texto):
        if not trozo:
            continue
        par._p.append(copy.deepcopy(normal))
        run = par.runs[-1]
        if trozo.startswith("{"):
            seccion = trozo.strip("{}")
            run.text = seccion
            par._p.replace(run._r, _hipervinculo(run._r, seccion, "sec_" + seccion.replace(".", "_")))
        else:
            run.text = trozo


def arquitectura(doc):
    cuerpo = [Paragraph(el, doc) for el in doc.element.body.iterchildren() if el.tag.endswith("}p")]
    ini = next(i for i, p in enumerate(cuerpo) if p.text.startswith(PRIMERO))
    fin = next(i for i, p in enumerate(cuerpo) if p.text.strip() == FIN)
    pares = [p for p in cuerpo[ini:fin] if p.text.strip()]
    assert len(pares) == len(ARQUITECTURA) == 16, len(pares)
    por_titulo = {p.text.split(" (")[0]: p for p in pares}
    # Gráficas pasa antes de Comparar: el orden del texto sigue la numeración del diagrama
    por_titulo["Gráficas"]._p.addnext(por_titulo["Comparar 10 modelos"]._p)
    for n, (titulo, texto) in enumerate(ARQUITECTURA, start=1):
        narrar(por_titulo[titulo.split(" (")[0]], n, titulo, texto)
    leer = cuerpo[ini - 1] if cuerpo[ini - 1].text.strip() else cuerpo[ini - 2]
    viejo = "El diagrama se lee de izquierda a derecha: primero la selección y el balanceo de las imágenes, después las siete etapas. "
    assert viejo in leer.text
    leer.runs[0].text = leer.runs[0].text.replace(viejo, INTRO)
    assert INTRO in leer.text


def glosario(doc):
    tabla = next(t for t in doc.tables if t.rows[0].cells[0].text.strip() == "Término"
                 and any(f.cells[0].text.strip() == "CSV" for f in t.rows))
    for termino, definicion in GLOSARIO:
        siguiente = next(f for f in tabla.rows[1:] if f.cells[0].text.strip().lower() > termino.lower())
        nueva = copy.deepcopy(siguiente._tr)
        siguiente._tr.addprevious(nueva)
        fila = next(f for f in tabla.rows if f._tr is nueva)
        poner(fila.cells[0], termino)
        poner(fila.cells[1], definicion)


def main():
    assert not list(DOC.parent.glob("~$cumentacion_Tecnica_10_Modelos*")), "cierra el documento en Word"
    doc = Document(DOC)
    arquitectura(doc)
    glosario(doc)
    entorno = next(t for t in doc.tables if t.cell(1, 0).text.strip() == "Hardware")
    poner(fila_de(entorno, "4 corridas de completar_30ep_*.log").cells[0],
          "Cuatro corridas (ViT-Base/16 fold 2, Swin-Base fold 4, EfficientNet-Lite0 fold 3, ResNet-50 fold 3)")
    poner(fila_de(entorno, "Deep learning").cells[0], "Librerías de deep learning")
    global_(doc, "La prueba se aplica en el documento de generalización, donde se compara ViT-Base/16 con DenseNet-121 "
                 "en Trucha y Rohu.",
            "La prueba se aplicará al final, al comparar con el modelo propuesto con adaptación de dominio.")
    global_(doc, "se aplica en el documento de generalización.",
            "se aplicará en la comparación final con el modelo propuesto con adaptación de dominio.")
    # Diagrama con los pasos numerados
    nueva = Path("6_documentacion/diagramas/arquitectura_sistema_con_seleccion.png").read_bytes()
    previa = Path(sys.argv[1]).read_text().strip()
    n = 0
    for parte in doc.part.package.iter_parts():
        if str(parte.partname).startswith("/word/media/") and hashlib.md5(parte.blob).hexdigest() == previa:
            parte._blob = nueva
            n += 1
    assert n == 1, n
    doc.save(DOC)


if __name__ == "__main__":
    main()
