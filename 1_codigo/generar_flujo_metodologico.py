"""Genera el borrador del Capitulo III (Flujo metodologico de la evaluacion de
los 10 modelos base) como .docx, incluida la Figura 1 (diagrama del flujo).

Todas las cifras del texto provienen de 3_resultados_test_por_modelo/
(metricas_pooled_por_modelo.csv, metricas_por_fold.csv y las matrices pooled)
y los hiperparametros de V3_CFG en run_freshness_v3.py.

Uso (desde la raiz del repo; requiere python-docx y matplotlib):
    python 1_codigo/generar_flujo_metodologico.py
Salida: 6_documentacion/Flujo_Metodologico_Capitulo_III.docx
        5_graficas_informe/flujo_metodologico.png (Figura 1 suelta)
"""
import io

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

OUT = "6_documentacion/Flujo_Metodologico_Capitulo_III.docx"
FIG_OUT = "5_graficas_informe/flujo_metodologico.png"

ETAPAS_FIGURA = [
    ("1. Preparación del conjunto de datos",
     "DNC: 8,000 imágenes balanceadas\n4,000 fresco / 4,000 no_fresco"),
    ("2. Validación cruzada estratificada 5-fold",
     "5 folds de 1,600 imágenes\npor corrida: 68% train · 12% val · 20% test"),
    ("3. Configuración de 10 modelos base",
     "CNN + Transformers · pesos preentrenados\nfine-tuning completo"),
    ("4. Entrenamiento y selección de checkpoint",
     "Máx. 30 epochs · early stopping = 7\ncriterio: mejor val_F1 macro por fold"),
    ("5. Evaluación sobre test independiente",
     "Mejor checkpoint → test del fold\nmatriz de confusión + probabilidades + métricas"),
    ("6. Consolidación OOF / pooled",
     "Concatenación de los 5 test\n8,000 predicciones out-of-fold, una por imagen"),
    ("7. Comparación multicriterio",
     "F1 macro · Accuracy · AUC · estabilidad\nparámetros/eficiencia · arquitectura"),
    ("8. Selección de modelos de referencia",
     "ViT-Base/16 + DenseNet-121\nsalida para las fases posteriores"),
]


def figura_flujo() -> io.BytesIO:
    n = len(ETAPAS_FIGURA)
    box_h, gap, top = 1.0, 0.45, 1.0
    total_h = top + n * box_h + (n - 1) * gap + 0.9
    fig, ax = plt.subplots(figsize=(8.5, total_h * 0.95))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, total_h)
    ax.axis("off")
    ax.text(5, total_h - 0.4, "Flujo metodológico — Evaluación y selección de modelos base",
            ha="center", va="center", fontsize=14, fontweight="bold")
    for i, (titulo, detalle) in enumerate(ETAPAS_FIGURA):
        y = total_h - top - i * (box_h + gap) - box_h
        ax.add_patch(FancyBboxPatch((0.6, y), 8.8, box_h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                    linewidth=1.5, edgecolor="#1F4E79", facecolor="#EAF1F8"))
        ax.text(0.9, y + box_h - 0.28, titulo, fontsize=11.5, fontweight="bold", va="center")
        ax.text(0.9, y + 0.33, detalle, fontsize=9.5, va="center", linespacing=1.3)
        if i < n - 1:
            ax.annotate("", xy=(5, y - gap + 0.04), xytext=(5, y - 0.04),
                        arrowprops=dict(arrowstyle="-|>", color="#444444", lw=1.5))
    ax.text(5, 0.35, "Resultado del flujo: dos modelos de referencia seleccionados con evidencia "
            "cuantitativa y trazable.", ha="center", va="center", fontsize=9.5)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    fig.savefig(FIG_OUT, dpi=200, bbox_inches="tight")  # misma figura, usada en Documentacion_Tecnica
    plt.close(fig)
    buf.seek(0)
    return buf


# --- helpers de formato -------------------------------------------------------

def parrafo(doc, texto, bold_prefix=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    if bold_prefix:
        p.add_run(bold_prefix).bold = True
    p.add_run(texto)
    return p


def vinetas(doc, items):
    for etiqueta, texto in items:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run(f"{etiqueta}: ").bold = True
        p.add_run(texto)


def alineacion(doc, texto):
    parrafo(doc, texto, bold_prefix="Alineación dentro del flujo. ")


# --- contenido ----------------------------------------------------------------

def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)  # A4
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(2.5))
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("FLUJO METODOLÓGICO")
    r.bold, r.font.size = True, Pt(16)
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("Capítulo III: Evaluación de Modelos Base")
    r.bold, r.font.size = True, Pt(13)

    # 1. Proposito
    doc.add_heading("1. Propósito del flujo metodológico", level=1)
    parrafo(doc,
        "El flujo metodológico establece la secuencia de actividades para evaluar y comparar los diez "
        "modelos base y seleccionar aquellos que serán utilizados como modelos de referencia en las "
        "fases posteriores de la investigación. Para garantizar una comparación consistente, todos los "
        "modelos son sometidos a un procedimiento común que comprende la preparación y partición de los "
        "datos, el entrenamiento, la validación, la evaluación sobre datos de prueba y la consolidación "
        "de los resultados.")
    parrafo(doc,
        "La evaluación se realiza mediante validación cruzada estratificada de cinco folds, permitiendo "
        "analizar el desempeño de cada arquitectura sobre diferentes particiones del conjunto de datos. "
        "En cada fold, el entrenamiento se realiza de manera independiente, utilizando el conjunto de "
        "validación para seleccionar el mejor checkpoint y reservando el conjunto de prueba "
        "exclusivamente para la evaluación final del modelo correspondiente.")
    parrafo(doc,
        "Posteriormente, las predicciones obtenidas en los cinco conjuntos de prueba se integran como "
        "predicciones out-of-fold (OOF), de modo que cada imagen dispone de una predicción generada por "
        "un modelo que no la utilizó durante su entrenamiento. Sobre estas predicciones se calculan "
        "métricas globales (pooled) como resumen complementario del desempeño. En paralelo, los "
        "resultados obtenidos en cada fold permiten caracterizar la variabilidad y estabilidad del "
        "modelo entre particiones.")
    parrafo(doc,
        "Finalmente, la evidencia obtenida permite seleccionar los modelos de referencia que continuarán "
        "en las siguientes fases experimentales. De esta manera, el flujo asegura que la selección no "
        "dependa de un único resultado o de una sola partición de los datos, sino de un proceso "
        "sistemático, documentado y sustentado en criterios cuantitativos de evaluación.")

    # 2. Figura
    doc.add_heading("2. Flujo metodológico", level=1)
    doc.add_picture(figura_flujo(), width=Cm(14.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.add_run("Figura 1. ").bold = True
    r = cap.add_run("Flujo metodológico para la evaluación y selección de los modelos base. "
                    "Fuente: elaboración propia.")
    r.italic = True
    parrafo(doc,
        "La lectura del flujo es secuencial. Cada etapa recibe la salida de la etapa anterior y agrega "
        "un control metodológico específico. Primero se fija el conjunto de datos; después se "
        "construyen particiones comparables; luego se someten los diez modelos al mismo principio de "
        "entrenamiento y selección del checkpoint; posteriormente se evalúan sobre datos de prueba "
        "independientes y se consolidan los resultados de los cinco folds; finalmente, esos resultados "
        "se comparan con varios criterios y se seleccionan los modelos de referencia.")

    # 3. Explicacion detallada
    doc.add_heading("3. Explicación detallada del flujo metodológico", level=1)

    # Etapa 1
    doc.add_heading("Etapa 1. Preparación y delimitación del conjunto de datos de evaluación", level=2)
    parrafo(doc,
        "El punto de partida de la evaluación es el dataset DNC (DataSetNotClean), compuesto por 8,000 "
        "imágenes. El conjunto está balanceado: 4,000 imágenes pertenecen a la clase «fresco» y 4,000 a "
        "la clase «no_fresco». Este balance es relevante porque establece una base equivalente para "
        "ambas clases y permite que la comparación de los modelos no quede dominada, en este "
        "experimento, por una clase mayoritaria.")
    parrafo(doc,
        "En esta etapa todavía no se entrena ningún modelo. Su función es fijar el universo experimental "
        "que será recorrido por la validación cruzada. Las 8,000 imágenes constituyen el conjunto sobre "
        "el cual se generarán las particiones de entrenamiento, validación y prueba. Asimismo, se deja "
        "establecido que las métricas finales se expresan en un problema binario y que las métricas por "
        "clase se calculan tratando cada clase como positiva por turno; posteriormente se obtiene el "
        "promedio macro.")
    parrafo(doc,
        "La métrica F1 se interpreta como F1 macro. Para cada clase, el F1 corresponde a la media "
        "armónica entre Precisión y Recall; posteriormente, el F1 macro se obtiene como el promedio no "
        "ponderado del F1 de «fresco» y del F1 de «no_fresco». De esta forma, ambas clases tienen el "
        "mismo peso en el resultado. El F1 macro se utiliza durante la validación para seleccionar el "
        "checkpoint y también como una de las métricas principales para comparar los modelos.")
    vinetas(doc, [
        ("Entrada", "8,000 imágenes del dataset DNC."),
        ("Clases", "4,000 fresco y 4,000 no_fresco."),
        ("Unidad que continúa al flujo", "Cada imagen individual con su etiqueta de clase."),
        ("Métrica central", "F1 macro, complementada con Accuracy, Precisión, Recall, Especificidad y "
         "AUC; adicionalmente, la variabilidad entre folds se caracteriza mediante la desviación "
         "estándar del F1."),
    ])
    alineacion(doc,
        "Esta etapa define exactamente qué datos entran al experimento y bajo qué formulación de "
        "clasificación. Sin esta delimitación, las métricas de los modelos no serían directamente "
        "comparables.")

    # Etapa 2
    doc.add_heading("Etapa 2. Construcción de la validación cruzada estratificada de cinco folds", level=2)
    parrafo(doc,
        "Una vez definido el conjunto de 8,000 imágenes, se aplica StratifiedKFold con n_splits=5, "
        "shuffle=True y random_state=42. La función de la estratificación es conservar la proporción de "
        "las dos clases dentro de las particiones. Al utilizar cinco folds, el conjunto completo se "
        "divide en cinco bloques de 20%, equivalentes a 1,600 imágenes por fold, sin solape entre esos "
        "bloques.")
    parrafo(doc,
        "La partición se realizó a nivel de imagen individual, sin agrupar las imágenes derivadas de una "
        "misma captura o ejemplar. En consecuencia, si el conjunto contiene varias imágenes de una misma "
        "captura, estas pueden quedar distribuidas entre entrenamiento, validación y prueba, por lo que "
        "no se descarta cierta correlación entre conjuntos que podría elevar las métricas de prueba. "
        "Esta condición se reconoce como una limitación del estudio.")
    parrafo(doc,
        "El procedimiento se repite cinco veces. En cada corrida, uno de los cinco folds se reserva como "
        "conjunto de prueba. Ese test representa el 20% de las 8,000 imágenes y permanece fuera del "
        "entrenamiento y de la validación de esa corrida. Sobre el 80% restante se aplica una separación "
        "adicional mediante train_test_split (test_size=0.15) con estratificación y la misma semilla. De "
        "esta forma, la distribución efectiva es 68% para entrenamiento (5,440 imágenes), 12% para "
        "validación (960) y 20% para prueba (1,600).")
    parrafo(doc,
        "La rotación es esencial para entender la metodología. En la primera corrida un fold actúa como "
        "test; en las siguientes corridas esa función pasa a los demás folds. Por tanto, al terminar las "
        "cinco corridas, cada imagen habrá formado parte del conjunto de prueba exactamente una vez. "
        "Dentro de una misma corrida, los conjuntos de entrenamiento, validación y prueba mantienen "
        "funciones diferenciadas y no se mezclan entre sí.")
    parrafo(doc,
        "El conjunto de entrenamiento sirve para actualizar los pesos del modelo. El conjunto de "
        "validación no actualiza los pesos: se utiliza para observar el comportamiento del modelo "
        "durante los epochs y decidir cuál checkpoint es el mejor. El conjunto de prueba queda reservado "
        "hasta que termina el entrenamiento de ese fold y se utiliza solo para medir el desempeño del "
        "checkpoint seleccionado.")
    vinetas(doc, [
        ("Train = 68%", "Se utiliza para aprender y actualizar los pesos."),
        ("Validation = 12%", "Se utiliza para calcular val_F1 y seleccionar el mejor checkpoint."),
        ("Test = 20%", "Permanece sin uso durante el ajuste y se emplea solo al final del fold."),
        ("Número de rotaciones", "Cinco; cada fold ocupa una vez el papel de test."),
    ])
    alineacion(doc,
        "La salida de esta etapa no es una sola división train/test, sino cinco escenarios de evaluación "
        "equivalentes. Esto permite observar tanto el desempeño global como la variabilidad del modelo "
        "frente a diferentes particiones.")

    # Etapa 3
    doc.add_heading("Etapa 3. Configuración homogénea de los diez modelos candidatos", level=2)
    parrafo(doc,
        "Con las particiones ya definidas, el flujo incorpora los diez modelos candidatos que serán "
        "sometidos al mismo principio experimental. Los modelos evaluados son: ViT-Base/16, "
        "DenseNet-121, Swin-Base, DINOv2-Small, EfficientNet-B0, EfficientNet-Lite0, MobileNetV3-Small, "
        "ResNet-50, MobileNetV2 y MobileNetV1. El conjunto incluye arquitecturas Transformer y "
        "arquitecturas convolucionales, lo que posteriormente permitirá realizar una selección que no "
        "dependa de una sola familia.")
    parrafo(doc,
        "Todos los modelos parten de pesos preentrenados (ImageNet; en DINOv2-Small, pesos "
        "auto-supervisados de DINOv2) y se someten a fine-tuning completo. Esto significa que el "
        "entrenamiento no se limita a reemplazar y ajustar únicamente la capa de clasificación final; "
        "los pesos de toda la arquitectura se actualizan durante el entrenamiento para adaptarse a la "
        "tarea de clasificación de frescura. En los modelos Transformer, el backbone se actualiza con una "
        "tasa de aprendizaje diez veces menor que la de la cabeza de clasificación.")
    parrafo(doc,
        "Se consideran además diferencias operativas asociadas a las familias arquitectónicas. Para las "
        "CNN se utilizó un tamaño de imagen de 256×256 y batch size de 16, mientras que los modelos "
        "Transformer (ViT-Base/16, Swin-Base y DINOv2-Small) utilizaron 224×224 y batch size de 8 "
        "durante el entrenamiento. Estas diferencias forman parte de la configuración real de ejecución "
        "y deben entenderse como parámetros operativos de las arquitecturas, no como cambios en el "
        "criterio con el que se decide qué modelo es mejor.")
    parrafo(doc,
        "El resto de la configuración de entrenamiento es común a los diez modelos: optimizador AdamW "
        "con tasa de aprendizaje de 1e-4 y weight decay de 0.05; tres epochs de calentamiento (warmup) "
        "lineal seguidos de un scheduler coseno; función de pérdida Focal Loss (γ = 2) con label "
        "smoothing de 0.1; y gradient clipping con norma máxima de 1.0. El data augmentation (recorte "
        "aleatorio, volteos, variación de color, desenfoque, rotación de ±15° y borrado aleatorio) se "
        "aplica solo al conjunto de entrenamiento; validación y prueba usan únicamente redimensionado y "
        "normalización.")
    parrafo(doc,
        "La implementación se realizó con Python y PyTorch/torchvision para entrenamiento y "
        "transformaciones; timm como repositorio de arquitecturas; scikit-learn para StratifiedKFold, "
        "train_test_split, matrices de confusión, ROC y AUC; pandas/numpy para tratamiento de "
        "resultados; y matplotlib para las visualizaciones. El entorno de ejecución fue una MacBook con "
        "chip Apple M2 y aceleración mediante Metal Performance Shaders (MPS). Las semillas se fijaron "
        "en 42; no obstante, algunas operaciones en MPS no son deterministas, por lo que un "
        "reentrenamiento puede producir diferencias del orden de décimas en las métricas.")
    vinetas(doc, [
        ("Modelos basados en arquitecturas Transformer", "ViT-Base/16, Swin-Base y DINOv2-Small."),
        ("CNN evaluadas", "DenseNet-121, EfficientNet-B0, EfficientNet-Lite0, MobileNetV3-Small, "
         "ResNet-50, MobileNetV2 y MobileNetV1."),
        ("Inicialización", "Pesos preentrenados."),
        ("Ajuste", "Fine-tuning completo."),
        ("Entrenamiento", "AdamW (lr 1e-4, weight decay 0.05), warmup + coseno, Focal Loss (γ = 2) con "
         "label smoothing 0.1."),
    ])
    alineacion(doc,
        "Esta etapa asegura que la comparación posterior tenga diversidad arquitectónica. El flujo no "
        "parte de un modelo previamente elegido: primero genera evidencia para diez candidatos y recién "
        "después aplica criterios de selección.")

    # Etapa 4
    doc.add_heading("Etapa 4. Entrenamiento por fold y selección del mejor checkpoint", level=2)
    parrafo(doc,
        "Esta es la etapa central de generación de modelos entrenados. Cada uno de los diez modelos se "
        "entrena en cada uno de los cinco folds, lo que produce 50 combinaciones modelo–fold. Para cada "
        "combinación se permite un máximo de 30 epochs. Sin embargo, el objetivo no es obligar a que "
        "todos los entrenamientos lleguen al epoch 30, sino conservar el estado del modelo que muestre "
        "el mejor comportamiento sobre el conjunto de validación.")
    parrafo(doc,
        "Durante cada epoch se calcula val_F1. Cuando el valor obtenido supera el mejor val_F1 observado "
        "previamente dentro de ese mismo fold, el checkpoint se sobrescribe y pasa a ser el nuevo «mejor "
        "modelo» de esa corrida. Si en los epochs siguientes el desempeño empeora, esos pesos no "
        "sustituyen al checkpoint guardado. Por ello, el modelo que finalmente se evalúa no es "
        "necesariamente el correspondiente al último epoch ejecutado.")
    parrafo(doc,
        "Se incorpora early stopping con una paciencia de siete epochs sin mejora. Operativamente, esto "
        "significa que si el entrenamiento acumula siete epochs sin superar el mejor val_F1 registrado, "
        "se detiene antes de alcanzar los 30 epochs. Si aparece una nueva mejora, el contador de "
        "paciencia vuelve a iniciar. Este mecanismo reduce entrenamiento innecesario y, sobre todo, "
        "mantiene la selección del modelo ligada al desempeño de validación.")
    parrafo(doc,
        "Es importante distinguir tres conceptos. Los pesos se aprenden con train; el checkpoint se "
        "selecciona con validation; y el desempeño final se mide con test. La separación de estas "
        "funciones evita utilizar el test para tomar decisiones durante el entrenamiento. Una vez "
        "finalizada la corrida, se recarga expresamente el mejor checkpoint almacenado.")
    vinetas(doc, [
        ("Cantidad máxima de entrenamientos", "10 modelos × 5 folds = 50 combinaciones modelo–fold."),
        ("Máximo de epochs", "30 por combinación."),
        ("Early stopping", "Paciencia de 7 epochs sin mejora."),
        ("Criterio para guardar checkpoint", "Mejor val_F1 macro observado en el fold."),
        ("Modelo que pasa a test", "El mejor checkpoint, no necesariamente el último epoch."),
    ])
    alineacion(doc,
        "La salida de esta etapa son cinco checkpoints seleccionados por modelo, uno por cada fold. "
        "Estos checkpoints son los únicos que pasan a la evaluación independiente de la etapa siguiente.")

    # Etapa 5
    doc.add_heading("Etapa 5. Evaluación del mejor checkpoint sobre el test independiente de cada fold",
                    level=2)
    parrafo(doc,
        "Al concluir el entrenamiento de un fold se recarga el checkpoint que obtuvo el mejor val_F1. "
        "Recién en ese momento se utiliza el 20% de test correspondiente a la corrida. Ese conjunto de "
        "prueba no se utiliza ni para actualizar los pesos ni para seleccionar el checkpoint. Las "
        "métricas reportadas se obtuvieron recargando esos mismos checkpoints y repitiendo únicamente la "
        "inferencia sobre el mismo test, sin reentrenar ni tomar decisiones con él.")
    parrafo(doc,
        "Para cada fold se generan predicciones de clase y probabilidades. A partir de las predicciones "
        "se construye la matriz de confusión y se identifican verdaderos positivos, falsos positivos, "
        "falsos negativos y verdaderos negativos. Sobre estos valores se calculan Accuracy, Precisión, "
        "Recall, F1 y Especificidad. Las probabilidades de salida (probabilidad softmax de la clase "
        "«no_fresco») se utilizan adicionalmente para construir la curva ROC y calcular el AUC.")
    parrafo(doc,
        "Las métricas por clase se interpretan mediante un esquema one-vs-rest: primero se considera "
        "«fresco» como clase positiva y después «no_fresco». Las métricas macro promedian ambas clases, "
        "de modo que las dos tienen el mismo peso en el resultado reportado. El AUC agrega otra "
        "dimensión: mientras Accuracy o F1 reflejan el comportamiento bajo la decisión de clasificación "
        "aplicada, el AUC evalúa la capacidad de discriminación a través de distintos umbrales.")
    parrafo(doc,
        "Al repetirse este procedimiento en los cinco folds, cada modelo obtiene cinco evaluaciones "
        "sobre conjuntos de prueba disjuntos. Estas evaluaciones permiten observar dos aspectos "
        "distintos: el desempeño alcanzado por el modelo y la variabilidad de dicho desempeño cuando "
        "cambia la partición de los datos.")
    vinetas(doc, [
        ("Productos por fold", "Predicciones, matriz de confusión, métricas y curva ROC."),
        ("Uso del test", "Solo para medir el checkpoint ya seleccionado; nunca para tomar decisiones."),
        ("Finalidad", "Obtener evidencia independiente de desempeño para cada fold."),
    ])
    alineacion(doc,
        "Esta etapa convierte los checkpoints seleccionados en resultados de prueba por fold. Estos "
        "resultados permiten analizar el desempeño y su variabilidad entre particiones y, "
        "adicionalmente, construir una evaluación global mediante la consolidación de las predicciones "
        "out-of-fold.")

    # Etapa 6
    doc.add_heading("Etapa 6. Consolidación de predicciones out-of-fold (OOF) y métricas globales pooled",
                    level=2)
    parrafo(doc,
        "Después de evaluar los cinco checkpoints de un modelo, las predicciones obtenidas en los cinco "
        "conjuntos de test se concatenan. Debido a que los folds de prueba son disjuntos y, en conjunto, "
        "cubren las 8,000 imágenes, se obtienen 8,000 predicciones out-of-fold (OOF): cada imagen "
        "aparece exactamente una vez como observación de prueba y su predicción procede de un modelo "
        "entrenado sin utilizar esa imagen.")
    parrafo(doc,
        "A partir de las predicciones OOF concatenadas se calculan métricas globales (pooled). Estas "
        "métricas constituyen un resumen complementario del comportamiento del modelo sobre el conjunto "
        "completo de predicciones fuera de muestra. No deben interpretarse como el promedio de las "
        "métricas de los cinco folds: métricas no lineales como F1 y AUC se recalculan directamente "
        "sobre las predicciones o puntuaciones concatenadas y, por tanto, pueden diferir del promedio "
        "aritmético de sus valores por fold. En el caso de la matriz de confusión, la matriz pooled "
        "equivale a la suma celda a celda de las cinco matrices de confusión.")
    parrafo(doc,
        "Sobre las predicciones OOF se calculan las métricas globales conservando los valores obtenidos "
        "experimentalmente. Para ViT-Base/16, el resultado pooled es F1 macro de 96.96%, Accuracy de "
        "96.96% y AUC de 0.9900. Para DenseNet-121, el resultado pooled es F1 macro de 95.66%, Accuracy "
        "de 95.66% y AUC de 0.9906. Estos valores proporcionan una visión global complementaria del "
        "desempeño de cada arquitectura.")
    parrafo(doc,
        "Paralelamente, se conserva el F1 obtenido en cada uno de los cinco folds y se calcula su "
        "desviación estándar muestral. Esta información caracteriza la estabilidad del modelo frente a "
        "cambios en la partición. Por ello, la interpretación combina dos perspectivas: las métricas "
        "globales OOF/pooled resumen el comportamiento agregado de las predicciones fuera de muestra, "
        "mientras que la variabilidad entre folds informa sobre la consistencia del desempeño. La "
        "evaluación por folds se resume mediante el F1 macro medio de las cinco particiones acompañado "
        "de su desviación estándar (media ± DE), mientras que el F1 OOF/pooled se mantiene como una "
        "medida global complementaria calculada sobre las 8,000 predicciones fuera de muestra.")
    vinetas(doc, [
        ("Resultado OOF/pooled", "Métricas recalculadas sobre las 8,000 predicciones out-of-fold "
         "concatenadas."),
        ("Matriz de confusión pooled", "Suma de las cinco matrices de confusión de test."),
        ("Información por fold", "F1 de cada partición, resumido como media ± desviación estándar."),
    ])
    alineacion(doc,
        "Esta etapa genera dos niveles de evidencia complementarios: resultados por fold para analizar "
        "la variabilidad del desempeño y resultados OOF/pooled para resumir globalmente las predicciones "
        "fuera de muestra. Ambos niveles, junto con la complejidad del modelo, alimentan la comparación "
        "multicriterio.")

    # Etapa 7
    doc.add_heading("Etapa 7. Comparación multicriterio de los diez modelos", level=2)
    parrafo(doc,
        "Con los resultados consolidados, el flujo deja de ser puramente operativo y entra en la fase de "
        "decisión. La selección de los modelos no se realiza mediante una única métrica, sino mediante "
        "cinco criterios complementarios: desempeño, discriminación, estabilidad, eficiencia y "
        "diversidad arquitectónica.")
    parrafo(doc,
        "El primer criterio es el desempeño predictivo. Se considera el F1 macro obtenido sobre las "
        "predicciones OOF/pooled, complementado por el comportamiento observado entre los cinco folds. "
        "Bajo el resultado pooled, ViT-Base/16 ocupa el primer lugar con 96.96% y DenseNet-121 el "
        "segundo con 95.66%. Swin-Base obtiene 95.56% y DINOv2-Small 95.44%.")
    parrafo(doc,
        "El segundo criterio es la capacidad discriminativa medida mediante AUC. DenseNet-121 obtiene el "
        "mayor AUC pooled de los diez modelos, con 0.9906, y ViT-Base/16 ocupa el segundo lugar con "
        "0.9900. El AUC se calcula a partir de las probabilidades de salida, no de las etiquetas de "
        "clase predichas, y resume la capacidad de discriminación a través de distintos umbrales. Al "
        "tratarse de un AUC calculado sobre predicciones OOF procedentes de cinco modelos entrenados en "
        "folds diferentes, se interpreta como una medida global complementaria al análisis por fold.")
    parrafo(doc,
        "El tercer criterio es la estabilidad entre folds. Se utiliza la desviación estándar del F1 de "
        "las cinco evaluaciones. ViT-Base/16 presenta una desviación estándar de 0.28, lo que evidencia "
        "una alta estabilidad entre los candidatos de mayor desempeño. DenseNet-121 registra 0.79; "
        "aunque no es la desviación más baja de todo el conjunto, mantiene un nivel de desempeño alto y "
        "una estabilidad comparable con otros modelos líderes.")
    parrafo(doc,
        "El cuarto criterio es la eficiencia, observada mediante la relación entre desempeño y número de "
        "parámetros. ViT-Base/16 tiene aproximadamente 85.8 millones de parámetros. DenseNet-121 alcanza "
        "un F1 macro de 95.66% con aproximadamente 7.0 millones. DenseNet-121 obtiene más del 98% del "
        "desempeño de ViT-Base/16 utilizando alrededor del 8% de sus parámetros. Por eso DenseNet no se "
        "selecciona simplemente por ser «más pequeño», sino porque combina tamaño reducido con "
        "desempeño y AUC muy altos.")
    parrafo(doc,
        "El quinto criterio es la diversidad arquitectónica. La selección busca conservar un "
        "representante Transformer y un representante CNN. Esto evita que las siguientes fases se apoyen "
        "en dos modelos muy parecidos desde el punto de vista de paradigma. Así, aunque Swin-Base tiene "
        "un F1 muy cercano a DenseNet-121, pertenece a la misma gran familia Transformer que ViT-Base/16 "
        "y no añade la comparación ViT frente a CNN que sí aporta DenseNet-121.")
    parrafo(doc,
        "La comparación de los otros modelos se interpreta con la misma lógica. EfficientNet-B0 presenta "
        "buen AUC y eficiencia, pero su F1 es inferior a DenseNet-121. MobileNetV3-Small es muy liviano, "
        "pero pierde desempeño. ResNet-50 posee más parámetros que DenseNet-121 y obtiene un F1 "
        "claramente menor. MobileNetV2 presenta una desviación estándar muy baja, pero su F1 de 89.01% "
        "no compensa la pérdida de desempeño. Por tanto, ningún criterio aislado determina la selección.")
    vinetas(doc, [
        ("Desempeño", "F1 macro OOF/pooled y Accuracy, complementados con el comportamiento del F1 entre "
         "folds."),
        ("Discriminación", "AUC OOF/pooled calculado a partir de las probabilidades de salida."),
        ("Estabilidad", "Desviación estándar del F1 entre los cinco folds."),
        ("Eficiencia", "Relación entre desempeño y cantidad de parámetros."),
        ("Diversidad", "Representación de las familias Transformer y CNN."),
    ])
    alineacion(doc,
        "Esta etapa transforma resultados experimentales en criterios explícitos de decisión. La "
        "selección final puede defenderse porque cada modelo elegido cumple un papel diferente y "
        "sustentado.")

    # Etapa 8
    doc.add_heading("Etapa 8. Selección de ViT-Base/16 y DenseNet-121 como modelos de referencia", level=2)
    parrafo(doc,
        "La última etapa formaliza la decisión producida por todo el flujo. Se seleccionan dos de los "
        "diez modelos: ViT-Base/16 como representante de las arquitecturas Transformer y DenseNet-121 "
        "como representante de las arquitecturas convolucionales. La elección no se realiza antes del "
        "experimento; es el resultado de la evidencia acumulada en las etapas anteriores.")
    parrafo(doc,
        "ViT-Base/16 queda seleccionado porque obtiene el mayor F1 macro OOF/pooled de los diez modelos, "
        "96.96%, con Accuracy de 96.96% y AUC pooled de 0.9900. Además, presenta una desviación estándar "
        "del F1 de 0.28 entre folds, lo que evidencia baja variabilidad entre particiones dentro del "
        "grupo de modelos de alto desempeño. Su matriz de confusión pooled produce F1 de 96.95% para "
        "fresco y 96.98% para no_fresco, cuyo promedio corresponde, salvo efectos de redondeo, al F1 "
        "macro reportado.")
    parrafo(doc,
        "DenseNet-121 queda seleccionado porque obtiene el segundo mayor F1 macro OOF/pooled, 95.66%, "
        "Accuracy de 95.66% y el mayor AUC pooled de todo el grupo, 0.9906. Su matriz de confusión "
        "pooled produce F1 de 95.62% para fresco y 95.70% para no_fresco, cuyo promedio es 95.66%. A "
        "esto se suma su menor número de parámetros, aproximadamente 7.0 millones, frente a 85.8 "
        "millones de ViT-Base/16.")
    parrafo(doc,
        "La combinación final tiene, por tanto, una lógica complementaria. ViT-Base/16 representa el "
        "mayor desempeño y una baja variabilidad entre particiones dentro del grupo de modelos de alto "
        "desempeño; DenseNet-121 representa una CNN con desempeño muy cercano, el mayor AUC OOF/pooled "
        "entre los modelos evaluados y una relación desempeño–parámetros favorable. Incorporar un segundo "
        "modelo basado en arquitectura Transformer, como Swin-Base, no aportaría la misma diversidad de "
        "paradigma que la combinación ViT-CNN.")
    parrafo(doc,
        "La salida metodológica del flujo no debe interpretarse como que uno de los dos modelos es el "
        "«modelo propuesto» de la tesis. Ambos constituyen modelos de referencia seleccionados a partir "
        "de la evaluación base. Su función es proporcionar una línea experimental sustentada para las "
        "fases siguientes del estudio.")
    vinetas(doc, [
        ("ViT-Base/16", "F1 macro OOF/pooled 96.96%; Accuracy OOF/pooled 96.96%; AUC OOF/pooled 0.9900; "
         "desviación estándar del F1 entre folds 0.28; 85.8M parámetros."),
        ("DenseNet-121", "F1 macro OOF/pooled 95.66%; Accuracy OOF/pooled 95.66%; AUC OOF/pooled 0.9906; "
         "desviación estándar del F1 entre folds 0.79; ~7.0M parámetros."),
        ("Resultado del flujo", "Dos modelos de referencia de familias diferentes, seleccionados con "
         "evidencia cuantitativa y trazable."),
    ])
    alineacion(doc,
        "La etapa cierra el proceso de evaluación base: las 8,000 imágenes originan cinco particiones; "
        "las particiones alimentan 50 entrenamientos modelo–fold; cada entrenamiento genera un mejor "
        "checkpoint; los checkpoints producen resultados de test; los resultados se consolidan y "
        "comparan; y esa comparación sustenta la selección final de ViT-Base/16 y DenseNet-121.")

    doc.save(OUT)
    print(f"Guardado en {OUT}")


if __name__ == "__main__":
    main()
