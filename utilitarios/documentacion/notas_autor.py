"""Amplía las notas «Explicación del autor» de Documentacion_Tecnica_10_Modelos.docx.

Uso (desde la raíz del repositorio): python utilitarios/documentacion/notas_autor.py
"""
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

DOC = Path("6_documentacion/Documentacion_Tecnica_10_Modelos.docx")
NOTAS = {
    "1.1 Etapa 1": "El proyecto busca detectar si un pescado está fresco a partir de la foto de su ojo. En este primer "
                   "experimento, los modelos aprenden con un conjunto grande: 4,390 fotos de 8 especies, que venían en "
                   "tres estados y se agruparon en dos, fresco y no fresco. Como había más del doble de fotos frescas, "
                   "un modelo podría acertar mucho solo por favorecer a esa clase. Por eso se crearon fotos nuevas con "
                   "pequeños cambios (giros, volteos, brillo, color) hasta tener 4,000 de cada clase. Así se llega a "
                   "8,000 fotos, y las dos clases pesan lo mismo.",
    "1.2 Etapa 2": "Las 8,000 fotos se reparten en 5 lotes de 1,600, cada uno con la misma cantidad de fotos frescas y "
                   "no frescas. En cada ronda, un lote queda guardado para el examen final y, con los otros cuatro, el "
                   "modelo aprende y se va midiendo. Al cambiar de lote en cada ronda, cada foto se usa una sola vez "
                   "como examen. También se preparan los 10 modelos, que ya vienen entrenados para reconocer imágenes "
                   "en general, con la misma configuración para que la comparación sea justa.",
    "1.3 Etapa 3": "Cada modelo aprende a distinguir fresco de no fresco revisando las fotos de entrenamiento 30 veces. "
                   "Después de cada revisión se mide su resultado con un grupo de fotos aparte, que no usa para "
                   "aprender. Si ese resultado es el mejor hasta ese momento, se guarda esa versión del modelo. Al "
                   "terminar queda la mejor versión, que no siempre es la última. Como son 10 modelos y 5 rondas, son "
                   "50 entrenamientos y 50 versiones guardadas.",
    "1.4 Etapa 4": "La mejor versión de cada modelo clasifica el lote guardado en la etapa 2, con 1,600 fotos que nunca "
                   "vio mientras aprendía. Este examen se hace una sola vez y ya no se ajusta nada del modelo. Por cada "
                   "foto queda su respuesta y qué tan seguro estaba. Así se comprueba si el modelo reconoce la frescura "
                   "en fotos nuevas, y no solo en las que ya conocía.",
    "1.5 Etapa 5": "Se juntan los resultados de los 5 exámenes. Como cada foto fue examen una sola vez, hay una respuesta "
                   "del modelo para cada una de las 8,000 fotos. Con esas respuestas se cuentan los aciertos y los "
                   "errores en fresco y en no fresco, y se calcula la nota final de cada modelo. Esa nota resume todas "
                   "las fotos, no solo las de una ronda.",
    "1.6 Etapa 6": "Se comparan los 10 modelos por su nota, por qué tan estables fueron sus resultados entre los 5 lotes "
                   "y por su tamaño. Se elige el mejor de cada tipo: ViT-Base/16 (97.04), el de mejor nota entre todos, y "
                   "DenseNet-121 (95.66), la mejor CNN, con un modelo mucho más pequeño. Tener uno de cada tipo permite "
                   "comparar los dos enfoques cuando se prueben después con Trucha y Rohu.",
    "1.7 Etapa 7": "Se ordenan los resultados en tablas y gráficos, junto con los registros que se generaron durante el "
                   "aprendizaje y el examen. Así, cada número del documento se puede ubicar en su archivo de origen y "
                   "revisar. Los enlaces del documento llevan de cada resultado a su fórmula y a su detalle.",
}


def poner_nota(doc, prefijo, texto):
    dentro = False
    for el in doc.element.body.iterchildren():
        if not el.tag.endswith("}p"):
            continue
        p = Paragraph(el, doc)
        if p.style.name.startswith("Heading"):
            if p.style.name != "Heading 3":  # 1.2.1, 1.6.1… siguen dentro de su etapa
                dentro = p.text.startswith(prefijo)
        elif dentro and p.text.startswith("Explicación del autor"):
            cuerpo = [r for r in p.runs if r.text.strip() and not r.bold]
            assert cuerpo, p.text
            cuerpo[0].text = texto
            for r in cuerpo[1:]:
                r.text = ""
            return
    raise AssertionError(prefijo)


if __name__ == "__main__":
    assert not list(DOC.parent.glob("~$cumentacion_Tecnica_10_Modelos*")), "cierra el documento en Word"
    documento = Document(DOC)
    for seccion, nota in NOTAS.items():
        poner_nota(documento, seccion, nota)
    documento.save(DOC)
