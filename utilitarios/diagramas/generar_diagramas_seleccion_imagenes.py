"""Diagramas del experimento Mendeley, incluido el paso previo de selección de imágenes.

Genera en 6_documentacion/diagramas/ (SVG editable + PNG en alta resolución):
  - arquitectura_sistema_con_seleccion: arquitectura completa (selección de imágenes y 7 etapas), con el estilo común de herramientas_docx/diagramas_svg.py;
  - flujo_metodologico_con_seleccion: flujo vertical de los pasos 0 a 7;
  - vista_simple_mendeley: el experimento en 6 pasos, para público general.

Cifras: 4,390 originales (3,084 fresco y 1,306 no fresco), 3,610 aumentadas (balance_dnc_4k.py),
5 folds, 10 modelos, F1 macro pooled de ViT-Base/16 (97.04) y DenseNet-121 (95.66).
Uso (desde la raíz del repositorio): python utilitarios/diagramas/generar_diagramas_seleccion_imagenes.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT.parent / "herramientas_docx"))
import diagramas_svg as d  # noqa: E402

OUT = ROOT / "6_documentacion" / "diagramas"
NARANJA, OCRE, AZUL, MAGENTA, VERDE = "#dd861a", "#c8930b", "#5154a1", "#a42b69", "#148f87"
ANCHO, ALTO = d.X0 + 8 * d.COL, 1640


# Tipo de cada caja (ícono del diagrama); la documentación lo usa para nombrar cada componente
# Orden de lectura: el número de cada caja coincide con la explicación del documento
ORDEN_CAJAS = ["Repositorio Mendeley", "Dos clases", "Aumento de datos", "Fotos DNC", "Cargador", "Dividir en 5 folds", "Preparar modelos", "Entrenar", "Mejor versión", "Examen con test", "Reunir 5 folds", "Tablas", "Gráficas", "Comparar 10 modelos", "Documentación"]
TIPO_CAJA = {
    "Repositorio Mendeley": "dataset",
    "Dos clases": "función",
    "Aumento de datos": "componente",
    "Fotos DNC": "dataset",
    "Cargador": "componente",
    "Dividir en 5 folds": "función",
    "Preparar modelos": "componente",
    "Mejor versión": "archivo",
    "Entrenar": "componente",
    "Examen con test": "componente",
    "Reunir 5 folds": "función",
    "Tablas": "archivo",
    "Gráficas": "archivo",
    "Comparar 10 modelos": "función",
    "Documentación": "archivo",
}


def arquitectura():
    L = d.Lienzo(ANCHO, ALTO)
    caja = lambda c, f, color, titulo, lineas, **kw: L.caja(  # noqa: E731
        c, f, color, f"{ORDEN_CAJAS.index(titulo) + 1}. {titulo}", lineas, tipo=TIPO_CAJA.get(titulo), **kw)
    for c in range(8):
        L.panel(c, 1030)
    cabeceras = [(d.GRIS, "ANTES DEL EXPERIMENTO", "Seleccionar imágenes"), (NARANJA, "ETAPA 1", "Preparar datos"),
                 (OCRE, "ETAPA 2", "Preparar experimento"), (AZUL, "ETAPA 3", "Entrenar"),
                 (MAGENTA, "ETAPA 4", "Evaluar"), (MAGENTA, "ETAPA 5", "Reunir resultados"),
                 (VERDE, "ETAPA 6", "Comparar y elegir"), (VERDE, "ETAPA 7", "Documentar")]
    for c, (color, l1, l2) in enumerate(cabeceras):
        L.cabecera(c, color, l1, l2)

    # Cajas: (columna, fila). La fila sigue el camino de los datos para que ninguna flecha se cruce.
    caja(0, 0, d.GRIS, "Repositorio Mendeley", ["4,390 fotos de ojos", "8 especies", "3 estados de frescura"])
    caja(0, 1, d.GRIS, "Dos clases", ["Muy fresco + fresco", "→ fresco: 3,084", "No fresco: 1,306"])
    caja(0, 2, d.GRIS, "Aumento de datos", ["Fresco: +916 → 4,000", "No fresco: +2,694 → 4,000",
                                              "Volteo, giro, brillo,", "color y desenfoque"])
    caja(1, 0, NARANJA, "Fotos DNC", ["8,000 fotos", "4,000 fresco", "4,000 no fresco"])
    caja(1, 1, NARANJA, "Cargador", ["Anota ruta y clase", "Abre las fotos por lotes", "Ajusta tamaño y color"])
    caja(2, 0, OCRE, "Dividir en 5 folds", ["En cada fold:", "entrenamiento 68 %", "validación 12 % · test 20 %"])
    caja(2, 1, OCRE, "Preparar modelos", ["10 modelos preentrenados", "3 Transformers", "7 CNN"])
    caja(3, 0, AZUL, "Mejor versión", ["Checkpoint con mejor F1", "1 por modelo y fold", "50 en total"])
    caja(3, 1, AZUL, "Entrenar", ["30 épocas por fold", "Mide el F1 en validación", "Guarda la mejor"])
    caja(4, 0, MAGENTA, "Examen con test", ["Fotos que el modelo", "nunca vio (20 %)", "Una sola evaluación"])
    caja(5, 0, MAGENTA, "Reunir 5 folds", ["8,000 respuestas", "por modelo", "Matriz · F1 · AUC"])
    caja(5, 1, MAGENTA, "Tablas", ["Métricas por fold", "y del total (CSV)"])
    caja(5, 2, MAGENTA, "Gráficas", ["Matrices de confusión", "Curvas ROC", "Curvas por época"])
    caja(6, 0, VERDE, "Comparar 10 modelos", ["F1 y AUC", "Estabilidad entre folds", "Tamaño del modelo"])
    caja(7, 1, VERDE, "Documentación", ["Comparación, gráficas", "y registros", "en este documento"])

    cx = lambda c: d.x_caja(c) + d.ANCHO / 2  # noqa: E731
    yc = lambda f: d.FILAS_Y[f] + d.ALTO / 2  # noqa: E731
    der = lambda c: d.x_caja(c) + d.ANCHO  # noqa: E731
    izq = d.x_caja
    arriba = lambda f: d.FILAS_Y[f]  # noqa: E731
    abajo = lambda f: d.FILAS_Y[f] + d.ALTO  # noqa: E731
    G = d.GRIS

    # Antes del experimento: 4,390 fotos → 2 clases → 8,000 fotos balanceadas
    L.flecha([(cx(0), abajo(0)), (cx(0), arriba(1))], G); L.etiqueta(cx(0) + 14, d.franja(0), "agrupar", G, "start")
    L.flecha([(cx(0), abajo(1)), (cx(0), arriba(2))], G); L.etiqueta(cx(0) + 14, d.franja(1), "balancear", G, "start")
    L.flecha([(der(0), yc(2)), (d.hueco(0), yc(2)), (d.hueco(0), yc(0)), (izq(1), yc(0))], G)
    L.etiqueta(d.hueco(0) + 12, d.franja(1), "8,000 fotos", G, "start")
    # Etapa 1: el cargador lista ruta y clase de cada foto; con esa lista se arman los folds
    L.flecha([(cx(1), abajo(0)), (cx(1), arriba(1))], NARANJA); L.etiqueta(cx(1) + 14, d.franja(0), "leer", NARANJA, "start")
    L.flecha([(der(1), yc(1)), (d.hueco(1), yc(1)), (d.hueco(1), yc(0)), (izq(2), yc(0))], NARANJA)
    L.etiqueta(d.hueco(1) + 12, d.franja(0), "ruta y clase", NARANJA, "start")
    # Etapa 2 → 3: entrenamiento y validación + modelo entran a «Entrenar»; el test va aparte, por arriba
    L.flecha([(der(2), yc(0)), (d.hueco(2), yc(0)), (d.hueco(2), yc(1) - 50), (izq(3), yc(1) - 50)], OCRE)
    L.etiqueta(d.hueco(2) + 12, d.franja(0), "entrenamiento\ny validación", OCRE, "start")
    L.flecha([(der(2), yc(1) + 50), (izq(3), yc(1) + 50)], OCRE); L.etiqueta(d.hueco(2), yc(1) + 92, "modelo", OCRE)
    L.flecha([(cx(2), arriba(0)), (cx(2), d.CARRIL_SUPERIOR), (cx(4), d.CARRIL_SUPERIOR), (cx(4), arriba(0))], OCRE)
    L.etiqueta(cx(3), d.CARRIL_SUPERIOR, "test reservado del mismo fold (20 %)", OCRE)
    # Etapa 3 → 4: se guarda la mejor versión y esa es la que rinde el examen
    L.flecha([(cx(3), arriba(1)), (cx(3), abajo(0))], AZUL)
    L.etiqueta(cx(3) + 14, d.franja(0), "guarda la mejor", AZUL, "start")
    L.flecha([(der(3), yc(0)), (izq(4), yc(0))], AZUL); L.etiqueta(d.hueco(3), yc(0) - 30, "cargar", AZUL)
    # Etapa 4 → 5: respuestas → tablas → gráficas
    L.flecha([(der(4), yc(0)), (izq(5), yc(0))], MAGENTA); L.etiqueta(d.hueco(4), yc(0) - 30, "respuestas", MAGENTA)
    L.flecha([(cx(5), abajo(0)), (cx(5), arriba(1))], MAGENTA); L.etiqueta(cx(5) + 14, d.franja(0), "guardar", MAGENTA, "start")
    L.flecha([(cx(5), abajo(1)), (cx(5), arriba(2))], MAGENTA); L.etiqueta(cx(5) + 14, d.franja(1), "datos", MAGENTA, "start")
    # Etapa 5 → 6 → documentación
    L.flecha([(der(5), yc(1) - 50), (d.hueco(5), yc(1) - 50), (d.hueco(5), yc(0) + 50), (izq(6), yc(0) + 50)], MAGENTA)
    L.etiqueta(d.hueco(5) + 12, d.franja(0), "métricas", MAGENTA, "start")
    L.flecha([(cx(6), abajo(0)), (cx(6), yc(1)), (izq(7), yc(1))], VERDE)
    L.etiqueta(cx(6) + 14, d.franja(0), "comparación", VERDE, "start")
    L.flecha([(der(5), yc(2)), (cx(7) - 60, yc(2)), (cx(7) - 60, abajo(1))], MAGENTA)
    L.etiqueta(cx(6), yc(2) - 30, "figuras", MAGENTA)

    # Registros (punteado): lo que anota el entrenamiento y el examen; respaldan la documentación
    yb = 1330
    L.banda(3, 7, yb, 110, "16. Registros de ejecución", "Entrenamiento época por época y examen de cada fold: "
                                                      "respaldan cada número del documento")
    L.flecha([(cx(3), abajo(1)), (cx(3), yb)], d.LINEA, True)
    L.etiqueta(cx(3) + 14, abajo(1) + 120, "registro por época", d.LINEA, "start")
    L.flecha([(cx(4), abajo(0)), (cx(4), yb)], d.LINEA, True)
    L.etiqueta(cx(4) + 14, abajo(1) + 120, "registro del examen", d.LINEA, "start")
    L.flecha([(cx(7) + 60, yb), (cx(7) + 60, abajo(1))], d.LINEA, True)
    L.etiqueta(cx(7) + 74, abajo(1) + 290, "respaldo", d.LINEA, "start")

    L.leyenda(1470, ["Flecha continua: paso de datos de una caja a otra. Flecha punteada: registro que se guarda o se "
                     "consulta. Cada color es una etapa; la columna gris ocurre antes del experimento.",
                     "Fold: cada una de las 5 partes en que se dividen las fotos; cada modelo se entrena 5 veces y cada "
                     "foto se usa una sola vez como examen."])
    svg = L.svg("EXPERIMENTO MENDELEY · ARQUITECTURA DEL SISTEMA",
                "De las fotos originales a los dos modelos elegidos: selección de imágenes y siete etapas")
    d.guardar(svg, OUT, "arquitectura_sistema_con_seleccion")


PASOS = [
    ("0", d.GRIS, "Selección y balanceo de imágenes",
     "4,390 fotos originales; 3 estados → 2 clases; aumento de datos hasta 4,000 por clase."),
    ("1", NARANJA, "Preparación de datos", "Se registran la ruta y la clase de las 8,000 fotos."),
    ("2", OCRE, "Preparación del experimento", "Las fotos se dividen en 5 folds y se configuran los 10 modelos."),
    ("3", AZUL, "Entrenamiento", "Cada modelo aprende en cada fold y se guarda su mejor versión (50 en total)."),
    ("4", MAGENTA, "Evaluación con fotos no vistas", "La mejor versión clasifica el 20 % reservado para el examen."),
    ("5", MAGENTA, "Reunir resultados", "Se juntan los 5 folds: 8,000 respuestas por modelo, una por foto."),
    ("6", VERDE, "Comparación y selección", "Se comparan F1, AUC, estabilidad y tamaño; se eligen 2 modelos."),
    ("7", VERDE, "Documentación y trazabilidad", "Tablas, gráficas y registros respaldan cada número."),
]


def flujo():
    """Flujo metodológico vertical con el paso previo «0»."""
    ancho, paso, y0, h = 1500, 190, 105, 140
    alto = y0 + paso * len(PASOS) + 20
    s = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{ancho}" height="{alto}" viewBox="0 0 {ancho} {alto}">',
         '<rect width="100%" height="100%" fill="#ffffff"/>',
         f'<text x="750" y="72" {d.FUENTE} font-size="38" font-weight="700" fill="#17394d" text-anchor="middle">'
         'Flujo metodológico · experimento Mendeley</text>']
    for i, (num, color, titulo, desc) in enumerate(PASOS):
        y = y0 + i * paso
        s += [f'<rect x="95" y="{y}" width="1310" height="{h}" rx="18" fill="#f4f8fa" stroke="{color}" stroke-width="4"/>',
              f'<circle cx="160" cy="{y + h / 2}" r="38" fill="{color}"/>',
              f'<text x="160" y="{y + h / 2 + 12}" {d.FUENTE} font-size="34" font-weight="700" fill="#fff" '
              f'text-anchor="middle">{num}</text>',
              f'<text x="226" y="{y + 58}" {d.FUENTE} font-size="33" font-weight="700" fill="#17394d">{titulo}</text>',
              f'<text x="226" y="{y + 104}" {d.FUENTE} font-size="26" fill="#3d5566">{desc}</text>']
        if i < len(PASOS) - 1:
            ya, yb = y + h + 2, y + paso - 4
            s.append(f'<path d="M750 {ya} V{yb - 18}" stroke="#5f7a89" stroke-width="5"/>'
                     f'<path d="M734 {yb - 22} L766 {yb - 22} L750 {yb} Z" fill="#5f7a89"/>')
    s.append("</svg>")
    d.guardar("\n".join(s), OUT, "flujo_metodologico_con_seleccion")


def vista_simple():
    d.guardar(d.vista_simple(
        "El experimento Mendeley en 6 pasos",
        "Objetivo: elegir los dos modelos que después se prueban con otros peces",
        [(d.GRIS, "Fotos", "4,390", "fotos de ojos\nde pescado\n(8 especies, Indonesia)"),
         (NARANJA, "Balancear", "8,000", "4,000 frescas y\n4,000 no frescas\n(con aumento de datos)"),
         (OCRE, "Dividir", "5 folds", "cada foto se usa\nuna vez como examen"),
         (AZUL, "Entrenar", "10 modelos", "aprenden con las\nmismas fotos y reglas"),
         (MAGENTA, "Examen", "Fotos\nno vistas", "cada modelo clasifica\nfotos que nunca vio"),
         (d.DORADO, "Elegir", "2 modelos", "ViT-Base/16 (97.04 %)\nDenseNet-121 (95.66 %)\nuno de cada familia")]),
        OUT, "vista_simple_mendeley")


if __name__ == "__main__":
    arquitectura()
    flujo()
    vista_simple()
