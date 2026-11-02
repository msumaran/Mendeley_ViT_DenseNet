"""Genera la version 2 de los diagramas del experimento Mendeley DNC (7 etapas):
arquitectura completa, recortes de detalle (etapas 1-3 y 4-7) y flujo
metodologico en forma de tabla. Escribe SVG y PNG con sufijo _v2 en
6_documentacion/diagramas/version_anterior/ sin tocar las versiones anteriores.

Uso: python generar_diagramas_7_etapas_v2.py   (requiere rsvg-convert)
"""
import subprocess
import textwrap
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parents[3] / "6_documentacion" / "diagramas" / "version_anterior"
RSVG = "/opt/homebrew/bin/rsvg-convert"
FONT = "Arial,Helvetica,sans-serif"
MONO = "Menlo,Consolas,monospace"
INK, MUTED, BG = "#243f51", "#597482", "#f9fbfc"
C = ["#dd861a", "#c8930b", "#5154a1", "#a42b69", "#a42b69", "#148f87", "#148f87"]
GOLD = "#e0a800"

# Rejilla: 8 carriles (la etapa 2 ocupa dos), 3 filas de cajas
LX = [54 + 466 * i for i in range(8)]
LW, BW = 440, 420
RA, HA = 420, 300
RB, HB = 810, 290
RC, HC = 1200, 280


def bx(col):
    return LX[col] + 10


def cx(col):
    return LX[col] + LW // 2


# ---------------------------------------------------------------- primitivas
def text(o, x, y, t, size=22, fill=INK, weight="normal", anchor="start", family=FONT):
    o.append(f'<text x="{x}" y="{y}" font-family="{family}" font-size="{size}" fill="{fill}" '
             f'font-weight="{weight}" text-anchor="{anchor}">{escape(t)}</text>')


def rect(o, x, y, w, h, fill, stroke="none", sw=2, rx=10, extra=""):
    o.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" '
             f'stroke="{stroke}" stroke-width="{sw}" {extra}/>')


def mid(col):
    return col.lstrip("#")


def arrow(o, d, col, dashed=False, w=4):
    dash = ' stroke-dasharray="14 9"' if dashed else ""
    o.append(f'<path d="{d}" fill="none" stroke="{col}" stroke-width="{w}" stroke-linejoin="round"'
             f'{dash} marker-end="url(#m{mid(col)})"/>')


def markers(colors):
    return "".join(
        f'<marker id="m{mid(c)}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" '
        f'markerHeight="7" orient="auto"><path d="M0 0 L10 5 L0 10Z" fill="{c}"/></marker>'
        for c in sorted(set(colors)))


# ---------------------------------------------------------------- iconos (lienzo -20..20)
ICONS = {
    "fish": '<path d="M-17 0 Q-4 -15 10 0 Q-4 15 -17 0Z"/><path d="M10 0 L19 -8 V8Z"/>'
            '<circle cx="-8" cy="-2" r="1.8" fill="{c}"/>',
    "list": '<circle cx="-13" cy="-10" r="2.2" fill="{c}"/><circle cx="-13" cy="0" r="2.2" fill="{c}"/>'
            '<circle cx="-13" cy="10" r="2.2" fill="{c}"/><path d="M-6 -10 H16 M-6 0 H16 M-6 10 H10"/>',
    "folds": '<rect x="-19" y="-14" width="6" height="28"/><rect x="-11" y="-14" width="6" height="28"/>'
             '<rect x="-3" y="-14" width="6" height="28" fill="{c}"/><rect x="5" y="-14" width="6" height="28"/>'
             '<rect x="13" y="-14" width="6" height="28"/>',
    "split": '<rect x="-19" y="-8" width="38" height="16"/><path d="M7 -8 V8 M12 -8 V8"/>'
             '<path d="M-19 14 H7" stroke-dasharray="3 3"/>',
    "net": '<path d="M-14 -9 L0 -14 M-14 -9 L0 0 M-14 -9 L0 14 M-14 9 L0 -14 M-14 9 L0 0 M-14 9 L0 14 '
           'M0 -14 L14 0 M0 0 L14 0 M0 14 L14 0"/>'
           '<circle cx="-14" cy="-9" r="3.5" fill="{c}"/><circle cx="-14" cy="9" r="3.5" fill="{c}"/>'
           '<circle cx="0" cy="-14" r="3.5" fill="{c}"/><circle cx="0" cy="0" r="3.5" fill="{c}"/>'
           '<circle cx="0" cy="14" r="3.5" fill="{c}"/><circle cx="14" cy="0" r="3.5" fill="{c}"/>',
    "stack": '<rect x="-8" y="-17" width="25" height="19" rx="2" fill="#fff"/>'
             '<rect x="-13" y="-11" width="25" height="19" rx="2" fill="#fff"/>'
             '<rect x="-18" y="-5" width="25" height="19" rx="2" fill="#fff"/>'
             '<path d="M-15 11 L-9 4 L-5 8 L-1 3 L4 11"/>',
    "chip": '<rect x="-12" y="-12" width="24" height="24" rx="3"/><rect x="-5" y="-5" width="10" height="10" fill="{c}"/>'
            '<path d="M-6 -12 V-18 M0 -12 V-18 M6 -12 V-18 M-6 12 V18 M0 12 V18 M6 12 V18 '
            'M-12 -6 H-18 M-12 0 H-18 M-12 6 H-18 M12 -6 H18 M12 0 H18 M12 6 H18"/>',
    "save": '<path d="M-16 -17 H9 L16 -10 V17 H-16Z"/><path d="M-9 -17 V-8 H7 V-17"/><rect x="-10" y="4" width="20" height="13"/>',
    "target": '<circle r="16"/><circle r="9.5"/><circle r="3" fill="{c}"/><path d="M3 -3 L17 -17 M11 -17 H17 V-11"/>',
    "sigma": '<path d="M12 -15 H-11 L3 0 L-11 15 H12" stroke-width="3.6"/>',
    "table": '<rect x="-18" y="-14" width="36" height="28" rx="2"/><path d="M-18 -5 H18 M-18 5 H18 M-6 -14 V14 M6 -14 V14"/>'
             '<rect x="-18" y="-14" width="36" height="9" fill="{c}"/>',
    "bars": '<path d="M-18 16 H18"/><rect x="-14" y="2" width="7" height="14" fill="{c}"/>'
            '<rect x="-3" y="-6" width="7" height="22" fill="{c}"/><rect x="8" y="-15" width="7" height="31" fill="{c}"/>',
    "balance": '<path d="M0 -16 V15 M-9 15 H9 M-15 -9 H15 M-15 -9 L-21 4 M-15 -9 L-9 4 M15 -9 L9 4 M15 -9 L21 4"/>'
               '<path d="M-22 4 Q-15 12 -8 4Z M8 4 Q15 12 22 4Z" fill="{c}"/>',
    "trophy": '<path d="M-11 -16 H11 V-5 A11 11 0 0 1 -11 -5Z" fill="{c}"/>'
              '<path d="M-11 -12 H-17 V-8 A7 7 0 0 0 -10 0 M11 -12 H17 V-8 A7 7 0 0 1 10 0 M0 6 V12"/>'
              '<rect x="-8" y="12" width="16" height="5" fill="{c}"/>',
    "doc": '<path d="M-12 -18 H6 L14 -10 V18 H-12Z M6 -18 V-10 H14"/><path d="M-6 -2 H8 M-6 5 H8 M-6 12 H4"/>',
    "log": '<rect x="-18" y="-14" width="36" height="28" rx="3"/><path d="M-10 -4 L-4 1 L-10 6 M0 7 H10"/>',
    "decision": '<path d="M0 -17 L17 0 L0 17 L-17 0Z"/><path d="M-4 -3 A4 4 0 1 1 1 1 V5 M0 9 V10"/>',
}


def icon(o, name, x, y, col, s=1.0):
    o.append(f'<g transform="translate({x} {y}) scale({s})" fill="none" stroke="{col}" stroke-width="3" '
             f'stroke-linecap="round" stroke-linejoin="round">{ICONS[name].format(c=col)}</g>')


# ---------------------------------------------------------------- componentes
def chip(o, x, y, t, col="#46606d", fill="#eef3f5"):
    w = len(t) * 10.9 + 26
    rect(o, x, y, w, 32, fill, rx=16)
    text(o, x + 13, y + 22, t, 18, col, family=MONO)


def box(o, col, y, h, color, title, ico, stereo, lines, path=None, log=False, border=None):
    x = bx(col)
    rect(o, x, y, BW, h, "#fff", border or color, 6 if border else 3, 14)
    o.append(f'<path d="M{x} {y + 64} V{y + 14} Q{x} {y} {x + 14} {y} H{x + BW - 14} '
             f'Q{x + BW} {y} {x + BW} {y + 14} V{y + 64}Z" fill="{color}"/>')
    o.append(f'<circle cx="{x + 38}" cy="{y + 32}" r="24" fill="#fff"/>')
    icon(o, ico, x + 38, y + 32, color, 0.95)
    text(o, x + 76, y + 42, title, 27, "#fff", "bold")
    text(o, x + 24, y + 100, f"«{stereo}»", 19, color, "bold")
    if log:
        rect(o, x + BW - 92, y + 79, 76, 30, "#eef1f8", rx=15)
        icon(o, "log", x + BW - 70, y + 94, "#435a87", 0.55)
        text(o, x + BW - 53, y + 100, "log", 17, "#435a87", "bold")
    for j, line in enumerate(lines):
        text(o, x + 24, y + 140 + j * 33, line, 22)
    if path:
        chip(o, x + 20, y + h - 48, path)


def label(o, x, y, t, col, anchor="start", size=18):
    text(o, x, y, t, size, col, "bold", anchor)


# ---------------------------------------------------------------- arquitectura
def header(o):
    text(o, 60, 92, "MENDELEY DNC · ARQUITECTURA DEL EXPERIMENTO Y SELECCIÓN ViT / DenseNet",
         46, "#102f45", "bold")
    text(o, 62, 142, "Siete etapas: de 8 000 imágenes a la selección de ViT-Base/16 y DenseNet-121 "
         "(validación cruzada estratificada 5-fold)", 27, MUTED)
    kpis = [("8 000", "imágenes"), ("10", "modelos"), ("5", "folds"), ("50", "checkpoints"), ("2", "elegidos")]
    for i, (v, k) in enumerate(kpis):
        x = 2620 + i * 236
        hi = i == 4
        rect(o, x, 44, 220, 108, "#fff8e1" if hi else "#fff", GOLD if hi else "#dbe5e9", 3 if hi else 2, 14)
        text(o, x + 110, 100, v, 44, "#8a6400" if hi else "#102f45", "bold", "middle")
        text(o, x + 110, 134, k, 21, MUTED, "normal", "middle")


def body(o):
    for x in LX:
        rect(o, x, 335, LW, 1215, "#fff", "#e1eaed", 2, 10)
    stages = [(0, 1, "Preparar datos"), (1, 2, "Preparación del experimento"), (3, 1, "Entrenar modelos"),
              (4, 1, "Evaluar en test"), (5, 1, "Reunir resultados"), (6, 1, "Comparar y seleccionar"),
              (7, 1, "Documentación")]
    for i, (c0, n, name) in enumerate(stages):
        x0, x1 = LX[c0], LX[c0 + n - 1] + LW
        o.append(f'<path d="M{x0} 195 H{x1 - 30} L{x1} 253 L{x1 - 30} 311 H{x0} L{x0 + 25} 253Z" fill="{C[i]}"/>')
        text(o, (x0 + x1) / 2, 244, f"ETAPA {i + 1}", 23, "#fff", "bold", "middle")
        text(o, (x0 + x1) / 2, 281, name, 26, "#fff", "bold", "middle")

    # bucle UML alrededor de entrenamiento + inferencia
    fx0, fx1, fy0, fy1 = LX[3] - 8, LX[4] + LW + 8, 372, 1112
    rect(o, fx0, fy0, fx1 - fx0, fy1 - fy0, "none", C[2], 3, 18, 'stroke-dasharray="12 8"')
    o.append(f'<path d="M{fx0} {fy0 + 12} Q{fx0} {fy0} {fx0 + 12} {fy0} H{fx0 + 470} V{fy0 + 24} '
             f'L{fx0 + 452} {fy0 + 40} H{fx0}Z" fill="{C[2]}"/>')
    text(o, fx0 + 16, fy0 + 28, "loop   10 modelos × 5 folds = 50 corridas", 20, "#fff", "bold")

    # Etapa 1
    box(o, 0, RA, HA, C[0], "Datos DNC", "fish", "datos",
        ["8 000 imágenes · todas las especies", "Etiqueta = carpeta de origen"], "DataSetNotClean_4k/")
    x = bx(0) + 24
    rect(o, x, RA + 194, 176, 34, "#e3f4ea", "#3d9a64", 2, 17)
    text(o, x + 88, RA + 218, "Fresco · 4 000", 19, "#1f6b40", "bold", "middle")
    rect(o, x + 188, RA + 194, 196, 34, "#fbe7e7", "#c0504d", 2, 17)
    text(o, x + 286, RA + 218, "No fresco · 4 000", 19, "#9b2f2c", "bold", "middle")
    box(o, 0, RB, HB, C[0], "Recolección", "list", "script",
        ["Lee rutas y etiquetas", "Sin abrir las imágenes", "Clases balanceadas 50 / 50"], "collect_dnc_only()")

    # Etapa 2
    box(o, 1, RA, HA, C[1], "Preparar folds", "folds", "script",
        ["StratifiedKFold · k = 5", "seed 42 (reproducible)", "Val = 15 % del resto, estratif."],
        "dnc_5fold_full_metrics.py")
    box(o, 1, RB, HB, C[1], "Partición por fold", "split", "datos", [])
    x, y, w = bx(1) + 24, RB + 122, 372
    segs = [(0.68, C[1], "68 %"), (0.12, "#e8c766", ""), (0.20, C[3], "20 %")]
    for f, col, t in segs:
        rect(o, x, y, w * f, 44, col, "#fff", 2, 6)
        if t:
            text(o, x + w * f / 2, y + 30, t, 20, "#fff", "bold", "middle")
        x += w * f
    text(o, bx(1) + 24, RB + 206, "Train 5 440 · Val 960", 22)
    text(o, bx(1) + 24, RB + 239, "Test 1 600 (reservado)", 22, C[3], "bold")
    text(o, bx(1) + 24, RB + 272, "por fold · 8 000 en total", 20, MUTED)
    box(o, 2, RA, HA, C[1], "Construir modelo", "net", "script",
        ["10 arquitecturas (timm)", "3 transformers · 7 CNN", "Pesos ImageNet · 2 clases"], "models.py")
    box(o, 2, RB, HB, C[1], "Cargador", "stack", "script",
        ["Abre las imágenes por lote", "ViT 224 px · CNN 256 px", "Aumentos solo en train"], "dataset_freshness.py")

    # Etapa 3
    box(o, 3, RA, HA, C[2], "Entrenar por fold", "chip", "script",
        ["Focal Loss γ=2 · AdamW 1e-4", "30 épocas · warmup 3", "Early stop: paciencia 7", "Guarda mejor val_F1"],
        "dnc_5fold_full_metrics.py", log=True)
    box(o, 3, RB, HB, C[2], "Checkpoints .pt", "save", "artefacto",
        ["Pesos de la mejor época", "1 por modelo y fold (50)", "Se reutilizan sin reentrenar"],
        "checkpoints_*_5fold/")

    # Etapa 4
    box(o, 4, RA, HA, C[3], "Inferencia de test", "target", "script",
        ["Carga el checkpoint del fold", "Predice el test reservado", "Matriz de confusión · ROC"],
        "dnc_5fold_confmat_roc.py", log=True)

    # Etapa 5
    box(o, 5, RA, HA, C[4], "Reunir resultados", "sigma", "script",
        ["Concatena los 5 folds", "8 000 predicciones / modelo", "Métricas pooled y por fold", "media ± std (ddof = 1)"],
        "metrics.py")
    box(o, 5, RB, HB, C[4], "Resultados CSV", "table", "artefacto",
        ["Métricas pooled por modelo", "Métricas por fold", "Matrices y ROC por modelo"],
        "3_resultados_test_por_modelo/")
    box(o, 5, RC, HC, C[4], "Gráficas PNG", "bars", "artefacto",
        ["ROC · matrices · AUC", "Curvas por época", "Gráficas del informe"], "4_graficas/ 5_graficas_informe/")

    # Etapa 6
    box(o, 6, RA, HA, C[5], "Comparar modelos", "balance", "análisis",
        ["F1 pooled y por fold", "AUC ROC", "Estabilidad (std entre folds)", "Parámetros vs. F1"])
    dx, dy = cx(6), RB + 145
    o.append(f'<path d="M{dx} {dy - 135} L{dx + 200} {dy} L{dx} {dy + 135} L{dx - 200} {dy}Z" '
             f'fill="#e9f6f4" stroke="{C[5]}" stroke-width="3"/>')
    icon(o, "decision", dx, dy - 80, C[5], 0.9)
    text(o, dx, dy - 34, "«decisión»", 19, C[5], "bold", "middle")
    text(o, dx, dy - 2, "Criterios", 24, INK, "bold", "middle")
    text(o, dx, dy + 28, "F1 · AUC · estabilidad", 21, INK, "normal", "middle")
    text(o, dx, dy + 56, "eficiencia", 21, INK, "normal", "middle")
    box(o, 6, RC, HC, C[5], "Selección vigente", "trophy", "decisión",
        ["ViT-Base/16 · F1 96.96 %", "DenseNet-121 · AUC 0.9906", "8 restantes: se conservan", "→ fuente de Generalización"],
        border=GOLD)
    rect(o, bx(6) + BW - 104, RC + 79, 88, 30, "#fff3c4", GOLD, 2, 15)
    text(o, bx(6) + BW - 60, RC + 100, "TOP-2", 17, "#8a6400", "bold", "middle")

    # Etapa 7
    box(o, 7, RA, HA, C[6], "Documentación", "doc", "artefacto",
        ["Documentación técnica", "Sustento de la selección", "Tabla comparativa (.docx)"], "6_documentacion/")

    # Registros: transversales a las etapas 1-4
    lx0, lx1 = bx(0), LX[4] + LW - 10
    rect(o, lx0, RC, lx1 - lx0, HC, "#eef1f8", "#8d9dbf", 2, 14)
    o.append(f'<circle cx="{lx0 + 70}" cy="{RC + 90}" r="44" fill="#fff" stroke="#8d9dbf" stroke-width="2"/>')
    icon(o, "log", lx0 + 70, RC + 90, "#435a87", 1.5)
    text(o, lx0 + 140, RC + 70, "REGISTROS DE EJECUCIÓN Y TRAZABILIDAD", 28, "#2f4470", "bold")
    text(o, lx0 + 140, RC + 112, "Escriben registros los componentes marcados con  [log] : entrenamiento por época, "
         "val_F1, early stopping y generación de matrices / ROC.", 22)
    text(o, lx0 + 140, RC + 146, "Sustentan la revisión de cada corrida y la documentación de la Etapa 7.", 22)
    x = lx0 + 140
    for f in ["2_logs_entrenamiento/", "dnc_5fold_resume.log", "dnc_reentreno_curvas.log", "dnc_confmat_roc_all.log"]:
        chip(o, x, RC + 186, f, "#435a87", "#fff")
        x += len(f) * 10.9 + 44

    # Flechas (continua = flujo; punteada = insumo consultado)
    arrow(o, f"M{cx(0)} {RA + HA} V{RB - 4}", C[0])
    label(o, cx(0) + 14, RA + HA + 52, "rutas y etiquetas", C[0])
    arrow(o, f"M{bx(0) + BW} {RB + 145} H{LX[1] - 13} V{RA + 190} H{bx(1) - 4}", C[0])
    arrow(o, f"M{cx(1)} {RA + HA} V{RB - 4}", C[1])
    label(o, cx(1) + 14, RA + HA + 52, "índices por fold", C[1])
    arrow(o, f"M{bx(1) + BW} {RB + 145} H{bx(2) - 4}", C[1])
    arrow(o, f"M{bx(2) + BW} {RA + 120} H{bx(3) - 4}", C[1])
    arrow(o, f"M{bx(2) + BW} {RB + 145} H{LX[3] - 22} V{RA + 200} H{bx(3) - 4}", C[1])
    arrow(o, f"M{cx(3)} {RA + HA} V{RB - 4}", C[2])
    label(o, cx(3) + 14, RA + HA + 52, "mejor val_F1", C[2])
    arrow(o, f"M{bx(3) + BW} {RB + 145} H{cx(4) - 100} V{RA + HA + 6}", C[2], dashed=True)
    label(o, bx(3) + BW + 40, RB + 132, "pesos", C[2])
    arrow(o, f"M{cx(1)} {RB + HB} V1150 H{cx(4)} V{RA + HA + 6}", C[1], dashed=True)
    label(o, cx(2), 1182, "test 20 % del mismo fold · 1 600 imágenes", C[1], "middle")
    arrow(o, f"M{bx(4) + BW} {RA + 150} H{bx(5) - 4}", C[3])
    arrow(o, f"M{cx(5)} {RA + HA} V{RB - 4}", C[4])
    label(o, cx(5) + 14, RA + HA + 52, "guardar", C[4])
    arrow(o, f"M{cx(5)} {RB + HB} V{RC - 4}", C[4])
    label(o, cx(5) + 14, RB + HB + 62, "dnc_graficas_organizadas.py", C[4], size=16)
    arrow(o, f"M{bx(5) + BW} {RB + 145} H{LX[6] - 13} V{RA + 200} H{bx(6) - 4}", C[4], dashed=True)
    arrow(o, f"M{cx(6)} {RA + HA} V{RB + 6}", C[5])
    arrow(o, f"M{cx(6)} {RB + 280} V{RC - 4}", C[5])
    label(o, cx(6) + 14, RC - 40, "top-2", C[5])
    arrow(o, f"M{bx(6) + BW} {RC + 140} H{cx(7) - 90} V{RA + HA + 6}", GOLD)
    label(o, cx(7) - 76, RB + 140, "modelos elegidos", "#8a6400")
    arrow(o, f"M{cx(5)} {RC + HC} V1520 H{cx(7) + 90} V{RA + HA + 6}", C[4])
    label(o, cx(7) + 104, RB + 140, "figuras", C[4])


def legend(o, x0, y0, w, full=True):
    rect(o, x0, y0, w, 170, "#edf4f6", "#c9d9de", 2, 12)
    text(o, x0 + 24, y0 + 40, "LECTURA DEL DIAGRAMA", 22, "#102f45", "bold")
    y = y0 + 80
    arrow(o, f"M{x0 + 24} {y} H{x0 + 104}", "#526c79")
    text(o, x0 + 118, y + 7, "flujo de datos / ejecución", 20)
    arrow(o, f"M{x0 + 24} {y + 50} H{x0 + 104}", "#526c79", dashed=True)
    text(o, x0 + 118, y + 57, "insumo consultado (pesos, test, métricas)", 20)
    x = x0 + 560
    rect(o, x, y - 16, 80, 34, "none", C[2], 3, 8, 'stroke-dasharray="10 6"')
    text(o, x + 96, y + 7, "repetición por modelo y fold", 20)
    o.append(f'<path d="M{x + 40} {y + 30} L{x + 70} {y + 50} L{x + 40} {y + 70} L{x + 10} {y + 50}Z" '
             f'fill="#e9f6f4" stroke="{C[5]}" stroke-width="3"/>')
    text(o, x + 96, y + 57, "decisión (criterios de selección)", 20)
    if not full:
        return
    x = x0 + 1060
    text(o, x, y0 + 40, "ESTEREOTIPOS", 20, "#102f45", "bold")
    for j, (s, d) in enumerate([("«datos»", "imágenes o particiones"), ("«script»", "código Python que se ejecuta"),
                                ("«artefacto»", "archivo generado"), ("«análisis» «decisión»", "comparación y elección")]):
        text(o, x + (j // 2) * 560, y + (j % 2) * 50 + 7, s, 20, "#102f45", "bold")
        text(o, x + (j // 2) * 560 + len(s) * 11 + 14, y + (j % 2) * 50 + 7, d, 20, MUTED)
    x = x0 + 2240
    text(o, x, y0 + 40, "ÍCONOS", 20, "#102f45", "bold")
    keys = [("fish", "datos"), ("folds", "folds"), ("net", "modelo"), ("chip", "entrenar"), ("save", "checkpoint"),
            ("target", "test"), ("sigma", "agregar"), ("table", "CSV"), ("bars", "gráficas"), ("trophy", "selección"),
            ("doc", "documento"), ("log", "registro")]
    for j, (k, d) in enumerate(keys):
        ix, iy = x + 20 + (j % 6) * 240, y + (j // 6) * 50
        icon(o, k, ix, iy, "#526c79", 0.75)
        text(o, ix + 26, iy + 7, d, 19, MUTED)


def svg(w, h, parts, view=None):
    vb = view or f"0 0 {w} {h}"
    cols = C + [GOLD, "#526c79"]
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="{vb}">'
            f"<defs>{markers(cols)}</defs>" + "\n".join(parts) + "</svg>")


def write(name, content):
    p = OUT / f"{name}.svg"
    p.write_text(content, encoding="utf-8")
    subprocess.run([RSVG, "-o", str(OUT / f"{name}.png"), str(p)], check=True)
    print("ok", p.with_suffix(".png"))


def continuation(o, x, y, t, col):
    w = len(t) * 10 + 28
    rect(o, x, y - 20, w, 36, "#fff", col, 2, 18)
    text(o, x + w / 2, y + 5, t, 17, col, "bold", "middle")


def arquitectura():
    W, H = 3840, 1810
    b = []
    body(b)

    o = [f'<rect width="{W}" height="{H}" fill="{BG}"/>']
    header(o)
    o += b
    legend(o, 54, 1600, 3732)
    write("arquitectura_sistema_7_etapas_v2", svg(W, H, o))

    # Recorte etapas 1-3: tapa lo que asoma de la etapa 4 y marca la continuación
    X0, X1, Y0 = 30, 2040, 100
    o = [f'<rect x="{X0}" y="{Y0}" width="{X1 - X0}" height="{H - Y0}" fill="{BG}"/>'] + b
    rect(o, LX[3] + LW + 14, 180, X1, H, BG)
    text(o, 60, 165, "Detalle · Etapas 1 a 3 (datos, experimento y entrenamiento)", 34, "#102f45", "bold")
    continuation(o, LX[3] + LW + 18, RB + 145, "→ Etapa 4", C[3])
    continuation(o, LX[3] + LW + 18, 1150, "→ Etapa 4", C[3])
    legend(o, 54, 1600, X1 - 84, full=False)
    write("arquitectura_detalle_1a3_v2", svg(X1 - X0, H - Y0, o, f"{X0} {Y0} {X1 - X0} {H - Y0}"))

    # Recorte etapas 4-7
    X0, X1 = 1760, 3810
    o = [f'<rect x="{X0}" y="{Y0}" width="{X1 - X0}" height="{H - Y0}" fill="{BG}"/>'] + b
    rect(o, X0, 180, LX[4] - 3 - X0, H, BG)
    text(o, LX[4], 165, "Detalle · Etapas 4 a 7 (evaluación, resultados, selección y documentación)",
         34, "#102f45", "bold")
    continuation(o, X0 + 12, RB + 145, "Etapa 3 →", C[2])
    continuation(o, X0 + 12, 1150, "Etapa 2 →", C[1])
    text(o, LX[4] + 20, RC + 70, "← Registros de ejecución", 22, "#2f4470", "bold")
    text(o, LX[4] + 20, RC + 104, "2_logs_entrenamiento/", 20, "#435a87", family=MONO)
    legend(o, LX[4], 1600, X1 - LX[4] - 24, full=False)
    write("arquitectura_detalle_4a7_v2", svg(X1 - X0, H - Y0, o, f"{X0} {Y0} {X1 - X0} {H - Y0}"))


# ---------------------------------------------------------------- flujo metodológico
def flujo():
    W, top, RH = 2400, 250, 176
    rows = [
        ("Preparación de datos", "fish",
         "Lectura de rutas y etiquetas de las 8 000 imágenes (4 000 fresco · 4 000 no fresco).",
         "Rutas y etiquetas", "collect_dnc_only()"),
        ("Preparación del experimento", "folds",
         "StratifiedKFold (k = 5, seed 42): train 68 % · val 12 % · test 20 % por fold. "
         "Construcción de los 10 modelos con pesos ImageNet.",
         "Índices por fold · modelos", "models.py"),
        ("Entrenamiento y checkpoint", "chip",
         "Focal Loss + AdamW hasta 30 épocas con early stopping (paciencia 7); "
         "se guarda la época con mejor val_F1.",
         "50 checkpoints .pt", "checkpoints_*_5fold/"),
        ("Evaluación sobre test", "target",
         "Inferencia sin reentrenar sobre el test reservado de cada fold: matriz de confusión, ROC y AUC.",
         "Predicciones por fold", "dnc_5fold_confmat_roc.py"),
        ("Reunir resultados", "sigma",
         "Se concatenan los 5 folds (8 000 predicciones por modelo) y se calculan métricas pooled "
         "y media ± std por fold.",
         "CSV de métricas", "3_resultados_test_por_modelo/"),
        ("Comparación y selección", "trophy",
         "Se comparan F1, AUC, estabilidad y parámetros; se eligen ViT-Base/16 (F1 96.96 %) "
         "y DenseNet-121 (AUC 0.9906).",
         "Selección vigente: 2 modelos", "5_graficas_informe/"),
        ("Documentación y trazabilidad", "doc",
         "Integración de resultados, gráficas y registros de ejecución en la documentación técnica.",
         "Documentos .docx · logs", "6_documentacion/"),
    ]
    H = top + RH * len(rows) + 60
    o = [f'<rect width="{W}" height="{H}" fill="{BG}"/>']
    text(o, 60, 88, "FLUJO METODOLÓGICO · MENDELEY DNC", 46, "#102f45", "bold")
    text(o, 62, 136, "Siete etapas, desde los datos hasta la selección de ViT-Base/16 y DenseNet-121", 27, MUTED)
    rect(o, 40, 168, W - 80, 64, "#102f45", rx=12)
    for x, t in [(110, "N.º"), (200, "ETAPA · QUÉ SE HACE"), (1560, "SALIDA · ARTEFACTO")]:
        text(o, x, 210, t, 22, "#fff", "bold")
    o.append(f'<path d="M136 {top + 60} V{top + RH * (len(rows) - 1) + 60}" stroke="#c9d9de" stroke-width="6"/>')
    for i, (name, ico, desc, out, path) in enumerate(rows):
        y, col = top + i * RH, C[i]
        hi = i == 5
        rect(o, 40, y + 12, W - 80, RH - 18, "#fff8e1" if hi else "#fff", GOLD if hi else col, 4 if hi else 2, 14)
        rect(o, 40, y + 12, 12, RH - 18, col, rx=6)
        o.append(f'<circle cx="136" cy="{y + 60}" r="40" fill="{col}"/>')
        text(o, 136, y + 75, str(i + 1), 40, "#fff", "bold", "middle")
        o.append(f'<circle cx="136" cy="{y + 126}" r="24" fill="#fff" stroke="{col}" stroke-width="2"/>')
        icon(o, ico, 136, y + 126, col, 0.9)
        text(o, 200, y + 72, name, 30, "#102f45", "bold")
        for j, line in enumerate(textwrap.wrap(desc, 95)):
            text(o, 200, y + 112 + j * 32, line, 23)
        o.append(f'<path d="M1530 {y + 32} V{y + RH - 26}" stroke="#e1eaed" stroke-width="2"/>')
        text(o, 1560, y + 76, out, 25, "#8a6400" if hi else col, "bold")
        chip(o, 1560, y + 100, path)
        if hi:
            rect(o, W - 200, y + 34, 130, 36, "#fff3c4", GOLD, 2, 18)
            text(o, W - 135, y + 59, "TOP-2", 19, "#8a6400", "bold", "middle")
    write("flujo_metodologico_7_etapas_v2", svg(W, H, o))


if __name__ == "__main__":
    arquitectura()
    flujo()
