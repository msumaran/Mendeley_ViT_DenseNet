"""Pasa Documentacion_Tecnica_10_Modelos.docx a las corridas que completan 30 épocas.

Cuatro corridas se habían detenido por paciencia y se repitieron hasta 30 épocas (logs
2_logs_entrenamiento/completar_30ep_*.log): ViT-Base/16 fold 2, Swin-Base fold 4,
EfficientNet-Lite0 fold 3 y ResNet-50 fold 3. Las tablas salen de esos logs y de
3_resultados_test_por_modelo/; las cifras anteriores se leen de git (`ANTES`).

Uso (desde la raíz del repositorio): python utilitarios/documentacion/actualizar_30_epocas.py [salida.docx]
"""
import copy
import hashlib
import io
import re
import statistics
import subprocess
import sys
import zipfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

DOC = Path("6_documentacion/Documentacion_Tecnica_10_Modelos.docx")
ANTES = "530a1bd"  # última versión con las cifras anteriores
RES = Path("3_resultados_test_por_modelo")
LOGS = Path("2_logs_entrenamiento")
GRAF = Path("7_graficas_por_modelo")
PARTE1, PARTE2 = "completar_30ep_parte1.log", "completar_30ep_parte2.log"
MODELOS = {  # clave: (nombre, sección, fold repetido, log, ¿cambió el resultado de test?)
    "vit_base_patch16_224": ("ViT-Base/16", "3.1", 2, PARTE2, True),
    "swin_base_patch4_window7_224": ("Swin-Base", "3.3", 4, PARTE2, True),
    "efficientnet_lite0": ("EfficientNet-Lite0", "3.6", 3, PARTE1, True),
    "resnet50": ("ResNet-50", "3.8", 3, PARTE1, False),
}
ANTERIORES = {"vit_base_patch16_224": 23, "swin_base_patch4_window7_224": 25, "efficientnet_lite0": 23, "resnet50": 28}


def miles(n):
    return f"{int(n):,}"


# --- datos ---------------------------------------------------------------------
def leer_curvas(log):
    """{(modelo, fold): [(época, val_F1, mejor, paciencia), ...]}; una corrida sin test (cortada) se ignora."""
    curvas, cur = {}, None
    for linea in (LOGS / log).read_text(errors="replace").splitlines():
        m = re.match(r"\[(\S+)\] fold (\d)/5", linea)
        if m:
            cur = (m[1], int(m[2]))
            curvas[cur] = []
        m = re.match(r"\s+epoch\s+(\d+)/\d+\s+val_F1=([\d.]+)\s+best=([\d.]+)\s+pat=(\d+)", linea)
        if m and cur:
            curvas[cur].append((int(m[1]), m[2], m[3], int(m[4])))
    return curvas


def csv_git(ruta):
    return pd.read_csv(io.StringIO(subprocess.run(["git", "show", f"{ANTES}:{ruta}"], capture_output=True, text=True,
                                                  check=True).stdout))


def cm(modelo, parte):
    return pd.read_csv(RES / modelo / f"confusion_matrix_{parte}.csv", index_col=0).values


def clases(c):
    tp, fn, fp, tn = (int(x) for x in c.ravel())
    pf, rf, pn, rn = tp / (tp + fp) * 100, tp / (tp + fn) * 100, tn / (tn + fn) * 100, tn / (tn + fp) * 100
    return dict(tp=tp, fn=fn, fp=fp, tn=tn, pf=pf, rf=rf, pn=pn, rn=rn, f1f=2 * pf * rf / (pf + rf),
                f1n=2 * pn * rn / (pn + rn))


# --- edición del documento -------------------------------------------------------
def parrafos(contenedor):
    yield from contenedor.paragraphs
    for t in contenedor.tables:
        for fila in t.rows:
            for celda in fila.cells:
                yield from celda.paragraphs


def cambiar(p, viejo, nuevo):
    """Reemplaza `viejo` aunque esté repartido entre varios runs; el formato es el del primer run tocado."""
    if viejo not in p.text:
        return False
    runs = p.runs
    texto_runs = "".join(r.text for r in runs)  # p.text también incluye los enlaces, que no están en p.runs
    assert viejo in texto_runs, f"cruza un enlace: {viejo}"
    i = texto_runs.index(viejo)
    pos = 0
    for k, r in enumerate(runs):
        fin = pos + len(r.text)
        if pos <= i < fin:
            resto = len(viejo)
            texto = r.text
            ini = i - pos
            r.text = texto[:ini] + nuevo + texto[ini + resto:] if ini + resto <= len(texto) else texto[:ini] + nuevo
            resto -= len(texto) - ini
            for s in runs[k + 1:]:
                if resto <= 0:
                    break
                t = s.text
                s.text = t[resto:]
                resto -= len(t)
            return True
        pos = fin
    raise AssertionError(f"texto fuera de runs (¿enlace?): {viejo}")


def global_(doc, viejo, nuevo):
    n = sum(cambiar(p, viejo, nuevo) for p in parrafos(doc))
    assert n, viejo


def poner(celda, texto):
    p = celda.paragraphs[0]
    runs = [r for r in p.runs if r.text] or p.runs
    assert runs, celda.text
    runs[0].text = str(texto)
    for r in runs[1:]:
        r.text = ""


def seccion(doc, prefijo):
    """Elementos (Paragraph o Table) entre el título que empieza con `prefijo` y el siguiente título."""
    out, dentro = [], False
    for el in doc.element.body.iterchildren():
        if el.tag.endswith("}p"):
            p = Paragraph(el, doc)
            if p.style.name.startswith("Heading"):
                if dentro:
                    break
                dentro = p.text.startswith(prefijo)
                continue
            if dentro:
                out.append(p)
        elif el.tag.endswith("}tbl") and dentro:
            out.append(Table(el, doc))
    assert out, prefijo
    return out


def tabla_con(doc, primera):
    ts = [t for t in doc.tables if t.rows[0].cells[0].text.strip() == primera[0]
          and all(c in [x.text.strip() for x in t.rows[0].cells] for c in primera)]
    assert len(ts) == 1, primera
    return ts[0]


def fila_de(tabla, clave):
    fs = [f for f in tabla.rows if f.cells[0].text.strip() == clave]
    assert len(fs) == 1, clave
    return fs[0]


def tabla_epocas(tabla, curva, elegida):
    filas = tabla.rows[1:]
    resaltada = next(f for f in filas if any(r.bold for c in f.cells for p in c.paragraphs for r in p.runs))
    normal = next(f for f in filas if f is not resaltada)
    moldes = copy.deepcopy(normal._tr), copy.deepcopy(resaltada._tr)
    for f in filas:
        tabla._tbl.remove(f._tr)
    for ep, val, mejor, pat in curva:
        tabla._tbl.append(copy.deepcopy(moldes[ep == elegida]))
        fila = tabla.rows[-1]
        for celda, v in zip(fila.cells, (ep, val, mejor, pat)):
            poner(celda, v)


def curva_png(modelo, nombre, curvas):
    colores = ["#3b5b8c", "#e8a33d", "#5f9163", "#b5495b", "#7f5a96"]
    fig, ax = plt.subplots(figsize=(9, 5), dpi=150)
    for k in range(1, 6):
        c = curvas[modelo, k]
        x, y = [e[0] for e in c], [float(e[1]) for e in c]
        ax.plot(x, y, color=colores[k - 1], alpha=.9, lw=1.6, label=f"Fold {k}")
        i = y.index(max(y))
        ax.plot(x[i], y[i], "o", color=colores[k - 1], mec="white", ms=7)
    ax.plot([], [], "o", color="black", label="= checkpoint elegido")
    ax.set(title=f"{nombre} — val_F1 por epoch", xlabel="Epoch", ylabel="val_F1 (%)")
    ax.legend(loc="lower right", fontsize=8)
    ruta = GRAF / f"epochs_{modelo}.png"
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return ruta


def reemplazar_imagenes(doc, pares):
    """pares = {md5 de la imagen anterior: ruta nueva}; devuelve cuántas imágenes del documento cambió."""
    n = 0
    for parte in doc.part.package.iter_parts():
        if str(parte.partname).startswith("/word/media/"):
            nuevo = pares.get(hashlib.md5(parte.blob).hexdigest())
            if nuevo is not None:
                parte._blob = Path(nuevo).read_bytes()
                n += 1
    return n


def imagen_de_seccion(doc, prefijo):
    """Parte de imagen de la primera figura de la sección (para las que no están como archivo en git)."""
    for el in seccion(doc, prefijo):
        if isinstance(el, Paragraph):
            blip = el._p.xpath(".//a:blip/@r:embed")
            if blip:
                return doc.part.related_parts[blip[0]]
    raise AssertionError(prefijo)


# --- cambios --------------------------------------------------------------------
def main(salida):
    assert not list(DOC.parent.glob("~$cumentacion_Tecnica_10_Modelos*")), "cierra el documento en Word"
    doc = Document(DOC)
    pooled_ant = csv_git(RES / "metricas_pooled_por_modelo.csv").set_index("model")
    pooled = pd.read_csv(RES / "metricas_pooled_por_modelo.csv").set_index("model")
    por_fold = pd.read_csv(RES / "metricas_por_fold.csv")

    curvas = {}
    for log in ("dnc_reentreno_curvas.log", "dnc_5fold_resume.log", PARTE1, PARTE2):  # el último gana
        for clave, c in leer_curvas(log).items():
            if c and (len(c) == 30 or clave not in curvas):
                curvas[clave] = c
    assert all(len(curvas[m, k]) == 30 for m in MODELOS for k in range(1, 6))

    # Cifras generales por modelo: tablas de 1.5 y 1.6
    t15 = tabla_con(doc, ("Modelo", "TP", "FN", "FP", "TN"))
    t16 = tabla_con(doc, ("Modelo", "Tipo", "Parámetros"))
    for modelo, (nombre, sec, fold, _, cambia) in MODELOS.items():
        if not cambia:
            continue
        v, a, c = pooled.loc[modelo], pooled_ant.loc[modelo], clases(cm(modelo, "pooled"))
        std = por_fold[por_fold.model == modelo].f1.std()
        for celda, val in zip(fila_de(t15, nombre).cells[1:], (miles(c["tp"]), miles(c["fn"]), miles(c["fp"]),
                                                              miles(c["tn"]), f"{v.accuracy:.2f}", f"{v.precision:.2f}",
                                                              f"{v.f1:.2f}")):
            poner(celda, val)
        celdas = fila_de(t16, nombre).cells
        for celda, val in zip(celdas[3:], (f"{v.accuracy:.2f}", f"{v.precision:.2f}", f"{v.f1:.2f}", f"{std:.2f}",
                                           f"{v.auc:.4f}")):
            poner(celda, val)

        # Sección 3.x: encabezado y pooled
        cab = next(e for e in seccion(doc, f"{sec}. ") if isinstance(e, Paragraph) and e.text.startswith("F1 macro ="))
        assert cambiar(cab, f"F1 macro = {a.f1:.2f}", f"F1 macro = {v.f1:.2f}")
        assert cambiar(cab, f"Accuracy = {a.accuracy:.2f}", f"Accuracy = {v.accuracy:.2f}")
        assert cambiar(cab, f"AUC = {a.auc:.4f}", f"AUC = {v.auc:.4f}")

        filas_f = por_fold[(por_fold.model == modelo) & (por_fold.fold == fold)].iloc[0]
        t2 = next(e for e in seccion(doc, f"{sec}.2 ") if isinstance(e, Table))
        for celda, col in zip(fila_de(t2, f"Fold {fold}").cells[4:], ("accuracy", "precision", "recall", "f1",
                                                                     "specificity")):
            poner(celda, f"{filas_f[col]:.2f}")
        poner(fila_de(t2, f"Fold {fold}").cells[9], f"{filas_f.auc:.4f}")

        folds = [clases(cm(modelo, f"fold{k}")) for k in range(1, 6)]
        for e in seccion(doc, f"{sec}.4 "):
            if isinstance(e, Table):
                for clase, (tp, fp, fn, tn, pr, re_, f1, es) in {
                    "fresco": (c["tp"], c["fp"], c["fn"], c["tn"], c["pf"], c["rf"], c["f1f"], c["rn"]),
                    "no fresco": (c["tn"], c["fn"], c["fp"], c["tp"], c["pn"], c["rn"], c["f1n"], c["rf"]),
                }.items():
                    for celda, val in zip(fila_de(e, clase).cells[1:], (tp, fp, fn, tn, f"{pr:.2f}", f"{re_:.2f}",
                                                                       f"{f1:.2f}", f"{es:.2f}")):
                        poner(celda, val)
                for celda, val in zip(fila_de(e, "MACRO (reportado)").cells[5:], (v.precision, v.recall, v.f1,
                                                                                 v.specificity)):
                    poner(celda, f"{val:.2f}")
                continue
            for pref, campo in (("TP (", "tp"), ("FN (", "fn"), ("FP (", "fp"), ("TN (", "tn")):
                if e.text.startswith(pref):
                    previo = e.text.split(": ")[1]
                    assert cambiar(e, previo, "+".join(str(f[campo]) for f in folds) + f" = {c[campo]}")
            if "Con fresco como clase positiva se leen" in e.text:
                m = re.search(r"TP = \S+, FN = \S+, FP = \S+ y TN = [\d,]+", e.text)
                assert cambiar(e, m[0], f"TP = {miles(c['tp'])}, FN = {miles(c['fn'])}, FP = {miles(c['fp'])} y TN = {miles(c['tn'])}")
                m = re.search(r"su TP es \S+, su FP es \S+ y su FN es [\d,]+", e.text)
                assert cambiar(e, m[0], f"su TP es {miles(c['tn'])}, su FP es {miles(c['fn'])} y su FN es {miles(c['fp'])}")
            if e.text.startswith("La matriz también muestra"):
                m = re.search(r"(\d+) imágenes fresco se clasificaron como no fresco y (\d+) imágenes no fresco", e.text)
                assert cambiar(e, m[0], f"{c['fn']} imágenes fresco se clasificaron como no fresco y {c['fp']} "
                                        f"imágenes no fresco")
            if e.text.startswith("F1 macro = ("):
                assert cambiar(e, e.text, f"F1 macro = ({c['f1f']:.2f} + {c['f1n']:.2f}) / 2 = {v.f1:.2f}")
            if e.text.startswith("(valor guardado:"):
                assert cambiar(e, e.text, f"(valor guardado: {v.f1:.2f})")
            if e.text.startswith("Accuracy = ("):
                assert cambiar(e, e.text, f"Accuracy = ({c['tp']}+{c['tn']}) / 8000 × 100 = {v.accuracy:.2f}")
            if "curva ROC pooled (der.), AUC =" in e.text:
                assert cambiar(e, f"AUC = {a.auc:.4f}", f"AUC = {v.auc:.4f}")

    # Sección 3.x.1 de las 4 corridas repetidas
    for modelo, (nombre, sec, fold, log, _) in MODELOS.items():
        curva = curvas[modelo, fold]
        val = [float(e[1]) for e in curva]
        elegida = val.index(max(val)) + 1
        els = seccion(doc, f"{sec}.1 ")
        resumen = next(e for e in els if isinstance(e, Table))
        fila = fila_de(resumen, f"Fold {fold}")
        for celda, x in zip(fila.cells[1:], (elegida, curva[elegida - 1][1], 30)):
            poner(celda, x)
        i = next(n for n, e in enumerate(els) if isinstance(e, Paragraph) and e.text.strip() == f"Fold {fold}")
        tabla_epocas(next(e for e in els[i + 1:] if isinstance(e, Table)), curva, elegida)
        arch = next(e for e in els if isinstance(e, Paragraph) and e.text.startswith("Archivo:"))
        arch.add_run(f" El fold {fold} viene de 2_logs_entrenamiento/{log}, que lo completa hasta 30 épocas.")
        arch.runs[-1].font.size = arch.runs[0].font.size

    global_(doc, " Único efecto de redondeo: EfficientNet-Lite0 tiene Accuracy = 7,282/8,000 = 91.025 %, guardado como "
                 "91.02.", "")

    # 1.3: todas las corridas completan 30 épocas
    global_(doc, "Cada entrenamiento dura hasta 30 épocas y se detiene antes si pasan 7 épocas sin mejora (early "
                 "stopping).", "Cada entrenamiento completa 30 épocas.")
    global_(doc, "; al llegar a 7 épocas consecutivas sin mejora se detiene el entrenamiento. También se detiene al "
                 "completar 30 épocas.",
            "; en la corrida original, al llegar a 7 épocas consecutivas sin mejora se detenía el entrenamiento. Las "
            "4 corridas que se detuvieron así se repitieron sin ese límite, hasta 30 épocas.")
    nota = next(p for p in doc.paragraphs if p.text.startswith("Por qué algunos folds terminan antes de 30 épocas"))
    assert cambiar(nota, "Por qué algunos folds terminan antes de 30 épocas.", "Por qué todas las corridas tienen 30 épocas.")
    assert cambiar(nota, nota.text[nota.text.index("Cada fold puede correr"):],
                   "Las 50 corridas completan 30 épocas; con 5 folds, cada modelo suma 150. En la corrida original, "
                   "46 llegaron a 30 épocas y 4 se detuvieron antes por la paciencia de 7. Esas 4 se repitieron "
                   "desde el inicio, sin paciencia, con los mismos folds, y sus resultados reemplazan a los anteriores:")
    t_cuatro = tabla_con(doc, ("Modelo", "Fold", "Comprobación"))
    for celda, x in zip(t_cuatro.rows[0].cells[2:], ("Épocas en la corrida original", "Época elegida (30 épocas)",
                                                     "F1 de test (antes → ahora)")):
        poner(celda, x)
    ant_fold = csv_git(RES / "metricas_por_fold.csv")
    for fila in t_cuatro.rows[1:]:
        nombre = fila.cells[0].text.strip()
        modelo = next(m for m, d in MODELOS.items() if d[0] == nombre)
        fold = MODELOS[modelo][2]
        val = [float(e[1]) for e in curvas[modelo, fold]]
        antes = ant_fold[(ant_fold.model == modelo) & (ant_fold.fold == fold)].f1.iloc[0]
        ahora = por_fold[(por_fold.model == modelo) & (por_fold.fold == fold)].f1.iloc[0]
        for celda, x in zip(fila.cells[2:], (ANTERIORES[modelo], val.index(max(val)) + 1, f"{antes:.2f} → {ahora:.2f}")):
            poner(celda, x)
    resto = next(p for p in doc.paragraphs if p.text.startswith("Por eso ViT-Base/16 suma 143"))
    assert cambiar(resto, resto.text[:resto.text.index("Las curvas de esos folds")],
                   "Siempre se usa el checkpoint de la mejor época, no el de la última. ")
    t_corr = tabla_con(doc, ("Modelo", "Corridas (folds)", "Épocas ejecutadas"))
    for fila in t_corr.rows[1:-1]:
        poner(fila.cells[2], 150)
    poner(t_corr.rows[-1].cells[2], "1,500")
    global_(doc, "si pasan 7 revisiones seguidas sin mejora, se detiene. El máximo es de 30 revisiones.",
            "cada modelo completa 30 revisiones.")
    global_(doc, "(por agotar las 30 épocas o por early stopping)", "(al completar las 30 épocas)")
    global_(doc, "Si el modelo no mejora durante 7 épocas seguidas, el entrenamiento se detiene antes de llegar a 30.",
            "Cuenta las épocas seguidas sin mejora. La paciencia original era 7; las 4 corridas que se detuvieron "
            "por ella se repitieron hasta 30 épocas.")
    global_(doc, "hasta 30 épocas. Al final de cada época mide el F1 en validación y para si pasan 7 épocas sin mejora.",
            "durante 30 épocas. Al final de cada época mide el F1 en validación y conserva la mejor versión.")
    global_(doc, "El límite de 30 épocas, junto con la paciencia de 7, da", "El límite de 30 épocas da")
    global_(doc, "durante un número de épocas (paciencia = 7).",
            "durante un número de épocas (paciencia = 7 en la corrida original; las 4 corridas detenidas se "
            "repitieron hasta 30 épocas).")
    global_(doc, "para elegir el checkpoint y decidir el early stopping.", "para elegir el checkpoint.")
    global_(doc, "2_logs_entrenamiento/dnc_5fold_resume.log y dnc_reentreno_curvas.log",
            "2_logs_entrenamiento/dnc_5fold_resume.log, dnc_reentreno_curvas.log y completar_30ep_parte{1,2}.log")

    # Las 4 corridas repetidas se ejecutaron en Mac (device=mps en sus logs)
    entorno = next(t for t in doc.tables if t.cell(1, 0).text.strip() == "Hardware")
    so = fila_de(entorno, "Sistema operativo")
    nueva = copy.deepcopy(so._tr)
    so._tr.addnext(nueva)
    fila = next(f for f in entorno.rows if f._tr is nueva)
    poner(fila.cells[0], "Repetición de 4 corridas a 30 épocas")
    poner(fila.cells[1], "Mac con aceleración Metal (PyTorch, dispositivo mps), según "
                         "2_logs_entrenamiento/completar_30ep_parte1.log y completar_30ep_parte2.log")

    # ViT-Base/16 como ejemplo en 1.6 y 2.2.1
    vit = clases(cm("vit_base_patch16_224", "pooled"))
    fvit = por_fold[por_fold.model == "vit_base_patch16_224"]
    p = pooled.loc["vit_base_patch16_224"]
    assert (vit["tp"], vit["fn"], vit["fp"], vit["tn"]) == (3861, 139, 98, 3902)
    tfold = tabla_con(doc, ("Fold", "TP", "FN", "FP", "TN", "AUC del fold"))
    c2 = clases(cm("vit_base_patch16_224", "fold2"))
    f2 = fvit[fvit.fold == 2].iloc[0]
    for celda, x in zip(fila_de(tfold, "Fold 2").cells[1:], (c2["tp"], c2["fn"], c2["fp"], c2["tn"], f"{f2.f1:.2f}",
                                                            f"{f2.auc:.4f}")):
        poner(celda, x)
    for celda, x in zip(fila_de(tfold, "Suma (pooled)").cells[1:5], ("3,861", "139", "98", "3,902")):
        poner(celda, x)
    tcol = tabla_con(doc, ("Columna", "Cómo se obtiene"))
    poner(fila_de(tcol, "Accuracy").cells[2], "(3,861 + 3,902) / 8,000 × 100")
    poner(fila_de(tcol, "Accuracy").cells[3], f"{p.accuracy:.2f}")
    calc = fila_de(tcol, "Precisión macro")
    poner(calc.cells[2], f"fresco 3,861 / (3,861 + 98) = {vit['pf']:.2f}\nno fresco 3,902 / (3,902 + 139) = "
                         f"{vit['pn']:.2f}\n({vit['pf']:.2f} + {vit['pn']:.2f}) / 2")
    poner(calc.cells[3], f"{p.precision:.2f}")
    calc = fila_de(tcol, "F1 macro (pooled)")
    for par, texto in zip(calc.cells[2].paragraphs,
                          (f"F1_fresco = 2×3,861 / (2×3,861 + 98 + 139) = {vit['f1f']:.2f}",
                           f"F1_no_fresco = 2×3,902 / (2×3,902 + 139 + 98) = {vit['f1n']:.2f}",
                           f"({vit['f1f']:.2f} + {vit['f1n']:.2f}) / 2")):
        par.runs[0].text = texto
    poner(calc.cells[3], f"{p.f1:.2f}")
    calc = fila_de(tcol, "Desv. std F1")
    poner(calc.cells[2], "stdev(" + ", ".join(f"{x:.2f}" for x in fvit.sort_values("fold").f1) + ")")
    poner(calc.cells[3], f"{statistics.stdev(fvit.f1):.2f}")
    poner(fila_de(tcol, "AUC (pooled)").cells[3], f"{p.auc:.4f}")
    media = round(fvit.f1.mean() + 1e-9, 2)
    global_(doc, "(media de los 5 = 96.96); el AUC pooled tampoco es el promedio de los AUC de fold (sería 0.9901)",
            f"(media de los 5 = {media:.2f}); el AUC pooled tampoco es el promedio de los AUC de fold "
            f"(sería {fvit.auc.mean():.4f})")
    for viejo, nuevo in (
            ("F1 macro de 96.96 %", "F1 macro de 97.04 %"),
            ("(96.96) con una desviación baja entre folds (0.28)", "(97.04) con una desviación baja entre folds (0.18)"),
            ("ViT-Base/16 (96.96) y DenseNet-121 (95.66)", "ViT-Base/16 (97.04) y DenseNet-121 (95.66)"),
            ("FP = 104 y TN = 3,896", "FP = 98 y TN = 3,902"),
            ("3,861 / (3,861 + 104) × 100 = 97.38 %", f"3,861 / (3,861 + 98) × 100 = {vit['pf']:.2f} %"),
            ("3,896 / (3,896 + 139) × 100 = 96.56 %", f"3,902 / (3,902 + 139) × 100 = {vit['pn']:.2f} %"),
            ("(97.38 + 96.56) / 2 = 96.97 %", f"({vit['pf']:.2f} + {vit['pn']:.2f}) / 2 = {p.precision:.2f} %"),
            ("TP = 3,896 y FP = 139", "TP = 3,902 y FP = 139"),
            ("(96.96 %) y es muy estable entre folds (desviación estándar de 0.28)",
             "(97.04 %) y es muy estable entre folds (desviación estándar de 0.18)")):
        global_(doc, viejo, nuevo)

    # Figuras: las del commit de resultados y las curvas por época de los 4 modelos
    subprocess.run(["git", "diff", "--quiet"], check=False)
    lista = subprocess.run(["git", "diff", "--name-only", ANTES, "--", "*.png"], capture_output=True, text=True).stdout
    pares = {}
    for ruta in lista.split("\n"):
        if ruta and Path(ruta).exists():
            viejo = subprocess.run(["git", "show", f"{ANTES}:{ruta}"], capture_output=True)
            if viejo.returncode == 0:
                pares[hashlib.md5(viejo.stdout).hexdigest()] = ruta
    # la imagen de Lite0 no está en git: se ubica por sección antes de regenerar las demás
    lite = imagen_de_seccion(doc, "3.6.1 ")
    for modelo, (nombre, *_ ) in MODELOS.items():
        ruta = curva_png(modelo, nombre, curvas)
        if modelo == "efficientnet_lite0":
            lite._blob = ruta.read_bytes()
        else:
            viejo = subprocess.run(["git", "show", f"{ANTES}:{ruta}"], capture_output=True, check=True).stdout
            pares[hashlib.md5(viejo).hexdigest()] = ruta
    print("imágenes reemplazadas:", reemplazar_imagenes(doc, pares) + 1)
    doc.save(salida)
    zipfile.ZipFile(salida).testzip()


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else DOC)
