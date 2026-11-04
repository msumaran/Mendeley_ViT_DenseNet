"""Regenera las graficas que cambian al completar los 4 folds cortados a 30 epochs.
Ejecutar desde la raiz del repo. Lee logs y CSV del repo; no entrena ni reevalua."""
import os
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Patch

LOGS = ["completar_30ep_parte2.log", "completar_30ep_parte1.log", "dnc_reentreno_curvas.log", "dnc_5fold_resume.log"]
RES, G7, G5 = "3_resultados_test_por_modelo", "7_graficas_por_modelo", "5_graficas_informe"
MODELOS = {  # id: (nombre, nombre corto, familia, parametros en millones), en orden de la leyenda original
    "mobilenetv3_small_100": ("MobileNetV3-Small", "MobileNetV3-S", "CNN", 1.52),
    "efficientnet_lite0": ("EfficientNet-Lite0", "EfficientNet-Lite0", "CNN", 3.37),
    "vit_base_patch16_224": ("ViT-Base/16", "ViT-Base/16", "ViT", 85.80),
    "efficientnet_b0": ("EfficientNet-B0", "EfficientNet-B0", "CNN", 4.01),
    "mobilenetv2_100": ("MobileNetV2", "MobileNetV2", "CNN", 2.23),
    "densenet121": ("DenseNet-121", "DenseNet-121", "CNN", 6.96),
    "resnet50": ("ResNet-50", "ResNet-50", "CNN", 23.51),
    "mobilenetv1_100": ("MobileNetV1", "MobileNetV1", "CNN", 3.21),
    "swin_base_patch4_window7_224": ("Swin-Base", "Swin-Base", "ViT", 86.75),
    "vit_small_patch14_dinov2": ("DINOv2-Small", "DINOv2-Small", "ViT", 21.63),
}
COLOR = {"CNN": "#4C72B0", "ViT": "#DD8452"}
COLOR_FOLD = ["#4C72B0", "#DD8452", "#55A868", "#C44E52", "#8172B3"]


def curvas():
    """{(modelo, fold): [(epoch, val_f1, mejor, paciencia)]}. Gana el primer log de LOGS con 30 epochs o el mas largo."""
    salida = {}
    for log in LOGS:
        actual = None
        por_log = {}
        for linea in open(os.path.join("2_logs_entrenamiento", log), errors="replace"):
            m = re.match(r"--- (\S+) fold (\d) ---", linea)
            if m:
                actual = (m.group(1), int(m.group(2)))
                por_log[actual] = []
            m = re.search(r"epoch\s+(\d+)/\d+\s+val_F1=([\d.]+)\s+best=([\d.]+)\s+pat=(\d+)", linea)
            if m and actual:
                por_log[actual].append((int(m.group(1)), float(m.group(2)), float(m.group(3)), int(m.group(4))))
        for k, v in por_log.items():
            if v and len(v) > len(salida.get(k, [])):
                salida[k] = v
    return salida


def epochs(m, c):
    plt.figure(figsize=(9, 5))
    for k in range(1, 6):
        filas = c[(m, k)]
        assert len(filas) == 30, (m, k, len(filas))
        x, y = [f[0] for f in filas], [f[1] for f in filas]
        i = y.index(max(y))
        plt.plot(x, y, color=COLOR_FOLD[k - 1], lw=1.6, alpha=.9, label=f"Fold {k}")
        plt.scatter(x[i], y[i], color=COLOR_FOLD[k - 1], edgecolor="white", s=45, zorder=5)
    plt.title(f"{MODELOS[m][0]} — val_F1 por epoch")
    plt.xlabel("Epoch")
    plt.ylabel("val_F1 (%)")
    plt.legend(title="● = checkpoint elegido", fontsize=8, loc="lower right")
    plt.tight_layout()
    plt.savefig(f"{G7}/epochs_{m}.png", dpi=150)
    plt.close()


def cm_folds(m, por_fold):
    fig, ejes = plt.subplots(1, 5, figsize=(17.9, 4.24))
    for k, ax in enumerate(ejes, 1):
        cm = pd.read_csv(f"{RES}/{m}/confusion_matrix_fold{k}.csv", index_col=0).values
        ax.imshow(cm, cmap="Blues", vmin=0)
        for i in range(2):
            for j in range(2):
                ax.text(j, i, cm[i, j], ha="center", va="center", color="white" if cm[i, j] > cm.max() / 2 else "black")
        ax.set_xticks([0, 1], ["fresco", "no_fresco"], fontsize=8)
        ax.set_yticks([0, 1], ["fresco", "no_fresco"], fontsize=8)
        ax.set_xlabel("Prediccion", fontsize=8)
        f1 = por_fold[(por_fold.model == m) & (por_fold.fold == k)].f1.iloc[0]
        ax.set_title(f"Fold {k}\nF1={f1:.2f}", fontsize=9)
    ejes[0].set_ylabel("Real", fontsize=8)
    fig.suptitle(f"{MODELOS[m][0]} — Matrices de confusión por fold")
    fig.tight_layout()
    fig.savefig(f"{G7}/cm_folds_{m}.png", dpi=150)
    plt.close(fig)


def roc_folds(m, por_fold, pooled):
    plt.figure(figsize=(6, 6))
    for k in range(1, 6):
        r = pd.read_csv(f"{RES}/{m}/roc_curve_fold{k}.csv")
        auc = por_fold[(por_fold.model == m) & (por_fold.fold == k)].auc.iloc[0]
        plt.plot(r.fpr, r.tpr, color=COLOR_FOLD[k - 1], lw=1.2, label=f"Fold {k} (AUC={auc:.4f})")
    r = pd.read_csv(f"{RES}/{m}/roc_curve_pooled.csv")
    plt.plot(r.fpr, r.tpr, "k", lw=2.5, label=f"Pooled (AUC={pooled.loc[m, 'auc']:.4f})")
    plt.plot([0, 1], [0, 1], "--", color="gray", lw=1)
    plt.title(f"{MODELOS[m][0]} — Curvas ROC: 5 folds + pooled")
    plt.xlabel("Tasa de falsos positivos")
    plt.ylabel("Tasa de verdaderos positivos")
    plt.legend(fontsize=8, loc="lower right")
    plt.tight_layout()
    plt.savefig(f"{G7}/roc_folds_{m}.png", dpi=150)
    plt.close()


def informe(por_fold, pooled):
    orden = list(pooled.sort_values("f1", ascending=False).index)
    leyenda = [Patch(color=COLOR[f], label=f) for f in COLOR]

    plt.figure(figsize=(8, 5))
    std = por_fold.groupby("model").f1.std()
    plt.bar([MODELOS[m][1] for m in orden], pooled.loc[orden, "f1"], yerr=std[orden], capsize=4,
            color=[COLOR[MODELOS[m][2]] for m in orden])
    plt.ylim(80, 100)
    plt.xticks(rotation=40, ha="right", fontsize=8)
    plt.ylabel("F1 (%) pooled  +/- desv. estandar entre folds")
    plt.title("F1 y estabilidad entre folds — 10 modelos")
    plt.legend(handles=leyenda)
    plt.tight_layout()
    plt.savefig(f"{G5}/f1_estabilidad.png", dpi=150)
    plt.close()

    plt.figure(figsize=(7, 6))
    for m in orden:
        x, y = MODELOS[m][3], pooled.loc[m, "f1"]
        plt.scatter(x, y, s=90, color=COLOR[MODELOS[m][2]], marker="^" if MODELOS[m][2] == "ViT" else "o")
        plt.annotate(MODELOS[m][1], (x, y), xytext=(5, 5), textcoords="offset points", fontsize=8)
    plt.xscale("log")
    plt.xlabel("Parametros (millones)")
    plt.ylabel("F1 (%) pooled")
    plt.title("Eficiencia: parametros vs. F1")
    plt.legend(handles=leyenda)
    plt.tight_layout()
    plt.savefig(f"{G5}/params_vs_f1.png", dpi=150)
    plt.close()

    plt.figure(figsize=(6, 6))
    for m in orden:
        x, y = pooled.loc[m, "recall"], pooled.loc[m, "precision"]
        plt.scatter(x, y, s=90, color=COLOR[MODELOS[m][2]], marker="^" if MODELOS[m][2] == "ViT" else "o")
        plt.annotate(MODELOS[m][1], (x, y), xytext=(5, 5), textcoords="offset points", fontsize=8)
    plt.xlim(85, 100)
    plt.ylim(85, 100)
    plt.xlabel("Recall (%)")
    plt.ylabel("Precision (%)")
    plt.title("Precision vs. Recall — 10 modelos (pooled)")
    plt.legend(handles=leyenda)
    plt.tight_layout()
    plt.savefig(f"{G5}/precision_vs_recall.png", dpi=150)
    plt.close()

    plt.figure(figsize=(7, 7))
    for m in MODELOS:
        r = pd.read_csv(f"{RES}/{m}/roc_curve_pooled.csv")
        plt.plot(r.fpr, r.tpr, lw=2, label=f"{MODELOS[m][0]} (AUC={pooled.loc[m, 'auc']:.4f})")
    plt.plot([0, 1], [0, 1], "--", color="gray", lw=1)
    plt.xlabel("Tasa de falsos positivos")
    plt.ylabel("Tasa de verdaderos positivos")
    plt.title("Curvas ROC pooled — los 10 modelos")
    plt.legend(fontsize=8, loc="lower right")
    plt.tight_layout()
    plt.savefig(f"{G5}/roc_todos.png", dpi=150)
    plt.close()

    fig, ejes = plt.subplots(2, 5, figsize=(16, 7))
    for m, ax in zip(orden, ejes.flat):
        cm = pd.read_csv(f"{RES}/{m}/confusion_matrix_pooled.csv", index_col=0).values
        ax.imshow(cm, cmap="Blues", vmin=0)
        for i in range(2):
            for j in range(2):
                ax.text(j, i, cm[i, j], ha="center", va="center", color="white" if cm[i, j] > cm.max() / 2 else "black")
        ax.set_xticks([0, 1], ["fresco", "no_fresco"], fontsize=7)
        ax.set_yticks([0, 1], ["fresco", "no_fresco"], fontsize=7)
        ax.set_title(f"{MODELOS[m][1]} ({MODELOS[m][2]})", fontsize=10)
    fig.suptitle("Matrices de confusion — 10 modelos (pooled, 8,000 imagenes)")
    fig.tight_layout()
    fig.savefig(f"{G5}/confmat_todos.png", dpi=150)
    plt.close(fig)

    for m in ("vit_base_patch16_224", "swin_base_patch4_window7_224"):
        r = pd.read_csv(f"{RES}/{m}/roc_curve_pooled.csv")
        plt.figure(figsize=(5, 5))
        plt.fill_between(r.fpr, r.tpr, color="#4C72B0", alpha=.25, label=f"Area = AUC = {pooled.loc[m, 'auc']:.4f}")
        plt.plot(r.fpr, r.tpr, color="#2B4C7E", lw=2)
        plt.plot([0, 1], [0, 1], "--", color="gray", lw=1, label="Modelo al azar (AUC=0.5)")
        plt.xlim(0, 1)
        plt.ylim(0, 1.02)
        plt.xlabel("Tasa de falsos positivos")
        plt.ylabel("Tasa de verdaderos positivos")
        plt.title(f"{MODELOS[m][0]} — Area bajo la curva (AUC)")
        plt.legend(loc="lower right")
        plt.tight_layout()
        plt.savefig(f"{G5}/roc_area_{m}.png", dpi=150)
        plt.close()


def auc_explicacion(pooled):
    m = "vit_base_patch16_224"
    r = pd.read_csv(f"{RES}/{m}/roc_curve_pooled.csv")
    plt.figure(figsize=(6, 6))
    plt.fill_between(r.fpr, r.tpr, color="#4C72B0", alpha=.25, label=f"Area sombreada = AUC = {pooled.loc[m, 'auc']:.4f}")
    plt.plot(r.fpr, r.tpr, color="#2B4C7E", lw=2, label="Curva ROC")
    plt.plot([0, 1], [0, 1], "--", color="gray", lw=1, label="Modelo al azar (AUC=0.5)")
    plt.xlim(0, 1)
    plt.ylim(0, 1.02)
    plt.xlabel("Tasa de falsos positivos (1 - especificidad)")
    plt.ylabel("Tasa de verdaderos positivos (recall)")
    plt.title("Que es el AUC — ejemplo con ViT-Base/16 (pooled, 5 folds)")
    plt.legend(loc="lower right")
    plt.tight_layout()
    plt.savefig(f"{G5}/auc_explicacion.png", dpi=150)
    plt.close()


def main():
    por_fold = pd.read_csv(f"{RES}/metricas_por_fold.csv")
    pooled = pd.read_csv(f"{RES}/metricas_pooled_por_modelo.csv").set_index("model")
    c = curvas()
    for m in ("vit_base_patch16_224", "swin_base_patch4_window7_224", "efficientnet_lite0", "resnet50"):
        epochs(m, c)
    for m in ("vit_base_patch16_224", "swin_base_patch4_window7_224", "efficientnet_lite0"):
        cm_folds(m, por_fold)
        roc_folds(m, por_fold, pooled)
    informe(por_fold, pooled)
    auc_explicacion(pooled)


if __name__ == "__main__":
    main()
