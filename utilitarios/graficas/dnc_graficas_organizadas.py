"""Genera y organiza en carpetas las graficas de los 10 modelos DNC 5-fold:
curvas ROC, matrices de confusion, y area bajo la curva (AUC sombreada).

No corre inferencia ni reentrena nada: lee los CSV ya calculados por
dnc_5fold_confmat_roc.py (roc_curve_pooled.csv y confusion_matrix_pooled.csv
por modelo) y solo genera/organiza las imagenes.

Uso: dnc_graficas_organizadas.py
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

IN_DIR = "results_freshness_v3/reportes/dnc_5fold_confmat_roc"
OUT_DIR = "results_freshness_v3/reportes/graficas"

NAMES = {
    "mobilenetv3_small_100": "MobileNetV3-Small", "efficientnet_lite0": "EfficientNet-Lite0",
    "vit_base_patch16_224": "ViT-Base-16", "efficientnet_b0": "EfficientNet-B0",
    "mobilenetv2_100": "MobileNetV2", "densenet121": "DenseNet-121", "resnet50": "ResNet-50",
    "mobilenetv1_100": "MobileNetV1", "swin_base_patch4_window7_224": "Swin-Base",
    "vit_small_patch14_dinov2": "DINOv2-Small",
}
CLASS_NAMES = ["fresco", "no_fresco"]


def plot_roc(fpr, tpr, auc, title, path):
    """Curva ROC simple (tasa de verdaderos positivos vs falsos positivos),
    AUC en escala 0-1 (no porcentaje) en la leyenda."""
    fig, ax = plt.subplots(figsize=(5, 5))
    ax.plot(fpr, tpr, color="#2A4D7A", linewidth=2, label=f"AUC = {auc:.4f}")
    ax.plot([0, 1], [0, 1], "--", color="gray", linewidth=1)
    ax.set_xlabel("Tasa de falsos positivos")
    ax.set_ylabel("Tasa de verdaderos positivos")
    ax.set_title(title, fontsize=11)
    ax.legend(loc="lower right")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def plot_confusion_matrix(cm_df, title, path):
    """Heatmap 2x2 (fresco/no_fresco) a partir de la matriz de confusion
    pooled ya calculada y guardada en confusion_matrix_pooled.csv."""
    cm = cm_df.values
    fig, ax = plt.subplots(figsize=(5, 5))
    im = ax.imshow(cm, cmap="Blues")
    ax.set_xticks([0, 1]); ax.set_xticklabels(CLASS_NAMES)
    ax.set_yticks([0, 1]); ax.set_yticklabels(CLASS_NAMES)
    ax.set_xlabel("Prediccion"); ax.set_ylabel("Real")
    ax.set_title(title, fontsize=11)
    for i in range(2):
        for j in range(2):
            ax.text(j, i, str(cm[i, j]), ha="center", va="center",
                     color="white" if cm[i, j] > cm.max() / 2 else "black")
    fig.colorbar(im)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def plot_area_bajo_curva(fpr, tpr, auc, title, path):
    """La curva ROC con el area (AUC) sombreada, para explicar visualmente que es el AUC."""
    fig, ax = plt.subplots(figsize=(5, 5))
    ax.fill_between(fpr, tpr, alpha=0.25, color="#4C72B0", label=f"Area = AUC = {auc:.4f}")
    ax.plot(fpr, tpr, color="#2A4D7A", linewidth=2)
    ax.plot([0, 1], [0, 1], "--", color="gray", linewidth=1, label="Modelo al azar (AUC=0.5)")
    ax.set_xlabel("Tasa de falsos positivos")
    ax.set_ylabel("Tasa de verdaderos positivos")
    ax.set_title(title, fontsize=11)
    ax.legend(loc="lower right", fontsize=9)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1.02)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def main():
    pooled = pd.read_csv(os.path.join(IN_DIR, "metricas_pooled_por_modelo.csv")).set_index("model")

    dir_roc = os.path.join(OUT_DIR, "curvas_roc")
    dir_cm = os.path.join(OUT_DIR, "matrices_confusion")
    dir_auc = os.path.join(OUT_DIR, "area_bajo_curva")
    for d in (dir_roc, dir_cm, dir_auc):
        os.makedirs(d, exist_ok=True)

    for model, nombre in NAMES.items():
        model_dir = os.path.join(IN_DIR, model)
        auc = pooled.loc[model, "auc"]

        roc = pd.read_csv(os.path.join(model_dir, "roc_curve_pooled.csv"))
        plot_roc(roc["fpr"], roc["tpr"], auc, f"{nombre} - Curva ROC (pooled, 5 folds)",
                 os.path.join(dir_roc, f"{model}.png"))

        cm = pd.read_csv(os.path.join(model_dir, "confusion_matrix_pooled.csv"), index_col=0)
        plot_confusion_matrix(cm, f"{nombre} - Matriz de confusion (pooled, 5 folds)",
                               os.path.join(dir_cm, f"{model}.png"))

        plot_area_bajo_curva(roc["fpr"], roc["tpr"], auc, f"{nombre} - Area bajo la curva (AUC)",
                              os.path.join(dir_auc, f"{model}.png"))

        print(f"{nombre}: AUC={auc:.4f}  -> guardado en curvas_roc/, matrices_confusion/, area_bajo_curva/")

    print(f"\nListo. Carpetas generadas en {OUT_DIR}/")


if __name__ == "__main__":
    main()
