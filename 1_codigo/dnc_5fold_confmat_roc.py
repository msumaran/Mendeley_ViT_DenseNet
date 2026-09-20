"""Inferencia SOLO (sin re-entrenar) sobre los checkpoints ya guardados del
batch DNC 5-fold (10 modelos x 5 folds, checkpoints_freshness_v3_4k_5fold/),
para obtener matriz de confusion y curva ROC + AUC, ademas de reconfirmar
las 5 metricas ya calculadas (accuracy, precision, recall, F1, especificidad).

Reproduce EXACTAMENTE el mismo split de test que uso dnc_5fold_full_metrics.py
(StratifiedKFold seed=42), asi que el test set de cada fold es identico al
que se uso quando se entreno cada checkpoint.

Para cada modelo genera:
  - resultados_por_fold: 5 filas con accuracy/precision/recall/f1/specificity/auc
  - matriz de confusion por fold + pooled (los 5 folds juntos)
  - curva ROC por fold + pooled (fpr, tpr, thresholds, auc)
  - PNG de la matriz de confusion pooled y la curva ROC pooled

Uso: dnc_5fold_confmat_roc.py [<modelo>]   (sin argumento = todos los 10 modelos)
"""
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import torch
from sklearn.metrics import confusion_matrix, roc_curve, roc_auc_score
from sklearn.model_selection import StratifiedKFold, train_test_split
from torch.utils.data import DataLoader

from config import CFG
from dataset_freshness import FreshnessDataset
from metrics import compute_metrics
from models import build_model, disable_fused_attention, is_vit
from run_freshness_v3 import V3_CFG, IMAGE_SIZE_CNN, IMAGE_SIZE_VIT, get_transforms_v3, collect_dnc_only

SEED = 42
NUM_CLASSES = 2
CLASS_NAMES = ["fresco", "no_fresco"]
CKPT_DIR = "checkpoints_freshness_v3_4k_5fold"
OUT_DIR = "results_freshness_v3/reportes/dnc_5fold_confmat_roc"
ALL_MODELS = [
    "mobilenetv3_small_100", "efficientnet_lite0", "vit_base_patch16_224",
    "efficientnet_b0", "mobilenetv2_100", "densenet121", "resnet50",
    "mobilenetv1_100", "swin_base_patch4_window7_224", "vit_small_patch14_dinov2",
]

dev = CFG.device


def test_split_for_fold(samples, labels, fold):
    """Reconstruye el 20% de test de `fold` con el MISMO seed y split usado
    en el entrenamiento (dnc_5fold_full_metrics.py) -> garantiza que se evalua
    contra las mismas imagenes que nunca vio ese checkpoint."""
    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
    _, te_idx = list(skf.split(samples, labels))[fold - 1]
    return [samples[i] for i in te_idx], [int(labels[i]) for i in te_idx]


@torch.no_grad()
def infer(model, samples, labels, img_size, bs):
    """Corre el modelo (ya en eval()) sobre `samples` y devuelve, para cada
    imagen: la etiqueta real, la clase predicha (argmax) y la probabilidad
    softmax de la clase "no_fresco" (necesaria para AUC/curva ROC, que
    evaluan todos los umbrales, no solo el corte de 50% que usa argmax)."""
    loader = DataLoader(FreshnessDataset(samples, labels, get_transforms_v3(img_size, False)),
                         batch_size=bs, shuffle=False, num_workers=2)
    probs, preds, trues = [], [], []
    for imgs, lbls in loader:
        logits = model(imgs.to(dev))
        p = torch.softmax(logits, dim=1)[:, 1].cpu().numpy()  # prob de "no_fresco"
        probs.extend(p)
        preds.extend(logits.argmax(1).cpu().numpy())
        trues.extend(lbls.numpy())
    return np.array(trues), np.array(preds), np.array(probs)


def plot_confusion_matrix(cm, title, path):
    fig, ax = plt.subplots(figsize=(4, 4))
    im = ax.imshow(cm, cmap="Blues")
    ax.set_xticks([0, 1]); ax.set_xticklabels(CLASS_NAMES)
    ax.set_yticks([0, 1]); ax.set_yticklabels(CLASS_NAMES)
    ax.set_xlabel("Prediccion"); ax.set_ylabel("Real")
    ax.set_title(title, fontsize=10)
    for i in range(2):
        for j in range(2):
            ax.text(j, i, str(cm[i, j]), ha="center", va="center",
                     color="white" if cm[i, j] > cm.max() / 2 else "black")
    fig.colorbar(im)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def plot_roc(fpr, tpr, auc, title, path):
    fig, ax = plt.subplots(figsize=(4, 4))
    ax.plot(fpr, tpr, label=f"AUC = {auc:.3f}")
    ax.plot([0, 1], [0, 1], "--", color="gray")
    ax.set_xlabel("Tasa de falsos positivos")
    ax.set_ylabel("Tasa de verdaderos positivos")
    ax.set_title(title, fontsize=10)
    ax.legend(loc="lower right")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def process_model(model_name, samples, labels, class_counts):
    """Para 1 modelo: recarga los 5 checkpoints ya entrenados (uno por fold),
    evalua cada uno contra el test set de SU fold, guarda matriz de confusion
    + curva ROC por fold, y al final junta (pooled) las 8,000 predicciones de
    los 5 folds -> matriz de confusion pooled + curva ROC pooled + metricas
    finales del modelo. No reentrena nada."""
    img_size = IMAGE_SIZE_VIT if is_vit(model_name) else IMAGE_SIZE_CNN
    bs = V3_CFG["batch_size_vit"] if is_vit(model_name) else V3_CFG["batch_size_cnn"]

    model_dir = os.path.join(OUT_DIR, model_name)
    os.makedirs(model_dir, exist_ok=True)

    fold_rows = []
    all_trues, all_preds, all_probs = [], [], []

    net = build_model(model_name, NUM_CLASSES, pretrained=False)
    if is_vit(model_name):
        disable_fused_attention(net)
    net.to(dev)

    for fold in range(1, 6):
        te_s, te_l = test_split_for_fold(samples, labels, fold)
        ckpt = os.path.join(CKPT_DIR, f"{model_name}_fold{fold}_best.pt")
        net.load_state_dict(torch.load(ckpt, map_location=dev))
        net.eval()

        trues, preds, probs = infer(net, te_s, te_l, img_size, bs)
        m = compute_metrics(trues, preds, NUM_CLASSES, probs=probs)
        auc = m["auc"]
        cm = confusion_matrix(trues, preds, labels=[0, 1])
        fpr, tpr, thr = roc_curve(trues, probs)

        pd.DataFrame(cm, index=CLASS_NAMES, columns=CLASS_NAMES).to_csv(
            os.path.join(model_dir, f"confusion_matrix_fold{fold}.csv"))
        pd.DataFrame({"fpr": fpr, "tpr": tpr, "threshold": thr}).to_csv(
            os.path.join(model_dir, f"roc_curve_fold{fold}.csv"), index=False)

        fold_rows.append({
            "model": model_name, "fold": fold, "n_test": len(te_s),
            "accuracy": round(m["accuracy"], 2), "precision": round(m["macro_precision"], 2),
            "recall": round(m["macro_recall"], 2), "f1": round(m["macro_f1"], 2),
            "specificity": round(m["macro_specificity"], 2), "auc": round(auc, 4),
        })
        print(f"  {model_name} fold {fold}/5: acc={m['accuracy']:.2f} f1={m['macro_f1']:.2f} "
              f"auc={auc:.4f}", flush=True)

        all_trues.append(trues); all_preds.append(preds); all_probs.append(probs)

    # pooled = concatenar las predicciones de los 5 folds (8,000 imagenes,
    # cada una evaluada 1 sola vez por el checkpoint que NUNCA la vio en
    # entrenamiento) y calcular cada metrica UNA vez sobre el total, en vez
    # de promediar 5 metricas por separado. Equivale a sumar las 5 matrices
    # de confusion celda a celda, porque los folds son disjuntos.
    trues = np.concatenate(all_trues); preds = np.concatenate(all_preds); probs = np.concatenate(all_probs)
    m = compute_metrics(trues, preds, NUM_CLASSES, probs=probs)
    auc = m["auc"]
    cm = confusion_matrix(trues, preds, labels=[0, 1])
    fpr, tpr, thr = roc_curve(trues, probs)

    pd.DataFrame(cm, index=CLASS_NAMES, columns=CLASS_NAMES).to_csv(
        os.path.join(model_dir, "confusion_matrix_pooled.csv"))
    pd.DataFrame({"fpr": fpr, "tpr": tpr, "threshold": thr}).to_csv(
        os.path.join(model_dir, "roc_curve_pooled.csv"), index=False)
    plot_confusion_matrix(cm, f"{model_name} - Matriz de confusion (pooled 5 folds)",
                            os.path.join(model_dir, "confusion_matrix_pooled.png"))
    plot_roc(fpr, tpr, auc, f"{model_name} - Curva ROC (pooled 5 folds)",
              os.path.join(model_dir, "roc_curve_pooled.png"))

    pooled_row = {
        "model": model_name, "fold": "pooled", "n_test": len(trues),
        "accuracy": round(m["accuracy"], 2), "precision": round(m["macro_precision"], 2),
        "recall": round(m["macro_recall"], 2), "f1": round(m["macro_f1"], 2),
        "specificity": round(m["macro_specificity"], 2), "auc": round(auc, 4),
    }
    print(f"  {model_name} POOLED: acc={m['accuracy']:.2f} f1={m['macro_f1']:.2f} "
          f"auc={auc:.4f}", flush=True)

    return fold_rows, pooled_row


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    models = [sys.argv[1]] if len(sys.argv) > 1 else ALL_MODELS

    samples, labels, _ = collect_dnc_only(two_class=True)
    labels = np.array(labels)
    class_counts = np.bincount(labels)

    all_fold_rows, all_pooled_rows = [], []
    for model_name in models:
        print(f"=== {model_name} ===", flush=True)
        fold_rows, pooled_row = process_model(model_name, samples, labels, class_counts)
        all_fold_rows.extend(fold_rows)
        all_pooled_rows.append(pooled_row)

    pd.DataFrame(all_fold_rows).to_csv(os.path.join(OUT_DIR, "metricas_por_fold.csv"), index=False)
    pd.DataFrame(all_pooled_rows).to_csv(os.path.join(OUT_DIR, "metricas_pooled_por_modelo.csv"), index=False)
    print(f"\nGuardado en {OUT_DIR}/", flush=True)


if __name__ == "__main__":
    main()
