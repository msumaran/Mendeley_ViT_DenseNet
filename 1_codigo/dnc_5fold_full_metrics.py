"""5-fold CV completo sobre DataSetNotClean_4k (8,000 imgs, balanceado),
full fine-tune, replicando el entrenamiento de run_freshness_v3.py pero con
StratifiedKFold(5) en vez de un unico split 70/15/15 fijo, para tener
detalle por fold documentado (misma metodologia del split unico ya
presentado: sin agrupar por foto base).
Uso: dnc_5fold_full_metrics.py <modelo> <fold 1-5>
"""
import os
import sys
import time

import numpy as np
import pandas as pd
import torch
from torch.utils.data import DataLoader
from sklearn.model_selection import StratifiedKFold, train_test_split

from config import CFG
from dataset_freshness import FreshnessDataset
from metrics import compute_metrics
from models import build_model, disable_fused_attention, is_vit
from run_freshness_v3 import (
    V3_CFG, IMAGE_SIZE_CNN, IMAGE_SIZE_VIT,
    get_transforms_v3, FocalLoss, build_optimizer, build_scheduler, collect_dnc_only,
)

MNAME = sys.argv[1]
FOLD = int(sys.argv[2])
NUM_CLASSES = 2
SEED = 42
np.random.seed(SEED)
torch.manual_seed(SEED)

dev = CFG.device


def main():
    """Entrena y evalua UN fold de UN modelo (recibe modelo y fold por argv).

    Flujo: reconstruye el split de 5 folds (StratifiedKFold seed=42) -> separa
    el 20% de test de este fold -> del 80% restante separa 15% para validacion
    -> entrena hasta V3_CFG['num_epochs'] epochs, guardando el checkpoint cada
    vez que val_F1 mejora (early stopping por paciencia) -> recarga el MEJOR
    checkpoint (no el ultimo epoch) y lo evalua UNA vez contra el test set de
    este fold -> guarda esa fila de metricas en un CSV.
    """
    t0 = time.time()
    samples, labels, _ = collect_dnc_only(two_class=True)
    labels = np.array(labels)

    # Mismo seed en los 5 procesos (uno por fold) -> cada uno reproduce el
    # split completo y solo toma el fold que le toca (list()[FOLD - 1]).
    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
    tr_full_idx, te_idx = list(skf.split(samples, labels))[FOLD - 1]
    tr_idx, va_idx = train_test_split(
        tr_full_idx, test_size=0.15, stratify=labels[tr_full_idx], random_state=SEED)

    def subset(idx):
        return [samples[i] for i in idx], [int(labels[i]) for i in idx]

    tr_s, tr_l = subset(tr_idx)
    va_s, va_l = subset(va_idx)
    te_s, te_l = subset(te_idx)

    img_size = IMAGE_SIZE_VIT if is_vit(MNAME) else IMAGE_SIZE_CNN
    bs = V3_CFG["batch_size_vit"] if is_vit(MNAME) else V3_CFG["batch_size_cnn"]
    nw = 2
    kw = dict(batch_size=bs, num_workers=nw, pin_memory=torch.cuda.is_available(),
              persistent_workers=nw > 0, prefetch_factor=2 if nw > 0 else None)
    train_loader = DataLoader(FreshnessDataset(tr_s, tr_l, get_transforms_v3(img_size, True)),
                               shuffle=True, **kw)
    val_loader = DataLoader(FreshnessDataset(va_s, va_l, get_transforms_v3(img_size, False)),
                             shuffle=False, **kw)
    test_loader = DataLoader(FreshnessDataset(te_s, te_l, get_transforms_v3(img_size, False)),
                              shuffle=False, **kw)

    class_counts = np.array([np.sum(np.array(tr_l) == c) for c in range(NUM_CLASSES)])
    print(f"[{MNAME}] fold {FOLD}/5  train={len(tr_s)} val={len(va_s)} test={len(te_s)}  "
          f"class_counts_train={class_counts.tolist()}  device={dev}", flush=True)

    model = build_model(MNAME, NUM_CLASSES, pretrained=True)
    if is_vit(MNAME):
        disable_fused_attention(model)
    model.to(dev)

    # Peso por clase inversamente proporcional a su frecuencia en TRAIN
    # (irrelevante aqui porque DNC esta balanceado 50/50, pero se deja
    # generico por si el split de un fold queda ligeramente desbalanceado).
    raw_w = 1.0 / (class_counts.astype(float) + 1)
    alpha = torch.tensor(raw_w / raw_w.sum() * NUM_CLASSES, dtype=torch.float32).to(dev)
    criterion = FocalLoss(alpha=alpha, gamma=V3_CFG["focal_gamma"],
                           label_smoothing=V3_CFG["label_smoothing"])
    optimizer = build_optimizer(model, MNAME)
    scheduler = build_scheduler(optimizer, V3_CFG["num_epochs"], V3_CFG["warmup_epochs"])

    best_val_f1 = -1.0
    patience = 0
    ckpt_dir = "checkpoints_freshness_v3_4k_5fold"
    os.makedirs(ckpt_dir, exist_ok=True)
    ckpt_path = os.path.join(ckpt_dir, f"{MNAME}_fold{FOLD}_best.pt")

    for epoch in range(V3_CFG["num_epochs"]):
        model.train()
        for imgs, lbls in train_loader:
            imgs, lbls = imgs.to(dev), lbls.to(dev)
            optimizer.zero_grad()
            criterion(model(imgs), lbls).backward()
            torch.nn.utils.clip_grad_norm_(model.parameters(), max_norm=1.0)
            optimizer.step()
        scheduler.step()

        model.eval()
        all_p, all_l = [], []
        with torch.no_grad():
            for imgs, lbls in val_loader:
                all_p.extend(model(imgs.to(dev)).argmax(1).cpu().numpy())
                all_l.extend(lbls.numpy())
        vm = compute_metrics(np.array(all_l), np.array(all_p), NUM_CLASSES)

        if dev == "mps":
            torch.mps.empty_cache()

        # Guarda (sobrescribe) el checkpoint SOLO si este epoch mejora el
        # mejor val_F1 visto hasta ahora en este fold -> al final del
        # entrenamiento, el archivo en disco es el peso del mejor epoch,
        # no el del ultimo.
        if vm["macro_f1"] > best_val_f1:
            best_val_f1 = vm["macro_f1"]
            torch.save(model.state_dict(), ckpt_path)
            patience = 0
        else:
            patience += 1

        print(f"  epoch {epoch + 1:3d}/{V3_CFG['num_epochs']}  val_F1={vm['macro_f1']:.2f}  "
              f"best={best_val_f1:.2f}  pat={patience}  ({time.time() - t0:.0f}s)", flush=True)

        if patience >= V3_CFG["early_stop_patience"]:
            print(f"  Early stop at epoch {epoch + 1}", flush=True)
            break

    # Recarga el MEJOR checkpoint (no el ultimo epoch entrenado) y lo evalua
    # UNA sola vez contra el test set de este fold -> asi se obtiene la fila
    # final reportada para (modelo, fold).
    model.load_state_dict(torch.load(ckpt_path, map_location=dev))
    model.eval()
    all_p, all_l = [], []
    with torch.no_grad():
        for imgs, lbls in test_loader:
            all_p.extend(model(imgs.to(dev)).argmax(1).cpu().numpy())
            all_l.extend(lbls.numpy())
    tm = compute_metrics(np.array(all_l), np.array(all_p), NUM_CLASSES)
    print(f"\n{MNAME} fold {FOLD}/5 TEST -> f1={tm['macro_f1']:.2f} acc={tm['accuracy']:.2f}", flush=True)

    row = {"model": MNAME, "fold": FOLD, "n_test": len(te_s), "best_val_f1": round(best_val_f1, 2)}
    for k in ("accuracy", "macro_precision", "macro_recall", "macro_f1", "macro_specificity"):
        row[k] = round(tm[k], 2)
    out = f"results_freshness_v3/fullmetrics_dnc_{MNAME}_fullfinetune_fold{FOLD}.csv"
    pd.DataFrame([row]).to_csv(out, index=False)
    print(f"Guardado en {out}  (total {time.time() - t0:.0f}s)", flush=True)


if __name__ == "__main__":
    main()
