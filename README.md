# Experimento Mendeley (DNC) — 10 modelos, selección reducida a 2

Entrenamiento y evaluación de 10 arquitecturas sobre el dataset DNC (DataSetNotClean,
Mendeley) — 8,000 imágenes balanceadas (4,000 fresco / 4,000 no_fresco), todas las
especies. Es el dominio **fuente** de todo el proyecto (los checkpoints resultantes se
usan como punto de partida en los experimentos cross-dataset de Trucha/Rohu).

## Selección de modelos — cambio respecto a la versión anterior

El proyecto trabajaba con **4 modelos** elegidos de Mendeley: ViT-Base/16, Swin-Base,
DenseNet-121 y EfficientNet-B0. Este repositorio documenta la reducción de esa selección
a **2 modelos: ViT-Base/16 y DenseNet-121** — se descartan Swin-Base y EfficientNet-B0 de
la selección activa del proyecto.

Los 10 modelos y sus resultados completos se conservan sin cambios en todas las carpetas
de este repositorio (ningún número se recalculó ni se volvió a entrenar) — se conserva
todo el experimento de los 10 modelos por trazabilidad, y el documento
[`Tabla_Mendeley_10_Modelos_ViT_DenseNet.docx`](Tabla_Mendeley_10_Modelos_ViT_DenseNet.docx)
en la raíz es la versión actualizada de la tabla comparativa con solo ViT-Base/16 y
DenseNet-121 resaltados como elegidos.

Justificación cuantitativa de por qué estos 2 (F1, AUC, estabilidad, eficiencia) en la
sección "Por qué ViT-Base/16 y DenseNet-121" de ese mismo documento, y en detalle en
[`6_documentacion/Sustento_Seleccion_Modelos.docx`](6_documentacion/Sustento_Seleccion_Modelos.docx)
(actualizado a la selección vigente de 2 modelos, con los mismos criterios).

## Contenido

| Carpeta | Contenido |
|---|---|
| `1_codigo/` | `config.py`, `dataset_freshness.py`, `dnc_5fold_confmat_roc.py`, `dnc_5fold_full_metrics.py`, `dnc_graficas_organizadas.py`, `metrics.py`, `models.py`, `run_freshness_v3.py` — utilidades compartidas y scripts de entrenamiento/evaluación de los 10 modelos |
| `2_logs_entrenamiento/` | Logs de la corrida de entrenamiento (5-fold) y de la generación de matrices de confusión/ROC de los 10 modelos |
| `3_resultados_test_por_modelo/` | Matrices de confusión y curvas ROC (por fold y pooled) de cada uno de los 10 modelos, más `metricas_pooled_por_modelo.csv` y `metricas_por_fold.csv` |
| `4_graficas/` | Gráficas agregadas: área bajo curva, curvas ROC, matrices de confusión de los 10 modelos |
| `5_graficas_informe/` | Gráficas usadas en el informe (AUC, F1 y estabilidad, parámetros vs. F1, precisión vs. recall) |
| `6_documentacion/` | `Documentacion_Tecnica_10_Modelos.docx` (detalle completo por epoch/fold) y `Sustento_Seleccion_Modelos.docx` (criterios de selección de los 2 modelos elegidos) |
| `7_graficas_por_modelo/` | Matrices de confusión, curvas ROC y curvas de épocas, una por modelo, para los 10 modelos |

## Entrenamiento y evaluación

```bash
# Entrenamiento (5-fold, 30 épocas, early_stop_patience=7), los 10 modelos
python run_freshness_v3.py

# Evaluación completa por fold
python dnc_5fold_full_metrics.py

# Matriz de confusión + curva ROC + AUC, por fold y pooled (solo inferencia,
# no reentrena -- recarga los 50 checkpoints ya entrenados, 10 modelos x 5 folds)
python dnc_5fold_confmat_roc.py

# Gráficas finales
python dnc_graficas_organizadas.py
```

`metrics.py` incluye además funciones de "Dimensión 2 — Robustez ante el Cambio de
Dominio" (`tasa_retencion_desempeno`, `brecha_error`, `reduccion_brecha_error`) que **no**
se usan en este experimento — pertenecen a los experimentos cross-dataset (Trucha/Rohu),
que comparten este mismo archivo de utilidades por ser un solo proyecto de tesis.
