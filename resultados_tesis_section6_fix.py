# -*- coding: utf-8 -*-
"""
Resultados Tesis – Python (Anaconda + VS Code)
------------------------------------------------
Genera tablas A, B, C, C2, D (OLS), E (GLM Binomial), F y F2 con p-values a 4 decimales.
Segmentación: 15/15/21/12 → G1,G2=Control (0), G3,G4=Tratamiento (1).

Requisitos: pandas, numpy, statsmodels, openpyxl
Instalar:  pip install pandas numpy statsmodels openpyxl
Ejecutar:  python resultados_tesis.py
"""

import sys
import re
from pathlib import Path
import numpy as np
import pandas as pd

# ---------- Configura aquí tu Excel ----------
# Cambia esta ruta si tu archivo está en otro lugar o tiene otro nombre.
EXCEL_PATH = r"D:\LOGAN\TIE\RESULTADOS.xlsx"  # intenta primero este
EXCEL_FALLBACKS = [
    r"D:\LOGAN\TIE\RESUTADOS.xlsx",           # fallback por el typo común
    r"D:\LOGAN\TIE\Combinación_Resultados.xlsx",
]

SHEET_PREFERRED = "Hoja1"  # si no existe, se usará la primera hoja

# ---------- Utilidades de nombres/columnas ----------
def canon(col: str) -> str:
    s = re.sub(r"[^0-9a-zA-Z]+", "_", str(col)).strip("_").lower()
    s = re.sub(r"_+", "_", s)
    return s

def find_col(df: pd.DataFrame, subs, prefer=None):
    """Busca una columna (en df.columns ya canonicalizadas) que contenga todas las subcadenas."""
    req = [s.lower() for s in subs]
    cand = [c for c in df.columns if all(s in c for s in req)]
    if not cand:
        return None
    if prefer:
        prefer = prefer.lower()
        for c in cand:
            if prefer in c:
                return c
    return cand[0]

def build_ok(df, ok_col, resp_col, idx_col, name):
    """Construye 1/0 de acierto por ítem usando *_ok==1 o comparando RESP vs IDX."""
    if ok_col is not None and ok_col in df.columns:
        x = pd.to_numeric(df[ok_col], errors="coerce")
        return (x == 1).astype(float).rename(name)
    if (resp_col is not None) and (idx_col is not None) and \
       (resp_col in df.columns) and (idx_col in df.columns):
        return (df[resp_col].astype(str) == df[idx_col].astype(str)).astype(float).rename(name)
    return pd.Series(np.nan, index=df.index, name=name)

def p4(x):
    return f"{x:.4f}" if pd.notnull(x) else ""

def fmt(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_float_dtype(out[c]):
            out[c] = out[c].astype(float).round(4)
    if "p" in out.columns:
        out["p"] = out["p"].map(p4)
    return out

def save_table(df: pd.DataFrame, name: str, folder: Path):
    folder.mkdir(parents=True, exist_ok=True)
    df_out = fmt(df)
    df_out.to_csv(folder / f"{name}.csv", index=False, encoding="utf-8-sig")
    return df_out

def append_html(html_parts: list, title: str, df: pd.DataFrame):
    df2 = fmt(df)
    html_parts.append(f"<h3>{title}</h3>\n{df2.to_html(index=False)}")

# ---------- Carga del Excel ----------
def load_excel() -> pd.DataFrame:
    paths = [Path(EXCEL_PATH)] + [Path(p) for p in EXCEL_FALLBACKS]
    chosen = None
    for p in paths:
        if p.exists():
            chosen = p
            break
    if chosen is None:
        print("No se encontró el Excel en las rutas predefinidas.")
        print("Edita EXCEL_PATH arriba o pásalo como argumento: python resultados_tesis.py D:\\ruta\\archivo.xlsx")
        if len(sys.argv) >= 2:
            p = Path(sys.argv[1])
            if not p.exists():
                raise FileNotFoundError(f"Ruta no válida: {p}")
            chosen = p
        else:
            raise FileNotFoundError("No se encontró archivo Excel.")
    try:
        xl = pd.ExcelFile(chosen)
    except Exception as e:
        raise RuntimeError(f"No pude abrir el Excel ({chosen}): {e}")

    sheet = SHEET_PREFERRED if SHEET_PREFERRED in xl.sheet_names else xl.sheet_names[0]
    df_raw = xl.parse(sheet_name=sheet)
    print(f"[INFO] Archivo: {chosen.name} | Hoja: {sheet} | Filas: {len(df_raw)} | Cols: {df_raw.shape[1]}")
    return df_raw, chosen.name, sheet

# ---------- Modelos ----------
def sanitize_for_model(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce nullable integers to vanilla float dtypes before modeling."""

    sanitized = df.copy()
    numeric_cols = ["grupo_num", "tratamiento", "nitems_row", "aciertos_total"]
    for column in numeric_cols:
        if column in sanitized.columns:
            sanitized[column] = pd.to_numeric(sanitized[column], errors="coerce")

    if "grupo_num" in sanitized.columns:
        sanitized["grupo_num"] = sanitized["grupo_num"].astype("float64")
    if "tratamiento" in sanitized.columns:
        sanitized["tratamiento"] = sanitized["tratamiento"].astype("float64")
    if "nitems_row" in sanitized.columns:
        sanitized["nitems_row"] = sanitized["nitems_row"].astype("float64")
    if "aciertos_total" in sanitized.columns:
        sanitized["aciertos_total"] = sanitized["aciertos_total"].astype("float64")

    return sanitized

def run_models(valid: pd.DataFrame):
    try:
        import statsmodels.api as sm
        import statsmodels.formula.api as smf
    except Exception as e:
        print("[WARN] statsmodels no disponible. Solo se generarán tablas descriptivas.")
        return None, None, None, None

    if valid["grupo_num"].nunique() < 2:
        print("[WARN] Menos de 2 grupos válidos. Se omiten modelos.")
        return None, None, None, None

    model_df = sanitize_for_model(valid)

    # OLS (HC1)
    ols = smf.ols("tasa_acierto ~ C(grupo_num)", data=model_df).fit(cov_type="HC1")
    ci = ols.conf_int()
    ols_tbl = pd.DataFrame({
        "term": ols.params.index,
        "coef": ols.params.values,
        "se": ols.bse.values,
        "t": ols.tvalues.values,
        "p": ols.pvalues.values,
        "ci_low": ci[0].values,
        "ci_high": ci[1].values,
        "nota": "OLS (HC1)"
    })

    # OLS (cluster por grupo) – con 4 clústeres es frágil, pero lo dejamos
    try:
        ols_cl = smf.ols("tasa_acierto ~ C(grupo_num)", data=model_df).fit(
            cov_type="cluster", cov_kwds={"groups": model_df["grupo_num"]}
        )
        ci2 = ols_cl.conf_int()
        ols_tbl2 = pd.DataFrame({
            "term": ols_cl.params.index,
            "coef": ols_cl.params.values,
            "se": ols_cl.bse.values,
            "t": ols_cl.tvalues.values,
            "p": ols_cl.pvalues.values,
            "ci_low": ci2[0].values,
            "ci_high": ci2[1].values,
            "nota": "OLS (VCE cluster por grupo)"
        })
        ols_tbl = pd.concat([ols_tbl, ols_tbl2], ignore_index=True)
    except Exception:
        pass

    # GLM Binomial (HC1) con pesos = # ítems
    glm = smf.glm("tasa_acierto ~ C(grupo_num)", data=model_df,
                  family=sm.families.Binomial()).fit(
                      cov_type="HC1", freq_weights=model_df["nitems_row"]
                  )
    ci = glm.conf_int()
    glm_tbl = pd.DataFrame({
        "term": glm.params.index,
        "coef": glm.params.values,
        "se": glm.bse.values,
        "z/t": glm.tvalues.values,
        "p": glm.pvalues.values,
        "ci_low": ci[0].values,
        "ci_high": ci[1].values,
        "nota": "GLM Binomial (HC1) + pesos"
    })

    # GLM (cluster por grupo)
    try:
        glm_cl = smf.glm("tasa_acierto ~ C(grupo_num)", data=model_df,
                         family=sm.families.Binomial()).fit(
                             cov_type="cluster", cov_kwds={"groups": model_df["grupo_num"]},
                             freq_weights=model_df["nitems_row"]
                         )
        ci2 = glm_cl.conf_int()
        glm_tbl2 = pd.DataFrame({
            "term": glm_cl.params.index,
            "coef": glm_cl.params.values,
            "se": glm_cl.bse.values,
            "z/t": glm_cl.tvalues.values,
            "p": glm_cl.pvalues.values,
            "ci_low": ci2[0].values,
            "ci_high": ci2[1].values,
            "nota": "GLM Binomial (cluster) + pesos"
        })
        glm_tbl = pd.concat([glm_tbl, glm_tbl2], ignore_index=True)
    except Exception:
        pass

    # Predicciones por grupo
    grupos = np.sort(model_df["grupo_num"].unique())
    grid = pd.DataFrame({"grupo_num": grupos})
    po = ols.get_prediction(grid)
    pred_ols = pd.DataFrame({
        "grupo": grid["grupo_num"].astype(int),
        "pred": po.predicted_mean,
        "se": po.se_mean,
        "ll": po.conf_int()[:, 0],
        "ul": po.conf_int()[:, 1],
    })

    pg = glm.get_prediction(grid)
    pred_glm = pd.DataFrame({
        "grupo": grid["grupo_num"].astype(int),
        "pred": pg.predicted_mean,
        "ll": pg.conf_int()[:, 0],
        "ul": pg.conf_int()[:, 1],
    })

    return ols_tbl, glm_tbl, pred_ols, pred_glm

# ---------- Pipeline principal ----------
def main():
    outdir = Path("./salida")
    html_parts = []

    # 1) Cargar Excel
    df_raw, archivo, hoja = load_excel()

    # 2) Canonicalizar columnas y asignar grupos
    df = df_raw.copy()
    df.columns = [canon(c) for c in df.columns]

    n = len(df)
    sizes = [15, 15, 21, 12]   # segmentos
    labels = [1, 2, 3, 4]
    cum = np.cumsum([0] + sizes)  # [0,15,30,51,63]
    grp = np.full(n, np.nan)
    for g in range(4):
        a, b = cum[g], cum[g+1]
        if a < n:
            grp[a:min(b, n)] = labels[g]
    df["grupo_num"] = pd.Series(grp).astype("Int64")
    df["tratamiento"] = df["grupo_num"].map({1:0, 2:0, 3:1, 4:1}).astype("Int64")

    # 3) Detectar columnas y construir gk_ok1..gk_ok4
    r1 = find_col(df, ["gk","1","resp"]) or find_col(df, ["1","resp"])
    r2 = find_col(df, ["gk","2","resp"]) or find_col(df, ["2","resp"])
    r3 = find_col(df, ["gk","3","resp"]) or find_col(df, ["3","resp"])
    r4 = find_col(df, ["gk","4","resp"]) or find_col(df, ["4","resp"])

    i1 = find_col(df, ["gk","idx","1"]) or find_col(df, ["idx","1"])
    i2 = find_col(df, ["gk","idx","2"]) or find_col(df, ["idx","2"])
    i3 = find_col(df, ["gk","idx","3"]) or find_col(df, ["idx","3"])
    i4 = find_col(df, ["gk","idx","4"]) or find_col(df, ["idx","4"])

    o1 = find_col(df, ["gk","1","ok"]) or find_col(df, ["1","ok"])
    o2 = find_col(df, ["gk","2","ok"]) or find_col(df, ["2","ok"])
    o3 = find_col(df, ["gk","3","ok"]) or find_col(df, ["3","ok"])
    o4 = find_col(df, ["gk","4","ok"]) or find_col(df, ["4","ok"])

    gk_ok1 = build_ok(df, o1, r1, i1, "gk_ok1")
    gk_ok2 = build_ok(df, o2, r2, i2, "gk_ok2")
    gk_ok3 = build_ok(df, o3, r3, i3, "gk_ok3")
    gk_ok4 = build_ok(df, o4, r4, i4, "gk_ok4")
    ok_df = pd.concat([gk_ok1, gk_ok2, gk_ok3, gk_ok4], axis=1)

    # 4) Métricas por fila
    df["nitems_row"]     = ok_df.notna().sum(axis=1)
    df["aciertos_total"] = ok_df.sum(axis=1, skipna=True)
    df["tasa_acierto"]   = df["aciertos_total"] / df["nitems_row"].replace(0, np.nan)

    # 5) Tablas descriptivas
    cuadro_A = pd.DataFrame({
        "archivo": [archivo],
        "hoja": [hoja],
        "N_total": [len(df)],
        "N_validos": [(df["nitems_row"] > 0).sum()],
        "items_prom": [df["nitems_row"].mean()],
        "aciertos_prom": [df["aciertos_total"].mean()],
        "tasa_acierto_prom": [df["tasa_acierto"].mean()],
        "tasa_acierto_sd": [df["tasa_acierto"].std()],
    })
    print("\n=== Cuadro A — Resumen general ===")
    print(fmt(cuadro_A).to_string(index=False))
    save_table(cuadro_A, "cuadro_A", outdir); append_html(html_parts, "Cuadro A — Resumen general", cuadro_A)

    per_item = []
    for j, col in enumerate(["gk_ok1","gk_ok2","gk_ok3","gk_ok4"], start=1):
        if col in ok_df.columns:
            per_item.append({
                "item": j,
                "n_disponible": int(ok_df[col].notna().sum()),
                "tasa_acierto_item": ok_df[col].mean(skipna=True)
            })
    cuadro_B = pd.DataFrame(per_item)
    print("\n=== Cuadro B — Exactitud por ítem ===")
    print(fmt(cuadro_B).to_string(index=False))
    save_table(cuadro_B, "cuadro_B", outdir); append_html(html_parts, "Cuadro B — Exactitud por ítem", cuadro_B)

    valid = df.loc[(df["nitems_row"]>0) & (df["grupo_num"].notna())].copy()

    cuadro_C = (
        valid.groupby("grupo_num", dropna=False)
             .agg(mean_tasa=("tasa_acierto","mean"),
                  sd_tasa=("tasa_acierto","std"),
                  mean_hits=("aciertos_total","mean"),
                  N=("tasa_acierto","size"))
             .reset_index().rename(columns={"grupo_num":"grupo"})
    )
    print("\n=== Cuadro C — Resumen por grupo ===")
    print(fmt(cuadro_C).to_string(index=False))
    save_table(cuadro_C, "cuadro_C", outdir); append_html(html_parts, "Cuadro C — Resumen por grupo", cuadro_C)

    cuadro_C2 = (
        valid.groupby("tratamiento", dropna=False)
             .agg(mean_tasa=("tasa_acierto","mean"),
                  sd_tasa=("tasa_acierto","std"),
                  mean_hits=("aciertos_total","mean"),
                  N=("tasa_acierto","size"))
             .reset_index()
    )
    cuadro_C2["etiqueta"] = cuadro_C2["tratamiento"].map({0:"Control", 1:"Tratamiento"})
    cuadro_C2 = cuadro_C2[["etiqueta","mean_tasa","sd_tasa","mean_hits","N"]]
    print("\n=== Cuadro C2 — Resumen por tratamiento (Control vs Tratamiento) ===")
    print(fmt(cuadro_C2).to_string(index=False))
    save_table(cuadro_C2, "cuadro_C2", outdir); append_html(html_parts, "Cuadro C2 — Resumen por tratamiento", cuadro_C2)

    # 6) Modelos y predicciones
    ols_tbl, glm_tbl, pred_ols, pred_glm = run_models(valid)

    if ols_tbl is not None:
        print("\n=== Cuadro D — OLS por grupo (HC1 y/o cluster) ===")
        print(fmt(ols_tbl).to_string(index=False))
        save_table(ols_tbl, "cuadro_D_OLS", outdir)
        append_html(html_parts, "Cuadro D — OLS por grupo", ols_tbl)

    if glm_tbl is not None:
        print("\n=== Cuadro E — GLM Binomial por grupo (HC1 y/o cluster) ===")
        print(fmt(glm_tbl).to_string(index=False))
        save_table(glm_tbl, "cuadro_E_GLM", outdir)
        append_html(html_parts, "Cuadro E — GLM Binomial por grupo", glm_tbl)

    if pred_ols is not None:
        print("\n=== Cuadro F — Medias predichas por grupo (OLS/HC1) ===")
        print(fmt(pred_ols).to_string(index=False))
        save_table(pred_ols, "cuadro_F_pred_OLS", outdir)
        append_html(html_parts, "Cuadro F — Medias predichas por grupo (OLS)", pred_ols)

    if pred_glm is not None:
        print("\n=== Cuadro F2 — Probabilidades predichas por grupo (GLM/HC1) ===")
        print(fmt(pred_glm).to_string(index=False))
        save_table(pred_glm, "cuadro_F2_pred_GLM", outdir)
        append_html(html_parts, "Cuadro F2 — Probabilidades predichas por grupo (GLM)", pred_glm)

    # 7) Guardar HTML consolidado
    html = """<html><head><meta charset="utf-8"><style>
    body { font-family: Arial, sans-serif; margin: 24px; }
    table { border-collapse: collapse; }
    th, td { border: 1px solid #999; padding: 6px 8px; }
    h3 { margin-top: 28px; }
    </style></head><body>
    <h2>Resultados – Tablas (Python)</h2>
    """ + "\n".join(html_parts) + "\n</body></html>"
    (outdir / "tablas_resultados.html").write_text(html, encoding="utf-8")
    print(f"\n[LISTO] Tablas guardadas en: {outdir.resolve()}")
    print("        - CSVs de cada cuadro")
    print("        - tablas_resultados.html (abrir en navegador)")

if __name__ == "__main__":
    main()
