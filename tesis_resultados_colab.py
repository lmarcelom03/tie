"""tesis_resultados_colab.py

Python translation of the Stata thesis results workflow so it can be run in
Google Colab.  The script mirrors the data preparation steps, descriptive
statistics, hypothesis tests, regression models, and output artifacts produced
by the Stata program.

Usage inside Google Colab
-------------------------
1. Mount Google Drive:
       from google.colab import drive
       drive.mount('/content/drive')

2. Adjust BASE_PATH below if your Excel workbook lives elsewhere in Drive.

3. Run this script:
       %run /content/drive/MyDrive/tie/tesis_resultados_colab.py

The script expects the workbook ``RESUTADOS.xlsx`` (note the missing ``L``)
inside ``BASE_PATH``.  Outputs are written to ``BASE_PATH / 'salidas_python'``.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence

import numpy as np
import pandas as pd
import seaborn as sns
import statsmodels.formula.api as smf
from matplotlib import pyplot as plt
from statsmodels.iolib.summary2 import summary_col
from statsmodels.stats.weightstats import ttest_ind

# ---------------------- CONFIGURATION ---------------------------------------
BASE_PATH = Path("/content/drive/MyDrive/LOGAN/TIE")
WORKBOOK_NAME = "RESUTADOS.xlsx"
PREFERRED_SHEETS: Sequence[str] = ("Hoja1", "Hoja 1", "Sheet1", "Sheet 1")
OUTPUT_DIR = BASE_PATH / "salidas_python"
LOG_FILE = OUTPUT_DIR / "tesis_resultados_colab.log"

# ---------------------- LOGGING --------------------------------------------

def configure_logging(log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, mode="w", encoding="utf-8"),
            logging.StreamHandler()
        ]
    )


# ---------------------- DATA LOADING ---------------------------------------

def discover_sheet(workbook: Path, preferred: Sequence[str]) -> str:
    """Return the first sheet that matches the preferred list or fallback to the
    workbook's first sheet."""
    xls = pd.ExcelFile(workbook)
    logging.info("Hojas disponibles: %s", ", ".join(xls.sheet_names))
    for name in preferred:
        if name in xls.sheet_names:
            logging.info("Usando hoja preferida: %s", name)
            return name
    logging.warning("Ninguna hoja preferida encontrada; usando la primera hoja: %s", xls.sheet_names[0])
    return xls.sheet_names[0]


def read_workbook(base_path: Path, workbook_name: str) -> pd.DataFrame:
    workbook_path = base_path / workbook_name
    if not workbook_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo {workbook_path}.")
    sheet = discover_sheet(workbook_path, PREFERRED_SHEETS)
    logging.info("Importando Excel: %s (hoja=%s)", workbook_path, sheet)
    df = pd.read_excel(workbook_path, sheet_name=sheet)
    logging.info("Importación exitosa con %d filas y %d columnas.", df.shape[0], df.shape[1])
    return df


# ---------------------- DATA PREPARATION -----------------------------------

def ensure_column(df: pd.DataFrame, current: str, alias: str) -> pd.DataFrame:
    if current in df.columns and alias not in df.columns:
        df = df.rename(columns={current: alias})
    return df


def build_group_label(df: pd.DataFrame) -> pd.Series:
    if "Unnamed: 0" in df.columns:
        label = df["Unnamed: 0"].astype("string")
    else:
        label = pd.Series(["" for _ in range(len(df))], index=df.index, dtype="string")
        string_columns = [c for c in df.columns if pd.api.types.is_string_dtype(df[c])]
        for col in string_columns:
            candidate = df[col].astype("string")
            mask = candidate.str.lower().str.contains("grupo|trat|control", na=False)
            label = label.mask((label == "") & mask, candidate)
    label = label.replace({"": pd.NA}).ffill()
    return label


def build_treatment_indicator(label: pd.Series) -> pd.Series:
    lower = label.str.lower()
    treat = pd.Series(np.nan, index=label.index)
    treat = treat.mask(lower.str.contains("trat"), 1)
    treat = treat.mask(lower.str.contains("control"), 0)
    return treat.astype("float")


def to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def pattern_columns(df: pd.DataFrame, suffix: str) -> List[str]:
    suffix = suffix.lower()
    cols = [
        col for col in df.columns
        if "p_blue" in col.lower() and col.lower().endswith(f"_{suffix}")
    ]
    return cols


def compute_row_mean(df: pd.DataFrame, cols: Sequence[str], name: str) -> pd.Series:
    if not cols:
        logging.warning("No se encontraron columnas para %s; se llenará con NaN.", name)
        return pd.Series(np.nan, index=df.index, name=name)
    return df[list(cols)].apply(pd.to_numeric, errors="coerce").mean(axis=1, skipna=True)


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if "participant_id_in_session" not in df.columns:
        raise KeyError("Falta la columna 'participant_id_in_session' en el Excel.")
    df = df.loc[~df["participant_id_in_session"].isna()].copy()
    df["group_label"] = build_group_label(df)
    df["treat"] = build_treatment_indicator(df["group_label"])

    df = ensure_column(df, "TESIS_TOTAL_C_1_player_edad", "edad")
    df = ensure_column(df, "TESIS_TOTAL_C_1_player_Sexo", "sexo_str")
    df = ensure_column(df, "TESIS_TOTAL_C_1_player_Preg_Optimismo", "preg_optimismo")
    df = ensure_column(df, "TESIS_TOTAL_C_1_player_Preg_Confianza", "preg_confianza")

    if "preg_optimismo" not in df.columns:
        df["preg_optimismo"] = np.nan
    if "preg_confianza" not in df.columns:
        df["preg_confianza"] = np.nan

    df["edad"] = to_numeric(df.get("edad", pd.Series(np.nan, index=df.index)))
    df["sexo_str"] = df.get("sexo_str", pd.Series(pd.NA, index=df.index)).astype("string")
    df["mujer"] = df["sexo_str"].str.lower().isin({"female", "mujer"}).astype("float")
    df.loc[df["sexo_str"].isna(), "mujer"] = np.nan

    gk_cols = [
        col for col in df.columns
        if col.startswith("TESIS_TOTAL_C_1_player_gk_") and col.endswith("_ok")
    ]
    if gk_cols:
        df["gk_ok_total"] = df[gk_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)
    else:
        df["gk_ok_total"] = np.nan

    for suffix, target in zip("ABCD", ("pA", "pB", "pC", "pD")):
        cols = pattern_columns(df, suffix)
        df[target] = compute_row_mean(df, cols, target)

    df["shift_B"] = df["pB"] - df["pA"]
    df["shift_C"] = df["pC"] - df["pA"]
    df["shift_D"] = df["pD"] - df["pA"]

    df["optim_pre"] = df["pA"]
    df["optim_post"] = df["pD"]
    df["d_optim"] = df["optim_post"] - df["optim_pre"]

    analysis_df = df.loc[df[["pA", "pD", "treat"]].notna().all(axis=1)].copy()
    analysis_df = analysis_df.loc[analysis_df["treat"].isin({0, 1})]
    logging.info("Muestra analítica: %d observaciones.", len(analysis_df))
    return analysis_df


# ---------------------- ANALYTICS ------------------------------------------

def save_descriptives(df: pd.DataFrame, output_dir: Path) -> None:
    numeric_cols = [
        col for col in ["pA", "pB", "pC", "pD", "shift_B", "shift_C", "shift_D",
                        "preg_optimismo", "preg_confianza", "edad", "mujer"]
        if col in df.columns
    ]
    if not numeric_cols:
        logging.warning("No hay columnas numéricas para descriptivos.")
        return
    desc = df[numeric_cols].describe().T
    desc_path = output_dir / "descriptivos.csv"
    desc.to_csv(desc_path)
    logging.info("Descriptivos guardados en %s", desc_path)


def run_t_tests(df: pd.DataFrame, output_dir: Path) -> None:
    results = []
    for var in ["shift_B", "shift_C"]:
        if var not in df.columns:
            continue
        groups = [group[var].dropna() for _, group in df.groupby("treat")]
        if len(groups) != 2 or any(len(g) == 0 for g in groups):
            logging.warning("No es posible realizar t-test para %s (grupos vacíos).", var)
            continue
        t_stat, pvalue, dfree = ttest_ind(groups[1], groups[0], usevar="unequal")
        results.append({
            "variable": var,
            "t_stat": t_stat,
            "p_value": pvalue,
            "df": dfree,
            "n_treat": len(groups[1]),
            "n_control": len(groups[0]),
        })
        logging.info("t-test %s: t=%.3f, p=%.4f, df=%.1f", var, t_stat, pvalue, dfree)
    if results:
        ttest_path = output_dir / "ttests_shift_by_treat.csv"
        pd.DataFrame(results).to_csv(ttest_path, index=False)
        logging.info("Resultados de t-test guardados en %s", ttest_path)


def plot_kdes(df: pd.DataFrame, var: str, output_dir: Path) -> None:
    plt.figure(figsize=(8, 5))
    sns.kdeplot(df.loc[df["treat"] == 0, var], label="Control", fill=False)
    sns.kdeplot(df.loc[df["treat"] == 1, var], label="Tratamiento", fill=False)
    plt.title(f"Distribución de {var}")
    plt.xlabel(var)
    plt.ylabel("Densidad")
    plt.legend()
    output_path = output_dir / f"kdensity_{var}.png"
    plt.tight_layout()
    plt.savefig(output_path, dpi=150)
    plt.close()
    logging.info("Gráfico KDE guardado en %s", output_path)


def plot_shift_bars(df: pd.DataFrame, output_dir: Path) -> None:
    means = df.groupby("treat")[["shift_B", "shift_C", "shift_D"]].mean()
    means.index = means.index.map({0: "Control", 1: "Tratamiento"})
    ax = means.plot(kind="bar", figsize=(8, 5))
    ax.set_title("Shifts promedio por grupo")
    ax.set_xlabel("Grupo")
    ax.set_ylabel("Promedio")
    plt.tight_layout()
    output_path = output_dir / "shifts_barplot.png"
    plt.savefig(output_path, dpi=150)
    plt.close()
    logging.info("Gráfico de barras guardado en %s", output_path)


@dataclass
class ModelResult:
    name: str
    model: object


def fit_model(formula: str, df: pd.DataFrame, name: str) -> Optional[ModelResult]:
    try:
        model = smf.ols(formula=formula, data=df).fit(cov_type="HC1")
        logging.info("Modelo %s estimado con %d observaciones.", name, int(model.nobs))
        return ModelResult(name, model)
    except Exception as exc:  # pylint: disable=broad-except
        logging.exception("No se pudo estimar el modelo %s: %s", name, exc)
        return None


def run_regressions(df: pd.DataFrame, output_dir: Path) -> None:
    models: List[ModelResult] = []
    models.append(fit_model("shift_D ~ C(treat) + shift_B + shift_C + edad + C(mujer)", df, "ANCOVA D-A"))
    models.append(fit_model("shift_D ~ shift_B*C(treat) + shift_C*C(treat) + edad + C(mujer)", df, "Diferencial"))
    models.append(fit_model("optim_post ~ C(treat) + optim_pre + edad + C(mujer)", df, "ANCOVA niveles"))
    models.append(fit_model("d_optim ~ C(treat) + edad + C(mujer)", df, "DiD niveles"))

    optional_models: List[ModelResult] = []
    if df["gk_ok_total"].notna().any():
        optional_models.append(fit_model("shift_D ~ shift_B*C(treat) + shift_C*C(treat) + gk_ok_total + edad + C(mujer)", df, "Diferencial + desempeño"))
    if df[["preg_optimismo", "preg_confianza"]].notna().any().any():
        optional_models.append(fit_model("shift_D ~ shift_B*C(treat) + shift_C*C(treat) + preg_optimismo + preg_confianza + edad + C(mujer)", df, "Diferencial + encuesta"))

    models = [m for m in models if m is not None]
    if models:
        table = summary_col(
            results=[m.model for m in models],
            float_format="{:.3f}".format,
            stars=True,
            model_names=[m.name for m in models],
            info_dict={"N": lambda x: f"{int(x.nobs)}"}
        )
        table_path = output_dir / "resultados_modelos_clave.txt"
        with open(table_path, "w", encoding="utf-8") as f:
            f.write(table.as_text())
        logging.info("Tabla principal exportada a %s", table_path)

    for opt in optional_models:
        if opt is None:
            continue
        name_slug = opt.name.lower().replace(" ", "_").replace("+", "mas")
        with open(output_dir / f"modelo_{name_slug}.txt", "w", encoding="utf-8") as f:
            f.write(opt.model.summary().as_text())
        logging.info("Tabla de %s exportada.", opt.name)

    # Wald test for equal feedback impacts
    diff_model = next((m.model for m in models if m.name == "Diferencial"), None)
    if diff_model is not None:
        try:
            wald = diff_model.wald_test("shift_B:C(treat)[T.1] = shift_C:C(treat)[T.1]")
        except Exception:
            wald = diff_model.wald_test("C(treat)[T.1]:shift_B = C(treat)[T.1]:shift_C")
        wald_path = output_dir / "wald_test_diferencias.csv"
        pd.DataFrame({
            "chi2": [float(wald.statistic)],
            "p_value": [float(wald.pvalue)],
            "df": [float(np.atleast_1d(wald.df_denom)[0]) if hasattr(wald, "df_denom") else np.nan],
        }).to_csv(wald_path, index=False)
        logging.info("Prueba Wald exportada a %s", wald_path)


# ---------------------- MAIN ------------------------------------------------

def main() -> None:
    configure_logging(LOG_FILE)
    logging.info("Base de trabajo: %s", BASE_PATH)
    logging.info("Generando salidas en: %s", OUTPUT_DIR)

    df_raw = read_workbook(BASE_PATH, WORKBOOK_NAME)
    analysis_df = prepare_dataframe(df_raw)

    save_descriptives(analysis_df, OUTPUT_DIR)
    run_t_tests(analysis_df, OUTPUT_DIR)
    if len(analysis_df) > 0:
        plot_kdes(analysis_df, "pA", OUTPUT_DIR)
        plot_kdes(analysis_df, "pD", OUTPUT_DIR)
        if {"shift_B", "shift_C", "shift_D"}.issubset(analysis_df.columns):
            plot_shift_bars(analysis_df, OUTPUT_DIR)

    run_regressions(analysis_df, OUTPUT_DIR)

    dataset_path_csv = OUTPUT_DIR / "resultados_analiticos_shifts.csv"
    dataset_path_dta = OUTPUT_DIR / "resultados_analiticos_shifts.dta"
    analysis_df.to_csv(dataset_path_csv, index=False)
    logging.info("Base analítica guardada en %s", dataset_path_csv)
    try:
        analysis_df.to_stata(dataset_path_dta, write_index=False, version=118)
        logging.info("Base analítica en formato Stata guardada en %s", dataset_path_dta)
    except Exception as exc:  # pylint: disable=broad-except
        logging.warning("No se pudo guardar la base en formato Stata: %s", exc)


if __name__ == "__main__":
    main()
