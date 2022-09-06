import xlwings as xw
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib import ticker
from scipy import stats
from Funciones import (
    Efw_caso1,
    Eo_caso1,
    F_caso1,
    Eo_caso3,
    Eg_caso3,
    F_caso3,
    Rp_corregido,
)

# Crete names for sheets
SHEET_SUMMARY = "Resumen"
SHEET_RESULTS = "Resultados"

# Ejemplo 1 (Campo Prudhoe Bay)
STOC_VALUES = "df"
VALUES_EFW = "datos_Efw"
DP = "dP"
Efw_columns = "Efw"
F_columns = "F_"
Eo_columns = "Eo"
Eo_Efw_columns = "Eo_Efw"
Swi_IDX, Bw_IDX, Cw_IDX, Pb_IDX, Cf_IDX = 0, 1, 2, 3, 4

# Ejemplo 2 (Exam Exercise)
DF_Exam = "Exam_Exercise_dataset"
VALUES_EXAM = "datos_examen"
Eg_exam_columns = "Eg_Exam"
Eo_exam_columns = "Eo_Exam"
F_exam_columns = "F_Exam"
Rpco_exam_columns = "Rp_corregido"
Eo_mEg_columns = "Eo_mEg"
N_volumetric_IDX, Rsi_IDX, M_IDX = 0, 1, 2

# Ejemplo 3 (Campo Valhall)
DF_VALHALL = "Valhall_dataset"
VALUES_EFW = "datos_Efw"
DP_VALHALL = "dP_3"
Efw_columns_3 = "Efw_3"
F_columns_3 = "F_3"
Eo_columns_3 = "Eo_3"
Eo_Efw_columns_3 = "Eo_Efw_3"


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[SHEET_SUMMARY]

    # Import dataframe from Ms. Excel
    df_Prudhoe = (
        sheet[STOC_VALUES].options(pd.DataFrame, index=False, expand="table").value
    )
    params = sheet[VALUES_EFW].options(np.array, transpose=True).value
    P = df_Prudhoe["p"].values
    Bo = df_Prudhoe["Bo"].values
    Boi = Bo[0]
    Np = df_Prudhoe["Np"].values
    Wp = df_Prudhoe["Wp"].values
    Rs = df_Prudhoe["Rs"].values
    Rsi = Rs[0]
    Bg = df_Prudhoe["Bg"].values

    # Calculo de diferencial de presión
    Pi = P[0]
    DP_value = Pi - P
    sheet[DP].options(transpose=True).value = DP_value

    # Calculo de Efw, Eo, F
    Efw = Efw_caso1(Boi, params[Cf_IDX], params[Swi_IDX], params[Cw_IDX], DP_value)
    sheet[Efw_columns].options(transpose=True).value = Efw
    Eo = Eo_caso1(Bo, Boi, Rsi, Rs, Bg)
    sheet[Eo_columns].options(transpose=True).value = Eo
    F = F_caso1(Np, Bo, Boi)
    sheet[F_columns].options(transpose=True).value = F
    Eo_Efw = Eo + Efw
    sheet[Eo_Efw_columns].options(transpose=True).value = Eo_Efw
    # Grafica
    slope, intercept, r_value, p_value, std_err = stats.linregress(Eo_Efw, F)
    y_fit = intercept + (Eo_Efw * slope)
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.scatter(Eo_Efw, F, label="Original Data")
    ax.plot(Eo_Efw, y_fit, c="g", label="Fitted Line")
    plt.legend()
    text = "Intercept: %.1f\nN: %.3f" % (intercept, slope / 1e6)
    plt.text(0.008, 1, text)
    plt.show()

    # Ejemplo #2 ejercicio del examen
    # -----------------------------------------------------------------------------------
    df_Exam = sheet[DF_Exam].options(pd.DataFrame, index=False, expand="table").value
    params_exam = sheet[VALUES_EXAM].options(np.array, transpose=True).value
    P_exam = df_Exam["p"].values
    Np_exam = df_Exam["Np"].values
    Bg_exam = df_Exam["Bg"].values
    Bgi_exam = Bg_exam[0]
    Bt_exam = df_Exam["Bt"].values
    Bti_exam = Bt_exam[0]
    Rp_exam = df_Exam["Rp"].values

    # Calculo de Eg, Eo
    Eg_exam = Eg_caso3(Bti_exam, Bg_exam, Bgi_exam)
    sheet[Eg_exam_columns].options(transpose=True).value = Eg_exam
    Eo_exam = Eo_caso3(Bt_exam, Bti_exam, 0, 0, 0)
    sheet[Eo_exam_columns].options(transpose=True).value = Eo_exam
    Rp_corregido_exam = Rp_corregido(
        params_exam[N_volumetric_IDX],
        Eo_exam,
        Eg_exam,
        params_exam[M_IDX],
        Np_exam,
        Bt_exam,
        params_exam[Rsi_IDX],
        Bg_exam,
    )
    sheet[Rpco_exam_columns].options(transpose=True).value = Rp_corregido_exam
    F_exam = F_caso3(Np_exam, Bt_exam, Rp_corregido_exam, params_exam[Rsi_IDX], Bg_exam)
    sheet[F_exam_columns].options(transpose=True).value = F_exam
    Eo_mEg_value = Eo_exam + params_exam[M_IDX] * Eg_exam
    sheet[Eo_mEg_columns].options(transpose=True).value = Eo_mEg_value
    slope2, intercept2, r_value2, p_value2, std_err2 = stats.linregress(
        Eo_mEg_value, F_exam
    )
    y_fit = intercept2 + (Eo_mEg_value * slope2)
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.scatter(Eo_mEg_value, F_exam, label="Original Data")
    ax.plot(Eo_mEg_value, y_fit, c="g", label="Fitted Line")
    plt.legend()
    text = "Intercept: %.1f\nN: %.3f" % (intercept2, slope2 / 1e6)
    plt.text(0.008, 1, text)
    plt.show()

    # Ejemplo #3 Campo Valhall
    # -----------------------------------------------------------------------------------
    df_Valhall = (
        sheet[DF_VALHALL].options(pd.DataFrame, index=False, expand="table").value
    )
    params = sheet[VALUES_EFW].options(np.array, transpose=True).value
    P_3 = df_Valhall["p"].values
    Bo_3 = df_Valhall["Bo"].values
    Boi_3 = Bo_3[0]
    Np_3 = df_Valhall["Np"].values
    Wp_3 = df_Valhall["Wp"].values
    Rs_3 = df_Valhall["Rs"].values
    Rsi_3 = Rs_3[0]

    # Calculo de diferencial de presión
    Pi_3 = P_3[0]
    DP_value_3 = Pi_3 - P_3
    sheet[DP_VALHALL].options(transpose=True).value = DP_value_3

    # Calculo de Efw, Eo, F
    Efw_3 = Efw_caso1(
        Boi_3, params[Cf_IDX], params[Swi_IDX], params[Cw_IDX], DP_value_3
    )
    sheet[Efw_columns_3].options(transpose=True).value = Efw_3
    Eo_3 = Eo_caso1(Bo_3, Boi_3, Rsi_3, Rs_3, 0)
    sheet[Eo_columns_3].options(transpose=True).value = Eo_3
    F_3 = F_caso1(Np_3, Bo_3, Boi_3)
    sheet[F_columns_3].options(transpose=True).value = F_3
    Eo_Efw_3 = Eo_3 + Efw_3
    sheet[Eo_Efw_columns_3].options(transpose=True).value = Eo_Efw_3
    # Grafica
    slope, intercept, r_value, p_value, std_err = stats.linregress(Eo_Efw_3, F_3)
    y_fit = intercept + (Eo_Efw_3 * slope)
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.scatter(Eo_Efw_3, F_3, label="Original Data")
    ax.plot(Eo_Efw_3, y_fit, c="g", label="Fitted Line")
    plt.legend()
    text = "Intercept: %.1f\nN: %.3f" % (intercept, slope / 1e6)
    plt.text(0.008, 1, text)
    plt.show()


if __name__ == "__main__":
    xw.Book("intro.xlsm").set_mock_caller()
    main()
