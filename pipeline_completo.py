import os
import calendar
from datetime import date
import zipfile
import pandas as pd
import numpy as np
import holidays

# ==============================================================================
# CONFIGURACIÓN FLEXIBLE DE INSUMOS (Manejo de tildes y sufijo '(1)')
# ==============================================================================
def buscar_archivo(opciones):
    for opc in opciones:
        if os.path.exists(opc):
            return opc
    archivos_reales = os.listdir(".")
    for opc in opciones:
        for real in archivos_reales:
            if real.lower().strip() == opc.lower().strip():
                return real
    return opciones[0]

FILE_CAPITA = buscar_archivo(["Nota_Técnica_Capita (1).xlsx", "Nota_Tecnica_Capita (1).xlsx", "Nota_Tecnica_Capita.xlsx", "Nota_Técnica_Capita.xlsx"])
FILE_PGP = buscar_archivo(["Nota_Técnica_PGP (1).xlsx", "Nota_Tecnica_PGP (1).xlsx", "Nota_Tecnica_PGP.xlsx", "Nota_Técnica_PGP.xlsx"])
FILE_TECHOS = buscar_archivo(["Nota_Técnica_Techos (1).xlsx", "Nota_Tecnica_Techos (1).xlsx", "Nota_Tecnica_Techos.xlsx", "Nota_Técnica_Techos.xlsx"])
FILE_PREMIUM = buscar_archivo(["Nota_Técnica_Premium (1).xlsx", "Nota_Tecnica_Premium (1).xlsx", "Nota_Tecnica_Premium.xlsx", "Nota_Técnica_Premium.xlsx"])
FILE_PROGRAMAS = buscar_archivo(["Nota_Técnica_Programas (1).xlsx", "Nota_Tecnica_Programas (1).xlsx", "Nota_Tecnica_Programas.xlsx", "Nota_Técnica_Programas.xlsx"])
FILE_TECHO_PROG = buscar_archivo(["NT_TECHO_PROGRAMAS (1).xlsx", "NT_TECHO_PROGRAMAS.xlsx"])
FILE_POBLACION = buscar_archivo(["poblacion_capita_cm (1).xlsx", "poblacion_capita_cm.xlsx"])
FILE_SERVICIOS = buscar_archivo(["servicios.xlsx"])
FILE_THUMANO = buscar_archivo(["8. T Humano.xlsx", "8. T Humano (1).xlsx"])
FILE_REDUCTORES = buscar_archivo(["9. reductores de capacidad.xlsx", "9. reductores de capacidad (1).xlsx"])
FILE_INFRA = buscar_archivo(["10. infraestructura.xlsx", "10. infraestructura (1).xlsx"])
FILE_PRIORIDADES = buscar_archivo(["13. Tabla prioridades modelo propuesta.xlsx", "13. Tabla prioridades modelo propuesta (1).xlsx"])

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

COLUMNAS_FINALES = [
    "Col_01_Agrupacion", "Col_02_CUPS", "Col_03_Codigo_NT", "Col_04_Descripcion",
    "Col_05_Frecuencia_Anual", "Col_06_Actividades_Ano", "Col_07_Actividades_Mes",
    "Col_08_Promedio_CxE", "Col_09_Valor_Mes", "Col_10_CxU", "Col_11_Poblacion",
    "Col_12_Duracion_Consulta", "Col_13_Profesional", "Col_14_Tipo_Consultorio",
    "Col_15_Modalidad", "Col_16_Tipo_Contrato", "Col_17_UEN", "Col_18_Ciudad",
    "Col_19_Centro_Medico", "Col_20_Porcentaje", "Col_21_Vacia",
    "Col_22_Poblacion_Centro_Medico", "Col_23_Minutos_1", "Col_24_Minutos_2",
    "Col_25_Eventos_Anual_Calculado", "Col_26_Eventos_Mes_Calculado", "Col_27_Minutos_Totales_Mes"
]

print("=== INICIANDO PIPELINE DE GESTIÓN DE LA DEMANDA Y CAPACIDAD ===")

# ==============================================================================
# BLOQUE 1: CÁPITA, PGP Y TECHOS
# ==============================================================================
print("\n[1/8] Procesando Bloque 1 (Cápita, PGP, Techos)...")

def procesar_nota_tecnica(ruta, contrato):
    df = pd.read_excel(ruta)
    df_temp = pd.DataFrame(index=df.index)
    for i in range(min(10, df.shape[1])):
        df_temp[COLUMNAS_FINALES[i]] = df.iloc[:, i]
    df_temp["Col_16_Tipo_Contrato"] = contrato
    df_temp["Col_20_Porcentaje"] = df.iloc[:, -1]
    return df_temp

df_capita = procesar_nota_tecnica(FILE_CAPITA, "cápita")
df_pgp = procesar_nota_tecnica(FILE_PGP, "PGP")
df_techos = procesar_nota_tecnica(FILE_TECHOS, "Techos")
df_b1 = pd.concat([df_capita, df_pgp, df_techos], ignore_index=True)

df_servicios = pd.read_excel(FILE_SERVICIOS)
col_serv_cups = df_servicios.columns[3]
col_serv_uen = df_servicios.columns[2]

df_b1["Col_02_CUPS"] = df_b1["Col_02_CUPS"].astype(str).str.strip()
df_servicios[col_serv_cups] = df_servicios[col_serv_cups].astype(str).str.strip()

cols_serv_b1 = {
    col_serv_cups: "Col_02_CUPS",
    df_servicios.columns[5]: "Col_12_Duracion_Consulta",
    df_servicios.columns[6]: "Col_13_Profesional",
    df_servicios.columns[7]: "Col_14_Tipo_Consultorio",
    df_servicios.columns[8]: "Col_15_Modalidad"
}
df_serv_sub1 = df_servicios[list(cols_serv_b1.keys())].rename(columns=cols_serv_b1).drop_duplicates(subset=["Col_02_CUPS"])
df_b1 = df_b1.merge(df_serv_sub1, on="Col_02_CUPS", how="left")

df_poblacion = pd.read_excel(FILE_POBLACION)
df_pob_sub = pd.DataFrame({
    "Col_17_UEN": df_poblacion.iloc[:, 0].astype(str).str.strip().str.upper(),
    "Col_18_Ciudad": df_poblacion.iloc[:, 1].astype(str).str.strip().str.upper(),
    "Col_19_Centro_Medico": df_poblacion.iloc[:, 2].astype(str).str.strip().str.upper(),
    "Col_22_Poblacion_Centro_Medico": pd.to_numeric(df_poblacion.iloc[:, 4], errors="coerce").fillna(0)
})
df_b1 = df_b1.merge(df_pob_sub, how="cross")

df_b1["Col_05_Frecuencia_Anual"] = pd.to_numeric(df_b1["Col_05_Frecuencia_Anual"], errors="coerce").fillna(0)
df_b1["Col_12_Duracion_Consulta"] = pd.to_numeric(df_b1["Col_12_Duracion_Consulta"], errors="coerce").fillna(0)
df_b1["Col_25_Eventos_Anual_Calculado"] = df_b1["Col_22_Poblacion_Centro_Medico"] * df_b1["Col_05_Frecuencia_Anual"]
df_b1["Col_26_Eventos_Mes_Calculado"] = df_b1["Col_25_Eventos_Anual_Calculado"] / 12.0
df_b1["Col_27_Minutos_Totales_Mes"] = df_b1["Col_26_Eventos_Mes_Calculado"] * df_b1["Col_12_Duracion_Consulta"]

for col in COLUMNAS_FINALES:
    if col not in df_b1.columns:
        df_b1[col] = np.nan
df_b1 = df_b1[COLUMNAS_FINALES]

# ==============================================================================
# BLOQUE 2: PROGRAMAS
# ==============================================================================
print("\n[2/8] Procesando Bloque 2 (Programas)...")
df_nt_prog = pd.read_excel(FILE_PROGRAMAS)
df_prog = pd.DataFrame(index=df_nt_prog.index)
df_prog["Col_18_Ciudad"] = df_nt_prog.iloc[:, 1].astype(str).str.strip().str.upper()
df_prog["Col_17_UEN"] = df_nt_prog.iloc[:, 2].astype(str).str.strip().str.upper()
df_prog["Col_19_Centro_Medico"] = df_nt_prog.iloc[:, 3].astype(str).str.strip().str.upper()
df_prog["Col_01_Agrupacion"] = df_nt_prog.iloc[:, 4]
df_prog["Col_03_Codigo_NT"] = df_nt_prog.iloc[:, 5].astype(str).str.strip()
df_prog["Col_04_Descripcion"] = df_nt_prog.iloc[:, 6]
df_prog["Col_13_Profesional"] = df_nt_prog.iloc[:, 7].astype(str).str.strip().str.upper()
df_prog["Col_05_Frecuencia_Anual"] = pd.to_numeric(df_nt_prog.iloc[:, 8], errors="coerce").fillna(0)
df_prog["Col_22_Poblacion_Centro_Medico"] = pd.to_numeric(df_nt_prog.iloc[:, 9], errors="coerce").fillna(0)
df_prog["Col_06_Actividades_Ano"] = pd.to_numeric(df_nt_prog.iloc[:, 10], errors="coerce").fillna(0)
df_prog["Col_07_Actividades_Mes"] = pd.to_numeric(df_nt_prog.iloc[:, 11], errors="coerce").fillna(0)
df_prog["Col_20_Porcentaje"] = df_nt_prog.iloc[:, 14]
df_prog["Col_16_Tipo_Contrato"] = "Programas"

cols_serv_b2 = {
    col_serv_cups: "Col_03_Codigo_NT",
    df_servicios.columns[5]: "Col_12_Duracion_Consulta",
    df_servicios.columns[6]: "Col_13_Profesional_Serv",
    df_servicios.columns[7]: "Col_14_Tipo_Consultorio",
    df_servicios.columns[8]: "Col_15_Modalidad"
}
df_serv_sub2 = df_servicios[list(cols_serv_b2.keys())].rename(columns=cols_serv_b2).drop_duplicates(subset=["Col_03_Codigo_NT"])
df_prog = df_prog.merge(df_serv_sub2, on="Col_03_Codigo_NT", how="left")
df_prog["Col_13_Profesional"] = df_prog["Col_13_Profesional_Serv"].combine_first(df_prog["Col_13_Profesional"])
df_prog.drop(columns=["Col_13_Profesional_Serv"], inplace=True)

df_techo_prog = pd.read_excel(FILE_TECHO_PROG)
col_cod_techo = df_techo_prog.columns[2]
df_techo_prog[col_cod_techo] = df_techo_prog[col_cod_techo].astype(str).str.strip()
cols_techo_map = {
    col_cod_techo: "Col_03_Codigo_NT",
    df_techo_prog.columns[1]: "Col_02_CUPS",
    df_techo_prog.columns[7]: "Col_08_Promedio_CxE"
}
df_techo_sub = df_techo_prog[list(cols_techo_map.keys())].rename(columns=cols_techo_map).drop_duplicates(subset=["Col_03_Codigo_NT"])
df_prog = df_prog.merge(df_techo_sub, on="Col_03_Codigo_NT", how="left")

df_prog["Col_08_Promedio_CxE"] = pd.to_numeric(df_prog["Col_08_Promedio_CxE"], errors="coerce").fillna(0)
df_prog["Col_12_Duracion_Consulta"] = pd.to_numeric(df_prog["Col_12_Duracion_Consulta"], errors="coerce").fillna(0)
df_prog["Col_09_Valor_Mes"] = df_prog["Col_07_Actividades_Mes"] * df_prog["Col_08_Promedio_CxE"]
df_prog["Col_10_CxU"] = np.where(df_prog["Col_22_Poblacion_Centro_Medico"] != 0, df_prog["Col_09_Valor_Mes"] / df_prog["Col_22_Poblacion_Centro_Medico"], 0)
df_prog["Col_25_Eventos_Anual_Calculado"] = df_prog["Col_22_Poblacion_Centro_Medico"] * df_prog["Col_05_Frecuencia_Anual"]
df_prog["Col_26_Eventos_Mes_Calculado"] = df_prog["Col_25_Eventos_Anual_Calculado"] / 12.0
df_prog["Col_27_Minutos_Totales_Mes"] = df_prog["Col_26_Eventos_Mes_Calculado"] * df_prog["Col_12_Duracion_Consulta"]

for col in COLUMNAS_FINALES:
    if col not in df_prog.columns:
        df_prog[col] = np.nan
df_prog = df_prog[COLUMNAS_FINALES]

# ==============================================================================
# BLOQUE 3: PREMIUM Y CONSOLIDACIÓN
# ==============================================================================
print("\n[3/8] Procesando Bloque 3 (Premium y Consolidado General)...")
df_prem = pd.read_excel(FILE_PREMIUM)
df_prem_init = pd.DataFrame()
df_prem_init["Col_16_Tipo_Contrato"] = df_prem.iloc[:, 1]
df_prem_init["Col_18_Ciudad"] = df_prem.iloc[:, 2].astype(str).str.strip().str.upper()
df_prem_init["Col_01_Agrupacion"] = df_prem.iloc[:, 5]
df_prem_init["Col_03_Codigo_NT"] = df_prem.iloc[:, 6].astype(str).str.strip()
df_prem_init["Col_04_Descripcion"] = df_prem.iloc[:, 7]
df_prem_init["Actividades_Originales"] = pd.to_numeric(df_prem.iloc[:, 8], errors="coerce").fillna(0)
df_prem_init["Col_20_Porcentaje"] = df_prem.iloc[:, 9]

cols_serv_prem = {
    col_serv_cups: "Col_03_Codigo_NT",
    df_servicios.columns[1]: "Col_02_CUPS",
    col_serv_uen: "Col_17_UEN",
    df_servicios.columns[5]: "Col_12_Duracion_Consulta",
    df_servicios.columns[6]: "Col_13_Profesional",
    df_servicios.columns[7]: "Col_14_Tipo_Consultorio",
    df_servicios.columns[8]: "Col_15_Modalidad"
}
df_serv_subprem = df_servicios[list(cols_serv_prem.keys())].rename(columns=cols_serv_prem).drop_duplicates(subset=["Col_03_Codigo_NT"])
df_prem_init = df_prem_init.merge(df_serv_subprem, on="Col_03_Codigo_NT", how="left")
df_prem_init["Col_17_UEN"] = df_prem_init["Col_17_UEN"].astype(str).str.strip().str.upper()

df_sedes_map = pd.DataFrame({
    "Col_17_UEN": df_poblacion.iloc[:, 0].astype(str).str.strip().str.upper(),
    "Col_19_Centro_Medico": df_poblacion.iloc[:, 2].astype(str).str.strip().str.upper(),
    "Col_22_Poblacion_Centro_Medico": pd.to_numeric(df_poblacion.iloc[:, 4], errors="coerce").fillna(0)
})
df_prem_final = df_prem_init.merge(df_sedes_map, on="Col_17_UEN", how="left")
conteo_sedes = df_prem_final.groupby(["Col_03_Codigo_NT", "Col_17_UEN"])["Col_19_Centro_Medico"].transform("count")
df_prem_final["N_Sedes"] = np.where(conteo_sedes > 0, conteo_sedes, 1)

df_prem_final["Col_26_Eventos_Mes_Calculado"] = df_prem_final["Actividades_Originales"] / df_prem_final["N_Sedes"]
df_prem_final["Col_12_Duracion_Consulta"] = pd.to_numeric(df_prem_final["Col_12_Duracion_Consulta"], errors="coerce").fillna(0)
df_prem_final["Col_25_Eventos_Anual_Calculado"] = df_prem_final["Col_26_Eventos_Mes_Calculado"] * 12.0
df_prem_final["Col_27_Minutos_Totales_Mes"] = df_prem_final["Col_26_Eventos_Mes_Calculado"] * df_prem_final["Col_12_Duracion_Consulta"]

for col in COLUMNAS_FINALES:
    if col not in df_prem_final.columns:
        df_prem_final[col] = np.nan
df_prem_final = df_prem_final[COLUMNAS_FINALES]

df_consolidado_total = pd.concat([df_b1, df_prog, df_prem_final], ignore_index=True)
df_consolidado_total.to_excel("Base_General_Consolidada.xlsx", index=False)
with zipfile.ZipFile("Base_General_Consolidada.zip", "w", compression=zipfile.ZIP_DEFLATED) as zf:
    zf.write("Base_General_Consolidada.xlsx", arcname="Base_General_Consolidada.xlsx")
print(f"-> Base consolidada: {df_consolidado_total.shape[0]} registros generados en Base_General_Consolidada.zip")

# ==============================================================================
# BLOQUE 4: TABLAS DE RESUMEN
# ==============================================================================
print("\n[4/8] Procesando Bloque 4 (Agregaciones y Resúmenes)...")
with pd.ExcelWriter("Resumenes_Consolidados_Bloque4.xlsx", engine="openpyxl") as w:
    df_consolidado_total.groupby(["Col_02_CUPS", "Col_04_Descripcion"], as_index=False)[["Col_07_Actividades_Mes", "Col_27_Minutos_Totales_Mes"]].sum().to_excel(w, sheet_name="Por_CUPS", index=False)
    df_consolidado_total.groupby(["Col_18_Ciudad"], as_index=False)[["Col_07_Actividades_Mes", "Col_27_Minutos_Totales_Mes"]].sum().to_excel(w, sheet_name="Por_Ciudad", index=False)
    df_consolidado_total.groupby(["Col_19_Centro_Medico"], as_index=False)[["Col_07_Actividades_Mes", "Col_27_Minutos_Totales_Mes"]].sum().to_excel(w, sheet_name="Por_Centro_Medico", index=False)

# ==============================================================================
# BLOQUE 5: TALENTO HUMANO CON FESTIVOS COLOMBIA POR MES
# ==============================================================================
print("\n[5/8] Procesando Bloque 5 (Talento Humano con Festivos Colombia)...")
df_th = pd.read_excel(FILE_THUMANO)
df_red = pd.read_excel(FILE_REDUCTORES)

df_th["Ciudad"] = df_th.iloc[:, 1].astype(str).str.strip().str.upper()
df_th["UEN"] = df_th.iloc[:, 2].astype(str).str.strip().str.upper()
df_th["TITULO DE CARGO"] = df_th.iloc[:, 7].astype(str).str.strip().str.upper()

df_red["UEN"] = df_red.iloc[:, 5].astype(str).str.strip().str.upper()
df_red["TITULO DE CARGO"] = df_red.iloc[:, 7].astype(str).str.strip().str.upper()
df_red["Valor"] = pd.to_numeric(df_red.iloc[:, 9], errors="coerce").fillna(0)

df_red_agrup = df_red.groupby(["UEN", "TITULO DE CARGO"], as_index=False)["Valor"].sum().rename(columns={"Valor": "reductor"})
df_th = df_th.merge(df_red_agrup, on=["UEN", "TITULO DE CARGO"], how="left")
df_th["reductor"] = df_th["reductor"].fillna(0)

horas_semanales = pd.to_numeric(df_th.iloc[:, 9], errors="coerce").fillna(0)
df_th["horas_netas_semanales"] = horas_semanales * ((100.0 - df_th["reductor"]) / 100.0)

co_holidays = holidays.Colombia(years=2026)
cal = calendar.Calendar()
for m_idx, m_nom in enumerate(MESES, start=1):
    dias_h = sum(1 for d, w in cal.itermonthdays2(2026, m_idx) if d != 0 and w < 6 and date(2026, m_idx, d) not in co_holidays)
    df_th[f"Horas_Disponibles_{m_nom}"] = (df_th["horas_netas_semanales"] / 6.0) * dias_h
    df_th[f"Minutos_Disponibles_{m_nom}"] = df_th[f"Horas_Disponibles_{m_nom}"] * 60.0

df_th.to_excel("8_T_Humano_Procesado.xlsx", index=False)

# ==============================================================================
# BLOQUE 6: INFRAESTRUCTURA CON FESTIVOS COLOMBIA POR MES
# ==============================================================================
print("\n[6/8] Procesando Bloque 6 (Infraestructura con Festivos Colombia)...")
df_inf = pd.read_excel(FILE_INFRA)
col_h_disp = df_inf.columns[8]
df_inf[col_h_disp] = pd.to_numeric(df_inf[col_h_disp], errors="coerce").fillna(0)

dias_cols = [df_inf.columns[i] for i in range(9, 16)]
nombres_dispo = [f"dispo_{d}" for d in ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"]]
for c_orig, c_disp in zip(dias_cols, nombres_dispo):
    df_inf[c_disp] = df_inf[col_h_disp] * pd.to_numeric(df_inf[c_orig], errors="coerce").fillna(0)

for m_idx, m_nom in enumerate(MESES, start=1):
    conteo_dias_semana = [0] * 7
    for d, w in cal.itermonthdays2(2026, m_idx):
        if d != 0 and date(2026, m_idx, d) not in co_holidays:
            conteo_dias_semana[w] += 1
    
    h_mes = np.zeros(len(df_inf))
    for w_idx, c_disp in enumerate(nombres_dispo):
        h_mes += df_inf[c_disp].values * conteo_dias_semana[w_idx]
    df_inf[f"Horas_Disponibles_{m_nom}"] = h_mes
    df_inf[f"Minutos_Disponibles_{m_nom}"] = h_mes * 60.0

df_inf.to_excel("10_Infraestructura_Procesada.xlsx", index=False)

# ==============================================================================
# BLOQUE 7: BALANCE MENSUAL DEMANDA VS OFERTAS
# ==============================================================================
print("\n[7/8] Procesando Bloque 7 (Balance Mensual Demanda vs Ofertas)...")
cols_min_m = [f"Minutos_Disponibles_{m}" for m in MESES]

df_oferta_th_mes = df_th.groupby(["Ciudad", "UEN", "TITULO DE CARGO"], as_index=False)[cols_min_m].sum()

col_inf_ciudad = df_inf.columns[1]
col_inf_uen = df_inf.columns[5]
col_inf_cons = df_inf.columns[7]
df_inf["Ciudad"] = df_inf[col_inf_ciudad].astype(str).str.strip().str.upper()
df_inf["UEN"] = df_inf[col_inf_uen].astype(str).str.strip().str.upper()
df_inf["Tipo Consultorio"] = df_inf[col_inf_cons].astype(str).str.strip().str.upper()

df_oferta_inf_mes = df_inf.groupby(["Ciudad", "UEN", "Tipo Consultorio"], as_index=False)[cols_min_m].sum()

with pd.ExcelWriter("Bloque7_Balance_Demanda_Oferta.xlsx", engine="openpyxl") as w:
    df_consolidado_total.to_excel(w, sheet_name="Demanda_Base", index=False)
    df_oferta_th_mes.to_excel(w, sheet_name="Oferta_Talento_Humano", index=False)
    df_oferta_inf_mes.to_excel(w, sheet_name="Oferta_Infraestructura", index=False)

# ==============================================================================
# BLOQUE 8: MOTOR DE ASIGNACIÓN Y DÉFICIT 12 MESES
# ==============================================================================
print("\n[8/8] Procesando Bloque 8 (Asignación y Diagnóstico de Capacidad para los 12 Meses)...")
dict_prio = {}
if os.path.exists(FILE_PRIORIDADES):
    df_prio = pd.read_excel(FILE_PRIORIDADES)
    c_cod, c_cob, c_ord = df_prio.columns[3], df_prio.columns[5], df_prio.columns[6]
    df_prio[c_cod] = df_prio[c_cod].astype(str).str.strip()
    cob = pd.to_numeric(df_prio[c_cob], errors="coerce").fillna(1.0)
    df_prio["Cob"] = np.where(cob > 1.0, cob / 100.0, cob)
    df_prio["Prio"] = pd.to_numeric(df_prio[c_ord], errors="coerce").fillna(999)
    dict_prio = df_prio.set_index(c_cod)[["Cob", "Prio"]].to_dict(orient="index")

dict_dur = dict(zip(
    df_servicios[col_serv_cups].astype(str).str.strip(),
    pd.to_numeric(df_servicios.iloc[:, 5], errors="coerce").fillna(20.0)
))

df_dem = df_consolidado_total.copy()
df_dem["Ciudad"] = df_dem["Col_18_Ciudad"].astype(str).str.strip().str.upper()
df_dem["UEN"] = df_dem["Col_17_UEN"].astype(str).str.strip().str.upper()
df_dem["Cargo"] = df_dem["Col_13_Profesional"].astype(str).str.strip().str.upper()
df_dem["Consultorio"] = df_dem["Col_14_Tipo_Consultorio"].astype(str).str.strip().str.upper()
df_dem["Codigo_NT"] = df_dem["Col_03_Codigo_NT"].astype(str).str.strip()
df_dem["CUPS"] = df_dem["Col_02_CUPS"].astype(str).str.strip()

df_dem["Prioridad"] = df_dem["Codigo_NT"].apply(lambda x: dict_prio.get(x, {}).get("Prio", 999))
df_dem["Cobertura"] = df_dem["Codigo_NT"].apply(lambda x: dict_prio.get(x, {}).get("Cob", 1.0))
df_dem["Duracion"] = df_dem["Codigo_NT"].map(dict_dur).fillna(df_dem["CUPS"].map(dict_dur)).fillna(df_dem["Col_12_Duracion_Consulta"]).fillna(20.0)
df_dem["Minutos_Objetivo"] = pd.to_numeric(df_dem["Col_27_Minutos_Totales_Mes"], errors="coerce").fillna(0) * df_dem["Cobertura"]

df_dem.sort_values(by=["Prioridad", "Col_20_Porcentaje", "Col_08_Promedio_CxE"], ascending=[True, False, True], inplace=True)

lista_asig = []
lista_no_asig = []

for mes in MESES:
    c_m = f"Minutos_Disponibles_{mes}"
    pool_th = {(r["Ciudad"], r["UEN"], r["TITULO DE CARGO"]): float(r[c_m]) for _, r in df_oferta_th_mes.iterrows()}
    pool_inf = {(r["Ciudad"], r["UEN"], r["Tipo Consultorio"]): float(r[c_m]) for _, r in df_oferta_inf_mes.iterrows()}
    
    for _, req in df_dem.iterrows():
        min_dem = req["Minutos_Objetivo"]
        if min_dem <= 0:
            continue
            
        ciudad, uen_o, cargo, cons, dur = req["Ciudad"], req["UEN"], req["Cargo"], req["Consultorio"], req["Duracion"]
        es_med_gen = "MEDICINA GENERAL" in str(cargo) or "MEDICINA GENERAL" in str(req["Col_04_Descripcion"]).upper()
        
        uens = [uen_o] if es_med_gen else [uen_o] + list({k[1] for k in pool_th.keys() if k[0] == ciudad and k[1] != uen_o})
        restante = min_dem
        
        for uen_d in uens:
            k_th, k_inf = (ciudad, uen_d, cargo), (ciudad, uen_d, cons)
            cuello = min(pool_th.get(k_th, 0.0), pool_inf.get(k_inf, 0.0))
            
            if cuello > 0:
                asig = min(restante, cuello)
                pool_th[k_th] -= asig
                pool_inf[k_inf] -= asig
                restante -= asig
                
                lista_asig.append({
                    "Mes": mes, "Ciudad": ciudad, "UEN": uen_d, "Consultorio": cons,
                    "Codigo_NT": req["Codigo_NT"], "CUPS": req["CUPS"], "Descripcion": req["Col_04_Descripcion"],
                    "Cargo": cargo, "Duracion_Min": dur, "Minutos_Asignados": asig, "Citas_Asignadas": asig / dur
                })
            if restante <= 0:
                break
                
        if restante > 0:
            d_th = pool_th.get((ciudad, uen_o, cargo), 0.0)
            d_inf = pool_inf.get((ciudad, uen_o, cons), 0.0)
            diag = "Déficit simultáneo TH e Infra" if (d_th <= 0 and d_inf <= 0) else ("Déficit TH" if d_th <= 0 else "Déficit Infra")
            
            lista_no_asig.append({
                "Mes": mes, "Ciudad": ciudad, "UEN": uen_o, "Codigo_NT": req["Codigo_NT"],
                "CUPS": req["CUPS"], "Descripcion": req["Col_04_Descripcion"], "Cargo": cargo,
                "Consultorio": cons, "Duracion_Min": dur, "Minutos_No_Asignados": restante,
                "Citas_No_Asignadas": restante / dur, "Horas_Faltantes": restante / 60.0, "Cuello_Botella": diag
            })

df_res_asig = pd.DataFrame(lista_asig)
df_res_no_asig = pd.DataFrame(lista_no_asig)

salida_b8 = "Bloque8_Resultados_Mensuales_Capacidad.xlsx"
with pd.ExcelWriter(salida_b8, engine="openpyxl") as w:
    df_res_asig.to_excel(w, sheet_name="Asignado_Mes", index=False)
    df_res_no_asig.to_excel(w, sheet_name="No_Asignado_Mes", index=False)
    
    df_res_asig.groupby("Mes", as_index=False).agg(
        Minutos_Asignados=("Minutos_Asignados", "sum"),
        Citas_Asignadas=("Citas_Asignadas", "sum")
    ).merge(
        df_res_no_asig.groupby("Mes", as_index=False).agg(
            Minutos_No_Asignados=("Minutos_No_Asignados", "sum"),
            Citas_No_Asignadas=("Citas_No_Asignadas", "sum")
        ), on="Mes", how="outer"
    ).fillna(0).to_excel(w, sheet_name="Resumen_Consolidado_Mes", index=False)

print("\n==================================================================")
print("PIPELINE FINALIZADO CON ÉXITO")
print(f"Archivo de resultados generado: {salida_b8}")
print("==================================================================")
