import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font



def generar_resumen_modelo_atencion(banco):
    
    # -------------------------------------------------
    # CÓDIGOS DEL MODELO
    # -------------------------------------------------
    codigos_modelo = [
        "IND-CTT-001","IND-CTT-002","IND-AUT-003","IND-RYC-001",
        "IND-RYC-002","IND-PYP-027","IND-PYP-028","IND-PYP-021",
        "IND-PYP-030","IND-CTT-003","IND-CTT-004","IND-CTT-005",
        "IND-CTT-006","IND-CTT-007","IND-CTT-008","IND-SISPI-002",
        "IND-SISPI-003","IND-SISPI-004","IND-GDR-001","IND-GDR-002",
    ]

    # -------------------------------------------------
    # REGLAS DE VALORACIÓN
    # -------------------------------------------------
    reglas_indicadores = {
        "IND-CTT-001": {"meta": 100, "op": ">="},
        "IND-CTT-002": {"meta": 100, "op": ">="},
        "IND-AUT-003": {"meta": 100, "op": ">="},

        "IND-RYC-001": {"meta": 80, "op": ">="},
        "IND-RYC-002": {"meta": 80, "op": ">="},
        "IND-PYP-027": {"meta": 80, "op": ">="},
        "IND-PYP-028": {"meta": 80, "op": ">="},

        "IND-PYP-021": {"meta": 50, "op": ">"},

        "IND-PYP-030": {"meta": 1, "op": ">="},
        "IND-CTT-003": {"meta": 1, "op": ">="},
        "IND-CTT-004": {"meta": 1, "op": ">="},
        "IND-CTT-005": {"meta": 1, "op": ">="},

        "IND-CTT-006": {"meta": 90, "op": ">="},
        "IND-CTT-007": {"meta": 90, "op": ">="},
        "IND-CTT-008": {"meta": 90, "op": ">="},
        "IND-SISPI-002": {"meta": 90, "op": ">="},
        "IND-SISPI-003": {"meta": 90, "op": ">="},
        "IND-SISPI-004": {"meta": 90, "op": ">="},

        "IND-GDR-001": {"meta": 100, "op": ">="},
        "IND-GDR-002": {"meta": 23.75, "op": ">="},
    }

    # -------------------------------------------------
    # NORMALIZAR COLUMNAS
    # -------------------------------------------------
    banco.columns = banco.columns.str.strip().str.upper()

    columnas_requeridas = [
        "CONSE","INDICADOR","ÁREA",
        "TIPO DE INDICADOR",
        "ACEPTABLE","SATISFACTORIO",
        "VALOR ANUAL","VALORACIÓN"
    ]

    faltantes = [c for c in columnas_requeridas if c not in banco.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas: {faltantes}")

    # -------------------------------------------------
    # FILTRAR MODELO
    # -------------------------------------------------
    modelo = banco[
        banco["CONSE"].astype(str).str.strip().isin(codigos_modelo)
    ].copy()

    # -------------------------------------------------
    # META
    # -------------------------------------------------
    modelo["META"] = modelo["ACEPTABLE"].astype(str).str.strip()

    modelo.loc[
        modelo["META"].isin(["", "N/A", "NA", "nan"]),
        "META"
    ] = modelo["SATISFACTORIO"]

    # -------------------------------------------------
    # LIMPIEZA NUMÉRICA
    # -------------------------------------------------
    def convertir_a_numero(valor):

        if pd.isna(valor):
            return np.nan

        valor = str(valor).strip()

        if valor == "" or valor.upper() == "N/A":
            return np.nan

        valor = valor.replace("%", "").replace(",", ".")

        try:
            return float(valor)
        except:
            return np.nan

    # -------------------------------------------------
    # TRIMESTRES
    # -------------------------------------------------
    t1 = ["ENE-25","FEB-25","MAR-25"]
    t2 = ["ABR-25","MAY-25","JUN-25"]
    t3 = ["JUL-25","AGO-25","SEPT-25"]
    t4 = ["OCT-25","NOV-25","DIC-25"]

    def promedio_trimestre(df, meses):

        meses_existentes = [m for m in meses if m in df.columns]
        if not meses_existentes:
            return pd.Series(np.nan, index=df.index)

        df_trim = df[meses_existentes].copy()

        for col in df_trim.columns:
            df_trim[col] = df_trim[col].apply(convertir_a_numero)

        return df_trim.mean(axis=1).round(2)

    modelo["T1"] = promedio_trimestre(modelo, t1)
    modelo["T2"] = promedio_trimestre(modelo, t2)
    modelo["T3"] = promedio_trimestre(modelo, t3)
    modelo["T4"] = promedio_trimestre(modelo, t4)

    # -------------------------------------------------
    # VALORACIÓN
    # -------------------------------------------------
    def calcular_valoracion(codigo, medicion):

        if pd.isna(medicion):
            return ""

        regla = reglas_indicadores.get(codigo)
        if not regla:
            return ""

        meta = regla["meta"]

        if regla["op"] == ">=":
            cumple = medicion >= meta
        else:
            cumple = medicion > meta

        return "Cumple" if cumple else "No cumple"

    modelo["VAL_T1"] = [
        calcular_valoracion(c, m)
        for c, m in zip(modelo["CONSE"], modelo["T1"])
    ]
    modelo["VAL_T2"] = [
        calcular_valoracion(c, m)
        for c, m in zip(modelo["CONSE"], modelo["T2"])
    ]
    modelo["VAL_T3"] = [
        calcular_valoracion(c, m)
        for c, m in zip(modelo["CONSE"], modelo["T3"])
    ]
    modelo["VAL_T4"] = [
        calcular_valoracion(c, m)
        for c, m in zip(modelo["CONSE"], modelo["T4"])
    ]

    # FORMATO %
    for col in ["T1","T2","T3","T4"]:
        modelo[col] = modelo[col].apply(
            lambda x: f"{x:.2f}%" if pd.notna(x) else ""
        )
        
    # -------------------------------------------------
    # DATOS ANUALES (DESDE BANCO)
    # -------------------------------------------------
    modelo["MEDICION_ANUAL"] = modelo["VALOR ANUAL"]

    modelo["VALORACION_ANUAL"] = (
        modelo["VALORACIÓN"]
        .astype(str)
        .str.strip()
        .str.title()
    )
    
    

    # -------------------------------------------------
    # CREAR TABLA RESUMEN
    # -------------------------------------------------
    resumen = pd.DataFrame({

        ("","CONSE"): modelo["CONSE"],
        ("","ÁREA RESPONSABLE"): modelo["ÁREA"],
        ("","TIPO DE INDICADOR"): modelo["TIPO DE INDICADOR"],
        ("","INDICADOR"): modelo["INDICADOR"],

        ("Registro trimestre 1","Meta"): modelo["META"],
        ("Registro trimestre 1","Medición"): modelo["T1"],
        ("Registro trimestre 1","Valoración"): modelo["VAL_T1"],

        ("Registro trimestre 2","Meta"): modelo["META"],
        ("Registro trimestre 2","Medición"): modelo["T2"],
        ("Registro trimestre 2","Valoración"): modelo["VAL_T2"],

        ("Registro trimestre 3","Meta"): modelo["META"],
        ("Registro trimestre 3","Medición"): modelo["T3"],
        ("Registro trimestre 3","Valoración"): modelo["VAL_T3"],

        ("Registro trimestre 4","Meta"): modelo["META"],
        ("Registro trimestre 4","Medición"): modelo["T4"],
        ("Registro trimestre 4","Valoración"): modelo["VAL_T4"],
        
        ("Registro anual","Medición anual"): modelo["MEDICION_ANUAL"],
        ("Registro anual","Valoración anual"): modelo["VALORACION_ANUAL"],
    })

    resumen.columns = pd.MultiIndex.from_tuples(resumen.columns)

    resumen = resumen.sort_values(
        by=[("","ÁREA RESPONSABLE")]
    ).reset_index(drop=True)

    # -------------------------------------------------
    # VALORACIÓN GENERAL
    # -------------------------------------------------
    def resumen_trimestre(col):

        total = resumen[col].isin(["Cumple","No cumple"]).sum()
        cumple = (resumen[col] == "Cumple").sum()
        no_cumple = (resumen[col] == "No cumple").sum()

        porcentaje = (cumple/total*100) if total > 0 else 0

        return cumple, no_cumple, round(porcentaje,2)

    c1,n1,p1 = resumen_trimestre(("Registro trimestre 1","Valoración"))
    c2,n2,p2 = resumen_trimestre(("Registro trimestre 2","Valoración"))
    c3,n3,p3 = resumen_trimestre(("Registro trimestre 3","Valoración"))
    c4,n4,p4 = resumen_trimestre(("Registro trimestre 4","Valoración"))

    cumple_anual = c1+c2+c3+c4
    no_cumple_anual = n1+n2+n3+n4
    total_anual = cumple_anual + no_cumple_anual

    porcentaje_anual = round(
        (cumple_anual/total_anual*100),2
    ) if total_anual>0 else 0

    estado_anual = "Cumple" if porcentaje_anual >= 50 else "No cumple"

    # -------------------------------------------------
    # FILAS RESUMEN (DEBAJO DE META)
    # -------------------------------------------------

    fila_vacia = {col: "" for col in resumen.columns}

    fila_cumple = fila_vacia.copy()
    fila_no = fila_vacia.copy()
    fila_pct = fila_vacia.copy()


    def llenar_trimestre(trimestre, c, n, p):

        # ---- COLUMNA META LLEVA EL TEXTO ----
        fila_cumple[(trimestre,"Meta")] = "Cumple"
        fila_no[(trimestre,"Meta")] = "No Cumple"
        fila_pct[(trimestre,"Meta")] = "Porcentaje de cumplimiento"

        # ---- VALORACIÓN LLEVA LOS NÚMEROS ----
        fila_cumple[(trimestre,"Valoración")] = c
        fila_no[(trimestre,"Valoración")] = n
        fila_pct[(trimestre,"Valoración")] = f"{p:.2f}%"

        # ---- MEDICIÓN VACÍA ----
        fila_cumple[(trimestre,"Medición")] = ""
        fila_no[(trimestre,"Medición")] = ""
        fila_pct[(trimestre,"Medición")] = ""


    # Llenar cada trimestre
    llenar_trimestre("Registro trimestre 1", c1, n1, p1)
    llenar_trimestre("Registro trimestre 2", c2, n2, p2)
    llenar_trimestre("Registro trimestre 3", c3, n3, p3)
    llenar_trimestre("Registro trimestre 4", c4, n4, p4)


    # -------------------------------------------------
    # VALORACIÓN GENERAL (ANUAL)
    # -------------------------------------------------

    # Fila Cumple
    fila_cumple[("Registro anual","Medición anual")] = "Cumple"
    fila_cumple[("Registro anual","Valoración anual")] = cumple_anual

    # Fila No Cumple
    fila_no[("Registro anual","Medición anual")] = "No cumple"
    fila_no[("Registro anual","Valoración anual")] = no_cumple_anual

    # Fila Porcentaje
    fila_pct[("Registro anual","Medición anual")] = "Porcentaje de cumplimiento anual"
    fila_pct[("Registro anual","Valoración anual")] = f"{porcentaje_anual:.2f}%"

    # -------------------------------------------------
    # UNIR AL DATAFRAME
    # -------------------------------------------------

    resumen_general = pd.DataFrame([fila_cumple, fila_no, fila_pct])
    resumen = pd.concat([resumen, resumen_general], ignore_index=True)

    return resumen