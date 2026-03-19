import pandas as pd



def generar_resumenes(banco):
    # =====================================================
    # 🔥 FILTRAR SOLO INDICADORES ACTIVOS
    # =====================================================
    banco_activos = banco[
        banco["ESTADO DEL INDICADOR"]
        .astype(str)
        .str.strip()
        .str.upper() == "ACTIVO"
    ].copy()

    # Normalizar periodicidad
    banco_activos["PERIODICIDAD MEDICION"] = (
        banco_activos["PERIODICIDAD MEDICION"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # =====================================================
    # 1️⃣ RESUMEN POR ÁREA
    # =====================================================

    resumen_area = pd.pivot_table(
        banco_activos,
        index="ÁREA",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    resumen_area["TOTAL"] = resumen_area.sum(axis=1)
    resumen_area.loc["TOTAL GENERAL"] = resumen_area.sum()
    resumen_area = resumen_area.reset_index()

    # =====================================================
    # 2️⃣ RESUMEN POR ESTADO
    # =====================================================

    resumen_estado = pd.pivot_table(
        banco_activos,
        index="ESTADO DEL INDICADOR",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    resumen_estado["TOTAL"] = resumen_estado.sum(axis=1)
    resumen_estado.loc["TOTAL GENERAL"] = resumen_estado.sum()
    resumen_estado = resumen_estado.reset_index()

    # =====================================================
    # 3️⃣ RESUMEN PERIODICIDAD + FICHAS
    # =====================================================

    pivot_conteo = pd.pivot_table(
        banco_activos,
        index="ÁREA",
        columns="PERIODICIDAD MEDICION",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    pivot_codigos = pd.pivot_table(
        banco_activos,
        index="ÁREA",
        columns="PERIODICIDAD MEDICION",
        values="CONSE",
        aggfunc=lambda x: ", ".join(sorted(x.astype(str).unique())),
        fill_value=""
    )

    pivot_codigos.columns = [f"{col} - FICHAS" for col in pivot_codigos.columns]

    resumen_periodicidad = pd.concat([pivot_conteo, pivot_codigos], axis=1)

    resumen_periodicidad["TOTAL"] = resumen_periodicidad.select_dtypes(
        include="number"
    ).sum(axis=1)

    totales = resumen_periodicidad.select_dtypes(include="number").sum()
    resumen_periodicidad.loc["TOTAL GENERAL"] = totales

    resumen_periodicidad = resumen_periodicidad.reset_index()

    # =====================================================
    # 4️⃣ RESUMEN GENERAL
    # =====================================================

    resumen_general = pd.DataFrame({
        "Indicador": ["Total Indicadores Activos"],
        "TOTAL": [len(banco_activos)]
    })

    # =====================================================
    # 5️⃣ RESUMEN CUMPLE / NO CUMPLE + PERIODICIDAD
    # =====================================================

    resumen_cumple = pd.pivot_table(
        banco_activos,
        index=["ÁREA", "PERIODICIDAD MEDICION"],
        columns="ESTADO DEL INDICADOR",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    for col in ["CUMPLE", "NO CUMPLE"]:
        if col not in resumen_cumple.columns:
            resumen_cumple[col] = 0

    resumen_cumple["TOTAL"] = resumen_cumple.sum(axis=1)

    totales = resumen_cumple.sum()
    resumen_cumple.loc[("TOTAL GENERAL", ""), :] = totales

    resumen_cumple = resumen_cumple.reset_index()
    
    
        # =====================================================
    # 5️⃣ RESUMEN POR JERARQUÍA
    # =====================================================

    resumen_jerarquia = pd.pivot_table(
        banco_activos,
        index="ÁREA",
        columns="JERARQUÍA",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    resumen_jerarquia["TOTAL"] = resumen_jerarquia.sum(axis=1)
    resumen_jerarquia.loc["TOTAL GENERAL"] = resumen_jerarquia.sum()
    resumen_jerarquia = resumen_jerarquia.reset_index()
    
    
        # =====================================================
    # 6️⃣ RESUMEN TIPO DE INDICADOR
    # =====================================================

    resumen_tipo = pd.pivot_table(
        banco_activos,
        index="TIPO DE INDICADOR",
        columns="ÁREA",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    resumen_tipo["TOTAL"] = resumen_tipo.sum(axis=1)
    resumen_tipo.loc["TOTAL GENERAL"] = resumen_tipo.sum()
    resumen_tipo = resumen_tipo.reset_index()


    # =====================================================
    # RETORNAR COMO DICCIONARIO (MEJOR PRÁCTICA)
    # =====================================================

    return {
        "resumen_area": resumen_area,
        "resumen_estado": resumen_estado,
        "resumen_periodicidad": resumen_periodicidad,
        "resumen_general": resumen_general,
        "resumen_jerarquia": resumen_jerarquia,
        "resumen_tipo": resumen_tipo,
        "resumen_cumple": resumen_cumple,
        
    }