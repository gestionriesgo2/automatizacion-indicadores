import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

import pandas as pd

def generar_resumenes(banco):

    # =====================================================
    # 1️⃣ RESUMEN POR ÁREA
    # =====================================================
    resumen_area = pd.pivot_table(
        banco,
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
        banco,
        index="ESTADO DEL INDICADOR",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    resumen_estado["TOTAL"] = resumen_estado.sum(axis=1)
    resumen_estado.loc["TOTAL GENERAL"] = resumen_estado.sum()
    resumen_estado = resumen_estado.reset_index()

    # =====================================================
    # 🔥 NORMALIZAR PERIODICIDAD (MUY IMPORTANTE)
    # =====================================================

    banco["PERIODICIDAD MEDICION"] = (
        banco["PERIODICIDAD MEDICION"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # =====================================================
    # 3️⃣ PERIODICIDAD + FICHAS POR CADA PERIODICIDAD
    # =====================================================

    pivot_conteo = pd.pivot_table(
        banco,
        index="ÁREA",
        columns="PERIODICIDAD MEDICION",
        values="CONSE",
        aggfunc="count",
        fill_value=0
    )

    pivot_codigos = pd.pivot_table(
        banco,
        index="ÁREA",
        columns="PERIODICIDAD MEDICION",
        values="CONSE",
        aggfunc=lambda x: ", ".join(sorted(x.astype(str).unique())),
        fill_value=""
    )

    pivot_codigos.columns = [f"{col} - FICHAS" for col in pivot_codigos.columns]

    resumen_periodicidad = pd.concat([pivot_conteo, pivot_codigos], axis=1)

    # =====================================================
    # ORDEN PERSONALIZADO SEGURO
    # =====================================================

    orden_periodicidad = ["MENSUAL", "BIMENSUAL", "TRIMESTRAL", "SEMESTRAL", "ANUAL"]

    columnas_ordenadas = []

    for periodo in orden_periodicidad:
        if periodo in resumen_periodicidad.columns:
            columnas_ordenadas.append(periodo)
        if f"{periodo} - FICHAS" in resumen_periodicidad.columns:
            columnas_ordenadas.append(f"{periodo} - FICHAS")

    # 🔥 Solo reordenar si encontró columnas
    if columnas_ordenadas:
        resumen_periodicidad = resumen_periodicidad[columnas_ordenadas]

    # =====================================================
    # TOTALES
    # =====================================================

    resumen_periodicidad["TOTAL"] = resumen_periodicidad.select_dtypes(include="number").sum(axis=1)

    totales = resumen_periodicidad.select_dtypes(include="number").sum()
    resumen_periodicidad.loc["TOTAL GENERAL"] = totales

    resumen_periodicidad = resumen_periodicidad.reset_index()


    # =====================================================
    # 4️⃣ GENERAL
    # =====================================================
    resumen_general = pd.DataFrame({
        "Indicador": ["Total Indicadores"],
        "TOTAL": [len(banco)]
    })


    # =====================================================
    # 5️⃣ JERARQUÍA
    # =====================================================
    resumen_jerarquia = pd.pivot_table(
        banco,
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
    # 6️⃣ TIPO DE INDICADOR
    # =====================================================
    resumen_tipo = pd.pivot_table(
        banco,
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
    return (
        resumen_area,
        resumen_estado,
        resumen_periodicidad,
        resumen_general,
        resumen_jerarquia,
        resumen_tipo
    )



def guardar_banco_con_estilos_drive(
    banco,
    create_or_update_file,
    banco_file_id,
    banco_folder_id,
    filename="Banco_Indicadores.xlsx"
):
    """
    Aplica estilos al banco y lo guarda en Google Drive.
    """

    print("🎨 Aplicando estilos y guardando Banco en Drive...")
    
    # Generar tablas resumen
    resumen_area, resumen_estado, resumen_periodicidad, resumen_general, resumen_jerarquia, resumen_tipo= generar_resumenes(banco)
    
    # ----------------------------
    # Guardar banco en memoria
    # ----------------------------
    buffer = io.BytesIO()
    banco.to_excel(buffer, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    # ----------------------------
    # Estilos
    # ----------------------------
    header_fill = PatternFill(start_color="A7D08C", end_color="A7D08C", fill_type="solid")
    header_font = Font(bold=True, color="000000")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    columnas_sin_color = ["Q", "R", "S"]

    # Encabezados
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=1, column=col)
        if col_letter not in columnas_sin_color:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

    # Colores especiales
    ws["Q1"].fill = PatternFill(start_color="FF0000", fill_type="solid")
    ws["R1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["S1"].fill = PatternFill(start_color="00FF00", fill_type="solid")

    # Celdas
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Ancho columnas
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    # Altura filas
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 30

    # Filtro
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    
    # ==================================================
    # 📊 AGREGAR HOJAS DE RESUMEN
    # ==================================================



    def agregar_hoja_resumen(nombre, df):
        ws_res = wb.create_sheet(title=nombre)

        # Escribir datos
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws_res.cell(row=r_idx, column=c_idx, value=value)

        # ----------------------------
        # Aplicar estilos iguales al banco
        # ----------------------------

        # Encabezado
        for col in range(1, ws_res.max_column + 1):
            cell = ws_res.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

        # Celdas
        for row in ws_res.iter_rows(min_row=2, max_row=ws_res.max_row, max_col=ws_res.max_column):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        # Auto ancho columnas
        for col in ws_res.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            col_letter = get_column_letter(col[0].column)
            ws_res.column_dimensions[col_letter].width = min(max_len + 2, 40)

        # Filtro
        ws_res.auto_filter.ref = f"A1:{get_column_letter(ws_res.max_column)}{ws_res.max_row}"

        # Congelar fila 1
        ws_res.freeze_panes = "A2"


    # Crear hojas
    agregar_hoja_resumen("Resumen_Area", resumen_area)
    agregar_hoja_resumen("Resumen_Estado", resumen_estado)
    agregar_hoja_resumen("Resumen_Periodicidad", resumen_periodicidad)
    agregar_hoja_resumen("Resumen_General", resumen_general)
    agregar_hoja_resumen("Resumen_jerarquia", resumen_jerarquia)
    agregar_hoja_resumen("Resumen_tipo", resumen_tipo)


    # ----------------------------
    # Subir a Drive
    # ----------------------------
    buffer_out = io.BytesIO()
    wb.save(buffer_out)
    buffer_out.seek(0)

    create_or_update_file(
        bytes_data=buffer_out.getvalue(),
        file_id=banco_file_id,        # None si no existía
        filename=filename,
        parent_folder_id=banco_folder_id
    )
    
    

    print("✅ Banco guardado correctamente en Drive")
