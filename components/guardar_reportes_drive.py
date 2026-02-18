import io
import pandas as pd


def guardar_reportes_drive(
    registros,
    get_file_id_by_name,
    create_or_update_file,
    reporte_folder_id,
    filename_excel="Reporte_Indicadores.xlsx",
    filename_csv="Reporte_Indicadores.csv"
):
    """
    Genera y guarda reportes (Excel y CSV) en Google Drive.
    """

    print("📄 Generando reportes en Drive...")

    # ----------------------------
    # DataFrame de reporte
    # ----------------------------
    columnas_base = ["archivo", "hoja", "codigo", "accion", "ok"]
    df_rep = (
        pd.DataFrame(registros)
        if registros
        else pd.DataFrame(columns=columnas_base)
    )

    # ----------------------------
    # REPORTE EXCEL
    # ----------------------------
    file_id_excel = get_file_id_by_name(reporte_folder_id, filename_excel)

    with io.BytesIO() as buffer_excel:
        with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
            df_rep.to_excel(writer, sheet_name="Reporte Completo", index=False)

            if not df_rep.empty:
                df_rep[df_rep["accion"] == "actualizado"].to_excel(
                    writer, sheet_name="Actualizados", index=False
                )
                df_rep[df_rep["accion"] == "agregado"].to_excel(
                    writer, sheet_name="Agregados", index=False
                )
                df_rep[df_rep["ok"] == False].to_excel(
                    writer, sheet_name="Errores", index=False
                )

        buffer_excel.seek(0)
        create_or_update_file(
            bytes_data=buffer_excel.read(),
            file_id=file_id_excel,
            filename=filename_excel,
            parent_folder_id=reporte_folder_id
        )

    # ----------------------------
    # REPORTE CSV
    # ----------------------------
    file_id_csv = get_file_id_by_name(reporte_folder_id, filename_csv)

    with io.BytesIO() as buffer_csv:
        df_rep.to_csv(buffer_csv, index=False)
        buffer_csv.seek(0)
        create_or_update_file(
            bytes_data=buffer_csv.read(),
            file_id=file_id_csv,
            filename=filename_csv,
            parent_folder_id=reporte_folder_id,
            mimetype="text/csv"
        )

    print("✔ Reportes Excel y CSV creados/actualizados correctamente en Drive ✅")
