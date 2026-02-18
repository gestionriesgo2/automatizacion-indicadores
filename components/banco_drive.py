import pandas as pd
from openpyxl.utils import get_column_letter

# --------------------------------
# GOOGLE DRIVE (IDS)
# --------------------------------
BANCO_FOLDER_ID   = "1eAkULUJ_suAe_L8quHzYshBpxGqW5Zey"
FICHAS_FOLDER_ID  = "1Ugu1ud21AneX82I6SMMOcQyCRrIr9U4B"
REPORTE_FOLDER_ID = "1tCGd7qZVgW_PvRiW3CrJTgy8XwwmCzeZ"
MANUAL_FILE_ID = "1QnGSEhKdwpFuYfLaKeICbEq-N3vK9o_L"

# ----------------------------
# FUNCIONES AUXILIARES
# ----------------------------
def clean_str(x):
    if x is None:
        return None
    s = str(x).strip()
    return s if s != "" else None


def norm_code(x):
    if x is None:
        return None
    return str(x).strip().upper().replace(" ", "")


def construir_formula_excel(row_idx, col_valor, formula_base):

    if formula_base is None:
        return None

    if isinstance(formula_base, float):
        return None

    formula_base = str(formula_base).strip()

    if formula_base == "":
        return None

    # Garantizar =
    if not formula_base.startswith("="):
        formula_base = "=" + formula_base

    letra_col = get_column_letter(col_valor + 1)
    return formula_base.replace("{VALOR}", f"{letra_col}{row_idx}")


# ----------------------------
# CARGAR / CREAR BANCO
# ----------------------------
def cargar_banco_drive(
    get_file_id_by_name,
    read_excel_from_drive
):

    BANCO_FILENAME = "Banco_Indicadores_BASE.xlsx"

    print("🔄 Verificando banco de indicadores...")

    BANCO_FILE_ID = get_file_id_by_name(BANCO_FOLDER_ID, BANCO_FILENAME)

    columnas_banco = [
        "ÁREA", "CONSE", "INDICADOR", "ESTADO DEL INDICADOR", "PROCESO",
        "OBJETIVO-DESCRIPCIÓN", "ORIGEN", "FÓRMULA", "FUENTE NUMERADOR",
        "FUENTE DENOMINADOR", "JERARQUÍA", "NORMA RELACIONADA",
        "TIPO DE INDICADOR", "TENDENCIA", "PERIODICIDAD MEDICION",
        "PERIODICIDAD ANÁLISIS", "OBSERVACIONES",
        "Critico", "Aceptable", "Satisfactorio",
        "DOCUMENTADO", "DE SEG CONTRACTUAL", "REVISADOS",
        "ene-25", "feb-25", "mar-25", "abr-25", "may-25", "jun-25",
        "jul-25", "ago-25", "sept-25", "oct-25", "nov-25", "dic-25",
        "VALOR ANUAL", "VALORACIÓN"
    ]

    if BANCO_FILE_ID:
        banco_stream = read_excel_from_drive(BANCO_FILE_ID)
        banco = pd.read_excel(banco_stream, dtype=str, keep_default_na=False)
    else:
        banco = pd.DataFrame(columns=columnas_banco)
        BANCO_FILE_ID = None

    banco["CONSE"] = banco["CONSE"].apply(norm_code)

    return banco, BANCO_FILE_ID


# ----------------------------
# CARGAR ARCHIVO MANUAL
# ----------------------------
def cargar_datos_manuales(read_excel_from_drive):

    print("📥 Leyendo archivo manual...")

    stream = read_excel_from_drive(MANUAL_FILE_ID)
    df = pd.read_excel(stream, dtype=str, keep_default_na=False)

    df["Cod. Indicador"] = df["Cod. Indicador"].apply(norm_code)

    columnas = {
        "Cod. Indicador": "CONSE",
        "ESTADO DEL INDICADOR": "ESTADO DEL INDICADOR",
        "ORIGEN": "ORIGEN",
        "DOCUMENTADO": "DOCUMENTADO",
        "DE SEG CONTRACTUAL": "DE SEG CONTRACTUAL",
        "REVISADOS": "REVISADOS",
        
    }

    df = df[list(columnas.keys())].rename(columns=columnas)

    return df

def convertir_formula_es_en(formula):

    if formula is None:
        return None

    # ⚠️ Si viene como float (NaN, etc.), no hay fórmula
    if isinstance(formula, float):
        return None

    formula = str(formula).strip()

    if formula == "":
        return None

    reemplazos = {
        "SI(": "IF(",
        "O(": "OR(",
        ";": ","
    }

    for es, en in reemplazos.items():
        formula = formula.replace(es, en)

    return formula



# ----------------------------
# UNIR BANCO + MANUAL
# ----------------------------
def unir_datos_manuales(banco, manual_df):

    print("🔗 Cruzando datos manuales...")

    banco = banco.merge(
        manual_df,
        on="CONSE",
        how="left",
        suffixes=("", "_MANUAL")
    )

    columnas_simple = [
        "ESTADO DEL INDICADOR",
        "ORIGEN",
        "DOCUMENTADO",
        "DE SEG CONTRACTUAL",
        "REVISADOS"
    ]

    for col in columnas_simple:
        col_m = f"{col}_MANUAL"
        if col_m in banco.columns:
            banco[col] = banco[col_m].where(banco[col_m] != "", banco[col])
            banco.drop(columns=[col_m], inplace=True)




    return banco
