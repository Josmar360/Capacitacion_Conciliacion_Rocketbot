import pandas as pd

# ============================
# RUTAS DE ARCHIVOS
# ============================
archivo_auxiliar = "libro_auxiliar.csv"
archivo_banco = "banco.csv"
archivo_salida = "Conciliacion.xlsx"

# ============================
# LEER CSV
# ============================
df_aux = pd.read_csv(archivo_auxiliar)
df_banco = pd.read_csv(archivo_banco)

# ============================
# CREAR ARCHIVO EXCEL
# ============================
with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:

    # Hoja1 - Auxiliar
    df_aux.to_excel(writer, sheet_name="Hoja1", index=False)

    # Hoja2 - Banco
    df_banco.to_excel(writer, sheet_name="Hoja2", index=False)

    # ============================
    # CONCILIACION (LEFT JOIN)
    # ============================
    df_merge = pd.merge(
        df_aux,
        df_banco,
        on="Referencia",
        how="left",
        suffixes=("_Aux", "_Banco")
    )

    # Crear columna vacía (espacio)
    df_merge[""] = ""

    # Construir DataFrame final EXACTO
    df_conciliacion = df_merge[[
        "Fecha_Aux",
        "Referencia",
        "Beneficiario",
        "Cargo",
        "Abono",
        "",  # columna espacio
        "Fecha_Banco",
        "Referencia",
        "Concepto_Banco",
        "Retiro",
        "Deposito"
    ]]

    # Renombrar columnas
    df_conciliacion.columns = [
        "Fecha",
        "Referencia",
        "Beneficiario",
        "Cargo",
        "Abono",
        "",
        "Fecha",
        "Referencia",
        "Concepto_Banco",
        "Retiro",
        "Deposito"
    ]

    # Hoja Conciliacion
    df_conciliacion.to_excel(writer, sheet_name="Conciliacion", index=False)

print("Archivo conciliacion.xlsx creado correctamente.")
