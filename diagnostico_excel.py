import streamlit as st
import pandas as pd
import openpyxl

st.title("🔍 Diagnóstico del Excel")

excel_file = st.file_uploader("Cargar Excel para diagnóstico", type=["xlsx", "xls"])

if excel_file:
    st.success("Excel cargado")
    
    # ===== DIAGNÓSTICO 1: Ver todas las pestañas =====
    st.header("1️⃣ Pestañas encontradas")
    try:
        xls = pd.ExcelFile(excel_file)
        st.write(xls.sheet_names)
    except Exception as e:
        st.error(f"Error: {e}")
    
    # ===== DIAGNÓSTICO 2: Ver encabezados de RESUMEN =====
    st.header("2️⃣ Encabezados de RESUMEN (fila 1)")
    try:
        excel_file.seek(0)
        wb = openpyxl.load_workbook(excel_file)
        
        if "RESUMEN" in wb.sheetnames:
            ws = wb["RESUMEN"]
            
            st.write("**Encabezados exactos (como están en el Excel):**")
            encabezados = {}
            for col in range(1, ws.max_column + 1):
                valor = ws.cell(row=1, column=col).value
                if valor:
                    encabezados[col] = str(valor)
                    st.write(f"- Columna {col}: `{valor}` (tipo: {type(valor).__name__})")
            
            st.write("\n**Encabezados en lowercase:**")
            for col, val in encabezados.items():
                st.write(f"- Columna {col}: `{val.lower()}`")
        else:
            st.error("No se encontró pestaña RESUMEN")
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    # ===== DIAGNÓSTICO 3: Ver columnas de ASISTENCIA =====
    st.header("3️⃣ Columnas de ASISTENCIA")
    try:
        excel_file.seek(0)
        if "ASISTENCIA" in pd.ExcelFile(excel_file).sheet_names:
            # Leer primeras 10 filas sin encabezado para ver la estructura
            df_raw = pd.read_excel(excel_file, sheet_name="ASISTENCIA", header=None, nrows=10)
            
            st.write("**Primeras 10 filas completas (para ver encabezados reales):**")
            st.dataframe(df_raw)
            
            # Buscar en qué fila están los nombres de módulos (MF0969, MF0970, etc.)
            st.write("\n**Buscando filas con 'MF' en el texto:**")
            for idx, row in df_raw.iterrows():
                row_text = ' '.join([str(x) for x in row.values if pd.notna(x)])
                if 'MF' in row_text.upper():
                    st.write(f"- Fila {idx}: {row_text[:200]}")
        else:
            st.error("No se encontró pestaña ASISTENCIA")
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    # ===== DIAGNÓSTICO 4: Ver columnas de CALIFICACIONES =====
    st.header("4️⃣ Columnas de CALIFICACIONES")
    try:
        excel_file.seek(0)
        if "CALIFICACIONES" in pd.ExcelFile(excel_file).sheet_names:
            # Leer primeras 10 filas sin encabezado para ver la estructura
            df_raw = pd.read_excel(excel_file, sheet_name="CALIFICACIONES", header=None, nrows=10)
            
            st.write("**Primeras 10 filas completas (para ver encabezados reales):**")
            st.dataframe(df_raw)
            
            # Buscar en qué fila están los nombres de módulos (MF0969, MF0970, etc.)
            st.write("\n**Buscando filas con 'MF' en el texto:**")
            for idx, row in df_raw.iterrows():
                row_text = ' '.join([str(x) for x in row.values if pd.notna(x)])
                if 'MF' in row_text.upper():
                    st.write(f"- Fila {idx}: {row_text[:200]}")
        else:
            st.error("No se encontró pestaña CALIFICACIONES")
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())