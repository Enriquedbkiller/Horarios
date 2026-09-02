import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
import io

st.set_page_config(page_title="Generador de Horarios - City Market", layout="wide")

st.title("Generador de Horarios Oficial - City Market")

# Cargar el catálogo de empleados automáticamente
@st.cache_data
def cargar_catalogo():
    try:
        df = pd.read_excel("catalogo empleados.xlsx")
        return df
    except Exception as e:
        return None

df_catalogo = cargar_catalogo()

if df_catalogo is not None:
    # Obtener lista única de departamentos y un diccionario de correspondencia
    dept_df = df_catalogo[['Clave Departamento', 'Departamento']].drop_duplicates().reset_index(drop=True)
    lista_departamentos = dept_df['Departamento'].tolist()
else:
    lista_departamentos = ["GERENCIA", "ABARROTES", "LACTEOS"]

# 1. Datos Generales
st.subheader("1. Datos Generales")
col1, col2, col3 = st.columns(3)

with col1:
    departamento_seleccionado = st.selectbox("Seleccione el Departamento", lista_departamentos)
    
    # Autocompletar el número de departamento según la selección
    no_departamento = ""
    if df_catalogo is not None:
        match_dept = dept_df[dept_df['Departamento'] == departamento_seleccionado]
        if not match_dept.empty:
            no_departamento = str(match_dept.iloc[0]['Clave Departamento'])

with col2:
    # Campo de texto bloqueado o visible con el número automático
    no_depto_input = st.text_input("Número de Departamento", value=no_departamento)
with col3:
    fecha_entrega = st.date_input("Fecha de Entrega", datetime.today())

# Horarios autorizados en formato de 24 horas
horarios_autorizados = [
    "Descanso",
    "Vacaciones",
    "05:00 - 13:30 (MIXTA)",
    "05:30 - 14:00 (MIXTA)",
    "06:00 - 15:00 (DIURNA)",
    "06:30 - 15:30 (DIURNA)",
    "07:00 - 16:00 (DIURNA)",
    "07:30 - 16:30 (DIURNA)",
    "08:00 - 17:00 (DIURNA)",
    "08:30 - 17:30 (DIURNA)",
    "09:00 - 18:00 (DIURNA)",
    "09:30 - 18:30 (DIURNA)",
    "10:00 - 19:00 (DIURNA)",
    "10:30 - 19:30 (DIURNA)",
    "11:00 - 20:00 (DIURNA)",
    "11:30 - 20:30 (DIURNA)",
    "12:30 - 21:00 (MIXTA)",
    "13:00 - 21:30 (MIXTA)",
    "13:30 - 22:00 (MIXTA)",
    "14:00 - 22:30 (MIXTA)",
    "14:30 - 23:00 (MIXTA)",
    "22:00 - 05:00 (NOCTURNA)",
    "22:30 - 05:30 (NOCTURNA)",
    "23:00 - 06:00 (NOCTURNA)"
]

# 2. Gestión de Empleados en Session State
if 'empleados' not in st.session_state:
    st.session_state.empleados = []

st.subheader("2. Agregar Empleado al Horario")
with st.form("form_empleado", clear_on_submit=True):
    c1, c2 = st.columns(2)
    with c1:
        no_empleado = st.text_input("No. de Empleado (Escribe y autocompleta el nombre)")
    with c2:
        # Buscar el nombre automáticamente si el número existe en el catálogo
        nombre_sugerido = ""
        if df_catalogo is not None and no_empleado:
            match_emp = df_catalogo[df_catalogo['Empleado'].astype(str) == str(no_empleado).strip()]
            if not match_emp.empty:
                nombre_sugerido = str(match_emp.iloc[0]['Nombre'])
        
        nombre_empleado = st.text_input("Nombre Completo", value=nombre_sugerido)
        
    st.markdown("**Horarios Autorizados por Día (Miércoles a Martes)**")
    d1, d2, d3, d4, d5, d6, d7 = st.columns(7)
    
    with d1: h_miercoles = st.selectbox("Miércoles", horarios_autorizados)
    with d2: h_jueves = st.selectbox("Jueves", horarios_autorizados)
    with d3: h_viernes = st.selectbox("Viernes", horarios_autorizados)
    with d4: h_sabado = st.selectbox("Sábado", horarios_autorizados)
    with d5: h_domingo = st.selectbox("Domingo", horarios_autorizados)
    with d6: h_lunes = st.selectbox("Lunes", horarios_autorizados)
    with d7: h_martes = st.selectbox("Martes", horarios_autorizados)
    
    c3, c4, c5 = st.columns(3)
    with c3:
        hora_comida = st.text_input("Hora de Comida (ej. 14:00 - 15:00)")
    with c4:
        fecha_aviso = st.date_input("Fecha de Aviso", datetime.today())
    with c5:
        horario_aviso = st.text_input("Horario de Aviso")
        
    submitted = st.form_submit_button("Agregar Empleado")
    if submitted and no_empleado and nombre_empleado:
        nuevo_emp = {
            "No. Empleado": no_empleado,
            "Nombre Completo": nombre_empleado,
            "Miércoles": h_miercoles,
            "Jueves": h_jueves,
            "Viernes": h_viernes,
            "Sábado": h_sabado,
            "Domingo": h_domingo,
            "Lunes": h_lunes,
            "Martes": h_martes,
            "Hora de Comida": hora_comida,
            "Fecha de Aviso": str(fecha_aviso),
            "Horario de Aviso": horario_aviso,
        }
        st.session_state.empleados.append(nuevo_emp)
        st.success(f"Empleado {nombre_empleado} agregado correctamente.")

# 3. Vista Previa y Exportación
st.subheader("3. Vista Previa y Descarga del Formato Oficial")
if st.session_state.empleados:
    df_preview = pd.DataFrame(st.session_state.empleados)
    st.dataframe(df_preview, use_container_width=True)
    
    if st.button("Generar y Descargar Archivo Excel"):
        template_file = "FORMATO DE HORARIO NUEVA ACTULIZACION.xlsx"
        wb = openpyxl.load_workbook(template_file)
        ws = wb['FORMATO ']
        
        # Insertar Departamento y Número de Departamento
        ws['C4'] = departamento_seleccionado
        ws['J4'] = no_depto_input
        ws['C5'] = str(fecha_entrega)
        
        start_row = 9
        for idx, emp in enumerate(st.session_state.empleados):
            row_num = start_row + idx
            
            if row_num >= 23:
                ws.insert_rows(row_num)
            
            ws.cell(row=row_num, column=2, value=emp["No. Empleado"])
            ws.cell(row=row_num, column=3, value=emp["Nombre Completo"])
            
            dias_keys = ["Miércoles", "Jueves", "Viernes", "Sábado", "Domingo", "Lunes", "Martes"]
            for col_idx, d_key in enumerate(dias_keys, start=4):
                val_horario = emp[d_key]
                cell = ws.cell(row=row_num, column=col_idx, value=val_horario)
                
                if val_horario in ["Descanso", "Vacaciones"]:
                    cell.font = Font(name="Calibri", size=12, bold=True, color="FF000000")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.font = Font(name="Calibri", size=10, bold=False)
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(row=row_num, column=11, value=emp["Hora de Comida"])
            ws.cell(row=row_num, column=12, value=emp["Fecha de Aviso"])
            ws.cell(row=row_num, column=13, value=emp["Horario de Aviso"])
            
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Descargar Horario Rellenado (Excel)",
            data=output,
            file_name=f"Horario_{departamento_seleccionado}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    if st.button("Limpiar Tabla"):
        st.session_state.empleados = []
        st.rerun()
else:
    st.info("Agrega al menos un empleado usando el formulario de arriba para generar tu archivo de Excel con el formato correcto.")
