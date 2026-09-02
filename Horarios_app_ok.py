import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import os
import json

st.set_page_config(page_title="Generador de Horarios - City Market", layout="wide")

# Archivo local para almacenar el historial
HISTORIAL_FILE = "historial_horarios.json"

def cargar_historial():
    if os.path.exists(HISTORIAL_FILE):
        try:
            with open(HISTORIAL_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def guardar_historial(historial):
    with open(HISTORIAL_FILE, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=4)

st.title("Generador de Horarios Oficial - City Market")

# Cargar el catálogo de empleados con caché robusta
@st.cache_data
def cargar_catalogo():
    try:
        df = pd.read_excel("catalogo empleados.xlsx")
        # Asegurar que la columna Empleado sea string limpia
        df['Empleado'] = df['Empleado'].astype(str).str.strip().str.replace('.0', '', regex=False)
        return df
    except Exception as e:
        return None

df_catalogo = cargar_catalogo()

if df_catalogo is not None:
    dept_df = df_catalogo[['Clave Departamento', 'Departamento']].drop_duplicates().reset_index(drop=True)
    lista_departamentos = dept_df['Departamento'].tolist()
else:
    lista_departamentos = ["GERENCIA", "ABARROTES", "LACTEOS", "RECURSOS HUMANOS"]

# 1. Datos Generales
st.subheader("1. Datos Generales")
col1, col2, col3 = st.columns(3)

with col1:
    departamento_seleccionado = st.selectbox("Seleccione el Departamento", lista_departamentos)
    
    no_departamento = ""
    if df_catalogo is not None:
        match_dept = dept_df[dept_df['Departamento'] == departamento_seleccionado]
        if not match_dept.empty:
            no_departamento = str(match_dept.iloc[0]['Clave Departamento'])

with col2:
    no_depto_input = st.text_input("Número de Departamento", value=no_departamento)
with col3:
    fecha_entrega = st.date_input("Fecha de Entrega", datetime.today())

# Cargar historial
historial = cargar_historial()

if 'depto_actual' not in st.session_state or st.session_state.depto_actual != departamento_seleccionado:
    st.session_state.depto_actual = departamento_seleccionado
    if departamento_seleccionado in historial:
        st.session_state.empleados = historial[departamento_seleccionado].copy()
    else:
        st.session_state.empleados = []

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

dias_keys = ["Miércoles", "Jueves", "Viernes", "Sábado", "Domingo", "Lunes", "Martes"]

def ajustar_horario_lactancia(horario_str):
    """Resta una hora a la hora de salida si es un turno normal."""
    if horario_str in ["Descanso", "Vacaciones"] or " - " not in horario_str:
        return horario_str
    try:
        partes = horario_str.split(" (")
        rango = partes[0]
        tipo = f" ({partes[1]}" if len(partes) > 1 else ""
        
        entrada, salida = rango.split(" - ")
        h_salida, m_salida = map(int, salida.split(":"))
        
        h_salida_nueva = (h_salida - 1) % 24
        salida_ajustada = f"{h_salida_nueva:02d}:{m_salida:02d}"
        return f"{entrada} - {salida_ajustada}{tipo} (LACTANCIA)"
    except:
        return horario_str

def calcular_sugerencia_comida(horario_str):
    """Calcula una hora de comida sugerida exactamente 4 horas después de la entrada."""
    if horario_str in ["Descanso", "Vacaciones"] or " - " not in horario_str:
        return "14:00 - 15:00"
    try:
        entrada = horario_str.split(" - ")[0]
        h, m = map(int, entrada.split(":")[:2])
        h_comida_inicio = (h + 4) % 24
        h_comida_fin = (h_comida_inicio + 1) % 24
        return f"{h_comida_inicio:02d}:{m:02d} - {h_comida_fin:02d}:{m:02d}"
    except:
        return "14:00 - 15:00"

# 2. Gestión de Empleados
st.subheader("2. Agregar o Editar Empleado")

c1, c2, c3 = st.columns([1, 2, 1])
with c1:
    no_empleado = st.text_input("No. de Empleado", key="input_no_emp")

with c2:
    # Búsqueda robusta del nombre en el catálogo
    nombre_sugerido = ""
    if df_catalogo is not None and no_empleado:
        emp_buscado = str(no_empleado).strip()
        match_emp = df_catalogo[df_catalogo['Empleado'] == emp_buscado]
        if not match_emp.empty:
            nombre_sugerido = str(match_emp.iloc[0]['Nombre'])
    
    nombre_empleado = st.text_input("Nombre Completo", value=nombre_sugerido, key="input_nombre_emp")

with c3:
    es_lactancia = st.checkbox("Hora de Lactancia (-1 hr salida)", key="input_lactancia")

st.markdown("**Horarios y Comidas por Día (Miércoles a Martes)**")

horarios_dias = {}
comidas_dias = {}

cols_dias = st.columns(7)
for idx, d_key in enumerate(dias_keys):
    with cols_dias[idx]:
        st.markdown(f"*{d_key}*")
        h_sel = st.selectbox(f"Turno {d_key}", horarios_autorizados, key=f"h_{d_key}")
        
        # Aplicar lactancia de forma automática si la casilla está marcada
        if es_lactancia:
            h_sel = ajustar_horario_lactancia(h_sel)
            
        horarios_dias[d_key] = h_sel
        
        # Sugerido dinámico basado en la hora de entrada de ese día
        comida_sugerida = calcular_sugerencia_comida(h_sel)
        comidas_dias[d_key] = st.text_input(f"Comida {d_key}", value=comida_sugerida, key=f"c_{d_key}")

fc1, fc2 = st.columns(2)
with fc1:
    fecha_aviso = st.date_input("Fecha de Aviso", datetime.today(), key="input_f_aviso")
with fc2:
    horario_aviso = st.text_input("Horario de Aviso", key="input_h_aviso")

if st.button("➕ Agregar / Actualizar Empleado en la Tabla", type="primary"):
    if no_empleado and nombre_empleado:
        nuevo_emp = {
            "No. Empleado": no_empleado,
            "Nombre Completo": nombre_empleado,
            "Fecha de Aviso": str(fecha_aviso),
            "Horario de Aviso": horario_aviso,
        }
        for d_key in dias_keys:
            nuevo_emp[d_key] = horarios_dias[d_key]
            nuevo_emp[f"Comida_{d_key}"] = comidas_dias[d_key]
            
        existente = False
        for i, emp in enumerate(st.session_state.empleados):
            if str(emp["No. Empleado"]) == str(no_empleado):
                st.session_state.empleados[i] = nuevo_emp
                existente = True
                break
        if not existente:
            st.session_state.empleados.append(nuevo_emp)
            
        historial[departamento_seleccionado] = st.session_state.empleados
        guardar_historial(historial)
        
        st.success(f"Empleado {nombre_empleado} registrado correctamente.")
    else:
        st.warning("Debe ingresar el número y nombre del empleado.")

# 3. Vista Previa y Exportación
st.subheader("3. Vista Previa y Descarga del Formato Oficial")
if st.session_state.empleados:
    df_preview = pd.DataFrame(st.session_state.empleados)
    st.dataframe(df_preview, use_container_width=True)
    
    col_acc1, col_acc2 = st.columns(2)
    with col_acc1:
        if st.button("Generar y Descargar Archivo Excel"):
            template_file = "FORMATO DE HORARIO NUEVA ACTULIZACION.xlsx"
            wb = openpyxl.load_workbook(template_file)
            ws = wb['FORMATO ']
            
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
                
                for col_idx, d_key in enumerate(dias_keys, start=4):
                    val_horario = emp[d_key]
                    cell = ws.cell(row=row_num, column=col_idx, value=val_horario)
                    
                    if val_horario in ["Descanso", "Vacaciones"]:
                        cell.font = Font(name="Calibri", size=12, bold=True, color="FF000000")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.font = Font(name="Calibri", size=10, bold=False)
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                comidas_resumen = " / ".join([emp.get(f"Comida_{d}", "") for d in dias_keys])
                ws.cell(row=row_num, column=11, value=comidas_resumen)
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
            
    with col_acc2:
        if st.button("Limpiar Tabla de este Departamento"):
            st.session_state.empleados = []
            if departamento_seleccionado in historial:
                del historial[departamento_seleccionado]
                guardar_historial(historial)
            st.rerun()
else:
    st.info("El departamento seleccionado no tiene horarios guardados o la tabla está vacía. Agrega empleados para comenzar.")
