import streamlit as st
import pandas as pd
import io
import json
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from streamlit_gsheets import GSheetsConnection

st.set_page_config(
    page_title="Generador de Informes - Instituto Acuarela (A-1090)",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header {
        background-color: #0e7490;
        color: white;
        padding: 1.5rem;
        border-radius: 0.75rem;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

ALL_AREAS = [
    "Prácticas del Lenguaje",
    "Matemática",
    "Cs. Sociales",
    "Cs. Naturales",
    "Area Social y Emocional",
    "Actividades de la Vida Diaria",
    "Taller de Psicopedagogía",
    "Taller de Comunicación",
    "Taller de Psicomotricidad",
    "Taller de Actividades Prácticas",
    "Taller de Musicoterapia",
    "Equipo de Orientación Escolar",
    "Objetivos para el próximo año"
]

# Conexión a Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# Cargar datos desde Google Sheets (o iniciar vacío si la planilla está vacía)
try:
    df_db = conn.read(worksheet="Informes", ttl=0)
    if df_db.empty:
        reports_db = []
    else:
        # Reconstruir la base de datos desde el DataFrame de Google Sheets
        reports_db = []
        for _, row in df_db.iterrows():
            try:
                reports_db.append({
                    "id": int(row["id"]),
                    "nombre": str(row["nombre"]),
                    "dni": str(row["dni"]),
                    "grupo": str(row["grupo"]),
                    "nivel": str(row["nivel"]),
                    "tipo_informe": str(row["tipo_informe"]),
                    "fecha": str(row["fecha"]),
                    "status": str(row["status"]),
                    "updated_at": str(row["updated_at"]),
                    "areas": json.loads(row["areas_json"])
                })
            except Exception:
                continue
    st.session_state.reports_db = reports_db
except Exception as e:
    st.session_state.reports_db = []

if "current_report" not in st.session_state:
    st.session_state.current_report = {
        "id": None,
        "nombre": "",
        "dni": "",
        "grupo": "",
        "nivel": "",
        "tipo_informe": "Final",
        "fecha": "Noviembre 2027",
        "status": "Borrador",
        "updated_at": "",
        "areas": {area: {"texto": "", "docente": ""} for area in ALL_AREAS}
    }

st.markdown("""
    <div class="main-header">
        <h1 style="margin:0; font-size: 1.75rem;">Instituto Acuarela (A-1090)</h1>
        <p style="margin:0; color: #cffafe; font-size: 0.95rem;">Sistema Integrado de Informes Pedagógicos Colaborativos en la Nube</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("📂 Gestión y Archivos")
    
    uploaded_excel = st.file_uploader("Cargar Excel Base (Alumnos)", type=["xlsx", "xls"])
    if uploaded_excel:
        try:
            xls = pd.ExcelFile(uploaded_excel)
            df = pd.read_excel(xls, sheet_name=0, header=None)
            st.success("¡Excel cargado correctamente!")
            sample_nombre = str(df.iloc[12, 1]) if df.shape[0] > 12 and df.shape[1] > 1 else "De la Vega, Diego"
            sample_dni = str(df.iloc[14, 1]) if df.shape[0] > 14 and df.shape[1] > 1 else "45875951"
            
            if st.button("Cargar Alumno desde Excel"):
                st.session_state.current_report.update({
                    "id": int(datetime.now().timestamp()),
                    "nombre": sample_nombre,
                    "dni": sample_dni,
                    "grupo": "2",
                    "nivel": "Primaria"
                })
                st.rerun()
        except Exception as e:
            st.error(f"Error al leer el Excel: {e}")

    st.divider()
    st.subheader(f"Borradores en la Nube ({len(st.session_state.reports_db)})")
    
    search_query = st.text_input("🔍 Buscar por Alumno o DNI", placeholder="Escribe nombre...")
    filtered_reports = st.session_state.reports_db
    if search_query:
        filtered_reports = [r for r in st.session_state.reports_db if search_query.lower() in r["nombre"].lower() or search_query in r["dni"]]
    
    if not filtered_reports:
        st.info("No hay registros en la nube.")
    else:
        for r in filtered_reports:
            badge_color = "🟢" if r["status"] == "Finalizado" else "🟠"
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.markdown(f"**{badge_color} {r['nombre']}**<br><span style='font-size:0.8rem; color:gray;'>DNI: {r['dni']}</span>", unsafe_allow_html=True)
            with col_b:
                if st.button("Retomar", key=f"load_{r['id']}"):
                    st.session_state.current_report = r.copy()
                    st.rerun()

    st.divider()
    if st.button("➕ Nuevo Informe / Limpiar", type="secondary"):
        st.session_state.current_report = {
            "id": None, "nombre": "", "dni": "", "grupo": "", "nivel": "",
            "tipo_informe": "Final", "fecha": "Noviembre 2027", "status": "Borrador", "updated_at": "",
            "areas": {area: {"texto": "", "docente": ""} for area in ALL_AREAS}
        }
        st.rerun()

st.markdown("### Datos de Identificación del Alumno")

with st.form("report_form"):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        nombre_input = st.text_input("Apellido y Nombre del Alumno *", value=st.session_state.current_report["nombre"], placeholder="Ej. De la Vega, Diego")
    with col2:
        dni_input = st.text_input("DNI *", value=st.session_state.current_report["dni"], placeholder="Ej. 45875951")
    with col3:
        col_g, col_n = st.columns(2)
        with col_g:
            grupo_input = st.text_input("Grupo *", value=st.session_state.current_report["grupo"], placeholder="Ej. 2")
        with col_n:
            nivel_input = st.text_input("Nivel *", value=st.session_state.current_report["nivel"], placeholder="Ej. Primaria")

    col4, col5, col6 = st.columns(3)
    with col4:
        tipo_informe_input = st.selectbox("Tipo de Informe", ["Semestral", "Final"], index=0 if st.session_state.current_report["tipo_informe"]=="Semestral" else 1)
    with col5:
        fecha_input = st.selectbox("Período / Fecha", ["Junio 2027", "Noviembre 2027"], index=0 if st.session_state.current_report["fecha"]=="Junio 2027" else 1)
    with col6:
        st.text_input("Plan de Estudios", value="1332/92", disabled=True)

    st.markdown("---")
    st.markdown("### Redacción de Áreas Pedagógicas")

    updated_areas = {}
    for idx, area_name in enumerate(ALL_AREAS):
        st.markdown(f"#### Área {idx + 1}: {area_name}")
        col_doc, col_space = st.columns([2, 1])
        with col_doc:
            docente_val = st.text_input(f"Docente a Cargo ({area_name})", value=st.session_state.current_report["areas"][area_name]["docente"], key=f"doc_{idx}")
        
        texto_val = st.text_area(f"Informe Pedagógico ({area_name})", value=st.session_state.current_report["areas"][area_name]["texto"], key=f"txt_{idx}", height=120)
        
        updated_areas[area_name] = {"texto": texto_val, "docente": docente_val}
        st.markdown("<br>", unsafe_allow_html=True)

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        submit_draft = st.form_submit_button("💾 Guardar Avance / Borrador", use_container_width=True)
    with col_btn2:
        submit_final = st.form_submit_button("✔️ Guardar como Finalizado", use_container_width=True)

    if submit_draft or submit_final:
        if not nombre_input.strip() or not dni_input.strip() or not grupo_input.strip() or not nivel_input.strip():
            st.error("⚠️ No se permite el guardado por falta de datos del alumno")
        else:
            status_val = "Finalizado" if submit_final else "Borrador"
            now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
            record_id = st.session_state.current_report["id"] or int(datetime.now().timestamp())
            
            saved_record = {
                "id": record_id,
                "nombre": nombre_input.strip(),
                "dni": dni_input.strip(),
                "grupo": grupo_input.strip(),
                "nivel": nivel_input.strip(),
                "tipo_informe": tipo_informe_input,
                "fecha": fecha_input,
                "status": status_val,
                "updated_at": now_str,
                "areas": updated_areas
            }
            
            # Actualizar en la lista local y guardar en Google Sheets
            db_list = st.session_state.reports_db
            existing_idx = next((i for i, r in enumerate(db_list) if r["id"] == record_id), None)
            if existing_idx is not None:
                db_list[existing_idx] = saved_record
            else:
                db_list.insert(0, saved_record)
                
            # Convertir toda la base a DataFrame para guardarla en Google Sheets
            rows_to_save = []
            for r in db_list:
                rows_to_save.append({
                    "id": r["id"],
                    "nombre": r["nombre"],
                    "dni": r["dni"],
                    "grupo": r["grupo"],
                    "nivel": r["nivel"],
                    "tipo_informe": r["tipo_informe"],
                    "fecha": r["fecha"],
                    "status": r["status"],
                    "updated_at": r["updated_at"],
                    "areas_json": json.dumps(r["areas"])
                })
            
            df_to_write = pd.DataFrame(rows_to_save)
            conn.update(worksheet="Informes", data=df_to_write)
            
            st.session_state.current_report = saved_record
            st.success(f"¡{status_val} guardado y sincronizado en la nube a las {now_str}!")

st.markdown("---")
st.subheader("📥 Exportación de Documento Oficial")

def generate_word_document(report_data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("INSTITUTO ACUARELA (A-1090)")
    r_title.bold = True
    r_title.font.name = "Arial"
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = RGBColor(14, 116, 144)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Escuela de Educación Especial Incorporada a la Enseñanza Oficial – CUE 020026900")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph()

    p_meta1 = doc.add_paragraph()
    p_meta1.add_run("Tipo de Informe: ").bold = True
    p_meta1.add_run(f"{report_data['tipo_informe']}    |    ")
    p_meta1.add_run("Período: ").bold = True
    p_meta1.add_run(f"{report_data['fecha']}    |    ")
    p_meta1.add_run("Plan de Estudios: ").bold = True
    p_meta1.add_run("1332/92")

    p_meta2 = doc.add_paragraph()
    p_meta2.add_run("Alumno/a: ").bold = True
    p_meta2.add_run(f"{report_data['nombre']}    |    ")
    p_meta2.add_run("DNI: ").bold = True
    p_meta2.add_run(f"{report_data['dni']}    |    ")
    p_meta2.add_run("Grupo: ").bold = True
    p_meta2.add_run(f"{report_data['grupo']}    |    ")
    p_meta2.add_run("Nivel: ").bold = True
    p_meta2.add_run(f"{report_data['nivel']}")

    doc.add_paragraph("_________________________________________________________________________________")

    for area_name, content in report_data["areas"].items():
        if content["texto"].strip() or content["docente"].strip():
            h2 = doc.add_heading(area_name, level=2)
            h2.style.font.name = "Arial"
            
            if content["docente"].strip():
                p_doc = doc.add_paragraph()
                p_doc.add_run("Docente a Cargo: ").bold = True
                p_doc.add_run(content["docente"])
                p_doc.paragraph_format.line_spacing = 1.15
            
            if content["texto"].strip():
                p_txt = doc.add_paragraph(content["texto"])
                p_txt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFIED
                p_txt.paragraph_format.line_spacing = 1.15

    doc.add_paragraph()
    p_foot1 = doc.add_paragraph()
    p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = p_foot1.add_run("Pringles 738/40 (1183) – CABA – Te: 4862-5307 – 4864-7572 – 4865-2256")
    r_f1.font.size = Pt(8)
    r_f1.font.color.rgb = RGBColor(100, 116, 139)

    p_foot2 = doc.add_paragraph()
    p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f2 = p_foot2.add_run("E-mail: institutoacuarela@yahoo.com.ar – www.acuarela1090.com")
    r_f2.font.size = Pt(8)
    r_f2.font.color.rgb = RGBColor(100, 116, 139)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.session_state.current_report["nombre"]:
    docx_file = generate_word_document(st.session_state.current_report)
    safe_name = st.session_state.current_report['nombre'].replace(" ", "_").replace(",", "")
    st.download_button(
        label="📄 Descargar Informe en Word (.docx)",
        data=docx_file,
        file_name=f"Informe_{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
else:
    st.warning("Completa al menos el Nombre del alumno para habilitar la descarga en Word.")
