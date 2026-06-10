"""
streamlit_app.py
Interfaz Streamlit para el Generador de Informes Mensuales – DISPOWER
Reutiliza utils/data_processor.py, utils/chart_generator.py, utils/word_generator.py
sin modificar ninguno.
"""

import streamlit as st
import json
import uuid
import shutil
import tempfile
import os
from pathlib import Path

# ── Imports internos ────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "config"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

import sys
sys.path.insert(0, str(BASE_DIR))
from utils.data_processor import DataProcessor, MESES
from utils.chart_generator import generar_todos_los_graficos
from utils.word_generator import WordGenerator

# ── Helpers ─────────────────────────────────────────────────────────────
def list_clients():
    return sorted([p.stem for p in CONFIG_DIR.glob("*.json")])

def load_config(cliente: str) -> dict:
    path = CONFIG_DIR / f"{cliente}.json"
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def fmt_cop(n):
    try:
        return f"$ {int(n):,}".replace(",", ".")
    except:
        return str(n)

def fmt_num(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except:
        return str(n)

# ── Page config ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Generador de Informes | DISPOWER",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizado (misma paleta del Flask) ───────────────────────────
st.markdown("""
<style>
/* Paleta corporativa */
:root {
    --dp-azul: #1B4F8A;
    --dp-azul-claro: #2E86C1;
    --dp-naranja: #F39C12;
}

/* Header principal */
.header-banner {
    background: linear-gradient(135deg, #1B4F8A, #2E86C1);
    color: white;
    padding: 22px 30px;
    border-radius: 12px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.header-banner h1 { margin: 0; font-size: 1.5rem; font-weight: 700; }
.header-banner p  { margin: 4px 0 0 0; font-size: .85rem; opacity: .85; }

/* Tarjetas de sección */
.section-card {
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,.07);
    border-left: 4px solid #1B4F8A;
}
.section-card h3 {
    color: #1B4F8A;
    font-size: 1rem;
    font-weight: 700;
    margin: 0 0 16px 0;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* Métricas resultado */
.metric-row {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
    margin: 16px 0;
}
.metric-box {
    flex: 1;
    min-width: 140px;
    background: linear-gradient(135deg, #1B4F8A, #2E86C1);
    color: white;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
}
.metric-box .num { font-size: 1.8rem; font-weight: 700; line-height: 1; }
.metric-box .lbl { font-size: .78rem; opacity: .85; margin-top: 4px; }

/* Capítulos árbol */
.cap-tree-item { padding: 2px 0; font-size: .9rem; }

/* Badge inconsistencia */
.badge-warn {
    background: #FDEBD0;
    color: #D35400;
    border-radius: 20px;
    padding: 5px 14px;
    font-size: .85rem;
    font-weight: 600;
    display: inline-block;
    margin-top: 10px;
}
.badge-ok {
    background: #D5F5E3;
    color: #1E8449;
    border-radius: 20px;
    padding: 5px 14px;
    font-size: .85rem;
    font-weight: 600;
    display: inline-block;
    margin-top: 10px;
}

/* Upload zone personalizada */
[data-testid="stFileUploader"] {
    border: 2px dashed #2E86C1 !important;
    border-radius: 10px !important;
    background: #F8FBFF !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #1B4F8A !important;
    background: #EBF5FB !important;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #F4F6F9;
}
[data-testid="stSidebar"] .stMarkdown h3 {
    color: #1B4F8A;
    font-size: .95rem;
}

/* Botón primario */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1B4F8A, #2E86C1) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 12px 32px !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    width: 100%;
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 18px rgba(27,79,138,.35) !important;
}

/* Línea separadora */
hr { border-color: #DEE2E6; margin: 20px 0; }

footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-banner">
  <div>
    <h1>📄 Generador de Informes Mensuales</h1>
    <p>DISPOWER S.A.S E.S.P. — Sistema automatizado de informes de concesión</p>
  </div>
  <div style="text-align:right; opacity:.8; font-size:.8rem;">
    Operaciones · SAC · Facturación · HSEQ
  </div>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════
# SIDEBAR — Capítulos
# ════════════════════════════════════════════════════════════════════════
clientes = list_clients()

with st.sidebar:
    st.markdown("### ⚙️ Parámetros del informe")

    cliente = st.selectbox(
        "Cliente",
        clientes,
        index=clientes.index("DISPAC") if "DISPAC" in clientes else 0,
        help="Cada cliente tiene su propia configuración de capítulos y estructura"
    )

    config = load_config(cliente)
    capitulos_config = config.get("capitulos", [])

    col1, col2 = st.columns(2)
    mes = col1.selectbox("Mes", list(MESES.keys()),
                         format_func=lambda x: MESES[x],
                         index=3)   # Abril por defecto
    ano = col2.number_input("Año", min_value=2020, max_value=2040,
                            value=2026, step=1)

    st.markdown("---")
    st.markdown("### 📋 Capítulos a incluir")

    # Botones Todos / Ninguno
    c1, c2 = st.columns(2)
    sel_todos   = c1.button("✅ Todos",   use_container_width=True)
    sel_ninguno = c2.button("⬜ Ninguno", use_container_width=True)

    # Inicializar estado de checkboxes
    if "caps_state" not in st.session_state or st.session_state.get("last_cliente") != cliente:
        st.session_state.caps_state = {c["id"]: True for c in capitulos_config}
        st.session_state.last_cliente = cliente

    if sel_todos:
        for c in capitulos_config:
            st.session_state.caps_state[c["id"]] = True
    if sel_ninguno:
        for c in capitulos_config:
            st.session_state.caps_state[c["id"]] = False

    # Renderizar árbol — solo checkboxes para capítulos raíz (sin parent)
    # Los hijos se agregan automáticamente en el backend si el padre está seleccionado
    top_level = [c for c in capitulos_config if c["parent"] is None]

    for cap in top_level:
        checked = st.session_state.caps_state.get(cap["id"], True)
        nuevo = st.checkbox(
            f"**{cap['num']}. {cap['titulo']}**",
            value=checked,
            key=f"cap_root_{cap['id']}"
        )
        st.session_state.caps_state[cap["id"]] = nuevo

        # Mostrar hijos (nivel 2) sangrados
        hijos = [c for c in capitulos_config if c["parent"] == cap["id"]]
        for hijo in hijos:
            checked_h = st.session_state.caps_state.get(hijo["id"], True)
            nuevo_h = st.checkbox(
                f"  {hijo['num']}. {hijo['titulo']}",
                value=checked_h,
                key=f"cap_h2_{hijo['id']}"
            )
            st.session_state.caps_state[hijo["id"]] = nuevo_h

            # Nietos (nivel 3)
            nietos = [c for c in capitulos_config if c["parent"] == hijo["id"]]
            for nieto in nietos:
                checked_n = st.session_state.caps_state.get(nieto["id"], True)
                nuevo_n = st.checkbox(
                    f"    {nieto['num']}. {nieto['titulo']}",
                    value=checked_n,
                    key=f"cap_h3_{nieto['id']}"
                )
                st.session_state.caps_state[nieto["id"]] = nuevo_n

    capitulos_seleccionados = [cid for cid, val in st.session_state.caps_state.items() if val]

# ════════════════════════════════════════════════════════════════════════
# CUERPO PRINCIPAL — Carga de archivos
# ════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-card"><h3>📂 Carga de Archivos</h3>', unsafe_allow_html=True)

col_word, col_right = st.columns([1, 1])

with col_word:
    st.markdown("##### 📝 Informe Word del mes anterior (plantilla base)")
    word_file = st.file_uploader(
        "Arrastra o selecciona el .docx del mes anterior",
        type=["docx"],
        key="word_upload",
        help="Se usa como plantilla de diseño. El mes/año se actualizará automáticamente."
    )
    if word_file:
        st.success(f"✅ {word_file.name}  ({round(word_file.size/1024)} KB)")

with col_right:
    st.markdown("##### 📊 Bases de datos Excel")

    ec1, ec2 = st.columns(2)
    with ec1:
        opex_file = st.file_uploader(
            "🔧 db_opex — Operaciones",
            type=["xlsx"], key="opex_upload",
            help="Hojas: Agendamiento y Reposiciones"
        )
        if opex_file:
            st.success(f"✅ {opex_file.name}")

        fact_file = st.file_uploader(
            "🧾 db_fact — Facturación",
            type=["xlsx"], key="fact_upload",
            help="Hoja principal con facturación por NIU"
        )
        if fact_file:
            st.success(f"✅ {fact_file.name}")

    with ec2:
        sac_file = st.file_uploader(
            "🎧 db_sac — Servicio al Cliente (PQR)",
            type=["xlsx"], key="sac_upload",
            help="Tickets, estados, hurtos por NIU"
        )
        if sac_file:
            st.success(f"✅ {sac_file.name}")

        asist_file = st.file_uploader(
            "👥 db_asistencia — Atención Sedes",
            type=["xlsx"], key="asist_upload",
            help="Visitas presenciales y contactos digitales"
        )
        if asist_file:
            st.success(f"✅ {asist_file.name}")

st.markdown('</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════
# BOTÓN GENERAR
# ════════════════════════════════════════════════════════════════════════
st.markdown("")
generar = st.button("⚡ Generar Informe Word", type="primary", use_container_width=False)

# ════════════════════════════════════════════════════════════════════════
# PROCESO DE GENERACIÓN
# ════════════════════════════════════════════════════════════════════════
if generar:
    if not capitulos_seleccionados:
        st.error("⚠️ Selecciona al menos un capítulo en el menú lateral.")
        st.stop()

    sess_dir = Path(tempfile.mkdtemp())
    charts_dir = sess_dir / "charts"
    charts_dir.mkdir()

    try:
        # ── Barra de progreso ────────────────────────────────────────────
        progress = st.progress(0)
        status   = st.empty()

        # Paso 1 — Guardar archivos subidos
        status.info("📁 Guardando archivos...")
        progress.progress(10)

        saved = {}

        if word_file:
            p = sess_dir / "template.docx"
            p.write_bytes(word_file.read())
            saved["template_word"] = str(p)

        for key, uploaded in [("db_opex", opex_file), ("db_sac", sac_file),
                               ("db_fact", fact_file), ("db_asistencia", asist_file)]:
            if uploaded:
                p = sess_dir / uploaded.name
                p.write_bytes(uploaded.read())
                saved[key] = str(p)

        # Paso 2 — Procesar Excel
        status.info("📊 Leyendo y procesando bases de datos Excel...")
        progress.progress(30)

        processor = DataProcessor(config, mes, ano)
        db_files = {k: v for k, v in saved.items() if k.startswith("db_")}
        if db_files:
            processor.cargar_excel(db_files)

        metricas = processor.get_all_metricas()
        inconsistencias = metricas.pop("inconsistencias", [])

        # Paso 3 — Validaciones
        status.info(f"🔍 Validando consistencia de datos... ({len(inconsistencias)} inconsistencias detectadas)")
        progress.progress(50)

        # Paso 4 — Gráficos
        status.info("📈 Generando gráficos automáticos...")
        progress.progress(65)
        graficos = generar_todos_los_graficos(metricas, str(charts_dir))

        # Paso 5 — Documento Word
        status.info("📝 Construyendo documento Word...")
        progress.progress(80)

        mes_nombre  = MESES.get(mes, str(mes))
        output_name = f"Informe_{cliente}_{mes_nombre}_{ano}.docx"
        output_path = str(OUTPUT_DIR / output_name)

        gen_word = WordGenerator(
            config=config,
            mes=mes,
            ano=ano,
            template_path=saved.get("template_word", ""),
            metricas=metricas,
            graficos=graficos,
            capitulos_seleccionados=capitulos_seleccionados
        )
        gen_word.agregar_seccion_inconsistencias(inconsistencias)
        gen_word.generar(output_path)

        progress.progress(100)
        status.empty()

        # ── Resultado ────────────────────────────────────────────────────
        ops   = metricas.get("operaciones", {})
        sac   = metricas.get("sac", {})
        asist = metricas.get("asistencia", {})
        fact  = metricas.get("facturacion", {})

        st.success(f"✅ ¡Informe **{output_name}** generado exitosamente!")

        # Métricas resumen
        st.markdown("""
        <div class="section-card"><h3>📊 Resumen de métricas procesadas</h3>
        """, unsafe_allow_html=True)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🔧 Visitas operaciones",  fmt_num(ops.get("total", 0)))
        m2.metric("🎧 PQR registradas",       fmt_num(sac.get("total", 0)))
        m3.metric("👥 Atenciones en sede",    fmt_num(asist.get("total", 0)))
        m4.metric("🧾 Facturación total",
                  f"$ {fact.get('total_facturado',0)/1e6:.1f} M")

        # Sub-métricas operaciones
        if ops.get("total", 0) > 0:
            st.markdown("**Operaciones:**")
            o1, o2, o3, o4 = st.columns(4)
            o1.metric("Preventivos",  fmt_num(ops.get("preventivos", 0)))
            o2.metric("Correctivos",  fmt_num(ops.get("correctivos", 0)))
            o3.metric("Funcionales",  fmt_num(ops.get("funcionales", 0)))
            o4.metric("No funcionales", fmt_num(ops.get("no_funcionales", 0)))

        # Sub-métricas SAC
        if sac.get("total", 0) > 0:
            st.markdown("**SAC / PQR:**")
            s1, s2, s3 = st.columns(3)
            s1.metric("Cerrados",  fmt_num(sac.get("cerrados", 0)))
            s2.metric("Abiertos",  fmt_num(sac.get("abiertos", 0)))
            s3.metric("Hurtos",    fmt_num(sac.get("hurtos", 0)))

        st.markdown("</div>", unsafe_allow_html=True)

        # Inconsistencias
        if inconsistencias:
            st.markdown(f"""
            <div class="badge-warn">
              ⚠️ {len(inconsistencias)} inconsistencias detectadas — incluidas en la última sección del informe
            </div>""", unsafe_allow_html=True)

            with st.expander(f"Ver detalle de inconsistencias ({len(inconsistencias)})"):
                import pandas as pd
                df_inc = pd.DataFrame(inconsistencias)
                df_inc.columns = ["Base origen", "Hoja/Archivo", "NIU",
                                  "Tipo", "Descripción", "Recomendación"]
                st.dataframe(df_inc, use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="badge-ok">✅ Sin inconsistencias detectadas</div>',
                        unsafe_allow_html=True)

        # Gráficos preview
        if graficos:
            st.markdown("---")
            st.markdown("**Vista previa de gráficos generados:**")
            gcols = st.columns(min(len(graficos), 2))
            for i, (nombre, ruta) in enumerate(graficos.items()):
                if ruta and Path(ruta).exists():
                    gcols[i % 2].image(ruta, use_container_width=True)

        # Botón de descarga
        st.markdown("---")
        with open(output_path, "rb") as f:
            docx_bytes = f.read()

        st.download_button(
            label=f"⬇️ Descargar  {output_name}",
            data=docx_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    except Exception as e:
        import traceback
        st.error(f"❌ Error durante la generación: {e}")
        with st.expander("Ver detalle del error (traceback)"):
            st.code(traceback.format_exc())

    finally:
        shutil.rmtree(str(sess_dir), ignore_errors=True)

# ── Footer ───────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#aaa; font-size:.8rem;'>"
    "DISPOWER S.A.S E.S.P. — Sistema Generador de Informes © 2026"
    "</p>",
    unsafe_allow_html=True
)
