"""
streamlit_app.py
Generador de Informes Mensuales – DISPOWER S.A.S E.S.P.
"""
import streamlit as st
import json, shutil, tempfile, os, io
from pathlib import Path

BASE_DIR   = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "config"

import sys
sys.path.insert(0, str(BASE_DIR))
from utils.data_processor import DataProcessor, MESES
from utils.chart_generator import generar_todos_los_graficos
from utils.word_generator   import WordGenerator

# ── helpers ──────────────────────────────────────────────────────────────
def list_clients():
    return sorted([p.stem for p in CONFIG_DIR.glob("*.json")])

def load_config(cliente: str) -> dict:
    path = CONFIG_DIR / f"{cliente}.json"
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def fmt_num(n):
    try:    return f"{int(n):,}".replace(",", ".")
    except: return str(n)

def fmt_cop(n):
    try:    return f"$ {int(n):,}".replace(",", ".")
    except: return str(n)

# ── page config ───────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Generador de Informes | DISPOWER",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.header-banner {
    background: linear-gradient(135deg,#1B4F8A,#2E86C1);
    color:#fff; padding:22px 30px; border-radius:12px;
    margin-bottom:24px; display:flex;
    align-items:center; justify-content:space-between;
}
.header-banner h1 { margin:0; font-size:1.5rem; font-weight:700; }
.header-banner p  { margin:4px 0 0; font-size:.85rem; opacity:.85; }
.badge-warn { background:#FDEBD0; color:#D35400; border-radius:20px;
    padding:5px 14px; font-size:.85rem; font-weight:600;
    display:inline-block; margin-top:10px; }
.badge-ok   { background:#D5F5E3; color:#1E8449; border-radius:20px;
    padding:5px 14px; font-size:.85rem; font-weight:600;
    display:inline-block; margin-top:10px; }
footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-banner">
  <div>
    <h1>📄 Generador de Informes Mensuales</h1>
    <p>DISPOWER S.A.S E.S.P. — Sistema automatizado de informes de concesión</p>
  </div>
  <div style="text-align:right;opacity:.8;font-size:.8rem;">
    Operaciones · SAC · Facturación · HSEQ
  </div>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════
clientes = list_clients()

with st.sidebar:
    st.markdown("### ⚙️ Parámetros del informe")

    # ── Cliente ──────────────────────────────────────────────────────────
    idx_default = clientes.index("DISPAC") if "DISPAC" in clientes else 0
    cliente = st.selectbox("Cliente", clientes, index=idx_default,
                           key="sel_cliente")

    # Cuando cambia el cliente → resetear capítulos
    if st.session_state.get("_ultimo_cliente") != cliente:
        st.session_state["_ultimo_cliente"] = cliente
        # Borrar todas las keys de capítulos del cliente anterior
        for k in list(st.session_state.keys()):
            if k.startswith("chk_"):
                del st.session_state[k]

    config           = load_config(cliente)
    capitulos_config = config.get("capitulos", [])

    # ── Mes / Año ─────────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    mes = c1.selectbox("Mes", list(MESES.keys()),
                       format_func=lambda x: MESES[x],
                       index=3, key="sel_mes")
    ano = c2.number_input("Año", min_value=2020, max_value=2040,
                          value=2026, step=1, key="sel_ano")

    st.markdown("---")
    st.markdown("### 📋 Capítulos a incluir")

    # ── Botones Todos / Ninguno ───────────────────────────────────────────
    b1, b2 = st.columns(2)
    if b1.button("✅ Todos",   use_container_width=True, key="btn_todos"):
        for cap in capitulos_config:
            st.session_state[f"chk_{cap['id']}"] = True
    if b2.button("⬜ Ninguno", use_container_width=True, key="btn_ninguno"):
        for cap in capitulos_config:
            st.session_state[f"chk_{cap['id']}"] = False

    # ── Árbol de capítulos ────────────────────────────────────────────────
    # Renderiza niveles 1 → 2 → 3 con sangría visual
    caps_por_padre = {}
    for cap in capitulos_config:
        padre = cap.get("parent")
        caps_por_padre.setdefault(padre, []).append(cap)

    def render_caps(padre_id, indent=0):
        hijos = caps_por_padre.get(padre_id, [])
        for cap in hijos:
            cid   = cap["id"]
            label = ("　" * indent) + f"{cap['num']}. {cap['titulo']}"
            # Valor por defecto: True (todos seleccionados al iniciar)
            key   = f"chk_{cid}"
            if key not in st.session_state:
                st.session_state[key] = True
            st.checkbox(label, key=key)
            render_caps(cid, indent + 1)

    render_caps(None)

    # Recoger IDs seleccionados
    capitulos_seleccionados = [
        cap["id"] for cap in capitulos_config
        if st.session_state.get(f"chk_{cap['id']}", True)
    ]

# ════════════════════════════════════════════════════════════════════════
# CUERPO — Carga de archivos
# ════════════════════════════════════════════════════════════════════════
st.markdown("#### 📂 Carga de archivos")

col_word, col_excel = st.columns([1, 1])

with col_word:
    st.markdown("**📝 Informe Word del mes anterior** *(plantilla base)*")
    word_file = st.file_uploader(
        "Arrastra o selecciona el .docx del mes anterior",
        type=["docx"], key="word_upload",
        help="Se usa como plantilla de diseño. El mes/año se actualizará automáticamente."
    )
    if word_file:
        st.success(f"✅ {word_file.name}  ({round(word_file.size/1024)} KB)")

with col_excel:
    st.markdown("**📊 Bases de datos Excel**")
    ec1, ec2 = st.columns(2)

    with ec1:
        opex_file = st.file_uploader(
            "🔧 db_opex — Operaciones", type=["xlsx"], key="opex_upload",
            help="Hojas: Agendamiento y Reposiciones"
        )
        if opex_file: st.success(f"✅ {opex_file.name}")

        fact_file = st.file_uploader(
            "🧾 db_fact — Facturación", type=["xlsx"], key="fact_upload",
            help="Hoja principal con facturación por NIU"
        )
        if fact_file: st.success(f"✅ {fact_file.name}")

    with ec2:
        sac_file = st.file_uploader(
            "🎧 db_sac — PQR / SAC", type=["xlsx"], key="sac_upload",
            help="Reporte asesores: tickets, semáforo, hurtos"
        )
        if sac_file: st.success(f"✅ {sac_file.name}")

        asist_file = st.file_uploader(
            "👥 db_asistencia — Sedes", type=["xlsx"], key="asist_upload",
            help="Visitas presenciales y contactos digitales"
        )
        if asist_file: st.success(f"✅ {asist_file.name}")

st.markdown("---")

# ════════════════════════════════════════════════════════════════════════
# BOTÓN GENERAR
# ════════════════════════════════════════════════════════════════════════
generar = st.button("⚡ Generar Informe Word", type="primary",
                    use_container_width=True)

# ════════════════════════════════════════════════════════════════════════
# PROCESO
# ════════════════════════════════════════════════════════════════════════
if generar:
    if not capitulos_seleccionados:
        st.error("⚠️ Selecciona al menos un capítulo en el menú lateral.")
        st.stop()

    sess_dir   = Path(tempfile.mkdtemp())
    charts_dir = sess_dir / "charts"
    charts_dir.mkdir()
    output_dir = sess_dir / "output"
    output_dir.mkdir()

    try:
        progress = st.progress(0)
        status   = st.empty()

        # ── 1. Guardar archivos subidos ───────────────────────────────────
        status.info("📁 Leyendo archivos cargados…")
        progress.progress(10)

        saved = {}
        if word_file:
            word_file.seek(0)
            p = sess_dir / "template.docx"
            p.write_bytes(word_file.read())
            saved["template_word"] = str(p)

        for key, uploaded in [("db_opex",       opex_file),
                               ("db_sac",        sac_file),
                               ("db_fact",        fact_file),
                               ("db_asistencia",  asist_file)]:
            if uploaded:
                uploaded.seek(0)
                p = sess_dir / uploaded.name
                p.write_bytes(uploaded.read())
                saved[key] = str(p)

        # ── 2. Procesar Excel ─────────────────────────────────────────────
        status.info("📊 Procesando bases de datos Excel…")
        progress.progress(30)

        processor = DataProcessor(config, mes, ano)
        db_files  = {k: v for k, v in saved.items() if k.startswith("db_")}
        if db_files:
            processor.cargar_excel(db_files)

        metricas        = processor.get_all_metricas()
        inconsistencias = metricas.pop("inconsistencias", [])

        # ── 3. Gráficos ───────────────────────────────────────────────────
        status.info(f"📈 Generando gráficos… ({len(inconsistencias)} inconsistencias detectadas)")
        progress.progress(55)
        graficos = generar_todos_los_graficos(metricas, str(charts_dir))

        # ── 4. Word ───────────────────────────────────────────────────────
        status.info("📝 Construyendo documento Word…")
        progress.progress(75)

        mes_nombre  = MESES.get(mes, str(mes))
        output_name = f"Informe_{cliente}_{mes_nombre}_{ano}.docx"
        output_path = str(output_dir / output_name)

        gen_word = WordGenerator(
            config=config, mes=mes, ano=ano,
            template_path=saved.get("template_word", ""),
            metricas=metricas, graficos=graficos,
            capitulos_seleccionados=capitulos_seleccionados
        )
        gen_word.agregar_seccion_inconsistencias(inconsistencias)
        gen_word.generar(output_path)

        # Leer bytes ANTES de borrar el directorio temporal
        with open(output_path, "rb") as fh:
            docx_bytes = fh.read()

        progress.progress(100)
        status.empty()

        # ── Resultado ─────────────────────────────────────────────────────
        ops  = metricas.get("operaciones", {})
        sac  = metricas.get("sac",         {})
        asist= metricas.get("asistencia",  {})
        fact = metricas.get("facturacion", {})

        st.success(f"✅ Informe **{output_name}** generado — "
                   f"{len(capitulos_seleccionados)} capítulos incluidos")

        # Métricas resumen
        st.markdown("#### 📊 Resumen de métricas procesadas")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🔧 Visitas operaciones", fmt_num(ops.get("total", 0)))
        m2.metric("🎧 PQR registradas",     fmt_num(sac.get("total", 0)))
        m3.metric("👥 Atenciones en sede",  fmt_num(asist.get("total", 0)))
        m4.metric("🧾 Facturación total",
                  f"$ {fact.get('total_facturado',0)/1e6:.1f} M"
                  if fact.get("total_facturado",0) else "—")

        if ops.get("total", 0) > 0:
            st.markdown("**Operaciones:**")
            o1, o2, o3, o4 = st.columns(4)
            o1.metric("Preventivos",    fmt_num(ops.get("preventivos", 0)))
            o2.metric("Correctivos",    fmt_num(ops.get("correctivos", 0)))
            o3.metric("Funcionales",    fmt_num(ops.get("funcionales", 0)))
            o4.metric("No funcionales", fmt_num(ops.get("no_funcionales", 0)))

        if sac.get("total", 0) > 0:
            st.markdown("**SAC / PQR:**")
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("Cerrados",  fmt_num(sac.get("cerrados",  0)))
            s2.metric("Abiertos",  fmt_num(sac.get("abiertos",  0)))
            s3.metric("Hurtos",    fmt_num(sac.get("hurtos",    0)))
            s4.metric("Con bloqueo facturación",
                      fmt_num(sac.get("bloqueos_facturacion", 0)))

        # Inconsistencias
        st.markdown("---")
        if inconsistencias:
            st.markdown(
                f'<div class="badge-warn">⚠️ {len(inconsistencias)} '
                f'inconsistencias detectadas — incluidas en la última sección del informe</div>',
                unsafe_allow_html=True)
            with st.expander(f"Ver detalle ({len(inconsistencias)} registros)"):
                import pandas as pd
                df_inc = pd.DataFrame(inconsistencias)
                df_inc.columns = ["Base origen","Hoja/Archivo","NUI",
                                  "Tipo","Descripción","Recomendación"]
                st.dataframe(df_inc, use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="badge-ok">✅ Sin inconsistencias detectadas</div>',
                        unsafe_allow_html=True)

        # Preview gráficos
        graficos_validos = {k: v for k, v in graficos.items()
                            if v and Path(v).exists() and k != "sac_estado"}
        if graficos_validos:
            st.markdown("---")
            st.markdown("**Vista previa de gráficos generados:**")
            cols = st.columns(2)
            for i, (nombre, ruta) in enumerate(graficos_validos.items()):
                cols[i % 2].image(ruta, caption=nombre.replace("_", " ").title(),
                                  use_container_width=True)

        # ── Descarga ──────────────────────────────────────────────────────
        st.markdown("---")
        st.download_button(
            label=f"⬇️  Descargar  {output_name}",
            data=docx_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    except Exception as e:
        import traceback
        st.error(f"❌ Error durante la generación: {e}")
        with st.expander("Ver traceback completo"):
            st.code(traceback.format_exc())

    finally:
        shutil.rmtree(str(sess_dir), ignore_errors=True)

# ── footer ────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#aaa;font-size:.8rem;'>"
    "DISPOWER S.A.S E.S.P. — Sistema Generador de Informes © 2026</p>",
    unsafe_allow_html=True
)
