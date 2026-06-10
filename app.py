"""
app.py - Backend Flask para generación automática de informes Word
"""
import os
import json
import uuid
import shutil
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

# ---- Importar utilidades propias
import sys
sys.path.insert(0, str(Path(__file__).parent))
from utils.data_processor import DataProcessor, MESES
from utils.chart_generator import generar_todos_los_graficos
from utils.word_generator import WordGenerator

# ── Configuración ──────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "output"
CONFIG_FOLDER = BASE_DIR / "config"
TEMP_FOLDER   = BASE_DIR / "temp"

for d in [UPLOAD_FOLDER, OUTPUT_FOLDER, CONFIG_FOLDER, TEMP_FOLDER]:
    d.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "docx"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_config(cliente: str) -> dict:
    path = CONFIG_FOLDER / f"{cliente}.json"
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def list_clients() -> list:
    return [p.stem for p in CONFIG_FOLDER.glob("*.json")]


# ── Rutas ──────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html", clientes=list_clients(), meses=MESES)


@app.route("/api/config/<cliente>")
def api_config(cliente):
    cfg = load_config(cliente)
    if not cfg:
        return jsonify({"error": "Cliente no encontrado"}), 404
    return jsonify(cfg)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """
    Recibe:
      - cliente (str)
      - mes (int)
      - ano (int)
      - capitulos (JSON array de IDs)
      - template_word (file, opcional)
      - db_opex (file, opcional)
      - db_sac (file, opcional)
      - db_fact (file, opcional)
      - db_asistencia (file, opcional)
    """
    try:
        cliente   = request.form.get("cliente", "")
        mes       = int(request.form.get("mes", 1))
        ano       = int(request.form.get("ano", 2026))
        caps_raw  = request.form.get("capitulos", "[]")
        capitulos = json.loads(caps_raw)

        config = load_config(cliente)
        if not config:
            return jsonify({"error": f"Config no encontrada para cliente '{cliente}'"}), 400

        # Crear directorio temporal para esta sesión
        session_id = uuid.uuid4().hex[:8]
        sess_dir   = TEMP_FOLDER / session_id
        sess_dir.mkdir()
        charts_dir = sess_dir / "charts"
        charts_dir.mkdir()

        # Guardar archivos subidos
        saved = {}
        for key in ["template_word", "db_opex", "db_sac", "db_fact", "db_asistencia"]:
            f = request.files.get(key)
            if f and f.filename and allowed(f.filename):
                fname = secure_filename(f.filename)
                dest  = sess_dir / fname
                f.save(str(dest))
                saved[key] = str(dest)

        # Procesar datos
        processor = DataProcessor(config, mes, ano)
        db_files  = {k: v for k, v in saved.items()
                     if k.startswith("db_") and v}
        if db_files:
            processor.cargar_excel(db_files)

        metricas      = processor.get_all_metricas()
        inconsistencias = metricas.pop("inconsistencias", [])

        # Generar gráficos
        graficos = generar_todos_los_graficos(metricas, str(charts_dir))

        # Generar Word
        template_path = saved.get("template_word", "")
        mes_nombre    = MESES.get(mes, str(mes))
        output_name   = f"Informe_{cliente}_{mes_nombre}_{ano}.docx"
        output_path   = str(OUTPUT_FOLDER / output_name)

        gen = WordGenerator(
            config=config,
            mes=mes,
            ano=ano,
            template_path=template_path,
            metricas=metricas,
            graficos=graficos,
            capitulos_seleccionados=capitulos
        )
        gen.agregar_seccion_inconsistencias(inconsistencias)
        gen.generar(output_path)

        # Limpiar temporales
        shutil.rmtree(str(sess_dir), ignore_errors=True)

        return jsonify({
            "success": True,
            "filename": output_name,
            "inconsistencias": len(inconsistencias),
            "metricas_resumen": {
                "operaciones_total": metricas.get("operaciones", {}).get("total", 0),
                "sac_total": metricas.get("sac", {}).get("total", 0),
                "asistencia_total": metricas.get("asistencia", {}).get("total", 0),
                "facturacion_total": metricas.get("facturacion", {}).get("total_facturado", 0),
            }
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/download/<filename>")
def api_download(filename):
    path = OUTPUT_FOLDER / secure_filename(filename)
    if not path.exists():
        return jsonify({"error": "Archivo no encontrado"}), 404
    return send_file(str(path), as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/api/preview/<filename>")
def api_preview(filename):
    """Retorna un resumen de métricas del archivo generado (placeholder)."""
    path = OUTPUT_FOLDER / secure_filename(filename)
    return jsonify({"exists": path.exists(), "size_kb": round(path.stat().st_size/1024, 1) if path.exists() else 0})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
