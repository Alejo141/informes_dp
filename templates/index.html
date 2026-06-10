<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Generador de Informes | DISPOWER</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet"/>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet"/>
  <style>
    :root {
      --dp-azul: #1B4F8A;
      --dp-azul-claro: #2E86C1;
      --dp-naranja: #F39C12;
      --dp-gris: #F4F6F9;
      --dp-borde: #DEE2E6;
    }
    body { background: var(--dp-gris); font-family: 'Segoe UI', sans-serif; }

    .navbar-brand img { height: 38px; }
    .navbar { background: var(--dp-azul) !important; }

    .card { border: none; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,.07); }
    .card-header {
      background: var(--dp-azul);
      color: #fff;
      border-radius: 12px 12px 0 0 !important;
      font-weight: 600;
      letter-spacing: .3px;
    }

    /* Upload zones */
    .drop-zone {
      border: 2px dashed var(--dp-azul-claro);
      border-radius: 10px;
      padding: 20px;
      text-align: center;
      cursor: pointer;
      transition: background .2s, border-color .2s;
      background: #fff;
      position: relative;
    }
    .drop-zone:hover, .drop-zone.dragover { background: #EBF5FB; border-color: var(--dp-azul); }
    .drop-zone input[type=file] { position:absolute; inset:0; opacity:0; cursor:pointer; }
    .drop-zone .badge-file {
      display:inline-flex; align-items:center; gap:6px;
      background: #D6EAF8; color: var(--dp-azul);
      border-radius:20px; padding:3px 12px; font-size:.8rem; margin-top:8px;
    }

    /* Capítulos tree */
    .cap-tree { max-height: 460px; overflow-y: auto; border-radius: 8px; padding: 8px; background:#fff; }
    .cap-item { padding: 4px 0; }
    .cap-item.level-1 { font-weight: 600; border-bottom: 1px solid #f0f0f0; margin-bottom:2px; }
    .cap-item.level-2 { padding-left: 20px; }
    .cap-item.level-3 { padding-left: 40px; font-size:.9rem; }
    .cap-item label { cursor:pointer; user-select:none; display:flex; align-items:center; gap:8px; }
    .cap-item input[type=checkbox]:checked + span { color: var(--dp-azul); font-weight:600; }

    /* Progress */
    #progress-wrap { display:none; }
    .step-dot {
      width:32px; height:32px; border-radius:50%;
      display:flex; align-items:center; justify-content:center;
      font-size:.8rem; font-weight:700; border:2px solid #ddd;
      color:#aaa; background:#fff;
    }
    .step-dot.active { border-color:var(--dp-azul); color:var(--dp-azul); }
    .step-dot.done   { background:var(--dp-azul); border-color:var(--dp-azul); color:#fff; }

    /* Result */
    #result-card { display:none; }
    .metric-box {
      background: linear-gradient(135deg, var(--dp-azul), var(--dp-azul-claro));
      color:#fff; border-radius:10px; padding:16px 20px; text-align:center;
    }
    .metric-box .num { font-size:2rem; font-weight:700; }
    .metric-box .lbl { font-size:.8rem; opacity:.85; }

    .inconsistencia-badge {
      background: #FDEBD0; color: #D35400; border-radius:20px;
      padding:4px 14px; font-size:.85rem; font-weight:600;
    }

    /* Select2-like select */
    select.form-select { border-radius:8px; }

    .btn-generar {
      background: linear-gradient(135deg, var(--dp-azul), var(--dp-azul-claro));
      color:#fff; border:none; border-radius:10px;
      padding:12px 32px; font-size:1rem; font-weight:600;
      transition: transform .15s, box-shadow .15s;
    }
    .btn-generar:hover { transform:translateY(-2px); box-shadow: 0 6px 18px rgba(27,79,138,.35); color:#fff; }
    .btn-generar:active { transform:translateY(0); }

    .btn-download {
      background: #27AE60; color:#fff; border:none; border-radius:10px;
      padding:10px 28px; font-weight:600; transition: background .2s;
    }
    .btn-download:hover { background:#1E8449; color:#fff; }

    footer { font-size:.8rem; color:#aaa; padding:20px 0; text-align:center; }
  </style>
</head>
<body>

<!-- Navbar -->
<nav class="navbar navbar-dark px-4 py-2 mb-4">
  <span class="navbar-brand fw-bold fs-5">
    <i class="bi bi-file-earmark-word me-2"></i>Generador de Informes Mensual
  </span>
  <span class="text-white-50 small">DISPOWER S.A.S E.S.P.</span>
</nav>

<div class="container-xl pb-5">

  <form id="report-form" enctype="multipart/form-data">

    <div class="row g-4">

      <!-- ── Col izquierda: parámetros + archivos ──────────────────── -->
      <div class="col-lg-7">

        <!-- Parámetros del informe -->
        <div class="card mb-4">
          <div class="card-header"><i class="bi bi-sliders me-2"></i>Parámetros del Informe</div>
          <div class="card-body row g-3 pt-3">

            <div class="col-md-4">
              <label class="form-label fw-semibold">Cliente</label>
              <select id="select-cliente" name="cliente" class="form-select">
                {% for c in clientes %}
                <option value="{{ c }}" {% if c == 'DISPAC' %}selected{% endif %}>{{ c }}</option>
                {% endfor %}
              </select>
            </div>

            <div class="col-md-4">
              <label class="form-label fw-semibold">Mes del informe</label>
              <select name="mes" class="form-select">
                {% for num, nombre in meses.items() %}
                <option value="{{ num }}" {% if num == 4 %}selected{% endif %}>{{ nombre }}</option>
                {% endfor %}
              </select>
            </div>

            <div class="col-md-4">
              <label class="form-label fw-semibold">Año del informe</label>
              <input type="number" name="ano" class="form-control" value="2026" min="2020" max="2040"/>
            </div>

          </div>
        </div>

        <!-- Carga de archivos -->
        <div class="card mb-4">
          <div class="card-header"><i class="bi bi-cloud-upload me-2"></i>Carga de Archivos</div>
          <div class="card-body row g-3 pt-3">

            <!-- Word base -->
            <div class="col-12">
              <label class="form-label fw-semibold">
                <i class="bi bi-file-word text-primary me-1"></i>
                Informe Word del mes anterior (plantilla base)
              </label>
              <div class="drop-zone" id="dz-word">
                <input type="file" name="template_word" accept=".docx" onchange="showFile(this,'dz-word')"/>
                <i class="bi bi-file-earmark-word fs-2 text-primary"></i>
                <div class="mt-1 small text-muted">Arrastra aquí o haz clic para seleccionar <strong>.docx</strong></div>
                <div id="dz-word-label"></div>
              </div>
            </div>

            <!-- Excel files grid -->
            <div class="col-sm-6">
              <label class="form-label fw-semibold">
                <i class="bi bi-tools me-1 text-success"></i>db_opex — Operaciones / Mantenimiento
              </label>
              <div class="drop-zone" id="dz-opex">
                <input type="file" name="db_opex" accept=".xlsx" onchange="showFile(this,'dz-opex')"/>
                <i class="bi bi-file-earmark-excel fs-2 text-success"></i>
                <div class="small text-muted">Agendamiento · Reposiciones</div>
                <div id="dz-opex-label"></div>
              </div>
            </div>

            <div class="col-sm-6">
              <label class="form-label fw-semibold">
                <i class="bi bi-headset me-1 text-warning"></i>db_sac — Servicio al Cliente (PQR)
              </label>
              <div class="drop-zone" id="dz-sac">
                <input type="file" name="db_sac" accept=".xlsx" onchange="showFile(this,'dz-sac')"/>
                <i class="bi bi-file-earmark-excel fs-2 text-warning"></i>
                <div class="small text-muted">Tickets · Estados · Hurtos</div>
                <div id="dz-sac-label"></div>
              </div>
            </div>

            <div class="col-sm-6">
              <label class="form-label fw-semibold">
                <i class="bi bi-receipt me-1 text-danger"></i>db_fact — Facturación
              </label>
              <div class="drop-zone" id="dz-fact">
                <input type="file" name="db_fact" accept=".xlsx" onchange="showFile(this,'dz-fact')"/>
                <i class="bi bi-file-earmark-excel fs-2 text-danger"></i>
                <div class="small text-muted">Facturas · Recaudo · Usuarios</div>
                <div id="dz-fact-label"></div>
              </div>
            </div>

            <div class="col-sm-6">
              <label class="form-label fw-semibold">
                <i class="bi bi-people me-1 text-info"></i>db_asistencia — Atención Sedes
              </label>
              <div class="drop-zone" id="dz-asistencia">
                <input type="file" name="db_asistencia" accept=".xlsx" onchange="showFile(this,'dz-asistencia')"/>
                <i class="bi bi-file-earmark-excel fs-2 text-info"></i>
                <div class="small text-muted">Visitas presenciales · Digitales</div>
                <div id="dz-asistencia-label"></div>
              </div>
            </div>

          </div>
        </div>

      </div><!-- /col-izquierda -->

      <!-- ── Col derecha: capítulos ────────────────────────────────── -->
      <div class="col-lg-5">
        <div class="card h-100">
          <div class="card-header d-flex justify-content-between align-items-center">
            <span><i class="bi bi-list-check me-2"></i>Capítulos a incluir</span>
            <div>
              <button type="button" class="btn btn-sm btn-outline-light me-1" onclick="toggleAll(true)">Todos</button>
              <button type="button" class="btn btn-sm btn-outline-light" onclick="toggleAll(false)">Ninguno</button>
            </div>
          </div>
          <div class="card-body p-2">
            <div class="cap-tree" id="cap-tree">
              <p class="text-muted text-center py-4">Selecciona un cliente para cargar los capítulos</p>
            </div>
          </div>
        </div>
      </div>

    </div><!-- /row -->

    <!-- Botón generar -->
    <div class="text-center mt-4">
      <button type="submit" class="btn-generar btn">
        <i class="bi bi-magic me-2"></i>Generar Informe Word
      </button>
    </div>

  </form>

  <!-- Progress -->
  <div id="progress-wrap" class="card mt-4 p-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
      <h6 class="mb-0 fw-bold">Procesando informe…</h6>
      <span id="progress-pct" class="fw-bold text-primary">0%</span>
    </div>
    <div class="progress mb-3" style="height:10px; border-radius:8px;">
      <div id="progress-bar" class="progress-bar progress-bar-striped progress-bar-animated"
           style="width:0%; background:var(--dp-azul);"></div>
    </div>
    <div class="d-flex justify-content-around">
      <div class="text-center">
        <div class="step-dot mx-auto" id="step-1"><i class="bi bi-file-earmark-spreadsheet"></i></div>
        <div class="small mt-1">Leyendo Excel</div>
      </div>
      <div class="text-center">
        <div class="step-dot mx-auto" id="step-2"><i class="bi bi-shield-check"></i></div>
        <div class="small mt-1">Validaciones</div>
      </div>
      <div class="text-center">
        <div class="step-dot mx-auto" id="step-3"><i class="bi bi-bar-chart"></i></div>
        <div class="small mt-1">Gráficos</div>
      </div>
      <div class="text-center">
        <div class="step-dot mx-auto" id="step-4"><i class="bi bi-file-word"></i></div>
        <div class="small mt-1">Documento</div>
      </div>
    </div>
  </div>

  <!-- Resultado -->
  <div id="result-card" class="card mt-4">
    <div class="card-header bg-success text-white">
      <i class="bi bi-check-circle me-2"></i>¡Informe generado exitosamente!
    </div>
    <div class="card-body">
      <div class="row g-3 mb-4" id="metric-boxes"></div>

      <div class="d-flex align-items-center gap-3 flex-wrap">
        <button id="btn-download" class="btn-download btn" onclick="downloadFile()">
          <i class="bi bi-download me-2"></i>Descargar Informe Word
        </button>
        <span id="incons-badge" class="inconsistencia-badge" style="display:none">
          <i class="bi bi-exclamation-triangle me-1"></i>
          <span id="incons-count"></span> inconsistencias detectadas (ver última sección del informe)
        </span>
      </div>
    </div>
  </div>

  <!-- Error -->
  <div id="error-card" class="alert alert-danger mt-4" style="display:none">
    <i class="bi bi-x-circle me-2"></i><strong>Error:</strong> <span id="error-msg"></span>
    <pre id="error-trace" class="mt-2 small text-muted" style="display:none"></pre>
  </div>

</div><!-- /container -->

<footer>DISPOWER S.A.S E.S.P. &mdash; Sistema Generador de Informes &copy; 2026</footer>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
let currentFilename = "";

/* ─── Mostrar nombre de archivo en zona de drop ─── */
function showFile(input, dzId) {
  const lbl = document.getElementById(dzId + '-label');
  if (input.files.length) {
    lbl.innerHTML = `<span class="badge-file"><i class="bi bi-check2"></i>${input.files[0].name}</span>`;
  }
}

/* ─── Drag & drop visual ─── */
document.querySelectorAll('.drop-zone').forEach(dz => {
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('dragover');
    const inp = dz.querySelector('input[type=file]');
    if (inp && e.dataTransfer.files.length) {
      inp.files = e.dataTransfer.files;
      const evt = new Event('change');
      inp.dispatchEvent(evt);
    }
  });
});

/* ─── Cargar capítulos según cliente ─── */
document.getElementById('select-cliente').addEventListener('change', loadChapters);
window.addEventListener('DOMContentLoaded', loadChapters);

async function loadChapters() {
  const cliente = document.getElementById('select-cliente').value;
  if (!cliente) return;
  try {
    const res = await fetch(`/api/config/${cliente}`);
    const cfg = await res.json();
    renderChapters(cfg.capitulos || []);
  } catch(e) {
    document.getElementById('cap-tree').innerHTML = '<p class="text-danger small">Error cargando capítulos</p>';
  }
}

function renderChapters(caps) {
  const tree = document.getElementById('cap-tree');
  const topLevel = caps.filter(c => !c.parent);
  tree.innerHTML = '';
  topLevel.forEach(c => {
    appendCap(tree, c, caps, 1);
  });
}

function appendCap(container, cap, all, level) {
  const div = document.createElement('div');
  div.className = `cap-item level-${level}`;
  div.innerHTML = `<label>
    <input type="checkbox" name="cap_${cap.id}" value="${cap.id}" checked
           onchange="syncChildren(${cap.id}, this.checked)">
    <span>${cap.num}. ${cap.titulo}</span>
  </label>`;
  container.appendChild(div);

  const children = all.filter(c => c.parent === cap.id);
  children.forEach(child => appendCap(container, child, all, Math.min(level + 1, 3)));
}

function syncChildren(parentId, checked) {
  // No se pueden seleccionar hijos solos sin padre (se agregan automáticamente en backend)
}

function toggleAll(checked) {
  document.querySelectorAll('#cap-tree input[type=checkbox]').forEach(cb => cb.checked = checked);
}

/* ─── Progress animation ─── */
function setStep(n, pct) {
  document.getElementById('progress-bar').style.width = pct + '%';
  document.getElementById('progress-pct').textContent = pct + '%';
  for (let i = 1; i <= 4; i++) {
    const el = document.getElementById('step-' + i);
    if (i < n) el.className = 'step-dot mx-auto done';
    else if (i === n) el.className = 'step-dot mx-auto active';
    else el.className = 'step-dot mx-auto';
  }
}

async function fakeProgress() {
  setStep(1, 15); await sleep(600);
  setStep(2, 40); await sleep(700);
  setStep(3, 65); await sleep(800);
  setStep(4, 85); await sleep(600);
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

/* ─── Submit del formulario ─── */
document.getElementById('report-form').addEventListener('submit', async function(e) {
  e.preventDefault();
  const form = e.target;

  // Recopilar capítulos seleccionados
  const caps = [];
  document.querySelectorAll('#cap-tree input[type=checkbox]:checked').forEach(cb => {
    caps.push(parseInt(cb.value));
  });

  const formData = new FormData(form);
  formData.set('capitulos', JSON.stringify(caps));

  // Ocultar resultados previos
  document.getElementById('result-card').style.display = 'none';
  document.getElementById('error-card').style.display = 'none';
  document.getElementById('progress-wrap').style.display = 'block';

  // Animar progreso
  const progPromise = fakeProgress();

  try {
    const res = await fetch('/api/generate', { method: 'POST', body: formData });
    const data = await res.json();
    await progPromise;
    setStep(5, 100);
    await sleep(300);

    document.getElementById('progress-wrap').style.display = 'none';

    if (data.success) {
      showResult(data);
    } else {
      showError(data.error, data.trace);
    }
  } catch(err) {
    await progPromise;
    document.getElementById('progress-wrap').style.display = 'none';
    showError(err.toString());
  }
});

function showResult(data) {
  currentFilename = data.filename;
  const rc = document.getElementById('result-card');
  rc.style.display = 'block';

  const mr = data.metricas_resumen || {};
  const boxes = document.getElementById('metric-boxes');
  const fmt = n => n ? n.toLocaleString('es-CO') : '—';
  boxes.innerHTML = `
    <div class="col-6 col-md-3">
      <div class="metric-box">
        <div class="num">${fmt(mr.operaciones_total)}</div>
        <div class="lbl">Visitas operaciones</div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="metric-box">
        <div class="num">${fmt(mr.sac_total)}</div>
        <div class="lbl">PQR registradas</div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="metric-box">
        <div class="num">${fmt(mr.asistencia_total)}</div>
        <div class="lbl">Atenciones en sede</div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="metric-box">
        <div class="num">$ ${mr.facturacion_total ? (mr.facturacion_total/1e6).toFixed(1)+'M' : '—'}</div>
        <div class="lbl">Facturación total</div>
      </div>
    </div>`;

  const badge = document.getElementById('incons-badge');
  if (data.inconsistencias > 0) {
    badge.style.display = 'inline-flex';
    document.getElementById('incons-count').textContent = data.inconsistencias;
  } else {
    badge.style.display = 'none';
  }

  rc.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function showError(msg, trace) {
  const ec = document.getElementById('error-card');
  ec.style.display = 'block';
  document.getElementById('error-msg').textContent = msg;
  const traceEl = document.getElementById('error-trace');
  if (trace) { traceEl.style.display = 'block'; traceEl.textContent = trace; }
  else { traceEl.style.display = 'none'; }
  ec.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function downloadFile() {
  if (currentFilename) {
    window.location.href = `/api/download/${encodeURIComponent(currentFilename)}`;
  }
}
</script>
</body>
</html>
