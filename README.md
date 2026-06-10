# Generador de Informes Mensuales – DISPOWER

## Descripción
Aplicación **Streamlit** que genera automáticamente informes Word mensuales
a partir de archivos Excel de operaciones, SAC, facturación y asistencia.

## Requisitos
- Python 3.10+
- pip

## Instalación

```bash
# 1. Descomprimir el ZIP
unzip informe_app.zip
cd informe_app

# 2. Crear entorno virtual (recomendado)
python -m venv venv
source venv/bin/activate        # Linux / Mac
venv\Scripts\activate           # Windows

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la app
streamlit run streamlit_app.py
```

Se abrirá automáticamente en: **http://localhost:8501**

## Cómo usar
1. En el panel lateral, selecciona **cliente**, **mes** y **año**
2. Marca los **capítulos** que deseas incluir (o usa Todos / Ninguno)
3. En el área principal, carga:
   - El **informe Word del mes anterior** (plantilla base, opcional)
   - Los archivos Excel: `db_opex`, `db_sac`, `db_fact`, `db_asistencia`
4. Haz clic en **⚡ Generar Informe Word**
5. Revisa el resumen de métricas y descarga el documento

## Estructura del proyecto
```
informe_app/
├── streamlit_app.py          ← App principal (ejecutar este)
├── requirements.txt
├── README.md
├── config/
│   ├── DISPAC.json           ← Configuración completa (32 capítulos)
│   ├── ISA.json
│   ├── GENSA.json
│   ├── CENDENAR.json
│   ├── CENS.json
│   └── IPSE.json
├── utils/
│   ├── data_processor.py     ← Lectura y validaciones cruzadas
│   ├── chart_generator.py    ← Generación de gráficos
│   └── word_generator.py     ← Construcción del Word final
└── output/                   ← Informes generados
```

## Agregar un nuevo cliente
Crea `config/NUEVO_CLIENTE.json` siguiendo el modelo de `DISPAC.json`.
No se modifica ningún archivo Python.

## Validaciones automáticas entre bases
| Validación | Bases cruzadas |
|-----------|---------------|
| Mantenimiento exitoso con PQR abierta | db_opex ↔ db_sac |
| Usuario con hurto/suspensión facturado | db_fact ↔ db_sac |
| NIU duplicado en base de hurtos | db_sac |
| NIU nulos | db_sac, db_fact |
