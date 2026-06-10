"""
utils/data_processor.py - Lee y procesa archivos Excel, ejecuta validaciones cruzadas
"""
import pandas as pd
import numpy as np
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
MESES_INV = {v: k for k, v in MESES.items()}


class DataProcessor:
    def __init__(self, config: dict, mes: int, ano: int):
        self.config = config
        self.mes = mes
        self.ano = ano
        self.mes_nombre = MESES.get(mes, str(mes))
        self.schemas = config.get("db_schemas", {})
        self.dfs = {}
        self.inconsistencias = []

    # ------------------------------------------------------------------ #
    # Helpers internos de estado SAC
    # ------------------------------------------------------------------ #
    def _sac_es_abierto(self, serie: pd.Series) -> pd.Series:
        """
        Retorna máscara booleana de registros abiertos.
        Compatible con el formato antiguo (Estado PQRS = 'abierto/pendiente')
        y el nuevo (Semaforo = 'Crítico' / 'Moderado' / 'Leve').
        """
        schema = self.schemas.get("db_sac", {})
        valores_abiertos = schema.get("col_estado_abierto_valores", [])
        if valores_abiertos:
            # Nuevo formato: Semaforo con valores explícitos
            return serie.astype(str).isin([str(v) for v in valores_abiertos])
        else:
            # Formato antiguo: texto libre
            return serie.astype(str).str.lower().str.contains(
                "abierto|open|pendiente|critico|moderado|leve", na=False
            )

    def _sac_es_cerrado(self, serie: pd.Series) -> pd.Series:
        """Retorna máscara booleana de registros cerrados."""
        schema = self.schemas.get("db_sac", {})
        valor_cerrado = schema.get("col_estado_cerrado_valor", "")
        if valor_cerrado:
            return serie.astype(str) == str(valor_cerrado)
        else:
            return serie.astype(str).str.lower().str.contains("cerrado|closed", na=False)

    def _sac_es_hurto(self, df: pd.DataFrame) -> pd.Series:
        """
        Retorna máscara booleana de registros de hurto.
        En el nuevo formato, 'HURTO' aparece en el campo Concatenado.
        En el antiguo, en Tipificacion.
        """
        schema = self.schemas.get("db_sac", {})
        col_concat = schema.get("col_concatenado", "")
        col_tipif  = schema.get("col_tipificacion", "")
        valor_hurto = schema.get("col_hurto_valor", "HURTO")

        mask = pd.Series(False, index=df.index)
        if col_concat and col_concat in df.columns:
            mask |= df[col_concat].astype(str).str.upper().str.contains(valor_hurto, na=False)
        if col_tipif and col_tipif in df.columns:
            mask |= df[col_tipif].astype(str).str.lower().str.contains("hurto", na=False)
        return mask

    def _sac_bloquea_facturacion(self, df: pd.DataFrame) -> pd.Series:
        """
        Retorna máscara de registros que bloquean facturación.
        En el nuevo formato, 'BLOQUEA FACTURACION' aparece en Concatenado.
        En el antiguo se detecta por Tipificacion.
        """
        schema = self.schemas.get("db_sac", {})
        col_concat = schema.get("col_concatenado", "")
        col_tipif  = schema.get("col_tipificacion", "")
        valor_bloqueo = schema.get("col_bloquea_facturacion_valor", "BLOQUEA FACTURACION")

        mask = pd.Series(False, index=df.index)
        if col_concat and col_concat in df.columns:
            mask |= df[col_concat].astype(str).str.upper().str.contains(
                valor_bloqueo.upper(), na=False
            )
        if col_tipif and col_tipif in df.columns:
            mask |= df[col_tipif].astype(str).str.lower().str.contains(
                "suspens|bloqueo|bloquea", na=False
            )
        return mask

    # ------------------------------------------------------------------ #
    # Carga de archivos
    # ------------------------------------------------------------------ #
    def cargar_excel(self, archivos: dict):
        for db_type, path in archivos.items():
            schema = self.schemas.get(db_type, {})
            if db_type == "db_opex":
                self._cargar_opex(path, schema)
            elif db_type == "db_sac":
                self._cargar_sac(path, schema)
            elif db_type == "db_fact":
                self._cargar_fact(path, schema)
            elif db_type == "db_asistencia":
                self._cargar_asistencia(path, schema)
            else:
                wb_sheets = pd.ExcelFile(path).sheet_names
                self.dfs[db_type] = {s: pd.read_excel(path, sheet_name=s) for s in wb_sheets}

    def _cargar_opex(self, path, schema):
        sheet_ag  = schema.get("sheet_agendamiento", "Agendamiento")
        sheet_rep = schema.get("sheet_reposiciones", "Reposiciones ")
        df_ag  = pd.read_excel(path, sheet_name=sheet_ag)
        df_rep = pd.read_excel(path, sheet_name=sheet_rep)
        self.dfs["db_opex"] = {"agendamiento": df_ag, "reposiciones": df_rep}

    def _cargar_sac(self, path, schema):
        # Acepta tanto el formato antiguo ('DISPAC') como el nuevo ('Reporte asesores')
        sheet = schema.get("sheet", "Reporte asesores")
        df = pd.read_excel(path, sheet_name=sheet)

        # Normalizar nombres de columnas clave si el archivo usa la estructura nueva
        # Crear columna derivada "Estado PQRS" para compatibilidad interna si no existe
        col_estado = schema.get("col_estado", "Semaforo")
        if col_estado in df.columns and "Estado PQRS" not in df.columns:
            df["_estado_pqrs"] = df[col_estado]
        else:
            df["_estado_pqrs"] = df.get("Estado PQRS", pd.Series(dtype=str))

        # Crear columna derivada de mes desde FechaCreacion si no existe columna Mes
        col_fecha_ap = schema.get("col_fecha_apertura", "FechaCreacion")
        if col_fecha_ap in df.columns and "Mes" not in df.columns:
            df["_fecha_dt"] = pd.to_datetime(df[col_fecha_ap], dayfirst=True, errors="coerce")
            df["Mes"] = df["_fecha_dt"].dt.month
            df["Ano"] = df["_fecha_dt"].dt.year

        self.dfs["db_sac"] = df

    def _cargar_fact(self, path, schema):
        sheet = schema.get("sheet", "Hoja1")
        df = pd.read_excel(path, sheet_name=sheet)
        self.dfs["db_fact"] = df

    def _cargar_asistencia(self, path, schema):
        sheet = schema.get("sheet", "DISPAC")
        df = pd.read_excel(path, sheet_name=sheet)
        self.dfs["db_asistencia"] = df

    # ------------------------------------------------------------------ #
    # Validaciones cruzadas
    # ------------------------------------------------------------------ #
    def validar(self):
        self.inconsistencias = []
        self._validar_nulos()
        self._validar_opex_sac()
        self._validar_fact_sac()
        self._validar_hurtos()
        self._validar_semaforo_critico()
        return self.inconsistencias

    def _agregar_inconsistencia(self, base, hoja, nui, tipo, descripcion, recomendacion=""):
        self.inconsistencias.append({
            "base": base, "hoja": hoja, "nui": str(nui),
            "tipo": tipo, "descripcion": descripcion,
            "recomendacion": recomendacion
        })

    def _validar_nulos(self):
        schema_sac = self.schemas.get("db_sac", {})
        col_nui = schema_sac.get("col_nui", "NUI")
        if "db_sac" in self.dfs:
            df = self.dfs["db_sac"]
            nulos = df[col_nui].isna().sum() if col_nui in df.columns else 0
            if nulos > 0:
                sheet_name = schema_sac.get("sheet", "db_sac")
                self._agregar_inconsistencia(
                    "db_sac", sheet_name, "N/A",
                    "Valores nulos en NUI",
                    f"Se encontraron {nulos} registros sin NUI en db_sac.",
                    "Completar el campo NUI en los registros afectados."
                )
        schema_fact = self.schemas.get("db_fact", {})
        col_nui_f = schema_fact.get("col_nui", "nui")
        if "db_fact" in self.dfs:
            df = self.dfs["db_fact"]
            nulos = df[col_nui_f].isna().sum() if col_nui_f in df.columns else 0
            if nulos > 0:
                self._agregar_inconsistencia(
                    "db_fact", "Hoja1", "N/A",
                    "Valores nulos en NUI",
                    f"Se encontraron {nulos} registros sin NUI en db_fact.",
                    "Completar el campo NUI en la base de facturación."
                )

    def _validar_opex_sac(self):
        if "db_opex" not in self.dfs or "db_sac" not in self.dfs:
            return
        schema_opex = self.schemas.get("db_opex", {})
        schema_sac  = self.schemas.get("db_sac", {})

        df_ag  = self.dfs["db_opex"]["agendamiento"].copy()
        df_sac = self.dfs["db_sac"].copy()

        col_nui_opex = schema_opex.get("col_nui", "Elementored")
        col_func     = schema_opex.get("col_estado_func", "Estado Funcionamiento")
        col_nui_sac  = schema_sac.get("col_nui", "NUI")
        col_estado   = schema_sac.get("col_estado", "Semaforo")

        # Visitas con SISFV funcional → la PQR asociada debería estar cerrada
        correctivos_func = df_ag[
            df_ag[col_func].astype(str).str.lower().str.contains("si|funcional", na=False)
        ]

        for _, row in correctivos_func.iterrows():
            nui = str(row.get(col_nui_opex, ""))
            if not nui or nui == "nan":
                continue
            tickets_sac = df_sac[df_sac[col_nui_sac].astype(str) == nui]
            if tickets_sac.empty:
                continue
            abiertos = tickets_sac[self._sac_es_abierto(tickets_sac[col_estado])]
            if len(abiertos) > 0:
                self._agregar_inconsistencia(
                    "db_opex / db_sac",
                    f"Agendamiento / {schema_sac.get('sheet', 'Reporte asesores')}",
                    nui,
                    "Visita correctiva exitosa con PQR abierta",
                    f"El NUI {nui} registra SISFV funcional en db_opex pero tiene "
                    f"{len(abiertos)} PQR(s) en estado abierto/crítico en db_sac.",
                    "Verificar y cerrar los tickets correspondientes en el sistema de PQR."
                )

    def _validar_fact_sac(self):
        if "db_fact" not in self.dfs or "db_sac" not in self.dfs:
            return
        schema_fact = self.schemas.get("db_fact", {})
        schema_sac  = self.schemas.get("db_sac", {})

        df_fact = self.dfs["db_fact"].copy()
        df_sac  = self.dfs["db_sac"].copy()

        col_nui_f  = schema_fact.get("col_nui", "nui")
        col_nui_s  = schema_sac.get("col_nui", "NUI")
        col_estado = schema_sac.get("col_estado", "Semaforo")

        # Casos abiertos que bloquean facturación
        mask_abierto  = self._sac_es_abierto(df_sac[col_estado])
        mask_bloqueo  = self._sac_bloquea_facturacion(df_sac)
        bloqueados    = df_sac[mask_abierto & mask_bloqueo]

        nuis_bloqueados = set(bloqueados[col_nui_s].astype(str).tolist())
        nuis_facturados = set(df_fact[col_nui_f].astype(str).tolist())

        for nui in nuis_bloqueados & nuis_facturados:
            self._agregar_inconsistencia(
                "db_fact / db_sac",
                f"Hoja1 / {schema_sac.get('sheet', 'Reporte asesores')}",
                nui,
                "Facturación con PQR abierta que bloquea facturación",
                f"El NUI {nui} tiene una PQR abierta con 'BLOQUEA FACTURACION' en db_sac "
                f"pero registra facturación en db_fact.",
                "Revisar el estado del servicio del usuario y corregir la factura si corresponde."
            )

    def _validar_hurtos(self):
        if "db_sac" not in self.dfs:
            return
        schema_sac = self.schemas.get("db_sac", {})
        df_sac = self.dfs["db_sac"].copy()
        col_nui    = schema_sac.get("col_nui", "NUI")
        col_estado = schema_sac.get("col_estado", "Semaforo")

        mask_hurto  = self._sac_es_hurto(df_sac)
        hurtos = df_sac[mask_hurto]

        # NIU duplicados en registros de hurto
        dupes = hurtos[hurtos.duplicated(subset=[col_nui], keep=False)]
        for nui in dupes[col_nui].unique():
            cnt = len(dupes[dupes[col_nui] == nui])
            self._agregar_inconsistencia(
                "db_sac",
                schema_sac.get("sheet", "Reporte asesores"),
                nui,
                "NUI duplicado en base de hurtos",
                f"El NUI {nui} aparece {cnt} veces en registros de hurto.",
                "Depurar los registros duplicados en la base de hurtos."
            )

        # Hurtos simultáneamente abiertos y en db_sac como caso activo
        hurtos_abiertos = hurtos[self._sac_es_abierto(hurtos[col_estado])]
        if len(hurtos_abiertos) > 0:
            nuis = hurtos_abiertos[col_nui].dropna().unique()[:5]  # limitamos a 5
            for nui in nuis:
                self._agregar_inconsistencia(
                    "db_sac",
                    schema_sac.get("sheet", "Reporte asesores"),
                    nui,
                    "Caso de hurto activo sin cerrar",
                    f"El NUI {nui} tiene un registro de hurto en estado abierto/crítico.",
                    "Verificar si el hurto fue resuelto y actualizar el estado del ticket."
                )

    def _validar_semaforo_critico(self):
        """
        Validación nueva del formato Reporte asesores:
        Detecta PQRs críticas con más de 30 días sin cerrar.
        """
        if "db_sac" not in self.dfs:
            return
        schema_sac = self.schemas.get("db_sac", {})
        col_semaforo_dias = schema_sac.get("col_semaforo_dias", "Semaforo en Dias")
        col_estado        = schema_sac.get("col_estado", "Semaforo")
        col_nui           = schema_sac.get("col_nui", "NUI")

        df_sac = self.dfs["db_sac"].copy()
        if col_semaforo_dias not in df_sac.columns:
            return

        criticos_vencidos = df_sac[
            self._sac_es_abierto(df_sac[col_estado]) &
            (pd.to_numeric(df_sac[col_semaforo_dias], errors="coerce") > 30)
        ]
        if len(criticos_vencidos) > 0:
            self._agregar_inconsistencia(
                "db_sac",
                schema_sac.get("sheet", "Reporte asesores"),
                f"{len(criticos_vencidos)} NUIs",
                "PQRs críticas con más de 30 días sin atención",
                f"Se detectaron {len(criticos_vencidos)} PQRs en estado abierto/crítico "
                f"con más de 30 días de antigüedad sin cierre.",
                "Priorizar la gestión y cierre de estas PQRs para evitar incumplimientos contractuales."
            )

    # ------------------------------------------------------------------ #
    # Métricas para el informe
    # ------------------------------------------------------------------ #
    def get_metricas_operaciones(self) -> dict:
        if "db_opex" not in self.dfs:
            return {}
        schema = self.schemas.get("db_opex", {})
        df = self.dfs["db_opex"]["agendamiento"].copy()

        col_tipo     = schema.get("col_tipo_tarea", "Nombres sitios afectados")
        col_func     = schema.get("col_estado_func", "Estado Funcionamiento")
        col_proyecto = schema.get("col_proyecto", "Proyecto")

        total       = len(df)
        preventivos = df[df[col_tipo].astype(str).str.upper().str.contains("PREVENTIVO", na=False)]
        correctivos = df[~df[col_tipo].astype(str).str.upper().str.contains("PREVENTIVO", na=False)]
        funcionales    = df[df[col_func].astype(str).str.lower().str.contains("si|funcional", na=False)]
        no_funcionales = df[~df[col_func].astype(str).str.lower().str.contains("si|funcional", na=False)]

        por_municipio = {}
        for proy, grp in df.groupby(col_proyecto):
            func    = grp[col_func].astype(str).str.lower().str.contains("si|funcional", na=False).sum()
            no_func = len(grp) - func
            por_municipio[str(proy)] = {"funcional": int(func), "no_funcional": int(no_func), "total": len(grp)}

        return {
            "total": total,
            "preventivos": len(preventivos),
            "correctivos": len(correctivos),
            "funcionales": len(funcionales),
            "no_funcionales": len(no_funcionales),
            "por_municipio": por_municipio
        }

    def get_metricas_reposiciones(self) -> dict:
        if "db_opex" not in self.dfs:
            return {}
        schema = self.schemas.get("db_opex", {})
        df = self.dfs["db_opex"]["reposiciones"].copy()

        col_municipio = "municipio" if "municipio" in df.columns else schema.get("col_proyecto", "municipio")
        total_rows    = len(df)

        baterias      = int(df["bateria"].fillna(0).sum())      if "bateria"      in df.columns else 0
        controladores = int(df["Controlador "].fillna(0).sum()) if "Controlador " in df.columns else 0
        inversores    = int(df["inversor "].fillna(0).sum())    if "inversor "    in df.columns else 0

        por_municipio = {}
        if col_municipio in df.columns:
            for mun, grp in df.groupby(col_municipio):
                bat  = int(grp["bateria"].fillna(0).sum())      if "bateria"      in grp.columns else 0
                ctrl = int(grp["Controlador "].fillna(0).sum()) if "Controlador " in grp.columns else 0
                inv  = int(grp["inversor "].fillna(0).sum())    if "inversor "    in grp.columns else 0
                por_municipio[str(mun)] = {"bateria": bat, "controlador": ctrl, "inversor": inv, "total": bat+ctrl+inv}

        return {
            "total_intervenciones": total_rows,
            "baterias": baterias,
            "controladores": controladores,
            "inversores": inversores,
            "total_componentes": baterias + controladores + inversores,
            "por_municipio": por_municipio
        }

    def get_metricas_sac(self) -> dict:
        if "db_sac" not in self.dfs:
            return {}
        schema = self.schemas.get("db_sac", {})
        df = self.dfs["db_sac"].copy()

        col_estado    = schema.get("col_estado", "Semaforo")
        col_tipo      = schema.get("col_tipo", "Menu")
        col_tipif     = schema.get("col_tipificacion", "SubMenu1")
        col_subtipif  = schema.get("col_subtipificacion", "SubMenu2")
        col_municipio = schema.get("col_municipio", "Municipio")
        col_canal     = schema.get("col_canal", "canal")
        col_proyecto  = schema.get("col_proyecto", "NombreSeccionales")
        col_semaforo_dias = schema.get("col_semaforo_dias", "Semaforo en Dias")
        col_concat    = schema.get("col_concatenado", "Concatenado")
        col_asesor    = schema.get("col_asesor", "Creador_gestion")

        abiertos  = df[self._sac_es_abierto(df[col_estado])]  if col_estado in df.columns else df.iloc[0:0]
        cerrados  = df[self._sac_es_cerrado(df[col_estado])]  if col_estado in df.columns else df.iloc[0:0]
        hurtos    = df[self._sac_es_hurto(df)]
        bloqueos  = df[self._sac_bloquea_facturacion(df)]

        # Distribución del semáforo (nuevo campo exclusivo del formato nuevo)
        por_semaforo = {}
        if col_estado in df.columns:
            por_semaforo = df[col_estado].value_counts().to_dict()

        # PQRs críticas > 30 días
        criticos_vencidos = 0
        if col_semaforo_dias in df.columns:
            criticos_vencidos = int(
                (
                    self._sac_es_abierto(df[col_estado]) &
                    (pd.to_numeric(df[col_semaforo_dias], errors="coerce") > 30)
                ).sum()
            )

        por_tipo      = df[col_tipif].value_counts().head(10).to_dict()    if col_tipif     in df.columns else {}
        por_tipo_menu = df[col_tipo].value_counts().to_dict()               if col_tipo      in df.columns else {}
        por_municipio = df[col_municipio].value_counts().to_dict()          if col_municipio in df.columns else {}
        por_canal_sac = df[col_canal].value_counts().to_dict()              if col_canal     in df.columns else {}
        por_proyecto  = df[col_proyecto].value_counts().head(20).to_dict()  if col_proyecto  in df.columns else {}
        por_asesor    = df[col_asesor].value_counts().head(10).to_dict()    if col_asesor    in df.columns else {}

        # Días promedio de cierre (solo registros cerrados con días registrados)
        dias_prom_cierre = 0
        col_dias = schema.get("col_dias_cierre", "Dias para Cierre")
        if col_dias in df.columns and len(cerrados) > 0:
            dias_prom_cierre = round(
                pd.to_numeric(cerrados[col_dias], errors="coerce").mean(), 1
            ) or 0

        return {
            "total": len(df),
            "abiertos": len(abiertos),
            "cerrados": len(cerrados),
            "hurtos": len(hurtos),
            "bloqueos_facturacion": len(bloqueos),
            "criticos_vencidos": criticos_vencidos,
            "dias_prom_cierre": dias_prom_cierre,
            "por_semaforo":   {str(k): int(v) for k, v in por_semaforo.items()},
            "por_tipo":       {str(k): int(v) for k, v in por_tipo.items()},
            "por_tipo_menu":  {str(k): int(v) for k, v in por_tipo_menu.items()},
            "por_municipio":  {str(k): int(v) for k, v in por_municipio.items()},
            "por_canal":      {str(k): int(v) for k, v in por_canal_sac.items()},
            "por_proyecto":   {str(k): int(v) for k, v in por_proyecto.items()},
            "por_asesor":     {str(k): int(v) for k, v in por_asesor.items()},
        }

    def get_metricas_asistencia(self) -> dict:
        if "db_asistencia" not in self.dfs:
            return {}
        schema = self.schemas.get("db_asistencia", {})
        df = self.dfs["db_asistencia"].copy()

        col_canal    = schema.get("col_canal", "Canal ")
        col_proyecto = schema.get("col_proyecto", "PROYECTO")

        total     = len(df)
        por_canal = df[col_canal].value_counts().to_dict() if col_canal in df.columns else {}
        presencial = int(df[df[col_canal].astype(str).str.lower().str.contains(
            "oficina|presencial", na=False)].shape[0]) if col_canal in df.columns else 0
        digital = int(df[df[col_canal].astype(str).str.lower().str.contains(
            "whatsapp|digital|web|email|mail|linea", na=False)].shape[0]) if col_canal in df.columns else 0

        return {
            "total": total,
            "presencial": presencial,
            "digital": digital,
            "por_canal":   {str(k): int(v) for k, v in por_canal.items()},
            "por_proyecto": {str(k): int(v) for k, v in
                             df[col_proyecto].value_counts().to_dict().items()} if col_proyecto in df.columns else {}
        }

    def get_metricas_facturacion(self) -> dict:
        if "db_fact" not in self.dfs:
            return {}
        schema = self.schemas.get("db_fact", {})
        df = self.dfs["db_fact"].copy()

        col_total     = schema.get("col_total", "total")
        col_descuento = schema.get("col_descuento", "descuento")
        col_proyecto  = schema.get("col_proyecto", "address")
        col_nui       = schema.get("col_nui", "nui")

        total_facturado = float(df[col_total].sum())     if col_total     in df.columns else 0
        total_descuento = float(df[col_descuento].sum()) if col_descuento in df.columns else 0
        num_usuarios    = df[col_nui].nunique()          if col_nui       in df.columns else 0

        por_proyecto = {}
        if col_proyecto in df.columns:
            for proy, grp in df.groupby(col_proyecto):
                por_proyecto[str(proy)] = {
                    "total":    float(grp[col_total].sum())  if col_total in grp.columns else 0,
                    "usuarios": grp[col_nui].nunique()       if col_nui   in grp.columns else 0
                }

        return {
            "total_facturado": total_facturado,
            "total_descuento": total_descuento,
            "neto": total_facturado - total_descuento,
            "num_usuarios": int(num_usuarios),
            "por_proyecto": por_proyecto
        }

    def get_all_metricas(self) -> dict:
        return {
            "operaciones":   self.get_metricas_operaciones(),
            "reposiciones":  self.get_metricas_reposiciones(),
            "sac":           self.get_metricas_sac(),
            "asistencia":    self.get_metricas_asistencia(),
            "facturacion":   self.get_metricas_facturacion(),
            "inconsistencias": self.validar()
        }
