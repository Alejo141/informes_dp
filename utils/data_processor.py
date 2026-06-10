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
        self.dfs = {}          # DataFrames cargados
        self.inconsistencias = []

    # ------------------------------------------------------------------ #
    # Carga de archivos
    # ------------------------------------------------------------------ #
    def cargar_excel(self, archivos: dict):
        """
        archivos: dict con clave db_type → ruta del archivo
        Ejemplo: {"db_opex": "/tmp/db_opex.xlsx"}
        """
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
                # Carga genérica
                wb_sheets = pd.ExcelFile(path).sheet_names
                self.dfs[db_type] = {s: pd.read_excel(path, sheet_name=s) for s in wb_sheets}

    def _cargar_opex(self, path, schema):
        sheet_ag = schema.get("sheet_agendamiento", "Agendamiento")
        sheet_rep = schema.get("sheet_reposiciones", "Reposiciones ")
        df_ag = pd.read_excel(path, sheet_name=sheet_ag)
        df_rep = pd.read_excel(path, sheet_name=sheet_rep)
        self.dfs["db_opex"] = {"agendamiento": df_ag, "reposiciones": df_rep}

    def _cargar_sac(self, path, schema):
        sheet = schema.get("sheet", "DISPAC")
        df = pd.read_excel(path, sheet_name=sheet)
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
    # Validaciones
    # ------------------------------------------------------------------ #
    def validar(self):
        """Ejecuta todas las validaciones y retorna lista de inconsistencias."""
        self.inconsistencias = []
        self._validar_nulos()
        self._validar_opex_sac()
        self._validar_fact_sac()
        self._validar_hurtos()
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
            nulos = df[col_nui].isna().sum()
            if nulos > 0:
                self._agregar_inconsistencia(
                    "db_sac", "DISPAC", "N/A",
                    "Valores nulos en NIU",
                    f"Se encontraron {nulos} registros sin NIU en db_sac.",
                    "Completar el campo NIU en los registros afectados."
                )
        schema_fact = self.schemas.get("db_fact", {})
        col_nui_f = schema_fact.get("col_nui", "nui")
        if "db_fact" in self.dfs:
            df = self.dfs["db_fact"]
            nulos = df[col_nui_f].isna().sum()
            if nulos > 0:
                self._agregar_inconsistencia(
                    "db_fact", "Hoja1", "N/A",
                    "Valores nulos en NIU",
                    f"Se encontraron {nulos} registros sin NIU en db_fact.",
                    "Completar el campo NIU en la base de facturación."
                )

    def _validar_opex_sac(self):
        if "db_opex" not in self.dfs or "db_sac" not in self.dfs:
            return
        schema_opex = self.schemas.get("db_opex", {})
        schema_sac = self.schemas.get("db_sac", {})

        df_ag = self.dfs["db_opex"]["agendamiento"].copy()
        df_sac = self.dfs["db_sac"].copy()

        col_nui_opex = schema_opex.get("col_nui", "Elementored")
        col_func = schema_opex.get("col_estado_func", "Estado Funcionamiento")
        col_nui_sac = schema_sac.get("col_nui", "NUI")
        col_estado_sac = schema_sac.get("col_estado", "Estado PQRS")

        # Visitas correctivas con SISFV funcional → PQR debería estar cerrada
        correctivos_func = df_ag[
            (df_ag[col_func].astype(str).str.lower().str.contains("si|funcional", na=False))
        ]

        for _, row in correctivos_func.iterrows():
            nui = str(row.get(col_nui_opex, ""))
            if not nui or nui == "nan":
                continue
            tickets_abiertos = df_sac[
                (df_sac[col_nui_sac].astype(str) == nui) &
                (df_sac[col_estado_sac].astype(str).str.lower().str.contains("abierto|open|pendiente", na=False))
            ]
            if len(tickets_abiertos) > 0:
                self._agregar_inconsistencia(
                    "db_opex / db_sac", "Agendamiento / DISPAC", nui,
                    "Visita correctiva exitosa con PQR abierta",
                    f"El NIU {nui} registra SISFV funcional en db_opex pero tiene {len(tickets_abiertos)} PQR(s) abierta(s) en db_sac.",
                    "Verificar y cerrar los tickets correspondientes en el sistema de PQR."
                )

    def _validar_fact_sac(self):
        if "db_fact" not in self.dfs or "db_sac" not in self.dfs:
            return
        schema_fact = self.schemas.get("db_fact", {})
        schema_sac = self.schemas.get("db_sac", {})

        df_fact = self.dfs["db_fact"].copy()
        df_sac = self.dfs["db_sac"].copy()

        col_nui_f = schema_fact.get("col_nui", "nui")
        col_nui_s = schema_sac.get("col_nui", "NUI")
        col_estado_s = schema_sac.get("col_estado", "Estado PQRS")
        col_tipo_s = schema_sac.get("col_tipificacion", "Tipificacion")

        # Hurtos / suspensión en SAC → no deberían estar facturando
        hurtos_abiertos = df_sac[
            (df_sac[col_estado_s].astype(str).str.lower().str.contains("abierto|open|pendiente", na=False)) &
            (df_sac[col_tipo_s].astype(str).str.lower().str.contains("hurto|suspens", na=False))
        ]
        nuis_bloqueados = set(hurtos_abiertos[col_nui_s].astype(str).tolist())
        nuis_facturados = set(df_fact[col_nui_f].astype(str).tolist())

        overlap = nuis_bloqueados & nuis_facturados
        for nui in overlap:
            self._agregar_inconsistencia(
                "db_fact / db_sac", "Hoja1 / DISPAC", nui,
                "Facturación con caso abierto de hurto/suspensión",
                f"El NIU {nui} tiene un caso abierto de hurto o suspensión en db_sac pero registra facturación en db_fact.",
                "Revisar el estado del servicio del usuario y anular/corregir la factura si corresponde."
            )

    def _validar_hurtos(self):
        if "db_sac" not in self.dfs:
            return
        schema_sac = self.schemas.get("db_sac", {})
        df_sac = self.dfs["db_sac"].copy()

        col_nui = schema_sac.get("col_nui", "NUI")
        col_tipo = schema_sac.get("col_tipificacion", "Tipificacion")
        col_estado = schema_sac.get("col_estado", "Estado PQRS")

        hurtos = df_sac[df_sac[col_tipo].astype(str).str.lower().str.contains("hurto", na=False)]
        # Duplicados por NIU en hurtos
        dupes = hurtos[hurtos.duplicated(subset=[col_nui], keep=False)]
        if len(dupes) > 0:
            for nui in dupes[col_nui].unique():
                self._agregar_inconsistencia(
                    "db_sac", "DISPAC", nui,
                    "NIU duplicado en base de hurtos",
                    f"El NIU {nui} aparece {len(dupes[dupes[col_nui]==nui])} veces en registros de hurto.",
                    "Depurar los registros duplicados en la base de hurtos."
                )

    # ------------------------------------------------------------------ #
    # Métricas para el informe
    # ------------------------------------------------------------------ #
    def get_metricas_operaciones(self) -> dict:
        if "db_opex" not in self.dfs:
            return {}
        schema = self.schemas.get("db_opex", {})
        df = self.dfs["db_opex"]["agendamiento"].copy()

        col_tipo = schema.get("col_tipo_tarea", "Nombres sitios afectados")
        col_func = schema.get("col_estado_func", "Estado Funcionamiento")
        col_proyecto = schema.get("col_proyecto", "Proyecto")
        col_estado_tarea = schema.get("col_estado_tarea", "Estado")

        total = len(df)
        preventivos = df[df[col_tipo].astype(str).str.upper().str.contains("PREVENTIVO", na=False)]
        correctivos = df[~df[col_tipo].astype(str).str.upper().str.contains("PREVENTIVO", na=False)]

        funcionales = df[df[col_func].astype(str).str.lower().str.contains("si|funcional", na=False)]
        no_funcionales = df[~df[col_func].astype(str).str.lower().str.contains("si|funcional", na=False)]

        # Por municipio
        por_municipio = {}
        for proy, grp in df.groupby(col_proyecto):
            func = grp[col_func].astype(str).str.lower().str.contains("si|funcional", na=False).sum()
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

        col_municipio = schema.get("col_proyecto", "municipio") if "municipio" in df.columns else "municipio"
        total_rows = len(df)

        # Contar componentes
        baterias = int(df["bateria"].fillna(0).sum()) if "bateria" in df.columns else 0
        controladores = int(df["Controlador "].fillna(0).sum()) if "Controlador " in df.columns else 0
        inversores = int(df["inversor "].fillna(0).sum()) if "inversor " in df.columns else 0

        # Por municipio
        por_municipio = {}
        if col_municipio in df.columns:
            for mun, grp in df.groupby(col_municipio):
                bat = int(grp["bateria"].fillna(0).sum()) if "bateria" in grp.columns else 0
                ctrl = int(grp["Controlador "].fillna(0).sum()) if "Controlador " in grp.columns else 0
                inv = int(grp["inversor "].fillna(0).sum()) if "inversor " in grp.columns else 0
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

        col_estado = schema.get("col_estado", "Estado PQRS")
        col_tipo = schema.get("col_tipo", "Tipo de PQRS")
        col_tipif = schema.get("col_tipificacion", "Tipificacion")
        col_municipio = schema.get("col_municipio", "Municipio")

        abiertos = df[df[col_estado].astype(str).str.lower().str.contains("abierto|open|pendiente", na=False)]
        cerrados = df[df[col_estado].astype(str).str.lower().str.contains("cerrado|closed", na=False)]

        hurtos = df[df[col_tipif].astype(str).str.lower().str.contains("hurto", na=False)]

        por_tipo = df[col_tipif].value_counts().head(10).to_dict() if col_tipif in df.columns else {}
        por_municipio = df[col_municipio].value_counts().to_dict() if col_municipio in df.columns else {}

        return {
            "total": len(df),
            "abiertos": len(abiertos),
            "cerrados": len(cerrados),
            "hurtos": len(hurtos),
            "por_tipo": {str(k): int(v) for k, v in por_tipo.items()},
            "por_municipio": {str(k): int(v) for k, v in por_municipio.items()}
        }

    def get_metricas_asistencia(self) -> dict:
        if "db_asistencia" not in self.dfs:
            return {}
        schema = self.schemas.get("db_asistencia", {})
        df = self.dfs["db_asistencia"].copy()

        col_canal = schema.get("col_canal", "Canal ")
        col_proyecto = schema.get("col_proyecto", "PROYECTO")
        col_tipif = schema.get("col_tipificacion", "Tipificación")

        total = len(df)
        por_canal = df[col_canal].value_counts().to_dict() if col_canal in df.columns else {}
        presencial = int(df[df[col_canal].astype(str).str.lower().str.contains("oficina|presencial", na=False)].shape[0]) if col_canal in df.columns else 0
        digital = int(df[df[col_canal].astype(str).str.lower().str.contains("whatsapp|digital|web|email|mail", na=False)].shape[0]) if col_canal in df.columns else 0

        return {
            "total": total,
            "presencial": presencial,
            "digital": digital,
            "por_canal": {str(k): int(v) for k, v in por_canal.items()},
            "por_proyecto": {str(k): int(v) for k, v in df[col_proyecto].value_counts().to_dict().items()} if col_proyecto in df.columns else {}
        }

    def get_metricas_facturacion(self) -> dict:
        if "db_fact" not in self.dfs:
            return {}
        schema = self.schemas.get("db_fact", {})
        df = self.dfs["db_fact"].copy()

        col_total = schema.get("col_total", "total")
        col_descuento = schema.get("col_descuento", "descuento")
        col_proyecto = schema.get("col_proyecto", "address")

        total_facturado = float(df[col_total].sum()) if col_total in df.columns else 0
        total_descuento = float(df[col_descuento].sum()) if col_descuento in df.columns else 0
        num_usuarios = df[schema.get("col_nui","nui")].nunique() if schema.get("col_nui","nui") in df.columns else 0

        por_proyecto = {}
        if col_proyecto in df.columns:
            for proy, grp in df.groupby(col_proyecto):
                por_proyecto[str(proy)] = {
                    "total": float(grp[col_total].sum()) if col_total in grp.columns else 0,
                    "usuarios": grp[schema.get("col_nui","nui")].nunique() if schema.get("col_nui","nui") in grp.columns else 0
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
            "operaciones": self.get_metricas_operaciones(),
            "reposiciones": self.get_metricas_reposiciones(),
            "sac": self.get_metricas_sac(),
            "asistencia": self.get_metricas_asistencia(),
            "facturacion": self.get_metricas_facturacion(),
            "inconsistencias": self.validar()
        }
