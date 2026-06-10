"""
utils/chart_generator.py - Genera gráficos PNG para insertar en el documento Word
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from pathlib import Path
import io

# Paleta corporativa Dispower
COLORS = ["#1B4F8A", "#2E86C1", "#85C1E9", "#F39C12", "#E74C3C", "#27AE60",
          "#8E44AD", "#17A589", "#D35400", "#2C3E50"]

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.titlesize": 10,
    "axes.titleweight": "bold",
    "figure.dpi": 120,
    "axes.spines.top": False,
    "axes.spines.right": False,
})


def _save_fig(fig, output_dir: str, name: str) -> str:
    path = Path(output_dir) / f"{name}.png"
    fig.savefig(str(path), bbox_inches="tight", dpi=120, facecolor="white")
    plt.close(fig)
    return str(path)


def grafico_barras_operaciones(metricas: dict, output_dir: str) -> str:
    """Mantenimientos preventivos vs correctivos por municipio."""
    por_mun = metricas.get("por_municipio", {})
    if not por_mun:
        return ""

    municipios = list(por_mun.keys())[:12]
    func = [por_mun[m]["funcional"] for m in municipios]
    no_func = [por_mun[m]["no_funcional"] for m in municipios]

    x = np.arange(len(municipios))
    w = 0.4
    fig, ax = plt.subplots(figsize=(10, 4.5))
    b1 = ax.bar(x - w/2, func, w, label="Funcional", color=COLORS[1], edgecolor="white")
    b2 = ax.bar(x + w/2, no_func, w, label="No funcional", color=COLORS[4], edgecolor="white")

    ax.set_xticks(x)
    ax.set_xticklabels([m[:14] for m in municipios], rotation=30, ha="right", fontsize=8)
    ax.set_title("Estado del parque por municipio")
    ax.set_ylabel("N° de visitas")
    ax.legend()
    ax.bar_label(b1, padding=2, fontsize=7)
    ax.bar_label(b2, padding=2, fontsize=7)
    fig.tight_layout()
    return _save_fig(fig, output_dir, "operaciones_municipio")


def grafico_torta_funcionalidad(metricas: dict, output_dir: str) -> str:
    func = metricas.get("funcionales", 0)
    no_func = metricas.get("no_funcionales", 0)
    if func + no_func == 0:
        return ""

    fig, ax = plt.subplots(figsize=(5, 4))
    vals = [func, no_func]
    labels = [f"Funcional\n({func})", f"No Funcional\n({no_func})"]
    ax.pie(vals, labels=labels, colors=[COLORS[5], COLORS[4]],
           autopct="%1.1f%%", startangle=90,
           wedgeprops={"edgecolor": "white", "linewidth": 1.5})
    ax.set_title("Funcionalidad de visitas")
    fig.tight_layout()
    return _save_fig(fig, output_dir, "operaciones_torta")


def grafico_pqr_tipo(metricas: dict, output_dir: str) -> str:
    por_tipo = metricas.get("por_tipo", {})
    if not por_tipo:
        return ""

    tipos = list(por_tipo.keys())[:8]
    valores = [por_tipo[t] for t in tipos]

    fig, ax = plt.subplots(figsize=(9, 4))
    bars = ax.barh(tipos[::-1], valores[::-1], color=COLORS[0], edgecolor="white")
    ax.set_title("PQR por tipificación")
    ax.set_xlabel("Cantidad")
    ax.bar_label(bars, padding=3, fontsize=8)
    fig.tight_layout()
    return _save_fig(fig, output_dir, "sac_por_tipo")


def grafico_sac_estado(metricas: dict, output_dir: str) -> str:
    abiertos = metricas.get("abiertos", 0)
    cerrados = metricas.get("cerrados", 0)
    hurtos = metricas.get("hurtos", 0)
    if abiertos + cerrados + hurtos == 0:
        return ""

    fig, ax = plt.subplots(figsize=(5, 4))
    labels = ["Abiertos", "Cerrados", "Hurtos"]
    vals = [abiertos, cerrados, hurtos]
    colors = [COLORS[3], COLORS[5], COLORS[4]]
    bars = ax.bar(labels, vals, color=colors, edgecolor="white", width=0.5)
    ax.set_title("Estado de PQR")
    ax.set_ylabel("Cantidad")
    ax.bar_label(bars, padding=3)
    fig.tight_layout()
    return _save_fig(fig, output_dir, "sac_estado")


def grafico_asistencia_canal(metricas: dict, output_dir: str) -> str:
    por_canal = metricas.get("por_canal", {})
    if not por_canal:
        return ""

    canales = list(por_canal.keys())[:6]
    valores = [por_canal[c] for c in canales]

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.pie(valores, labels=canales, autopct="%1.1f%%",
           colors=COLORS[:len(canales)],
           wedgeprops={"edgecolor": "white", "linewidth": 1.5},
           startangle=90)
    ax.set_title("Atención usuarios por canal")
    fig.tight_layout()
    return _save_fig(fig, output_dir, "asistencia_canal")


def grafico_facturacion_proyecto(metricas: dict, output_dir: str) -> str:
    por_proy = metricas.get("por_proyecto", {})
    if not por_proy:
        return ""

    proyectos = list(por_proy.keys())[:10]
    totales = [por_proy[p]["total"] / 1_000_000 for p in proyectos]

    fig, ax = plt.subplots(figsize=(10, 4.5))
    bars = ax.bar(proyectos, totales, color=COLORS[0], edgecolor="white")
    ax.set_xticks(range(len(proyectos)))
    ax.set_xticklabels([p[:14] for p in proyectos], rotation=30, ha="right", fontsize=8)
    ax.set_title("Facturación por proyecto (millones COP)")
    ax.set_ylabel("Millones COP")
    ax.bar_label(bars, fmt="%.1f M", padding=3, fontsize=7)
    fig.tight_layout()
    return _save_fig(fig, output_dir, "facturacion_proyecto")


def grafico_reposiciones(metricas: dict, output_dir: str) -> str:
    por_mun = metricas.get("por_municipio", {})
    if not por_mun:
        return ""

    municipios = list(por_mun.keys())[:10]
    baterias = [por_mun[m]["bateria"] for m in municipios]
    controladores = [por_mun[m]["controlador"] for m in municipios]
    inversores = [por_mun[m]["inversor"] for m in municipios]

    x = np.arange(len(municipios))
    w = 0.25
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(x - w, baterias, w, label="Batería", color=COLORS[1])
    ax.bar(x, controladores, w, label="Controlador", color=COLORS[3])
    ax.bar(x + w, inversores, w, label="Inversor", color=COLORS[4])
    ax.set_xticks(x)
    ax.set_xticklabels([m[:12] for m in municipios], rotation=30, ha="right", fontsize=8)
    ax.set_title("Reposiciones por componente y municipio")
    ax.set_ylabel("Unidades")
    ax.legend()
    fig.tight_layout()
    return _save_fig(fig, output_dir, "reposiciones")


def generar_todos_los_graficos(metricas: dict, output_dir: str) -> dict:
    """Genera todos los gráficos disponibles y retorna rutas."""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    graficos = {}

    ops = metricas.get("operaciones", {})
    sac = metricas.get("sac", {})
    asist = metricas.get("asistencia", {})
    fact = metricas.get("facturacion", {})
    repos = metricas.get("reposiciones", {})

    if ops:
        graficos["operaciones_municipio"] = grafico_barras_operaciones(ops, output_dir)
        graficos["operaciones_torta"] = grafico_torta_funcionalidad(ops, output_dir)

    if sac:
        graficos["sac_tipo"] = grafico_pqr_tipo(sac, output_dir)
        graficos["sac_estado"] = grafico_sac_estado(sac, output_dir)

    if asist:
        graficos["asistencia_canal"] = grafico_asistencia_canal(asist, output_dir)

    if fact:
        graficos["facturacion_proyecto"] = grafico_facturacion_proyecto(fact, output_dir)

    if repos:
        graficos["reposiciones"] = grafico_reposiciones(repos, output_dir)

    return {k: v for k, v in graficos.items() if v}
