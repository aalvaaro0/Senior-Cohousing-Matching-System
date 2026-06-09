#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
interfaz.py  —  Interfaz visual del Sistema Co-Housing Senior
Ejecutar con:  python -m streamlit run interfaz.py
"""

import streamlit as st
import pandas as pd
from collections import Counter
from algoritmo_compatibilidad import (
    TFG_CargadorOntologia, TFG_MotorCompatibilidad,
    NIVEL_LABEL, NIVEL_SHORT, NIVEL_COLOR_BG, NIVEL_COLOR_TXT,
    ETIQ_BARTHEL, ETIQ_MECLOBO,
    nivel_barthel, nivel_meclobo, riesgo_tinetti,
    interp_gds, interp_goldberg, interp_norton,
    interp_npi, interp_pfeiffer, interp_necpal,
    n_adl, TTL_FILE,
)

# CONFIGURACIÓN DE PÁGINA

st.set_page_config(
    page_title="Co-Housing Senior · TFG",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown("""
<style>
/* ── Fondo general oscuro ── */
html, body, .stApp, [class*="css"] {
    background-color: #0f1923 !important;
    color: #e8edf2 !important;
    font-family: 'Segoe UI', Arial, sans-serif !important;
}

/* ── Fondo del contenido principal ── */
.main .block-container {
    background-color: #0f1923 !important;
    padding: 2rem 2.5rem 3rem 2.5rem !important;
    max-width: 1300px !important;
}

/* ── Sidebar oscuro ── */
section[data-testid="stSidebar"] {
    background-color: #111d2b !important;
    border-right: 1px solid #1e3248;
}
section[data-testid="stSidebar"] * {
    color: #dce6f0 !important;
}
section[data-testid="stSidebar"] label {
    color: #a0bcd4 !important;
    font-size: 13px !important;
    font-weight: 600 !important;
}
/* Selectbox y slider dentro del sidebar */
section[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div,
section[data-testid="stSidebar"] [data-baseweb="select"] {
    background-color: #1a2e42 !important;
    border: 1px solid #2a4a6a !important;
    color: #e8edf2 !important;
    border-radius: 6px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] span {
    color: #e8edf2 !important;
}
section[data-testid="stSidebar"] .stSlider [data-testid="stTickBar"] {
    color: #5a7a9a !important;
}

/* ── Selectbox principal (fuera sidebar) ── */
[data-testid="stSelectbox"] > div > div,
[data-baseweb="select"] {
    background-color: #1a2e42 !important;
    border: 1px solid #2a4a6a !important;
    color: #e8edf2 !important;
    border-radius: 8px !important;
}
[data-baseweb="select"] span {
    color: #e8edf2 !important;
}
[data-baseweb="popover"] {
    background-color: #1a2e42 !important;
    border: 1px solid #2a4a6a !important;
}
[data-baseweb="menu"] {
    background-color: #1a2e42 !important;
}
[data-baseweb="menu"] li {
    color: #e8edf2 !important;
}
[data-baseweb="menu"] li:hover {
    background-color: #243d56 !important;
}

/* ── Sliders ── */
[data-testid="stSlider"] [data-baseweb="slider"] [role="slider"] {
    background-color: #3b82c4 !important;
}

/* ── Divider ── */
hr {
    border-color: #1e3248 !important;
    margin: 20px 0 !important;
}

/* ── Cabecera principal ── */
.header-banner {
    background: linear-gradient(135deg, #0d2137 0%, #1a4a7a 60%, #1e5fa8 100%);
    color: #ffffff;
    padding: 30px 40px;
    border-radius: 12px;
    margin-bottom: 28px;
    border: 1px solid #1e3a5c;
    box-shadow: 0 4px 20px rgba(0,0,0,0.45);
}
.header-banner h1 {
    margin: 0 0 6px 0;
    font-size: 24px;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: 0.3px;
}
.header-banner p {
    margin: 0;
    font-size: 13px;
    color: rgba(255,255,255,0.78);
    letter-spacing: 0.2px;
}

/* ── Título de sección ── */
.sec-title {
    font-size: 11px;
    font-weight: 700;
    color: #5aaae8;
    text-transform: uppercase;
    letter-spacing: 1.4px;
    border-bottom: 2px solid #1e3a5c;
    padding-bottom: 7px;
    margin: 24px 0 16px 0;
}

/* ── Badge de nivel ── */
.badge-nivel {
    display: inline-block;
    padding: 5px 16px;
    border-radius: 20px;
    font-weight: 700;
    font-size: 12px;
    letter-spacing: 0.4px;
    margin-bottom: 18px;
}

/* ── Título de grupo de columna ── */
.grupo-titulo {
    font-size: 12px;
    font-weight: 700;
    color: #7ab8e0;
    text-transform: uppercase;
    letter-spacing: 0.9px;
    margin: 0 0 10px 0;
    padding-bottom: 7px;
    border-bottom: 1px solid #1e3248;
}

/* ── Tarjeta de dato ── */
.dato-card {
    background: #162433;
    border: 1px solid #1e3248;
    border-left: 3px solid #2a6096;
    border-radius: 7px;
    padding: 9px 14px;
    margin-bottom: 6px;
    display: flex;
    align-items: baseline;
    gap: 10px;
}
.dato-card .dk {
    font-size: 11px;
    font-weight: 700;
    color: #6a9abf;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    min-width: 112px;
    flex-shrink: 0;
}
.dato-card .dv {
    font-size: 13.5px;
    font-weight: 600;
    color: #dce8f2;
}

/* ── Métricas en sidebar ── */
.metric-box {
    background: #162433;
    border: 1px solid #1e3248;
    border-radius: 10px;
    padding: 12px 8px;
    text-align: center;
    margin-bottom: 8px;
}
.metric-box .mval {
    font-size: 22px;
    font-weight: 800;
    color: #5aaae8;
    line-height: 1.1;
}
.metric-box .mlbl {
    font-size: 10px;
    color: #6a8fac;
    text-transform: uppercase;
    letter-spacing: 0.7px;
    margin-top: 3px;
}

/* ── Panel de residente (detalle pareja) ── */
.panel-residente {
    background: #162433;
    border: 1px solid #1e3a5c;
    border-radius: 10px;
    padding: 20px 18px;
    text-align: center;
}
.panel-residente .pr-nombre {
    font-size: 17px;
    font-weight: 700;
    color: #dce8f2;
    margin-bottom: 3px;
}
.panel-residente .pr-codigo {
    font-size: 12px;
    color: #5a7a9a;
    margin-bottom: 10px;
}

/* ── VS separador ── */
.vs-box {
    display: flex;
    align-items: center;
    justify-content: center;
    padding-top: 24px;
    font-size: 22px;
    font-weight: 800;
    color: #3b82c4;
}

/* ── Barra de afinidad ── */
.barra-wrap {
    background: #1e3248;
    border-radius: 6px;
    height: 14px;
    width: 100%;
    margin-top: 6px;
    overflow: hidden;
}
.barra-fill {
    height: 14px;
    border-radius: 6px;
}

/* ── Score grande ── */
.score-grande {
    font-size: 44px;
    font-weight: 900;
    text-align: center;
    line-height: 1;
    padding-top: 8px;
}
.score-estado {
    text-align: center;
    font-size: 12px;
    color: #6a8fac;
    margin-top: 4px;
    letter-spacing: 0.5px;
}

/* ── Nombre del residente en ficha ── */
.residente-nombre {
    font-size: 22px;
    font-weight: 800;
    color: #c8dff0;
}
.residente-codigo {
    font-size: 13px;
    color: #5a7a9a;
    font-weight: 500;
}

/* ── Tablas / DataFrames ── */
[data-testid="stDataFrame"] {
    border: 1px solid #1e3248 !important;
    border-radius: 10px !important;
    overflow: hidden;
    background-color: #111d2b !important;
}
[data-testid="stDataFrame"] * {
    color: #dce8f2 !important;
    background-color: transparent !important;
}

/* ── Caption / texto auxiliar ── */
[data-testid="stCaptionContainer"] p,
.stCaption {
    color: #5a7a9a !important;
}

/* ── Spinner ── */
[data-testid="stSpinner"] p {
    color: #7ab8e0 !important;
}

/* ── Info box ── */
[data-testid="stAlert"] {
    background-color: #162433 !important;
    border: 1px solid #1e3a5c !important;
    color: #a0c4e8 !important;
    border-radius: 8px !important;
}

/* ── Barra de progreso en columnas (ProgressColumn) ── */
[data-testid="stDataFrameResizable"] [role="gridcell"] {
    color: #dce8f2 !important;
}

/* ── Ocultar elementos de Streamlit ── */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
</style>
""", unsafe_allow_html=True)



# CARGA DE DATOS


@st.cache_resource
def cargar_sistema():
    cargador  = TFG_CargadorOntologia()
    motor     = TFG_MotorCompatibilidad()
    grafo     = cargador.cargar_grafo_ttl(TTL_FILE)
    datos_xls = cargador.cargar_datos_excel_xls()
    perfiles  = cargador.cargar_todos_perfiles(grafo, datos_xls)
    matriz, parejas, ranking = motor.calcular_matriz_afinidad_completa(perfiles)
    return perfiles, matriz, parejas, ranking




def badge_nivel(nivel: int) -> str:
    bg  = NIVEL_COLOR_BG.get(nivel, "334455")
    txt = NIVEL_COLOR_TXT.get(nivel, "ffffff")
    return (f'<span class="badge-nivel" style="background:#{bg}; color:#{txt};">'
            f'{NIVEL_LABEL.get(nivel, "—")}</span>')

def color_score(pct: float) -> str:
    if pct == 0:   return "#e05252"
    if pct >= 70:  return "#4caf82"
    if pct >= 40:  return "#e8a830"
    return "#e07040"

def barra_afinidad(pct: float) -> str:
    color = color_score(pct)
    ancho = min(pct, 100)
    return (f'<div class="barra-wrap">'
            f'<div class="barra-fill" style="width:{ancho:.1f}%; background:{color};"></div>'
            f'</div>')

def estado_texto(pct: float) -> str:
    if pct == 0:   return "Vetado"
    if pct >= 70:  return "Ideal"
    if pct >= 40:  return "Aceptable"
    return "Bajo"

def bool_txt(v) -> str:
    if v is True:  return "Si"
    if v is False: return "No"
    return "—"

def seccion(titulo: str):
    st.markdown(f'<div class="sec-title">{titulo}</div>', unsafe_allow_html=True)

def dato(etiqueta: str, valor) -> str:
    return (f'<div class="dato-card">'
            f'<span class="dk">{etiqueta}</span>'
            f'<span class="dv">{valor}</span>'
            f'</div>')

def metrica_sidebar(valor, etiqueta: str) -> str:
    return (f'<div class="metric-box">'
            f'<div class="mval">{valor}</div>'
            f'<div class="mlbl">{etiqueta}</div>'
            f'</div>')




st.markdown("""
<div class="header-banner">
    <h1>Sistema de Emparejamiento Co-Housing Senior</h1>
    <p>Prototipo de Filtrado Inteligente basado en Ontologias OWL y Reglas SWRL
       &nbsp;·&nbsp; TFG &nbsp;·&nbsp; Alvaro &nbsp;·&nbsp; Ingenieria Informatica &nbsp;·&nbsp; 2026</p>
</div>
""", unsafe_allow_html=True)

with st.spinner("Cargando ontologia y calculando compatibilidades..."):
    perfiles, matriz, parejas, ranking = cargar_sistema()

codigos = sorted(perfiles)



with st.sidebar:
    st.markdown("### Filtros y navegacion")
    st.markdown("---")

    niveles_opciones = ["Todos los niveles"] + [NIVEL_LABEL[n] for n in range(1, 7)]
    filtro_nivel = st.selectbox("Nivel de convivencia", niveles_opciones)

    if filtro_nivel == "Todos los niveles":
        codigos_filtrados = codigos
    else:
        nivel_num = next(n for n in range(1, 7) if NIVEL_LABEL[n] == filtro_nivel)
        codigos_filtrados = [c for c in codigos if perfiles[c]["nivel"] == nivel_num]

    opciones_selector = [
        f"{c}  —  {perfiles[c]['nombre']}  ({NIVEL_SHORT[perfiles[c]['nivel']]})"
        for c in codigos_filtrados
    ]

    seleccion = st.selectbox(
        f"Residente  ({len(codigos_filtrados)} disponibles)",
        opciones_selector,
    )
    codigo_sel = seleccion.split("  —  ")[0].strip()
    perfil     = perfiles[codigo_sel]

    top_n = st.slider("Candidatos a mostrar", min_value=5, max_value=30, value=10)

    st.markdown("---")
    st.markdown("### Estadisticas globales")

    scores_validos = [p["pct"] for p in parejas if p["pct"] > 0]
    vetadas        = sum(1 for p in parejas if p["pct"] == 0)
    media          = sum(scores_validos) / len(scores_validos) if scores_validos else 0

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.markdown(metrica_sidebar(len(perfiles),      "Residentes"),    unsafe_allow_html=True)
        st.markdown(metrica_sidebar(f"{vetadas:,}",     "Vetadas"),       unsafe_allow_html=True)
    with col_s2:
        st.markdown(metrica_sidebar(f"{len(parejas):,}", "Parejas"),      unsafe_allow_html=True)
        st.markdown(metrica_sidebar(f"{media:.1f}%",    "Afinidad media"),unsafe_allow_html=True)




seccion("Ficha del residente")

st.markdown(
    f'<div style="display:flex; align-items:center; gap:14px; margin-bottom:4px;">'
    f'<span class="residente-nombre">{perfil["nombre"]}</span>'
    f'<span class="residente-codigo">{codigo_sel}</span>'
    f'</div>',
    unsafe_allow_html=True,
)
st.markdown(badge_nivel(perfil["nivel"]), unsafe_allow_html=True)

col1, col2, col3 = st.columns(3, gap="medium")

with col1:
    st.markdown('<div class="grupo-titulo">Datos personales</div>', unsafe_allow_html=True)
    nb     = nivel_barthel(perfil.get("barthel"))
    nc     = nivel_meclobo(perfil.get("meclobo"))
    dietas = perfil["dieta"] - {"Ninguna"}
    st.markdown(dato("Edad",            f"{perfil.get('edad','—')} años"),       unsafe_allow_html=True)
    st.markdown(dato("Genero",          perfil.get("genero","—")),               unsafe_allow_html=True)
    st.markdown(dato("Fumador",         bool_txt(perfil.get("esFumador"))),      unsafe_allow_html=True)
    st.markdown(dato("Diabetes",        bool_txt(perfil.get("tieneDiabetes"))),  unsafe_allow_html=True)
    st.markdown(dato("Silla de ruedas", bool_txt(perfil.get("usaSillaRuedas"))),unsafe_allow_html=True)
    st.markdown(dato("Perd. memoria",   bool_txt(perfil.get("perdidaMemoria"))), unsafe_allow_html=True)
    st.markdown(dato("Mascota",         bool_txt(perfil.get("tieneMascota"))),   unsafe_allow_html=True)
    st.markdown(dato("Presupuesto",     f"{perfil.get('presupuesto','—')} EUR"), unsafe_allow_html=True)
    st.markdown(dato("Dieta",           ", ".join(dietas) or "Estandar"),        unsafe_allow_html=True)

with col2:
    st.markdown('<div class="grupo-titulo">Escalas clinicas</div>', unsafe_allow_html=True)
    st.markdown(dato("Barthel",   f"{perfil.get('barthel','—')}  ·  {ETIQ_BARTHEL.get(nb,'—') if nb is not None else '—'}"), unsafe_allow_html=True)
    st.markdown(dato("MEC-Lobo",  f"{perfil.get('meclobo','—')}  ·  {ETIQ_MECLOBO.get(nc,'—') if nc is not None else '—'}"), unsafe_allow_html=True)
    st.markdown(dato("Tinetti",   f"{perfil.get('tinetti','—')}  ·  {riesgo_tinetti(perfil.get('tinetti')) or '—'} riesgo caida"), unsafe_allow_html=True)
    st.markdown(dato("GDS",       f"{perfil.get('gds','—')}  ·  {interp_gds(perfil.get('gds'))}"),             unsafe_allow_html=True)
    st.markdown(dato("Goldberg",  f"{perfil.get('goldberg','—')}  ·  {interp_goldberg(perfil.get('goldberg'))}"), unsafe_allow_html=True)
    st.markdown(dato("Norton",    f"{perfil.get('norton','—')}  ·  {interp_norton(perfil.get('norton'))}"),     unsafe_allow_html=True)
    st.markdown(dato("NPI",       f"{perfil.get('npi','—')}  ·  {interp_npi(perfil.get('npi'))}"),             unsafe_allow_html=True)
    st.markdown(dato("Pfeiffer",  f"{perfil.get('pfeiffer','—')}  ·  {interp_pfeiffer(perfil.get('pfeiffer'))}"), unsafe_allow_html=True)
    st.markdown(dato("NECPAL",    f"{perfil.get('necpal','—')}  ·  {interp_necpal(perfil.get('necpal'))}"),     unsafe_allow_html=True)

with col3:
    st.markdown('<div class="grupo-titulo">Estilo de vida</div>', unsafe_allow_html=True)
    ciclo = list(perfil["cicloSueno"])[0] if perfil["cicloSueno"] else "—"
    st.markdown(dato("Ciclo sueno",     ciclo),                                          unsafe_allow_html=True)
    st.markdown(dato("Idiomas",         ", ".join(perfil["idiomas"]) or "—"),            unsafe_allow_html=True)
    st.markdown(dato("Nivel actividad", ", ".join(perfil["nivelActividad"]) or "—"),    unsafe_allow_html=True)
    st.markdown(dato("Dias activ./sem.",str(perfil.get("diasActividad","—"))),           unsafe_allow_html=True)
    st.markdown(dato("Zonas comunes",   ", ".join(perfil["zonasComunes"]) or "—"),      unsafe_allow_html=True)
    st.markdown(dato("Hab. compartida", bool_txt(perfil.get("quiereHabComp"))),         unsafe_allow_html=True)
    st.markdown(dato("Visitas frec.",   bool_txt(perfil.get("recibeVisitas"))),          unsafe_allow_html=True)
    st.markdown(dato("AVD",             f"{n_adl(perfil)} / 3"),                         unsafe_allow_html=True)
    if perfil.get("hobby"):
        st.markdown(dato("Hobby",       perfil["hobby"]),    unsafe_allow_html=True)
    if perfil.get("religion"):
        st.markdown(dato("Religion",    perfil["religion"]), unsafe_allow_html=True)




st.divider()
seccion("Candidatos mas compatibles")

candidatos = ranking.get(codigo_sel, [])
n_validos  = sum(1 for s, _, _ in candidatos if s > 0)
n_vetados  = len(candidatos) - n_validos

st.caption(f"{n_validos} candidatos con afinidad positiva   ·   {n_vetados} vetados")

filas_tabla = []
for pos, (puntuacion, codigo_b, detalle) in enumerate(candidatos[:top_n], 1):
    pb      = perfiles[codigo_b]
    motivos = sorted(detalle, key=lambda x: abs(x[1]), reverse=True)[:2]
    razones = "  ·  ".join(
        f"{'[+]' if pts > 0 else '[-]'} {txt[:40]}" for txt, pts in motivos
    )
    filas_tabla.append({
        "Pos.":               pos,
        "Codigo":             codigo_b,
        "Nombre":             pb["nombre"],
        "Nivel":              NIVEL_SHORT[pb["nivel"]],
        "Afinidad %":         puntuacion,
        "Estado":             estado_texto(puntuacion),
        "Razones principales": razones,
    })

st.dataframe(
    pd.DataFrame(filas_tabla),
    use_container_width=True,
    hide_index=True,
    column_config={
        "Pos.":    st.column_config.NumberColumn(width=60),
        "Codigo":  st.column_config.TextColumn(width=100),
        "Nombre":  st.column_config.TextColumn(width=130),
        "Nivel":   st.column_config.TextColumn(width=180),
        "Afinidad %": st.column_config.ProgressColumn(
            "Afinidad %", min_value=0, max_value=100,
            format="%.1f%%", width="medium",
        ),
        "Estado":  st.column_config.TextColumn(width=90),
        "Razones principales": st.column_config.TextColumn(width="large"),
    },
)




st.divider()
seccion("Analisis detallado de pareja")

opciones_pareja = [
    f"{codigo_b}  —  {perfiles[codigo_b]['nombre']}  ({punt:.1f}%)"
    for punt, codigo_b, _ in candidatos if punt > 0
]

if opciones_pareja:
    seleccion_b  = st.selectbox("Selecciona el candidato a analizar:", opciones_pareja)
    codigo_b_sel = seleccion_b.split("  —  ")[0].strip()
    punt_sel, _, detalle_sel = next(
        (s, c, d) for s, c, d in candidatos if c == codigo_b_sel
    )
    pb_sel = perfiles[codigo_b_sel]

    col_pa, col_vs, col_pb = st.columns([5, 1, 5], gap="small")

    with col_pa:
        st.markdown(
            f'<div class="panel-residente">'
            f'<div class="pr-nombre">{perfil["nombre"]}</div>'
            f'<div class="pr-codigo">{codigo_sel}</div>'
            f'{badge_nivel(perfil["nivel"])}'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col_vs:
        st.markdown('<div class="vs-box">&#8596;</div>', unsafe_allow_html=True)

    with col_pb:
        st.markdown(
            f'<div class="panel-residente">'
            f'<div class="pr-nombre">{pb_sel["nombre"]}</div>'
            f'<div class="pr-codigo">{codigo_b_sel}</div>'
            f'{badge_nivel(pb_sel["nivel"])}'
            f'</div>',
            unsafe_allow_html=True,
        )

   
    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
    col_pct, col_bar = st.columns([1, 5], gap="medium")

    with col_pct:
        color_p = color_score(punt_sel)
        st.markdown(
            f'<div class="score-grande" style="color:{color_p};">{punt_sel:.1f}%</div>'
            f'<div class="score-estado">{estado_texto(punt_sel)}</div>',
            unsafe_allow_html=True,
        )

    with col_bar:
        st.markdown(
            f'<div style="margin-top:26px;">{barra_afinidad(punt_sel)}</div>',
            unsafe_allow_html=True,
        )

    
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    st.markdown(
        '<p style="font-weight:700; color:#a0c4e0; margin-bottom:8px;">'
        'Reglas aplicadas en el calculo:</p>',
        unsafe_allow_html=True,
    )

    filas_reglas = []
    for texto, pts in sorted(detalle_sel, key=lambda x: x[1], reverse=True):
        if texto.startswith("VETO"):
            efecto = "Veto"
        elif pts > 0:
            efecto = "Positivo"
        else:
            efecto = "Negativo"
        filas_reglas.append({"Regla": texto, "Puntos": pts, "Efecto": efecto})

    st.dataframe(
        pd.DataFrame(filas_reglas),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Regla":  st.column_config.TextColumn(width="large"),
            "Puntos": st.column_config.NumberColumn(width=100, format="%+d pts"),
            "Efecto": st.column_config.TextColumn(width=100),
        },
    )

else:
    st.info("Este residente tiene todos los candidatos vetados.")



st.divider()
col_top, col_sep, col_dist = st.columns([3, 0.08, 2], gap="medium")

with col_top:
    seccion("Top 10 parejas mas compatibles")
    top10 = [p for p in parejas if p["pct"] > 0][:10]
    filas_top = []
    for i, par in enumerate(top10, 1):
        filas_top.append({
            "Pos.":        i,
            "Residente A": f"{par['A']} — {par['nombre_A']}",
            "Nivel A":     NIVEL_SHORT[par["nivel_A"]],
            "Residente B": f"{par['B']} — {par['nombre_B']}",
            "Nivel B":     NIVEL_SHORT[par["nivel_B"]],
            "Afinidad %":  par["pct"],
        })
    st.dataframe(
        pd.DataFrame(filas_top),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Pos.":        st.column_config.NumberColumn(width=55),
            "Residente A": st.column_config.TextColumn(width="medium"),
            "Nivel A":     st.column_config.TextColumn(width=140),
            "Residente B": st.column_config.TextColumn(width="medium"),
            "Nivel B":     st.column_config.TextColumn(width=140),
            "Afinidad %":  st.column_config.ProgressColumn(
                "Afinidad %", min_value=0, max_value=100,
                format="%.1f%%", width="medium",
            ),
        },
    )

with col_dist:
    seccion("Distribucion por nivel de convivencia")
    conteo = Counter(p["nivel"] for p in perfiles.values())
    df_dist = pd.DataFrame({
        "Nivel":      [NIVEL_SHORT[n] for n in range(1, 7)],
        "Residentes": [conteo.get(n, 0) for n in range(1, 7)],
    }).set_index("Nivel")
    st.bar_chart(df_dist, color="#3b82c4", use_container_width=True)

