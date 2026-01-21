import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os
import re

# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="Dashboard - AVP - Manutenção - Processo",
    layout="wide"
)

# =====================================================
# PALETA
# =====================================================
AMARELO = "#f1c40f"
PRETO = "#0e1117"
CINZA_ESCURO = "#1c1f26"
VERDE = "#1dd268"

# =====================================================
# CSS
# =====================================================
st.markdown(f"""
<style>
html, body, [class*="css"] {{
    background-color: {PRETO};
    color: white;
}}
[data-testid="metric-container"] {{
    background-color: {CINZA_ESCURO};
    padding: 15px;
    border-radius: 10px;
}}
</style>
""", unsafe_allow_html=True)

# =====================================================
# CAMINHOS
# =====================================================
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CAMINHO_IW28 = os.path.join(BASE_DIR, "data1", "IW28.xlsx")
CAMINHO_IW47 = os.path.join(BASE_DIR, "data1", "IW47.xlsx")
CAMINHO_SPLAN = os.path.join(BASE_DIR, "data1", "SPLAN.xlsx")

PASTA_FOTOS = os.path.join(BASE_DIR, "assets")
# =====================================================
# MAPA DE FOTOS
# =====================================================
MAPA_FOTOS = {
    "ANDRE SAVIO NADAL": "Andre.png",
    "CARLOS DANIEL ANTUNES": "Carlos.jpeg",
    "DEJAIR JOSE SANTOS LIVRAMENTO": "Dejair.jpeg",
    "JEAN WILLIAN SANTOS": "Jean.jpeg",
    "RODRIGO PINHEIRO": "Rodrigo_Pinheiro.jpeg",
    "NATALIA BARRETO": "Natalia.jpeg",
    "EZEQUIEL ALEIXO": "Ezequiel.jpg",
    "RODRIGO CHARLES VIEIRA": "Rodrigo_Charles.jpeg",
    "RONALDO CORREA DA ROCHA": "Ronaldo.PNG",
    "ALISSON PAULO GASTAO": "Alisson.jpeg",
    "ROMARIO KASPCHAK": "Romario.jpeg",
    "THIAGO MAURICIO AZEVEDO": "Thiago.jpg",
    "CRISTIANO IATCZAKI": "Cristiano.jpg",
    "EVANDRO LOPES SANTANA": "Evandro.jpg",
    "VICTOR EMANUEL PAES DE MELLO": "Victor.jpeg",
}

def obter_foto(nome):
    arq = MAPA_FOTOS.get(nome)
    if not arq:
        return None
    caminho = os.path.join(PASTA_FOTOS, arq)
    return caminho if os.path.exists(caminho) else None

# =====================================================
# CARREGAMENTO
# =====================================================
@st.cache_data(ttl=300)
def carregar_iw28():
    df = pd.read_excel(CAMINHO_IW28)
    df.columns = [c.strip() for c in df.columns]
    df["Conclusão desejada"] = pd.to_datetime(df["Conclusão desejada"], errors="coerce")
    df["Criado em"] = pd.to_datetime(df["Criado em"], errors="coerce")
    return df

@st.cache_data(ttl=300)
def carregar_iw47():
    df = pd.read_excel(CAMINHO_IW47)
    df.columns = [c.strip() for c in df.columns]
    df["Data de lançamento"] = pd.to_datetime(df["Data de lançamento"], errors="coerce")
    df["Minutos"] = pd.to_numeric(df["Trabalho real"], errors="coerce")
    df["Horas"] = df["Minutos"] / 60
    return df.dropna(subset=["Nome do empregado", "Horas", "Data de lançamento"])

@st.cache_data(ttl=300)
def carregar_splan():
    df = pd.read_excel(CAMINHO_SPLAN)
    df.columns = [c.strip() for c in df.columns]
    df["Data da Investigação"] = pd.to_datetime(df["Data da Investigação"], errors="coerce")
    df["Mês"] = df["Data da Investigação"].dt.to_period("M").astype(str)
    return df.dropna(subset=["Criador da Investigação", "Data da Investigação"])

# =====================================================
# IW28 – STATUS
# =====================================================
def ajustar_status(df):
    def mapa(s):
        s = str(s)
        if "CNF" in s:
            return "Confirmada"
        if "ORDA" in s or "MSPR" in s:
            return "Em aberto"
        return "Outros"
    df["Status do sistema"] = df["Status do sistema"].apply(mapa)
    return df

CORES_STATUS = {
    "Confirmada": VERDE,
    "Em aberto": AMARELO,
    "Outros": "#ec7c40"
}

# =====================================================
# LEITURA
# =====================================================
df_iw28 = ajustar_status(carregar_iw28())
df_iw47 = carregar_iw47()
df_splan = carregar_splan()

# =====================================================
# CÓDIGO MACRO
# =====================================================
def manter_codigo_macro(txt):
    if pd.isna(txt):
        return None
    m = re.search(r"[A-Z]{2}-\d{5}", str(txt))
    return m.group() if m else None

df_iw28["Local de instalação"] = df_iw28["Local de instalação"].apply(manter_codigo_macro)

# =====================================================
# ABAS
# =====================================================
aba_iw28, aba_iw47, aba_splan = st.tabs(
    ["📋 Programação – IW28", "⏱️ IW47 – Apropriação", "🕵️ Splan – Investigações"]
)

# =====================================================
# ================= IW28 ===============================
# =====================================================
with aba_iw28:
    st.title("Dashboard — AVP | Manutenção | Processo")
    st.caption("Fonte: SAP - IW28")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        f_local = st.multiselect("Código Macro", sorted(df_iw28["Local de instalação"].dropna().unique()))
    with c2:
        f_status = st.multiselect("Status", sorted(df_iw28["Status do sistema"].unique()))
    with c3:
        f_centro = st.multiselect("Centro de Trabalho", sorted(df_iw28["CenTrab.principal"].dropna().unique()))
    with c4:
        f_ord = st.multiselect("Campo de ordenação", sorted(df_iw28["Campo de ordenação"].dropna().unique()))

    df_f = df_iw28.copy()
    if f_local: df_f = df_f[df_f["Local de instalação"].isin(f_local)]
    if f_status: df_f = df_f[df_f["Status do sistema"].isin(f_status)]
    if f_centro: df_f = df_f[df_f["CenTrab.principal"].isin(f_centro)]
    if f_ord: df_f = df_f[df_f["Campo de ordenação"].isin(f_ord)]

    hoje = pd.Timestamp(datetime.today().date())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total de Ordens", len(df_f))
    k2.metric("Confirmadas", (df_f["Status do sistema"] == "Confirmada").sum())
    k3.metric("Em aberto", (df_f["Status do sistema"] == "Em aberto").sum())
    k4.metric("Atrasadas", ((df_f["Conclusão desejada"] < hoje) & (df_f["Status do sistema"] != "Confirmada")).sum())

    st.divider()

    aba_vis = st.radio(
        "Visualização",
        ["Por Status", "Por Código Macro", "Por Centro de Trabalho", "Por Campo de ordenação"],
        horizontal=True,
        label_visibility="collapsed"
    )

    if aba_vis == "Por Status":
        fig = px.histogram(df_f, x="Status do sistema", color="Status do sistema",
                           color_discrete_map=CORES_STATUS, text_auto=True)
        st.plotly_chart(fig, use_container_width=True)

    elif aba_vis == "Por Código Macro":
        base = df_f.groupby(["Local de instalação", "Status do sistema"]).size().reset_index(name="Qtd")
        fig = px.bar(base, x="Local de instalação", y="Qtd", color="Status do sistema",
                     color_discrete_map=CORES_STATUS, text="Qtd")
        st.plotly_chart(fig, use_container_width=True)

    elif aba_vis == "Por Centro de Trabalho":
        base = df_f.groupby(["CenTrab.principal", "Status do sistema"]).size().reset_index(name="Qtd")
        fig = px.bar(base, x="CenTrab.principal", y="Qtd", color="Status do sistema",
                     color_discrete_map=CORES_STATUS, text="Qtd")
        st.plotly_chart(fig, use_container_width=True)

    elif aba_vis == "Por Campo de ordenação":
        base = df_f.groupby(["Campo de ordenação", "Status do sistema"]).size().reset_index(name="Qtd")
        fig = px.bar(base, x="Campo de ordenação", y="Qtd", color="Status do sistema",
                     color_discrete_map=CORES_STATUS, text="Qtd")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Programação")
    st.dataframe(df_f, use_container_width=True)

# =====================================================
# ================= IW47 ===============================
# =====================================================
with aba_iw47:
    st.title("IW47 — Apropriação de Horas (Individual)")

    c1, c2, c3 = st.columns(3)
    with c1:
        f_colab = st.multiselect("Colaborador", sorted(df_iw47["Nome do empregado"].unique()))
    with c2:
        f_mes = st.selectbox(
            "Mês",
            ["Todos"] + sorted(df_iw47["Data de lançamento"].dt.to_period("M").astype(str).unique()),
            key="mes_iw47"
        )
    with c3:
        f_periodo = st.date_input(
            "Período",
            (df_iw47["Data de lançamento"].min().date(),
             df_iw47["Data de lançamento"].max().date())
        )

    df47 = df_iw47.copy()
    if f_colab:
        df47 = df47[df47["Nome do empregado"].isin(f_colab)]
    if f_mes != "Todos":
        df47 = df47[df47["Data de lançamento"].dt.to_period("M").astype(str) == f_mes]

    df47 = df47[
        (df47["Data de lançamento"].dt.date >= f_periodo[0]) &
        (df47["Data de lançamento"].dt.date <= f_periodo[1])
    ]

    resumo = df47.groupby("Nome do empregado", as_index=False)["Horas"].sum()

    fig = px.bar(resumo, x="Nome do empregado", y="Horas", text="Horas")
    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Detalhamento Diário")
    st.dataframe(
        df47[["Data de lançamento", "Nome do empregado", "Minutos", "Horas", "Texto de confirmação"]],
        use_container_width=True
    )

    PREVENTIVA_120H = [
        "ANDRE SAVIO NADAL", "CARLOS DANIEL ANTUNES", "DEJAIR JOSE SANTOS LIVRAMENTO",
        "JEAN WILLIAN SANTOS", "RODRIGO PINHEIRO", "NATALIA BARRETO"
    ]

    CORRETIVA_90H = [
        "EZEQUIEL ALEIXO", "RODRIGO CHARLES VIEIRA", "RONALDO CORREA DA ROCHA",
        "ALISSON PAULO GASTAO", "ROMARIO KASPCHAK", "THIAGO MAURICIO AZEVEDO",
        "CRISTIANO IATCZAKI", "EVANDRO LOPES SANTANA", "VICTOR EMANUEL PAES DE MELLO"
    ]

    def definir_meta(nome):
        if nome in PREVENTIVA_120H:
            return 120
        if nome in CORRETIVA_90H:
            return 90

    st.subheader("Indicador de Meta Mensal por Colaborador")

    if f_mes == "Todos":
        st.warning("Selecione um mês para validar a meta mensal.")
    else:
        resumo["Meta (h)"] = resumo["Nome do empregado"].apply(definir_meta)
        resumo["Ating (%)"] = (resumo["Horas"] / resumo["Meta (h)"] * 100).round(1)
        resumo["Status"] = resumo.apply(
            lambda x: "🟢 Atingida" if x["Horas"] >= x["Meta (h)"] else "🔴 Abaixo",
            axis=1
        )

        for _, row in resumo.iterrows():
            col_foto, col_info = st.columns([1, 6])
            with col_foto:
                foto = obter_foto(row["Nome do empregado"])
                if foto:
                    st.image(foto, width=70)
                else:
                    st.markdown("👤")
            with col_info:
                st.markdown(f"""
                <div style="background-color:#1c1f26;padding:12px;border-radius:10px;margin-bottom:8px;">
                    <b>{row['Nome do empregado']}</b><br>
                    Horas: <b>{row['Horas']:.1f}h</b> |
                    Meta: <b>{row['Meta (h)']}h</b> |
                    Atingimento: <b>{row['Ating (%)']}%</b><br>
                    <span style="font-size:18px">{row['Status']}</span>
                </div>
                """, unsafe_allow_html=True)

# =====================================================
# ================= SPLAN ==============================
# =====================================================
with aba_splan:
    st.title("Splan — Investigações por Colaborador")

    c1, c2, c3 = st.columns(3)
    with c1:
        f_colab = st.multiselect(
            "Criador da Investigação",
            sorted(df_splan["Criador da Investigação"].unique())
        )
    with c2:
        f_mes = st.selectbox(
            "Mês",
            ["Todos"] + sorted(df_splan["Mês"].unique()),
            key="mes_splan"
        )
    with c3:
        f_metodo = st.multiselect(
            "Método de Investigação",
            sorted(df_splan["Método de Investigação"].dropna().unique())
        )

    df_sp = df_splan.copy()
    if f_colab:
        df_sp = df_sp[df_sp["Criador da Investigação"].isin(f_colab)]
    if f_mes != "Todos":
        df_sp = df_sp[df_sp["Mês"] == f_mes]
    if f_metodo:
        df_sp = df_sp[df_sp["Método de Investigação"].isin(f_metodo)]

    resumo = df_sp.groupby("Criador da Investigação").size().reset_index(name="Qtd Investigações")

    fig = px.bar(resumo, x="Criador da Investigação", y="Qtd Investigações", text="Qtd Investigações")
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Detalhamento")
    st.dataframe(
        df_sp[["Data da Investigação", "Criador da Investigação", "Método de Investigação"]],
        use_container_width=True
    )
