import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import shutil
import subprocess
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
from streamlit_paste_button import paste_image_button
from PIL import Image
import platform
import time
import calendar
import json
from pathlib import Path

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.7.12", layout="wide")

# --- CONSTANTES DO CONTRATO ---
META_DIARIA_CONTRATO = 250

# --- CUSTOM CSS (ORIGINAL) ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f5; }
    .dashboard-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        border-left: 5px solid #28a745;
    }
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
        border-radius: 8px !important;
    }
    div.stButton > button[key*="del_"] {
        border: 1px solid #dc3545 !important;
        color: #dc3545 !important;
        background-color: transparent !important;
        font-size: 0.8em !important;
        height: 2em !important;
    }
    /* Estilo para o botão excluir na gerência */
    div.stButton > button[key="btn_excluir_relatorio"] {
        border: 1px solid #dc3545 !important;
        color: #dc3545 !important;
        background-color: transparent !important;
        height: 3em !important;
        width: 100% !important;
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "IMAGEM_PRINT_ATENDIMENTO": 165, "PRINT_CLASSIFICACAO": 160,
    "IMAGEM_DOCUMENTO_RAIO_X": 165, "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 160, "TABELA_OBITO": 180,
    "TABELA_CCIH": 160, "IMAGEM_NEP": 160,
    "IMAGEM_TREINAMENTO_INTERNO": 160, "IMAGEM_MELHORIAS": 160,
    "GRAFICO_OUVIDORIA": 155, "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 160
}

# --- DIRETÓRIO DE RELATÓRIOS SALVOS ---
BASE_RELATORIOS_DIR = Path("relatorios_salvos")
BASE_RELATORIOS_DIR.mkdir(exist_ok=True)

# --- CHAVES DE CAMPOS QUE SERÃO PERSISTIDAS ---
FORM_KEYS = [
    "sel_mes", "sel_ano",
    "in_total", "in_rx", "in_mc", "in_mp",
    "in_oc", "in_op", "in_ccih",
    "in_oi", "in_oe", "in_taxa",
    "in_tt", "in_to", "in_to_menor", "in_to_maior"
]

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

if 'relatorio_atual' not in st.session_state:
    st.session_state.relatorio_atual = ""

# --- FUNÇÕES DE PERSISTÊNCIA ---

def _normalizar_nome_relatorio(nome: str) -> str:
    nome = nome.strip()
    for ch in r'\/:*?"<>|': nome = nome.replace(ch, "_")
    return nome or "relatorio_sem_nome"

def _caminho_relatorio(nome_normalizado: str) -> Path:
    return BASE_RELATORIOS_DIR / nome_normalizado

def listar_relatorios_salvos():
    return sorted([p.name for p in BASE_RELATORIOS_DIR.iterdir() if p.is_dir()])

def salvar_relatorio(nome_relatorio: str):
    if not nome_relatorio: return
    nome_norm = _normalizar_nome_relatorio(nome_relatorio)
    pasta_rel = _caminho_relatorio(nome_norm)
    pasta_rel.mkdir(parents=True, exist_ok=True)
    pasta_evid = pasta_rel / "evidencias"
    pasta_evid.mkdir(exist_ok=True)

    evidencias_meta = {}
    for marcador, itens in st.session_state.dados_sessao.items():
        evidencias_meta[marcador] = []
        for idx, item in enumerate(itens):
            _, ext = os.path.splitext(item["name"])
            ext = ext.lower() or ".png"
            nome_arq = f"{marcador}_{idx}{ext}"
            caminho_dest = pasta_evid / nome_arq
            conteudo = item["content"]
            if isinstance(conteudo, Image.Image):
                buf = io.BytesIO(); conteudo.save(buf, format="PNG"); data = buf.getvalue()
            else:
                if hasattr(conteudo, "getvalue"): data = conteudo.getvalue()
                elif hasattr(conteudo, "read"): data = conteudo.read()
                else: data = conteudo
            with open(caminho_dest, "wb") as f: f.write(data)
            evidencias_meta[marcador].append({"name": item["name"], "file": f"evidencias/{nome_arq}", "type": item["type"]})

    estado = {"form_state": {k: st.session_state.get(k) for k in FORM_KEYS}, "evidencias": evidencias_meta}
    with open(pasta_rel / "estado.json", "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)
    st.session_state.relatorio_atual = nome_norm
    st.success(f"Relatório '{nome_relatorio}' salvo.")

def carregar_relatorio(nome_relatorio: str):
    nome_norm = _normalizar_nome_relatorio(nome_relatorio)
    pasta_rel = _caminho_relatorio(nome_norm)
    estado_path = pasta_rel / "estado.json"
    if not estado_path.exists(): return
    with open(estado_path, "r", encoding="utf-8") as f: estado = json.load(f)
    for k, v in estado.get("form_state", {}).items(): st.session_state[k] = v
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
    for m, lista in estado.get("evidencias", {}).items():
        for meta in lista:
            p = pasta_rel / meta["file"]
            if p.exists():
                with open(p, "rb") as f: bio = io.BytesIO(f.read()); bio.name = meta["name"]
                st.session_state.dados_sessao[m].append({"name": meta["name"], "content": bio, "type": meta["type"]})
    st.session_state.relatorio_atual = nome_norm
    st.success(f"Relatório '{nome_relatorio}' carregado.")

def excluir_relatorio(nome_relatorio: str):
    nome_norm = _normalizar_nome_relatorio(nome_relatorio)
    pasta_rel = _caminho_relatorio(nome_norm)
    if pasta_rel.exists():
        shutil.rmtree(pasta_rel)
        if st.session_state.relatorio_atual == nome_norm: st.session_state.relatorio_atual = ""
        st.success(f"Relatório '{nome_relatorio}' excluído.")

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3208/3208726.png", width=100)
    st.title("Painel de Controlo")
    st.markdown("---")
    total_anexos = sum(len(v) for v in st.session_state.dados_sessao.values())
    st.metric("Total de Anexos", total_anexos)
    if st.button("🗑 Limpar Todos os Dados", key="btn_limpar_tudo"):
        st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
        st.rerun()

# --- UI PRINCIPAL ---
st.title("Automação de Relatórios V0.7.12")

# --- GERENCIAMENTO DE RELATÓRIOS (LAYOUT PRESERVADO) ---
with st.container(border=True):
    st.markdown("#### Gerenciamento de Relatórios")
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        rel_existentes = listar_relatorios_salvos()
        opcao_rel = st.selectbox("Relatórios salvos", ["(Novo relatório)"] + rel_existentes, index=0)

    with col2:
        nome_input = st.text_input("Nome do relatório", value=st.session_state.relatorio_atual or "")

    with col3:
        if st.button("Carregar", key="btn_carregar"):
            if opcao_rel != "(Novo relatório)": carregar_relatorio(opcao_rel); st.rerun()
        
        if st.button("Salvar", key="btn_salvar"):
            n = nome_input or (opcao_rel if opcao_rel != "(Novo relatório)" else None)
            if n: salvar_relatorio(n); st.rerun()

        if opcao_rel != "(Novo relatório)":
            if st.button("Excluir", key="btn_excluir_relatorio"):
                excluir_relatorio(opcao_rel); st.rerun()

# --- TABS DE DADOS (CONFORME SUA NECESSIDADE) ---
t_dados, t_arquivos = st.tabs(["📊 Dados Mensais", "📁 Arquivos"])

with t_dados:
    with st.container(border=True):
        st.markdown("### Período e Produção")
        c1, c2, c3 = st.columns(3)
        with c1: st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], key="sel_mes")
        with c2: st.selectbox("Ano", [2025, 2026, 2027], key="sel_ano")
        with c3: st.number_input("Total Atendimentos", key="in_total", step=1)

    with st.container(border=True):
        st.markdown("### Hospitalar")
        c1, c2, c3 = st.columns(3)
        with c1: st.number_input("Clínica Médica", key="in_mc", step=1)
        with c2: st.number_input("Pediatria", key="in_mp", step=1)
        with c3: st.number_input("Raio-X", key="in_rx", step=1)

with t_arquivos:
    st.info("Utilize esta aba para anexar as evidências conforme os marcadores do sistema.")
