import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
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

# --- CUSTOM CSS ---
st.markdown("""
    &lt;style&gt;
    .main { background-color: #f0f2f5; }
    .dashboard-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        border-left: 5px solid #28a745;
    }
    div.stButton &gt; button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
        border-radius: 8px !important;
    }
    div.stButton &gt; button[key*="del_"] {
        border: 1px solid #dc3545 !important;
        color: #dc3545 !important;
        background-color: transparent !important;
        font-size: 0.8em !important;
        height: 2em !important;
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    &lt;/style&gt;
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
    "in_total",
    "in_rx", "in_mc", "in_mp",
    "in_oc", "in_op", "in_ccih",
    "in_oi", "in_oe", "in_taxa",
    "in_tt", "in_to", "in_to_menor", "in_to_maior"
]

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

if 'relatorio_atual' not in st.session_state:
    st.session_state.relatorio_atual = ""

# Valores padrão iniciais para período de referência
if "sel_mes" not in st.session_state:
    st.session_state.sel_mes = "Janeiro"
if "sel_ano" not in st.session_state:
    st.session_state.sel_ano = 2026

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

