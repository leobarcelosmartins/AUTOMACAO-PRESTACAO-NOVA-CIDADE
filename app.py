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

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.4.3", layout="wide")

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .dashboard-section {
        background-color: #ffffff;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        margin-bottom: 30px;
        border: 1px solid #f0f0f0;
    }
    .upload-label {
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 10px;
        font-size: 1rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "EXCEL_META_ATENDIMENTOS": 165, "IMAGEM_PRINT_ATENDIMENTO": 160,
    "PRINT_CLASSIFICACAO": 155, "IMAGEM_DOCUMENTO_RAIO_X": 150, 
    "TABELA_TRANSFERENCIA": 120, "GRAFICO_TRANSFERENCIA": 155,
    "TABELA_TOTAL_OBITO": 150, "TABELA_OBITO": 150, 
    "TABELA_CCIH": 150, "TABELA_QUALITATIVA_IMG": 155,
    "IMAGEM_NEP": 165, "IMAGEM_TREINAMENTO_INTERNO": 165, 
    "IMAGEM_MELHORIAS": 165, "GRAFICO_OUVIDORIA": 155, 
    "PDF_OUVIDORIA_INTERNA": 165
}

# --- ESTADO DA SESSÃO ---
if 'arquivos_por_marcador' not in st.session_state:
    st.session_state.arquivos_por_marcador = {m: [] for m in DIMENSOES_CAMPOS.keys()}

def excel_para_imagem(doc_template, arquivo_excel):
    try:
        df = pd.read_excel(arquivo_excel, sheet_name="TRANSFERENCIAS", usecols=[3, 4], skiprows=2, nrows=14, header=None)
        df = df.fillna('')
        def format_inteiro(val):
            if val == '' or val is None: return ''
            try: return str(int(float(val)))
            except: return str(val)
        if df.shape[1] > 1:
            df.iloc[:, 1] = df.iloc[:, 1].apply(format_inteiro)
        
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('off')
        tabela = ax.table(cellText=df.values, loc='center', cellLoc='center', colWidths=[0.45, 0.45])
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(11)
        tabela.scale(1.2, 1.8)
        
        for (row, col), cell in tabela.get_celld().items():
            cell.get_text().set_weight('bold')
            cell.set_edgecolor('#000000')
            cell.set_linewidth(1)
            if row == 0:
                cell.set_facecolor('#D3D3D3')
                if col == 1: cell.get_text().set_text('')
        
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        img_buf.seek(0)
        return InlineImage(doc_template, img_buf, width=Mm(DIMENSOES_CAMPOS["TABELA_TRANSFERENCIA"]))
    except Exception as e:
        st.error(f"Erro Excel: {e}")
        return None

def processar_item(doc_template, item, marcador):
    largura_mm = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        # Se for bytes (preview armazenado no state para prints colados)
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura_mm))]
        
        # Se for imagem PIL direta
        if isinstance(item, Image.Image):
            buf = io.BytesIO()
            item.save(buf, format="PNG")
            buf.seek(0)
            return [InlineImage(doc_template, buf, width=Mm(largura_mm))]

        # Se for upload de arquivo
        extensao = getattr(item, 'name', '').lower()
        if marcador == "TABELA_TRANSFERENCIA" and (extensao.endswith(".xlsx") or extensao.endswith(".xls")):
            res = excel_para_imagem(doc_template, item)
            return [res] if res else []
        
        if extensao.endswith(".pdf"):
            pdf_doc = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf_doc:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                buf = io.BytesIO(pix.tobytes())
                imgs.append(InlineImage(doc_template, buf, width=Mm(largura_mm)))
            pdf_doc.close()
            return imgs
        
        return [InlineImage(doc_template, item, width=Mm(largura_mm))]
    except Exception as e:
        st.error(f"Erro no marcador {marcador}: {e}")
        return []

def gerar_pdf(docx_path, output_dir):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True, capture_output=True)
        return os.path.join(output_dir, os.path.basename(docx_path).replace('.docx', '.pdf'))
    except:
        return None

# --- UI ---
st.title("Automação de Relatórios Assistenciais")
st.caption("Versão 0.4.3 - Estabilidade Master")

tab_manual, tab_arquivos = st.tabs(["📝 Dados Manuais", "📁 Gestão de Evidências"])
ctx_manual = {}

with tab_manual:
    with st.container():
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        ctx_manual["SISTEMA_MES_REFERENCIA"] = st.text_input("Mês de Referência (Ex: Janeiro/2026)")
        c1, c2 = st.columns(2)
        ctx_manual["ANALISTA_TOTAL_ATENDIMENTOS"] = c1.text_input("Total de Atendimentos")
        ctx_manual["TOTAL_RAIO_X"] = c2.text_input("Total Raio-X")
        c3, c4, c5 = st.columns(3)
        ctx_manual["ANALISTA_MEDICO_CLINICO"] = c3.text_input("Médicos Clínicos")
        ctx_manual["ANALISTA_MEDICO_PEDIATRA"] = c4.text_input("Médicos Pediatras")
        ctx_manual["ANALISTA_ODONTO_CLINICO"] = c5.text_input("Odonto Clínico")
        c6, c7, c8 = st.columns(3)
        ctx_manual["ANALISTA_ODONTO_PED"] = c6.text_input("Odonto Ped")
        ctx_manual["TOTAL_PACIENTES_CCIH"] = c7.text_input("Pacientes CCIH")
        ctx_manual["OUVIDORIA_INTERNA"] = c8.text_input("Ouvidoria Interna")
        c9, c10, c11 = st.columns(3)
        ctx_manual["OUVIDORIA_EXTERNA"] = c9.text_input("Ouvidoria Externa")
        ctx_manual["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c10.number_input("Total de Transferências", step=1, value=0)
        ctx_manual["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c11.text_input("Taxa de Transferência (%)", value="0,00%")
        st.markdown('</div>', unsafe_allow_html=True)

with tab_arquivos:
    labels = {
        "EXCEL_META_ATENDIMENTOS": "Grade de Metas", "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento", 
        "PRINT_CLASSIFICACAO": "Classificação de Risco", "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X", 
        "TABELA_TRANSFERENCIA": "Tabela Transferência (Excel)", "GRAFICO_TRANSFERENCIA": "Gráfico Transferência",
        "TABELA_TOTAL_OBITO": "Tab. Total Óbito", "TABELA_OBITO": "Tab. Óbito", 
        "TABELA_CCIH": "Tabela CCIH", "TABELA_QUALITATIVA_IMG": "Tab. Qualitativa",
        "IMAGEM_NEP": "Imagens NEP", "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno", 
        "IMAGEM_MELHORIAS": "Melhorias", "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria", 
        "PDF_OUVIDORIA_INTERNA": "Relatório Ouvidoria"
    }
    blocos = [
        ["EXCEL_META_ATENDIMENTOS", "IMAGEM_PRINT_ATENDIMENTO", "PRINT_CLASSIFICACAO", "IMAGEM_DOCUMENTO_RAIO_X"],
        ["TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA"],
        ["TABELA_TOTAL_OBITO", "TABELA_OBITO", "TABELA_CCIH", "TABELA_QUALITATIVA_IMG"],
        ["IMAGEM_NEP", "IMAGEM_TREINAMENTO_INTERNO", "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA"]
    ]

    for b_idx, lista_m in enumerate(blocos):
        st.markdown('<div class="dashboard-section">', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        for idx, m in enumerate(lista_m):
            target_col = col1 if idx % 2 == 0 else col2
            with target_col:
                with st.container(border=True):
                    st.markdown(f"<div class='upload-label'>{labels.get(m, m)}</div>", unsafe_allow_html=True)
                    c_act1, c_act2 = st.columns([1, 1.2])
                    with c_act1:
                        pasted = paste_image_button(label="Colar print", key=f"p_{m}_{b_idx}")
                        if pasted:
                            # Tratamento robusto para extrair bytes da imagem colada
                            buf = io.BytesIO()
                            pasted.save(buf, format="PNG")
                            img_bytes = buf.getvalue()
                            nome_p = f"Captura_{len(st.session_state.arquivos_por_marcador[m]) + 1}.png"
                            st.session_state.arquivos_por_marcador[m].append({
                                "name": nome_p, "content": img_bytes, "preview": img_bytes, "type": "print"
                            })
                            st.rerun()
                    with c_act2:
                        tipo_f = ['png', 'jpg', 'pdf', 'xlsx', 'xls'] if m == "TABELA_TRANSFERENCIA" else ['png', 'jpg', 'pdf']
                        f_up = st.file_uploader("Upload", type=tipo_f, key=f"f_{m}_{b_idx}", accept_multiple_files=True, label_visibility="collapsed")
                        if f_up:
                            for f in f_up:
                                if f.name not in [x["name"] for x in st.session_state.arquivos_por_marcador[m]]:
                                    st.session_state.arquivos_por_marcador[m].append({
                                        "name": f.name, "content": f, "preview": f if not f.name.lower().endswith(('.pdf', '.xlsx', '.xls')) else None, "type": "file"
                                    })
                            st.rerun()

                    if st.session_state.arquivos_por_marcador[m]:
                        for i_idx, item in enumerate(st.session_state.arquivos_por_marcador[m]):
                            with st.expander(f"📄 {item['name']}"):
                                if item['preview']: st.image(item['preview'], use_container_width=True)
                                if st.button("Remover", key=f"del_{m}_{i_idx}_{b_idx}", use_container_width=True):
                                    st.session_state.arquivos_por_marcador[m].pop(i_idx)
                                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if st.button("🚀 FINALIZAR E GERAR RELATÓRIO PDF", use_container_width=True):
    if not ctx_manual.get("SISTEMA_MES_REFERENCIA"):
        st.error("Mês de Referência é obrigatório.")
    else:
        try:
            try:
                mc = int(ctx_manual.get("ANALISTA_MEDICO_CLINICO") or 0)
                mp = int(ctx_manual.get("ANALISTA_MEDICO_PEDIATRA") or 0)
                ctx_manual["SISTEMA_TOTAL_MEDICOS"] = mc + mp
            except:
                ctx_manual["SISTEMA_TOTAL_MEDICOS"] = 0

            with tempfile.TemporaryDirectory() as tmp:
                docx_path = os.path.join(tmp, "temp.docx")
                doc = DocxTemplate("template.docx")
                with st.spinner("Gerando documento..."):
                    dados_finais = ctx_manual.copy()
                    for m in DIMENSOES_CAMPOS.keys():
                        imgs = []
                        for item in st.session_state.arquivos_por_marcador[m]:
                            res = processar_item(doc, item['content'], m)
                            if res: imgs.extend(res)
                        dados_finais[m] = imgs
                doc.render(dados_finais)
                doc.save(docx_path)
                pdf_res = gerar_pdf(docx_path, tmp)
                if pdf_res:
                    with open(pdf_res, "rb") as f:
                        st.success("Relatório gerado com sucesso.")
                        st.download_button("📥 Baixar Relatório PDF", f.read(), f"Relatorio_{ctx_manual['SISTEMA_MES_REFERENCIA']}.pdf", "application/pdf")
        except Exception as e:
            st.error(f"Erro Crítico: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins | Backup Tático")
