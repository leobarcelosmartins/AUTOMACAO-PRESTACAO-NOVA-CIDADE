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

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.6.7", layout="wide")

# --- CUSTOM CSS PARA DASHBOARD ---
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
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES (Conforme Referência V0.6.4) ---
DIMENSOES_CAMPOS = {
    "EXCEL_META_ATENDIMENTOS": 165, "IMAGEM_PRINT_ATENDIMENTO": 165,
    "IMAGEM_DOCUMENTO_RAIO_X": 165, "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 160, "TABELA_TOTAL_OBITO": 165,
    "TABELA_OBITO": 180, "TABELA_CCIH": 180, "IMAGEM_NEP": 160,
    "IMAGEM_TREINAMENTO_INTERNO": 160, "IMAGEM_MELHORIAS": 160,
    "GRAFICO_OUVIDORIA": 155, "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 160, "PRINT_CLASSIFICACAO": 160
}

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

def excel_para_imagem(doc_template, arquivo_excel):
    try:
        if hasattr(arquivo_excel, 'seek'): arquivo_excel.seek(0)
        df = pd.read_excel(arquivo_excel, sheet_name="TRANSFERENCIAS", usecols=[3, 4], skiprows=2, nrows=14, header=None)
        df = df.fillna('')
        def fmt(v):
            try: return str(int(float(v)))
            except: return str(v)
        if df.shape[1] > 1: df.iloc[:, 1] = df.iloc[:, 1].apply(fmt)
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('off')
        tabela = ax.table(cellText=df.values, loc='center', cellLoc='center', colWidths=[0.45, 0.45])
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(11)
        tabela.scale(1.2, 1.8)
        for (r, c), cell in tabela.get_celld().items():
            cell.get_text().set_weight('bold')
            if r == 0: cell.set_facecolor('#D3D3D3')
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        img_buf.seek(0)
        return InlineImage(doc_template, img_buf, width=Mm(DIMENSOES_CAMPOS["TABELA_TRANSFERENCIA"]))
    except Exception as e:
        st.error(f"Erro Excel: {e}")
        return None

def processar_item_lista(doc_template, item, marcador):
    largura = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        if hasattr(item, 'seek'): item.seek(0)
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura))]
        ext = getattr(item, 'name', '').lower()
        if marcador == "TABELA_TRANSFERENCIA" and (ext.endswith(".xlsx") or ext.endswith(".xls")):
            res = excel_para_imagem(doc_template, item)
            return [res] if res else []
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close()
            return imgs
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except Exception: return []

def converter_para_pdf(docx_path, output_dir):
    """Auxiliar para ambiente localhost encontrar o LibreOffice."""
    comando = 'libreoffice'
    if platform.system() == "Windows":
        caminhos = [
            'libreoffice',
            r'C:\Program Files\LibreOffice\program\soffice.exe',
            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
        ]
        for p in caminhos:
            try:
                subprocess.run([p, '--version'], capture_output=True, check=True)
                comando = p
                break
            except: continue
    subprocess.run([comando, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True)

# --- UI ---
st.title("Automação de Relatórios - UPA Nova Cidade")
st.caption("Versão 0.6.7 - Ajuste de Layout e Campo Raio-X")

t_manual, t_evidencia = st.tabs(["📝 Dados", "📁 Evidências"])
ctx_manual = {}

with t_manual:
    st.markdown("### Preencha os campos de texto")
    
    # Layout em 3 colunas por linha
    c1, c2, c3 = st.columns(3)
    ctx_manual["SISTEMA_MES_REFERENCIA"] = c1.text_input("Mês de Referência", key="in_mes")
    ctx_manual["ANALISTA_TOTAL_ATENDIMENTOS"] = c2.text_input("Total de Atendimentos", key="in_total")
    ctx_manual["TOTAL_RAIO_X"] = c3.text_input("Total Raio-X", key="in_rx")
    
    c4, c5, c6 = st.columns(3)
    ctx_manual["ANALISTA_MEDICO_CLINICO"] = c4.text_input("Médicos Clínicos", key="in_mc")
    ctx_manual["ANALISTA_MEDICO_PEDIATRA"] = c5.text_input("Médicos Pediatras", key="in_mp")
    ctx_manual["ANALISTA_ODONTO_CLINICO"] = c6.text_input("Odonto Clínico", key="in_oc")
    
    c7, c8, c9 = st.columns(3)
    ctx_manual["ANALISTA_ODONTO_PED"] = c7.text_input("Odonto Ped", key="in_op")
    ctx_manual["TOTAL_PACIENTES_CCIH"] = c8.text_input("Pacientes CCIH", key="in_ccih")
    ctx_manual["OUVIDORIA_INTERNA"] = c9.text_input("Ouvidoria Interna", key="in_oi")
    
    c10, c11, c12 = st.columns(3)
    ctx_manual["OUVIDORIA_EXTERNA"] = c10.text_input("Ouvidoria Externa", key="in_oe")
    ctx_manual["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c11.number_input("Total de Transferências", step=1, key="in_tt")
    ctx_manual["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c12.text_input("Taxa de Transferência (%)", key="in_taxa")

with t_evidencia:
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
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        col_esq, col_dir = st.columns(2)
        for idx, m in enumerate(lista_m):
            target = col_esq if idx % 2 == 0 else col_dir
            with target:
                st.markdown(f"<span class='upload-label'>{labels.get(m, m)}</span>", unsafe_allow_html=True)
                ca, cb = st.columns([1, 1])
                with ca:
                    # Lógica de Captura restaurada conforme V0.6.4
                    key_p = f"p_{m}_{len(st.session_state.dados_sessao[m])}"
                    pasted = paste_image_button(label="Colar Print", key=key_p)
                    
                    if pasted is not None and pasted.image_data is not None:
                        try:
                            img_pil = pasted.image_data
                            buf = io.BytesIO()
                            img_pil.save(buf, format="PNG")
                            b_data = buf.getvalue()
                            nome = f"Captura_{len(st.session_state.dados_sessao[m]) + 1}.png"
                            st.session_state.dados_sessao[m].append({"name": nome, "content": b_data, "type": "p"})
                            st.rerun()
                        except: pass

                with cb:
                    f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf', 'xlsx'], key=f"f_{m}_{b_idx}", label_visibility="collapsed")
                    if f_up:
                        if f_up.name not in [x['name'] for x in st.session_state.dados_sessao[m]]:
                            st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                            st.rerun()

                # Lógica de Listagem restaurada conforme V0.6.4
                if st.session_state.dados_sessao[m]:
                    for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                        with st.expander(f"📄 {item['name']}", expanded=False):
                            if item['type'] == "p" or not item['name'].lower().endswith(('.pdf', '.xlsx')):
                                st.image(item['content'], use_container_width=True)
                            if st.button("Remover", key=f"del_{m}_{i_idx}_{b_idx}"):
                                st.session_state.dados_sessao[m].pop(i_idx)
                                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if st.button("🚀 FINALIZAR E GERAR RELATÓRIO", type="primary", use_container_width=True):
    if not ctx_manual.get("SISTEMA_MES_REFERENCIA"):
        st.error("Mês de Referência é obrigatório.")
    else:
        try:
            # Cálculos automáticos
            mc = int(ctx_manual.get("ANALISTA_MEDICO_CLINICO") or 0)
            mp = int(ctx_manual.get("ANALISTA_MEDICO_PEDIATRA") or 0)
            ctx_manual["SISTEMA_TOTAL_MEDICOS"] = mc + mp

            with tempfile.TemporaryDirectory() as tmp:
                docx_p = os.path.join(tmp, "relatorio.docx")
                doc = DocxTemplate("template.docx")
                with st.spinner("Construindo relatório..."):
                    for m in DIMENSOES_CAMPOS.keys():
                        imgs_doc = []
                        for item in st.session_state.dados_sessao[m]:
                            res = processar_item_lista(doc, item['content'], m)
                            if res: imgs_doc.extend(res)
                        ctx_manual[m] = imgs_doc
                    
                    doc.render(ctx_manual)
                    doc.save(docx_p)
                    
                    st.success("✅ Relatório renderizado!")
                    c_down1, c_down2 = st.columns(2)
                    
                    with c_down1:
                        with open(docx_p, "rb") as f_w:
                            st.download_button(
                                label="📥 Baixar WORD (.docx)",
                                data=f_w.read(),
                                file_name=f"Relatorio_{ctx_manual['SISTEMA_MES_REFERENCIA']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    
                    with c_down2:
                        try:
                            converter_para_pdf(docx_p, tmp)
                            pdf_p = os.path.join(tmp, "relatorio.pdf")
                            if os.path.exists(pdf_p):
                                with open(pdf_p, "rb") as f_p:
                                    st.download_button(
                                        label="📥 Baixar PDF",
                                        data=f_p.read(),
                                        file_name=f"Relatorio_{ctx_manual['SISTEMA_MES_REFERENCIA']}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                        except Exception:
                            st.warning("⚠️ LibreOffice não detectado localmente. Baixe o Word para conversão manual.")
        except Exception as e: st.error(f"Erro Crítico: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins | Backup Tático")
