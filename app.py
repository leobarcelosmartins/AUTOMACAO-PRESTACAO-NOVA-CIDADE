import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import subprocess
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
from streamlit_paste_button import paste_image_button

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.4.2", layout="wide")

# --- DICIONÁRIO DE DIMENSÕES POR CAMPO (LARGURAS EM MM) ---
# Definimos larguras específicas para cada marcador para otimizar o layout
DIMENSOES_CAMPOS = {
    "EXCEL_META_ATENDIMENTOS": 165,
    "IMAGEM_PRINT_ATENDIMENTO": 160,
    "IMAGEM_DOCUMENTO_RAIO_X": 150,
    "TABELA_TRANSFERENCIA": 120,   # Tabela Excel mais estreita para evitar quebra
    "GRAFICO_TRANSFERENCIA": 155,
    "TABELA_TOTAL_OBITO": 150,
    "TABELA_OBITO": 150,
    "TABELA_CCIH": 150,
    "IMAGEM_NEP": 165,
    "IMAGEM_TREINAMENTO_INTERNO": 165,
    "IMAGEM_MELHORIAS": 165,
    "GRAFICO_OUVIDORIA": 155,
    "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 155,
    "PRINT_CLASSIFICACAO": 155
}

def excel_para_imagem(doc_template, arquivo_excel):
    """
    Extrai o intervalo D3:E16 da aba TRANSFERENCIAS com formatação profissional.
    """
    try:
        df = pd.read_excel(
            arquivo_excel, 
            sheet_name="TRANSFERENCIAS", 
            usecols=[3, 4], 
            skiprows=2, 
            nrows=14, 
            header=None
        )
        
        df = df.fillna('')
        
        def format_inteiro(val):
            if val == '' or val is None: return ''
            try:
                return str(int(float(val)))
            except:
                return str(val)
        
        # Formatação da segunda coluna para inteiros puros
        if df.shape[1] > 1:
            df.iloc[:, 1] = df.iloc[:, 1].apply(format_inteiro)
        
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('off')
        
        tabela = ax.table(
            cellText=df.values, 
            loc='center', 
            cellLoc='center',
            colWidths=[0.45, 0.45]
        )
        
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(11)
        tabela.scale(1.2, 1.8)
        
        for (row, col), cell in tabela.get_celld().items():
            cell.get_text().set_weight('bold')
            cell.set_edgecolor('#000000')
            cell.set_linewidth(1)
            
            if row == 0:
                cell.set_facecolor('#D3D3D3')
                if col == 1:
                    cell.get_text().set_text('')
                if col == 0:
                    cell.get_text().set_position((0.5, 0.5))

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        img_buf.seek(0)
        
        # Usa a largura específica definida no dicionário
        largura_mm = DIMENSOES_CAMPOS.get("TABELA_TRANSFERENCIA", 120)
        return [InlineImage(doc_template, img_buf, width=Mm(largura_mm))]
    except Exception as e:
        st.error(f"Erro no processamento da tabela Excel: {e}")
        return []

def processar_conteudo(doc_template, conteudo, marcador=None):
    """Processa ficheiros, PDFs ou imagens coladas do clipboard com larguras dinâmicas."""
    if not conteudo:
        return []
    
    imagens = []
    # Obtém a largura definida para este marcador ou usa 165mm como padrão
    largura_mm = DIMENSOES_CAMPOS.get(marcador, 165)
    
    try:
        # Se for imagem colada (objeto PIL Image vindo do streamlit-paste-button)
        if hasattr(conteudo, 'save') and not hasattr(conteudo, 'name'):
            img_byte_arr = io.BytesIO()
            conteudo.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            imagens.append(InlineImage(doc_template, img_byte_arr, width=Mm(largura_mm)))
            return imagens

        # Se for ficheiro carregado
        extensao = getattr(conteudo, 'name', '').lower()
        
        if marcador == "TABELA_TRANSFERENCIA" and (extensao.endswith(".xlsx") or extensao.endswith(".xls")):
            return excel_para_imagem(doc_template, conteudo)

        if extensao.endswith(".pdf"):
            pdf_doc = fitz.open(stream=conteudo.read(), filetype="pdf")
            for pagina in pdf_doc:
                pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_byte_arr = io.BytesIO(pix.tobytes())
                imagens.append(InlineImage(doc_template, img_byte_arr, width=Mm(largura_mm)))
            pdf_doc.close()
            return imagens

        imagens.append(InlineImage(doc_template, conteudo, width=Mm(largura_mm)))
        return imagens
    except Exception as e:
        st.error(f"Erro no marcador {marcador}: {e}")
        return []

def gerar_pdf(docx_path, output_dir):
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path],
            check=True, capture_output=True
        )
        return os.path.join(output_dir, os.path.basename(docx_path).replace('.docx', '.pdf'))
    except:
        return None

# --- UI ---
st.title("Automação de Relatório de Prestação - UPA Nova Cidade")
st.caption("Versão 0.4.2 - Estabilização de Dependências")

col_t1 = ["SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO", "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO"]
col_t2 = ["ANALISTA_ODONTO_PED", "TOTAL_RAIO_X", "TOTAL_PACIENTES_CCIH", "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"]

marcadores = {
    "EXCEL_META_ATENDIMENTOS": "Grade de Metas",
    "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento",
    "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X",
    "TABELA_TRANSFERENCIA": "Tabela Transferência",
    "GRAFICO_TRANSFERENCIA": "Gráfico Transferência",
    "TABELA_TOTAL_OBITO": "Tabela Total Óbito",
    "TABELA_OBITO": "Tabela Óbito",
    "TABELA_CCIH": "Tabela CCIH",
    "IMAGEM_NEP": "Imagens NEP",
    "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno",
    "IMAGEM_MELHORIAS": "Imagens de Melhorias",
    "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria",
    "PDF_OUVIDORIA_INTERNA": "Relatório Ouvidoria (PDF)",
    "TABELA_QUALITATIVA_IMG": "Tabela Qualitativa",
    "PRINT_CLASSIFICACAO": "Classificação de Risco"
}

if 'pasted_files' not in st.session_state:
    st.session_state.pasted_files = {}

with st.form("form_v4_2"):
    t1, t2 = st.tabs(["Dados Manuais", "Arquivos"])
    ctx = {}
    
    with t1:
        c1, c2 = st.columns(2)
        for f in col_t1: ctx[f] = c1.text_input(f.replace("_", " "))
        for f in col_t2: ctx[f] = c2.text_input(f.replace("_", " "))
        st.write("---")
        st.subheader("Indicadores de Transferência")
        c3, c4 = st.columns(2)
        ctx["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c3.number_input("Total de Transferências", step=1, value=0)
        ctx["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c4.text_input("Taxa de Transferência (Ex: 0,76%)", value="0,00%")

    with t2:
        uploads = {}
        u1, u2 = st.columns(2)
        for i, (m, label) in enumerate(marcadores.items()):
            col = u1 if i % 2 == 0 else u2
            with col:
                st.write(f"**{label}**")
                pasted = paste_image_button(label=f"Colar print", key=f"p_{m}")
                if pasted:
                    st.session_state.pasted_files[m] = pasted.image_data
                
                # Mensagem de confirmação de recebimento do print
                if m in st.session_state.pasted_files:
                    st.info("Print recebido.")
                
                tipos = ['png', 'jpg', 'pdf', 'xlsx', 'xls'] if m == "TABELA_TRANSFERENCIA" else ['png', 'jpg', 'pdf']
                uploads[m] = st.file_uploader("Ou ficheiro", type=tipos, key=f"f_{m}", label_visibility="collapsed")

    btn = st.form_submit_button("GERAR RELATÓRIO PDF FINAL")

if btn:
    if not ctx["SISTEMA_MES_REFERENCIA"]:
        st.error("Mês de Referência é obrigatório.")
    else:
        try:
            m_c = int(ctx.get("ANALISTA_MEDICO_CLINICO") or 0)
            m_p = int(ctx.get("ANALISTA_MEDICO_PEDIATRA") or 0)
            ctx["SISTEMA_TOTAL_MEDICOS"] = m_c + m_p

            with tempfile.TemporaryDirectory() as tmp:
                docx_p = os.path.join(tmp, "temp.docx")
                doc = DocxTemplate("template.docx")

                with st.spinner("Processando dados e arquivos..."):
                    for m in marcadores.keys():
                        raw = uploads.get(m) or st.session_state.pasted_files.get(m)
                        ctx[m] = processar_conteudo(doc, raw, m)

                doc.render(ctx)
                doc.save(docx_p)
                
                with st.spinner("Convertendo para PDF..."):
                    pdf_p = gerar_pdf(docx_p, tmp)
                    if pdf_p:
                        with open(pdf_p, "rb") as f:
                            st.download_button("Baixar PDF", f.read(), f"Relatorio_{ctx['SISTEMA_MES_REFERENCIA']}.pdf", "application/pdf")
                            st.success("Relatório gerado com sucesso.")
                    else:
                        st.error("Erro na conversão para PDF.")
        except Exception as e:
            st.error(f"Erro: {e}")

st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins")
