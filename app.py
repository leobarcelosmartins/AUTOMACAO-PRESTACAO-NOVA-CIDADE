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

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.4.3", layout="wide", page_icon="📑")

# Largura de 130mm para manter a harmonia visual com títulos
LARGURA_OTIMIZADA = Mm(130)

def excel_para_imagem(doc_template, arquivo_excel):
    """Lê o intervalo D3:E16 da aba TRANSFERENCIAS e converte em imagem."""
    try:
        df = pd.read_excel(
            arquivo_excel, 
            sheet_name="TRANSFERENCIAS", 
            usecols="D:E", 
            skiprows=2, 
            nrows=14, 
            header=None
        )
        
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.axis('off')
        
        tabela = ax.table(
            cellText=df.values, 
            loc='center', 
            cellLoc='center',
            colWidths=[0.5, 0.5]
        )
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(10)
        tabela.scale(1.2, 1.5)
        
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=150, transparent=True)
        plt.close(fig)
        img_buf.seek(0)
        
        return [InlineImage(doc_template, img_buf, width=LARGURA_OTIMIZADA)]
    except Exception as e:
        st.error(f"Erro ao processar intervalo Excel: {e}")
        return []

def processar_anexo(doc_template, arquivo, marcador):
    """Detecta se é arquivo (PDF/Img/Excel) ou imagem colada e retorna InlineImages."""
    if not arquivo:
        return []
    
    imagens = []
    try:
        # Lógica para imagem vinda do Clipboard (Objeto PIL Image)
        if hasattr(arquivo, 'save') and not hasattr(arquivo, 'name'):
            img_byte_arr = io.BytesIO()
            arquivo.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            imagens.append(InlineImage(doc_template, img_byte_arr, width=LARGURA_OTIMIZADA))
            return imagens

        # Lógica para ficheiros carregados (UploadedFile)
        extensao = arquivo.name.lower()
        
        if marcador == "TABELA_TRANSFERENCIA" and (extensao.endswith(".xlsx") or extensao.endswith(".xls")):
            return excel_para_imagem(doc_template, arquivo)
            
        if extensao.endswith(".pdf"):
            pdf_stream = arquivo.read()
            pdf_doc = fitz.open(stream=pdf_stream, filetype="pdf")
            for pagina in pdf_doc:
                pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_byte_arr = io.BytesIO(pix.tobytes())
                imagens.append(InlineImage(doc_template, img_byte_arr, width=LARGURA_OTIMIZADA))
            pdf_doc.close()
        else:
            imagens.append(InlineImage(doc_template, arquivo, width=LARGURA_OTIMIZADA))
        return imagens
    except Exception as e:
        st.error(f"Erro no processamento do anexo: {e}")
        return []

def gerar_pdf(docx_path, output_dir):
    """Conversão via LibreOffice Headless."""
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path],
            check=True, capture_output=True
        )
        pdf_path = os.path.join(output_dir, os.path.basename(docx_path).replace('.docx', '.pdf'))
        return pdf_path
    except Exception as e:
        st.error(f"Erro na conversão PDF: {e}")
        return None

# --- INTERFACE (UI) ---
st.title("📑 Automação de Relatórios - Backup Tático")
st.caption("Versão 0.4.3 - Ajuste de UX no Clipboard")

# Inicialização do estado para imagens coladas
if 'pasted_images' not in st.session_state:
    st.session_state.pasted_images = {}

# Estrutura de campos
campos_texto_col1 = ["SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO", "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO"]
campos_texto_col2 = ["ANALISTA_ODONTO_PED", "TOTAL_RAIO_X", "TOTAL_PACIENTES_CCIH", "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"]

campos_upload = {
    "EXCEL_META_ATENDIMENTOS": "Grade de Metas",
    "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento",
    "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X",
    "TABELA_TRANSFERENCIA": "Tabela Transferência (Excel)",
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

with st.form("form_v4_3"):
    tab1, tab2 = st.tabs(["📝 Dados Manuais e Cálculos", "🖼️ Evidências Digitais"])
    contexto = {}
    
    with tab1:
        c1, c2 = st.columns(2)
        for campo in campos_texto_col1:
            contexto[campo] = c1.text_input(campo.replace("_", " "))
        for campo in campos_texto_col2:
            contexto[campo] = c2.text_input(campo.replace("_", " "))
        
        st.write("---")
        st.subheader("📊 Indicadores de Transferência")
        c3, c4 = st.columns(2)
        contexto["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c3.number_input("Total de Transferências", step=1, value=0)
        contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c4.text_input("Taxa de Transferência (Ex: 0,76%)", value="0,00%")

    with tab2:
        st.info("💡 Você pode carregar um ficheiro ou colar um print diretamente do clipboard.")
        uploads = {}
        c_up1, c_up2 = st.columns(2)
        
        for i, (marcador, label) in enumerate(campos_upload.items()):
            col = c_up1 if i % 2 == 0 else c_up2
            with col:
                st.write(f"**{label}**")
                
                # Botão de Colar (Clipboard)
                pasted = paste_image_button(
                    label=f"📋 Colar para {label}", 
                    key=f"paste_{marcador}"
                )
                
                # Lógica de anexo com feedback transitório (Toast)
                if pasted and pasted.image_data:
                    st.session_state.pasted_images[marcador] = pasted.image_data
                    # O toast aparece no canto e some sozinho, resolvendo o problema da persistência
                    st.toast(f"✅ Print anexado em: {label}")
                
                # Uploader de Ficheiro (Tradicional)
                uploads[marcador] = st.file_uploader(
                    "Ou escolha um ficheiro", 
                    type=['png', 'jpg', 'pdf', 'xlsx', 'xls'], 
                    key=f"file_{marcador}",
                    label_visibility="collapsed"
                )
                
                # Indicador visual discreto de estado
                if marcador in st.session_state.pasted_images and not uploads[marcador]:
                    st.caption("📎 *Imagem capturada do clipboard*")
            st.write("---")

    btn_gerar = st.form_submit_button("🚀 GERAR RELATÓRIO PDF FINAL")

if btn_gerar:
    if not contexto["SISTEMA_MES_REFERENCIA"]:
        st.error("O campo 'Mês de Referência' é obrigatório.")
    else:
        try:
            # Cálculo Automático: Soma de Médicos
            m_clinico = int(contexto.get("ANALISTA_MEDICO_CLINICO", 0) or 0)
            m_pediatra = int(contexto.get("ANALISTA_MEDICO_PEDIATRA", 0) or 0)
            contexto["SISTEMA_TOTAL_MEDICOS"] = m_clinico + m_pediatra

            with tempfile.TemporaryDirectory() as pasta_temp:
                docx_temp = os.path.join(pasta_temp, "relatorio.docx")
                doc = DocxTemplate("template.docx")

                with st.spinner("Processando arquivos e capturas de tela..."):
                    for marcador in campos_upload.keys():
                        # Prioridade: Ficheiro carregado > Imagem colada
                        arquivo_final = uploads.get(marcador)
                        if not arquivo_final:
                            arquivo_final = st.session_state.pasted_images.get(marcador)
                        
                        contexto[marcador] = processar_anexo(doc, arquivo_final, marcador)

                doc.render(contexto)
                doc.save(docx_temp)
                
                with st.spinner("Convertendo para PDF..."):
                    pdf_final = gerar_pdf(docx_temp, pasta_temp)
                    
                    if pdf_final and os.path.exists(pdf_final):
                        with open(pdf_final, "rb") as f:
                            st.success("Relatório gerado com sucesso.")
                            nome_arquivo = f"Relatorio_{contexto['SISTEMA_MES_REFERENCIA'].replace('/', '-')}.pdf"
                            st.download_button("📥 Baixar Relatório PDF", f.read(), nome_arquivo, "application/pdf")
                    else:
                        st.error("Falha na conversão para PDF.")
        except Exception as e:
            st.error(f"Erro Crítico: {e}")
