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
from openpyxl import load_workbook

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.4.2", layout="wide")

# Largura de 165mm conforme solicitado para preenchimento da página
LARGURA_OTIMIZADA = Mm(165)

def excel_para_imagem(doc_template, arquivo_excel):
    """
    Lê o intervalo D3:E16 da aba TRANSFERENCIAS, limpa dados nulos,
    aplica negrito, destaca o cabeçalho e gera uma imagem para o Word.
    """
    try:
        # Leitura do intervalo D3:E16
        # skiprows=2 (pula linhas 1 e 2)
        # nrows=14 (lê 14 linhas a partir da 3)
        df = pd.read_excel(
            arquivo_excel, 
            sheet_name="TRANSFERENCIAS", 
            usecols="D:E", 
            skiprows=2, 
            nrows=14, 
            header=None
        )
        
        # Substitui valores nulos (NaN) por strings vazias
        df = df.fillna('')
        
        # Configuração da figura para renderização
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.axis('off')
        
        # Criação da tabela
        tabela = ax.table(
            cellText=df.values, 
            loc='center', 
            cellLoc='center',
            colWidths=[0.45, 0.45]
        )
        
        # Estilização técnica da tabela
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(11)
        tabela.scale(1.2, 1.8)
        
        # Iteração pelas células para aplicar formatação específica
        for (row, col), cell in tabela.get_celld().items():
            # Aplicar Negrito em todo o texto
            cell.get_text().set_weight('bold')
            
            # Formatação da primeira linha (Cabeçalho destacado)
            if row == 0:
                cell.set_facecolor('#E0E0E0')  # Cor de destaque (Cinza claro)
                # Simulação de mesclagem: se for a segunda célula da primeira linha, removemos o texto
                if col == 1:
                    cell.get_text().set_text('')
                # Centralizamos o texto da primeira célula como título do intervalo
                if col == 0:
                    cell.get_text().set_position((0.5, 0.5)) # Tenta centralizar visualmente
            
            # Bordas da tabela
            cell.set_edgecolor('#000000')
            cell.set_linewidth(1)

        # Salvar em buffer de memória com alta resolução
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200, transparent=False)
        plt.close(fig)
        img_buf.seek(0)
        
        return [InlineImage(doc_template, img_buf, width=LARGURA_OTIMIZADA)]
    except Exception as e:
        st.error(f"Erro no processamento da tabela Excel: {e}")
        return []

def processar_anexo(doc_template, arquivo, marcador=None):
    """Processa PDFs, Imagens e Excel de forma específica por marcador."""
    if not arquivo:
        return []
    
    imagens = []
    try:
        extensao = arquivo.name.lower()
        
        # Regra específica para a Tabela de Transferência vinda do Excel
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
        st.error(f"Erro no processamento do arquivo {arquivo.name}: {e}")
        return []

def gerar_pdf(docx_path, output_dir):
    """Conversão DOCX para PDF via LibreOffice."""
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path],
            check=True, capture_output=True
        )
        pdf_path = os.path.join(output_dir, os.path.basename(docx_path).replace('.docx', '.pdf'))
        return pdf_path
    except Exception as e:
        st.error(f"Erro na conversão para PDF: {e}")
        return None

# --- INTERFACE (UI LIMPA - SEM ÍCONES/EMOJIS) ---
st.title("Automacao de Relatorio de Prestacao - UPA Nova Cidade")
st.caption("Versao 0.4.3")

campos_texto_col1 = [
    "SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO",
    "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO"
]
campos_texto_col2 = [
    "ANALISTA_ODONTO_PED", "TOTAL_RAIO_X", "TOTAL_PACIENTES_CCIH", 
    "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"
]

campos_upload = {
    "EXCEL_META_ATENDIMENTOS": "Grade de Metas",
    "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento",
    "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X",
    "TABELA_TRANSFERENCIA": "Tabela Transferencia (Excel - Aba TRANSFERENCIAS)",
    "GRAFICO_TRANSFERENCIA": "Grafico Transferencia",
    "TABELA_TOTAL_OBITO": "Tabela Total Obito",
    "TABELA_OBITO": "Tabela Obito",
    "TABELA_CCIH": "Tabela CCIH",
    "IMAGEM_NEP": "Imagens NEP",
    "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno",
    "IMAGEM_MELHORIAS": "Imagens de Melhorias",
    "GRAFICO_OUVIDORIA": "Grafico Ouvidoria",
    "PDF_OUVIDORIA_INTERNA": "Relatorio Ouvidoria (PDF)",
    "TABELA_QUALITATIVA_IMG": "Tabela Qualitativa",
    "PRINT_CLASSIFICACAO": "Classificacao de Risco"
}

with st.form("form_v4_2"):
    tab1, tab2 = st.tabs(["Dados Manuais", "Arquivos"])
    contexto = {}
    
    with tab1:
        c1, c2 = st.columns(2)
        for campo in campos_texto_col1:
            contexto[campo] = c1.text_input(campo.replace("_", " "))
        for campo in campos_texto_col2:
            contexto[campo] = c2.text_input(campo.replace("_", " "))
        
        st.write("---")
        st.subheader("Indicadores de Transferencia")
        c3, c4 = st.columns(2)
        contexto["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c3.number_input("Total de Transferencias (Inteiro)", step=1, value=0)
        contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c4.text_input("Taxa de Transferencia (Ex: 0,76%)", value="0,00%")

    with tab2:
        uploads = {}
        c_up1, c_up2 = st.columns(2)
        for i, (marcador, label) in enumerate(campos_upload.items()):
            col = c_up1 if i % 2 == 0 else c_up2
            formatos = ['png', 'jpg', 'pdf', 'xlsx', 'xls'] if marcador == "TABELA_TRANSFERENCIA" else ['png', 'jpg', 'pdf']
            uploads[marcador] = col.file_uploader(label, type=formatos, key=marcador)

    btn_gerar = st.form_submit_button("GERAR RELATORIO PDF FINAL")

if btn_gerar:
    if not contexto["SISTEMA_MES_REFERENCIA"]:
        st.error("O campo Mês de Referência é obrigatório.")
    else:
        try:
            # Cálculo Automático de Médicos
            try:
                m_clinico = int(contexto.get("ANALISTA_MEDICO_CLINICO", 0) or 0)
                m_pediatra = int(contexto.get("ANALISTA_MEDICO_PEDIATRA", 0) or 0)
                contexto["SISTEMA_TOTAL_MEDICOS"] = m_clinico + m_pediatra
            except:
                contexto["SISTEMA_TOTAL_MEDICOS"] = "Erro no calculo"

            with tempfile.TemporaryDirectory() as pasta_temp:
                docx_temp = os.path.join(pasta_temp, "relatorio_final.docx")
                doc = DocxTemplate("template.docx")

                with st.spinner("Processando anexos e extraindo dados..."):
                    for marcador, arquivo in uploads.items():
                        # Passa o marcador para tratamento específico do Excel
                        contexto[marcador] = processar_anexo(doc, arquivo, marcador)

                # Mapeamento para garantir compatibilidade com tags acentuadas no Word
                if "PRINT_CLASSIFICACAO" in contexto:
                    contexto["PRINT_CLASSIFICAÇÃO"] = contexto["PRINT_CLASSIFICACAO"]

                doc.render(contexto)
                doc.save(docx_temp)
                
                with st.spinner("Convertendo para PDF..."):
                    pdf_final = gerar_pdf(docx_temp, pasta_temp)
                    
                    if pdf_final and os.path.exists(pdf_final):
                        with open(pdf_final, "rb") as f:
                            pdf_bytes = f.read()
                            st.success("✅ Relatorio gerado com sucesso.")
                            
                            nome_arquivo = f"Relatorio_{contexto['SISTEMA_MES_REFERENCIA'].replace('/', '-')}.pdf"
                            st.download_button(
                                label="Baixar Relatorio PDF",
                                data=pdf_bytes,
                                file_name=nome_arquivo,
                                mime="application/pdf"
                            )
                    else:
                        st.error("A conversão para PDF falhou.")

        except Exception as e:
            st.error(f"Erro Crítico no Sistema: {e}")

st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins")




