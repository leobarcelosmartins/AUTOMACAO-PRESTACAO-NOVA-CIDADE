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
from st_paste_button import paste_image_button

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.4.1", layout="wide")

# Largura de 165mm para garantir o preenchimento da página nas imagens padrão
LARGURA_OTIMIZADA = Mm(165)
# Largura reduzida para a tabela Excel conforme solicitado
LARGURA_TABELA = Mm(120)

def excel_para_imagem(doc_template, arquivo_excel):
    """
    Lê o intervalo D3:E16 da aba TRANSFERENCIAS, limpa nulos, 
    formata a segunda coluna como inteiro, aplica negrito e gera imagem.
    """
    try:
        # Leitura do intervalo D3:E16 (Colunas D e E são índices 3 e 4)
        df = pd.read_excel(
            arquivo_excel, 
            sheet_name="TRANSFERENCIAS", 
            usecols=[3, 4], 
            skiprows=2, 
            nrows=14, 
            header=None
        )
        
        # Substitui valores nulos por vazio
        df = df.fillna('')
        
        col_labels = df.columns
        
        def format_inteiro(val):
            if val == '' or val is None: return ''
            try:
                # Converte para float e depois int para remover decimais
                return str(int(float(val)))
            except (ValueError, TypeError):
                return str(val)
        
        # Formata a segunda coluna (índice 1 do DataFrame resultante)
        if len(col_labels) > 1:
            df[col_labels[1]] = df[col_labels[1]].apply(format_inteiro)
        
        # Configuração da figura para renderização
        fig, ax = plt.subplots(figsize=(8, 6))
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
        
        # Iteração pelas células para aplicar formatação
        for (row, col), cell in tabela.get_celld().items():
            # Aplicar Negrito em todo o texto
            cell.get_text().set_weight('bold')
            
            # Formatação da primeira linha (Cabeçalho destacado / Simulação de Mesclagem)
            if row == 0:
                cell.set_facecolor('#D3D3D3')  # Cinza claro de destaque
                if col == 1:
                    cell.get_text().set_text('')
                if col == 0:
                    cell.get_text().set_position((0.5, 0.5))
            
            # Bordas pretas nítidas
            cell.set_edgecolor('#000000')
            cell.set_linewidth(1)

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200, transparent=False)
        plt.close(fig)
        img_buf.seek(0)
        
        return [InlineImage(doc_template, img_buf, width=LARGURA_TABELA)]
    except Exception as e:
        st.error(f"Erro no processamento da tabela Excel: {e}")
        return []

def processar_conteudo(doc_template, conteudo, marcador=None):
    """Processa tanto ficheiros carregados como imagens coladas do clipboard."""
    if not conteudo:
        return []
    
    imagens = []
    try:
        # Verifica se é uma imagem colada (objeto Image do PIL)
        if hasattr(conteudo, 'save'):
            img_byte_arr = io.BytesIO()
            conteudo.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            imagens.append(InlineImage(doc_template, img_byte_arr, width=LARGURA_OTIMIZADA))
            return imagens

        # Caso contrário, trata como ficheiro carregado (UploadedFile)
        extensao = conteudo.name.lower()
        
        if marcador == "TABELA_TRANSFERENCIA" and (extensao.endswith(".xlsx") or extensao.endswith(".xls")):
            return excel_para_imagem(doc_template, conteudo)

        if extensao.endswith(".pdf"):
            pdf_stream = conteudo.read()
            pdf_doc = fitz.open(stream=pdf_stream, filetype="pdf")
            for pagina in pdf_doc:
                pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_byte_arr = io.BytesIO(pix.tobytes())
                imagens.append(InlineImage(doc_template, img_byte_arr, width=LARGURA_OTIMIZADA))
            pdf_doc.close()
        else:
            imagens.append(InlineImage(doc_template, conteudo, width=LARGURA_OTIMIZADA))
        return imagens
    except Exception as e:
        st.error(f"Erro no processamento de conteúdo: {e}")
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
st.title("Automação de Relatório de Prestação - UPA Nova Cidade")
st.caption("Versão 0.4.1")

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

# Inicialização do estado para imagens coladas
if 'pasted_images' not in st.session_state:
    st.session_state.pasted_images = {}

with st.form("form_v4_1"):
    tab1, tab2 = st.tabs(["Dados Manuais", "Arquivos"])
    contexto = {}
    
    with tab1:
        c1, c2 = st.columns(2)
        for campo in campos_texto_col1:
            contexto[campo] = c1.text_input(campo.replace("_", " "))
        for campo in campos_texto_col2:
            contexto[campo] = c2.text_input(campo.replace("_", " "))
        
        st.write("---")
        st.subheader("Indicadores de Transferência")
        c3, c4 = st.columns(2)
        contexto["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c3.number_input("Total de Transferências (Inteiro)", step=1, value=0)
        contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c4.text_input("Taxa de Transferência (Ex: 0,76%)", value="0,00%")

    with tab2:
        uploads = {}
        c_up1, c_up2 = st.columns(2)
        for i, (marcador, label) in enumerate(campos_upload.items()):
            col = c_up1 if i % 2 == 0 else c_up2
            with col:
                st.write(f"**{label}**")
                # Botão para colar imagem do clipboard
                pasted_img = paste_image_button(
                    label=f"Colar print para {label}",
                    key=f"paste_{marcador}"
                )
                if pasted_img:
                    st.session_state.pasted_images[marcador] = pasted_img.image_data
                    st.info("Imagem capturada do clipboard.")

                # Upload de ficheiro tradicional
                tipos = ['png', 'jpg', 'pdf', 'xlsx', 'xls'] if marcador == "TABELA_TRANSFERENCIA" else ['png', 'jpg', 'pdf']
                uploads[marcador] = st.file_uploader(
                    "Ou escolha um ficheiro", 
                    type=tipos, 
                    key=f"file_{marcador}",
                    label_visibility="collapsed"
                )

    btn_gerar = st.form_submit_button("GERAR RELATÓRIO PDF FINAL")

if btn_gerar:
    if not contexto["SISTEMA_MES_REFERENCIA"]:
        st.error("O campo 'Mês de Referência' é obrigatório.")
    else:
        try:
            # Cálculo Automático: Soma de Médicos
            try:
                m_clinico = int(contexto.get("ANALISTA_MEDICO_CLINICO", 0) or 0)
                m_pediatra = int(contexto.get("ANALISTA_MEDICO_PEDIATRA", 0) or 0)
                contexto["SISTEMA_TOTAL_MEDICOS"] = m_clinico + m_pediatra
            except Exception:
                contexto["SISTEMA_TOTAL_MEDICOS"] = "Erro no cálculo"

            with tempfile.TemporaryDirectory() as pasta_temp:
                docx_temp = os.path.join(pasta_temp, "relatorio_final.docx")
                doc = DocxTemplate("template.docx")

                with st.spinner("Processando anexos e cálculos..."):
                    for marcador in campos_upload.keys():
                        # Prioriza o ficheiro carregado, se não existir, tenta a imagem colada
                        conteudo = uploads.get(marcador) or st.session_state.pasted_images.get(marcador)
                        contexto[marcador] = processar_conteudo(doc, conteudo, marcador)

                doc.render(contexto)
                doc.save(docx_temp)
                
                with st.spinner("A converter para PDF..."):
                    pdf_final = gerar_pdf(docx_temp, pasta_temp)
                    
                    if pdf_final and os.path.exists(pdf_final):
                        with open(pdf_final, "rb") as f:
                            pdf_bytes = f.read()
                            st.success("Relatório gerado com sucesso.")
                            
                            nome_arquivo = f"Relatorio_{contexto['SISTEMA_MES_REFERENCIA'].replace('/', '-')}.pdf"
                            st.download_button(
                                label="Baixar Relatório PDF",
                                data=pdf_bytes,
                                file_name=nome_arquivo,
                                mime="application/pdf"
                            )
                    else:
                        st.error("A conversão para PDF falhou.")

        except Exception as e:
            st.error(f"Erro Crítico no Sistema: {e}")
            
# --- RODAPÉ ---
st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins")
