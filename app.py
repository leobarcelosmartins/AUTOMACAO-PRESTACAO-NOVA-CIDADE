import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import subprocess
import tempfile

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(
    page_title="Gerador de Relatórios Assistenciais",
    page_icon="📑",
    layout="wide"
)

# --- FUNÇÕES AUXILIARES ---

def converter_pdf_para_imagens(doc_template, arquivo_pdf):
    """
    Converte cada página de um PDF enviado em objetos InlineImage para o Word.
    """
    imagens = []
    try:
        # Lê o conteúdo do ficheiro enviado
        pdf_stream = arquivo_pdf.read()
        pdf_doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        for pagina in pdf_doc:
            # Renderiza a página como imagem (zoom de 2x para manter legibilidade)
            pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_byte_arr = io.BytesIO(pix.tobytes())
            # Define a largura padrão (160mm cabe bem em A4 com margens)
            imagens.append(InlineImage(doc_template, img_byte_arr, width=Mm(160)))
            
        pdf_doc.close()
        return imagens
    except Exception as e:
        st.error(f"Erro ao processar o PDF anexado: {e}")
        return []

def preparar_imagem_simples(doc_template, arquivo_img):
    """
    Prepara uma imagem (PNG/JPG) como uma lista contendo um objeto InlineImage.
    """
    try:
        return [InlineImage(doc_template, arquivo_img, width=Mm(160))]
    except Exception as e:
        st.error(f"Erro ao processar a imagem: {e}")
        return []

def converter_docx_para_pdf(docx_path, output_dir):
    """
    Usa o LibreOffice instalado no servidor (via packages.txt) para converter DOCX em PDF.
    """
    try:
        # Executa o comando headless do LibreOffice
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path],
            check=True,
            capture_output=True,
            text=True
        )
        # O LibreOffice gera o PDF com o mesmo nome na pasta de saída
        nome_pdf = os.path.basename(docx_path).replace('.docx', '.pdf')
        return os.path.join(output_dir, nome_pdf)
    except Exception as e:
        st.error(f"Erro na conversão para PDF: {e}")
        st.info("Verifique se o ficheiro 'packages.txt' contém 'libreoffice' e se o deploy foi concluído.")
        return None

# --- INTERFACE DO UTILIZADOR ---

st.title("📑 Automação de Relatórios: Backup Tático")
st.markdown("Preencha os dados abaixo e anexe as evidências para gerar o relatório em **PDF**.")

# Definição dos campos conforme o Relatório Modelo
campos_manuais = [
    "SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO",
    "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO", "ANALISTA_ODONTO_PED",
    "TOTAL_RAIO_X", "SISTEMA_TOTAL_DE_TRANSFERENCIA", "TOTAL_PACIENTES_CCIH",
    "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"
]

campos_upload = {
    "EXCEL_META_ATENDIMENTOS": "Grade de Metas (Excel/Print)",
    "IMAGEM_PRINT_ATENDIMENTO": "Print de Atendimento",
    "IMAGEM_DOCUMENTO_RAIO_X": "Documento Raio-X",
    "TABELA_TRANSFERENCIA": "Tabela de Transferência",
    "GRAFICO_TRANSFERENCIA": "Gráfico de Transferência",
    "TABELA_TOTAL_OBITO": "Tabela Total de Óbitos",
    "TABELA_OBITO": "Tabela de Óbitos Detalhada",
    "TABELA_CCIH": "Tabela CCIH",
    "IMAGEM_NEP": "Imagens NEP",
    "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno",
    "IMAGEM_MELHORIAS": "Imagens de Melhorias",
    "GRAFICO_OUVIDORIA": "Gráfico de Ouvidoria",
    "PDF_OUVIDORIA_INTERNA": "Relatório de Ouvidoria (PDF)",
    "TABELA_QUALITATIVA_IMG": "Tabela Qualitativa",
    "PRINT_CLASSIFICACAO": "Relatório de Classificação de Risco"
}

with st.form("form_gerador"):
    col1, col2 = st.columns(2)
    contexto = {}

    with col1:
        st.subheader("✍️ Dados da Produção")
        for campo in campos_manuais:
            contexto[campo] = st.text_input(campo.replace("_", " "), placeholder=f"Introduza {campo.lower()}")
        
        st.write("---")
        st.subheader("🏥 Transferências")
        destinos_input = st.text_area("Destinos de Transferência (Um por linha)", height=100)
        # Lógica solicitada: múltiplos nomes separados por " / "
        contexto["MANUAL_DESTINO_TRANSFERENCIA"] = " / ".join([d.strip() for d in destinos_input.split('\n') if d.strip()])

    with col2:
        st.subheader("📁 Anexos e Evidências")
        uploads = {}
        for marcador, label in campos_upload.items():
            uploads[marcador] = st.file_uploader(f"{label}", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"up_{marcador}")

    st.write("---")
    botao_gerar = st.form_submit_button("🚀 GERAR RELATÓRIO PDF")

# --- PROCESSAMENTO DOS DADOS ---

if botao_gerar:
    if not contexto["SISTEMA_MES_REFERENCIA"]:
        st.error("O campo 'SISTEMA MES REFERENCIA' é obrigatório.")
    else:
        try:
            # Caminho do template no repositório
            template_path = "template.docx"
            
            if not os.path.exists(template_path):
                st.error("Ficheiro 'template.docx' não encontrado no repositório.")
                st.stop()

            # Usamos uma pasta temporária para segurança dos dados
            with tempfile.TemporaryDirectory() as pasta_temp:
                caminho_docx_temp = os.path.join(pasta_temp, "processando.docx")
                
                # Inicia o motor do template
                doc = DocxTemplate(template_path)
                
                # 1. Cálculo Automático da Taxa de Transferência
                try:
                    total_aten = float(contexto.get("ANALISTA_TOTAL_ATENDIMENTOS", 0))
                    total_trans = float(contexto.get("SISTEMA_TOTAL_DE_TRANSFERENCIA", 0))
                    taxa = (total_trans / total_aten * 100) if total_aten > 0 else 0
                    contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = f"{taxa:.2f}%"
                except ValueError:
                    contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = "0.00%"

                # 2. Processamento de Imagens e PDFs
                with st.spinner("A processar anexos e a converter PDFs..."):
                    for marcador, arquivo in uploads.items():
                        if arquivo:
                            if arquivo.name.lower().endswith(".pdf"):
                                contexto[marcador] = converter_pdf_para_imagens(doc, arquivo)
                            else:
                                contexto[marcador] = preparar_imagem_simples(doc, arquivo)
                        else:
                            # Se não houver upload, enviamos lista vazia para o loop {% for %} não falhar
                            contexto[marcador] = []

                # 3. Renderização do Word
                doc.render(contexto)
                doc.save(caminho_docx_temp)
                
                # 4. Conversão para PDF
                with st.spinner("A converter para PDF (LibreOffice)..."):
                    caminho_pdf_final = converter_docx_para_pdf(caminho_docx_temp, pasta_temp)
                    
                    if caminho_pdf_final and os.path.exists(caminho_pdf_final):
                        with open(caminho_pdf_final, "rb") as f:
                            pdf_bytes = f.read()
                        
                        st.success("✅ Relatório gerado com sucesso!")
                        
                        # Nome do ficheiro de saída
                        nome_download = f"Relatorio_Assistencial_{contexto['SISTEMA_MES_REFERENCIA'].replace('/', '-')}.pdf"
                        
                        st.download_button(
                            label="📥 Baixar Relatório em PDF",
                            data=pdf_bytes,
                            file_name=nome_download,
                            mime="application/pdf"
                        )
                    else:
                        st.error("A conversão para PDF falhou. Verifique os logs.")
        
        except Exception as e:
            st.error(f"Ocorreu um erro inesperado: {e}")

# --- RODAPÉ ---
st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins - Backup Tático")
