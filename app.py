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
    page_title="Gerador de Relatório Assistencial - UPA Nova Cidade",
    page_icon="📑",
    layout="wide"
)

# --- CONFIGURAÇÕES DE DESIGN (PO DECISION) ---
# Redução de 160mm para 140mm para garantir que a imagem caiba logo abaixo do título
LARGURA_PADRAO = Mm(140)

def converter_pdf_para_imagens(doc_template, arquivo_pdf):
    """Converte cada página de um PDF em objetos InlineImage para o Word."""
    imagens = []
    try:
        pdf_stream = arquivo_pdf.read()
        pdf_doc = fitz.open(stream=pdf_stream, filetype="pdf")
        for pagina in pdf_doc:
            # Renderização de alta qualidade (2x zoom)
            pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_byte_arr = io.BytesIO(pix.tobytes())
            imagens.append(InlineImage(doc_template, img_byte_arr, width=LARGURA_PADRAO))
        pdf_doc.close()
        return imagens
    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
        return []

def preparar_imagem_simples(doc_template, arquivo_img):
    """Prepara imagem JPG/PNG para inserção via loop no Word com largura otimizada."""
    try:
        return [InlineImage(doc_template, arquivo_img, width=LARGURA_PADRAO)]
    except:
        return []

def gerar_pdf_via_libreoffice(docx_path, output_dir):
    """Converte o DOCX resultante para PDF usando LibreOffice Headless (Ambiente Linux)."""
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path],
            check=True, capture_output=True, text=True
        )
        pdf_name = os.path.basename(docx_path).replace('.docx', '.pdf')
        return os.path.join(output_dir, pdf_name)
    except Exception as e:
        st.error(f"Erro na conversão PDF: {e}")
        return None

# --- INTERFACE DE UTILIZADOR ---

st.title("📑 Automação de Relatórios")
st.markdown("Preencha os dados e anexe as evidências. O sistema otimizará o tamanho das imagens para o layout.")

# Definição de campos para organização em Abas
campos_manuais = [
    "SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO",
    "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO", "ANALISTA_ODONTO_PED",
    "TOTAL_RAIO_X", "SISTEMA_TOTAL_DE_TRANSFERENCIA", "TOTAL_PACIENTES_CCIH",
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

with st.form("main_form"):
    tab1, tab2 = st.tabs(["✍️ Informações de Texto", "📁 Upload de Anexos"])
    contexto = {}

    with tab1:
        c1, c2 = st.columns(2)
        for i, campo in enumerate(campos_manuais):
            col = c1 if i < 6 else c2
            contexto[campo] = col.text_input(campo.replace("_", " "), key=f"txt_{campo}")
        
        st.write("---")
        destinos = st.text_area("MANUAL DESTINO TRANSFERÊNCIA (Um por linha)", height=100)
        contexto["MANUAL_DESTINO_TRANSFERENCIA"] = " / ".join([d.strip() for d in destinos.split("\n") if d.strip()])

    with tab2:
        uploads = {}
        c3, c4 = st.columns(2)
        for i, (marcador, label) in enumerate(campos_upload.items()):
            col = c3 if i % 2 == 0 else c4
            uploads[marcador] = col.file_uploader(label, type=['png', 'jpg', 'pdf'], key=f"up_{marcador}")

    botao_gerar = st.form_submit_button("🚀 GERAR RELATÓRIO PDF OTIMIZADO")

# --- LÓGICA DE PROCESSAMENTO ---

if botao_gerar:
    if not contexto["SISTEMA_MES_REFERENCIA"]:
        st.error("O campo 'Mês de Referência' é obrigatório.")
    else:
        try:
            with tempfile.TemporaryDirectory() as pasta_temp:
                docx_path = os.path.join(pasta_temp, "relatorio.docx")
                doc = DocxTemplate("template.docx")

                # Cálculo de Indicadores
                try:
                    total = float(contexto.get("ANALISTA_TOTAL_ATENDIMENTOS", 0))
                    trans = float(contexto.get("SISTEMA_TOTAL_DE_TRANSFERENCIA", 0))
                    taxa = (trans / total) * 100 if total > 0 else 0
                    contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = f"{taxa:.2f}%"
                except:
                    contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = "0.00%"

                # Processamento de Ficheiros com Largura de 140mm
                with st.spinner("Otimizando dimensões das imagens..."):
                    for marcador, arquivo in uploads.items():
                        if arquivo:
                            if arquivo.name.lower().endswith(".pdf"):
                                contexto[marcador] = converter_pdf_para_imagens(doc, arquivo)
                            else:
                                contexto[marcador] = preparar_imagem_simples(doc, arquivo)
                        else:
                            contexto[marcador] = []

                doc.render(contexto)
                doc.save(docx_path)
                
                with st.spinner("Convertendo para PDF final..."):
                    pdf_path = gerar_pdf_via_libreoffice(docx_path, pasta_temp)
                    
                    if pdf_path and os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()
                        
                        st.success("✅ Relatório gerado com sucesso!")
                        nome_pdf = f"Relatorio_{contexto['SISTEMA_MES_REFERENCIA'].replace('/', '-')}.pdf"
                        st.download_button("📥 Baixar PDF", pdf_bytes, nome_pdf, "application/pdf")
                    else:
                        st.error("Falha na conversão para PDF.")
        
        except Exception as e:
            st.error(f"Erro Crítico: {e}")


# --- RODAPÉ ---
st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins")
