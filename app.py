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
st.set_page_config(page_title="Gerador de Relatórios V0.4.3", layout="wide")

# --- DICIONÁRIO DE DIMENSÕES POR CAMPO (LARGURAS EM MM) ---
DIMENSOES_CAMPOS = {
    "EXCEL_META_ATENDIMENTOS": 165,
    "IMAGEM_PRINT_ATENDIMENTO": 165,
    "IMAGEM_DOCUMENTO_RAIO_X": 165,
    "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 160,
    "TABELA_TOTAL_OBITO": 165,
    "TABELA_OBITO": 180,
    "TABELA_CCIH": 180,
    "IMAGEM_NEP": 180,
    "IMAGEM_TREINAMENTO_INTERNO": 180,
    "IMAGEM_MELHORIAS": 180,
    "GRAFICO_OUVIDORIA": 155,
    "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 170,
    "PRINT_CLASSIFICACAO": 160
}

# --- INICIALIZAÇÃO DO ESTADO ---
# Usamos o session_state para manter a lista de arquivos de cada marcador
if 'lista_arquivos' not in st.session_state:
    st.session_state.lista_arquivos = {m: [] for m in [
        "EXCEL_META_ATENDIMENTOS", "IMAGEM_PRINT_ATENDIMENTO", "IMAGEM_DOCUMENTO_RAIO_X",
        "TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA", "TABELA_TOTAL_OBITO",
        "TABELA_OBITO", "TABELA_CCIH", "IMAGEM_NEP", "IMAGEM_TREINAMENTO_INTERNO",
        "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA",
        "TABELA_QUALITATIVA_IMG", "PRINT_CLASSIFICACAO"
    ]}

def excel_para_imagem(doc_template, arquivo_excel):
    """Extrai o intervalo D3:E16 da aba TRANSFERENCIAS com formatação profissional."""
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
                if col == 0: cell.get_text().set_position((0.5, 0.5))

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        img_buf.seek(0)
        
        largura_mm = DIMENSOES_CAMPOS.get("TABELA_TRANSFERENCIA", 120)
        return InlineImage(doc_template, img_buf, width=Mm(largura_mm))
    except Exception as e:
        st.error(f"Erro no processamento da tabela Excel: {e}")
        return None

def processar_item(doc_template, item, marcador):
    """Processa um único item (arquivo ou imagem colada) e retorna InlineImage ou lista de InlineImages."""
    largura_mm = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        # Se for imagem colada (objeto PIL Image)
        if hasattr(item, 'save') and not hasattr(item, 'name'):
            img_byte_arr = io.BytesIO()
            item.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            return [InlineImage(doc_template, img_byte_arr, width=Mm(largura_mm))]

        # Se for Excel (apenas para o marcador específico)
        extensao = getattr(item, 'name', '').lower()
        if marcador == "TABELA_TRANSFERENCIA" and (extensao.endswith(".xlsx") or extensao.endswith(".xls")):
            resultado = excel_para_imagem(doc_template, item)
            return [resultado] if resultado else []

        # Se for PDF
        if extensao.endswith(".pdf"):
            pdf_doc = fitz.open(stream=item.read(), filetype="pdf")
            imgs_pdf = []
            for pagina in pdf_doc:
                pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_byte_arr = io.BytesIO(pix.tobytes())
                imgs_pdf.append(InlineImage(doc_template, img_byte_arr, width=Mm(largura_mm)))
            pdf_doc.close()
            return imgs_pdf

        # Imagem padrão
        return [InlineImage(doc_template, item, width=Mm(largura_mm))]
    except Exception as e:
        st.error(f"Erro no item do marcador {marcador}: {e}")
        return []

def gerar_pdf(docx_path, output_dir):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True, capture_output=True)
        return os.path.join(output_dir, os.path.basename(docx_path).replace('.docx', '.pdf'))
    except:
        return None

# --- UI PRINCIPAL ---
st.title("Automação de Relatório de Prestação - UPA Nova Cidade")
st.caption("Versão 0.4.3 - Gestor de Multi-Evidências")

col_t1_campos = ["SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO", "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO"]
col_t2_campos = ["ANALISTA_ODONTO_PED", "TOTAL_RAIO_X", "TOTAL_PACIENTES_CCIH", "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"]

marcadores_evidencia = {
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

# --- ABAS ---
tab_manual, tab_arquivos = st.tabs(["Dados Manuais", "Gestor de Arquivos"])
contexto_manual = {}

with tab_manual:
    with st.form("manual_form"):
        c1, c2 = st.columns(2)
        for f in col_t1_campos: contexto_manual[f] = c1.text_input(f.replace("_", " "))
        for f in col_t2_campos: contexto_manual[f] = c2.text_input(f.replace("_", " "))
        st.write("---")
        c3, c4 = st.columns(2)
        contexto_manual["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c3.number_input("Total de Transferências", step=1, value=0)
        contexto_manual["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c4.text_input("Taxa de Transferência (Ex: 0,76%)", value="0,00%")
        submit_manual = st.form_submit_button("Salvar Dados de Texto")
        if submit_manual:
            st.success("Dados de texto salvos temporariamente.")

with tab_arquivos:
    st.write("Anexe arquivos ou cole prints para cada campo abaixo.")
    
    cup1, cup2 = st.columns(2)
    for i, (m, label) in enumerate(marcadores_evidencia.items()):
        col_alvo = cup1 if i % 2 == 0 else cup2
        with col_alvo:
            st.markdown(f"#### {label}")
            
            # 1. Botão de Colar
            pasted = paste_image_button(label=f"Colar print", key=f"p_{m}")
            if pasted:
                nome_p = f"Captura_{len(st.session_state.lista_arquivos[m]) + 1}"
                st.session_state.lista_arquivos[m].append({"name": nome_p, "content": pasted.image_data, "type": "pasted"})
                st.toast(f"Print recebido para {label}")

            # 2. Uploader de Ficheiro (Multi-seleção)
            tipo_f = ['png', 'jpg', 'pdf', 'xlsx', 'xls'] if m == "TABELA_TRANSFERENCIA" else ['png', 'jpg', 'pdf']
            files = st.file_uploader("Adicionar arquivos", type=tipo_f, key=f"f_{m}", accept_multiple_files=True, label_visibility="collapsed")
            
            if files:
                for f in files:
                    # Evita duplicidade simples pelo nome
                    if f.name not in [x["name"] for x in st.session_state.lista_arquivos[m]]:
                        st.session_state.lista_arquivos[m].append({"name": f.name, "content": f, "type": "uploaded"})

            # 3. Lista de arquivos enviados com Preview e Delete
            if st.session_state.lista_arquivos[m]:
                st.write("Arquivos enviados:")
                for idx, item in enumerate(st.session_state.lista_arquivos[m]):
                    with st.expander(f"📄 {item['name']}"):
                        # Preview
                        if item['type'] == "pasted":
                            st.image(item['content'], caption="Visualização do Print", width=300)
                        elif not item['name'].lower().endswith(('.pdf', '.xlsx', '.xls')):
                            st.image(item['content'], caption=item['name'], width=300)
                        else:
                            st.info("Visualização prévia não disponível para este formato.")
                        
                        # Botão Excluir
                        if st.button("Remover arquivo", key=f"del_{m}_{idx}"):
                            st.session_state.lista_arquivos[m].pop(idx)
                            st.rerun()
            st.write("---")

# --- BOTÃO FINAL DE GERAÇÃO ---
st.write("")
if st.button("🚀 GERAR RELATÓRIO PDF FINAL", use_container_width=True):
    if not contexto_manual.get("SISTEMA_MES_REFERENCIA"):
        st.error("O campo 'Mês de Referência' é obrigatório.")
    else:
        try:
            # Cálculos Médicos
            try:
                mc = int(contexto_manual.get("ANALISTA_MEDICO_CLINICO") or 0)
                mp = int(contexto_manual.get("ANALISTA_MEDICO_PEDIATRA") or 0)
                contexto_manual["SISTEMA_TOTAL_MEDICOS"] = mc + mp
            except:
                contexto_manual["SISTEMA_TOTAL_MEDICOS"] = 0

            with tempfile.TemporaryDirectory() as tmp_dir:
                docx_out = os.path.join(tmp_dir, "temp_relatorio.docx")
                doc_tpl = DocxTemplate("template.docx")

                # Consolidação de Conteúdo (Injeção no Word)
                with st.spinner("Consolidando evidências e gerando documento..."):
                    # Unimos os dados manuais com as listas de imagens processadas
                    dados_finais = contexto_manual.copy()
                    for m in marcadores_evidencia.keys():
                        lista_imgs_processadas = []
                        for item in st.session_state.lista_arquivos[m]:
                            # Cada processar_item retorna uma LISTA (porque PDFs podem ter várias páginas)
                            resultado = processar_item(doc_tpl, item['content'], m)
                            if resultado:
                                lista_imgs_processadas.extend(resultado)
                        dados_finais[m] = lista_imgs_processadas

                doc_tpl.render(dados_finais)
                doc_tpl.save(docx_out)
                
                with st.spinner("Convertendo para formato PDF..."):
                    pdf_res = gerar_pdf(docx_out, tmp_dir)
                    if pdf_res:
                        with open(pdf_res, "rb") as f:
                            pdf_bytes = f.read()
                            st.success("Relatório gerado com sucesso.")
                            st.download_button("📥 Baixar Relatório PDF", pdf_bytes, f"Relatorio_{contexto_manual['SISTEMA_MES_REFERENCIA']}.pdf", "application/pdf")
                    else:
                        st.error("Falha na conversão para PDF via LibreOffice.")
        except Exception as e:
            st.error(f"Erro Crítico: {e}")

st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins | Backup Tático")
