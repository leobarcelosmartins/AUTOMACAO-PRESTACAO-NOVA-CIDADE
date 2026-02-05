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
    "IMAGEM_PRINT_ATENDIMENTO": 160,
    "IMAGEM_DOCUMENTO_RAIO_X": 150,
    "TABELA_TRANSFERENCIA": 120,   
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

# --- INICIALIZAÇÃO DO ESTADO DA SESSÃO ---
if 'arquivos_selecionados' not in st.session_state:
    st.session_state.arquivos_selecionados = {m: [] for m in DIMENSOES_CAMPOS.keys()}

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
    """Processa um único item (arquivo ou print) para inserção no Word."""
    largura_mm = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        # CORREÇÃO: Extração do dado de imagem se for um PasteResult
        if hasattr(item, 'image_data'):
            item = item.image_data

        # Se for imagem PIL (tem o método save mas não tem atributo name)
        if hasattr(item, 'save') and not hasattr(item, 'name'):
            img_byte_arr = io.BytesIO()
            item.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            return [InlineImage(doc_template, img_byte_arr, width=Mm(largura_mm))]

        # Se for ficheiro carregado
        extensao = getattr(item, 'name', '').lower()
        if marcador == "TABELA_TRANSFERENCIA" and (extensao.endswith(".xlsx") or extensao.endswith(".xls")):
            res = excel_para_imagem(doc_template, item)
            return [res] if res else []

        if extensao.endswith(".pdf"):
            pdf_doc = fitz.open(stream=item.read(), filetype="pdf")
            imgs_pdf = []
            for pagina in pdf_doc:
                pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_byte_arr = io.BytesIO(pix.tobytes())
                imgs_pdf.append(InlineImage(doc_template, img_byte_arr, width=Mm(largura_mm)))
            pdf_doc.close()
            return imgs_pdf

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

# --- UI PRINCIPAL ---
st.title("Automação de Relatório de Prestação - UPA Nova Cidade")
st.caption("Versão 0.4.3 - Edição Corrigida (Clipboard)")

col_t1 = ["SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO", "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO"]
col_t2 = ["ANALISTA_ODONTO_PED", "TOTAL_RAIO_X", "TOTAL_PACIENTES_CCIH", "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"]

marcadores_config = {
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

tab_man, tab_arq = st.tabs(["Dados Manuais", "Arquivos de Evidência"])
ctx = {}

with tab_man:
    with st.form("form_dados"):
        c1, c2 = st.columns(2)
        for f in col_t1: ctx[f] = c1.text_input(f.replace("_", " "))
        for f in col_t2: ctx[f] = c2.text_input(f.replace("_", " "))
        st.write("---")
        c3, c4 = st.columns(2)
        ctx["SISTEMA_TOTAL_DE_TRANSFERENCIA"] = c3.number_input("Total de Transferências", step=1, value=0)
        ctx["SISTEMA_TAXA_DE_TRANSFERENCIA"] = c4.text_input("Taxa de Transferência (Ex: 0,76%)", value="0,00%")
        st.form_submit_button("Salvar Textos")

with tab_arq:
    u1, u2 = st.columns(2)
    for i, (m, label) in enumerate(marcadores_config.items()):
        col = u1 if i % 2 == 0 else u2
        with col:
            st.markdown(f"**{label}**")
            
            # 1. Colar do Clipboard
            pasted = paste_image_button(label="Colar print", key=f"p_{m}")
            if pasted:
                nome_print = f"Captura_{len(st.session_state.arquivos_selecionados[m]) + 1}"
                # Armazenamos o image_data diretamente se for um PasteResult
                img_data = pasted.image_data if hasattr(pasted, 'image_data') else pasted
                st.session_state.arquivos_selecionados[m].append({"name": nome_print, "content": img_data, "type": "print"})
                st.rerun()

            # 2. Upload de Arquivo
            tipo_f = ['png', 'jpg', 'pdf', 'xlsx', 'xls'] if m == "TABELA_TRANSFERENCIA" else ['png', 'jpg', 'pdf']
            files = st.file_uploader("Ficheiro", type=tipo_f, key=f"f_{m}", accept_multiple_files=True, label_visibility="collapsed")
            if files:
                for f in files:
                    if f.name not in [x["name"] for x in st.session_state.arquivos_selecionados[m]]:
                        st.session_state.arquivos_selecionados[m].append({"name": f.name, "content": f, "type": "file"})
                st.rerun()

            # 3. Lista Dinâmica de Itens Recebidos
            if st.session_state.arquivos_selecionados[m]:
                for idx, item in enumerate(st.session_state.arquivos_selecionados[m]):
                    with st.expander(f"📄 {item['name']}"):
                        if item['type'] == "print" or not item['name'].lower().endswith(('.pdf', '.xlsx', '.xls')):
                            st.image(item['content'], caption=item['name'], use_container_width=True)
                        else:
                            st.info("Pré-visualização não disponível para este formato.")
                        
                        if st.button("Remover", key=f"del_{m}_{idx}", use_container_width=True):
                            st.session_state.arquivos_selecionados[m].pop(idx)
                            st.rerun()
            st.write("---")

# --- BOTÃO DE GERAÇÃO ---
if st.button("🚀 GERAR RELATÓRIO PDF FINAL", use_container_width=True):
    if not ctx.get("SISTEMA_MES_REFERENCIA"):
        st.error("O campo 'Mês de Referência' é obrigatório.")
    else:
        try:
            mc = int(ctx.get("ANALISTA_MEDICO_CLINICO") or 0)
            mp = int(ctx.get("ANALISTA_MEDICO_PEDIATRA") or 0)
            ctx["SISTEMA_TOTAL_MEDICOS"] = mc + mp

            with tempfile.TemporaryDirectory() as tmp:
                docx_p = os.path.join(tmp, "temp.docx")
                doc = DocxTemplate("template.docx")

                with st.spinner("Consolidando todos os arquivos e prints..."):
                    dados_finais = ctx.copy()
                    for m in marcadores_config.keys():
                        lista_imgs = []
                        for item in st.session_state.arquivos_selecionados[m]:
                            processado = processar_item(doc, item['content'], m)
                            if processado:
                                lista_imgs.extend(processado)
                        dados_finais[m] = lista_imgs

                doc.render(dados_finais)
                doc.save(docx_p)
                
                pdf_p = gerar_pdf(docx_p, tmp)
                if pdf_p:
                    with open(pdf_p, "rb") as f:
                        st.success("Relatório gerado com sucesso.")
                        st.download_button("📥 Baixar PDF", f.read(), f"Relatorio_{ctx['SISTEMA_MES_REFERENCIA']}.pdf", "application/pdf")
        except Exception as e:
            st.error(f"Erro Crítico: {e}")

st.markdown("---")
st.caption("Desenvolvido por Leonardo Barcelos Martins | Backup Tático")
