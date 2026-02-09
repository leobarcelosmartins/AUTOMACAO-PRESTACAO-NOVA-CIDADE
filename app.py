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
st.set_page_config(page_title="Gerador de Relatórios V0.6.0", layout="wide")

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
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    .anexo-badge { 
        background-color: #28a745; 
        color: white; 
        padding: 2px 8px; 
        border-radius: 12px; 
        font-size: 11px; 
        font-weight: bold;
        margin-left: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "EXCEL_META_ATENDIMENTOS": 165, "IMAGEM_PRINT_ATENDIMENTO": 165,
    "IMAGEM_DOCUMENTO_RAIO_X": 165, "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 160, "TABELA_TOTAL_OBITO": 165,
    "TABELA_OBITO": 180, "TABELA_CCIH": 180, "IMAGEM_NEP": 180,
    "IMAGEM_TREINAMENTO_INTERNO": 180, "IMAGEM_MELHORIAS": 180,
    "GRAFICO_OUVIDORIA": 155, "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 170, "PRINT_CLASSIFICACAO": 160
}

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

if 'ultimo_print_time' not in st.session_state:
    st.session_state.ultimo_print_time = {m: 0 for m in DIMENSOES_CAMPOS.keys()}

def excel_para_imagem(doc_template, arquivo_excel):
    try:
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
    """Processa item e retorna lista de InlineImage para o template"""
    largura = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        # CORREÇÃO 1: Se item é bytes (prints colados), processar diretamente
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura))]

        # CORREÇÃO 2: Se item tem atributo 'name', é um arquivo uploadado
        ext = getattr(item, 'name', '').lower()

        # Processar Excel especial
        if marcador == "TABELA_TRANSFERENCIA" and (ext.endswith(".xlsx") or ext.endswith(".xls")):
            res = excel_para_imagem(doc_template, item)
            return [res] if res else []

        # Processar PDF
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close()
            return imgs

        # Processar imagens normais (PNG, JPG)
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except Exception as e:
        st.warning(f"Erro ao processar item: {e}")
        return []

# --- UI ---
st.title("Automação de Relatórios - UPA Nova Cidade")
st.caption("Versão 0.6.0 - Correção Sistema de Prints")
t_manual, t_evidencia = st.tabs(["📝 Dados", "📁 Evidências"])

with t_manual:
    st.markdown("### Preencha os campos de texto")

    # Linha 1: 2 colunas
    c1, c2 = st.columns(2)
    with c1: st.text_input("Mês de Referência", key="in_mes")
    with c2: st.text_input("Total de Atendimentos", key="in_total")

    # Linha 2: 3 colunas
    c3, c4, c5 = st.columns(3)
    with c3: st.text_input("Médicos Clínicos", key="in_mc")
    with c4: st.text_input("Médicos Pediatras", key="in_mp")
    with c5: st.text_input("Odonto Clínico", key="in_oc")

    # Linha 3: 3 colunas
    c6, c7, c8 = st.columns(3)
    with c6: st.text_input("Odonto Ped", key="in_op")
    with c7: st.text_input("Pacientes CCIH", key="in_ccih")
    with c8: st.text_input("Ouvidoria Interna", key="in_oi")

    # Linha 4: 3 colunas
    c9, c10, c11 = st.columns(3)
    with c9: st.text_input("Ouvidoria Externa", key="in_oe")
    with c10: st.number_input("Total de Transferências", step=1, key="in_tt")
    with c11: st.text_input("Taxa de Transferência (%)", key="in_taxa")

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
                # CORREÇÃO 3: Badge visual mostrando quantidade de anexos
                num_anexos = len(st.session_state.dados_sessao[m])
                badge_html = f'<span class="anexo-badge">{num_anexos}</span>' if num_anexos > 0 else ''
                st.markdown(f"<span class='upload-label'>{labels.get(m, m)}{badge_html}</span>", unsafe_allow_html=True)

                ca, cb = st.columns([1, 1])
                with ca:
                    pasted = paste_image_button(label="Colar Print", key=f"p_{m}_{b_idx}")

                    if pasted is not None:
                        img_pil = getattr(pasted, 'image_data', None)
                        p_time = getattr(pasted, 'time_now', 0)

                        # CORREÇÃO 4: Verificação de timestamp e salvamento correto
                        if img_pil is not None and p_time > st.session_state.ultimo_print_time.get(m, 0):
                            try:
                                buf = io.BytesIO()
                                img_pil.save(buf, format="PNG")
                                b_data = buf.getvalue()
                                nome = f"Print_{len(st.session_state.dados_sessao[m]) + 1}.png"

                                # IMPORTANTE: Salvar como bytes diretamente
                                st.session_state.dados_sessao[m].append({
                                    "name": nome, 
                                    "content": b_data,  # bytes puros
                                    "type": "p"
                                })

                                st.session_state.ultimo_print_time[m] = p_time
                                st.toast(f"✅ Print anexado: {labels[m]}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao processar print: {e}")

                with cb:
                    f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf', 'xlsx'], key=f"f_{m}_{b_idx}", label_visibility="collapsed")
                    if f_up:
                        if f_up.name not in [x['name'] for x in st.session_state.dados_sessao[m]]:
                            st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                            st.rerun()

                # CORREÇÃO 5: Visualização melhorada dos anexos
                if st.session_state.dados_sessao[m]:
                    for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                        tipo_icon = "📸" if item['type'] == "p" else "📎"
                        with st.expander(f"{tipo_icon} {item['name']}", expanded=False):
                            # Visualizar prints e imagens
                            if item['type'] == "p":
                                st.image(item['content'], use_container_width=True)
                            elif not item['name'].lower().endswith(('.pdf', '.xlsx')):
                                st.image(item['content'], use_container_width=True)
                            else:
                                st.info(f"Arquivo: {item['name']}")

                            if st.button("🗑️ Remover", key=f"del_{m}_{i_idx}_{b_idx}"):
                                st.session_state.dados_sessao[m].pop(i_idx)
                                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if st.button("🚀 FINALIZAR E GERAR RELATÓRIO PDF", type="primary", use_container_width=True):
    # Recuperação dos dados via session_state
    mes_ref = st.session_state.get("in_mes", "").strip()

    if not mes_ref:
        st.error("⚠️ Mês de Referência é obrigatório.")
    else:
        try:
            # Cálculos médicos
            mc = int(st.session_state.get("in_mc", 0) or 0)
            mp = int(st.session_state.get("in_mp", 0) or 0)
            total_med = mc + mp

            with tempfile.TemporaryDirectory() as tmp:
                docx_p = os.path.join(tmp, "relatorio.docx")
                doc = DocxTemplate("template.docx")
                with st.spinner("🔧 Construindo relatório..."):
                    dados_finais = {
                        "SISTEMA_MES_REFERENCIA": mes_ref,
                        "ANALISTA_TOTAL_ATENDIMENTOS": st.session_state.get("in_total", ""),
                        "ANALISTA_MEDICO_CLINICO": st.session_state.get("in_mc", ""),
                        "ANALISTA_MEDICO_PEDIATRA": st.session_state.get("in_mp", ""),
                        "ANALISTA_ODONTO_CLINICO": st.session_state.get("in_oc", ""),
                        "ANALISTA_ODONTO_PED": st.session_state.get("in_op", ""),
                        "TOTAL_RAIO_X": st.session_state.get("in_rx", ""),
                        "TOTAL_PACIENTES_CCIH": st.session_state.get("in_ccih", ""),
                        "OUVIDORIA_INTERNA": st.session_state.get("in_oi", ""),
                        "OUVIDORIA_EXTERNA": st.session_state.get("in_oe", ""),
                        "SISTEMA_TOTAL_DE_TRANSFERENCIA": st.session_state.get("in_tt", 0),
                        "SISTEMA_TAXA_DE_TRANSFERENCIA": st.session_state.get("in_taxa", ""),
                        "SISTEMA_TOTAL_MEDICOS": total_med
                    }

                    # CORREÇÃO 6: Processar todos os anexos (prints e uploads)
                    total_imagens = 0
                    for m in DIMENSOES_CAMPOS.keys():
                        imgs = []
                        for item in st.session_state.dados_sessao[m]:
                            # Passa o conteúdo correto (bytes para prints, file object para uploads)
                            res = processar_item_lista(doc, item['content'], m)
                            if res: 
                                imgs.extend(res)
                                total_imagens += len(res)
                        dados_finais[m] = imgs

                    st.info(f"📊 Processando {total_imagens} imagens...")

                    doc.render(dados_finais)
                    doc.save(docx_p)

                    # Converter para PDF
                    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, docx_p], check=True)
                    pdf_final = os.path.join(tmp, "relatorio.pdf")

                    if os.path.exists(pdf_final):
                        # CORREÇÃO 7: Ler arquivo ANTES do download_button (erro crítico anterior)
                        with open(pdf_final, "rb") as f:
                            pdf_bytes = f.read()

                        st.success("✅ Relatório gerado com sucesso!")
                        nome_arquivo = f"Relatorio_{mes_ref.replace('/', '-')}.pdf"
                        st.download_button(
                            "📥 Descarregar PDF", 
                            pdf_bytes,  # Usar variável, não f.read()
                            nome_arquivo, 
                            "application/pdf"
                        )
                    else:
                        st.error("❌ Falha na conversão para PDF.")
        except Exception as e: 
            st.error(f"❌ Erro Crítico: {e}")
            import traceback
            st.code(traceback.format_exc())

st.caption("Desenvolvido por Leonardo Barcelos Martins | V0.6.0")
