import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.cell.cell import MergedCell
from datetime import datetime, timedelta
from io import BytesIO
import zipfile
import tempfile
import os
import re
import unicodedata

# =========================================
# CONFIG
# =========================================
st.set_page_config(page_title="Consolidador Gerencial CN", layout="wide")

COLUNA_L = 12

ORDEM_ABAS = ["COB", "CUI", "EDE", "NOB", "PVE", "SOB", "XAM"]

# =========================================
# UTIL
# =========================================
def normalizar(texto):
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ASCII", "ignore").decode("ASCII")
    return texto.lower()

# =========================================
# IDENTIFICAR TIPO DE ABA
# =========================================
def identificar_tipo_aba(nome_arquivo):

    nome = normalizar(nome_arquivo)

    if "cuiaba" in nome:
        return "CUI"
    if "ede" in nome:
        return "EDE"
    if "sobradinho" in nome:
        return "SOB"
    if "gerencial_" in nome:
        return "XAM"
    if re.search(r"gerencial \\d{2}-\\d{2}-\\d{4}", nome):
        return "COB"
    if re.search(r"gerencial \\d{1,2}", nome):
        return "PVE"
    if "resumo gerencial" in nome:
        return "NOB"

    return "OUTRO"

# =========================================
# LIMPAR COLUNA L (CORRIGIDO PVE)
# =========================================
def limpar_coluna_l(aba):

    if aba.max_column < COLUNA_L:
        return

    # limpar mescladas
    for intervalo in list(aba.merged_cells.ranges):
        if intervalo.min_col >= COLUNA_L:
            celula = aba.cell(intervalo.min_row, intervalo.min_col)
            if not isinstance(celula, MergedCell):
                celula.value = None

    # limpar normal
    for row in aba.iter_rows(
        min_row=1,
        max_row=aba.max_row,
        min_col=COLUNA_L,
        max_col=aba.max_column
    ):
        for c in row:
            if isinstance(c, MergedCell):
                continue
            c.value = None
            c.comment = None
            c._hyperlink = None

# =========================================
# LAYOUT
# =========================================
def aplicar_layout(aba):
    aba.freeze_panes = "A5"
    aba.sheet_view.showGridLines = False
    aba.sheet_view.zoomScale = 80

# =========================================
# COPIAR ABA
# =========================================
def copiar_aba(origem, destino):

    for row in origem.iter_rows():
        for c in row:
            destino[c.coordinate].value = c.value

    for merged in origem.merged_cells.ranges:
        destino.merge_cells(str(merged))

# =========================================
# ESCOLHER ABA CORRETA
# =========================================
def escolher_aba(wb, data):

    for nome in wb.sheetnames:
        if nome.isdigit() and int(nome) == data.day:
            return wb[nome]

    for nome in wb.sheetnames:
        if normalizar(nome) == "gerencial":
            return wb[nome]

    return None

# =========================================
# ORDENAR ABAS
# =========================================
def ordenar_abas(wb):

    def prioridade(nome):
        nome_upper = nome.upper()

        for i, prefixo in enumerate(ORDEM_ABAS):
            if prefixo in nome_upper:
                return i

        return 99

    wb._sheets = sorted(wb.worksheets, key=lambda aba: prioridade(aba.title))

# =========================================
# PROCESSAR UMA DATA
# =========================================
def processar_uma_data(arquivos, data):

    wb_final = Workbook()
    wb_final.remove(wb_final.active)

    for arq in arquivos:

        wb = load_workbook(arq["caminho"], data_only=True)

        aba_origem = escolher_aba(wb, data)

        if aba_origem is None:
            wb.close()
            continue

        tipo = identificar_tipo_aba(arq["nome"])

        nome_aba = tipo

        # evitar duplicadas
        contador = 1
        nome_final = nome_aba

        while nome_final in wb_final.sheetnames:
            nome_final = f"{nome_aba}_{contador}"
            contador += 1

        aba_destino = wb_final.create_sheet(nome_final)

        copiar_aba(aba_origem, aba_destino)

        limpar_coluna_l(aba_destino)
        aplicar_layout(aba_destino)

        wb.close()

    # reforço final (garante PVE)
    for aba in wb_final.worksheets:
        limpar_coluna_l(aba)
        aplicar_layout(aba)

    # ordenar!!!
    ordenar_abas(wb_final)

    # salvar
    output = BytesIO()
    wb_final.save(output)
    output.seek(0)

    return {
        "arquivo_excel": output,
        "nome": f"Resumo Gerencial CN - {data.strftime('%d.%m.%Y')}.xlsx"
    }

# =========================================
# MULTI DATA
# =========================================
def processar_multiplas_datas(arquivos, datas):

    resultados = []

    for item in datas:
        data = item["data"]
        resultados.append(processar_uma_data(arquivos, data))

    return resultados

# =========================================
# ZIP
# =========================================
def gerar_zip(resultados):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for r in resultados:
            z.writestr(r["nome"], r["arquivo_excel"].getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# =========================================
# UI
# =========================================
st.title("📊 Consolidador Gerencial CN")

# UPLOAD
arquivos_upload = st.file_uploader(
    "Envie todos os arquivos Excel",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

arquivos_salvos = []

if arquivos_upload:

    pasta = tempfile.mkdtemp()

    for f in arquivos_upload:
        caminho = os.path.join(pasta, f.name)

        with open(caminho, "wb") as file:
            file.write(f.getvalue())

        arquivos_salvos.append({
            "nome": f.name,
            "caminho": caminho
        })

    st.success(f"{len(arquivos_salvos)} arquivos carregados")

# DATAS
hoje = datetime.today()
data_base = hoje - timedelta(days=1)

datas = []

for i in range(1, 32):
    try:
        d = datetime(data_base.year, data_base.month, i)
        datas.append({
            "label": d.strftime("%d/%m/%Y"),
            "data": d
        })
    except:
        pass

datas_selecionadas = st.multiselect(
    "Selecione uma ou mais datas",
    datas,
    format_func=lambda x: x["label"]
)

# PROCESSAR
if st.button("🚀 Processar"):

    if not arquivos_salvos:
        st.warning("Envie arquivos primeiro")
        st.stop()

    if not datas_selecionadas:
        st.warning("Selecione pelo menos uma data")
        st.stop()

    if len(datas_selecionadas) > 5:
        st.warning("Máximo de 5 datas por execução")
        st.stop()

    with st.spinner("Processando..."):

        resultados = processar_multiplas_datas(
            arquivos_salvos,
            datas_selecionadas
        )

    st.success("✅ Arquivos gerados")

    # ZIP
    zip_file = gerar_zip(resultados)

    st.download_button(
        "⬇️ Baixar TODOS (ZIP)",
        zip_file,
        "consolidados.zip"
    )

    # individuais
    st.subheader("Downloads individuais")

    for r in resultados:
        st.download_button(
            r["nome"],
            r["arquivo_excel"],
            r["nome"]
        )
``
