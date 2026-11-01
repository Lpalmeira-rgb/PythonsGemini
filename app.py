import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tratamento de Folha de Cálculo", layout="centered"
)

st.title("🛠️ Processador e Padronizador de Mangueiras")
st.write(
    "Carregue o seu ficheiro Excel (`Pasta1.xlsx`) para discriminar e padronizar as colunas automaticamente."
)


# Funções de extração de dados
def extrair_bitola(desc):
    if not isinstance(desc, str):
        return ""
    m_frac = re.search(
        r'(\d+(?:\.\d+)?(?:\/\d+)?(?:\s*\d+\/\d+)?)\s*(?:"|\'\'|”|in|ID|pol)',
        desc,
        re.IGNORECASE,
    )
    if m_frac:
        return f'{m_frac.group(1).strip()}"'
    m_dash = re.search(r"-(\d{1,2})\s+(?:JIC|HOSE)", desc, re.IGNORECASE)
    if m_dash:
        dash_map = {
            "4": '1/4"',
            "6": '3/8"',
            "8": '1/2"',
            "10": '5/8"',
            "12": '3/4"',
            "16": '1"',
            "20": '1.1/4"',
        }
        return dash_map.get(m_dash.group(1), "")
    m_dn = re.search(r"DN\s*(\d+)", desc, re.IGNORECASE)
    if m_dn:
        dn_map = {
            "6": '1/4"',
            "10": '3/8"',
            "12": '1/2"',
            "19": '3/4"',
            "25": '1"',
        }
        return dn_map.get(m_dn.group(1), "")
    return ""


def extrair_comprimento_mm(row):
    comp = str(row["Comprimento"]) if pd.notna(row["Comprimento"]) else ""
    desc = str(row["Descrição"]) if pd.notna(row["Descrição"]) else ""
    texto = f"{comp} {desc}"

    m_m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:metram|metros|metro|M|m)\b", texto)
    if m_m:
        val = float(m_m.group(1).replace(",", "."))
        return f"{int(round(val * 1000))} mm"

    m_mm = re.search(r"(\d+(?:[.,]\d+)?)\s*mm\b", texto, re.IGNORECASE)
    if m_mm:
        val = float(m_mm.group(1).replace(",", "."))
        return f"{int(round(val))} mm"

    m_lg = re.search(r"(\d+(?:[.,]\d+)?)\s*LG\b", texto, re.IGNORECASE)
    if m_lg:
        val = float(m_lg.group(1).replace(",", "."))
        return f"{int(round(val * 25.4))} mm"

    return ""


def extrair_conexao(desc):
    if not isinstance(desc, str):
        return "", ""
    c = ""
    if re.search(r"JIC\s*0?4", desc, re.IGNORECASE) or "-4 JIC" in desc:
        c = "JIC 04 Fêmea 90°" if "90" in desc else "JIC 04 Fêmea Reta"
    elif re.search(r"JIC\s*0?6", desc, re.IGNORECASE) or "-6 JIC" in desc:
        c = "JIC 06 Fêmea Reta"
    elif re.search(r"JIC\s*0?8", desc, re.IGNORECASE) or "-8 JIC" in desc:
        c = "JIC 08 Fêmea 90°" if "90" in desc else "JIC 08 Fêmea Reta"
    elif re.search(r"JIC\s*10", desc, re.IGNORECASE) or "-10 JIC" in desc:
        c = "JIC 10 Fêmea Reta"
    elif re.search(r"JIC\s*12", desc, re.IGNORECASE) or "-12 JIC" in desc:
        c = "JIC 12 Fêmea 90°" if "90" in desc else "JIC 12 Fêmea Reta"
    elif re.search(r"JIC\s*16", desc, re.IGNORECASE) or "-16 JIC" in desc:
        c = "JIC 16 Fêmea Reta"
    elif re.search(r"JIC\s*20", desc, re.IGNORECASE) or "-20 JIC" in desc:
        c = "JIC 20 Fêmea Reta"
    elif "20 ORFS" in desc.upper() or "200FOI" in desc.upper():
        c = "20 ORFS"
    elif "BSP" in desc.upper():
        c = 'BSP 1/2"'
    elif "FG 1" in desc.upper() or "FG" in desc.upper():
        c = 'FG 1" INOX'
    elif "FLANGE" in desc.upper():
        c = "Flange SAE"
    elif "M9" in desc.upper():
        c = "M9 Fêmea" if "FÊMEA" in desc.upper() else "M9 Macho"
    elif "ABRAÇADEIRA" in desc.upper() or "SUCÇÃO" in desc.upper():
        c = "Abraçadeira"
    return c, c


def extrair_pressao(desc):
    if not isinstance(desc, str):
        return ""
    m = re.search(
        r"(\d+\s*(?:PSI|BAR|psi|bar)(?:\s*\/\s*\d+\s*(?:PSI|BAR|psi|bar))?)",
        desc,
    )
    return m.group(1).upper() if m else ""


# Interface de Upload
uploaded_file = st.file_uploader(
    "Escolha o ficheiro Excel", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # Aplicar transformações
    df["Bitola"] = df["Descrição"].apply(extrair_bitola)
    df["Comprimento"] = df.apply(extrair_comprimento_mm, axis=1)

    conexoes = df["Descrição"].apply(extrair_conexao)
    df["Conexão 1"] = [c[0] for c in conexoes]
    df["Conexão 2"] = [c[1] for c in conexoes]
    df["Pressão de Trabalho"] = df["Descrição"].apply(extrair_pressao)

    st.success("Ficheiro processado com sucesso!")
    st.dataframe(df.head(10))

    # Gerar ficheiro para download na memória
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados_Tratados")
    buffer.seek(0)

    st.download_button(
        label="📥 Descarregar Folha de Cálculo Tratada",
        data=buffer,
        file_name="Pasta1_Processada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
