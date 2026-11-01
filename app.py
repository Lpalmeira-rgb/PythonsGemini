import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Processador e Padronizador de Mangueiras", layout="centered"
)

st.title("🛠️ Processador e Padronizador de Mangueiras")
st.write(
    "Carregue a sua planilha Excel para realizar a limpeza, conversão de pressão, padronização de conexões e descrições."
)


# 1. Função Robusta de Extração e Padronização da Bitola
def padronizar_bitola(val_existente, desc):
    texto_base = (
        str(val_existente).strip()
        if pd.notna(val_existente) and str(val_existente).strip() not in ["", "nan"]
        else str(desc).strip()
    )

    if not texto_base or texto_base == "nan":
        return ""

    text_norm = re.sub(r"(\d+);(\d+)", r"\1/\2", texto_base)
    text_norm = text_norm.replace("”", '"').replace("''", '"').replace("`", "'")

    text_clean = re.sub(
        r"\bSAE\s*100R(?:12|19|1|2|4|7)(?:[A-Z]*)\b",
        "SAE_SPEC",
        text_norm,
        flags=re.IGNORECASE,
    )
    text_clean = re.sub(
        r"\b006-\d+-\d+\b", "PART_SPEC", text_clean, flags=re.IGNORECASE
    )

    m_mixed = re.search(
        r'(\d+)[\s\-\;\.]+(\d+\/\d+)\s*(?:"|\'|in|ID|pol)?',
        text_clean,
        re.IGNORECASE,
    )
    if m_mixed:
        return f'{m_mixed.group(1)}.{m_mixed.group(2)}"'

    m_frac = re.search(
        r'(\d+\/\d+)\s*(?:"|\'|in|ID|pol)?', text_clean, re.IGNORECASE
    )
    if m_frac:
        return f'{m_frac.group(1)}"'

    m_mm = re.search(r"(\d+(?:[.,]\d+)?)\s*mm", texto_base, re.IGNORECASE)
    if m_mm:
        val_mm = float(m_mm.group(1).replace(",", "."))
        mm_map = {
            6.35: '1/4"',
            9.52: '3/8"',
            12.7: '1/2"',
            15.87: '5/8"',
            19.05: '3/4"',
            25.4: '1"',
            31.75: '1.1/4"',
        }
        for k, v in mm_map.items():
            if abs(val_mm - k) < 0.5:
                return v

    dash_map = {
        "4": '1/4"',
        "04": '1/4"',
        "6": '3/8"',
        "06": '3/8"',
        "8": '1/2"',
        "08": '1/2"',
        "10": '5/8"',
        "12": '3/4"',
        "16": '1"',
        "20": '1.1/4"',
    }

    m_dash_q = re.search(r"\b(4|6|8|10|12|16)\s*(?:\"|')\b", text_clean)
    if m_dash_q and m_dash_q.group(1) in dash_map:
        return dash_map[m_dash_q.group(1)]

    for m in re.finditer(
        r"(?:JIC\s*|#|,\s*-|-\s*|^-\s*)(\d{1,2})\b", text_clean, re.IGNORECASE
    ):
        if m.group(1) in dash_map:
            return dash_map[m.group(1)]

    m_whole = re.search(
        r'(?:Bitola:\s*|ID\s*|x\s*|\b|^)([1-3])\s*(?:"|\'|in|pol)\b',
        text_clean,
        re.IGNORECASE,
    )
    if m_whole:
        return f'{m_whole.group(1)}"'

    m_dn = re.search(r"\bDN\s*(\d+)", texto_base, re.IGNORECASE)
    if m_dn:
        dn_map = {
            "6": '1/4"',
            "10": '3/8"',
            "12": '1/2"',
            "19": '3/4"',
            "25": '1"',
        }
        if m_dn.group(1) in dn_map:
            return dn_map[m_dn.group(1)]

    return ""


# 2. Função de Padronização e Conversão de Pressão de Trabalho
def padronizar_pressao(val_existente, desc):
    val_str = (
        str(val_existente).strip()
        if pd.notna(val_existente)
        and str(val_existente).strip() not in ["", "nan", "-", "???????????"]
        else ""
    )

    if not val_str:
        m_desc = re.search(
            r"(\d+(?:\.\d+)?\s*(?:PSI|BAR|psi|bar)(?:\s*\/\s*\d+(?:\.\d+)?\s*(?:PSI|BAR|psi|bar))?)",
            str(desc),
        )
        val_str = m_desc.group(1) if m_desc else ""

    if not val_str:
        return ""

    m_psi = re.search(r"(\d+(?:\.\d+)?)\s*PSI", val_str, re.IGNORECASE)
    if m_psi:
        num = float(m_psi.group(1))
        num_int = int(round(num))
        if num_int % 100 != 0 and abs(num_int - round(num_int, -2)) < 15:
            num_int = int(round(num_int, -2))
        return f"{num_int} PSI"

    m_bar = re.search(r"(\d+(?:\.\d+)?)\s*BAR", val_str, re.IGNORECASE)
    if m_bar:
        num_bar = float(m_bar.group(1))
        num_psi = num_bar * 14.50377
        num_int = (
            int(round(num_psi, -2))
            if num_psi > 1000
            else int(round(num_psi))
        )
        return f"{num_int} PSI"

    m_num = re.search(r"(\d+(?:\.\d+)?)", val_str)
    if m_num:
        num = float(m_num.group(1))
        num_int = int(round(num))
        if num_int % 100 != 0 and abs(num_int - round(num_int, -2)) < 15:
            num_int = int(round(num_int, -2))
        return f"{num_int} PSI"

    return val_str


# 3. Função de Padronização de Conexão (Considerando Reta quando sem ângulo)
def padronizar_conexao(val_existente, desc, bitola):
    texto = (
        str(val_existente).strip()
        if pd.notna(val_existente) and str(val_existente).strip() not in ["", "nan"]
        else str(desc).strip()
    )
    if not texto or texto == "nan":
        return ""

    texto_upper = texto.upper()
    has_90 = bool(re.search(r"90|CURVA|ANGULO|ÂNGULO", texto_upper))

    if "20 ORFS" in texto_upper or "200FOI" in texto_upper:
        return "20 ORFS"
    elif "BSP" in texto_upper:
        return 'BSP 1/2"'
    elif "FG" in texto_upper:
        return 'FG 1" INOX'
    elif "FLANGE" in texto_upper:
        return "Flange SAE"
    elif "M9" in texto_upper:
        return "M9 Fêmea" if "FÊMEA" in texto_upper else "M9 Macho"
    elif "ABRAÇADEIRA" in texto_upper or "SUCÇÃO" in texto_upper:
        return "Abraçadeira"

    jic_num = ""
    m_jic = re.search(r"JIC\s*0?(\d{1,2})", texto_upper)
    if m_jic:
        jic_num = m_jic.group(1)
    elif bitola:
        bit_map = {
            '1/4"': "04",
            '3/8"': "06",
            '1/2"': "08",
            '5/8"': "10",
            '3/4"': "12",
            '1"': "16",
            '1.1/4"': "20",
        }
        jic_num = bit_map.get(bitola, "")

    if jic_num:
        jic_formatted = f"JIC {int(jic_num):02d}"
        if has_90 and jic_formatted in ["JIC 04", "JIC 08", "JIC 12"]:
            return f"{jic_formatted} Fêmea 90°"
        else:
            return f"{jic_formatted} Fêmea Reta"

    return texto


# 4. Função de Comprimento
def padronizar_comprimento(val_existente, desc):
    comp = (
        str(val_existente).strip()
        if pd.notna(val_existente) and str(val_existente).strip() not in ["", "nan"]
        else ""
    )
    desc_str = str(desc).strip() if pd.notna(desc) else ""
    texto = f"{comp} {desc_str}"

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

    return comp


# Interface de Upload
uploaded_file = st.file_uploader(
    "Escolha a planilha Excel (`.xlsx`)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # CORREÇÃO DO ERRO: Garantir que todas as colunas aceitem strings sem conflito de dtypes (float64 -> object)
    colunas_alvo = [
        "Bitola",
        "Comprimento",
        "Conexão 1",
        "Conexão 2",
        "Pressão de Trabalho",
        "Descrição",
    ]
    for col in colunas_alvo:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(object)

    # Processamento linha a linha
    for idx, row in df.iterrows():
        desc_orig = row["Descrição"] if pd.notna(row["Descrição"]) else ""

        bitola_val = padronizar_bitola(row.get("Bitola", ""), desc_orig)
        pressao_val = padronizar_pressao(
            row.get("Pressão de Trabalho", ""), desc_orig
        )
        comp_val = padronizar_comprimento(row.get("Comprimento", ""), desc_orig)
        c1_val = padronizar_conexao(
            row.get("Conexão 1", ""), desc_orig, bitola_val
        )
        c2_val = padronizar_conexao(
            row.get("Conexão 2", ""), desc_orig, bitola_val
        )

        df.at[idx, "Bitola"] = bitola_val
        df.at[idx, "Pressão de Trabalho"] = pressao_val
        df.at[idx, "Comprimento"] = comp_val
        df.at[idx, "Conexão 1"] = c1_val
        df.at[idx, "Conexão 2"] = c2_val

        parts = ["Mangueira"]
        if bitola_val:
            parts.append(bitola_val)
        if pressao_val:
            parts.append(pressao_val)

        if len(parts) > 1:
            df.at[idx, "Descrição"] = " ".join(parts)

    st.success("Planilha padronizada e processada com sucesso!")
    st.dataframe(df.head(15))

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados_Padronizados")
    buffer.seek(0)

    st.download_button(
        label="📥 Baixar Planilha Padronizada",
        data=buffer,
        file_name="Planilha_Padronizada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
