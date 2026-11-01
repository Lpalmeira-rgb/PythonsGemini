import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Processador e Padronizador de Mangueiras", layout="centered"
)

st.title("🛠️ Processador e Padronizador de Mangueiras")
st.write(
    "Carregue a sua planilha Excel para realizar a extração, padronização da descrição e estruturação dos dados automaticamente."
)


# 1. Função Robusta para Extração da Bitola
def extrair_bitola(desc):
    if not isinstance(desc, str) or not desc.strip():
        return ""

    text = desc.strip()

    # Normalizações prévias de typos e caracteres especiais
    text_norm = re.sub(r"(\d+);(\d+)", r"\1/\2", text)  # Troca 2-1;2 por 2-1/2
    text_norm = text_norm.replace("”", '"').replace("''", '"').replace("`", "'")

    # Limpa nomes de especificações/normas para evitar falsos positivos
    text_clean = re.sub(
        r"\bSAE\s*100R(?:12|19|1|2|4|7)(?:[A-Z]*)\b",
        "SAE_SPEC",
        text_norm,
        flags=re.IGNORECASE,
    )
    text_clean = re.sub(
        r"\b006-\d+-\d+\b", "PART_SPEC", text_clean, flags=re.IGNORECASE
    )

    # Frações compostas (ex: 1 1/4", 1" 1/4, 2-1/2", 1.1/4")
    m_mixed = re.search(
        r'(\d+)[\s\-\;\.]+(\d+\/\d+)\s*(?:"|\'|in|ID|pol)?',
        text_clean,
        re.IGNORECASE,
    )
    if m_mixed:
        return f'{m_mixed.group(1)}.{m_mixed.group(2)}"'

    m_inch_frac = re.search(
        r'(\d+)\s*(?:"|\')?\s*(\d+\/\d+)', text_clean, re.IGNORECASE
    )
    if m_inch_frac:
        return f'{m_inch_frac.group(1)}.{m_inch_frac.group(2)}"'

    # Frações simples (ex: 1/4", 1/4', 3/8", 1/2', 3/16")
    m_frac = re.search(
        r'(\d+\/\d+)\s*(?:"|\'|in|ID|pol)?', text_clean, re.IGNORECASE,
    )
    if m_frac:
        return f'{m_frac.group(1)}"'

    # Notação em milímetros (ex: 12,70mm -> 1/2", 31,75mm -> 1.1/4")
    m_mm = re.search(r"(\d+(?:[.,]\d+)?)\s*mm", text, re.IGNORECASE)
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

    # Dash sizes anotados com aspas (ex: 4', 4", 6', 6", 8', 8")
    m_dash_q = re.search(r"\b(4|6|8|10|12|16)\s*(?:\"|')\b", text_clean)
    if m_dash_q and m_dash_q.group(1) in dash_map:
        return dash_map[m_dash_q.group(1)]

    # Códigos JIC / # / - (ex: JIC 4, JIC 04, -4 JIC, #12, GH781-6)
    for m in re.finditer(
        r"(?:JIC\s*|#|,\s*-|-\s*|^-\s*)(\d{1,2})\b", text_clean, re.IGNORECASE
    ):
        if m.group(1) in dash_map:
            return dash_map[m.group(1)]

    # Polegadas inteiras (ex: 1", 1', 2", 3")
    m_whole = re.search(
        r'(?:Bitola:\s*|ID\s*|x\s*|\b|^)([1-3])\s*(?:"|\'|in|pol)\b',
        text_clean,
        re.IGNORECASE,
    )
    if m_whole:
        return f'{m_whole.group(1)}"'

    # Notação DN (DN 6 -> 1/4", DN 12 -> 1/2", DN 19 -> 3/4", DN 25 -> 1")
    m_dn = re.search(r"\bDN\s*(\d+)", text, re.IGNORECASE)
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


# 2. Extração de Pressão de Trabalho
def extrair_pressao(desc):
    if not isinstance(desc, str):
        return ""
    m = re.search(
        r"(\d+(?:\.\d+)?\s*(?:PSI|BAR|psi|bar)(?:\s*\/\s*\d+(?:\.\d+)?\s*(?:PSI|BAR|psi|bar))?)",
        desc,
    )
    return m.group(1).upper() if m else ""


# 3. Extração de Comprimento em mm
def extrair_comprimento_mm(row):
    comp = (
        str(row["Comprimento"])
        if "Comprimento" in row and pd.notna(row["Comprimento"])
        else ""
    )
    desc = (
        str(row["Descrição"])
        if "Descrição" in row and pd.notna(row["Descrição"])
        else ""
    )
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


# 4. Mapeamento de Conexões Conforme o Padrão da Imagem
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


# Interface do Streamlit
uploaded_file = st.file_uploader(
    "Escolha a planilha Excel (`.xlsx`)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # Preserva texto original da Descrição para análise de dados
    desc_original = df["Descrição"].astype(str)

    # 1. Extrai a bitola
    df["Bitola"] = desc_original.apply(extrair_bitola)

    # 2. Extrai/Consolida a Pressão de Trabalho
    if "Pressão de Trabalho" in df.columns:
        df["Pressão de Trabalho"] = df.apply(
            lambda r: str(r["Pressão de Trabalho"]).strip()
            if pd.notna(r["Pressão de Trabalho"])
            and str(r["Pressão de Trabalho"]).strip() != "nan"
            else extrair_pressao(r["Descrição"]),
            axis=1,
        )
    else:
        df["Pressão de Trabalho"] = desc_original.apply(extrair_pressao)

    # 3. Extrai Comprimento e Conexões
    df["Comprimento"] = df.apply(extrair_comprimento_mm, axis=1)

    conexoes = desc_original.apply(extrair_conexao)
    df["Conexão 1"] = [c[0] for c in conexoes]
    df["Conexão 2"] = [c[1] for c in conexoes]

    # 4. Padroniza a própria coluna "Descrição" no formato: Mangueira + Bitola + Pressão
    def padronizar_descricao_final(row):
        bitola = row["Bitola"]
        pressao = row["Pressão de Trabalho"]

        parts = ["Mangueira"]
        if bitola:
            parts.append(bitola)
        if pressao:
            p_str = str(pressao).strip()
            if not re.search(r"PSI|BAR", p_str, re.IGNORECASE) and p_str != "":
                p_str += " PSI"
            parts.append(p_str)

        # Se não houver bitola nem pressão identificadas, mantém a descrição original
        if len(parts) == 1:
            return row["Descrição"]

        return " ".join(parts)

    df["Descrição"] = df.apply(padronizar_descricao_final, axis=1)

    st.success("Planilha processada e Descrição padronizada com sucesso!")
    st.dataframe(df.head(15))

    # Download do arquivo Excel
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
