import pandas as pd
import streamlit as st

st.set_page_config(page_title="Calculadora de Emissões e Redução de CO₂", layout="centered")
st.title("🔍 Calculadora de Emissões e Redução de CO₂ com Hidrogênio")

st.markdown("""
Este aplicativo lê a planilha **Calculos.xlsx** e calcula:
- Emissão do combustível original (kg CO₂);
- Quantidade de H₂ necessária para substituir;
- Emissões do H₂ por rota de produção (**Eletrólise**, **Biomassa**, **SMR**, **SMR+CCS**);
- Redução de CO₂ em cada cenário.

> Compatível com dois formatos de planilha:
> 1) **Formato "arrumado" (tidy)** com colunas: `Combustível`, `Fator_CO2 (kg/kg)`, `H2_equivalente (kg/kg)`
> 2) **Formato "matricial" (como no seu Excel)**: usa índices de linha/coluna configuráveis abaixo.
""")

# ---------------------- Parâmetros configuráveis p/ formato matricial ----------------------
with st.expander("⚙️ Opções avançadas para planilhas 'matriciais' (opcional)"):
    st.write("Use estes campos apenas se sua planilha **não** tiver colunas nomeadas (formato tidy).")
    linha_qtd_comb_default = 0   # linha com quantidade base do combustível (ex.: 1 kg)
    linha_h2_equiv_default = 2   # linha com 'kg H2 por kg de combustível'
    linha_emissao_prefix_default = 4  # linha onde estão as emissões de referência (uma coluna por combustível), com cabeçalhos tipo 'Emissão Gás Natural'

    linha_qtd_comb = st.number_input("Linha da quantidade base do combustível", min_value=0, value=linha_qtd_comb_default, step=1)
    linha_h2_equiv = st.number_input("Linha do fator H2 equivalente (kg H2/kg combustível)", min_value=0, value=linha_h2_equiv_default, step=1)
    linha_emissao_prefix = st.number_input("Linha das emissões de referência", min_value=0, value=linha_emissao_prefix_default, step=1)
    prefixo_emissao = st.text_input("Prefixo do cabeçalho de emissão", value="Emissão ")

# ---------------------- Entrada do usuário ----------------------
arquivo = st.file_uploader("📥 Envie a sua planilha Calculos.xlsx (ou deixe vazia para usar o arquivo local).", type=["xlsx"])

@st.cache_data
def carregar_excel(conteudo: bytes | None):
    if conteudo is None:
        # Ler arquivo local
        return pd.read_excel("Calculos.xlsx", header=None), True
    else:
        # Ler upload (tentamos sem header para cobrir o caso matricial)
        return pd.read_excel(io.BytesIO(conteudo), header=None), False

df_raw, usando_arquivo_local = carregar_excel(arquivo.read() if arquivo is not None else None)

st.caption(f"Planilha carregada ({'local' if usando_arquivo_local else 'upload do usuário'}). Mostrando prévia (8 primeiras linhas):")
try:
    st.dataframe(df_raw.head(8))
except Exception:
    st.write(df_raw.head(8))

st.divider()

# ---------------------- Escolha de combustível e quantidade ----------------------
combustiveis_sugeridos = ["Gás Natural", "Óleo Combustível", "Carvão"]
combustivel = st.selectbox("Selecione o combustível:", combustiveis_sugeridos)
quantidade = st.number_input("Quantidade de combustível (kg):", min_value=0.0, value=1000.0, step=10.0)

# ---------------------- Detectar formato tidy ----------------------
def tentar_formato_tidy(df_any: pd.DataFrame) -> pd.DataFrame | None:
    # Procurar por cabeçalhos em alguma linha
    for header_row in range(min(5, len(df_any))):
        df_try = pd.read_excel("Calculos.xlsx", header=header_row) if usando_arquivo_local else pd.read_excel(io.BytesIO(arquivo.read()), header=header_row)
        cols = [str(c).strip() for c in df_try.columns]
        if ("Combustível" in cols and 
            any("Fator_CO2" in c or "Emissao_CO2" in c for c in cols) and 
            any("H2_equivalente" in c for c in cols)):
            return df_try
    return None

df_tidy = None
try:
    df_tidy = tentar_formato_tidy(df_raw)
except Exception:
    df_tidy = None

# ---------------------- Funções de extração ----------------------
def extrair_por_tidy(df_tidy: pd.DataFrame, combustivel_escolhido: str):
    # Normalizar nomes de colunas esperados
    cols = {c: str(c) for c in df_tidy.columns}
    col_comb = [c for c in df_tidy.columns if str(c).strip() == "Combustível"][0]
    col_fe = [c for c in df_tidy.columns if "Fator_CO2" in str(c) or "Emissao_CO2" in str(c)][0]
    col_h2 = [c for c in df_tidy.columns if "H2_equivalente" in str(c)][0]
    linha = df_tidy[df_tidy[col_comb].astype(str).str.strip().str.lower() == combustivel_escolhido.lower()].iloc[0]
    fator_emissao = float(linha[col_fe])
    fator_h2eq = float(linha[col_h2])
    qtd_base = 1.0
    return qtd_base, fator_h2eq, fator_emissao

def extrair_por_matriz(df_mat: pd.DataFrame, combustivel_escolhido: str, 
                       linha_qtd: int, linha_h2: int, linha_em_ref: int, prefixo_emissao: str):
    # Assumimos que a linha 0 tem rótulos na coluna 0 e os combustíveis aparecem como cabeçalhos em alguma linha superior.
    # Procurar a linha onde aparecem os cabeçalhos (combustíveis).
    header_row = None
    for r in range(0, min(5, len(df_mat))):
        valores = [str(v).strip() for v in df_mat.iloc[r].tolist()]
        if any(c.lower() in ["gás natural", "gas natural", "óleo combustível", "oleo combustivel", "carvão", "carvao"] for c in valores):
            header_row = r
            break
    if header_row is None:
        header_row = 0
    headers = df_mat.iloc[header_row].tolist()
    # Mapear índice da coluna do combustível
    def idx_comb(nome):
        for i, h in enumerate(headers):
            if str(h).strip().lower() == nome.lower():
                return i
        # tentar variações simples (acentos)
        mapa = {
            "gas natural": "gás natural",
            "oleo combustivel": "óleo combustível",
            "carvao": "carvão"
        }
        alvo = mapa.get(nome.lower(), nome.lower())
        for i, h in enumerate(headers):
            if str(h).strip().lower() == alvo:
                return i
        return None

    col_idx = idx_comb(combustivel_escolhido)
    if col_idx is None:
        raise ValueError("Não encontrei a coluna do combustível na planilha. Verifique os cabeçalhos.")

    # Quantidade base (kg)
    qtd_base = float(str(df_mat.iat[linha_qtd, col_idx]).replace(",", ".").strip())

    # Fator H2 equivalente (kg H2/kg comb)
    fator_h2eq = float(str(df_mat.iat[linha_h2, col_idx]).replace(",", ".").strip())

    # Emissão de referência: procurar coluna com título "Emissão <combustível>"
    # Procurar na linha 'linha_em_ref' por colunas cujo cabeçalho, algumas linhas acima, contenham esse prefixo
    # Alternativa: na própria linha_em_ref as células podem conter o valor; vamos tentar direto:
    # Buscar índice da coluna cuja célula superior tenha o texto "Emissão <Combustível>" em alguma linha
    # Estratégia simples: se existir uma coluna com exatamente o título "Emissão <Combustível>", use-a;
    # senão, use o mesmo col_idx da matriz base.
    col_em_ref = None
    for j, h in enumerate(headers):
        if str(h).strip().lower() == f"{prefixo_emissao}{combustivel_escolhido}".strip().lower():
            col_em_ref = j
            break
    if col_em_ref is None:
        col_em_ref = col_idx  # fallback

    fator_emissao = float(str(df_mat.iat[linha_em_ref, col_em_ref]).replace(",", ".").strip()) / max(qtd_base, 1e-9)
    return qtd_base, fator_h2eq, fator_emissao

# ---------------------- Extrair fatores ----------------------
try:
    if df_tidy is not None:
        qtd_base, fator_h2eq, fator_emissao = extrair_por_tidy(df_tidy, combustivel)
    else:
        qtd_base, fator_h2eq, fator_emissao = extrair_por_matriz(df_raw, combustivel, linha_qtd_comb, linha_h2_equiv, linha_emissao_prefix, prefixo_emissao)
    st.success("Fatores obtidos com sucesso a partir da planilha.")
except Exception as e:
    st.error(f"Não foi possível extrair os fatores automaticamente. Motivo: {e}")
    st.stop()

# ---------------------- Entrada dos fatores de H2 (editáveis) ----------------------
st.subheader("Fatores de emissão do H₂ (kg CO₂ por kg H₂ produzido)")
c1, c2 = st.columns(2)
with c1:
    ef_elet = st.number_input("Eletrólise", min_value=0.0, value=9.97, step=0.1)
    ef_bio  = st.number_input("Biomassa", min_value=-100.0, value=2.5, step=0.1, help="Pode ser negativo em cenários com captura (BECCS).")
with c2:
    ef_smr  = st.number_input("Reforma a Vapor (SMR)", min_value=0.0, value=10.0, step=0.1)
    ef_smr_ccs = st.number_input("SMR com CCS", min_value=0.0, value=3.0, step=0.1)

st.divider()

# ---------------------- Cálculos ----------------------
emissao_original_total = quantidade * fator_emissao
h2_necessario = quantidade * fator_h2eq

em_h2 = {
    "H₂ por Eletrólise": h2_necessario * ef_elet,
    "H₂ por Biomassa":   h2_necessario * ef_bio,
    "H₂ por SMR":        h2_necessario * ef_smr,
    "H₂ por SMR+CCS":    h2_necessario * ef_smr_ccs,
}
reduc = {k: emissao_original_total - v for k, v in em_h2.items()}

st.subheader("Resultados")
st.write(f"**Combustível:** {combustivel}")
st.write(f"**Quantidade (kg):** {quantidade:,.2f}")
st.write(f"**Emissão de referência:** {emissao_original_total:,.2f} kg CO₂")
st.write(f"**H₂ necessário para substituir:** {h2_necessario:,.2f} kg")

df_out = pd.DataFrame({
    "Cenário": list(em_h2.keys()),
    "Emissão com H₂ (kg CO₂)": list(em_h2.values()),
    "Redução vs. referência (kg CO₂)": list(reduc.values())
})
st.dataframe(df_out, use_container_width=True)

st.bar_chart(df_out.set_index("Cenário")["Emissão com H₂ (kg CO₂)"])

# ---------------------- Download ----------------------
st.download_button(
    "📥 Baixar resultados (CSV)",
    data=df_out.to_csv(index=False).encode("utf-8"),
    file_name="resultados_substituicao_h2.csv",
    mime="text/csv"
)
