# app.py
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Concat Excel", layout="wide")
st.title("📊 Concatenador de planilhas Excel")
st.subheader("OLÁ SABRINA!!!")
st.caption("Envie vários arquivos .xlsx ou arraste uma paste, aplique pré-processamento e baixe tudo concatenado.")

# -------------------- Sidebar (opções) --------------------
st.sidebar.header("⚙️ Opções")
skiprows = st.sidebar.number_input("Linhas a pular no início", min_value=0, value=9, step=1)
skipcols_left = st.sidebar.number_input("Colunas a remover (esquerda → direita)", min_value=0, value=0, step=1)
remover_unnamed = st.sidebar.checkbox("Remover colunas 'Unnamed:*'", value=True)
adicionar_coluna = st.sidebar.checkbox("Adicionar coluna fixa", value=False)
if adicionar_coluna:
    nome_col = st.sidebar.text_input("Nome da coluna", value="Homem Aranha")
    valor_col = st.sidebar.text_input("Valor da coluna", value="1")

# -------------------- Upload --------------------
up_files = st.file_uploader("Selecione um ou mais .xlsx", type=["xlsx"], accept_multiple_files=True)

# -------------------- Utils --------------------
def limpar_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

@st.cache_data(show_spinner=False)
def ler_xlsx(file, skiprows: int, skipcols_left: int,
             remover_unnamed_flag: bool, add_col: bool, nome_coluna: str, valor_coluna: str):
    df = pd.read_excel(file, skiprows=skiprows, engine="openpyxl")

    if skipcols_left > 0:
        n = min(skipcols_left, df.shape[1])
        df = df.iloc[:, n:]

    if remover_unnamed_flag:
        df = limpar_unnamed(df)

    if add_col and nome_coluna:
        df.insert(0, nome_coluna, valor_coluna)

    df.insert(0, "arquivoOrigem", file.name)
    return df

def baixar_excel(df: pd.DataFrame, sheet_name: str = "dados") -> bytes:
    # Gera um .xlsx com "autofit" simples de largura de colunas
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        # Autoajuste básico das colunas (limite de 50)
        for idx, col in enumerate(df.columns):
            serie = df[col].astype(str)
            max_len = max([len(col)] + [len(s) for s in serie.head(1000)])  # amostra p/ performance
            ws.set_column(idx, idx, min(max_len + 2, 50))
    buffer.seek(0)
    return buffer.read()

# -------------------- Processamento --------------------
st.subheader("🛠️ Processar")
if st.button("Concatenar"):
    if not up_files:
        st.warning("Envie pelo menos um arquivo.")
    else:
        dfs = []
        progress = st.progress(0)
        for i, f in enumerate(up_files, start=1):
            df = ler_xlsx(
                f,
                skiprows,
                skipcols_left,
                remover_unnamed,
                adicionar_coluna,
                nome_col if adicionar_coluna else "",
                valor_col if adicionar_coluna else ""
            )
            dfs.append(df)
            progress.progress(i / len(up_files))

        if dfs:
            df_final = pd.concat(dfs, ignore_index=True)
            st.success(f"Concluído! Dimensões finais: {df_final.shape[0]} linhas × {df_final.shape[1]} colunas")

            st.write("### 🔎 Preview")
            st.dataframe(df_final.head(100), use_container_width=True)

            # ✅ Exporta apenas Excel (nada de CSV com .xlsx)
            st.download_button(
                "⬇️ Baixar Excel",
                data=baixar_excel(df_final),
                file_name="concatenado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Nenhum DataFrame processado.")
            
# -------------------- Dicas --------------------
with st.expander("💡 Suporte"):
     st.markdown(
"""
- Carlos Massato Horibe Chinen 👨‍💻  
- Guilherme Amato 👨‍💼
- Maura Chagas 👩‍💻
"""
)
