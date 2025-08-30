import io
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

st.set_page_config(page_title="Ordens por Mês", page_icon="📊", layout="centered")

st.title("📊 Ordens de Serviço por Mês")
st.write("Envie a planilha Excel ou use o arquivo de exemplo na pasta `dados/`.")

# ===== Entrada de dados =====
uploaded = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"])

# fallback: dados/Planilha_base_python.xlsx se existir
default_path = Path("dados/Planilha_base_python.xlsx")
file_to_read = None

if uploaded is not None:
    file_to_read = uploaded
elif default_path.exists():
    file_to_read = default_path.open("rb")
    st.info("Usando arquivo padrão: `dados/Planilha_base_python.xlsx`")

if file_to_read is None:
    st.warning("Envie um arquivo para continuar.")
    st.stop()

# ===== Leitura e preparo =====
try:
    df = pd.read_excel(file_to_read)
except Exception as e:
    st.error(f"Erro ao ler o Excel: {e}")
    st.stop()

# Validação básica
if "Abertura" not in df.columns:
    st.error("A planilha precisa ter a coluna **'Abertura'**.")
    st.stop()

# Converte a coluna 'Abertura' para datetime (formato dd/mm/aaaa hh:mm)
df["Abertura"] = pd.to_datetime(df["Abertura"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Abertura"]).copy()
df["Ano"] = df["Abertura"].dt.year
df["MesNum"] = df["Abertura"].dt.month

# Sidebar: filtro de ano
anos = sorted(df["Ano"].dropna().unique())
ano_sel = st.sidebar.selectbox("Filtrar por ano", anos, index=len(anos)-1 if anos else 0)

df_filtrado = df[df["Ano"] == ano_sel].copy()

# Agrupa por mês
ordens_por_mes = df_filtrado.groupby("MesNum").size().reindex(range(1, 13), fill_value=0)

# Mapeia número do mês para nome em português e mantém ordem correta
nomes_meses = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}
ordens_por_mes.index = [nomes_meses[m] for m in ordens_por_mes.index]

# ===== Exibição =====
st.subheader(f"Resumo ({ano_sel})")
col1, col2 = st.columns(2)
col1.metric("Total de OS", int(ordens_por_mes.sum()))
col2.metric("Média mensal", f"{ordens_por_mes.mean():.1f}")

with st.expander("Prévia dos dados (primeiras 100 linhas)"):
    st.dataframe(df_filtrado.head(100))

st.subheader("Gráfico de OS por mês")
fig, ax = plt.subplots()
ordens_por_mes.plot(kind="bar", ax=ax)
ax.set_xlabel("Mês")
ax.set_ylabel("Ordens de Serviço")
ax.set_title(f"Ordens de Serviço Abertas por Mês ({ano_sel})")
plt.xticks(rotation=45)
plt.tight_layout()
st.pyplot(fig, clear_figure=True)

# ===== Download do gráfico =====
buf = io.BytesIO()
fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
buf.seek(0)
st.download_button("Baixar gráfico (PNG)", data=buf, file_name=f"ordens_por_mes_{ano_sel}.png", mime="image/png")

