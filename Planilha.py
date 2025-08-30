import io
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

st.set_page_config(page_title="Análise de Ordens de Serviço", page_icon="📊", layout="wide")

st.title("📊 Análise de Ordens de Serviço")
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
required_columns = ["Abertura", "Tipo de Manutenção", "Plano de Manutenção", "Técnico Resolvedor", "Setor"]
if not all(col in df.columns for col in required_columns):
    missing_cols = [col for col in required_columns if col not in df.columns]
    st.error(f"A planilha precisa ter as colunas: {', '.join(missing_cols)}.")
    st.stop()

# Prepara os dados
df["Abertura"] = pd.to_datetime(df["Abertura"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Abertura"]).copy()
df["Ano"] = df["Abertura"].dt.year
df["MesNum"] = df["Abertura"].dt.month

# Normaliza e cria a coluna 'Unidade'
df["Unidade"] = df["Setor"].str.split('_').str[0].str.strip()
df["Unidade"] = df["Unidade"].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

# Sidebar: filtros
st.sidebar.header("Filtros")
anos = sorted(df["Ano"].dropna().unique())
ano_sel = st.sidebar.selectbox("Filtrar por ano", anos, index=len(anos) - 1 if anos else 0)

unidades = sorted(df["Unidade"].dropna().unique())
unidade_sel = st.sidebar.selectbox("Filtrar por unidade", ["Todas"] + unidades)

# Aplica os filtros
df_filtrado = df[df["Ano"] == ano_sel].copy()
if unidade_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Unidade"] == unidade_sel].copy()

if df_filtrado.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()

# ===== Exibição =====
st.subheader(f"Resumo ({ano_sel})")
col1, col2 = st.columns(2)
ordens_por_mes = df_filtrado.groupby("MesNum").size().reindex(range(1, 13), fill_value=0)
col1.metric("Total de OS", int(ordens_por_mes.sum()))
col2.metric("Média mensal", f"{ordens_por_mes.mean():.1f}")

with st.expander("Prévia dos dados (primeiras 100 linhas)"):
    st.dataframe(df_filtrado.head(100))

# Função para gerar gráficos de barras
def plot_bar_chart(data, title, x_label, y_label):
    fig, ax = plt.subplots(figsize=(10, 6))
    data.plot(kind="bar", ax=ax, color='skyblue')
    
    # Adiciona rótulos de dados
    for i, valor in enumerate(data):
        ax.text(i, valor, str(valor), ha='center', va='bottom', fontsize=10)
    
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig, clear_figure=True)
    return fig

# ===== Geração dos Gráficos =====
st.subheader("Gráficos de Análise")
col_graf1, col_graf2 = st.columns(2)

# Gráfico 1: Ordens de Serviço por Mês
with col_graf1:
    nomes_meses = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
                   5: "maio", 6: "junho", 7: "julho", 8: "agosto",
                   9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
    ordens_por_mes.index = [nomes_meses[m] for m in ordens_por_mes.index]
    fig1 = plot_bar_chart(ordens_por_mes,
                          f"Ordens de Serviço por Mês ({ano_sel})",
                          "Mês", "Quantidade de OS")

# Gráfico 2: Quantidade de OS por Tipo
with col_graf2:
    ordens_por_tipo = df_filtrado["Tipo de Manutenção"].value_counts().sort_index()
    fig2 = plot_bar_chart(ordens_por_tipo,
                          f"OS por Tipo de Manutenção ({ano_sel})",
                          "Tipo", "Quantidade de OS")

# Gráfico 3: OS Programada x Não Programada
st.subheader("Programadas x Não Programadas")
col_prog1, col_prog2 = st.columns(2)

with col_prog1:
    # Cria a coluna 'Status' para o gráfico de Programadas x Não Programadas
    # Se 'Plano de Manutenção' for nulo, é 'Não Programada'. Caso contrário, é 'Programada'.
    df_filtrado['Status'] = df_filtrado['Plano de Manutenção'].apply(lambda x: 'Não Programada' if pd.isna(x) else 'Programada')
    os_planejamento = df_filtrado['Status'].value_counts()
    fig3 = plot_bar_chart(os_planejamento,
                          f"OS Programada x Não Programada ({ano_sel})",
                          "Status", "Quantidade de OS")

# Gráfico 4: OS por Técnico Resolvedor
with col_prog2:
    ordens_por_tecnico = df_filtrado["Técnico Resolvedor"].value_counts().sort_index()
    fig4 = plot_bar_chart(ordens_por_tecnico,
                          f"OS por Técnico Resolvedor ({ano_sel})",
                          "Técnico", "Quantidade de OS")


# ===== Download dos Gráficos =====
st.sidebar.header("Download")
for title, fig in [("Mês", fig1), ("Tipo", fig2), ("Planejamento", fig3), ("Técnico", fig4)]:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    buf.seek(0)
    st.sidebar.download_button(
        f"Baixar Gráfico de {title} (PNG)",
        data=buf,
        file_name=f"ordens_por_{title.lower().replace(' ', '_')}_{ano_sel}.png",
        mime="image/png"
    )
