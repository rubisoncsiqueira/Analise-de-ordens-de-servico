import io
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

st.set_page_config(page_title="Análise de Ordens de Serviço", page_icon="📊", layout="wide")

st.title("📊 Análise de Ordens de Serviço")
st.write("Envie a planilha Excel ou use o arquivo de exemplo na pasta `dados/`.")

# Inicializa uma flag na session_state para controlar o upload
if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False
    st.session_state.file_to_read = None

# ===== Entrada de dados =====
file_to_read = None

# Tenta carregar o arquivo padrão primeiro
default_path = Path("dados/Planilha_base_python.xlsx")
if default_path.exists():
    file_to_read = default_path.open("rb")
    st.session_state.file_to_read = default_path.open("rb")
    st.info("Usando arquivo padrão: `dados/Planilha_base_python.xlsx`")
else:
    st.session_state.file_to_read = None

# Se o arquivo padrão não existe ou não foi carregado, exibe o uploader
if st.session_state.file_to_read is None:
    # Exibe o file_uploader
    uploaded = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"])
    if uploaded is not None:
        file_to_read = uploaded
        st.session_state.file_to_read = uploaded
    else:
        st.warning("Nenhum arquivo padrão encontrado. Por favor, faça o upload de uma planilha para continuar.")
        st.stop()
else:
    # Se já há um arquivo na session_state (do upload ou do arquivo padrão), usa ele
    file_to_read = st.session_state.file_to_read

# ===== Leitura e preparo =====
try:
    # Reinicia o cursor do arquivo para o início, caso ele já tenha sido lido
    if hasattr(file_to_read, 'seek'):
        file_to_read.seek(0)
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

nomes_meses = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
               5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
               9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
meses_list = sorted(df_filtrado['MesNum'].unique())
meses_sel = {num: nomes_meses[num] for num in meses_list}
mes_sel_text = st.sidebar.selectbox("Filtrar por mês", ["Todos"] + list(meses_sel.values()))

# Aplica os filtros
df_filtrado = df[df["Ano"] == ano_sel].copy()
if unidade_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Unidade"] == unidade_sel].copy()
if mes_sel_text != "Todos":
    mes_num = list(nomes_meses.keys())[list(nomes_meses.values()).index(mes_sel_text)]
    df_filtrado = df_filtrado[df_filtrado['MesNum'] == mes_num].copy()

if df_filtrado.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()

# ===== Exibição =====
st.subheader(f"Resumo ({mes_sel_text}, {ano_sel})")
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
    ordens_por_tipo = df_filtrado["Tipo de Manutenção"].value_counts().head(10)
    fig2 = plot_bar_chart(ordens_por_tipo,
                          f"Top 10 OS por Tipo de Manutenção ({ano_sel})",
                          "Tipo", "Quantidade de OS")

# Gráfico 3: OS Programada x Não Programada
st.subheader("Programadas x Não Programadas")
col_prog1, col_prog2 = st.columns(2)

with col_prog1:
    df_filtrado['Status'] = df_filtrado['Plano de Manutenção'].apply(lambda x: 'Não Programada' if pd.isna(x) else 'Programada')
    os_planejamento = df_filtrado['Status'].value_counts()
    fig3 = plot_bar_chart(os_planejamento,
                          f"OS Programada x Não Programada ({ano_sel})",
                          "Status", "Quantidade de OS")

# Gráfico 4: OS por Técnico Resolvedor
with col_prog2:
    ordens_por_tecnico = df_filtrado["Técnico Resolvedor"].value_counts().head(10)
    fig4 = plot_bar_chart(ordens_por_tecnico,
                          f"Top 10 OS por Técnico Resolvedor ({ano_sel})",
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
