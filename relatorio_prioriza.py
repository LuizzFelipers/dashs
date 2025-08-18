import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px

#Carregar os dados

st.set_page_config(page_title="Contas a Vencer", layout="wide")
st.title("📅 Contas a Vencer (Priorização por Impacto)")

st.markdown("-----------")

@st.cache_data
def carregar_dados():

    df = pd.read_excel("contas_a_pagar_receber.xlsx")

    return df

df = carregar_dados()

df.info()

df["Data Vencimento"] = pd.to_datetime(df["Data Vencimento"],format="%d/%m/%Y")

with st.sidebar:
    st.header("Filtros")
    date_range = st.date_input(
        "Período de Vencimento",
        value=(datetime.now(), datetime.now() + pd.DateOffset(months=1)),
        key="date_filter"
    )
    
    tipo_conta = st.multiselect(
        "Tipo de Conta",
        options=["Pagar","Receber"],
        default=["Pagar","Receber"],
        key="tipo_filter"
    )

filtered_df = df[
    (df["Data Vencimento"].between(pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]))) &
    (df["Tipo de Conta"].isin(tipo_conta))
]
#priorização por data de vencimento e impacto no caixa

df["Dias ate o Vencimento"] = (df["Data Vencimento"] - datetime.now()).dt.days
df["Desconto Antecipação (%)"] = df["Desconto Antecipação (%)"].str.replace('%','').astype(float)
df["Impacto no caixa"] = df["Valor (R$)"] * (df["Desconto Antecipação (%)"]/100)
df["Prioridade"] = df.apply(
    lambda row: row["Valor (R$)"] / row["Dias ate o Vencimento"] if row["Dias ate o Vencimento"] > 0 else 9999,
    axis=1
)

def formatar_moeda_br(valor: float) -> str:
    """
    Formata um float como moeda brasileira (R$ 1.234.567,89).
    Exemplo:
        >>> formatar_moeda_br(1576456.58)
        'R$ 1.576.456,58'
    """
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

st.subheader("Principais Métricas")
col1, col2, col3 = st.columns(3)

with col1:

    total_pagar = filtered_df[filtered_df["Tipo de Conta"] == "Pagar"]["Valor (R$)"].sum()
    st.metric("Total a Pagar:", f"{formatar_moeda_br(total_pagar)}")

with col2:

    total_receber = filtered_df[filtered_df["Tipo de Conta"] == "Receber"]["Valor (R$)"].sum()
    st.metric("Total a Receber:",f" {formatar_moeda_br(total_receber)}")

with col3:

    st.metric("Total Disponível em Caixa:",f"")

st.markdown("-----------------")

tab1, tab2 = st.tabs(["Resumo de Contas a Pagar","Resumo de Contas a Receber"])

with tab1:
    st.subheader("Resumo de Contas a Pagar")
    fig_comparacao = px.bar(
        filtered_df[filtered_df["Tipo de Conta"] == "Pagar"].groupby("Descrição")["Valor (R$)"].sum().reset_index(name="Valor Total (R$)"),
        x="Descrição",
        y="Valor Total (R$)",
        title="Distribuição de Contas a Pagar por Descrição (R$)",
        color=["Antecipação","Cancelamentos","Manutenção","Serviço"]
    )
    st.plotly_chart(fig_comparacao,use_container_width=True)

    st.markdown("------")

    st.subheader("Tabela completa do Tipo de Conta Pagar")
    # Passo 1: Agrupar e somar (mantendo "Descrição" como coluna)
    df_pagar = (
        filtered_df[filtered_df["Tipo de Conta"] == "Pagar"]
        .groupby("Descrição", as_index=False)  # ← as_index=False mantém "Descrição" como coluna
        .agg({"Valor (R$)": "sum"})
        .rename(columns={"Valor (R$)": "Valor Total (R$)"})
    )

    # Passo 2: Criar coluna formatada (sem perder a original)
    df_pagar["Valor Total (R$)"] = df_pagar["Valor Total (R$)"].apply(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

    st.dataframe(df_pagar)


with tab2:
    st.subheader("Resumo de Contas a Receber")

    fig_receber = px.bar(
        filtered_df[filtered_df["Tipo de Conta"] == "Receber"].groupby("Descrição")["Valor (R$)"].sum().reset_index(name="Valor Total (R$)"),
        x= "Descrição",
        y="Valor Total (R$)",
        title="Distribuição de Contas a Receber por Descrição(R$)",
        color=["Antecipação","Cancelamentos","Manutenção","Serviço"]
    )
    st.plotly_chart(fig_receber,use_container_width=True)

    st.markdown('----------------')

    st.subheader("Tabela completa do Tipo de Conta a Receber")
    df_receber = (
        filtered_df[filtered_df["Tipo de Conta"] == "Receber"]
        .groupby("Descrição", as_index=False)  # ← as_index=False mantém "Descrição" como coluna
        .agg({"Valor (R$)": "sum"})
        .rename(columns={"Valor (R$)": "Valor Total (R$)"})
    )

    df_receber["Valor Total (R$)"] = df_receber["Valor Total (R$)"].apply(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

    st.dataframe(df_receber)