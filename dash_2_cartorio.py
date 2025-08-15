import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="Dashboard 2 Cartório", layout="wide")

st.title("Dashboard 2° Cartório")

df_emissoes = pd.read_excel("CRC_2_OFICIO.xlsx")

df_emissoes["DATA"] = pd.to_datetime(df_emissoes["DATA"], format="%Y-%m-%d", errors="coerce")

df_emissoes["MES"] = df_emissoes["DATA"].dt.month

 
df_emissoes["VALOR"] = df_emissoes["VALOR"].str.replace(",",".")

df_emissoes["VALOR"] = pd.to_numeric(df_emissoes["VALOR"], errors="coerce")

tab1, tab2 = st.tabs(["Análise das Emissões","Análise dos Apostilamentos"])

def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

df_emissoes["MES-ANO"] = df_emissoes["DATA"].dt.to_period("M")
receita_mensal_emissoes = df_emissoes.groupby("MES-ANO")["VALOR"].sum().reset_index()
receita_mensal_emissoes["MES-ANO"] = receita_mensal_emissoes["MES-ANO"].astype(str)

with tab1:

    col1, col2 = st.columns(2)

    with col1:
        receita_total = df_emissoes["VALOR"].sum()
        st.metric("Receita Total Emissões", formatar_valor(receita_total))
        st.write("---------------------------")
    
        receita_mensal_emissoes_copia = receita_mensal_emissoes.copy()
        receita_mensal_emissoes_copia["VALOR"] = receita_mensal_emissoes_copia["VALOR"].apply(formatar_valor)
        st.dataframe(receita_mensal_emissoes_copia, use_container_width=True)

    with col2:
        st.metric("Preço Médio Unitário por **Emissão**", formatar_valor(df_emissoes["VALOR"].mean()))

        st.write("---------------------------")

        fig_receita_emissoes = px.line(
            receita_mensal_emissoes,
             x="MES-ANO",
             y="VALOR",
             title="Receita Mensal de Emissões",
             labels={"MES-ANO": "Mês-Ano", "VALOR": "Receita (R$)"},
        )
        
        st.plotly_chart(fig_receita_emissoes, use_container_width=True)

    st.write("---------------------------")

    col3, col4 = st.columns(2)

    with col3:
        st.metric("**Maior** Preço Registrado",formatar_valor(df_emissoes["VALOR"].max()))

        st.write("---------------------------")

        maiores_precos_registrados = df_emissoes.groupby("MES-ANO")["VALOR"].max().reset_index()
        maiores_precos_registrados["MES-ANO"] = maiores_precos_registrados["MES-ANO"].astype(str)

        fig_maiores_precos = px.scatter(
            maiores_precos_registrados,
            x="MES-ANO",
            y="VALOR",
            title="Maiores Preços Registrados por Mês",
            labels={"MES-ANO": "Mês-Ano", "VALOR": "Valor (R$)"},
            color = "VALOR"
        )

        st.plotly_chart(fig_maiores_precos, use_container_width=True)

    with col4:
        st.metric("**Menor** Preço Registrado", formatar_valor(df_emissoes["VALOR"].min()))

        st.write("---------------------------")

        maiores_precos_registrados_COPIA = maiores_precos_registrados.copy()
        maiores_precos_registrados_COPIA["VALOR"] = maiores_precos_registrados_COPIA["VALOR"].apply(formatar_valor)
        st.dataframe(maiores_precos_registrados_COPIA, use_container_width=True)

with tab2:

    col1, col2 = st.columns(2)

    with col1:
        df_apostilamentos = pd.read_excel("Apostilamento_copy.xlsx")
        df_apostilamentos["MES-ANO"] = df_apostilamentos["DATA DE APOSTILAMENTO"].dt.to_period("M")
        receita_mensal_apostilamentos = df_apostilamentos.groupby("MES-ANO")["VALOR"].sum().reset_index()
        receita_mensal_apostilamentos["MES-ANO"] = receita_mensal_apostilamentos["MES-ANO"].astype(str)

        st.metric("Receita Total Apostilamentos", formatar_valor(df_apostilamentos["VALOR"].sum()))

        st.write("---------------------------")

        st.subheader("Receitas Mensais de Apostilamentos")

        receita_mensal_apostilamentos_copia = receita_mensal_apostilamentos.copy()
        receita_mensal_apostilamentos_copia["VALOR"] = receita_mensal_apostilamentos_copia["VALOR"].apply(formatar_valor)
        st.dataframe(receita_mensal_apostilamentos_copia,use_container_width=True)

    with col2:
        st.metric("Preço Médio de Apostilamentos", formatar_valor(df_apostilamentos["VALOR"].mean()))

        st.write("------------------------------")

        fig_receita_apostilamentos = px.line(
            receita_mensal_apostilamentos,
            x="MES-ANO",
            y="VALOR",
            title="Receita Mensal de Apostilamentos",
            labels={"MES-ANO": "Mês-Ano", "VALOR": "Receita (R$)"},
        )
        st.plotly_chart(fig_receita_apostilamentos, use_container_width=True)


    st.write("O **Preço** de Apostilamentos é o mesmo para todos os tipos de documentos, portanto, não há necessidade de análise de maiores e menores preços.")

