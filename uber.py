import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title= "Uber",layout="wide")
st.title("📊 Dashboard de Viagens Corporativas - Uber")
st.markdown("-----------------")
#ETL dos arquivos

df_geral = pd.read_excel("Registros Uber Original.xlsx", sheet_name="TOTAL")

jan_fev = pd.read_excel("Registros Uber Original.xlsx",sheet_name="Janeiro-fevereiro")

març_abril_maio = pd.read_excel("Registros Uber Original.xlsx", sheet_name="Março - abril- Maio")

jun_jul_agost = pd.read_excel("Registros Uber Original.xlsx",sheet_name="Jun- jul- Agosto")



#Análise dos Dados

def formatar_reais(valor):
    return f"R$ {valor:,.2f}".replace(",","X").replace(".",",").replace("X",".")

jan_fev["Data da solicitação (local)"] = pd.to_datetime(jan_fev["Data da solicitação (local)"])

març_abril_maio["Data da solicitação (local)"] = pd.to_datetime(març_abril_maio["Data da solicitação (UTC)"])

jun_jul_agost["Data da solicitação (local)"] = pd.to_datetime(jun_jul_agost["Data da solicitação (UTC)"])

df_geral["Data da solicitação (local)"] = pd.to_datetime(df_geral["Data da solicitação (local)"])


gasto_dias_jan = jan_fev.groupby(["Data da solicitação (local)","Nome"])["Valor da transação em BRL (com tributos)"].sum().reset_index(name="Gasto Diário")
gastos_dias_març = març_abril_maio.groupby(["Data da solicitação (UTC)","Nome"])["Valor da transação em BRL (com tributos)"].sum().reset_index(name="Gasto Diário")
gasto_jun = jun_jul_agost.groupby(["Data da solicitação (UTC)","Nome"])["Valor da transação em BRL (com tributos)"].sum().reset_index(name="Gasto Diário")

gasto_geral_mensal = df_geral.groupby("Mês")["Valor da transação em BRL (com tributos)"].sum().reset_index(name="Gasto Mensal")

tipos_viagens_jan = jan_fev["Serviço"].value_counts().reset_index(name="Quantidade de Solicitações")

tipos_viagens_març = març_abril_maio["Serviço"].value_counts().reset_index(name="Quantidade de Solicitações")

tipos_viagens_jun = jun_jul_agost["Serviço"].value_counts().reset_index(name="Quantidade de solicitações")

#Dashboard
tab1, tab3 = st.tabs(["Visão Geral","Consultar Pessoa"])

with tab1:

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🔎**Visão Geral dos Gastos**")
        st.markdown("---------------")
        gasto_geral_mensal["Gasto Mensal"] = gasto_geral_mensal["Gasto Mensal"].apply(formatar_reais)
        mes_ordem = {'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4,
                'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
                'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12}
        gasto_geral_mensal['ordem_mes'] = gasto_geral_mensal['Mês'].map(mes_ordem)
        gasto_geral_mensal = gasto_geral_mensal.sort_values('ordem_mes')
        gasto_geral_mensal = gasto_geral_mensal[["Mês","Gasto Mensal"]]
        st.subheader("💵Gasto total por Mês")
        st.dataframe(gasto_geral_mensal)

    with col2:
        fig_bar_mes = px.bar(gasto_geral_mensal,
                            x="Mês",
                            y="Gasto Mensal",
                            title="Gasto Mensal (R$)"
                            )
        
        st.plotly_chart(fig_bar_mes,use_container_width=True)

    gasto_total_individual = df_geral.groupby("Nome")["Valor da transação em BRL (com tributos)"].sum().reset_index(name="Gasto Total")
    gasto_total_individual["Gasto Total Formatado"] = gasto_total_individual["Gasto Total"].apply(formatar_reais)

    # Paleta de roxos dark
    paleta_roxo_dark = [
        '#2D1B69', '#3D28A8', '#4A36D9', '#5F4AE8', '#7261F2',
        '#8A7DFF', '#9E8EFF', '#B3A5FF', '#C8BCFF', '#DCD4FF'
    ]

    fig_pie = px.pie(
        gasto_total_individual,
        names="Nome",
        values="Gasto Total",
        title="<b> Distribuição de Gastos por Colaborador</b>",
        hole=0.4,
        color_discrete_sequence=paleta_roxo_dark
    )

    # Layout dark elegante
    fig_pie.update_layout(
        plot_bgcolor='#1E1E2E',  # Fundo dark
        paper_bgcolor='#1E1E2E',  # Fundo dark
        font=dict(color='#E0DEF4', size=12, family="Arial"),
        legend=dict(
            title=dict(text='<b>👥 Colaboradores</b>', font=dict(color='#E0DEF4')),
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="right",
            x=1.2,
            font=dict(color='#E0DEF4')
        ),
        title_x=0.5,
        title_font_size=18,
        title_font_color='#C9B6F2',
        margin=dict(t=80, b=40, l=40, r=40)
    )

    fig_pie.update_traces(
        textinfo='percent+label',
        textposition='inside',
        textfont=dict(color='white', size=11),
        texttemplate='<b>%{label}</b><br>%{percent:.1%}',
        hovertemplate='<b>%{label}</b><br>💸 Valor: %{customdata}<br>📊 Percentual: %{percent:.1%}<extra></extra>',
        customdata=gasto_total_individual["Gasto Total Formatado"],
        marker=dict(line=dict(color='#1E1E2E', width=2)),
        insidetextfont=dict(color='white')
    )

    st.plotly_chart(fig_pie, use_container_width=True)
    st.subheader("📋 Resumo em Tabela")
    gasto_total_individual = gasto_total_individual.sort_values(by="Gasto Total",ascending=False)
    st.dataframe(gasto_total_individual[["Nome", "Gasto Total Formatado"]], use_container_width=True)

with tab3:
    st.subheader("🔎Consultar Colaborador")
    pessoa_input = st.text_input("Digite o nome do Colaborador")

    if pessoa_input:
        resultados = df_geral[df_geral["Nome"].str.contains(pessoa_input,case=False,na=False)]
        if not resultados.empty:

            st.subheader(f"📊 Resumo do Colaborador: {pessoa_input.title()}")

            col1,col2,col3, col4, col8 = st.columns(5)

            with col1:
                total_gasto = resultados["Valor da transação em BRL (com tributos)"].sum()
                st.metric("**Total Gasto**", formatar_reais(total_gasto))

            with col2:
                dias_unico = resultados["Data da solicitação (local)"].nunique()
                gasto_diario_medio = total_gasto/ dias_unico if  dias_unico > 0 else 0
                st.metric("**Gasto Médio Diário**", formatar_reais(gasto_diario_medio))

            with col3:
                resultados["Ano-Mes"] = resultados["Data da solicitação (local)"].dt.to_period("M")
                meses_unico = resultados["Ano-Mes"].nunique()
                gasto_mensal_medio = total_gasto / meses_unico if meses_unico > 0 else 0
                st.metric("**Gasto Médio Mensal**", formatar_reais(gasto_mensal_medio))

            with col4:
                qtd_solicitacoes = len(resultados)
                st.metric("**Quantidade de Solicitações**", qtd_solicitacoes)

            with col8:
                servico_mais_usado = resultados["Serviço"].mode()
                servico_mais_frequente = servico_mais_usado.iloc[0] if not servico_mais_usado.empty else "Nenhum"
                freq_servico = resultados["Serviço"].value_counts().iloc[0] if not resultados.empty else 0

                st.metric("**Serviço mais utilizado**", servico_mais_frequente)
                st.caption(f"Utilizado {freq_servico} vezes")
            col5, col6,col7 = st.columns(3)
            with col5:
                data_inicial = st.date_input("**Data Inicial**", value= resultados["Data da solicitação (local)"].min())
            with col6:
                data_final = st.date_input("**Data Final**", value=resultados["Data da solicitação (local)"].max())
            resultados_filtrados = resultados[
                (resultados["Data da solicitação (local)"] >= pd.to_datetime(data_inicial)) &
                (resultados["Data da solicitação (local)"] <= pd.to_datetime(data_final))
            ]
            if not resultados_filtrados.empty:

                total_periodo = resultados_filtrados["Valor da transação em BRL (com tributos)"].sum()
                qtd_solicitacoes_individuais = len(resultados_filtrados)
                
                dias_periodo = (pd.to_datetime(data_final) - pd.to_datetime(data_inicial)).days + 1
                dias_com_gasto = resultados_filtrados["Data da solicitação (local)"].nunique()
                gasto_diario_medio = total_periodo / dias_com_gasto if dias_com_gasto > 0 else 0

                meses_periodo = dias_periodo / 30.44
                gasto_mensal_medio = total_periodo / meses_periodo if meses_periodo > 0 else 0

                st.subheader(f"📊 Métricas do Período: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}")


                col1,col2,col3,col4 = st.columns(4)

                with col1:
                    st.metric("💰 Montante Gasto", formatar_reais(total_periodo))
                    st.caption("Total no Período")

                with col2:
                    st.metric("Solicitações", qtd_solicitacoes)
                    st.caption("Quantidade de Solicitações")
                
                with col3:
                    st.metric("📅Gasto Médio Diário", formatar_reais(gasto_diario_medio))
                    st.caption(f"Em {dias_com_gasto} dias com gasto")
                with col4:
                    st.metric("📊Gasto Médio Diário",formatar_reais(gasto_mensal_medio))
                    st.caption(f"Período de {dias_periodo} dias")


                resultados_filtrados["Valor da transação em BRL (com tributos)"] = resultados_filtrados["Valor da transação em BRL (com tributos)"].apply(formatar_reais)
                st.dataframe(resultados_filtrados, use_container_width=True)
            else:
                st.warning("Nenhuma Transação encontrada neste Período")
            resultados["Valor da transação em BRL (com tributos)"] = resultados["Valor da transação em BRL (com tributos)"].apply(formatar_reais)

        else:
            st.warning("Ops...Nenhum Colaborador encontrado")


    
