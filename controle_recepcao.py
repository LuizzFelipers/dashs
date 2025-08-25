import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="Controle de Estoque Geral", layout = "wide")
st.title("📌**Controle de Estoque Geral**📌")

df = pd.read_excel("Controle de estoque recepção.xlsx")

tab1, tab2, tab3, tab4 = st.tabs(["Análise Geral do Facilites","Estoque Recepção - 3° Andar","🧼Limpeza 4° Andar e Copa☕","Gastos Joinville"])

def formatar_reais(valor):

    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
with tab1:

    st.subheader("🪜**Análise Geral do Facilites**🛠️")
    st.write("Este dashboard apresenta uma visão geral dos gastos mensais e controle de estoque do 4° e 3° Andar.")
    st.write("-------------------------------")

    col1, col2 = st.columns(2)

    gastos_mensais = {
                      "SOS/Manutenção em Média":2330.67,
                      "Insight (Ar-Condicionado)":3600,
                      "Telefonia e Internet":4000,
                      "Estacionamento":4800,
                      "Aluguel 4° Andar":16000,
                      "Condomínio Total":40344.68,
                        'Quality': 15553.95,
                        'Copa': 3975.22,
                        'Limpeza':12224.03,
                        'Energia Elétrica': 16000,
                        'Materiais de Escritórios': 7364.02
                      }
    
    items = list(gastos_mensais.items())
    df_gastos = pd.DataFrame(items, columns=["Tipo de gasto", "Valor Mensal (R$)"])
    df_gastos = df_gastos.sort_values(by="Valor Mensal (R$)", ascending=False)


    with col1:
        st.metric("**Gasto Mensal Total**", formatar_reais(df_gastos["Valor Mensal (R$)"].sum()))
        st.write('-------------------------------')
        st.subheader("💸**Gastos Totais Mensais**💸")
        df_display = df_gastos.copy()
        df_display["Valor Mensal (R$)"] = df_display["Valor Mensal (R$)"].apply(formatar_reais)
        st.dataframe(df_display, use_container_width=True)


    with col2:

        st.subheader("📊**Gráfico de Gastos Mensais**📊")
        fig = px.bar(df_gastos,
                    x="Tipo de gasto",
                    y="Valor Mensal (R$)",
                    color="Tipo de gasto",
                    title="Gastos Mensais do Facilites")

        st.plotly_chart(fig, use_container_width=True)

    st.markdown("-------------------------------")
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("**💸Gastos com o 3° Andar💸**")
       
        gastos_3_andar = {
            "SOS/Manutenção em Média": 1165.34,
            "Insight (Ar-Condicionado)": 1800,
            "Telefonia e Internet": 2000,
            "Estacionamento": 2400,
            "Condomínio Total": 21564.60,
            'Quality': 7776.98,
            'Copa': 1987.61,
            'Limpeza': 6112.01,
            'Energia Elétrica': 11000,
            'Materiais de Escritórios': 3682.01
        }

        df_gastos_3_andar = pd.DataFrame(gastos_3_andar.items(), columns=["Tipo de gasto", "Valor Mensal (R$)"])
        df_gastos_3_andar = df_gastos_3_andar.sort_values(by="Valor Mensal (R$)", ascending=False)
        df_display_3_andar = df_gastos_3_andar.copy()
        df_display_3_andar["Valor Mensal (R$)"] = df_display_3_andar["Valor Mensal (R$)"].apply(formatar_reais)
        
        st.metric("💰**Gasto Mensal Total com o 3° Andar**💰", formatar_reais(sum(gastos_3_andar.values())))
        st.dataframe(df_display_3_andar, use_container_width=True)

    with col4:

        st.subheader("📊**Gráfico de Gastos Mensais do 3° Andar**📊")
        fig_3_andar = px.pie(df_gastos_3_andar,
                            names="Tipo de gasto",
                            values="Valor Mensal (R$)",
                            title="Gastos Mensais do 3° Andar",
                            hole=0.3)
        fig_3_andar.update_traces(textinfo='percent+label')

        st.plotly_chart(fig_3_andar, use_container_width=True)

    st.markdown("-------------------------------")
    col4, col5 = st.columns(2)
    with col4:

        st.subheader("**💸Gastos com o 4° Andar💸**")
        gastos_4_andar ={
            "SOS/Manutenção em Média": 1165.33,
            "Aluguel 4° Andar": 16000,
            "Insight (Ar-Condicionado)": 1800,
            "Telefonia e Internet": 2000,
            "Estacionamento": 2400,
            "Condomínio Total": 18780.08,
            'Quality': 7776.97,
            'Copa': 1987.61,
            'Limpeza': 6112.02,
            'Energia Elétrica': 5000,
            'Materiais de Escritórios': 3682.01
        }

        df_gastos_4_andar = pd.DataFrame(gastos_4_andar.items(), columns=["Tipo de gasto", "Valor Mensal (R$)"])
        df_gastos_4_andar = df_gastos_4_andar.sort_values(by="Valor Mensal (R$)", ascending=False)
        df_display_4_andar = df_gastos_4_andar.copy()
        df_display_4_andar["Valor Mensal (R$)"] = df_display_4_andar["Valor Mensal (R$)"].apply(formatar_reais)
        st.metric("💰**Gasto Mensal Total com o 4° Andar**💰", formatar_reais(sum(gastos_4_andar.values())))
        st.dataframe(df_display_4_andar, use_container_width=True)

    with col5:
        st.subheader("📊**Gráfico de Gastos Mensais do 4° Andar**📊")
        fig_4_andar = px.pie(df_gastos_4_andar,
                            names="Tipo de gasto",
                            values="Valor Mensal (R$)",
                            title="Gastos Mensais do 4° Andar",
                            hole=0.3)
        fig_4_andar.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_4_andar, use_container_width=True)
with tab2:

    st.subheader("📦**Controle de Estoque Recepção - 3° Andar**📦")
    st.write("Este dashboard apresenta o controle de estoque da **Recepção do 3° andar**, incluindo materiais solicitados, gastos mensais e setores com mais solicitações.")
    st.write("-------------------------------")
    col1, col2 = st.columns(2)

    with col1:
    
        st.metric("💰**Gasto Mensal com Materiais**💰","R$ 7.364,02")
        st.subheader("📋**Top 10** Materiais mais solicitados")
        mais_solicitados = df.groupby(["Material","Mês"])["Quantidade"].sum().reset_index(name="Total")
        mais_solicitados= mais_solicitados.sort_values(by="Total", ascending=False).head(10)
        st.dataframe(mais_solicitados, use_container_width=True)

    with col2:
        solicitados_mensais = df.groupby("Mês")["Quantidade"].sum().reset_index(name="Total")
        st.subheader("📅**Solicitações Mensais**")
        fig = px.bar(solicitados_mensais,
                    x="Mês",
                    y="Total",
                    color="Mês",
                    title="Solicitações Mensais de Materiais")
        st.plotly_chart(fig, use_container_width=True)

    st.write("----------------------------")

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("📚**Setores com mais Materiais Solicitados**📚")

        setores_solicitados = df.groupby(["Setor",'Mês'])["Quantidade"].sum().reset_index(name="Total")
        setores_solicitados = setores_solicitados.sort_values(by="Total", ascending=False).head(5)
        st.dataframe(setores_solicitados, use_container_width=True)

    
    with col4:
        fig= px.pie(setores_solicitados,
                    names="Setor",
                    values="Total",
                    title="Distribuição de Materiais por Setor",
                    hole=0.3)
        st.plotly_chart(fig, use_container_width=True)


with tab3:
    st.subheader("🧼**Controle de Estoque Limpeza e Copa☕**")

    st.write("Este dashboard apresenta o controle de estoque de materiais de **Limpeza e da Copa**.")
    st.write("-------------------------------")

    gastos_limpeza = {
        'Quality': 15553.95,
        'Limpeza':12224.03
    }

    df_gastos_limpeza = pd.DataFrame(gastos_limpeza.items(), columns=['Setor', 'Gasto Mensal (R$)'])
    col1, col2 = st.columns(2)
    with col1:
        st.metric("💰**Gasto Mensal Total com Limpeza do 3° e 4° Andar**", formatar_reais(df_gastos_limpeza["Gasto Mensal (R$)"].sum()))
        df_display_limpeza = df_gastos_limpeza.copy()
        df_display_limpeza = df_display_limpeza.sort_values(by="Gasto Mensal (R$)", ascending=False)
        df_display_limpeza["Gasto Mensal (R$)"] = df_display_limpeza["Gasto Mensal (R$)"].apply(formatar_reais)
        st.subheader("💸**Tabela dos Gastos Mensais com Limpeza 3° Andar e 4° Andar**💸")
        st.dataframe(df_display_limpeza, use_container_width=True)

    with col2:
        st.metric("💰**Gasto Mensal Médio da Copa (Maio, Junho, Julho)**", formatar_reais(3975.22))
    st.write("-------------------------------")

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Gasto com a **Alvorada**")
        st.metric("**Total Gasto com a Alvorada em Junho (R$)**",formatar_reais(2776.65))
        dados_alvorada = {"Material":["AÇUCAR CRISTAL 2KG TAI",
                                    "ALCOOL 1 LITRO",
                                    "BLOQUEADOR AEROSSOL 100ML",
                                    "COADOR CAFÉ TEXTIL BOM LAR",
                                    "DESINF. AZULIM",
                                    "DESINF. USO GERAL 360ML",
                                    "ESPONJA D FACE TININDO 3M C/10UN",
                                    "LUVA FORRADA AMAR VERNIZ SLIM",
                                    "MEXEDOR DRINK GDE 11 CM 240UN",
                                    "MULTI USO CLORO AT TIR LIMO ZUPP 1X500ML",
                                    "MULTI USO LIMPEZA PESADA 1L YPE",
                                    "PANO M.USO AZUL",
                                    "PATO PURIFIC GERMINEX 500ML",
                                    "PURIFICADOR LEV&UZE 400ML",
                                    "REFIL P/AO 269ML",
                                    "SABAO EM PO BRILHANTE 4KG",
                                    "SACO ALVEJADO EXTRA 45X70CM",
                                    "SACO PTO 40L",
                                    "SACO PTO 110L",
                                    "SAPONACEO CREMOSO CIF 250ML",
                                    "TELA DESOD. P/MIC",
                                    "Chá"],
                          "Quantidade":[5,12,10,4,8,10,1,4,5,12,10,2,12,7,12,1,10,5,8,10,5,70],
                          "Valor Total (R$)":[43.25,65.28,164.10,80.36,180,200.6,5.47,18.60,49.6,55.32,115.1,143.14,138.96,58.31,339.12,38.94,55.2,59.05,274,74.6,13.3,604.35]}
        df_alvorada = pd.DataFrame(dados_alvorada)
        df_alvorada_col = df_alvorada.copy()
        df_alvorada_col["Valor Total (R$)"] = df_alvorada_col["Valor Total (R$)"].apply(formatar_reais)
        st.subheader("💸**Tabela de Gastos com a Alvorada em Junho**💸")
        st.dataframe(df_alvorada_col, use_container_width=True)

    with col4:
        st.subheader("Gastos com a **Brago**")
        dados_brago = {"Material":["Papel Higiênico","Papel Toalha","Sabão Líquido 475ml","Sabão Líquido 1000ml","Álcool em Gel 1000ml"],
                       "Quantidade":[20,44,20,4,10],
                       "Valor Total (R$)":[3953.4,7601.88,804.8,278.64,731.2]}
        df_brago = pd.DataFrame(dados_brago)
        df_brago_col = df_brago.copy()
        df_brago_col["Valor Total (R$)"] = df_brago_col["Valor Total (R$)"].apply(formatar_reais)
        st.metric("**Total Gasto com a Brago em Junho (R$)**",formatar_reais(13369.92))
        st.subheader("💸**Tabela de Gastos com a Brago em Junho**💸")
        st.dataframe(df_brago_col, use_container_width=True)
        st.markdown("**OBS: Em Julho, não houve compras com a Brago.**")

        dados_brago_agosto = {"Material":["Mini Higienizador Assento LIQ. 475ML","Papel Higiênico","Papel Toalha","Perfumador Aerossol Haroma 100ml", "Sabão Liq. Extra SUAVE MINI 475ml"],
                              "Quantidade":[10,33,30,10,30],
                              "Valor Total (R$)":[956.90,3953.4,5183.10,449.00,1207.20]}
        df_brago_agosto = pd.DataFrame(dados_brago_agosto)
        df_brago_agosto_col = df_brago_agosto.copy()
        df_brago_agosto_col["Valor Total (R$)"] = df_brago_agosto_col["Valor Total (R$)"].apply(formatar_reais)
        st.metric("**Total Gasto com a Brago em Agosto (R$)**",formatar_reais(11749.60))
        st.subheader("💸**Tabela de Gastos com a Brago em Agosto**💸")
        st.dataframe(df_brago_agosto_col, use_container_width=True)
with tab4:
    st.subheader("💰**Gastos Joinville**💰")
    st.write("Este dashboard apresenta os gastos mensais de **Joinville**.")
    st.write("-------------------------------")

    dados_joinville = {"Aluguel": 7500,
                       "Condomínio(Conta de Água incluída)": 1837.23,
                       "Limpeza": 5865,
                       "Materiais da Copa/Escritório/Limpeza": 1883.05,
                       "Conta de Luz": 1250,
                       "Internet": 300
                       }
    df_joinville = pd.DataFrame(dados_joinville.items(), columns=["Tipo de Gasto", "Valor Mensal (R$)"])
    df_joinville = df_joinville.sort_values(by="Valor Mensal (R$)", ascending=False)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("**Gasto Mensal Total Joinville**", formatar_reais(df_joinville["Valor Mensal (R$)"].sum()))
    st.subheader("📊**Gráfico de Gastos Mensais Joinville**📊")
    fig_joinville = px.bar(
                            df_joinville,
                              x="Tipo de Gasto",
                              y="Valor Mensal (R$)",
                              color="Tipo de Gasto",
                              labels={"x": "Tipo de Gasto", "y": "Valor (R$)"},
                              title="Gastos Mensais Joinville")
    st.plotly_chart(fig_joinville, use_container_width=True)
    st.write("-------------------------------")
    df_joinville["Valor Mensal (R$)"] = df_joinville["Valor Mensal (R$)"].apply(formatar_reais)
    st.subheader("💸**Tabela de Gastos Mensais Joinville**💸")
    st.dataframe(df_joinville, use_container_width=True)

    with col2:
        st.metric("**Valor da Cláusula Contratual do Aluguel** (3 vezes o **Valor** do Aluguel)", formatar_reais(22500))

    with col3:
        st.metric("**Valor do Frete para Transportar os Equipamentos**", formatar_reais(23000))