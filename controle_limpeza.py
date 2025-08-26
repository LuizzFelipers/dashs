# import pandas as pd
# import streamlit as st
# import plotly.express as px
# from datetime import datetime
# import os

# st.set_page_config(page_title="Controle de Limpeza",
#                    layout="wide",
#                    page_icon="🧼",
#                    initial_sidebar_state="expanded")

# st.title("🧼Controle de Limpeza🧼")
# st.markdown("-------")


# df = pd.DataFrame()
# arquivo_excel = "produtos_limpeza.xlsx"

# col1, col2 = st.columns(2)

# with col1:
#     st.subheader("➕ Novo Registro de Saída")
#     with st.form("form_novo_registro"):
       
#         colaborador = st.text_input("Colaborador")
#         produto = st.selectbox("Produto:",[
#                                         "ÁGUA SANITÁRIA KI JÓIA 5L",
#                                         "ALCOOL DESINF LIQ 70",
#                                         "ÁLCOOL EM GEL TORK 1000L",
#                                         "BRILHA INOX AZULIM GATILHO",
#                                         "BUCHA FIBRA VERDE DE LIMPEZA PESADA - USO GERAL",
#                                         "COLETOR DE ABSORVENTE FEMININO CAIXA COM 18 UN",
#                                         "DESINFERANTE AEROSSOL - LYSOFORM",
#                                         "DESINFETANTE AZULIM 5L",
#                                         "DESINFETANTE PARA SUPERFÍCIES", 
#                                         "DETERGENTE LÍQUIDO LAVA LOUÇAS",
#                                         "ESPONJA MULTIUSO DUPLA FACE", 
#                                         "GLADE 269ML",
#                                         "INSERTICIDA AEROSOOL 360ML BAYGON",
#                                         "LIMPADOR LIMPEZA PESADA",
#                                         "LIMPADOR MULTIUSO 500ML",
#                                         "LIMPADOR MULTIUSO CIF",
#                                         "LIMPADOR PATO PURIFIC LAVANDA",
#                                         "ODORIZADOR AEROSSOL 360ML",
#                                         "ODORIZADOR DE AMBIENTE LAVANDA FLORAL 400ML",
#                                         "PAPEL HIGIENICO FOLHA SIMPLES 8 ROLOS 10CMX400M",
#                                         "PAPEL TOALHA INT ADV FS 16X260",
#                                         "SABONETE LIQUIDO MÃOS SOAP - TAM G",
#                                         "SABONETE LIQUIDO MÃOS SOAP  - TAM P", 
#                                         "SACOS DE LIXO 110L PA 100 UN",
#                                         "TOILET SEAT CLEANER",
#                                         "ZUPP CLORO ATIVO 500 ML",
#                                         "SACOS DE LIXO 40 LITROS",
#                                         "TELA PARA MICTORIO"
#                                         ])
#         meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
#         mes = st.selectbox("Mês", meses)
#         quantidade = st.number_input("Quantidade", min_value=1, step=1)
#         andar = st.selectbox("Andar",["3° Andar","4° Andar"])

#         submitted = st.form_submit_button("✅ Adicionar Registro")
#         if submitted:
#             if colaborador and produto:
#                     # Cria novo registro com a mesma estrutura da planilha
#                 novo_registro = {
#                         'Nome do colaborador': colaborador,
#                         'Produto': produto,
#                         'Andar': andar,
#                         'Quantidade': quantidade,
#                         'Mês': mes
#                     }
                    
#                     # Adiciona ao DataFrame
#                 df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                    
#                     # Salva na planilha
#                 (df, arquivo_excel)
                    
#                 st.success("✅ Registro salvo com sucesso!")
#                 st.balloons()
#             else:
#                 st.error("❌ Preencha todos os campos obrigatórios!")        

# with col2:
#     st.subheader("📊 Estatísticas do Estoque")

#     if not df.empty:
#         # Estatísticas Gerais
#         total_itens = df['Quantidade'].sum()
#         total_registros = len(df)

#         col3, col4 = st.columns(2)
#         with col3:
#             st.metric("Total de Itens", total_itens)
#         with col4:
#             st.metric("Total de Registros", total_registros)

#         st.markdown("-----")

#         # Gráfico de Barras - Itens por Mês
#         itens_por_mes = df.groupby('Mês')['Quantidade'].sum().reindex(meses)
#         fig_mes = px.bar(itens_por_mes, x=itens_por_mes.index, y='Quantidade', title="Itens Saídos por Mês", labels={"Mês": "Mês", "Quantidade": "Total de Itens"})
#         st.plotly_chart(fig_mes, use_container_width=True)

#         # Gráfico de Barras - Itens por Andar
#         itens_por_andar = df.groupby('Andar')['Quantidade'].sum()
#         fig_andar = px.bar(itens_por_andar, x=itens_por_andar.index, y='Quantidade', title="Itens Saídos por Andar", labels={"Andar": "Andar", "Quantidade": "Total de Itens"})
#         st.plotly_chart(fig_andar, use_container_width=True)

#         # Tabela Resumo por Item
#         resumo_item = df.groupby('Produto')['Quantidade'].sum().reset_index().sort_values(by='Quantidade', ascending=False)
#         st.subheader("Resumo de Itens")
#         st.dataframe(resumo_item, use_container_width=True)

# with st.sidebar:
#     st.header("ℹ️Informações sobre o Sistema")
#     st.markdown("""
#                 1. Preencha todos os campos do formulário para adicionar um novo registro de saída de produto.
#                 2. Utilize o campo "Quantidade" para especificar o número de unidades retiradas)
#                 3. As estatísticas e gráficos serão atualizados automaticamente com base nos dados inseridos.
#                 """)
#     st.markdown("-------")

#     if not df.empty:
#         st.markdown("**📈 Estatísticas Gerais:**")
#         ultimo_registro = df.iloc[-1]['Nome do colaborador'] if 'Nome do colaborador' in df.columns else "N/A"
#         ultimo_mes = df.iloc[-1]['Mês'] if 'Mês' in df.columns else "N/A"
        
#         st.write(f"Último registro: **{ultimo_registro}**")
#         st.write(f"Último mês: **{ultimo_mes}**")
        
#         # Materiais mais retirados
#         st.markdown("---")
#         st.markdown("**📦 Materiais Mais Retirados:**")
#         top_materiais = df.groupby('Produto')['Quantidade'].sum().nlargest(5)
#         for material, qtd in top_materiais.items():
#             st.write(f"- {material}: {qtd} unidades")

# st.markdown("-------")
# st.subheader("📋Dados da Planilha")

# with open(arquivo_excel, "rb") as file:
#     st.download_button(
#         label="📥 Baixar Planilha Excel",
#         data=file,
#         file_name="produtos_limpeza.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#     )

# import pandas as pd
# import streamlit as st
# import plotly.express as px
# from datetime import datetime
# import os

# # Configuração da página
# st.set_page_config(
#     page_title="Controle de Limpeza",
#     page_icon="🧼", 
#     layout="wide",
#     initial_sidebar_state="expanded"
# )

# # Título da aplicação
# st.title("🧼 Controle de Limpeza 🧼")
# st.markdown("---")

# # Função para carregar a planilha existente
# @st.cache_data
# def carregar_planilha():
#     arquivo_excel = "produtos_limpeza.xlsx"
    
#     try:
#         df = pd.read_excel(arquivo_excel)
#     except FileNotFoundError:
#         # Se não existir, cria um novo DataFrame com a estrutura da planilha
#         df = pd.DataFrame(columns=[
#             'Nome do colaborador', 'Produto', 'Andar', 'Quantidade', 'Mês'
#         ])
    
#     return df, arquivo_excel

# # Função para salvar na planilha
# def salvar_planilha(df, arquivo_excel):
#     df.to_excel(arquivo_excel, index=False)
#     st.cache_data.clear()

# # Carrega a planilha
# df, arquivo_excel = carregar_planilha()

# # Layout em colunas
# col1, col2 = st.columns(2)

# with col1:
#     st.subheader("➕ Novo Registro de Saída")
    
#     with st.form("form_novo_registro"):
#         colaborador = st.text_input("Colaborador:", max_chars=100)
        
#         produto = st.selectbox("Produto:", [
#             "ÁGUA SANITÁRIA KI JÓIA 5L",
#             "ALCOOL DESINF LIQ 70",
#             "ÁLCOOL EM GEL TORK 1000L",
#             "BRILHA INOX AZULIM GATILHO",
#             "BUCHA FIBRA VERDE DE LIMPEZA PESADA - USO GERAL",
#             "COLETOR DE ABSORVENTE FEMININO CAIXA COM 18 UN",
#             "DESINFERANTE AEROSSOL - LYSOFORM",
#             "DESINFETANTE AZULIM 5L",
#             "DESINFETANTE PARA SUPERFÍCIES", 
#             "DETERGENTE LÍQUIDO LAVA LOUÇAS",
#             "ESPONJA MULTIUSO DUPLA FACE", 
#             "GLADE 269ML",
#             "INSERTICIDA AEROSOOL 360ML BAYGON",
#             "LIMPADOR LIMPEZA PESADA",
#             "LIMPADOR MULTIUSO 500ML",
#             "LIMPADOR MULTIUSO CIF",
#             "LIMPADOR PATO PURIFIC LAVANDA",
#             "ODORIZADOR AEROSSOL 360ML",
#             "ODORIZADOR DE AMBIENTE LAVANDA FLORAL 400ML",
#             "PAPEL HIGIENICO FOLHA SIMPLES 8 ROLOS 10CMX400M",
#             "PAPEL TOALHA INT ADV FS 16X260",
#             "SABONETE LIQUIDO MÃOS SOAP - TAM G",
#             "SABONETE LIQUIDO MÃOS SOAP  - TAM P", 
#             "SACOS DE LIXO 110L PA 100 UN",
#             "TOILET SEAT CLEANER",
#             "ZUPP CLORO ATIVO 500 ML",
#             "SACOS DE LIXO 40 LITROS",
#             "TELA PARA MICTORIO"
#         ])
        
#         meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
#                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
#         mes = st.selectbox("Mês:", meses, index=datetime.now().month-1)
        
#         quantidade = st.number_input("Quantidade:", min_value=1, step=1, value=1)
        
#         andar = st.selectbox("Andar:", ["3° Andar", "4° Andar"])
        
#         submitted = st.form_submit_button("✅ Adicionar Registro")
        
#         if submitted:
#             if colaborador and produto:
#                 # Cria novo registro com a mesma estrutura da planilha
#                 novo_registro = {
#                     'Nome do colaborador': colaborador,
#                     'Produto': produto,
#                     'Andar': andar,
#                     'Quantidade': quantidade,
#                     'Mês': mes
#                 }
                
#                 # Adiciona ao DataFrame
#                 df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                
#                 # Salva na planilha
#                 salvar_planilha(df, arquivo_excel)
                
#                 st.success("✅ Registro salvo com sucesso!")
#                 st.balloons()
#             else:
#                 st.error("❌ Preencha todos os campos obrigatórios!")

# with col2:
#     st.subheader("📊 Estatísticas do Estoque")
    
#     if not df.empty:
#         # Estatísticas básicas
#         total_itens = df['Quantidade'].sum()
#         total_registros = len(df)
#         meses_unicos = df['Mês'].nunique()
        
#         col3, col4, col5 = st.columns(3)
        
#         with col3:
#             st.metric("Total de Itens", f"{total_itens:,}")
        
#         with col4:
#             st.metric("Total de Registros", total_registros)
        
#         with col5:
#             st.metric("Meses com Registros", meses_unicos)
        
#         # Gráfico de itens por mês
#         itens_por_mes = df.groupby('Mês')['Quantidade'].sum().reset_index()
#         # Ordenar meses cronologicamente
#         ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
#                       "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
#         itens_por_mes['Mês'] = pd.Categorical(itens_por_mes['Mês'], categories=ordem_meses, ordered=True)
#         itens_por_mes = itens_por_mes.sort_values('Mês')
        
#         fig_mes = px.line(
#             itens_por_mes, 
#             x='Mês', 
#             y='Quantidade', 
#             title='Itens por Mês',
#             markers=True
#         )
#         st.plotly_chart(fig_mes, use_container_width=True)
        
#         # Gráfico de itens por andar
#         itens_por_andar = df.groupby('Andar')['Quantidade'].sum().reset_index()
#         fig_andar = px.bar(
#             itens_por_andar,
#             x='Andar',
#             y='Quantidade',
#             title="Itens por Andar",
#             color='Quantidade',
#             color_continuous_scale='Greens'
#         )
#         st.plotly_chart(fig_andar, use_container_width=True)

# # Seção de visualização de dados
# st.markdown("---")
# st.subheader("📋 Dados da Planilha")

# if not df.empty:
#     # Filtros
#     col1, col2, col3 = st.columns(3)
    
#     with col1:
#         filtro_andar = st.selectbox(
#             "Filtrar por andar:",
#             ["Todos"] + list(df['Andar'].unique())
#         )
    
#     with col2:
#         filtro_mes = st.selectbox(
#             "Filtrar por mês:",
#             ["Todos"] + list(df['Mês'].unique())
#         )
    
#     with col3:
#         filtro_produto = st.selectbox(
#             "Filtrar por produto:",
#             ["Todos"] + list(df['Produto'].unique())
#         )
    
#     # Aplicar filtros
#     df_filtrado = df.copy()
    
#     if filtro_andar != "Todos":
#         df_filtrado = df_filtrado[df_filtrado['Andar'] == filtro_andar]
    
#     if filtro_mes != "Todos":
#         df_filtrado = df_filtrado[df_filtrado['Mês'] == filtro_mes]
    
#     if filtro_produto != "Todos":
#         df_filtrado = df_filtrado[df_filtrado['Produto'] == filtro_produto]
    
#     # Exibe dados filtrados
#     st.dataframe(
#         df_filtrado, 
#         use_container_width=True,
#         hide_index=True
#     )
    
#     # Opção para baixar a planilha
#     with open(arquivo_excel, "rb") as file:
#         st.download_button(
#             label="📥 Baixar Planilha Completa",
#             data=file,
#             file_name=arquivo_excel,
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             use_container_width=True
#         )
# else:
#     st.info("📝 Nenhum registro encontrado. Use o formulário acima para adicionar movimentos de estoque.")

# # Sidebar com informações
# with st.sidebar:
#     st.header("ℹ️ Informações do Sistema")
    
#     st.markdown("""
#     **Como usar:**
#     1. Preencha o formulário à esquerda
#     2. Os campos devem corresponder aos da planilha
#     3. Clique em 'Adicionar Registro'
#     4. Os dados serão salvos na planilha Excel
#     5. Use os filtros para visualizar dados específicos
#     """)
    
#     st.markdown("---")
    
#     if not df.empty:
#         st.markdown("**📈 Estatísticas Gerais:**")
#         ultimo_registro = df.iloc[-1]['Nome do colaborador'] if 'Nome do colaborador' in df.columns else "N/A"
#         ultimo_mes = df.iloc[-1]['Mês'] if 'Mês' in df.columns else "N/A"
        
#         st.write(f"Último registro: **{ultimo_registro}**")
#         st.write(f"Último mês: **{ultimo_mes}**")
        
#         # Produtos mais retirados
#         st.markdown("---")
#         st.markdown("**🧼 Produtos Mais Retirados:**")
#         top_produtos = df.groupby('Produto')['Quantidade'].sum().nlargest(5)
#         for produto, qtd in top_produtos.items():
#             st.write(f"- {produto}: {qtd} unidades")

import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime
import os

# Configuração da página
st.set_page_config(
    page_title="Controle de Limpeza",
    page_icon="🧼", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título da aplicação
st.title("🧼 Controle de Limpeza 🧼")
st.markdown("---")

# Função para carregar a planilha existente
@st.cache_data
def carregar_planilha():
    arquivo_excel = "produtos_limpeza.xlsx"
    
    try:
        df = pd.read_excel(arquivo_excel)
    except FileNotFoundError:
        # Se não existir, cria um novo DataFrame com a estrutura da planilha
        df = pd.DataFrame(columns=[
            'Nome do colaborador', 'Produto', 'Andar', 'Quantidade', 'Mês'
        ])
    
    return df, arquivo_excel

# Função para salvar na planilha
def salvar_planilha(df, arquivo_excel):
    df.to_excel(arquivo_excel, index=False)
    st.cache_data.clear()

# Carrega a planilha
df, arquivo_excel = carregar_planilha()

# Layout em colunas
col1, col2 = st.columns(2)

with col1:
    st.subheader("➕ Novo Registro de Saída")
    
    with st.form("form_novo_registro"):
        colaborador = st.text_input("Colaborador:", max_chars=100)
        
        produto = st.selectbox("Produto:", [
            "ÁGUA SANITÁRIA KI JÓIA 5L",
            "ALCOOL DESINF LIQ 70",
            "ÁLCOOL EM GEL TORK 1000L",
            "BRILHA INOX AZULIM GATILHO",
            "BUCHA FIBRA VERDE DE LIMPEZA PESADA - USO GERAL",
            "COLETOR DE ABSORVENTE FEMININO CAIXA COM 18 UN",
            "DESINFERANTE AEROSSOL - LYSOFORM",
            "DESINFETANTE AZULIM 5L",
            "DESINFETANTE PARA SUPERFÍCIES", 
            "DETERGENTE LÍQUIDO LAVA LOUÇAS",
            "ESPONJA MULTIUSO DUPLA FACE", 
            "GLADE 269ML",
            "INSERTICIDA AEROSOOL 360ML BAYGON",
            "LIMPADOR LIMPEZA PESADA",
            "LIMPADOR MULTIUSO 500ML",
            "LIMPADOR MULTIUSO CIF",
            "LIMPADOR PATO PURIFIC LAVANDA",
            "ODORIZADOR AEROSSOL 360ML",
            "ODORIZADOR DE AMBIENTE LAVANDA FLORAL 400ML",
            "PAPEL HIGIENICO FOLHA SIMPLES 8 ROLOS 10CMX400M",
            "PAPEL TOALHA INT ADV FS 16X260",
            "SABONETE LIQUIDO MÃOS SOAP - TAM G",
            "SABONETE LIQUIDO MÃOS SOAP  - TAM P", 
            "SACOS DE LIXO 110L PA 100 UN",
            "TOILET SEAT CLEANER",
            "ZUPP CLORO ATIVO 500 ML",
            "SACOS DE LIXO 40 LITROS",
            "TELA PARA MICTORIO"
        ])
        
        meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes = st.selectbox("Mês:", meses, index=datetime.now().month-1)
        
        quantidade = st.number_input("Quantidade:", min_value=1, step=1, value=1)
        
        andar = st.selectbox("Andar:", ["3° Andar", "4° Andar"])
        
        submitted = st.form_submit_button("✅ Adicionar Registro")
        
        if submitted:
            if colaborador and produto:
                # Cria novo registro com a mesma estrutura da planilha
                novo_registro = {
                    'Nome do colaborador': colaborador,
                    'Produto': produto,
                    'Andar': andar,
                    'Quantidade': quantidade,
                    'Mês': mes
                }
                
                # Adiciona ao DataFrame
                df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                
                # Salva na planilha
                salvar_planilha(df, arquivo_excel)
                
                st.success("✅ Registro salvo com sucesso!")
                st.balloons()
            else:
                st.error("❌ Preencha todos os campos obrigatórios!")

with col2:
    st.subheader("📊 Estatísticas do Estoque")
    
    if not df.empty:
        # Estatísticas básicas
        total_itens = df['Quantidade'].sum()
        total_registros = len(df)
        meses_unicos = df['Mês'].nunique()
        
        col3, col4, col5 = st.columns(3)
        
        with col3:
            st.metric("Total de Itens", f"{total_itens:,}")
        
        with col4:
            st.metric("Total de Registros", total_registros)
        
        with col5:
            st.metric("Meses com Registros", meses_unicos)
        
        # Gráfico de itens por mês
        itens_por_mes = df.groupby('Mês')['Quantidade'].sum().reset_index()
        # Ordenar meses cronologicamente
        ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                      "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        itens_por_mes['Mês'] = pd.Categorical(itens_por_mes['Mês'], categories=ordem_meses, ordered=True)
        itens_por_mes = itens_por_mes.sort_values('Mês')
        
        fig_mes = px.line(
            itens_por_mes, 
            x='Mês', 
            y='Quantidade', 
            title='Itens por Mês',
            markers=True
        )
        st.plotly_chart(fig_mes, use_container_width=True)
        
        # Gráfico de itens por andar
        itens_por_andar = df.groupby('Andar')['Quantidade'].sum().reset_index()
        fig_andar = px.bar(
            itens_por_andar,
            x='Andar',
            y='Quantidade',
            title="Itens por Andar",
            color='Quantidade',
            color_continuous_scale='Greens'
        )
        st.plotly_chart(fig_andar, use_container_width=True)

# Seção de visualização de dados
st.markdown("---")
st.subheader("📋 Dados da Planilha")

if not df.empty:
    # Filtros
    col1, col2, col3 = st.columns(3)
    
    with col1:
        filtro_andar = st.selectbox(
            "Filtrar por andar:",
            ["Todos"] + list(df['Andar'].unique())
        )
    
    with col2:
        filtro_mes = st.selectbox(
            "Filtrar por mês:",
            ["Todos"] + list(df['Mês'].unique())
        )
    
    with col3:
        filtro_produto = st.selectbox(
            "Filtrar por produto:",
            ["Todos"] + list(df['Produto'].unique())
        )
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    if filtro_andar != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Andar'] == filtro_andar]
    
    if filtro_mes != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Mês'] == filtro_mes]
    
    if filtro_produto != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Produto'] == filtro_produto]
    
    # Exibe dados filtrados com checkbox para seleção
    st.markdown("**Selecione os registros para exclusão:**")
    
    # Adiciona checkbox para cada linha
    indices_para_excluir = []
    for idx, row in df_filtrado.iterrows():
        col1, col2, col3, col4, col5 = st.columns([1, 3, 3, 2, 2])
        with col1:
            excluir = st.checkbox("Excluir", key=f"excluir_{idx}")
        with col2:
            st.write(row['Nome do colaborador'])
        with col3:
            st.write(row['Produto'])
        with col4:
            st.write(row['Andar'])
        with col5:
            st.write(f"{row['Quantidade']} unid.")
        
        if excluir:
            indices_para_excluir.append(idx)
    
    # Botão para confirmar exclusão
    if indices_para_excluir:
        if st.button("🗑️ Excluir Registros Selecionados", type="primary"):
            # Remove os registros selecionados
            df = df.drop(indices_para_excluir).reset_index(drop=True)
            salvar_planilha(df, arquivo_excel)
            st.success(f"✅ {len(indices_para_excluir)} registro(s) excluído(s) com sucesso!")
            st.rerun()
    
    # Opção para baixar a planilha
    with open(arquivo_excel, "rb") as file:
        st.download_button(
            label="📥 Baixar Planilha Completa",
            data=file,
            file_name=arquivo_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("📝 Nenhum registro encontrado. Use o formulário acima para adicionar movimentos de estoque.")

# Sidebar com informações
with st.sidebar:
    st.header("ℹ️ Informações do Sistema")
    
    st.markdown("""
    **Como usar:**
    1. Preencha o formulário à esquerda
    2. Os campos devem corresponder aos da planilha
    3. Clique em 'Adicionar Registro'
    4. Os dados serão salvos na planilha Excel
    5. Use os filtros para visualizar dados específicos
    6. Marque os checkboxes e clique em 'Excluir Registros' para remover
    """)
    
    st.markdown("---")
    
    if not df.empty:
        st.markdown("**📈 Estatísticas Gerais:**")
        ultimo_registro = df.iloc[-1]['Nome do colaborador'] if 'Nome do colaborador' in df.columns else "N/A"
        ultimo_mes = df.iloc[-1]['Mês'] if 'Mês' in df.columns else "N/A"
        
        st.write(f"Último registro: **{ultimo_registro}**")
        st.write(f"Último mês: **{ultimo_mes}**")
        
        # Produtos mais retirados
        st.markdown("---")
        st.markdown("**🧼 Produtos Mais Retirados:**")
        top_produtos = df.groupby('Produto')['Quantidade'].sum().nlargest(5)
        for produto, qtd in top_produtos.items():
            st.write(f"- {produto}: {qtd} unidades")