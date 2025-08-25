import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# Configuração da página
st.set_page_config(
    page_title="Controle de Estoque Recepção", 
    page_icon="📦", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título da aplicação
st.title("📦 Controle de Estoque - Recepção")
st.markdown("---")

# Função para carregar a planilha existente
@st.cache_data
def carregar_planilha():
    arquivo_excel = "Controle de estoque recepção.xlsx"
    
    try:
        df = pd.read_excel(arquivo_excel)
        # Garante que a coluna de valor seja tratada como string para preservar fórmulas
        if 'valor' in df.columns:
            df['valor'] = df['valor'].astype(str)
    except FileNotFoundError:
        # Se não existir, cria um novo DataFrame com a estrutura da planilha
        df = pd.DataFrame(columns=[
            'Nome do colaborador', 'Setor', 'Material', 'Quantidade', 'Mês', 'valor'
        ])
    
    return df, arquivo_excel

# Função para salvar na planilha
def salvar_planilha(df, arquivo_excel):
    # Preserva fórmulas existentes na coluna 'valor'
    if 'valor' in df.columns:
        for idx, val in enumerate(df['valor']):
            if isinstance(val, str) and val.startswith('='):
                # Mantém as fórmulas intactas
                pass
            elif pd.notna(val) and val != '':
                # Converte valores numéricos para float
                try:
                    df.at[idx, 'valor'] = float(val)
                except (ValueError, TypeError):
                    pass
    
    df.to_excel(arquivo_excel, index=False)
    st.cache_data.clear()

# Carrega a planilha
df, arquivo_excel = carregar_planilha()

# Layout em colunas



st.subheader("➕ Novo Registro de Estoque")
    
with st.form("form_estoque"):
        # Campos do formulário baseados na planilha
    colaborador = st.text_input("Nome do colaborador:", max_chars=100)
        
    setor = st.selectbox(
            "Setor:",
            ["BackOffice", "TI", "Comercial", "Financeiro", "GG", "Marketing", 
             "Governança", "Processos e Qualidade", "Tesouraria", "Be Civis", "Cordenação Back Office"]
)
        
    material = st.text_input("Material:", max_chars=100)
        
    quantidade = st.number_input("Quantidade:", min_value=1, max_value=1000, step=1, value=1)
        
        # Input para mês apenas (como na planilha)
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes = st.selectbox("Mês:", meses, index=datetime.now().month-1)
        
    valor = st.text_input("Valor (opcional, use = para fórmulas):", value="")
        
    submitted = st.form_submit_button("✅ Registrar Movimento")
        
    if submitted:
        if colaborador and material:
                # Cria novo registro com a mesma estrutura da planilha
            novo_registro = {
                    'Nome do colaborador': colaborador,
                    'Setor': setor,
                    'Material': material,
                    'Quantidade': quantidade,
                    'Mês': mes,
                    'valor': valor if valor else ""
                }
                
                # Adiciona ao DataFrame
            df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                
                # Salva na planilha
            salvar_planilha(df, arquivo_excel)
                
            st.success("✅ Registro salvo com sucesso!")
            st.balloons()
        else:
            st.error("❌ Preencha todos os campos obrigatórios!")

# with col2:
#     st.subheader("📊 Estatísticas do Estoque")
    
#     if not df.empty:
#         # Estatísticas básicas
#         total_itens = df['Quantidade'].sum()
#         total_registros = len(df)
#         meses_unicos = df['Mês'].nunique()
        
#         col1, col2, col3 = st.columns(3)
        
#         with col1:
#             st.metric("Total de Itens", f"{total_itens:,}")
        
#         with col2:
#             st.metric("Total de Registros", total_registros)
        
#         with col3:
#             st.metric("Meses com Registros", meses_unicos)
        
#         # Gráfico de itens por setor
#         itens_por_setor = df.groupby('Setor')['Quantidade'].sum().reset_index()
#         fig = px.bar(
#             itens_por_setor.sort_values('Quantidade', ascending=False),
#             x='Setor',
#             y='Quantidade',
#             title="Itens por Setor",
#             color='Quantidade',
#             color_continuous_scale='Blues'
#         )
#         fig.update_layout(xaxis_tickangle=-45)
#         st.plotly_chart(fig, use_container_width=True)
        
#         # Gráfico de itens por mês
#         itens_por_mes = df.groupby('Mês')['Quantidade'].sum().reset_index()
#         # Ordenar meses cronologicamente
#         ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
#                       "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
#         itens_por_mes['Mês'] = pd.Categorical(itens_por_mes['Mês'], categories=ordem_meses, ordered=True)
#         itens_por_mes = itens_por_mes.sort_values('Mês')
        
#         fig2 = px.line(
#             itens_por_mes, 
#             x='Mês', 
#             y='Quantidade', 
#             title='Itens por Mês',
#             markers=True
#         )
#         st.plotly_chart(fig2, use_container_width=True)

# Seção de visualização de dados
st.markdown("---")
st.subheader("📋 Dados da Planilha")

if not df.empty:
    # Filtros
    col1, col2, col3 = st.columns(3)
    
    with col1:
        filtro_setor = st.selectbox(
            "Filtrar por setor:",
            ["Todos"] + list(df['Setor'].unique())
        )
    
    with col2:
        filtro_mes = st.selectbox(
            "Filtrar por mês:",
            ["Todos"] + list(df['Mês'].unique())
        )
    
    with col3:
        filtro_material = st.selectbox(
            "Filtrar por material:",
            ["Todos"] + list(df['Material'].unique())
        )
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    if filtro_setor != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Setor'] == filtro_setor]
    
    if filtro_mes != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Mês'] == filtro_mes]
    
    if filtro_material != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Material'] == filtro_material]
    
    # Exibe dados filtrados
    st.dataframe(
        df_filtrado, 
        use_container_width=True,
        hide_index=True
    )
    
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
    3. Clique em 'Registrar Movimento'
    4. Os dados serão salvos na planilha Excel
    5. Use os filtros para visualizar dados específicos
    """)
    
    st.markdown("---")
    
    if not df.empty:
        st.markdown("**📈 Estatísticas Gerais:**")
        ultimo_registro = df.iloc[-1]['Nome do colaborador'] if 'Nome do colaborador' in df.columns else "N/A"
        ultimo_mes = df.iloc[-1]['Mês'] if 'Mês' in df.columns else "N/A"
        
        st.write(f"Último registro: **{ultimo_registro}**")
        st.write(f"Último mês: **{ultimo_mes}**")
        
        # Materiais mais retirados
        st.markdown("---")
        st.markdown("**📦 Materiais Mais Retirados:**")
        top_materiais = df.groupby('Material')['Quantidade'].sum().nlargest(5)
        for material, qtd in top_materiais.items():
            st.write(f"- {material}: {qtd} unidades")





