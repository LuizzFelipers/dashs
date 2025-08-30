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

# Função para carregar a planilha existente SEM cache
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

# Carrega a planilha - SEM cache para garantir que sempre carrega os dados mais recentes
df, arquivo_excel = carregar_planilha()

# Garantir que a coluna Material seja tratada corretamente
if not df.empty and 'Material' in df.columns:
    df["Material"] = df["Material"].str.strip()

# Usar session_state para manter os dados durante a sessão
if 'dados_estoque' not in st.session_state:
    st.session_state.dados_estoque = df.copy()

# Layout em colunas
col1, col2 = st.columns(2)

with col1:
    st.subheader("➕ Novo Registro de Estoque")
        
    with st.form("form_estoque"):
        # Campos do formulário baseados na planilha
        colaborador = st.text_input("Nome do colaborador:", max_chars=100)
            
        setor = st.selectbox(
            "Setor:",
            ["BackOffice", "TI", "Comercial", "Financeiro", "GG", "Marketing", 
            "Governança", "Processos e Qualidade", "Tesouraria", "Be Civis", "Cordenação Back Office"]
        )
            
        material = st.selectbox("Material:", [
            "Pilha AA", "Pilha AAA", "Saco Plástico PP 240mm X 320mm: 50 unidades", 
            "RESMA DE PAPEL A4 - 500FLS", "Post-it Pequeno", "Post-It Grande", 
            "Caderno", "Caneta Azul", "Caneta Preta", "Caneta Vermelha", 
            "Caneta Colorida", "Agenda", "Lápis", "Borracha", "Apontador", 
            "Marca-Texto", "Lapizeira", "Grafite", "Corretivo", "Clip", "Pilha C", 
            "Grampo", "Grampeador", "Apoio de Pé", "Apoio de Notebook", 
            "Papel Timbrado", "COLA EM BASTAO 40G", "TESOUSA", "CURATIVO - JOELHO", 
            "CURATIVO TRANSPARENTE", "FITA CREPE - 12mmX30M", "Álcool em Gel 50g", 
            "Extrator de Grampo Galvanizado", "Envelope Personalizado", 
            "Envelope A4 Saco Kraf Pardo 240x340 cm", "APOIO/SUPORTE DE MONITOR", 
            "Cartão Presente - SPOTIFY PREMIUM", 
            "Etiqueta Adesiva A4 350 - 100 Folhas - 3000 Etiq.", "AGUA COM GAS", 
            "ÁGUA MINERAL 250ML", "REDBUD - ZERO", "LEITE", "REFRIGERANTE - SCHWEPPES", 
            "CERVEJA", "REFRIGERANTE COCA-COLA 310ML", "REFRIGERANTE COCA-COLA ZERO 220ML", 
            "REFRIGERANTE COCA-COLA ZERO 310ML", "REFRIGERANTE FANTA UVA", 
            "REFRIGERANTE GUARANA 350ML", "SUCO 290ML", "PAPEL TIMBRADO", 
            "PAPEL TOALHA", "MÁSCARA", 
            "REABASTECEDOR PARA PINCEL DE QUADRO BRANCO - PRETO", 
            "REABASTECEDOR PARA PINCEL DE QUADRO BRANCO - AZUL", 
            "REABASTECEDOR PARA PINCEL DE QUADRO BRANCO - VERMELHO", 
            "MINI-GRAMPEADOR GENMES P/25FLS CORES DIVERSAS", 
            "PASTA PLASTICA EM L PP 0,15 A4 - TRANSPARENTE", 
            "FITA ADESIVA PP 12MMX30M DUREX HB0041744262 3M PT 10 UM", 
            "FITA CREPE - 12mmX30M", 
            "FITA DUPLA FACE 3M SCOTCH FIXA FORTE FIXAÇÃO EXTREMA - 24mm x 2"
        ])
            
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
                    
                # Adiciona ao DataFrame na session_state
                novo_df = pd.concat([st.session_state.dados_estoque, pd.DataFrame([novo_registro])], ignore_index=True)
                st.session_state.dados_estoque = novo_df
                    
                # Salva na planilha
                salvar_planilha(st.session_state.dados_estoque, arquivo_excel)
                    
                st.success("✅ Registro salvo com sucesso!")
                st.balloons()
            else:
                st.error("❌ Preencha todos os campos obrigatórios!")

with col2:
    st.subheader("📊 Estatísticas do Estoque")
    
    if not st.session_state.dados_estoque.empty:
        # Estatísticas básicas
        total_itens = st.session_state.dados_estoque['Quantidade'].sum()
        total_registros = len(st.session_state.dados_estoque)
        meses_unicos = st.session_state.dados_estoque['Mês'].nunique()
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total de Itens", f"{total_itens:,}")
        
        with col2:
            st.metric("Total de Registros", total_registros)
        
        with col3:
            st.metric("Meses com Registros", meses_unicos)
        
        # Gráfico de itens por setor
        itens_por_setor = st.session_state.dados_estoque.groupby('Setor')['Quantidade'].sum().reset_index()
        fig = px.bar(
            itens_por_setor.sort_values('Quantidade', ascending=False),
            x='Setor',
            y='Quantidade',
            title="Itens por Setor",
            color='Quantidade',
            color_continuous_scale='Blues'
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        
        # Gráfico de itens por mês
        itens_por_mes = st.session_state.dados_estoque.groupby('Mês')['Quantidade'].sum().reset_index()
        # Ordenar meses cronologicamente
        ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                      "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        itens_por_mes['Mês'] = pd.Categorical(itens_por_mes['Mês'], categories=ordem_meses, ordered=True)
        itens_por_mes = itens_por_mes.sort_values('Mês')
        
        fig2 = px.line(
            itens_por_mes, 
            x='Mês', 
            y='Quantidade', 
            title='Itens por Mês',
            markers=True
        )
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Nenhum dado disponível para exibir estatísticas.")

# Seção de visualização de dados
st.markdown("---")
st.subheader("📋 Dados da Planilha")

if not st.session_state.dados_estoque.empty:
    # Filtros
    col1, col2, col3 = st.columns(3)
    
    with col1:
        filtro_setor = st.selectbox(
            "Filtrar por setor:",
            ["Todos"] + list(st.session_state.dados_estoque['Setor'].unique())
        )
    
    with col2:
        filtro_mes = st.selectbox(
            "Filtrar por mês:",
            ["Todos"] + list(st.session_state.dados_estoque['Mês'].unique())
        )
    
    with col3:
        filtro_material = st.selectbox(
            "Filtrar por material:",
            ["Todos"] + list(st.session_state.dados_estoque['Material'].unique())
        )
    
    # Aplicar filtros
    df_filtrado = st.session_state.dados_estoque.copy()
    
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
    
    # Interface para exclusão de registros
    st.subheader("🗑️ Excluir Registros")
    
    with st.form("form_excluir"):
        indices_para_excluir = st.multiselect(
            "Selecione os índices dos registros a excluir:",
            options=df_filtrado.index.tolist(),
            format_func=lambda x: f"Índice {x}: {df_filtrado.loc[x, 'Nome do colaborador']} - {df_filtrado.loc[x, 'Material']}"
        )
        
        confirmar_exclusao = st.form_submit_button("Confirmar Exclusão")
        
        if confirmar_exclusao and indices_para_excluir:
            # Remove os registros selecionados
            novo_df = st.session_state.dados_estoque.drop(indices_para_excluir).reset_index(drop=True)
            st.session_state.dados_estoque = novo_df
            salvar_planilha(st.session_state.dados_estoque, arquivo_excel)
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
    3. Clique em 'Registrar Movimento'
    4. Os dados serão salvos na planilha Excel
    5. Use os filtros para visualizar dados específicos
    """)
    
    st.markdown("---")
    
    if not st.session_state.dados_estoque.empty:
        st.markdown("**📈 Estatísticas Gerais:**")
        ultimo_registro = st.session_state.dados_estoque.iloc[-1]['Nome do colaborador'] if 'Nome do colaborador' in st.session_state.dados_estoque.columns else "N/A"
        ultimo_mes = st.session_state.dados_estoque.iloc[-1]['Mês'] if 'Mês' in st.session_state.dados_estoque.columns else "N/A"
        
        st.write(f"Último registro: **{ultimo_registro}**")
        st.write(f"Último mês: **{ultimo_mes}**")
        
        # Materiais mais retirados
        st.markdown("---")
        st.markdown("**📦 Materiais Mais Retirados:**")
        top_materiais = st.session_state.dados_estoque.groupby('Material')['Quantidade'].sum().nlargest(5)
        for material, qtd in top_materiais.items():
            st.write(f"- {material}: {qtd} unidades")










