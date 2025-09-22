import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# -----------------------------
# Configuração inicial
# -----------------------------
st.set_page_config(page_title="Painel SGE – Notas e Alertas", layout="wide")

MEDIA_APROVACAO = 6.0
MEDIA_FINAL_ALVO = 6.0   # média final desejada após 4 bimestres
SOMA_FINAL_ALVO = MEDIA_FINAL_ALVO * 4  # 24 pontos no ano

# -----------------------------
# Utilidades
# -----------------------------
def detectar_tipo_planilha(df):
    """
    Detecta automaticamente o tipo de planilha baseado nas colunas disponíveis
    Retorna: 'notas_frequencia', 'conteudo_aplicado' ou 'censo_escolar'
    """
    colunas = [col.lower().strip() for col in df.columns]

    # Verificar se é planilha de censo escolar
    censo_indicators = [
        'código', 'superv', 'convên', 'entidade', 'inep', 'situação', 'classific',
        'nome', 'endereço', 'bairro', 'distrito', 'cep', 'cnpj', 'telefone', 'email',
        'nível de', 'categoria', 'tipo de estrutura', 'etapas', 'ano letivo', 'calendário',
        'curso', 'avaliação', 'conceito', 'servidor', 'turno', 'horário', 'tempo',
        'média', 'salário', 'língua', 'professor', 'área de cargo', 'data na', 'cpf'
    ]

    # Verificar se é planilha de conteúdo aplicado
    conteudo_indicators = [
        'componente curricu', 'atividade/conteúdo', 'situação', 'data', 'horário'
    ]

    # Verificar se é planilha de notas/frequência
    notas_indicators = [
        'aluno', 'nota', 'frequencia', 'turma', 'escola', 'disciplina', 'periodo'
    ]

    censo_score = sum(1 for indicator in censo_indicators
                      if any(indicator in col for col in colunas))
    conteudo_score = sum(1 for indicator in conteudo_indicators
                         if any(indicator in col for col in colunas))
    notas_score = sum(1 for indicator in notas_indicators
                      if any(indicator in col for col in colunas))

    # Se tem mais indicadores de censo escolar, é esse tipo
    if censo_score >= 8:
        return 'censo_escolar'
    elif conteudo_score >= 3:
        return 'conteudo_aplicado'
    elif notas_score >= 3:
        return 'notas_frequencia'
    else:
        # Se não conseguir detectar claramente, assume notas/frequência como padrão
        return 'notas_frequencia'

@st.cache_data(show_spinner=False)
def carregar_dados(arquivo, sheet=None):
    if arquivo is None:
        # Tenta ler o padrão local "dados.xlsx"
        df = pd.read_excel("dados.xlsx", sheet_name=sheet) if sheet else pd.read_excel("dados.xlsx")
    else:
        df = pd.read_excel(arquivo, sheet_name=sheet) if sheet else pd.read_excel(arquivo)

    # Normalizar nomes de colunas
    df.columns = [c.strip() for c in df.columns]
    
    # Detectar tipo de planilha
    tipo_planilha = detectar_tipo_planilha(df)
    
    if tipo_planilha == 'conteudo_aplicado':
        # Processar planilha de conteúdo aplicado
        return processar_conteudo_aplicado(df)
    elif tipo_planilha == 'censo_escolar':
        # Processar planilha do censo escolar
        return processar_censo_escolar(df)
    else:
        # Processar planilha de notas/frequência (padrão atual)
        return processar_notas_frequencia(df)

def processar_conteudo_aplicado(df):
    """Processa planilha de conteúdo aplicado"""
    # Mapear colunas para nomes padronizados
    mapeamento_colunas = {}
    
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'componente curricu' in col_lower:
            mapeamento_colunas[col] = 'Disciplina'
        elif 'atividade/conteúdo' in col_lower or 'atividade' in col_lower:
            mapeamento_colunas[col] = 'Atividade'
        elif 'situação' in col_lower:
            mapeamento_colunas[col] = 'Status'
        elif 'data' in col_lower:
            mapeamento_colunas[col] = 'Data'
        elif 'horário' in col_lower:
            mapeamento_colunas[col] = 'Horario'
    
    df = df.rename(columns=mapeamento_colunas)
    
    # Converter Data para datetime se possível
    if 'Data' in df.columns:
        # Tentar diferentes formatos de data
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
        # Se não funcionar, tentar formato automático
        if df['Data'].isna().all():
            df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    
    # Padronizar texto dos campos principais
    for col in ['Disciplina', 'Atividade', 'Status']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    
    # Adicionar tipo de planilha para identificação
    df.attrs['tipo_planilha'] = 'conteudo_aplicado'
    
    return df

def processar_notas_frequencia(df):
    """Processa planilha de notas/frequência (processamento atual)"""
    # Garantir colunas esperadas (flexível aos nomes encontrados)
    # Esperados: Escola, Turma, Turno, Aluno, Periodo, Disciplina, Nota, Falta, Frequência, Frequência Anual
    # Algumas planilhas têm "Período" com acento; vamos padronizar para "Periodo"
    if "Período" in df.columns and "Periodo" not in df.columns:
        df = df.rename(columns={"Período": "Periodo"})
    if "Frequência" in df.columns and "Frequencia" not in df.columns:
        df = df.rename(columns={"Frequência": "Frequencia"})
    if "Frequência Anual" in df.columns and "Frequencia Anual" not in df.columns:
        df = df.rename(columns={"Frequência Anual": "Frequencia Anual"})

    # Converter Nota (vírgula -> ponto, texto -> float)
    if "Nota" in df.columns:
        df["Nota"] = (
            df["Nota"]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(" ", "", regex=False)
        )
        df["Nota"] = pd.to_numeric(df["Nota"], errors="coerce")

    # Falta -> numérico
    if "Falta" in df.columns:
        df["Falta"] = pd.to_numeric(df["Falta"], errors="coerce").fillna(0).astype(int)

    # Frequências -> numérico
    if "Frequencia" in df.columns:
        df["Frequencia"] = pd.to_numeric(df["Frequencia"], errors="coerce")
    if "Frequencia Anual" in df.columns:
        df["Frequencia Anual"] = pd.to_numeric(df["Frequencia Anual"], errors="coerce")

    # Mapear colunas específicas se necessário
    if "Estudante" in df.columns and "Aluno" not in df.columns:
        df = df.rename(columns={"Estudante": "Aluno"})
    
    # Padronizar texto dos campos principais (evita diferenças por espaços)
    for col in ["Escola", "Turma", "Turno", "Aluno", "Status", "Periodo", "Disciplina"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    
    # Adicionar tipo de planilha para identificação
    df.attrs['tipo_planilha'] = 'notas_frequencia'
    
    return df

def processar_censo_escolar(df):
    """
    Processa dados do Censo Escolar - Lista de Estudantes
    """
    # Normalizar nomes das colunas
    df.columns = df.columns.str.strip()
    
    # Mapear colunas específicas da planilha ListaDeEstudantes_TurmaEscolarização
    colunas_mapeadas = {}
    for col in df.columns:
        col_lower = col.lower()
        if col == 'Nome':
            colunas_mapeadas[col] = 'Nome_Estudante'
        elif col == 'Escola':
            colunas_mapeadas[col] = 'Escola'
        elif col == 'CPF':
            colunas_mapeadas[col] = 'CPF'
        elif col == 'INEP':
            colunas_mapeadas[col] = 'Codigo_Estudante'
        elif col == 'Situação da Matrícula':
            colunas_mapeadas[col] = 'Situacao'
        elif col == 'Turno':
            colunas_mapeadas[col] = 'Turno'
        elif col == 'Data Nascimento':
            colunas_mapeadas[col] = 'Data_Nascimento'
        elif col == 'Nível de Ensino':
            colunas_mapeadas[col] = 'Nivel_Educacao'
        elif col == 'Ano/Série':
            colunas_mapeadas[col] = 'Ano_Serie'
        elif col == 'Descrição Turma':
            colunas_mapeadas[col] = 'Turma'
        elif col == 'Entidade Conveniada':
            colunas_mapeadas[col] = 'Entidade'
        elif col == 'Superintendência Regional':
            colunas_mapeadas[col] = 'Supervisao'
        elif col == 'Convênio':
            colunas_mapeadas[col] = 'Convenio'
        elif col == 'INEP da Escola':
            colunas_mapeadas[col] = 'INEP_Escola'
        elif col == 'Classificação da Escola':
            colunas_mapeadas[col] = 'Classificacao'
        elif col == 'Endereço':
            colunas_mapeadas[col] = 'Endereco'
        elif col == 'Bairro':
            colunas_mapeadas[col] = 'Bairro'
        elif col == 'Distrito':
            colunas_mapeadas[col] = 'Distrito'
        elif col == 'Cep':
            colunas_mapeadas[col] = 'CEP'
        elif col == 'Telefone Principal':
            colunas_mapeadas[col] = 'Telefone'
        elif col == 'E-mail':
            colunas_mapeadas[col] = 'Email'
        elif col == 'CNPJ':
            colunas_mapeadas[col] = 'CNPJ'
        elif col == 'Carga Horária':
            colunas_mapeadas[col] = 'Carga_Horaria'
        elif col == 'Entrada':
            colunas_mapeadas[col] = 'Data_Entrada'
        elif col == 'Data de saída':
            colunas_mapeadas[col] = 'Data_Saida'
        elif col == 'Cor/Raça':
            colunas_mapeadas[col] = 'Cor_Raca'
    
    # Renomear colunas
    df = df.rename(columns=colunas_mapeadas)
    
    # Converter tipos de dados
    if 'Data_Nascimento' in df.columns:
        df['Data_Nascimento'] = pd.to_datetime(df['Data_Nascimento'], dayfirst=True, errors='coerce')
    
    if 'Data_Entrada' in df.columns:
        df['Data_Entrada'] = pd.to_datetime(df['Data_Entrada'], dayfirst=True, errors='coerce')
    
    if 'Data_Saida' in df.columns:
        df['Data_Saida'] = pd.to_datetime(df['Data_Saida'], dayfirst=True, errors='coerce')
    
    # Padronizar texto dos campos principais
    for col in ['Nome_Estudante', 'Escola', 'Situacao', 'Turno', 'Nivel_Educacao', 'Ano_Serie', 'Turma']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    
    # Marcar tipo de planilha
    df.attrs['tipo_planilha'] = 'censo_escolar'
    
    return df

def criar_interface_censo_escolar(df):
    """Cria interface específica para análise do Censo Escolar"""
    
    # Header específico para censo escolar
    st.markdown("""
    <div style="background: linear-gradient(90deg, #1e40af 0%, #3b82f6 100%); 
                padding: 2rem; border-radius: 10px; margin-bottom: 2rem; text-align: center;">
        <h1 style="color: white; margin: 0; font-size: 2.5rem; font-weight: bold;">
            📊 Painel Censo Escolar
        </h1>
        <p style="color: #e0e7ff; margin: 0.5rem 0 0 0; font-size: 1.2rem;">
            Identificação de Duplicatas - Estudantes em Múltiplas Escolas/Turmas
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Resumo Geral Simples
    st.markdown("### 📊 Resumo Geral")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total de Registros", f"{len(df):,}")
    
    with col2:
        escolas_unicas = df['Escola'].nunique() if 'Escola' in df.columns else 0
        st.metric("Escolas", escolas_unicas)
    
    with col3:
        estudantes_unicos = df['Nome_Estudante'].nunique() if 'Nome_Estudante' in df.columns else 0
        st.metric("Estudantes Únicos", estudantes_unicos)
    
    with col4:
        turmas_unicas = df['Turma'].nunique() if 'Turma' in df.columns else 0
        st.metric("Turmas", turmas_unicas)
    
    # Filtros Simples
    st.sidebar.markdown("### 🔍 Filtros")
    
    # Filtro por Escola
    if 'Escola' in df.columns:
        escolas_disponiveis = ['Todas as Escolas'] + sorted(df['Escola'].dropna().unique().tolist())
        escola_sel = st.sidebar.selectbox("Escola", escolas_disponiveis)
        
        if escola_sel != 'Todas as Escolas':
            df_filt = df[df['Escola'] == escola_sel].copy()
        else:
            df_filt = df.copy()
    else:
        df_filt = df.copy()
        escola_sel = 'Todas as Escolas'
    
    # Filtro por Situação (apenas Matriculado)
    if 'Situacao' in df.columns:
        situacoes_disponiveis = ['Todas as Situações'] + sorted(df_filt['Situacao'].dropna().unique().tolist())
        situacao_sel = st.sidebar.selectbox("Situação", situacoes_disponiveis)
        
        if situacao_sel != 'Todas as Situações':
            df_filt = df_filt[df_filt['Situacao'] == situacao_sel].copy()
    else:
        situacao_sel = 'Todas as Situações'
    
    # Resumo dos Dados Filtrados
    st.markdown("### 📋 Dados Após Filtros")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Registros", f"{len(df_filt):,}")
    
    with col2:
        estudantes_filtrados = df_filt['Nome_Estudante'].nunique() if 'Nome_Estudante' in df_filt.columns else 0
        st.metric("Estudantes", estudantes_filtrados)
    
    with col3:
        escolas_filtradas = df_filt['Escola'].nunique() if 'Escola' in df_filt.columns else 0
        st.metric("Escolas", escolas_filtradas)
    
    # Análise de Duplicatas - Foco Principal
    st.markdown("### 🔍 Duplicatas Encontradas")
    
    if 'Nome_Estudante' in df_filt.columns and 'Escola' in df_filt.columns:
        # 1. Duplicatas por Escola (estudante em múltiplas escolas)
        duplicatas_escola = df_filt.groupby('Nome_Estudante').agg({
            'Escola': 'nunique',
            'Turma': 'nunique' if 'Turma' in df_filt.columns else 'count'
        }).reset_index()
        
        estudantes_multiplas_escolas = duplicatas_escola[duplicatas_escola['Escola'] > 1]
        
        # 2. Duplicatas por Turma (estudante em múltiplas turmas na mesma escola)
        duplicatas_turma = df_filt.groupby(['Nome_Estudante', 'Escola']).agg({
            'Turma': 'nunique' if 'Turma' in df_filt.columns else 'count'
        }).reset_index()
        
        estudantes_multiplas_turmas = duplicatas_turma[duplicatas_turma['Turma'] > 1]
        
        # Métricas Principais
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Em Múltiplas Escolas", len(estudantes_multiplas_escolas))
        
        with col2:
            st.metric("Em Múltiplas Turmas", len(estudantes_multiplas_turmas))
        
        with col3:
            total_duplicatas = len(estudantes_multiplas_escolas) + len(estudantes_multiplas_turmas)
            st.metric("Total Duplicatas", total_duplicatas)
        
        with col4:
            percentual = (total_duplicatas / len(df_filt['Nome_Estudante'].unique())) * 100 if len(df_filt['Nome_Estudante'].unique()) > 0 else 0
            st.metric("Percentual", f"{percentual:.1f}%")
        
        # Tabelas Detalhadas
        if len(estudantes_multiplas_escolas) > 0 or len(estudantes_multiplas_turmas) > 0:
            
            # 1. Estudantes em Múltiplas Escolas (Detalhado)
            if len(estudantes_multiplas_escolas) > 0:
                st.markdown("#### 🏫 Estudantes em Múltiplas Escolas")
                
                # Criar tabela detalhada mostrando escola + turma para cada estudante
                duplicatas_escola_detalhadas = []
                for _, row in estudantes_multiplas_escolas.iterrows():
                    nome = row['Nome_Estudante']
                    dados_estudante = df_filt[df_filt['Nome_Estudante'] == nome]
                    
                    # Para cada escola do estudante, mostrar a turma correspondente
                    for _, linha in dados_estudante.iterrows():
                        duplicatas_escola_detalhadas.append({
                            'Nome': nome,
                            'Escola': linha['Escola'],
                            'Turma': linha['Turma'] if 'Turma' in linha else 'N/A',
                            'CPF': linha['CPF'] if 'CPF' in linha else 'N/A',
                            'Situacao': linha['Situacao'] if 'Situacao' in linha else 'N/A'
                        })
                
                df_duplicatas_escola = pd.DataFrame(duplicatas_escola_detalhadas)
                st.dataframe(df_duplicatas_escola, use_container_width=True)
            
            # 2. Estudantes em Múltiplas Turmas (mesma escola) - Detalhado
            if len(estudantes_multiplas_turmas) > 0:
                st.markdown("#### 🎓 Estudantes em Múltiplas Turmas (Mesma Escola)")
                
                # Criar tabela detalhada mostrando cada linha de turma
                duplicatas_turma_detalhadas = []
                for _, row in estudantes_multiplas_turmas.iterrows():
                    nome = row['Nome_Estudante']
                    escola = row['Escola']
                    dados_estudante = df_filt[(df_filt['Nome_Estudante'] == nome) & 
                                            (df_filt['Escola'] == escola)]
                    
                    # Para cada turma do estudante na mesma escola
                    for _, linha in dados_estudante.iterrows():
                        duplicatas_turma_detalhadas.append({
                            'Nome': nome,
                            'Escola': escola,
                            'Turma': linha['Turma'] if 'Turma' in linha else 'N/A',
                            'CPF': linha['CPF'] if 'CPF' in linha else 'N/A',
                            'Situacao': linha['Situacao'] if 'Situacao' in linha else 'N/A'
                        })
                
                df_duplicatas_turma = pd.DataFrame(duplicatas_turma_detalhadas)
                st.dataframe(df_duplicatas_turma, use_container_width=True)
            else:
                st.info("ℹ️ Nenhum estudante encontrado em múltiplas turmas da mesma escola.")
            
            # Botão de Download com Abas Separadas
            st.markdown("#### 💾 Download dos Dados")
            
            # Preparar dados para download em abas separadas
            if len(estudantes_multiplas_escolas) > 0 or len(estudantes_multiplas_turmas) > 0:
                
                # Converter para Excel com abas separadas
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    # Aba 1: Duplicatas por Escola (Detalhado)
                    if len(estudantes_multiplas_escolas) > 0:
                        duplicatas_escola_download = []
                        for _, row in estudantes_multiplas_escolas.iterrows():
                            nome = row['Nome_Estudante']
                            dados_estudante = df_filt[df_filt['Nome_Estudante'] == nome]
                            
                            # Para cada escola do estudante, mostrar a turma correspondente
                            for _, linha in dados_estudante.iterrows():
                                duplicatas_escola_download.append({
                                    'Nome': nome,
                                    'Escola': linha['Escola'],
                                    'Turma': linha['Turma'] if 'Turma' in linha else 'N/A',
                                    'CPF': linha['CPF'] if 'CPF' in linha else 'N/A',
                                    'Situacao': linha['Situacao'] if 'Situacao' in linha else 'N/A'
                                })
                        
                        df_escola_download = pd.DataFrame(duplicatas_escola_download)
                        df_escola_download.to_excel(writer, sheet_name='Múltiplas_Escolas', index=False)
                    
                    # Aba 2: Duplicatas por Turma (Detalhado)
                    if len(estudantes_multiplas_turmas) > 0:
                        duplicatas_turma_download = []
                        for _, row in estudantes_multiplas_turmas.iterrows():
                            nome = row['Nome_Estudante']
                            escola = row['Escola']
                            dados_estudante = df_filt[(df_filt['Nome_Estudante'] == nome) & 
                                                    (df_filt['Escola'] == escola)]
                            
                            # Para cada turma do estudante na mesma escola
                            for _, linha in dados_estudante.iterrows():
                                duplicatas_turma_download.append({
                                    'Nome': nome,
                                    'Escola': escola,
                                    'Turma': linha['Turma'] if 'Turma' in linha else 'N/A',
                                    'CPF': linha['CPF'] if 'CPF' in linha else 'N/A',
                                    'Situacao': linha['Situacao'] if 'Situacao' in linha else 'N/A'
                                })
                        
                        df_turma_download = pd.DataFrame(duplicatas_turma_download)
                        df_turma_download.to_excel(writer, sheet_name='Múltiplas_Turmas', index=False)
                    
                    # Aba 3: Resumo Geral
                    resumo_geral = pd.DataFrame({
                        'Tipo_Duplicata': ['Múltiplas Escolas', 'Múltiplas Turmas', 'Total'],
                        'Quantidade': [
                            len(estudantes_multiplas_escolas),
                            len(estudantes_multiplas_turmas),
                            len(estudantes_multiplas_escolas) + len(estudantes_multiplas_turmas)
                        ],
                        'Percentual': [
                            f"{(len(estudantes_multiplas_escolas) / len(df_filt['Nome_Estudante'].unique())) * 100:.1f}%" if len(df_filt['Nome_Estudante'].unique()) > 0 else "0%",
                            f"{(len(estudantes_multiplas_turmas) / len(df_filt['Nome_Estudante'].unique())) * 100:.1f}%" if len(df_filt['Nome_Estudante'].unique()) > 0 else "0%",
                            f"{((len(estudantes_multiplas_escolas) + len(estudantes_multiplas_turmas)) / len(df_filt['Nome_Estudante'].unique())) * 100:.1f}%" if len(df_filt['Nome_Estudante'].unique()) > 0 else "0%"
                        ]
                    })
                    resumo_geral.to_excel(writer, sheet_name='Resumo', index=False)
                
                st.download_button(
                    label="📥 Baixar Relatório Completo (Excel com Abas)",
                    data=output.getvalue(),
                    file_name=f"duplicatas_censo_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.success("✅ Nenhuma duplicata encontrada nos dados filtrados!")
    
    
    # Dados Brutos (Opcional)
    with st.expander("📄 Ver todos os dados", expanded=False):
        st.dataframe(df_filt, use_container_width=True)

def criar_interface_conteudo_aplicado(df):
    """Cria interface específica para análise de conteúdo aplicado"""
    
    # Header específico para conteúdo aplicado
    st.markdown("""
    <div style="text-align: center; padding: 40px 20px; background: linear-gradient(135deg, #059669, #10b981); border-radius: 15px; margin-bottom: 30px; box-shadow: 0 8px 25px rgba(5, 150, 105, 0.3);">
        <h1 style="color: white; margin: 0; font-size: 2.2em; font-weight: 700; text-shadow: 0 2px 4px rgba(0,0,0,0.3);">Superintendência Regional de Educação de Gurupi TO</h1>
        <h2 style="color: white; margin: 15px 0 0 0; font-weight: 600; font-size: 1.8em; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Painel SGE - Conteúdo Aplicado</h2>
        <h3 style="color: rgba(255,255,255,0.95); margin: 10px 0 0 0; font-weight: 500; font-size: 1.4em;">Análise de Atividades e Conteúdos Registrados</h3>
        <p style="color: rgba(255,255,255,0.8); margin: 10px 0 0 0; font-size: 1.1em; font-weight: 400;">Registros de Conteúdo Aplicado</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Métricas gerais
    st.markdown("""
    <div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
        <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Visão Geral dos Registros</h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            label="Total de Registros", 
            value=f"{len(df):,}".replace(",", "."),
            help="Total de atividades/conteúdos registrados"
        )
    
    with col2:
        disciplinas_unicas = df["Disciplina"].nunique() if "Disciplina" in df.columns else 0
        st.metric(
            label="Disciplinas", 
            value=disciplinas_unicas,
            help="Número de disciplinas diferentes"
        )
    
    with col3:
        status_unicos = df["Status"].nunique() if "Status" in df.columns else 0
        st.metric(
            label="Status Diferentes", 
            value=status_unicos,
            help="Número de status diferentes"
        )
    
    with col4:
        if "Data" in df.columns:
            periodo_cobertura = f"{df['Data'].min().strftime('%d/%m/%Y')} a {df['Data'].max().strftime('%d/%m/%Y')}"
            st.metric(
                label="Período", 
                value=periodo_cobertura,
                help="Período coberto pelos registros"
            )
        else:
            st.metric("Período", "N/A")
    
    with col5:
        # Mostrar disciplina com mais registros
        if "Disciplina" in df.columns:
            disciplina_top = df["Disciplina"].value_counts().index[0] if len(df) > 0 else "N/A"
            qtd_top = df["Disciplina"].value_counts().iloc[0] if len(df) > 0 else 0
            st.metric(
                label="Disciplina Top", 
                value=f"{disciplina_top}",
                delta=f"{qtd_top} registros",
                help="Disciplina com maior número de registros"
            )
        else:
            st.metric("Disciplina Top", "N/A")
    
    # Função para classificar por bimestre baseado nas datas
    def classificar_bimestre(data):
        """Classifica a data em bimestre baseado nos períodos definidos"""
        if pd.isna(data):
            return "Sem Data"
        
        # Converter para datetime se necessário
        if not isinstance(data, pd.Timestamp):
            data = pd.to_datetime(data, errors='coerce')
        
        if pd.isna(data):
            return "Sem Data"
        
        # Definir períodos dos bimestres (ano 2025)
        bimestre1_inicio = pd.to_datetime("2025-02-03")
        bimestre1_fim = pd.to_datetime("2025-04-03")
        
        bimestre2_inicio = pd.to_datetime("2025-04-04")
        bimestre2_fim = pd.to_datetime("2025-06-27")
        
        bimestre3_inicio = pd.to_datetime("2025-08-04")
        bimestre3_fim = pd.to_datetime("2025-10-11")
        
        bimestre4_inicio = pd.to_datetime("2025-10-12")
        bimestre4_fim = pd.to_datetime("2025-12-19")
        
        # Classificar por bimestre
        if bimestre1_inicio <= data <= bimestre1_fim:
            return "1º Bimestre"
        elif bimestre2_inicio <= data <= bimestre2_fim:
            return "2º Bimestre"
        elif bimestre3_inicio <= data <= bimestre3_fim:
            return "3º Bimestre"
        elif bimestre4_inicio <= data <= bimestre4_fim:
            return "4º Bimestre"
        else:
            return "Fora do Período Letivo"
    
    # Adicionar coluna de bimestre se houver dados de data
    if "Data" in df.columns:
        df["Bimestre"] = df["Data"].apply(classificar_bimestre)
        
        
        
        # Análise por Bimestres
        st.markdown("""
        <div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
            <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Análise por Bimestres</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Contagem por bimestre
        contagem_bimestres = df["Bimestre"].value_counts().reset_index()
        contagem_bimestres.columns = ["Bimestre", "Quantidade"]
        
        # Ordenar por bimestre (1º, 2º, 3º, 4º)
        ordem_bimestres = ["1º Bimestre", "2º Bimestre", "3º Bimestre", "4º Bimestre", "Fora do Período Letivo", "Sem Data"]
        contagem_bimestres["Ordem"] = contagem_bimestres["Bimestre"].map({b: i for i, b in enumerate(ordem_bimestres)})
        contagem_bimestres = contagem_bimestres.sort_values("Ordem").reset_index(drop=True)
        
        # Criar colunas para mostrar bimestres
        num_bimestres = len(contagem_bimestres)
        num_colunas_bim = min(num_bimestres, 6)
        cols_bimestres = st.columns(num_colunas_bim)
        
        # Mostrar bimestres em cards
        for i, (_, row) in enumerate(contagem_bimestres.iterrows()):
            col_index = i % num_colunas_bim
            with cols_bimestres[col_index]:
                # Definir cor baseada no bimestre
                if "1º" in row['Bimestre']:
                    cor_borda = "#3b82f6"  # Azul
                elif "2º" in row['Bimestre']:
                    cor_borda = "#10b981"  # Verde
                elif "3º" in row['Bimestre']:
                    cor_borda = "#f59e0b"  # Amarelo
                elif "4º" in row['Bimestre']:
                    cor_borda = "#ef4444"  # Vermelho
                else:
                    cor_borda = "#6b7280"  # Cinza
                
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #f0f9ff, #e0f2fe); border-radius: 8px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid {cor_borda};">
                    <div style="font-size: 0.9em; font-weight: 600; color: #1e40af; margin-bottom: 8px;">{row['Bimestre']}</div>
                    <div style="font-size: 1.8em; font-weight: 700; color: #1e40af; margin: 8px 0;">{row['Quantidade']}</div>
                    <div style="font-size: 1.1em; color: #64748b; font-weight: 600;">registros</div>
                </div>
                """, unsafe_allow_html=True)
        
        # Gráfico de barras por bimestre
        fig_bimestres = px.bar(contagem_bimestres, x="Bimestre", y="Quantidade", 
                              title="Registros por Bimestre",
                              color="Bimestre",
                              color_discrete_map={
                                  "1º Bimestre": "#3b82f6",
                                  "2º Bimestre": "#10b981", 
                                  "3º Bimestre": "#f59e0b",
                                  "4º Bimestre": "#ef4444",
                                  "Fora do Período Letivo": "#6b7280",
                                  "Sem Data": "#9ca3af"
                              })
        fig_bimestres.update_layout(xaxis_tickangle=45, showlegend=False)
        st.plotly_chart(fig_bimestres, use_container_width=True)
        
        # Análise detalhada por bimestre - disciplinas em cada bimestre
        st.markdown("""
        <div style="background: linear-gradient(135deg, #7c3aed, #a855f7); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(124, 58, 237, 0.2);">
            <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Disciplinas por Bimestre</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Criar análise por bimestre e disciplina
        bimestre_disciplina = df.groupby(['Bimestre', 'Disciplina']).size().reset_index(name='Quantidade')
        
        # Ordenar por bimestre e quantidade
        ordem_bimestres = ["1º Bimestre", "2º Bimestre", "3º Bimestre", "4º Bimestre", "Fora do Período Letivo", "Sem Data"]
        bimestre_disciplina['Ordem_Bimestre'] = bimestre_disciplina['Bimestre'].map({b: i for i, b in enumerate(ordem_bimestres)})
        bimestre_disciplina = bimestre_disciplina.sort_values(['Ordem_Bimestre', 'Quantidade'], ascending=[True, False])
        
        # Mostrar cada bimestre com suas disciplinas
        for bimestre in ordem_bimestres:
            if bimestre in bimestre_disciplina['Bimestre'].values:
                disciplinas_bimestre = bimestre_disciplina[bimestre_disciplina['Bimestre'] == bimestre]
                
                # Definir cor do bimestre
                if "1º" in bimestre:
                    cor_bimestre = "#3b82f6"  # Azul
                elif "2º" in bimestre:
                    cor_bimestre = "#10b981"  # Verde
                elif "3º" in bimestre:
                    cor_bimestre = "#f59e0b"  # Amarelo
                elif "4º" in bimestre:
                    cor_bimestre = "#ef4444"  # Vermelho
                else:
                    cor_bimestre = "#6b7280"  # Cinza
                
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #f8fafc, #f1f5f9); border-radius: 8px; padding: 20px; margin: 15px 0; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border-left: 4px solid {cor_bimestre};">
                    <h4 style="color: {cor_bimestre}; margin: 0 0 15px 0; font-size: 1.3em; font-weight: 700;">{bimestre}</h4>
                </div>
                """, unsafe_allow_html=True)
                
                # Criar colunas para as disciplinas deste bimestre
                num_disciplinas = len(disciplinas_bimestre)
                num_colunas_disc = min(num_disciplinas, 4)  # Máximo 4 colunas
                cols_disciplinas = st.columns(num_colunas_disc)
                
                # Mostrar disciplinas em cards
                for i, (_, row) in enumerate(disciplinas_bimestre.iterrows()):
                    col_index = i % num_colunas_disc
                    with cols_disciplinas[col_index]:
                        st.markdown(f"""
                        <div style="background: linear-gradient(135deg, #ffffff, #f8fafc); border-radius: 6px; padding: 12px; margin: 5px 0; box-shadow: 0 1px 4px rgba(0,0,0,0.1); border-left: 3px solid {cor_bimestre};">
                            <div style="font-size: 0.9em; font-weight: 600; color: #374151; margin-bottom: 6px;">{row['Disciplina']}</div>
                            <div style="font-size: 1.5em; font-weight: 700; color: {cor_bimestre}; margin: 6px 0;">{row['Quantidade']}</div>
                            <div style="font-size: 0.9em; color: #6b7280; font-weight: 500;">registros</div>
                        </div>
                        """, unsafe_allow_html=True)
                
                # Gráfico de barras para este bimestre
                fig_bimestre_disc = px.bar(disciplinas_bimestre, x="Disciplina", y="Quantidade", 
                                          title=f"Disciplinas - {bimestre}",
                                          color="Disciplina",
                                          color_discrete_sequence=px.colors.qualitative.Set3)
                fig_bimestre_disc.update_layout(xaxis_tickangle=45, showlegend=False, height=300)
                st.plotly_chart(fig_bimestre_disc, use_container_width=True)
    
    # Adicionar seção com disciplinas (todas ou filtradas) - será movida para depois dos filtros
    
    # Filtros específicos para conteúdo aplicado
    st.sidebar.markdown("""
    <div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
        <h2 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Filtros - Conteúdo</h2>
        <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1em; font-weight: 500;">Filtre os registros de conteúdo</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Filtros
    disciplinas_opcoes = sorted(df["Disciplina"].dropna().unique().tolist()) if "Disciplina" in df.columns else []
    status_opcoes = sorted(df["Status"].dropna().unique().tolist()) if "Status" in df.columns else []
    bimestres_opcoes = sorted(df["Bimestre"].dropna().unique().tolist()) if "Bimestre" in df.columns else []
    
    # Filtro de Data
    if "Data" in df.columns:
        st.sidebar.markdown("""
        <div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
            <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">📅 Período</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Obter datas mínima e máxima
        data_min = df["Data"].min()
        data_max = df["Data"].max()
        
        # Filtro de data com slider
        data_range = st.sidebar.date_input(
            "Selecione o período:",
            value=(data_min.date(), data_max.date()),
            min_value=data_min.date(),
            max_value=data_max.date(),
            help="Selecione o período para filtrar os registros"
        )
        
        # Converter para datetime se necessário
        if len(data_range) == 2:
            data_inicio = pd.to_datetime(data_range[0])
            data_fim = pd.to_datetime(data_range[1])
        else:
            data_inicio = data_min
            data_fim = data_max
    
    # Filtro de Disciplina
    st.sidebar.markdown("""
    <div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
        <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">📚 Disciplina</h3>
    </div>
    """, unsafe_allow_html=True)
    
    disciplina_sel = st.sidebar.multiselect(
        "Selecione as disciplinas:", 
        disciplinas_opcoes, 
        help="Filtre por disciplinas específicas"
    )
    
    # Filtro de Status
    st.sidebar.markdown("""
    <div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
        <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">✅ Status</h3>
    </div>
    """, unsafe_allow_html=True)
    
    status_sel = st.sidebar.multiselect(
        "Selecione os status:", 
        status_opcoes, 
        help="Filtre por status específicos"
    )
    
    # Filtro de Bimestre
    if "Bimestre" in df.columns and len(bimestres_opcoes) > 0:
        st.sidebar.markdown("""
        <div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
            <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">📅 Bimestre</h3>
        </div>
        """, unsafe_allow_html=True)
        
        bimestre_sel = st.sidebar.multiselect(
            "Selecione os bimestres:", 
            bimestres_opcoes, 
            help="Filtre por bimestres específicos"
        )
    else:
        bimestre_sel = []
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    # Filtro por data
    if "Data" in df.columns and 'data_inicio' in locals() and 'data_fim' in locals():
        df_filtrado = df_filtrado[
            (df_filtrado["Data"] >= data_inicio) & 
            (df_filtrado["Data"] <= data_fim)
        ]
    
    # Filtro por disciplina
    if disciplina_sel:
        df_filtrado = df_filtrado[df_filtrado["Disciplina"].isin(disciplina_sel)]
    
    # Filtro por status
    if status_sel:
        df_filtrado = df_filtrado[df_filtrado["Status"].isin(status_sel)]
    
    # Filtro por bimestre
    if bimestre_sel:
        df_filtrado = df_filtrado[df_filtrado["Bimestre"].isin(bimestre_sel)]
    
    # Verificar se há filtros aplicados (agora que as variáveis estão definidas)
    tem_filtros = (
        ('data_inicio' in locals() and 'data_fim' in locals() and 
         (data_inicio != df["Data"].min() or data_fim != df["Data"].max())) or
        disciplina_sel or 
        status_sel or
        bimestre_sel
    )
    
    # Determinar título e dados baseado nos filtros
    if tem_filtros:
        titulo_secao = "Disciplinas Filtradas"
        dados_disciplinas = df_filtrado["Disciplina"].value_counts().reset_index() if len(df_filtrado) > 0 else pd.DataFrame()
    else:
        titulo_secao = "Todas as Disciplinas"
        dados_disciplinas = df["Disciplina"].value_counts().reset_index()
    
    dados_disciplinas.columns = ["Disciplina", "Quantidade"]
    
    # Adicionar seção com disciplinas (todas ou filtradas)
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
        <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">{titulo_secao}</h3>
    </div>
    """, unsafe_allow_html=True)
    
    if len(dados_disciplinas) > 0:
        # Calcular número de colunas necessárias (máximo 6 para não ficar muito pequeno)
        num_disciplinas = len(dados_disciplinas)
        num_colunas = min(num_disciplinas, 6)
        
        # Criar colunas dinamicamente
        cols_disciplinas = st.columns(num_colunas)
        
        # Mostrar disciplinas em cards
        for i, (_, row) in enumerate(dados_disciplinas.iterrows()):
            col_index = i % num_colunas
            with cols_disciplinas[col_index]:
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 8px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(5, 150, 105, 0.15); border-left: 4px solid #059669;">
                    <div style="font-size: 0.9em; font-weight: 600; color: #047857; margin-bottom: 8px;">{row['Disciplina']}</div>
                    <div style="font-size: 1.8em; font-weight: 700; color: #047857; margin: 8px 0;">{row['Quantidade']}</div>
                    <div style="font-size: 1.1em; color: #64748b; font-weight: 600;">registros</div>
                </div>
                """, unsafe_allow_html=True)
        
        # Se há mais de 6 disciplinas, mostrar aviso
        if num_disciplinas > 6:
            st.info(f"Mostrando as primeiras 6 disciplinas de {num_disciplinas} total. Use os filtros para focar em disciplinas específicas.")
    
    # Mostrar informações dos filtros aplicados
    st.markdown("""
    <div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
        <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Dados Filtrados</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Métricas dos dados filtrados
    col_filt1, col_filt2, col_filt3 = st.columns(3)
    
    with col_filt1:
        st.metric(
            label="Registros Filtrados", 
            value=f"{len(df_filtrado):,}".replace(",", "."),
            delta=f"{len(df_filtrado) - len(df)}" if len(df_filtrado) != len(df) else "0",
            help="Total de registros após aplicar os filtros"
        )
    
    with col_filt2:
        if len(df_filtrado) > 0 and "Disciplina" in df_filtrado.columns:
            disciplinas_filtradas = df_filtrado["Disciplina"].nunique()
            st.metric(
                label="Disciplinas no Filtro", 
                value=disciplinas_filtradas,
                help="Número de disciplinas nos dados filtrados"
            )
        else:
            st.metric("Disciplinas no Filtro", "0")
    
    with col_filt3:
        if len(df_filtrado) > 0 and "Data" in df_filtrado.columns:
            periodo_filtrado = f"{df_filtrado['Data'].min().strftime('%d/%m/%Y')} a {df_filtrado['Data'].max().strftime('%d/%m/%Y')}"
            st.metric(
                label="Período Filtrado", 
                value=periodo_filtrado,
                help="Período dos dados filtrados"
            )
        else:
            st.metric("Período Filtrado", "N/A")
    
    # Análise por Disciplina
    st.markdown("""
    <div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
        <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Análise por Disciplina</h3>
    </div>
    """, unsafe_allow_html=True)
    
    if len(df_filtrado) > 0:
        # Contagem por disciplina
        contagem_disciplina = df_filtrado["Disciplina"].value_counts().reset_index()
        contagem_disciplina.columns = ["Disciplina", "Quantidade"]
        
        # Gráfico de barras
        fig = px.bar(contagem_disciplina, x="Disciplina", y="Quantidade", 
                    title="Registros por Disciplina",
                    color="Quantidade",
                    color_continuous_scale="Viridis")
        fig.update_layout(xaxis_tickangle=45)
        st.plotly_chart(fig, use_container_width=True)
        
        # Tabela detalhada
        st.markdown("### Registros Detalhados")
        st.dataframe(df_filtrado, use_container_width=True)
        
        # Botão de exportação
        col_export1, col_export2 = st.columns([1, 4])
        with col_export1:
            if st.button("📊 Exportar Dados", key="export_conteudo", help="Baixar planilha com análise de conteúdo aplicado"):
                excel_data = criar_excel_formatado(df_filtrado, "Conteudo_Aplicado")
                st.download_button(
                    label="Baixar Excel",
                    data=excel_data,
                    file_name="conteudo_aplicado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Nenhum registro encontrado com os filtros aplicados.")

def mapear_bimestre(periodo: str) -> int | None:
    """Mapeia 'Primeiro Bimestre' -> 1, 'Segundo Bimestre' -> 2, etc."""
    if not isinstance(periodo, str):
        return None
    p = periodo.lower()
    if "primeiro" in p or "1º" in p or "1o" in p:
        return 1
    if "segundo" in p or "2º" in p or "2o" in p:
        return 2
    if "terceiro" in p or "3º" in p or "3o" in p:
        return 3
    if "quarto" in p or "4º" in p or "4o" in p:
        return 4
    return None

def classificar_status_b1_b2(n1, n2, media12):
    """
    Regras:
      - 'Vermelho Duplo': n1<6 e n2<6
      - 'Queda p/ Vermelho': n1>=6 e n2<6
      - 'Recuperou': n1<6 e n2>=6
      - 'Verde': n1>=6 e n2>=6
      - Se faltar n1 ou n2, retorna 'Incompleto'
    """
    if pd.isna(n1) or pd.isna(n2):
        return "Incompleto"
    if n1 < MEDIA_APROVACAO and n2 < MEDIA_APROVACAO:
        return "Vermelho Duplo"
    if n1 >= MEDIA_APROVACAO and n2 < MEDIA_APROVACAO:
        return "Queda p/ Vermelho"
    if n1 < MEDIA_APROVACAO and n2 >= MEDIA_APROVACAO:
        return "Recuperou"
    return "Verde"

def criar_excel_formatado(df, nome_planilha="Dados"):
    """
    Cria um arquivo Excel formatado usando pandas (método mais simples e confiável)
    """
    # Usar pandas para criar o Excel diretamente
    output = BytesIO()
    
    # Criar o arquivo Excel usando pandas
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nome_planilha, index=False)
        
        # Acessar a planilha para formatação
        workbook = writer.book
        worksheet = writer.sheets[nome_planilha]
        
        # Formatar cabeçalho
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Ajustar largura das colunas
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()

def calcula_indicadores(df):
    """
    Cria um dataframe por Aluno-Disciplina com:
      N1, N2, N3, N4, Media12, Soma12, ReqMediaProx2 (quanto precisa em média nos próximos 2 bimestres para fechar 6 no ano), Classificacao
    """
    # Criar coluna Bimestre
    df = df.copy()
    df["Bimestre"] = df["Periodo"].apply(mapear_bimestre)

    # Pivot por (Aluno, Turma, Disciplina)
    pivot = df.pivot_table(
        index=["Escola", "Turma", "Aluno", "Disciplina"],
        columns="Bimestre",
        values="Nota",
        aggfunc="mean"
    ).reset_index()

    # Renomear colunas 1..4 para N1..N4 (se existirem)
    rename_cols = {}
    for b in [1, 2, 3, 4]:
        if b in pivot.columns:
            rename_cols[b] = f"N{b}"
    pivot = pivot.rename(columns=rename_cols)

    # Calcular métricas dos 2 primeiros bimestres
    n1 = pivot.get("N1", pd.Series([np.nan] * len(pivot)))
    n2 = pivot.get("N2", pd.Series([np.nan] * len(pivot)))
    
    # Se não existir a coluna, criar uma série de NaN
    if isinstance(n1, float):
        n1 = pd.Series([np.nan] * len(pivot))
    if isinstance(n2, float):
        n2 = pd.Series([np.nan] * len(pivot))
    
    pivot["Soma12"] = n1.fillna(0) + n2.fillna(0)
    # Se um dos dois for NaN, a média 12 fica NaN (melhor do que assumir 0)
    pivot["Media12"] = (n1 + n2) / 2

    # Quanto precisa nos próximos dois bimestres (N3+N4) para fechar soma >= 24
    pivot["PrecisaSomarProx2"] = SOMA_FINAL_ALVO - pivot["Soma12"]
    pivot["ReqMediaProx2"] = pivot["PrecisaSomarProx2"] / 2

    # Classificação b1-b2
    pivot["Classificacao"] = [
        classificar_status_b1_b2(_n1, _n2, _m12)
        for _n1, _n2, _m12 in zip(pivot.get("N1", np.nan), pivot.get("N2", np.nan), pivot["Media12"])
    ]

    # Flags de alerta
    # "Corda Bamba": precisa de média >= 7 nos próximos dois bimestres
    pivot["CordaBamba"] = pivot["ReqMediaProx2"] >= 7

    # "Alerta": qualquer Vermelho Duplo ou Queda p/ Vermelho ou Corda Bamba
    pivot["Alerta"] = pivot["Classificacao"].isin(["Vermelho Duplo", "Queda p/ Vermelho"]) | pivot["CordaBamba"]

    return pivot

# -----------------------------
# UI – Entrada de dados
# -----------------------------
st.markdown("""
<div style="text-align: center; padding: 40px 20px; background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 15px; margin-bottom: 30px; box-shadow: 0 8px 25px rgba(30, 64, 175, 0.3);">
    <h1 style="color: white; margin: 0; font-size: 2.2em; font-weight: 700; text-shadow: 0 2px 4px rgba(0,0,0,0.3);">Superintendência Regional de Educação de Gurupi TO</h1>
    <h2 style="color: white; margin: 15px 0 0 0; font-weight: 600; font-size: 1.8em; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Painel SGE</h2>
    <h3 style="color: rgba(255,255,255,0.95); margin: 10px 0 0 0; font-weight: 500; font-size: 1.4em;">Notas, Frequência, Riscos e Alertas</h3>
    <p style="color: rgba(255,255,255,0.8); margin: 10px 0 0 0; font-size: 1.1em; font-weight: 400;">Análise dos 1º e 2º Bimestres</p>
</div>
""", unsafe_allow_html=True)

col_upl, col_info = st.columns([1, 2])
with col_upl:
    st.markdown("### Carregar Dados")
    arquivo = st.file_uploader("Planilha (.xlsx) do SGE", type=["xlsx"], help="Faça upload da planilha ou salve como 'dados.xlsx' na pasta")
with col_info:
    st.markdown("### Como usar")
    st.markdown("""
    **1.** Carregue sua planilha no uploader ou salve como `dados.xlsx`  
    **2.** Use os filtros na barra lateral para focar em turmas/disciplinas específicas  
    **3.** Analise os alertas, frequência e riscos dos alunos  
    **4.** Identifique quem precisa de atenção imediata
    """)

# Carregar
try:
    df = carregar_dados(arquivo)
    
    # Verificar tipo de planilha e rotear para interface apropriada
    tipo_planilha = df.attrs.get('tipo_planilha', 'notas_frequencia')
    
    if tipo_planilha == 'conteudo_aplicado':
        # Mostrar interface específica para conteúdo aplicado
        criar_interface_conteudo_aplicado(df)
        
        # Assinatura discreta do criador
        st.markdown("---")
        st.markdown(
            """
            <div style="text-align: center; margin-top: 40px; padding: 20px;">
                <p style="margin: 0;">
                    Desenvolvido por <strong style="color: #059669;">Alexandre Tolentino</strong> • 
                    <em>Painel SGE - Conteúdo Aplicado</em>
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )
    elif tipo_planilha == 'censo_escolar':
        # Mostrar interface específica para censo escolar
        criar_interface_censo_escolar(df)
        
        # Assinatura discreta do criador
        st.markdown("---")
        st.markdown(
            """
            <div style="text-align: center; margin-top: 40px; padding: 20px;">
                <p style="margin: 0;">
                    Desenvolvido por <strong style="color: #059669;">Alexandre Tolentino</strong> • 
                    <em>Painel SGE - Censo Escolar</em>
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.stop()
    else:
        # Continuar com interface padrão de notas/frequência
        pass
        
except FileNotFoundError:
    st.error("Não encontrei `dados.xlsx` na pasta e nenhum arquivo foi enviado no uploader.")
    
    # Assinatura discreta do criador (quando não há dados)
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; margin-top: 40px; padding: 20px;">
            <p style="margin: 0;">
                Desenvolvido por <strong style="color: #1e40af;">Alexandre Tolentino</strong> • 
                <em>Painel SGE</em>
            </p>
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.stop()

# Conferência mínima - Dados Gerais
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Visão Geral dos Dados</h3>
</div>
""", unsafe_allow_html=True)

colA, colB, colC, colD, colE = st.columns(5)

with colA:
    st.metric(
        label="Registros", 
        value=f"{len(df):,}".replace(",", "."),
        help="Total de linhas de dados na planilha"
    )
with colB:
    st.metric(
        label="Escolas", 
        value=df["Escola"].nunique() if "Escola" in df.columns else 0,
        help="Número de escolas diferentes"
    )
with colC:
    st.metric(
        label="Turmas", 
        value=df["Turma"].nunique() if "Turma" in df.columns else 0,
        help="Número de turmas diferentes"
    )
with colD:
    st.metric(
        label="Disciplinas", 
        value=df["Disciplina"].nunique() if "Disciplina" in df.columns else 0,
        help="Número de disciplinas diferentes"
    )
with colE:
    st.metric(
        label="Status", 
        value=df["Status"].nunique() if "Status" in df.columns else 0,
        help="Número de status diferentes"
    )

# Adicionar métrica de total de estudantes únicos
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">👥 Total de Estudantes</h3>
</div>
""", unsafe_allow_html=True)

col_total = st.columns(1)[0]
with col_total:
    total_estudantes = df["Aluno"].nunique() if "Aluno" in df.columns else 0
    st.metric(
        label="Estudantes Únicos", 
        value=f"{total_estudantes:,}".replace(",", "."),
        help="Total de estudantes únicos na escola (sem repetição por disciplina)"
    )


# -----------------------------
# Filtros laterais
# -----------------------------
st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Filtros</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1em; font-weight: 500;">Filtre os dados para análise específica</p>
</div>
""", unsafe_allow_html=True)

escolas = sorted(df["Escola"].dropna().unique().tolist()) if "Escola" in df.columns else []
status_opcoes = sorted(df["Status"].dropna().unique().tolist()) if "Status" in df.columns else []

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
    <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">Escola</h3>
</div>
""", unsafe_allow_html=True)
escola_sel = st.sidebar.selectbox("Selecione a escola:", ["Todas"] + escolas, help="Filtre por escola específica")

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
    <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">Status</h3>
</div>
""", unsafe_allow_html=True)
# Botões de ação rápida para status
col_s1, col_s2 = st.sidebar.columns(2)
with col_s1:
    if st.button("Todas", key="btn_todas_status", help="Selecionar todos os status"):
        st.session_state.status_selecionados = status_opcoes
with col_s2:
    if st.button("Limpar", key="btn_limpar_status", help="Limpar seleção"):
        st.session_state.status_selecionados = []

# Inicializar estado se não existir
if 'status_selecionados' not in st.session_state:
    st.session_state.status_selecionados = []

status_sel = st.sidebar.multiselect(
    "Selecione os status:", 
    status_opcoes, 
    default=st.session_state.status_selecionados,
    help="Use os botões acima para seleção rápida"
)

# Filtrar dados baseado na escola e status selecionados para mostrar opções relevantes
df_temp = df.copy()
if escola_sel != "Todas":
    df_temp = df_temp[df_temp["Escola"] == escola_sel]
if status_sel:  # Se algum status foi selecionado
    df_temp = df_temp[df_temp["Status"].isin(status_sel)]
else:  # Se nenhum status selecionado, mostra todos
    pass  # Mantém todos os status

turmas = sorted(df_temp["Turma"].dropna().unique().tolist()) if "Turma" in df_temp.columns else []
disciplinas = sorted(df_temp["Disciplina"].dropna().unique().tolist()) if "Disciplina" in df_temp.columns else []
alunos = sorted(df_temp["Aluno"].dropna().unique().tolist()) if "Aluno" in df_temp.columns else []

# Filtros com interface melhorada
st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
    <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">Turmas</h3>
</div>
""", unsafe_allow_html=True)
# Botões de ação rápida para turmas
col_t1, col_t2 = st.sidebar.columns(2)
with col_t1:
    if st.button("Todas", key="btn_todas_turmas", help="Selecionar todas as turmas"):
        st.session_state.turmas_selecionadas = turmas
with col_t2:
    if st.button("Limpar", key="btn_limpar_turmas", help="Limpar seleção"):
        st.session_state.turmas_selecionadas = []

# Inicializar estado se não existir
if 'turmas_selecionadas' not in st.session_state:
    st.session_state.turmas_selecionadas = []

turma_sel = st.sidebar.multiselect(
    "Selecione as turmas:", 
    turmas, 
    default=st.session_state.turmas_selecionadas,
    help="Use os botões acima para seleção rápida"
)

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
    <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">Disciplinas</h3>
</div>
""", unsafe_allow_html=True)
# Botões de ação rápida para disciplinas
col_d1, col_d2 = st.sidebar.columns(2)
with col_d1:
    if st.button("Todas", key="btn_todas_disc", help="Selecionar todas as disciplinas"):
        st.session_state.disciplinas_selecionadas = disciplinas
with col_d2:
    if st.button("Limpar", key="btn_limpar_disc", help="Limpar seleção"):
        st.session_state.disciplinas_selecionadas = []

# Inicializar estado se não existir
if 'disciplinas_selecionadas' not in st.session_state:
    st.session_state.disciplinas_selecionadas = []

disc_sel = st.sidebar.multiselect(
    "Selecione as disciplinas:", 
    disciplinas, 
    default=st.session_state.disciplinas_selecionadas,
    help="Use os botões acima para seleção rápida"
)

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 6px; padding: 8px 12px; margin: 6px 0; box-shadow: 0 1px 4px rgba(5, 150, 105, 0.1); border-left: 3px solid #059669;">
    <h3 style="color: #047857; margin: 0; font-size: 1em; font-weight: 600;">👤 Aluno</h3>
</div>
""", unsafe_allow_html=True)
aluno_sel = st.sidebar.selectbox("Selecione o aluno:", ["Todos"] + alunos, help="Filtre por aluno específico")

df_filt = df.copy()
if escola_sel != "Todas":
    df_filt = df_filt[df_filt["Escola"] == escola_sel]
if status_sel:  # Se algum status foi selecionado
    df_filt = df_filt[df_filt["Status"].isin(status_sel)]
else:  # Se nenhum status selecionado, mostra todos
    pass  # Mantém todos os status
if turma_sel:  # Se alguma turma foi selecionada
    df_filt = df_filt[df_filt["Turma"].isin(turma_sel)]
else:  # Se nenhuma turma selecionada, mostra todas
    pass  # Mantém todas as turmas

if disc_sel:  # Se alguma disciplina foi selecionada
    df_filt = df_filt[df_filt["Disciplina"].isin(disc_sel)]
else:  # Se nenhuma disciplina selecionada, mostra todas
    pass  # Mantém todas as disciplinas
if aluno_sel != "Todos":
    df_filt = df_filt[df_filt["Aluno"] == aluno_sel]

# Total de Estudantes Únicos (após filtros)
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Total de Estudantes (Filtrado)</h3>
</div>
""", unsafe_allow_html=True)

col_total_filt = st.columns(1)[0]
with col_total_filt:
    total_estudantes_filt = df_filt["Aluno"].nunique() if "Aluno" in df_filt.columns else 0
    st.metric(
        label="Estudantes Únicos", 
        value=f"{total_estudantes_filt:,}".replace(",", "."),
        help="Total de estudantes únicos considerando os filtros aplicados"
    )

# Métricas de Frequência na Visão Geral (após filtros)
if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
        <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Resumo de Frequência</h3>
    </div>
    """, unsafe_allow_html=True)
    
    colF1, colF2, colF3, colF4, colF5 = st.columns(5)
    
    # Função para classificar frequência (reutilizando a existente)
    def classificar_frequencia_geral(freq):
        if pd.isna(freq):
            return "Sem dados"
        elif freq < 75:
            return "Reprovado"
        elif freq < 80:
            return "Alto Risco"
        elif freq < 90:
            return "Risco Moderado"
        elif freq < 95:
            return "Ponto de Atenção"
        else:
            return "Meta Favorável"
    
    # Calcular frequências para visão geral (usando dados filtrados)
    # Agrupar apenas por Aluno para evitar duplicação quando aluno está em múltiplas turmas
    if "Frequencia Anual" in df_filt.columns:
        freq_geral = df_filt.groupby("Aluno")["Frequencia Anual"].last().reset_index()
        freq_geral = freq_geral.rename(columns={"Frequencia Anual": "Frequencia"})
    else:
        freq_geral = df_filt.groupby("Aluno")["Frequencia"].last().reset_index()
    
    freq_geral["Classificacao_Freq"] = freq_geral["Frequencia"].apply(classificar_frequencia_geral)
    contagem_freq_geral = freq_geral["Classificacao_Freq"].value_counts()
    
    # Calcular total de alunos para porcentagem
    total_alunos_freq = contagem_freq_geral.sum()
    
    with colF1:
        valor_reprovado = contagem_freq_geral.get("Reprovado", 0)
        percent_reprovado = (valor_reprovado / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #dbeafe, #bfdbfe); border-radius: 10px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
            <div style="font-size: 0.9em; font-weight: 600; color: #1e40af; margin-bottom: 8px;">< 75% (Reprovado)</div>
            <div style="font-size: 1.8em; font-weight: 700; color: #1e40af; margin: 8px 0;">{valor_reprovado}</div>
            <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_reprovado:.1f}%)</div>
        </div>
        """, unsafe_allow_html=True)
    with colF2:
        valor_alto_risco = contagem_freq_geral.get("Alto Risco", 0)
        percent_alto_risco = (valor_alto_risco / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #e0f2fe, #b3e5fc); border-radius: 10px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(14, 165, 233, 0.15); border-left: 4px solid #0ea5e9;">
            <div style="font-size: 0.9em; font-weight: 600; color: #0c4a6e; margin-bottom: 8px;">< 80% (Alto Risco)</div>
            <div style="font-size: 1.8em; font-weight: 700; color: #0c4a6e; margin: 8px 0;">{valor_alto_risco}</div>
            <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_alto_risco:.1f}%)</div>
        </div>
        """, unsafe_allow_html=True)
    with colF3:
        valor_risco_moderado = contagem_freq_geral.get("Risco Moderado", 0)
        percent_risco_moderado = (valor_risco_moderado / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #f0f9ff, #dbeafe); border-radius: 10px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(30, 64, 175, 0.15); border-left: 4px solid #1e40af;">
            <div style="font-size: 0.9em; font-weight: 600; color: #1e40af; margin-bottom: 8px;">< 90% (Risco Moderado)</div>
            <div style="font-size: 1.8em; font-weight: 700; color: #1e40af; margin: 8px 0;">{valor_risco_moderado}</div>
            <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_risco_moderado:.1f}%)</div>
        </div>
        """, unsafe_allow_html=True)
    with colF4:
        valor_ponto_atencao = contagem_freq_geral.get("Ponto de Atenção", 0)
        percent_ponto_atencao = (valor_ponto_atencao / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #eff6ff, #dbeafe); border-radius: 10px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
            <div style="font-size: 0.9em; font-weight: 600; color: #1e40af; margin-bottom: 8px;">< 95% (Ponto Atenção)</div>
            <div style="font-size: 1.8em; font-weight: 700; color: #1e40af; margin: 8px 0;">{valor_ponto_atencao}</div>
            <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_ponto_atencao:.1f}%)</div>
        </div>
        """, unsafe_allow_html=True)
    with colF5:
        valor_meta_favoravel = contagem_freq_geral.get("Meta Favorável", 0)
        percent_meta_favoravel = (valor_meta_favoravel / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #dbeafe, #bfdbfe); border-radius: 10px; padding: 15px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
            <div style="font-size: 0.9em; font-weight: 600; color: #1e40af; margin-bottom: 8px;">≥ 95% (Meta Favorável)</div>
            <div style="font-size: 1.8em; font-weight: 700; color: #1e40af; margin: 8px 0;">{valor_meta_favoravel}</div>
            <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_meta_favoravel:.1f}%)</div>
        </div>
        """, unsafe_allow_html=True)

# -----------------------------
# Indicadores e tabelas de risco
# -----------------------------
indic = calcula_indicadores(df_filt)

# KPIs - Análise de Notas Baixas
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h3 style="color: white; text-align: center; margin: 0; font-size: 1.5em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Análise de Notas Abaixo da Média</h3>
</div>
""", unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

notas_baixas_b1 = df_filt[df_filt["Periodo"].str.contains("Primeiro", case=False, na=False) & (df_filt["Nota"] < MEDIA_APROVACAO)]
notas_baixas_b2 = df_filt[df_filt["Periodo"].str.contains("Segundo", case=False, na=False) & (df_filt["Nota"] < MEDIA_APROVACAO)]

# Número de alunos únicos com notas baixas (não disciplinas)
alunos_notas_baixas_b1 = notas_baixas_b1["Aluno"].nunique() if "Aluno" in notas_baixas_b1.columns else 0
alunos_notas_baixas_b2 = notas_baixas_b2["Aluno"].nunique() if "Aluno" in notas_baixas_b2.columns else 0

# Calcular porcentagens baseadas no total de estudantes filtrados
total_estudantes_para_percent = total_estudantes_filt

with col1:
    percent_notas_b1 = (len(notas_baixas_b1) / len(df_filt) * 100) if len(df_filt) > 0 else 0
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #dbeafe, #bfdbfe); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
            <div style="font-size: 0.95em; font-weight: 600; color: #1e40af;">Notas < 6 – 1º Bim</div>
            <div style="background: rgba(30, 64, 175, 0.1); border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; color: #1e40af;">?</div>
        </div>
        <div style="font-size: 2em; font-weight: 700; color: #1e40af; margin: 8px 0;">{len(notas_baixas_b1)}</div>
        <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_notas_b1:.1f}%)</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Adicionar tooltip
    st.metric("", "", help="Total de notas abaixo de 6 no 1º bimestre. Inclui todas as disciplinas e alunos.")

with col2:
    percent_notas_b2 = (len(notas_baixas_b2) / len(df_filt) * 100) if len(df_filt) > 0 else 0
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #e0f2fe, #b3e5fc); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(14, 165, 233, 0.15); border-left: 4px solid #0ea5e9;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
            <div style="font-size: 0.95em; font-weight: 600; color: #0c4a6e;">Notas < 6 – 2º Bim</div>
            <div style="background: rgba(12, 74, 110, 0.1); border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; color: #0c4a6e;">?</div>
        </div>
        <div style="font-size: 2em; font-weight: 700; color: #0c4a6e; margin: 8px 0;">{len(notas_baixas_b2)}</div>
        <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_notas_b2:.1f}%)</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Adicionar tooltip
    st.metric("", "", help="Total de notas abaixo de 6 no 2º bimestre. Inclui todas as disciplinas e alunos.")

with col3:
    percent_alunos_b1 = (alunos_notas_baixas_b1 / total_estudantes_para_percent * 100) if total_estudantes_para_percent > 0 else 0
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #f0f9ff, #dbeafe); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(30, 64, 175, 0.15); border-left: 4px solid #1e40af;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
            <div style="font-size: 0.95em; font-weight: 600; color: #1e40af;">Alunos < 6 – 1º Bim</div>
            <div style="background: rgba(30, 64, 175, 0.1); border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; color: #1e40af;">?</div>
        </div>
        <div style="font-size: 2em; font-weight: 700; color: #1e40af; margin: 8px 0;">{alunos_notas_baixas_b1}</div>
        <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_alunos_b1:.1f}%)</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Adicionar tooltip
    st.metric("", "", help="Número de alunos únicos que tiveram pelo menos uma nota abaixo de 6 no 1º bimestre.")

with col4:
    percent_alunos_b2 = (alunos_notas_baixas_b2 / total_estudantes_para_percent * 100) if total_estudantes_para_percent > 0 else 0
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #eff6ff, #dbeafe); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
            <div style="font-size: 0.95em; font-weight: 600; color: #1e40af;">Alunos < 6 – 2º Bim</div>
            <div style="background: rgba(30, 64, 175, 0.1); border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; color: #1e40af;">?</div>
        </div>
        <div style="font-size: 2em; font-weight: 700; color: #1e40af; margin: 8px 0;">{alunos_notas_baixas_b2}</div>
        <div style="font-size: 1.3em; color: #64748b; font-weight: 600;">({percent_alunos_b2:.1f}%)</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Adicionar tooltip
    st.metric("", "", help="Número de alunos únicos que tiveram pelo menos uma nota abaixo de 6 no 2º bimestre.")

# KPIs - Alertas Críticos (com destaque visual)
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Alertas Críticos</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Situações que precisam de atenção imediata</p>
</div>
""", unsafe_allow_html=True)

col5, col6 = st.columns(2)

# Métricas de alerta com destaque visual
alerta_count = int(indic["Alerta"].sum())
corda_bamba_count = int(indic["CordaBamba"].sum())

# Calcular alunos únicos em alerta e corda bamba
alunos_unicos_alerta = indic[indic["Alerta"]]["Aluno"].nunique()
alunos_unicos_corda_bamba = indic[indic["CordaBamba"]]["Aluno"].nunique()

with col5:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #dbeafe, #bfdbfe); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
        <h3 style="color: #1e40af; margin: 0 0 15px 0; font-size: 1.1em; font-weight: 600;">Alunos-Disciplinas em ALERTA</h3>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div style="font-size: 2.5em; font-weight: 700; color: #1e40af;">{}</div>
            <div style="font-size: 2.5em; font-weight: 700; color: #64748b;">{} alunos</div>
        </div>
    </div>
    """.format(alerta_count, alunos_unicos_alerta), unsafe_allow_html=True)
    
    # Adicionar tooltip funcional
    st.metric("", "", help="Alunos-disciplinas em situação de risco (Vermelho Duplo, Queda p/ Vermelho ou Corda Bamba). O número entre parênteses mostra quantos alunos únicos estão em alerta.")

with col6:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #e0f2fe, #b3e5fc); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(14, 165, 233, 0.15); border-left: 4px solid #0ea5e9;">
        <h3 style="color: #0c4a6e; margin: 0 0 15px 0; font-size: 1.1em; font-weight: 600;">Corda Bamba</h3>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div style="font-size: 2.5em; font-weight: 700; color: #0c4a6e;">{}</div>
            <div style="font-size: 2.5em; font-weight: 700; color: #64748b;">{} alunos</div>
        </div>
    </div>
    """.format(corda_bamba_count, alunos_unicos_corda_bamba), unsafe_allow_html=True)
    
    # Adicionar tooltip funcional
    st.metric("", "", help="Corda Bamba são alunos que precisam tirar 7 ou mais nos próximos bimestres para recuperar e sair do limite da média mínima. O número maior mostra em quantas disciplinas eles aparecem; o número entre parênteses mostra quantos alunos diferentes estão nessa condição.")

# Resumo Executivo - Dashboard Principal
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Resumo Executivo</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Visão consolidada dos principais indicadores</p>
</div>
""", unsafe_allow_html=True)

# Métricas consolidadas em cards
col_res1, col_res2, col_res3, col_res4 = st.columns(4)

with col_res1:
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #dbeafe, #bfdbfe); border-radius: 8px; padding: 15px; margin: 10px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
        <h3 style="color: #1e40af; margin: 0 0 5px 0; font-size: 1em; font-weight: 600;">Alertas Críticos</h3>
        <p style="color: #64748b; margin: 0 0 8px 0; font-size: 0.85em;">Situações que precisam de atenção imediata</p>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div style="font-size: 1.5em; font-weight: 700; color: #1e40af;">{alerta_count}</div>
            <div style="font-size: 1.5em; font-weight: 700; color: #64748b;">{alunos_unicos_alerta} alunos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_res2:
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #e0f2fe, #b3e5fc); border-radius: 8px; padding: 15px; margin: 10px 0; box-shadow: 0 2px 8px rgba(14, 165, 233, 0.15); border-left: 4px solid #0ea5e9;">
        <h3 style="color: #0c4a6e; margin: 0 0 5px 0; font-size: 1em; font-weight: 600;">Corda Bamba</h3>
        <p style="color: #64748b; margin: 0 0 8px 0; font-size: 0.85em;">Precisam de média ≥ 7 nos próximos bimestres</p>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div style="font-size: 1.5em; font-weight: 700; color: #0c4a6e;">{corda_bamba_count}</div>
            <div style="font-size: 1.5em; font-weight: 700; color: #64748b;">{alunos_unicos_corda_bamba} alunos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_res3:
    # Calcular total de alunos com notas baixas
    total_alunos_notas_baixas = max(alunos_notas_baixas_b1, alunos_notas_baixas_b2)
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #f0f9ff, #dbeafe); border-radius: 8px; padding: 15px; margin: 10px 0; box-shadow: 0 2px 8px rgba(30, 64, 175, 0.15); border-left: 4px solid #1e40af;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
            <h3 style="color: #1e40af; margin: 0; font-size: 1em; font-weight: 600;">Notas Baixas</h3>
            <div style="background: rgba(30, 64, 175, 0.1); border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; color: #1e40af;">?</div>
        </div>
        <p style="color: #64748b; margin: 0 0 8px 0; font-size: 0.85em;">Alunos com notas abaixo de 6</p>
        <div style="font-size: 1.5em; font-weight: 700; color: #1e40af;">{total_alunos_notas_baixas}</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Adicionar tooltip usando st.metric
    st.metric("", "", help="Alunos únicos que tiveram pelo menos uma nota abaixo de 6 em qualquer bimestre. Considera o maior número entre 1º e 2º bimestres.")

with col_res4:
    # Calcular alunos com frequência baixa se disponível
    if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
        if "Frequencia Anual" in df_filt.columns:
            freq_baixa_count = len(df_filt[df_filt["Frequencia Anual"] < 95]["Aluno"].unique())
        else:
            freq_baixa_count = len(df_filt[df_filt["Frequencia"] < 95]["Aluno"].unique())
        
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #eff6ff, #dbeafe); border-radius: 10px; padding: 18px; margin: 5px 0; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.15); border-left: 4px solid #3b82f6;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <h3 style="color: #1e40af; margin: 0; font-size: 1.1em; font-weight: 600;">Frequência Baixa</h3>
                <div style="background: rgba(30, 64, 175, 0.1); border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; color: #1e40af;">?</div>
            </div>
            <p style="color: #64748b; margin: 0 0 8px 0; font-size: 0.85em;">Alunos com frequência < 95%</p>
            <div style="font-size: 2em; font-weight: 700; color: #1e40af;">{freq_baixa_count}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Adicionar tooltip usando st.metric
        st.metric("", "", help="Alunos únicos com frequência menor que 95%. Meta favorável é ≥ 95% de frequência.")
    else:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #f8fafc, #e2e8f0); border-radius: 8px; padding: 15px; margin: 10px 0; box-shadow: 0 2px 8px rgba(107, 114, 128, 0.1); border-left: 4px solid #64748b;">
            <h3 style="color: #374151; margin: 0 0 5px 0; font-size: 1em; font-weight: 600;">Frequência</h3>
            <p style="color: #64748b; margin: 0 0 8px 0; font-size: 0.85em;">Dados não disponíveis</p>
            <div style="font-size: 1.5em; font-weight: 700; color: #64748b;">N/A</div>
        </div>
        """, unsafe_allow_html=True)

# KPIs - Análise de Frequência
if "Frequencia Anual" in df_filt.columns:
    freq_title = "Análise de Frequência (Anual)"
    freq_subtitle = "Baseada na frequência anual dos alunos"
elif "Frequencia" in df_filt.columns:
    freq_title = "Análise de Frequência (Por Período)"
    freq_subtitle = "Baseada na frequência por período"
else:
    freq_title = "Análise de Frequência"
    freq_subtitle = "Dados de frequência não disponíveis"

st.markdown(f"""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">{freq_title}</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">{freq_subtitle}</p>
</div>
""", unsafe_allow_html=True)

col7, col8, col9, col10, col11 = st.columns(5)

# Função para classificar frequência
def classificar_frequencia(freq):
    if pd.isna(freq):
        return "Sem dados"
    elif freq < 75:
        return "Reprovado"
    elif freq < 80:
        return "Alto Risco"
    elif freq < 90:
        return "Risco Moderado"
    elif freq < 95:
        return "Ponto de Atenção"
    else:
        return "Meta Favorável"

# Calcular frequências se a coluna existir
if "Frequencia Anual" in df_filt.columns:
    # Usar frequência anual se disponível
    freq_atual = df_filt.groupby("Aluno")["Frequencia Anual"].last().reset_index()
    freq_atual = freq_atual.rename(columns={"Frequencia Anual": "Frequencia"})
    freq_atual["Classificacao_Freq"] = freq_atual["Frequencia"].apply(classificar_frequencia)
elif "Frequencia" in df_filt.columns:
    # Usar frequência do período se anual não estiver disponível
    freq_atual = df_filt.groupby("Aluno")["Frequencia"].last().reset_index()
    freq_atual["Classificacao_Freq"] = freq_atual["Frequencia"].apply(classificar_frequencia)
    
    # Contar por classificação
    contagem_freq = freq_atual["Classificacao_Freq"].value_counts()
    
    with col7:
        st.metric(
            label="< 75% (Reprovado)", 
            value=contagem_freq.get("Reprovado", 0),
            help="Alunos reprovados por frequência (abaixo de 75%)"
        )
    with col8:
        st.metric(
            label="< 80% (Alto Risco)", 
            value=contagem_freq.get("Alto Risco", 0),
            help="Alunos em alto risco de reprovação por frequência"
        )
    with col9:
        st.metric(
            label="< 90% (Risco Moderado)", 
            value=contagem_freq.get("Risco Moderado", 0),
            help="Alunos com risco moderado de reprovação"
        )
    with col10:
        st.metric(
            label="< 95% (Ponto Atenção)", 
            value=contagem_freq.get("Ponto de Atenção", 0),
            help="Alunos que precisam de atenção na frequência"
        )
    with col11:
        st.metric(
            label="≥ 95% (Meta Favorável)", 
            value=contagem_freq.get("Meta Favorável", 0),
            help="Alunos com frequência dentro da meta"
        )
else:
    col7.metric("< 75% (Reprovado)", "N/A")
    col8.metric("< 80% (Alto Risco)", "N/A")
    col9.metric("< 90% (Risco Moderado)", "N/A")
    col10.metric("< 95% (Ponto Atenção)", "N/A")
    col11.metric("≥ 95% (Meta Favorável)", "N/A")

# Seção expandível: Análise Detalhada de Frequência
if "Frequencia Anual" in df_filt.columns:
    expander_title = "Análise Detalhada de Frequência (Anual)"
elif "Frequencia" in df_filt.columns:
    expander_title = "Análise Detalhada de Frequência (Por Período)"
else:
    expander_title = "Análise Detalhada de Frequência"

with st.expander(expander_title):
    if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
        # Tabela de frequência por aluno (agrupando apenas por aluno para evitar duplicação)
        if "Frequencia Anual" in df_filt.columns:
            freq_detalhada = df_filt.groupby("Aluno")["Frequencia Anual"].last().reset_index()
            freq_detalhada = freq_detalhada.rename(columns={"Frequencia Anual": "Frequencia"})
        else:
            freq_detalhada = df_filt.groupby("Aluno")["Frequencia"].last().reset_index()
        freq_detalhada["Classificacao_Freq"] = freq_detalhada["Frequencia"].apply(classificar_frequencia)
        freq_detalhada = freq_detalhada.sort_values("Aluno")
        
        # Função para colorir frequência
        def color_frequencia(val):
            if val == "Reprovado":
                return "background-color: #f8d7da; color: #721c24"  # Vermelho
            elif val == "Alto Risco":
                return "background-color: #f5c6cb; color: #721c24"  # Vermelho claro
            elif val == "Risco Moderado":
                return "background-color: #fff3cd; color: #856404"  # Amarelo
            elif val == "Ponto de Atenção":
                return "background-color: #ffeaa7; color: #856404"  # Amarelo claro
            elif val == "Meta Favorável":
                return "background-color: #d4edda; color: #155724"  # Verde
            else:
                return "background-color: #e2e3e5; color: #383d41"  # Cinza
        
        # Formatar frequência
        freq_detalhada["Frequencia_Formatada"] = freq_detalhada["Frequencia"].apply(
            lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A"
        )
        
        # Aplicar cores
        styled_freq = freq_detalhada[["Aluno", "Frequencia_Formatada", "Classificacao_Freq"]]\
            .style.applymap(color_frequencia, subset=["Classificacao_Freq"])
        
        st.dataframe(styled_freq, use_container_width=True)
        
        # Botão de exportação para frequência
        col_export5, col_export6 = st.columns([1, 4])
        with col_export5:
            if st.button("📊 Exportar Frequência", key="export_frequencia", help="Baixar planilha com análise de frequência"):
                excel_data = criar_excel_formatado(freq_detalhada[["Aluno", "Turma", "Frequencia_Formatada", "Classificacao_Freq"]], "Analise_Frequencia")
                st.download_button(
                    label="Baixar Excel",
                    data=excel_data,
                    file_name="analise_frequencia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        # Legenda de frequência
        st.markdown("###  Legenda de Frequência")
        col_leg1, col_leg2, col_leg3 = st.columns(3)
        with col_leg1:
            st.markdown("""
            **< 75%**: Reprovado por frequência  
            **< 80%**: Alto risco de reprovação
            """)
        with col_leg2:
            st.markdown("""
            **< 90%**: Risco moderado  
            **< 95%**: Ponto de atenção
            """)
        with col_leg3:
            st.markdown("""
            **≥ 95%**: Meta favorável  
            **Sem dados**: Frequência não informada
            """)
    else:
        st.info("Dados de frequência não disponíveis na planilha.")


st.markdown("---")

# Tabela: Alunos-Disciplinas em ALERTA (com cálculo de necessidade para 3º e 4º)
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Alunos/Disciplinas em ALERTA</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Situações que precisam de atenção imediata</p>
</div>
""", unsafe_allow_html=True)
cols_visiveis = ["Aluno", "Turma", "Disciplina", "N1", "N2", "Media12", "Classificacao", "ReqMediaProx2", "CordaBamba"]
tabela_alerta = (indic[indic["Alerta"]]
                 .copy()
                 .sort_values(["Turma", "Aluno", "Disciplina"]))
for c in ["N1", "N2", "Media12", "ReqMediaProx2"]:
    if c in tabela_alerta.columns:
        # Formatar para 1 casa decimal, removendo .0 desnecessário
        tabela_alerta[c] = tabela_alerta[c].round(1)
        tabela_alerta[c] = tabela_alerta[c].apply(lambda x: f"{x:.1f}".rstrip('0').rstrip('.') if pd.notna(x) else x)

# Função para aplicar cores na classificação (definida antes de usar)
def color_classification(val):
    if val == "Verde":
        return "background-color: #d4edda; color: #155724"  # Verde claro
    elif val == "Vermelho Duplo":
        return "background-color: #f8d7da; color: #721c24"  # Vermelho claro
    elif val == "Queda p/ Vermelho":
        return "background-color: #fff3cd; color: #856404"  # Amarelo claro
    elif val == "Recuperou":
        return "background-color: #cce5ff; color: #004085"  # Azul claro
    elif val == "Incompleto":
        return "background-color: #e2e3e5; color: #383d41"  # Cinza claro
    else:
        return ""

# Aplicar cores na tabela de alertas também
if len(tabela_alerta) > 0:
    styled_alerta = tabela_alerta[cols_visiveis].style.applymap(color_classification, subset=["Classificacao"])
    st.dataframe(styled_alerta, use_container_width=True)
    
    # Botão de exportação para alertas
    col_export1, col_export2 = st.columns([1, 4])
    with col_export1:
        if st.button("📊 Exportar Alertas", key="export_alertas", help="Baixar planilha com alunos em alerta"):
            excel_data = criar_excel_formatado(tabela_alerta[cols_visiveis], "Alunos_em_Alerta")
            st.download_button(
                label="Baixar Excel",
                data=excel_data,
                file_name="alunos_em_alerta.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.dataframe(pd.DataFrame(columns=cols_visiveis), use_container_width=True)

# Tabela: Panorama Geral de Notas (todos para diagnóstico rápido)
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Panorama Geral de Notas (B1→B2)</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Visão completa de todos os alunos e disciplinas</p>
</div>
""", unsafe_allow_html=True)
tab_diag = indic.copy()
for c in ["N1", "N2", "Media12", "ReqMediaProx2"]:
    if c in tab_diag.columns:
        # Formatar para 1 casa decimal, removendo .0 desnecessário
        tab_diag[c] = tab_diag[c].round(1)
        tab_diag[c] = tab_diag[c].apply(lambda x: f"{x:.1f}".rstrip('0').rstrip('.') if pd.notna(x) else x)



# Aplicar estilização
styled_table = tab_diag[["Aluno", "Turma", "Disciplina", "N1", "N2", "Media12", "Classificacao", "ReqMediaProx2"]]\
    .sort_values(["Turma", "Aluno", "Disciplina"])\
    .style.applymap(color_classification, subset=["Classificacao"])

st.dataframe(styled_table, use_container_width=True)

# Botão de exportação para panorama de notas
col_export3, col_export4 = st.columns([1, 4])
with col_export3:
        if st.button("📊 Exportar Panorama", key="export_panorama", help="Baixar planilha com panorama geral de notas"):
            excel_data = criar_excel_formatado(tab_diag[["Aluno", "Turma", "Disciplina", "N1", "N2", "Media12", "Classificacao", "ReqMediaProx2"]], "Panorama_Geral_Notas")
            st.download_button(
                label="Baixar Excel",
                data=excel_data,
                file_name="panorama_notas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Legenda de cores
st.markdown("### Legenda de Cores")
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("""
    **Verde**: Aluno está bem (N1≥6 e N2≥6)  
    **Vermelho Duplo**: Risco alto (N1<6 e N2<6)
    """)
with col2:
    st.markdown("""
    **Queda p/ Vermelho**: Piorou (N1≥6 e N2<6)  
    **Recuperou**: Melhorou (N1<6 e N2≥6)
    """)
with col3:
    st.markdown("""
    **Incompleto**: Falta nota  
    **Corda Bamba**: Precisa ≥7 nos próximos 2
    """)

st.markdown(
    """
    **Interpretação rápida**  
    - *Vermelho Duplo*: segue risco alto (dois bimestres < 6).  
    - *Queda p/ Vermelho*: atenção no 3º bimestre (piora do 1º para o 2º).  
    - *Recuperou*: saiu do vermelho no 2º.  
    - *Corda Bamba*: para fechar média 6 no ano, precisa tirar **≥ 7,0** em média no 3º e 4º.
    """
)

# Gráficos: Notas e Frequência por Disciplina (movidos para o final)
st.markdown("---")
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Análises Gráficas</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Visualizações complementares dos dados</p>
</div>
""", unsafe_allow_html=True)

col_graf1, col_graf2 = st.columns(2)

# Gráfico: Notas abaixo de 6 por Disciplina (1º e 2º bimestres)
with col_graf1:
    with st.expander("Notas Abaixo da Média por Disciplina"):
        base_baixas = pd.concat([notas_baixas_b1, notas_baixas_b2], ignore_index=True)
        if len(base_baixas) > 0:
            # Contar notas por disciplina
            contagem = base_baixas.groupby("Disciplina")["Nota"].count().reset_index()
            contagem = contagem.rename(columns={"Nota": "Qtd Notas < 6"})
            
            # Ordenar em ordem decrescente (maior para menor)
            contagem = contagem.sort_values("Qtd Notas < 6", ascending=False).reset_index(drop=True)
            
            # Adicionar coluna de cores intercaladas baseada na posição após ordenação
            contagem['Cor'] = ['#1e40af' if i % 2 == 0 else '#059669' for i in range(len(contagem))]
            
            fig = px.bar(contagem, x="Disciplina", y="Qtd Notas < 6", 
                        title="Notas abaixo da média (1º + 2º Bimestre)",
                        color="Cor",
                        color_discrete_map={'#1e40af': '#1e40af', '#059669': '#059669'})
            
            # Forçar a ordem das disciplinas no eixo X
            fig.update_layout(
                xaxis_title=None, 
                yaxis_title="Quantidade", 
                bargap=0.25, 
                showlegend=False, 
                xaxis_tickangle=45,
                xaxis={'categoryorder': 'array', 'categoryarray': contagem['Disciplina'].tolist()}
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Botão de exportação para dados do gráfico
            col_export_graf1, col_export_graf2 = st.columns([1, 4])
            with col_export_graf1:
                if st.button("📊 Exportar Dados do Gráfico", key="export_grafico_notas", help="Baixar planilha com dados do gráfico de notas por disciplina"):
                    # Preparar dados para exportação (remover coluna de cor)
                    dados_export = contagem[['Disciplina', 'Qtd Notas < 6']].copy()
                    dados_export = dados_export.rename(columns={'Qtd Notas < 6': 'Quantidade_Notas_Abaixo_6'})
                    
                    excel_data = criar_excel_formatado(dados_export, "Notas_Por_Disciplina")
                    st.download_button(
                        label="Baixar Excel",
                        data=excel_data,
                        file_name="notas_por_disciplina.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("Sem notas abaixo da média para os filtros atuais.")

# Gráfico: Distribuição de Frequência por Faixas
with col_graf2:
    with st.expander("Distribuição de Frequência por Faixas"):
        if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
            # Usar os mesmos dados do Resumo de Frequência
            if "Frequencia Anual" in df_filt.columns:
                freq_geral = df_filt.groupby(["Aluno", "Turma"])["Frequencia Anual"].last().reset_index()
                freq_geral = freq_geral.rename(columns={"Frequencia Anual": "Frequencia"})
            else:
                freq_geral = df_filt.groupby(["Aluno", "Turma"])["Frequencia"].last().reset_index()
            
            freq_geral["Classificacao_Freq"] = freq_geral["Frequencia"].apply(classificar_frequencia_geral)
            contagem_freq_geral = freq_geral["Classificacao_Freq"].value_counts()
            
            # Preparar dados para o gráfico
            dados_grafico = []
            cores = {
                "Reprovado": "#dc2626",
                "Alto Risco": "#ea580c", 
                "Risco Moderado": "#d97706",
                "Ponto de Atenção": "#f59e0b",
                "Meta Favorável": "#16a34a"
            }
            
            for categoria, quantidade in contagem_freq_geral.items():
                if categoria != "Sem dados":  # Excluir "Sem dados" do gráfico
                    dados_grafico.append({
                        "Categoria": categoria,
                        "Quantidade": quantidade,
                        "Cor": cores.get(categoria, "#6b7280")
                    })
            
            if dados_grafico:
                df_grafico = pd.DataFrame(dados_grafico)
                
                # Criar gráfico de barras
                fig_freq = px.bar(df_grafico, x="Categoria", y="Quantidade", 
                                 title="Distribuição de Alunos por Faixa de Frequência",
                                 color="Categoria", 
                                 color_discrete_map=cores)
                fig_freq.update_layout(xaxis_title=None, yaxis_title="Número de Alunos", 
                                     bargap=0.25, showlegend=False, xaxis_tickangle=45)
                st.plotly_chart(fig_freq, use_container_width=True)
                
                # Botão de exportação para dados do gráfico de frequência
                col_export_graf3, col_export_graf4 = st.columns([1, 4])
                with col_export_graf3:
                    if st.button("📊 Exportar Dados do Gráfico", key="export_grafico_freq", help="Baixar planilha com dados do gráfico de frequência"):
                        # Preparar dados para exportação
                        dados_export_freq = df_grafico[['Categoria', 'Quantidade']].copy()
                        dados_export_freq = dados_export_freq.rename(columns={'Quantidade': 'Numero_Alunos'})
                        
                        excel_data = criar_excel_formatado(dados_export_freq, "Frequencia_Por_Faixa")
                        st.download_button(
                            label="Baixar Excel",
                            data=excel_data,
                            file_name="frequencia_por_faixa.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                # Estatísticas adicionais
                st.markdown("**Resumo das Faixas de Frequência:**")
                col_stat1, col_stat2, col_stat3 = st.columns(3)
                with col_stat1:
                    total_alunos = contagem_freq_geral.sum()
                    st.metric("Total de Alunos", total_alunos, help="Total de alunos considerados na análise de frequência")
                with col_stat2:
                    alunos_risco = contagem_freq_geral.get("Reprovado", 0) + contagem_freq_geral.get("Alto Risco", 0)
                    st.metric("Alunos em Risco", alunos_risco, help="Alunos reprovados ou em alto risco de reprovação por frequência")
                with col_stat3:
                    alunos_meta = contagem_freq_geral.get("Meta Favorável", 0)
                    percentual_meta = (alunos_meta / total_alunos * 100) if total_alunos > 0 else 0
                    st.metric("Meta Favorável", f"{percentual_meta:.1f}%", help="Percentual de alunos com frequência ≥ 95% (meta favorável)")
            else:
                st.info("Sem dados de frequência para exibir.")
        else:
            st.info("Dados de frequência não disponíveis na planilha.")

# Seção expandível: Análise Cruzada Nota x Frequência (movida para o final)
st.markdown("---")
st.markdown("""
<div style="background: linear-gradient(135deg, #1e40af, #3b82f6); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">Análise Cruzada</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Cruzamento entre Notas e Frequência</p>
</div>
""", unsafe_allow_html=True)

with st.expander("Análise Cruzada: Notas x Frequência"):
    if ("Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns) and len(indic) > 0:
        # Combinar dados de notas e frequência (priorizando Frequencia Anual)
        if "Frequencia Anual" in df_filt.columns:
            freq_alunos = df_filt.groupby(["Aluno", "Turma"])["Frequencia Anual"].last().reset_index()
            freq_alunos = freq_alunos.rename(columns={"Frequencia Anual": "Frequencia"})
        else:
            freq_alunos = df_filt.groupby(["Aluno", "Turma"])["Frequencia"].last().reset_index()
        freq_alunos["Classificacao_Freq"] = freq_alunos["Frequencia"].apply(classificar_frequencia)
        
        # Merge com indicadores de notas
        cruzada = indic.merge(freq_alunos, on=["Aluno", "Turma"], how="left")
        
        # Criar matriz de cruzamento
        matriz_cruzada = cruzada.groupby(["Classificacao", "Classificacao_Freq"]).size().unstack(fill_value=0)
        
        if not matriz_cruzada.empty:
            st.markdown("**Matriz de Cruzamento: Classificação de Notas x Frequência**")
            st.dataframe(matriz_cruzada, use_container_width=True)
            
            # Análise de alunos com frequência abaixo de 95%
            freq_baixa = cruzada[cruzada["Frequencia"] < 95]
            
            if len(freq_baixa) > 0:
                st.markdown("### Alunos com Frequência Abaixo de 95% (Cruzamento Notas x Frequência)")
                # Mostrar apenas colunas relevantes para frequência baixa
                freq_baixa_display = freq_baixa[["Aluno", "Turma", "Disciplina", "Classificacao", "Classificacao_Freq", "Frequencia"]].copy()
                # Formatar frequência
                freq_baixa_display["Frequencia"] = freq_baixa_display["Frequencia"].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A"
                )
                st.dataframe(freq_baixa_display, use_container_width=True)
                
                # Botão de exportação para alunos com frequência baixa
                col_export_freq_baixa1, col_export_freq_baixa2 = st.columns([1, 4])
                with col_export_freq_baixa1:
                    if st.button("📊 Exportar Cruzamento", key="export_freq_baixa", help="Baixar planilha com cruzamento de notas e frequência (alunos com frequência < 95%)"):
                        excel_data = criar_excel_formatado(freq_baixa_display, "Cruzamento_Notas_Freq")
                        st.download_button(
                            label="Baixar Excel",
                            data=excel_data,
                            file_name="cruzamento_notas_frequencia.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.info("Todos os alunos têm frequência ≥ 95% (Meta Favorável).")
        else:
            st.info("Dados insuficientes para análise cruzada.")
    else:
        st.info("Dados de frequência ou notas não disponíveis para análise cruzada.")

# Botão para baixar todas as planilhas em uma única planilha Excel
st.markdown("---")
st.markdown("""
<div style="background: linear-gradient(135deg, #059669, #10b981); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(5, 150, 105, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">📊 Exportação Completa</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Baixar todas as análises em uma única planilha Excel</p>
</div>
""", unsafe_allow_html=True)

col_export_all1, col_export_all2 = st.columns([1, 4])
with col_export_all1:
    if st.button("📊 Baixar Tudo", key="export_tudo", help="Baixar todas as análises em uma única planilha Excel com múltiplas abas"):
        # Criar arquivo Excel com múltiplas abas
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Aba 1: Alunos em Alerta
            if len(tabela_alerta) > 0:
                tabela_alerta[cols_visiveis].to_excel(writer, sheet_name="Alunos_em_Alerta", index=False)
            
            # Aba 2: Panorama Geral de Notas
            tab_diag[["Aluno", "Turma", "Disciplina", "N1", "N2", "Media12", "Classificacao", "ReqMediaProx2"]].to_excel(
                writer, sheet_name="Panorama_Geral_Notas", index=False)
            
            # Aba 3: Análise de Frequência (se disponível)
            if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
                if "Frequencia Anual" in df_filt.columns:
                    freq_detalhada = df_filt.groupby(["Aluno", "Turma"])["Frequencia Anual"].last().reset_index()
                    freq_detalhada = freq_detalhada.rename(columns={"Frequencia Anual": "Frequencia"})
                else:
                    freq_detalhada = df_filt.groupby(["Aluno", "Turma"])["Frequencia"].last().reset_index()
                
                freq_detalhada["Classificacao_Freq"] = freq_detalhada["Frequencia"].apply(classificar_frequencia)
                freq_detalhada["Frequencia_Formatada"] = freq_detalhada["Frequencia"].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A"
                )
                freq_detalhada[["Aluno", "Turma", "Frequencia_Formatada", "Classificacao_Freq"]].to_excel(
                    writer, sheet_name="Analise_Frequencia", index=False)
            
            # Aba 4: Notas por Disciplina (se houver dados)
            base_baixas = pd.concat([notas_baixas_b1, notas_baixas_b2], ignore_index=True)
            if len(base_baixas) > 0:
                contagem = base_baixas.groupby("Disciplina")["Nota"].count().reset_index()
                contagem = contagem.rename(columns={"Nota": "Quantidade_Notas_Abaixo_6"})
                contagem = contagem.sort_values("Quantidade_Notas_Abaixo_6", ascending=False).reset_index(drop=True)
                contagem.to_excel(writer, sheet_name="Notas_Por_Disciplina", index=False)
            
            # Aba 5: Frequência por Faixas (se disponível)
            if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
                if "Frequencia Anual" in df_filt.columns:
                    freq_geral = df_filt.groupby(["Aluno", "Turma"])["Frequencia Anual"].last().reset_index()
                    freq_geral = freq_geral.rename(columns={"Frequencia Anual": "Frequencia"})
                else:
                    freq_geral = df_filt.groupby(["Aluno", "Turma"])["Frequencia"].last().reset_index()
                
                freq_geral["Classificacao_Freq"] = freq_geral["Frequencia"].apply(classificar_frequencia_geral)
                contagem_freq_geral = freq_geral["Classificacao_Freq"].value_counts()
                
                dados_grafico = []
                for categoria, quantidade in contagem_freq_geral.items():
                    if categoria != "Sem dados":
                        dados_grafico.append({
                            "Categoria": categoria,
                            "Numero_Alunos": quantidade
                        })
                
                if dados_grafico:
                    df_grafico = pd.DataFrame(dados_grafico)
                    df_grafico.to_excel(writer, sheet_name="Frequencia_Por_Faixa", index=False)
            
            # Aba 6: Cruzamento Notas x Frequência (se disponível)
            if ("Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns) and len(indic) > 0:
                if "Frequencia Anual" in df_filt.columns:
                    freq_alunos = df_filt.groupby(["Aluno", "Turma"])["Frequencia Anual"].last().reset_index()
                    freq_alunos = freq_alunos.rename(columns={"Frequencia Anual": "Frequencia"})
                else:
                    freq_alunos = df_filt.groupby(["Aluno", "Turma"])["Frequencia"].last().reset_index()
                
                freq_alunos["Classificacao_Freq"] = freq_alunos["Frequencia"].apply(classificar_frequencia)
                cruzada = indic.merge(freq_alunos, on=["Aluno", "Turma"], how="left")
                freq_baixa = cruzada[cruzada["Frequencia"] < 95]
                
                if len(freq_baixa) > 0:
                    freq_baixa_display = freq_baixa[["Aluno", "Turma", "Disciplina", "Classificacao", "Classificacao_Freq", "Frequencia"]].copy()
                    freq_baixa_display["Frequencia"] = freq_baixa_display["Frequencia"].apply(
                        lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A"
                    )
                    freq_baixa_display.to_excel(writer, sheet_name="Cruzamento_Notas_Freq", index=False)
            
            # Aba 7: Alunos Duplicados (se houver)
            alunos_turmas = df_filt.groupby("Aluno")["Turma"].nunique().reset_index()
            alunos_turmas = alunos_turmas.rename(columns={"Turma": "Qtd_Turmas"})
            alunos_duplicados = alunos_turmas[alunos_turmas["Qtd_Turmas"] > 1].copy()
            
            if len(alunos_duplicados) > 0:
                # Criar formato com colunas separadas para cada turma
                export_data = []
                for _, row in alunos_duplicados.iterrows():
                    aluno = row["Aluno"]
                    qtd_turmas = row["Qtd_Turmas"]
                    turmas_aluno = df_filt[df_filt["Aluno"] == aluno]["Turma"].unique().tolist()
                    turmas_aluno = sorted(turmas_aluno)
                    
                    # Criar linha com colunas separadas
                    linha = {
                        "Aluno": aluno,
                        "Qtd_Turmas": qtd_turmas
                    }
                    
                    # Adicionar cada turma em uma coluna separada
                    for i, turma in enumerate(turmas_aluno, 1):
                        linha[f"Turma_{i}"] = turma
                    
                    # Preencher colunas vazias com None para alunos com menos turmas
                    max_turmas = alunos_duplicados["Qtd_Turmas"].max()
                    for i in range(len(turmas_aluno) + 1, max_turmas + 1):
                        linha[f"Turma_{i}"] = None
                    
                    export_data.append(linha)
                
                df_export = pd.DataFrame(export_data)
                df_export = df_export.sort_values(["Qtd_Turmas", "Aluno"], ascending=[False, True])
                df_export.to_excel(writer, sheet_name="Alunos_Duplicados", index=False)
        
        output.seek(0)
        st.download_button(
            label="📥 Baixar Planilha Completa",
            data=output.getvalue(),
            file_name="painel_sge_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Seção: Identificação de Alunos em Múltiplas Turmas
st.markdown("---")
st.markdown("""
<div style="background: linear-gradient(135deg, #dc2626, #ef4444); border-radius: 12px; padding: 25px; margin: 20px 0; box-shadow: 0 4px 15px rgba(220, 38, 38, 0.2);">
    <h2 style="color: white; text-align: center; margin: 0; font-size: 1.7em; font-weight: 700; text-shadow: 0 1px 3px rgba(0,0,0,0.3);">🔍 Identificação de Alunos Duplicados</h2>
    <p style="color: rgba(255,255,255,0.9); text-align: center; margin: 8px 0 0 0; font-size: 1.1em; font-weight: 500;">Detecção de alunos que aparecem em múltiplas turmas</p>
</div>
""", unsafe_allow_html=True)

# Identificar alunos em múltiplas turmas
alunos_turmas = df_filt.groupby("Aluno")["Turma"].nunique().reset_index()
alunos_turmas = alunos_turmas.rename(columns={"Turma": "Qtd_Turmas"})

# Filtrar apenas alunos com mais de uma turma
alunos_duplicados = alunos_turmas[alunos_turmas["Qtd_Turmas"] > 1].copy()

if len(alunos_duplicados) > 0:
    # Criar dataframe detalhado com todas as turmas de cada aluno duplicado
    alunos_detalhados = []
    
    for _, row in alunos_duplicados.iterrows():
        aluno = row["Aluno"]
        qtd_turmas = row["Qtd_Turmas"]
        
        # Obter todas as turmas deste aluno
        turmas_aluno = df_filt[df_filt["Aluno"] == aluno]["Turma"].unique().tolist()
        turmas_str = ", ".join(sorted(turmas_aluno))
        
        alunos_detalhados.append({
            "Aluno": aluno,
            "Qtd_Turmas": qtd_turmas,
            "Turmas": turmas_str
        })
    
    df_alunos_duplicados = pd.DataFrame(alunos_detalhados)
    df_alunos_duplicados = df_alunos_duplicados.sort_values(["Qtd_Turmas", "Aluno"], ascending=[False, True])
    
    # Função para colorir quantidade de turmas
    def color_qtd_turmas(val):
        if val == 2:
            return "background-color: #fef3c7; color: #92400e"  # Amarelo para duplicidade
        elif val == 3:
            return "background-color: #fed7aa; color: #9a3412"  # Laranja para triplicidade
        elif val >= 4:
            return "background-color: #fecaca; color: #991b1b"  # Vermelho para 4+ turmas
        else:
            return ""
    
    # Aplicar cores
    styled_duplicados = df_alunos_duplicados.style.applymap(color_qtd_turmas, subset=["Qtd_Turmas"])
    
    st.dataframe(styled_duplicados, use_container_width=True)
    
    # Métricas resumidas
    col_dup1, col_dup2, col_dup3 = st.columns(3)
    
    with col_dup1:
        total_duplicados = len(df_alunos_duplicados)
        st.metric(
            label="Total de Alunos Duplicados", 
            value=total_duplicados,
            help="Alunos que aparecem em mais de uma turma"
        )
    
    with col_dup2:
        duplicidade = len(df_alunos_duplicados[df_alunos_duplicados["Qtd_Turmas"] == 2])
        st.metric(
            label="Duplicidade (2 turmas)", 
            value=duplicidade,
            help="Alunos que aparecem em exatamente 2 turmas"
        )
    
    with col_dup3:
        triplicidade_mais = len(df_alunos_duplicados[df_alunos_duplicados["Qtd_Turmas"] >= 3])
        st.metric(
            label="Triplicidade+ (3+ turmas)", 
            value=triplicidade_mais,
            help="Alunos que aparecem em 3 ou mais turmas"
        )
    
    # Botão de exportação
    col_export_dup1, col_export_dup2 = st.columns([1, 4])
    with col_export_dup1:
        if st.button("📊 Exportar Duplicados", key="export_duplicados", help="Baixar planilha com alunos em múltiplas turmas"):
            # Criar formato com colunas separadas para cada turma
            export_data = []
            for _, row in df_alunos_duplicados.iterrows():
                aluno = row["Aluno"]
                qtd_turmas = row["Qtd_Turmas"]
                turmas_aluno = df_filt[df_filt["Aluno"] == aluno]["Turma"].unique().tolist()
                turmas_aluno = sorted(turmas_aluno)
                
                # Criar linha com colunas separadas
                linha = {
                    "Aluno": aluno,
                    "Qtd_Turmas": qtd_turmas
                }
                
                # Adicionar cada turma em uma coluna separada
                for i, turma in enumerate(turmas_aluno, 1):
                    linha[f"Turma_{i}"] = turma
                
                # Preencher colunas vazias com None para alunos com menos turmas
                max_turmas = df_alunos_duplicados["Qtd_Turmas"].max()
                for i in range(len(turmas_aluno) + 1, max_turmas + 1):
                    linha[f"Turma_{i}"] = None
                
                export_data.append(linha)
            
            df_export = pd.DataFrame(export_data)
            excel_data = criar_excel_formatado(df_export, "Alunos_Duplicados")
            st.download_button(
                label="Baixar Excel",
                data=excel_data,
                file_name="alunos_duplicados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Legenda
    st.markdown("### Legenda de Cores")
    col_leg_dup1, col_leg_dup2, col_leg_dup3 = st.columns(3)
    with col_leg_dup1:
        st.markdown("""
        **2 turmas**: Duplicidade (amarelo)  
        **3 turmas**: Triplicidade (laranja)
        """)
    with col_leg_dup2:
        st.markdown("""
        **4+ turmas**: Múltiplas turmas (vermelho)  
        **Ação**: Verificar dados
        """)
    with col_leg_dup3:
        st.markdown("""
        **Possíveis causas**:  
        • Erro de digitação  
        • Transferência não registrada
        """)
    
    # Aviso importante
    st.warning("""
    ⚠️ **Atenção**: Alunos em múltiplas turmas podem indicar:
    - Erros de digitação nos dados
    - Transferências não registradas adequadamente
    - Inconsistências na base de dados
    
    Recomenda-se verificar e corrigir essas situações.
    """)
    
else:
    st.success("✅ **Excelente!** Não foram encontrados alunos em múltiplas turmas. Os dados estão consistentes.")
    
    # Mostrar estatística geral
    col_stats1, col_stats2 = st.columns(2)
    with col_stats1:
        total_alunos_unicos = df_filt["Aluno"].nunique()
        st.metric("Total de Alunos Únicos", total_alunos_unicos, help="Número total de alunos únicos nos dados filtrados")
    
    with col_stats2:
        total_turmas = df_filt["Turma"].nunique()
        st.metric("Total de Turmas", total_turmas, help="Número total de turmas nos dados filtrados")

# Assinatura discreta do criador
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; margin-top: 40px; padding: 20px;">
        <p style="margin: 0;">
            Desenvolvido por <strong style="color: #1e40af;">Alexandre Tolentino</strong> • 
            <em>Painel SGE - Sistema de Gestão Escolar</em>
        </p>
    </div>
    """, 
    unsafe_allow_html=True
)
