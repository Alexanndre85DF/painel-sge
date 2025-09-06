import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

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
@st.cache_data(show_spinner=False)
def carregar_dados(arquivo, sheet=None):
    if arquivo is None:
        # Tenta ler o padrão local "dados.xlsx"
        df = pd.read_excel("dados.xlsx", sheet_name=sheet) if sheet else pd.read_excel("dados.xlsx")
    else:
        df = pd.read_excel(arquivo, sheet_name=sheet) if sheet else pd.read_excel(arquivo)

    # Normalizar nomes de colunas
    df.columns = [c.strip() for c in df.columns]

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

    # Padronizar texto dos campos principais (evita diferenças por espaços)
    for col in ["Escola", "Turma", "Turno", "Aluno", "Status", "Periodo", "Disciplina"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df

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
<div style="text-align: center; padding: 40px 20px; background: #ffffff; border: 2px solid #1e3a8a; border-radius: 8px; margin-bottom: 30px;">
    <h1 style="color: #1e3a8a; margin: 0; font-size: 2.2em; font-weight: 700;">Superintendência Regional de Educação de Gurupi TO</h1>
    <h2 style="color: #1e40af; margin: 15px 0 0 0; font-weight: 600; font-size: 1.8em;">Painel SGE</h2>
    <h3 style="color: #1e40af; margin: 10px 0 0 0; font-weight: 400; font-size: 1.4em;">Notas, Frequência, Riscos e Alertas</h3>
    <p style="color: #64748b; margin: 10px 0 0 0; font-size: 1.1em;">Análise dos 1º e 2º Bimestres</p>
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
<div style="background: #f8fafc; border: 2px solid #1e3a8a; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h3 style="color: #1e3a8a; text-align: center; margin: 0 0 20px 0; font-size: 1.4em; font-weight: 600;">📊 Visão Geral dos Dados</h3>
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
<div style="background: #f0f9ff; border: 2px solid #0ea5e9; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h3 style="color: #0c4a6e; text-align: center; margin: 0 0 20px 0; font-size: 1.4em; font-weight: 600;">👥 Total de Estudantes</h3>
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
<div style="background: #f8fafc; border: 1px solid #1e3a8a; padding: 20px; border-radius: 6px; margin-bottom: 20px;">
    <h2 style="color: #1e3a8a; text-align: center; margin: 0; font-size: 1.4em; font-weight: 600;">Filtros</h2>
    <p style="color: #64748b; text-align: center; margin: 8px 0 0 0; font-size: 0.9em;">Filtre os dados para análise específica</p>
</div>
""", unsafe_allow_html=True)

escolas = sorted(df["Escola"].dropna().unique().tolist()) if "Escola" in df.columns else []
status_opcoes = sorted(df["Status"].dropna().unique().tolist()) if "Status" in df.columns else []

st.sidebar.markdown("### Escola")
escola_sel = st.sidebar.selectbox("Selecione a escola:", ["Todas"] + escolas, help="Filtre por escola específica")

st.sidebar.markdown("### Status")
status_sel = st.sidebar.selectbox("Selecione o status:", ["Todos"] + status_opcoes, help="Filtre por status do aluno")

# Filtrar dados baseado na escola e status selecionados para mostrar opções relevantes
df_temp = df.copy()
if escola_sel != "Todas":
    df_temp = df_temp[df_temp["Escola"] == escola_sel]
if status_sel != "Todos":
    df_temp = df_temp[df_temp["Status"] == status_sel]

turmas = sorted(df_temp["Turma"].dropna().unique().tolist()) if "Turma" in df_temp.columns else []
disciplinas = sorted(df_temp["Disciplina"].dropna().unique().tolist()) if "Disciplina" in df_temp.columns else []
alunos = sorted(df_temp["Aluno"].dropna().unique().tolist()) if "Aluno" in df_temp.columns else []

# Filtros com interface melhorada
st.sidebar.markdown("### Turmas")
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

st.sidebar.markdown("### Disciplinas")
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

st.sidebar.markdown("### Aluno")
aluno_sel = st.sidebar.selectbox("Selecione o aluno:", ["Todos"] + alunos, help="Filtre por aluno específico")

df_filt = df.copy()
if escola_sel != "Todas":
    df_filt = df_filt[df_filt["Escola"] == escola_sel]
if status_sel != "Todos":
    df_filt = df_filt[df_filt["Status"] == status_sel]
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
<div style="background: #fef3c7; border: 2px solid #f59e0b; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h3 style="color: #92400e; text-align: center; margin: 0 0 20px 0; font-size: 1.4em; font-weight: 600;">🔍 Total de Estudantes (Filtrado)</h3>
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
    <div style="background: #f0fdf4; border: 2px solid #22c55e; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="color: #15803d; text-align: center; margin: 0 0 20px 0; font-size: 1.4em; font-weight: 600;">📈 Resumo de Frequência</h3>
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
    if "Frequencia Anual" in df_filt.columns:
        freq_geral = df_filt.groupby(["Aluno"])["Frequencia Anual"].last().reset_index()
        freq_geral = freq_geral.rename(columns={"Frequencia Anual": "Frequencia"})
    else:
        freq_geral = df_filt.groupby(["Aluno"])["Frequencia"].last().reset_index()
    
    freq_geral["Classificacao_Freq"] = freq_geral["Frequencia"].apply(classificar_frequencia_geral)
    contagem_freq_geral = freq_geral["Classificacao_Freq"].value_counts()
    
    # Calcular total de alunos para porcentagem
    total_alunos_freq = contagem_freq_geral.sum()
    
    with colF1:
        valor_reprovado = contagem_freq_geral.get("Reprovado", 0)
        percent_reprovado = (valor_reprovado / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.metric(
            label="< 75% (Reprovado)", 
            value=f"{valor_reprovado} ({percent_reprovado:.1f}%)",
            help="Alunos reprovados por frequência"
        )
    with colF2:
        valor_alto_risco = contagem_freq_geral.get("Alto Risco", 0)
        percent_alto_risco = (valor_alto_risco / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.metric(
            label="< 80% (Alto Risco)", 
            value=f"{valor_alto_risco} ({percent_alto_risco:.1f}%)",
            help="Alunos em alto risco de reprovação"
        )
    with colF3:
        valor_risco_moderado = contagem_freq_geral.get("Risco Moderado", 0)
        percent_risco_moderado = (valor_risco_moderado / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.metric(
            label="< 90% (Risco Moderado)", 
            value=f"{valor_risco_moderado} ({percent_risco_moderado:.1f}%)",
            help="Alunos com risco moderado"
        )
    with colF4:
        valor_ponto_atencao = contagem_freq_geral.get("Ponto de Atenção", 0)
        percent_ponto_atencao = (valor_ponto_atencao / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.metric(
            label="< 95% (Ponto Atenção)", 
            value=f"{valor_ponto_atencao} ({percent_ponto_atencao:.1f}%)",
            help="Alunos que precisam de atenção"
        )
    with colF5:
        valor_meta_favoravel = contagem_freq_geral.get("Meta Favorável", 0)
        percent_meta_favoravel = (valor_meta_favoravel / total_alunos_freq * 100) if total_alunos_freq > 0 else 0
        st.metric(
            label="≥ 95% (Meta Favorável)", 
            value=f"{valor_meta_favoravel} ({percent_meta_favoravel:.1f}%)",
            help="Alunos com frequência dentro da meta"
        )

# -----------------------------
# Indicadores e tabelas de risco
# -----------------------------
indic = calcula_indicadores(df_filt)

# KPIs - Análise de Notas Baixas
st.markdown("""
<div style="background: #fef2f2; border: 2px solid #dc2626; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h3 style="color: #991b1b; text-align: center; margin: 0 0 20px 0; font-size: 1.4em; font-weight: 600;">📚 Análise de Notas Abaixo da Média</h3>
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
    st.metric(
        label="Notas < 6 – 1º Bim", 
        value=f"{len(notas_baixas_b1)} ({percent_notas_b1:.1f}%)",
        help="Total de registros com notas abaixo de 6 no 1º bimestre"
    )
with col2:
    percent_notas_b2 = (len(notas_baixas_b2) / len(df_filt) * 100) if len(df_filt) > 0 else 0
    st.metric(
        label="Notas < 6 – 2º Bim", 
        value=f"{len(notas_baixas_b2)} ({percent_notas_b2:.1f}%)",
        help="Total de registros com notas abaixo de 6 no 2º bimestre"
    )
with col3:
    percent_alunos_b1 = (alunos_notas_baixas_b1 / total_estudantes_para_percent * 100) if total_estudantes_para_percent > 0 else 0
    st.metric(
        label="Alunos < 6 – 1º Bim", 
        value=f"{alunos_notas_baixas_b1} ({percent_alunos_b1:.1f}%)",
        help="Número de alunos únicos com notas abaixo de 6 no 1º bimestre"
    )
with col4:
    percent_alunos_b2 = (alunos_notas_baixas_b2 / total_estudantes_para_percent * 100) if total_estudantes_para_percent > 0 else 0
    st.metric(
        label="Alunos < 6 – 2º Bim", 
        value=f"{alunos_notas_baixas_b2} ({percent_alunos_b2:.1f}%)",
        help="Número de alunos únicos com notas abaixo de 6 no 2º bimestre"
    )

# KPIs - Alertas Críticos (com destaque visual)
st.markdown("""
<div style="background: #fef2f2; border: 1px solid #dc2626; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h2 style="color: #dc2626; text-align: center; margin: 0; font-size: 1.6em; font-weight: 600;">Alertas Críticos</h2>
    <p style="color: #ef4444; text-align: center; margin: 8px 0 0 0; font-size: 1em;">Situações que precisam de atenção imediata</p>
</div>
""", unsafe_allow_html=True)

col5, col6 = st.columns(2)

# Métricas de alerta com destaque visual
alerta_count = int(indic["Alerta"].sum())
corda_bamba_count = int(indic["CordaBamba"].sum())

with col5:
    st.markdown("""
    <div style="background: #fef2f2; border-left: 4px solid #dc2626; padding: 15px; border-radius: 4px; margin: 10px 0;">
        <h3 style="color: #dc2626; margin: 0 0 10px 0; font-size: 1.1em;">Alunos-Disciplinas em ALERTA</h3>
    </div>
    """, unsafe_allow_html=True)
    st.metric(
        label="", 
        value=alerta_count,
        help="Alunos em situação de risco (Vermelho Duplo, Queda p/ Vermelho ou Corda Bamba)"
    )

with col6:
    st.markdown("""
    <div style="background: #fffbeb; border-left: 4px solid #d97706; padding: 15px; border-radius: 4px; margin: 10px 0;">
        <h3 style="color: #d97706; margin: 0 0 10px 0; font-size: 1.1em;">Corda Bamba</h3>
    </div>
    """, unsafe_allow_html=True)
    st.metric(
        label="", 
        value=corda_bamba_count,
        help="Alunos que precisam de média ≥ 7 nos próximos 2 bimestres para não reprovar"
    )

# Resumo Executivo - Dashboard Principal
st.markdown("""
<div style="background: #f0f9ff; border: 1px solid #0ea5e9; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h2 style="color: #0c4a6e; text-align: center; margin: 0; font-size: 1.6em; font-weight: 600;">📊 Resumo Executivo</h2>
    <p style="color: #0369a1; text-align: center; margin: 8px 0 0 0; font-size: 1em;">Visão consolidada dos principais indicadores</p>
</div>
""", unsafe_allow_html=True)

# Métricas consolidadas em cards
col_res1, col_res2, col_res3, col_res4 = st.columns(4)

with col_res1:
    st.markdown("""
    <div style="background: #fef2f2; border-left: 4px solid #dc2626; padding: 15px; border-radius: 4px; margin: 10px 0;">
        <h3 style="color: #dc2626; margin: 0 0 5px 0; font-size: 1em;">🚨 Alertas Críticos</h3>
        <p style="color: #991b1b; margin: 0; font-size: 0.9em;">Situações que precisam de atenção imediata</p>
    </div>
    """, unsafe_allow_html=True)
    st.metric("Alunos em Alerta", alerta_count, help="Total de alunos-disciplinas em situação de risco")

with col_res2:
    st.markdown("""
    <div style="background: #fffbeb; border-left: 4px solid #d97706; padding: 15px; border-radius: 4px; margin: 10px 0;">
        <h3 style="color: #d97706; margin: 0 0 5px 0; font-size: 1em;">⚖️ Corda Bamba</h3>
        <p style="color: #92400e; margin: 0; font-size: 0.9em;">Precisam de média ≥ 7 nos próximos bimestres</p>
    </div>
    """, unsafe_allow_html=True)
    st.metric("Alunos em Corda Bamba", corda_bamba_count, help="Alunos que precisam de média ≥ 7 nos próximos 2 bimestres")

with col_res3:
    # Calcular total de alunos com notas baixas
    total_alunos_notas_baixas = max(alunos_notas_baixas_b1, alunos_notas_baixas_b2)
    st.markdown("""
    <div style="background: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; border-radius: 4px; margin: 10px 0;">
        <h3 style="color: #d97706; margin: 0 0 5px 0; font-size: 1em;">📉 Notas Baixas</h3>
        <p style="color: #92400e; margin: 0; font-size: 0.9em;">Alunos com notas abaixo de 6</p>
    </div>
    """, unsafe_allow_html=True)
    st.metric("Alunos com Notas < 6", total_alunos_notas_baixas, help="Máximo entre 1º e 2º bimestres")

with col_res4:
    # Calcular alunos com frequência baixa se disponível
    if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
        if "Frequencia Anual" in df_filt.columns:
            freq_baixa_count = len(df_filt[df_filt["Frequencia Anual"] < 95]["Aluno"].unique())
        else:
            freq_baixa_count = len(df_filt[df_filt["Frequencia"] < 95]["Aluno"].unique())
        
        st.markdown("""
        <div style="background: #f3e8ff; border-left: 4px solid #8b5cf6; padding: 15px; border-radius: 4px; margin: 10px 0;">
            <h3 style="color: #7c3aed; margin: 0 0 5px 0; font-size: 1em;">📊 Frequência Baixa</h3>
            <p style="color: #6b21a8; margin: 0; font-size: 0.9em;">Alunos com frequência < 95%</p>
        </div>
        """, unsafe_allow_html=True)
        st.metric("Alunos com Freq < 95%", freq_baixa_count, help="Alunos com frequência abaixo da meta")
    else:
        st.markdown("""
        <div style="background: #f3f4f6; border-left: 4px solid #6b7280; padding: 15px; border-radius: 4px; margin: 10px 0;">
            <h3 style="color: #374151; margin: 0 0 5px 0; font-size: 1em;">📊 Frequência</h3>
            <p style="color: #4b5563; margin: 0; font-size: 0.9em;">Dados não disponíveis</p>
        </div>
        """, unsafe_allow_html=True)
        st.metric("Dados de Frequência", "N/A", help="Dados de frequência não disponíveis")

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
<div style="background: #eff6ff; border: 1px solid #3b82f6; padding: 20px; border-radius: 8px; margin: 20px 0;">
    <h2 style="color: #1e40af; text-align: center; margin: 0; font-size: 1.6em; font-weight: 600;">{freq_title}</h2>
    <p style="color: #3b82f6; text-align: center; margin: 8px 0 0 0; font-size: 1em;">{freq_subtitle}</p>
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
    freq_atual = df_filt.groupby(["Aluno"])["Frequencia Anual"].last().reset_index()
    freq_atual = freq_atual.rename(columns={"Frequencia Anual": "Frequencia"})
    freq_atual["Classificacao_Freq"] = freq_atual["Frequencia"].apply(classificar_frequencia)
elif "Frequencia" in df_filt.columns:
    # Usar frequência do período se anual não estiver disponível
    freq_atual = df_filt.groupby(["Aluno"])["Frequencia"].last().reset_index()
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
        # Tabela de frequência por aluno
        if "Frequencia Anual" in df_filt.columns:
            freq_detalhada = df_filt.groupby(["Aluno"])["Frequencia Anual"].last().reset_index()
            freq_detalhada = freq_detalhada.rename(columns={"Frequencia Anual": "Frequencia"})
        else:
            freq_detalhada = df_filt.groupby(["Aluno"])["Frequencia"].last().reset_index()
        freq_detalhada["Classificacao_Freq"] = freq_detalhada["Frequencia"].apply(classificar_frequencia)
        freq_detalhada = freq_detalhada.sort_values(["Turma", "Aluno"])
        
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
        styled_freq = freq_detalhada[["Aluno", "Turma", "Frequencia_Formatada", "Classificacao_Freq"]]\
            .style.applymap(color_frequencia, subset=["Classificacao_Freq"])
        
        st.dataframe(styled_freq, use_container_width=True)
        
        # Legenda de frequência
        st.markdown("###  Legenda de Frequência")
        col_leg1, col_leg2, col_leg3 = st.columns(3)
        with col_leg1:
            st.markdown("""
            **📉 < 75%**: Reprovado por frequência  
            **🔴 < 80%**: Alto risco de reprovação
            """)
        with col_leg2:
            st.markdown("""
            **🟡 < 90%**: Risco moderado  
            **🟠 < 95%**: Ponto de atenção
            """)
        with col_leg3:
            st.markdown("""
            **🟢 ≥ 95%**: Meta favorável  
            **⚪ Sem dados**: Frequência não informada
            """)
    else:
        st.info("Dados de frequência não disponíveis na planilha.")

# Seção expandível: Análise Cruzada Nota x Frequência
with st.expander("Análise Cruzada: Notas x Frequência"):
    if ("Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns) and len(indic) > 0:
        # Combinar dados de notas e frequência (priorizando Frequencia Anual)
        if "Frequencia Anual" in df_filt.columns:
            freq_alunos = df_filt.groupby(["Aluno"])["Frequencia Anual"].last().reset_index()
            freq_alunos = freq_alunos.rename(columns={"Frequencia Anual": "Frequencia"})
        else:
            freq_alunos = df_filt.groupby(["Aluno"])["Frequencia"].last().reset_index()
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
                st.markdown("### Alunos com Frequência Abaixo de 95%")
                # Mostrar apenas colunas relevantes para frequência baixa
                freq_baixa_display = freq_baixa[["Aluno", "Turma", "Disciplina", "Classificacao", "Classificacao_Freq", "Frequencia"]].copy()
                # Formatar frequência
                freq_baixa_display["Frequencia"] = freq_baixa_display["Frequencia"].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A"
                )
                st.dataframe(freq_baixa_display, use_container_width=True)
            else:
                st.info("Todos os alunos têm frequência ≥ 95% (Meta Favorável).")
        else:
            st.info("Dados insuficientes para análise cruzada.")
    else:
        st.info("Dados de frequência ou notas não disponíveis para análise cruzada.")

st.markdown("---")

# Gráficos: Notas e Frequência por Disciplina
col_graf1, col_graf2 = st.columns(2)

# Gráfico: Notas abaixo de 6 por Disciplina (1º e 2º bimestres)
with col_graf1:
    with st.expander("📊 Notas Abaixo da Média por Disciplina"):
        base_baixas = pd.concat([notas_baixas_b1, notas_baixas_b2], ignore_index=True)
        if len(base_baixas) > 0:
            # Contar notas por disciplina
            contagem = base_baixas.groupby("Disciplina")["Nota"].count().reset_index()
            contagem = contagem.rename(columns={"Nota": "Qtd Notas < 6"})
            
            # Ordenar em ordem decrescente (maior para menor)
            contagem = contagem.sort_values("Qtd Notas < 6", ascending=False).reset_index(drop=True)
            
            # Adicionar coluna de cores intercaladas baseada na posição após ordenação
            contagem['Cor'] = ['#1e40af' if i % 2 == 0 else '#059669' for i in range(len(contagem))]
            
            # Debug: mostrar a ordenação
            st.write("**Debug - Ordenação das disciplinas:**")
            st.write(contagem[['Disciplina', 'Qtd Notas < 6', 'Cor']])
            
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
        else:
            st.info("Sem notas abaixo da média para os filtros atuais.")

# Gráfico: Distribuição de Frequência por Faixas
with col_graf2:
    with st.expander("📈 Distribuição de Frequência por Faixas"):
        if "Frequencia Anual" in df_filt.columns or "Frequencia" in df_filt.columns:
            # Usar os mesmos dados do Resumo de Frequência
            if "Frequencia Anual" in df_filt.columns:
                freq_geral = df_filt.groupby(["Aluno"])["Frequencia Anual"].last().reset_index()
                freq_geral = freq_geral.rename(columns={"Frequencia Anual": "Frequencia"})
            else:
                freq_geral = df_filt.groupby(["Aluno"])["Frequencia"].last().reset_index()
            
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
                
                # Estatísticas adicionais
                st.markdown("**📊 Resumo das Faixas de Frequência:**")
                col_stat1, col_stat2, col_stat3 = st.columns(3)
                with col_stat1:
                    total_alunos = contagem_freq_geral.sum()
                    st.metric("Total de Alunos", total_alunos)
                with col_stat2:
                    alunos_risco = contagem_freq_geral.get("Reprovado", 0) + contagem_freq_geral.get("Alto Risco", 0)
                    st.metric("Alunos em Risco", alunos_risco)
                with col_stat3:
                    alunos_meta = contagem_freq_geral.get("Meta Favorável", 0)
                    percentual_meta = (alunos_meta / total_alunos * 100) if total_alunos > 0 else 0
                    st.metric("Meta Favorável", f"{percentual_meta:.1f}%")
            else:
                st.info("Sem dados de frequência para exibir.")
        else:
            st.info("Dados de frequência não disponíveis na planilha.")

# Tabela: Alunos-Disciplinas em ALERTA (com cálculo de necessidade para 3º e 4º)
st.subheader("Alunos/Disciplinas em ALERTA")
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
else:
    st.dataframe(pd.DataFrame(columns=cols_visiveis), use_container_width=True)

# Tabela: Quedas e Recuperações (todos para diagnóstico rápido)
st.subheader("Quedas e Recuperações (Panorama B1→B2)")
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
