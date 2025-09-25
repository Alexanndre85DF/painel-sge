"""
Página Admin - Monitoramento de Acessos
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from io import BytesIO
import json
from firebase_config import firebase_manager
from ip_utils import get_client_info

def tela_admin():
    """Tela de login para administradores"""
    st.markdown("""
    <div style="text-align: center; padding: 40px 20px; background: linear-gradient(135deg, #dc2626, #ef4444); border-radius: 15px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0; font-size: 2.5em; font-weight: 700;">🔐 Painel Administrativo</h1>
        <h2 style="color: white; margin: 15px 0 0 0; font-weight: 600;">Monitoramento de Acessos</h2>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### Acesso Administrativo")
        st.warning("⚠️ Esta área é restrita apenas para administradores")
        
        with st.form("admin_login_form"):
            admin_user = st.text_input("Usuário Admin:", placeholder="admin")
            admin_password = st.text_input("Senha Admin:", type="password", placeholder="Digite a senha administrativa")
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                login_btn = st.form_submit_button("Entrar como Admin", use_container_width=True, type="primary")
            with col_btn2:
                if st.form_submit_button("Voltar", use_container_width=True):
                    st.session_state.admin_logado = False
                    st.session_state.mostrar_admin = False
                    st.rerun()
        
        if login_btn:
            # Verificação simples de admin (você pode melhorar isso)
            if admin_user == "admin" and admin_password == "admin123":
                st.session_state.admin_logado = True
                st.success("Login administrativo realizado com sucesso!")
                st.rerun()
            else:
                st.error("Usuário ou senha administrativa incorretos!")

def dashboard_admin():
    """Dashboard principal do administrador"""
    st.markdown("""
    <div style="text-align: center; padding: 30px 20px; background: linear-gradient(135deg, #dc2626, #ef4444); border-radius: 15px; margin-bottom: 30px;">
        <h1 style="color: white; margin: 0; font-size: 2.2em; font-weight: 700;">📊 Dashboard Administrativo</h1>
        <h2 style="color: white; margin: 10px 0 0 0; font-weight: 600;">Monitoramento de Acessos em Tempo Real</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Botões de controle
    col_control1, col_control2, col_control3, col_control4 = st.columns(4)
    
    with col_control1:
        if st.button("🔄 Atualizar Dados", use_container_width=True, type="primary"):
            st.rerun()
    
    with col_control2:
        if st.button("📊 Relatório Completo", use_container_width=True):
            st.session_state.mostrar_relatorio = True
            st.rerun()
    
    with col_control3:
        if st.button("👥 Estatísticas por Usuário", use_container_width=True):
            st.session_state.mostrar_stats_usuario = True
            st.rerun()
    
    with col_control4:
        if st.button("🚪 Sair do Admin", use_container_width=True):
            st.session_state.admin_logado = False
            st.session_state.mostrar_admin = False
            st.rerun()
    
    st.markdown("---")
    
    try:
        # Carregar dados do Firebase
        with st.spinner("Carregando dados de monitoramento..."):
            logs = firebase_manager.get_access_logs(limit=500)
        
        if not logs:
            st.warning("Nenhum log de acesso encontrado ainda.")
            return
        
        # Converter para DataFrame
        df_logs = pd.DataFrame(logs)
        df_logs['timestamp'] = pd.to_datetime(df_logs['timestamp'])
        df_logs['data'] = df_logs['timestamp'].dt.date
        df_logs['hora'] = df_logs['timestamp'].dt.strftime('%H:%M')
        
        # Métricas principais
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_acessos = len(df_logs)
            st.metric("Total de Acessos", total_acessos)
        
        with col2:
            usuarios_unicos = df_logs['usuario'].nunique()
            st.metric("Usuários Únicos", usuarios_unicos)
        
        with col3:
            ips_unicos = df_logs['ip'].nunique()
            st.metric("IPs Únicos", ips_unicos)
        
        with col4:
            hoje = datetime.now().date()
            acessos_hoje = len(df_logs[df_logs['data'] == hoje])
            st.metric("Acessos Hoje", acessos_hoje)
        
        st.markdown("---")
        
        # Filtros
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        with col_filter1:
            usuarios_disponiveis = ['Todos'] + sorted(df_logs['usuario'].unique().tolist())
            usuario_filtro = st.selectbox("Filtrar por Usuário:", usuarios_disponiveis)
        
        with col_filter2:
            datas_disponiveis = sorted(df_logs['data'].unique(), reverse=True)
            data_filtro = st.selectbox("Filtrar por Data:", ['Todas'] + [str(d) for d in datas_disponiveis])
        
        with col_filter3:
            ips_disponiveis = ['Todos'] + sorted(df_logs['ip'].unique().tolist())
            ip_filtro = st.selectbox("Filtrar por IP:", ips_disponiveis)
        
        # Aplicar filtros
        df_filtrado = df_logs.copy()
        
        if usuario_filtro != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['usuario'] == usuario_filtro]
        
        if data_filtro != 'Todas':
            data_selecionada = pd.to_datetime(data_filtro).date()
            df_filtrado = df_filtrado[df_filtrado['data'] == data_selecionada]
        
        if ip_filtro != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['ip'] == ip_filtro]
        
        # Gráficos
        col_graph1, col_graph2 = st.columns(2)
        
        with col_graph1:
            # Gráfico de acessos por dia
            acessos_por_dia = df_filtrado.groupby('data').size().reset_index(name='acessos')
            fig_dia = px.line(acessos_por_dia, x='data', y='acessos', 
                             title='Acessos por Dia', markers=True)
            fig_dia.update_layout(xaxis_title="Data", yaxis_title="Número de Acessos")
            st.plotly_chart(fig_dia, use_container_width=True)
        
        with col_graph2:
            # Gráfico de acessos por usuário
            acessos_por_usuario = df_filtrado.groupby('usuario').size().reset_index(name='acessos')
            fig_usuario = px.bar(acessos_por_usuario, x='usuario', y='acessos',
                                title='Acessos por Usuário')
            fig_usuario.update_layout(xaxis_title="Usuário", yaxis_title="Número de Acessos")
            fig_usuario.update_xaxis(tickangle=45)
            st.plotly_chart(fig_usuario, use_container_width=True)
        
        # Gráfico de acessos por hora
        df_filtrado['hora_int'] = pd.to_datetime(df_filtrado['hora'], format='%H:%M').dt.hour
        acessos_por_hora = df_filtrado.groupby('hora_int').size().reset_index(name='acessos')
        fig_hora = px.bar(acessos_por_hora, x='hora_int', y='acessos',
                         title='Acessos por Hora do Dia')
        fig_hora.update_layout(xaxis_title="Hora", yaxis_title="Número de Acessos")
        st.plotly_chart(fig_hora, use_container_width=True)
        
        st.markdown("---")
        
        # Tabela de logs recentes
        st.markdown("### 📋 Logs de Acesso Recentes")
        
        # Preparar dados para exibição
        df_exibicao = df_filtrado[['data_hora', 'usuario', 'ip', 'user_agent']].copy()
        df_exibicao.columns = ['Data/Hora', 'Usuário', 'IP', 'Navegador']
        df_exibicao = df_exibicao.sort_values('Data/Hora', ascending=False)
        
        st.dataframe(df_exibicao, use_container_width=True, height=400)
        
        # Botões de ação
        col_export, col_clean = st.columns(2)
        
        with col_export:
            if st.button("📥 Exportar Logs para Excel"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_exibicao.to_excel(writer, sheet_name='Logs de Acesso', index=False)
                
                st.download_button(
                    label="⬇️ Baixar Arquivo Excel",
                    data=output.getvalue(),
                    file_name=f"logs_acesso_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col_clean:
            if st.button("🧹 Limpar Logs Duplicados"):
                try:
                    # Limpar logs duplicados (manter apenas um por usuário a cada 2 minutos)
                    logs_limpos = []
                    for log in logs:
                        usuario = log.get('usuario', '')
                        timestamp = log.get('timestamp', '')
                        
                        # Verificar se já existe um log similar recente
                        log_similar = False
                        for log_existente in logs_limpos:
                            if (log_existente.get('usuario') == usuario and 
                                abs((datetime.fromisoformat(timestamp.replace('Z', '')) - 
                                     datetime.fromisoformat(log_existente.get('timestamp', '').replace('Z', ''))).seconds) < 120):
                                log_similar = True
                                break
                        
                        if not log_similar:
                            logs_limpos.append(log)
                    
                    # Salvar logs limpos
                    with open('local_access_log.json', 'w', encoding='utf-8') as f:
                        json.dump(logs_limpos, f, ensure_ascii=False, indent=2)
                    
                    st.success(f"Logs limpos! Removidos {len(logs) - len(logs_limpos)} duplicados.")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Erro ao limpar logs: {e}")
    
    except Exception as e:
        st.error(f"Erro ao carregar dados: {str(e)}")
        st.info("Verifique se o Firebase está configurado corretamente.")

def relatorio_completo():
    """Relatório completo de acessos"""
    st.markdown("### 📊 Relatório Completo de Acessos")
    
    try:
        logs = firebase_manager.get_access_logs(limit=1000)
        
        if not logs:
            st.warning("Nenhum log encontrado.")
            return
        
        df_logs = pd.DataFrame(logs)
        df_logs['timestamp'] = pd.to_datetime(df_logs['timestamp'])
        
        # Estatísticas gerais
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📈 Estatísticas Gerais")
            
            total_acessos = len(df_logs)
            usuarios_unicos = df_logs['usuario'].nunique()
            ips_unicos = df_logs['ip'].nunique()
            
            st.metric("Total de Acessos", total_acessos)
            st.metric("Usuários Únicos", usuarios_unicos)
            st.metric("IPs Únicos", ips_unicos)
            
            # Período de atividade
            primeiro_acesso = df_logs['timestamp'].min()
            ultimo_acesso = df_logs['timestamp'].max()
            
            st.info(f"**Período:** {primeiro_acesso.strftime('%d/%m/%Y')} até {ultimo_acesso.strftime('%d/%m/%Y')}")
        
        with col2:
            st.markdown("#### 🏆 Top Usuários")
            
            top_usuarios = df_logs.groupby('usuario').size().sort_values(ascending=False).head(10)
            
            for i, (usuario, acessos) in enumerate(top_usuarios.items(), 1):
                st.write(f"{i}. **{usuario}**: {acessos} acessos")
        
        # Gráfico de evolução temporal
        st.markdown("#### 📈 Evolução Temporal")
        
        df_logs['data'] = df_logs['timestamp'].dt.date
        evolucao = df_logs.groupby('data').agg({
            'usuario': 'count',
            'ip': 'nunique'
        }).rename(columns={'usuario': 'total_acessos', 'ip': 'ips_unicos'})
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=evolucao.index, y=evolucao['total_acessos'], 
                                mode='lines+markers', name='Total de Acessos'))
        fig.add_trace(go.Scatter(x=evolucao.index, y=evolucao['ips_unicos'], 
                                mode='lines+markers', name='IPs Únicos'))
        
        fig.update_layout(title='Evolução dos Acessos', xaxis_title='Data', yaxis_title='Quantidade')
        st.plotly_chart(fig, use_container_width=True)
        
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {str(e)}")

def estatisticas_usuario():
    """Estatísticas detalhadas por usuário"""
    st.markdown("### 👥 Estatísticas por Usuário")
    
    try:
        logs = firebase_manager.get_access_logs(limit=1000)
        
        if not logs:
            st.warning("Nenhum log encontrado.")
            return
        
        df_logs = pd.DataFrame(logs)
        
        # Selecionar usuário
        usuarios = sorted(df_logs['usuario'].unique())
        usuario_selecionado = st.selectbox("Selecionar usuário:", usuarios)
        
        if usuario_selecionado:
            # Estatísticas do usuário
            stats = firebase_manager.get_user_access_stats(usuario_selecionado)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total de Acessos", stats['total_acessos'])
            
            with col2:
                if stats['ultimo_acesso']:
                    ultimo_acesso = pd.to_datetime(stats['ultimo_acesso'])
                    st.metric("Último Acesso", ultimo_acesso.strftime('%d/%m/%Y %H:%M'))
                else:
                    st.metric("Último Acesso", "N/A")
            
            with col3:
                st.metric("IPs Utilizados", len(stats['ips_utilizados']))
            
            # IPs utilizados
            st.markdown("#### 🌐 IPs Utilizados")
            for ip in stats['ips_utilizados']:
                st.write(f"• {ip}")
            
            # Histórico do usuário
            st.markdown("#### 📋 Histórico de Acessos")
            
            df_usuario = df_logs[df_logs['usuario'] == usuario_selecionado].copy()
            df_usuario = df_usuario.sort_values('timestamp', ascending=False)
            
            df_exibicao = df_usuario[['data_hora', 'ip', 'user_agent']].copy()
            df_exibicao.columns = ['Data/Hora', 'IP', 'Navegador']
            
            st.dataframe(df_exibicao, use_container_width=True)
    
    except Exception as e:
        st.error(f"Erro ao carregar estatísticas: {str(e)}")
