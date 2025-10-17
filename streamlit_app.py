import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="Monitor AI - Carglass",
    page_icon="🔴",
    layout="wide",
    initial_sidebar_state="expanded"
)

CARGLASS_RED = "#DC0A0A"
CARGLASS_DARK_RED = "#B00000"
CARGLASS_BLUE = "#4A90E2"
CARGLASS_DARK_BLUE = "#2C5AA0"
CARGLASS_GRAY = "#6C757D"
CARGLASS_LIGHT_GRAY = "#F8F9FA"
CARGLASS_GREEN = "#28A745"
CARGLASS_YELLOW = "#FFC107"
CARGLASS_ORANGE = "#FD7E14"

custom_css = """
<style>
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
    .stMetric {
        background: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        border-left: 4px solid """ + CARGLASS_RED + """;
    }
    
    .stMetric [data-testid="metric-container"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
    }
    
    div[data-testid="metric-container"] > label[data-testid="stMetricLabel"] {
        color: white !important;
        font-weight: 600;
        font-size: 14px;
    }
    
    div[data-testid="metric-container"] > div[data-testid="stMetricValue"] {
        color: white !important;
        font-weight: bold;
        font-size: 28px;
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
    }
    
    h1 {
        color: """ + CARGLASS_RED + """;
        font-weight: bold;
        text-align: center;
        padding: 20px;
        background: white;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin-bottom: 30px;
    }
    
    h2 {
        color: """ + CARGLASS_DARK_RED + """;
        border-bottom: 3px solid """ + CARGLASS_RED + """;
        padding-bottom: 10px;
        margin-top: 30px;
    }
    
    .stSelectbox label {
        color: """ + CARGLASS_DARK_RED + """ !important;
        font-weight: 600;
    }
    
    .stMultiSelect label {
        color: """ + CARGLASS_DARK_RED + """ !important;
        font-weight: 600;
    }
    
    .stDateInput label {
        color: """ + CARGLASS_DARK_RED + """ !important;
        font-weight: 600;
    }
    
    .card {
        background: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    
    .header-container {
        background: linear-gradient(135deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
        color: white;
        padding: 30px;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 30px;
    }
    
    .kpi-card {
        background: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.1);
        transition: transform 0.3s;
        border-top: 4px solid """ + CARGLASS_RED + """;
    }
    
    .kpi-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 25px rgba(0,0,0,0.15);
    }
</style>
"""

st.markdown(custom_css, unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    try:
        xls = pd.ExcelFile(file)
        if 'Consulta1' in xls.sheet_names:
            df = pd.read_excel(file, sheet_name='Consulta1')
            df['AnalysisDateTime'] = pd.to_datetime(df['AnalysisDateTime'])
            df['CallDate'] = pd.to_datetime(df['CallDate'])
            return df
        else:
            st.error("A planilha 'Consulta1' não foi encontrada no arquivo.")
            return None
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return None

def create_gauge_chart(value, title, color):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = value,
        title = {'text': title, 'font': {'size': 16, 'color': CARGLASS_DARK_RED}},
        delta = {'reference': 70, 'increasing': {'color': CARGLASS_GREEN}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': CARGLASS_GRAY},
            'bar': {'color': color},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': CARGLASS_GRAY,
            'steps': [
                {'range': [0, 50], 'color': '#FFE5E5'},
                {'range': [50, 70], 'color': '#FFF5E5'},
                {'range': [70, 100], 'color': '#E5FFE5'}
            ],
            'threshold': {
                'line': {'color': CARGLASS_DARK_RED, 'width': 4},
                'thickness': 0.75,
                'value': 70
            }
        }
    ))
    
    fig.update_layout(
        height=250,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor='white',
        font={'color': CARGLASS_DARK_RED}
    )
    
    return fig

def create_performance_chart(df):
    questions = [f'Question{i}' for i in range(1, 13)]
    question_labels = [
        'Q1: Saudação', 'Q2: Dados', 'Q3: LGPD', 'Q4: Eco',
        'Q5: Escuta', 'Q6: Produto', 'Q7: Confirmação', 'Q8: Loja',
        'Q9: Comunicação', 'Q10: Conduta', 'Q11: Encerramento', 'Q12: Pesquisa'
    ]
    
    performance = []
    for q in questions:
        if q in df.columns:
            performance.append(df[q].mean() * 100)
        else:
            performance.append(0)
    
    colors = [CARGLASS_GREEN if p >= 70 else CARGLASS_ORANGE if p >= 50 else CARGLASS_RED for p in performance]
    
    fig = go.Figure(data=[
        go.Bar(
            x=question_labels,
            y=performance,
            marker_color=colors,
            text=[f'{p:.1f}%' for p in performance],
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Performance: %{y:.1f}%<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title={
            'text': '✅ Performance por Critério de Avaliação',
            'font': {'size': 20, 'color': CARGLASS_DARK_RED}
        },
        xaxis_title='Critérios',
        yaxis_title='Performance (%)',
        yaxis={'range': [0, 110]},
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=400,
        showlegend=False,
        font={'color': CARGLASS_DARK_RED}
    )
    
    fig.add_hline(y=70, line_dash="dash", line_color=CARGLASS_GREEN,
                  annotation_text="Meta: 70%", annotation_position="right")
    
    return fig

def create_satisfaction_donut(df):
    satisfaction_map = {
        'BOA': 'Boa',
        'MODERADA': 'Moderada',
        'BAIXA': 'Baixa',
        'ALTA': 'Alta'
    }
    
    if 'Client' in df.columns:
        satisfaction_counts = df['Client'].value_counts()
        labels = [satisfaction_map.get(label, label) for label in satisfaction_counts.index]
        values = satisfaction_counts.values
        
        colors = {
            'Boa': CARGLASS_GREEN,
            'Alta': CARGLASS_GREEN,
            'Moderada': CARGLASS_YELLOW,
            'Baixa': CARGLASS_RED
        }
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=.5,
            marker=dict(colors=[colors.get(label, CARGLASS_GRAY) for label in labels]),
            textinfo='label+percent',
            textposition='outside',
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title={
                'text': '😊 Distribuição de Satisfação do Cliente',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED}
            },
            height=400,
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED},
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.05
            )
        )
        
        return fig
    return None

def create_risk_analysis(df):
    if 'ClientRisk' in df.columns:
        risk_counts = df['ClientRisk'].value_counts()
        
        fig = go.Figure()
        
        colors = {
            'BAIXO': CARGLASS_GREEN,
            'MEDIO': CARGLASS_YELLOW,
            'ALTO': CARGLASS_RED
        }
        
        for risk in ['BAIXO', 'MEDIO', 'ALTO']:
            if risk in risk_counts.index:
                fig.add_trace(go.Bar(
                    x=[risk_counts[risk]],
                    y=[risk],
                    orientation='h',
                    name=risk.capitalize(),
                    marker_color=colors.get(risk, CARGLASS_GRAY),
                    text=f'{risk_counts[risk]} ({risk_counts[risk]/len(df)*100:.1f}%)',
                    textposition='outside',
                    hovertemplate=f'<b>{risk.capitalize()}</b><br>Casos: %{{x}}<br>Percentual: {risk_counts[risk]/len(df)*100:.1f}%<extra></extra>'
                ))
        
        fig.update_layout(
            title={
                'text': '⚠️ Análise de Risco',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED}
            },
            xaxis_title='Número de Casos',
            yaxis_title='Nível de Risco',
            height=300,
            showlegend=False,
            plot_bgcolor='white',
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED}
        )
        
        return fig
    return None

def create_agent_ranking(df, top_n=5):
    if 'CustomerAgent' in df.columns and 'NOTAS' in df.columns:
        agent_scores = df.groupby('CustomerAgent').agg({
            'NOTAS': 'mean',
            'IdAnalysis': 'count'
        }).round(1)
        
        agent_scores.columns = ['Score Médio', 'Total Ligações']
        agent_scores = agent_scores.sort_values('Score Médio', ascending=False).head(top_n)
        
        fig = go.Figure()
        
        for idx, (agent, row) in enumerate(agent_scores.iterrows()):
            color = CARGLASS_GREEN if idx < 3 else CARGLASS_BLUE
            fig.add_trace(go.Bar(
                x=[row['Score Médio']],
                y=[agent],
                orientation='h',
                name=agent,
                marker_color=color,
                text=f"{row['Score Médio']:.1f} pts<br>{int(row['Total Ligações'])} ligações",
                textposition='outside',
                hovertemplate=f"<b>{agent}</b><br>Score: {row['Score Médio']:.1f}<br>Ligações: {int(row['Total Ligações'])}<extra></extra>"
            ))
        
        fig.update_layout(
            title={
                'text': f'👥 Top {top_n} Agentes por Score',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED}
            },
            xaxis_title='Score Médio',
            yaxis_title='',
            height=400,
            showlegend=False,
            plot_bgcolor='white',
            paper_bgcolor='white',
            xaxis={'range': [0, 100]},
            font={'color': CARGLASS_DARK_RED}
        )
        
        return fig
    return None

def create_timeline_chart(df):
    if 'AnalysisDateTime' in df.columns and 'NOTAS' in df.columns:
        df_timeline = df.set_index('AnalysisDateTime').resample('D')['NOTAS'].agg(['mean', 'count']).reset_index()
        df_timeline.columns = ['Data', 'Score Médio', 'Quantidade']
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=df_timeline['Data'],
            y=df_timeline['Score Médio'],
            mode='lines+markers',
            name='Score Médio',
            line=dict(color=CARGLASS_RED, width=3),
            marker=dict(size=8, color=CARGLASS_RED),
            yaxis='y',
            hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Score: %{y:.1f}<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            x=df_timeline['Data'],
            y=df_timeline['Quantidade'],
            name='Quantidade de Análises',
            marker_color=CARGLASS_BLUE,
            opacity=0.3,
            yaxis='y2',
            hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Análises: %{y}<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': '📈 Evolução Temporal do Score',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED}
            },
            xaxis_title='Data',
            yaxis=dict(
                title='Score Médio',
                titlefont=dict(color=CARGLASS_RED),
                tickfont=dict(color=CARGLASS_RED)
            ),
            yaxis2=dict(
                title='Quantidade de Análises',
                titlefont=dict(color=CARGLASS_BLUE),
                tickfont=dict(color=CARGLASS_BLUE),
                anchor='x',
                overlaying='y',
                side='right'
            ),
            height=400,
            plot_bgcolor='white',
            paper_bgcolor='white',
            hovermode='x unified',
            font={'color': CARGLASS_DARK_RED}
        )
        
        fig.add_hline(y=70, line_dash="dash", line_color=CARGLASS_GREEN,
                      annotation_text="Meta: 70", annotation_position="left")
        
        return fig
    return None

with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 20px; background: white; border-radius: 15px; margin-bottom: 20px;'>
        <h2 style='color: """ + CARGLASS_RED + """; margin: 0;'>🔴 Monitor AI</h2>
        <p style='color: """ + CARGLASS_GRAY + """; margin-top: 10px;'>Sistema de Análise Inteligente</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### ⚙️ Configurações")
    
    uploaded_file = st.file_uploader(
        "📁 Carregar arquivo Excel",
        type=['xlsx', 'xls'],
        help="Selecione o arquivo de análise de monitoria"
    )
    
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        
        if df is not None:
            st.success(f"✅ {len(df)} registros carregados")
            
            st.markdown("---")
            st.markdown("### 🔍 Filtros")
            
            date_range = st.date_input(
                "📅 Período de Análise",
                value=(df['AnalysisDateTime'].min().date(), df['AnalysisDateTime'].max().date()),
                min_value=df['AnalysisDateTime'].min().date(),
                max_value=df['AnalysisDateTime'].max().date(),
                format="DD/MM/YYYY"
            )
            
            if len(date_range) == 2:
                start_date, end_date = date_range
                mask = (df['AnalysisDateTime'].dt.date >= start_date) & (df['AnalysisDateTime'].dt.date <= end_date)
                df = df[mask]
            
            agents = st.multiselect(
                "👤 Selecionar Agentes",
                options=df['CustomerAgent'].unique() if 'CustomerAgent' in df.columns else [],
                default=None
            )
            
            if agents:
                df = df[df['CustomerAgent'].isin(agents)]
            
            risk_levels = st.multiselect(
                "⚠️ Nível de Risco",
                options=df['ClientRisk'].unique() if 'ClientRisk' in df.columns else [],
                default=None
            )
            
            if risk_levels:
                df = df[df['ClientRisk'].isin(risk_levels)]
            
            satisfaction_levels = st.multiselect(
                "😊 Nível de Satisfação",
                options=df['Client'].unique() if 'Client' in df.columns else [],
                default=None
            )
            
            if satisfaction_levels:
                df = df[df['Client'].isin(satisfaction_levels)]
            
            st.markdown("---")
            
            if st.button("🔄 Resetar Filtros", use_container_width=True):
                st.rerun()
            
            if st.button("📥 Exportar Dados", use_container_width=True):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Dados_Filtrados', index=False)
                output.seek(0)
                st.download_button(
                    label="💾 Download Excel",
                    data=output,
                    file_name=f"monitoria_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        df = None
        st.info("👆 Por favor, carregue um arquivo Excel para começar")

st.markdown("""
<div class='header-container'>
    <h1 style='color: white; margin: 0; font-size: 36px;'>🔴 Monitor AI</h1>
    <p style='color: white; font-size: 18px; margin-top: 10px;'>Dashboard de Indicadores de Gestão - Análise de Atendimento</p>
</div>
""", unsafe_allow_html=True)

if df is not None and len(df) > 0:
    
    col1, col2, col3, col4 = st.columns(4)
    
    total_analyses = len(df)
    avg_score = df['NOTAS'].mean() if 'NOTAS' in df.columns else 0
    
    low_risk_pct = (df['ClientRisk'] == 'BAIXO').sum() / len(df) * 100 if 'ClientRisk' in df.columns else 0
    
    satisfaction_pct = 0
    if 'Client' in df.columns:
        good_satisfaction = df[df['Client'].isin(['BOA', 'ALTA'])].shape[0]
        satisfaction_pct = (good_satisfaction / len(df)) * 100 if len(df) > 0 else 0
    
    with col1:
        st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
        st.metric(
            label="📊 Total de Análises",
            value=f"{total_analyses:,}".replace(',', '.'),
            delta=f"+{len(df[df['AnalysisDateTime'] >= (datetime.now() - timedelta(days=7))])} esta semana"
        )
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
        st.metric(
            label="📈 Score Médio",
            value=f"{avg_score:.1f}",
            delta=f"{'↑' if avg_score >= 70 else '↓'} Meta: 70"
        )
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
        st.metric(
            label="✅ Risco Baixo",
            value=f"{low_risk_pct:.1f}%",
            delta="Dentro do objetivo" if low_risk_pct >= 60 else "Abaixo da meta"
        )
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col4:
        st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
        st.metric(
            label="😊 Taxa Satisfação",
            value=f"{satisfaction_pct:.1f}%",
            delta="Excelente performance" if satisfaction_pct >= 80 else "Necessita melhoria"
        )
        st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        satisfaction_chart = create_satisfaction_donut(df)
        if satisfaction_chart:
            st.plotly_chart(satisfaction_chart, use_container_width=True)
    
    with col2:
        performance_chart = create_performance_chart(df)
        if performance_chart:
            st.plotly_chart(performance_chart, use_container_width=True)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        risk_chart = create_risk_analysis(df)
        if risk_chart:
            st.plotly_chart(risk_chart, use_container_width=True)
        
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        st.markdown("### 🎯 Pontos de Melhoria")
        
        weak_questions = []
        for i in range(1, 13):
            q = f'Question{i}'
            if q in df.columns:
                performance = df[q].mean() * 100
                if performance < 70:
                    weak_questions.append((q, performance))
        
        weak_questions.sort(key=lambda x: x[1])
        
        question_names = {
            'Question1': 'Saudação',
            'Question2': 'Dados Cadastrais',
            'Question3': 'LGPD',
            'Question4': 'Técnica do Eco',
            'Question5': 'Escuta Ativa',
            'Question6': 'Conhecimento',
            'Question7': 'Confirmação',
            'Question8': 'Seleção Loja',
            'Question9': 'Comunicação',
            'Question10': 'Conduta',
            'Question11': 'Encerramento',
            'Question12': 'Pesquisa'
        }
        
        for q, perf in weak_questions[:3]:
            q_name = question_names.get(q, q)
            color = CARGLASS_ORANGE if perf >= 50 else CARGLASS_RED
            st.markdown(f"<p style='color: {color}; font-weight: bold;'>• {q_name}: {perf:.1f}%</p>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        agent_ranking = create_agent_ranking(df)
        if agent_ranking:
            st.plotly_chart(agent_ranking, use_container_width=True)
        
        outcome_pie = None
        if 'ClientOutcome' in df.columns:
            outcome_counts = df['ClientOutcome'].value_counts()
            
            colors_map = {
                'RESOLVIDO': CARGLASS_GREEN,
                'PARCIALMENTE_RESOLVIDO': CARGLASS_YELLOW,
                'PENDENTE': CARGLASS_ORANGE,
                'INCOMPLETO': CARGLASS_RED,
                'NAO_RESOLVIDO': CARGLASS_RED
            }
            
            outcome_pie = go.Figure(data=[go.Pie(
                labels=outcome_counts.index,
                values=outcome_counts.values,
                hole=.5,
                marker=dict(colors=[colors_map.get(x, CARGLASS_GRAY) for x in outcome_counts.index]),
                textinfo='label+percent',
                hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
            )])
            
            outcome_pie.update_layout(
                title={
                    'text': '📞 Desfecho das Ligações',
                    'font': {'size': 20, 'color': CARGLASS_DARK_RED}
                },
                height=350,
                paper_bgcolor='white',
                font={'color': CARGLASS_DARK_RED}
            )
            
            st.plotly_chart(outcome_pie, use_container_width=True)
    
    timeline = create_timeline_chart(df)
    if timeline:
        st.plotly_chart(timeline, use_container_width=True)
    
    st.markdown("## 📊 Análise Detalhada por Agente")
    
    tab1, tab2, tab3 = st.tabs(["📈 Performance Individual", "🎯 Comparativo", "📝 Detalhes"])
    
    with tab1:
        selected_agent = st.selectbox(
            "Selecione o Agente",
            options=df['CustomerAgent'].unique() if 'CustomerAgent' in df.columns else []
        )
        
        if selected_agent:
            agent_df = df[df['CustomerAgent'] == selected_agent]
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Ligações", len(agent_df))
            with col2:
                st.metric("Score Médio", f"{agent_df['NOTAS'].mean():.1f}")
            with col3:
                risk_baixo = (agent_df['ClientRisk'] == 'BAIXO').sum() / len(agent_df) * 100 if 'ClientRisk' in agent_df.columns else 0
                st.metric("Risco Baixo", f"{risk_baixo:.1f}%")
            with col4:
                satisfaction = agent_df[agent_df['Client'].isin(['BOA', 'ALTA'])].shape[0] / len(agent_df) * 100 if 'Client' in agent_df.columns else 0
                st.metric("Satisfação", f"{satisfaction:.1f}%")
            
            questions_performance = []
            for i in range(1, 13):
                q = f'Question{i}'
                if q in agent_df.columns:
                    perf = agent_df[q].mean() * 100
                    questions_performance.append({
                        'Critério': question_names.get(q, q),
                        'Performance': perf
                    })
            
            if questions_performance:
                perf_df = pd.DataFrame(questions_performance)
                
                fig = go.Figure(go.Bar(
                    x=perf_df['Performance'],
                    y=perf_df['Critério'],
                    orientation='h',
                    marker_color=[CARGLASS_GREEN if p >= 70 else CARGLASS_ORANGE if p >= 50 else CARGLASS_RED 
                                 for p in perf_df['Performance']],
                    text=[f'{p:.1f}%' for p in perf_df['Performance']],
                    textposition='outside'
                ))
                
                fig.update_layout(
                    title=f'Performance de {selected_agent}',
                    xaxis_title='Performance (%)',
                    height=400,
                    xaxis={'range': [0, 110]},
                    plot_bgcolor='white',
                    paper_bgcolor='white'
                )
                
                fig.add_vline(x=70, line_dash="dash", line_color=CARGLASS_GREEN,
                            annotation_text="Meta", annotation_position="top")
                
                st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        if 'CustomerAgent' in df.columns and 'NOTAS' in df.columns:
            agent_comparison = df.groupby('CustomerAgent').agg({
                'NOTAS': 'mean',
                'IdAnalysis': 'count',
                'ClientRisk': lambda x: (x == 'BAIXO').sum() / len(x) * 100 if len(x) > 0 else 0
            }).round(1)
            
            agent_comparison.columns = ['Score Médio', 'Total Ligações', '% Risco Baixo']
            agent_comparison = agent_comparison.sort_values('Score Médio', ascending=False)
            
            fig = go.Figure()
            
            fig.add_trace(go.Scatter(
                x=agent_comparison['Total Ligações'],
                y=agent_comparison['Score Médio'],
                mode='markers+text',
                marker=dict(
                    size=agent_comparison['% Risco Baixo'],
                    color=agent_comparison['Score Médio'],
                    colorscale=[[0, CARGLASS_RED], [0.5, CARGLASS_YELLOW], [1, CARGLASS_GREEN]],
                    showscale=True,
                    colorbar=dict(title="Score<br>Médio"),
                    line=dict(width=2, color='white')
                ),
                text=[name.split()[0] for name in agent_comparison.index],
                textposition='top center',
                hovertemplate='<b>%{text}</b><br>Ligações: %{x}<br>Score: %{y:.1f}<br>Risco Baixo: %{marker.size:.1f}%<extra></extra>'
            ))
            
            fig.update_layout(
                title='Análise Comparativa de Agentes',
                xaxis_title='Total de Ligações',
                yaxis_title='Score Médio',
                height=500,
                plot_bgcolor='white',
                paper_bgcolor='white'
            )
            
            fig.add_hline(y=70, line_dash="dash", line_color=CARGLASS_GREEN,
                         annotation_text="Meta Score", annotation_position="right")
            
            st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(
                agent_comparison.style.background_gradient(subset=['Score Médio'], cmap='RdYlGn', vmin=0, vmax=100)
                                     .background_gradient(subset=['% Risco Baixo'], cmap='RdYlGn', vmin=0, vmax=100),
                use_container_width=True
            )
    
    with tab3:
        st.markdown("### 📋 Dados Detalhados")
        
        display_columns = ['AnalysisDateTime', 'CustomerAgent', 'Client', 'ClientRisk', 'NOTAS', 'Summary']
        available_columns = [col for col in display_columns if col in df.columns]
        
        if available_columns:
            st.dataframe(
                df[available_columns].sort_values('AnalysisDateTime', ascending=False).head(100),
                use_container_width=True,
                hide_index=True
            )
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📊 Estatísticas Gerais")
            stats_df = pd.DataFrame({
                'Métrica': ['Score Médio', 'Score Mediano', 'Desvio Padrão', 'Score Mínimo', 'Score Máximo'],
                'Valor': [
                    f"{df['NOTAS'].mean():.2f}",
                    f"{df['NOTAS'].median():.2f}",
                    f"{df['NOTAS'].std():.2f}",
                    f"{df['NOTAS'].min():.2f}",
                    f"{df['NOTAS'].max():.2f}"
                ]
            })
            st.dataframe(stats_df, use_container_width=True, hide_index=True)
        
        with col2:
            st.markdown("### 🏆 Rankings")
            if 'CustomerAgent' in df.columns:
                best_agents = df.groupby('CustomerAgent')['NOTAS'].mean().sort_values(ascending=False).head(3)
                worst_agents = df.groupby('CustomerAgent')['NOTAS'].mean().sort_values(ascending=True).head(3)
                
                st.markdown("**Top 3 Melhores:**")
                for i, (agent, score) in enumerate(best_agents.items(), 1):
                    medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉"
                    st.markdown(f"{medal} {agent}: {score:.1f}")
                
                st.markdown("**3 Para Melhorar:**")
                for agent, score in worst_agents.items():
                    st.markdown(f"📈 {agent}: {score:.1f}")

else:
    st.info("📁 Por favor, carregue um arquivo Excel na barra lateral para visualizar o dashboard")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class='card'>
            <h3 style='color: """ + CARGLASS_RED + """;'>📊 Análise Completa</h3>
            <p>Visualize métricas detalhadas de performance, satisfação e risco dos atendimentos</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class='card'>
            <h3 style='color: """ + CARGLASS_RED + """;'>👥 Performance Individual</h3>
            <p>Acompanhe o desempenho de cada agente com indicadores personalizados</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class='card'>
            <h3 style='color: """ + CARGLASS_RED + """;'>📈 Tendências</h3>
            <p>Identifique padrões e oportunidades de melhoria ao longo do tempo</p>
        </div>
        """, unsafe_allow_html=True)
