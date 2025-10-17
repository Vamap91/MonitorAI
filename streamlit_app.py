import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Monitor AI - Carglass",
    page_icon="🔴",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de cores Carglass
CARGLASS_RED = "#DC0A0A"
CARGLASS_DARK_RED = "#B00000"
CARGLASS_BLUE = "#4A90E2"
CARGLASS_DARK_BLUE = "#2C5AA0"
CARGLASS_PURPLE = "#6B5B95"
CARGLASS_LIGHT_PURPLE = "#8B7AB8"
CARGLASS_GRAY = "#6C757D"
CARGLASS_LIGHT_GRAY = "#F8F9FA"
CARGLASS_GREEN = "#28A745"
CARGLASS_YELLOW = "#FFC107"
CARGLASS_ORANGE = "#FD7E14"

# CSS personalizado com design moderno
custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background: #F5F7FA;
    }
    
    /* Header com gradiente */
    .header-gradient {
        background: linear-gradient(135deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 50%, """ + CARGLASS_DARK_BLUE + """ 100%);
        padding: 40px;
        border-radius: 20px;
        margin-bottom: 30px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.15);
        position: relative;
        overflow: hidden;
    }
    
    .header-gradient::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 400px;
        height: 400px;
        background: rgba(255,255,255,0.1);
        border-radius: 50%;
    }
    
    .header-gradient h1 {
        color: white;
        margin: 0;
        font-size: 42px;
        font-weight: 700;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .header-gradient p {
        color: rgba(255,255,255,0.95);
        font-size: 18px;
        margin-top: 10px;
        font-weight: 400;
    }
    
    /* KPI Cards modernos */
    .kpi-card-modern {
        background: linear-gradient(135deg, """ + CARGLASS_PURPLE + """ 0%, """ + CARGLASS_LIGHT_PURPLE + """ 100%);
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
        margin-bottom: 20px;
    }
    
    .kpi-card-modern::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -20%;
        width: 200px;
        height: 200px;
        background: rgba(255,255,255,0.1);
        border-radius: 50%;
    }
    
    .kpi-card-modern:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 32px rgba(0,0,0,0.18);
    }
    
    .kpi-value {
        font-size: 48px;
        font-weight: 700;
        color: white;
        margin: 10px 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .kpi-label {
        font-size: 14px;
        font-weight: 600;
        color: rgba(255,255,255,0.9);
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .kpi-delta {
        font-size: 13px;
        color: rgba(255,255,255,0.85);
        margin-top: 8px;
        display: flex;
        align-items: center;
        gap: 5px;
    }
    
    /* Cards de conteúdo */
    .content-card {
        background: white;
        padding: 25px;
        border-radius: 16px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.08);
        margin-bottom: 20px;
        border: 1px solid rgba(0,0,0,0.05);
    }
    
    .content-card h3 {
        color: """ + CARGLASS_DARK_RED + """;
        font-size: 20px;
        font-weight: 600;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    /* Sidebar toggle button */
    .sidebar-toggle {
        position: fixed;
        left: 10px;
        top: 80px;
        z-index: 999;
        background: """ + CARGLASS_RED + """;
        color: white;
        border: none;
        border-radius: 50%;
        width: 45px;
        height: 45px;
        cursor: pointer;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        transition: all 0.3s ease;
    }
    
    .sidebar-toggle:hover {
        background: """ + CARGLASS_DARK_RED + """;
        transform: scale(1.1);
    }
    
    /* Métricas Streamlit customizadas */
    div[data-testid="metric-container"] {
        background: transparent;
        padding: 0;
    }
    
    div[data-testid="metric-container"] > label {
        color: rgba(255,255,255,0.9) !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    div[data-testid="metric-container"] > div[data-testid="stMetricValue"] {
        color: white !important;
        font-weight: 700 !important;
        font-size: 42px !important;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    div[data-testid="metric-container"] > div[data-testid="stMetricDelta"] {
        color: rgba(255,255,255,0.85) !important;
        font-size: 13px !important;
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
    }
    
    section[data-testid="stSidebar"] .stMarkdown {
        color: white;
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: white;
        border-radius: 8px;
        padding: 12px 24px;
        font-weight: 600;
        border: 2px solid transparent;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
        color: white;
    }
    
    /* Botões */
    .stButton > button {
        background: linear-gradient(135deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
    }
    
    /* Selectbox e inputs */
    .stSelectbox label, .stMultiSelect label, .stDateInput label {
        color: white !important;
        font-weight: 600 !important;
    }
    
    /* Ajustes para gráficos */
    .js-plotly-plot {
        border-radius: 12px;
        overflow: hidden;
    }
</style>
"""

st.markdown(custom_css, unsafe_allow_html=True)

# Função para carregar dados
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

# Função para criar gráfico de gauge
def create_gauge_chart(value, title, color, reference=70):
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value,
        title={'text': title, 'font': {'size': 18, 'color': CARGLASS_DARK_RED, 'family': 'Inter'}},
        delta={'reference': reference, 'increasing': {'color': CARGLASS_GREEN}, 'decreasing': {'color': CARGLASS_RED}},
        gauge={
            'axis': {'range': [None, 100], 'tickwidth': 2, 'tickcolor': CARGLASS_GRAY},
            'bar': {'color': color, 'thickness': 0.75},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': CARGLASS_LIGHT_GRAY,
            'steps': [
                {'range': [0, 50], 'color': '#FFE5E5'},
                {'range': [50, 70], 'color': '#FFF5E5'},
                {'range': [70, 100], 'color': '#E5FFE5'}
            ],
            'threshold': {
                'line': {'color': CARGLASS_DARK_RED, 'width': 4},
                'thickness': 0.75,
                'value': reference
            }
        }
    ))
    
    fig.update_layout(
        height=280,
        margin=dict(l=30, r=30, t=60, b=30),
        paper_bgcolor='white',
        font={'color': CARGLASS_DARK_RED, 'family': 'Inter'}
    )
    
    return fig

# Função para criar gráfico de performance por critério
def create_performance_chart(df):
    questions = [f'Question{i}' for i in range(1, 13)]
    question_labels = [
        'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 
        'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12'
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
            marker=dict(
                color=colors,
                line=dict(color='white', width=2)
            ),
            text=[f'{p:.1f}%' for p in performance],
            textposition='outside',
            textfont=dict(size=12, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{x}</b><br>Performance: %{y:.1f}%<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title={
            'text': '✅ Performance do Checklist por Critério',
            'font': {'size': 22, 'color': CARGLASS_DARK_RED, 'family': 'Inter', 'weight': 'bold'},
            'x': 0.5,
            'xanchor': 'center'
        },
        xaxis=dict(
            title='Critérios de Avaliação',
            titlefont=dict(size=14, color=CARGLASS_GRAY, family='Inter'),
            tickfont=dict(size=12, color=CARGLASS_GRAY, family='Inter')
        ),
        yaxis=dict(
            title='Performance (%)',
            range=[0, 110],
            titlefont=dict(size=14, color=CARGLASS_GRAY, family='Inter'),
            tickfont=dict(size=12, color=CARGLASS_GRAY, family='Inter')
        ),
        plot_bgcolor='#FAFBFC',
        paper_bgcolor='white',
        height=450,
        showlegend=False,
        font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
        margin=dict(l=60, r=40, t=80, b=60)
    )
    
    fig.add_hline(
        y=70, 
        line_dash="dash", 
        line_color=CARGLASS_GREEN,
        line_width=2,
        annotation_text="Meta: 70%", 
        annotation_position="right",
        annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
    )
    
    return fig

# Função para criar gráfico donut de satisfação
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
        
        colors_list = []
        for label in labels:
            if label in ['Boa', 'Alta']:
                colors_list.append('#4ECDC4')
            elif label == 'Moderada':
                colors_list.append('#FFE66D')
            else:
                colors_list.append('#FF6B6B')
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=.6,
            marker=dict(
                colors=colors_list,
                line=dict(color='white', width=3)
            ),
            textinfo='label+percent',
            textposition='outside',
            textfont=dict(size=13, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title={
                'text': '😊 Distribuição de Satisfação do Cliente',
                'font': {'size': 22, 'color': CARGLASS_DARK_RED, 'family': 'Inter', 'weight': 'bold'},
                'x': 0.5,
                'xanchor': 'center'
            },
            height=450,
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.15,
                xanchor="center",
                x=0.5,
                font=dict(size=12, family='Inter')
            ),
            margin=dict(l=40, r=40, t=80, b=100)
        )
        
        return fig
    return None

# Função para criar análise de risco
def create_risk_analysis(df):
    if 'ClientRisk' in df.columns:
        risk_counts = df['ClientRisk'].value_counts()
        total = len(df)
        
        risk_data = []
        risk_labels = []
        risk_colors = []
        
        for risk in ['BAIXO', 'MEDIO', 'ALTO']:
            if risk in risk_counts.index:
                count = risk_counts[risk]
                pct = count / total * 100
                risk_data.append(count)
                risk_labels.append(f'{risk.capitalize()}\n{count} casos ({pct:.1f}%)')
                
                if risk == 'BAIXO':
                    risk_colors.append(CARGLASS_GREEN)
                elif risk == 'MEDIO':
                    risk_colors.append(CARGLASS_YELLOW)
                else:
                    risk_colors.append(CARGLASS_RED)
        
        fig = go.Figure(data=[
            go.Bar(
                x=['Risco Baixo', 'Risco Médio', 'Risco Alto'][:len(risk_data)],
                y=risk_data,
                marker=dict(
                    color=risk_colors,
                    line=dict(color='white', width=2)
                ),
                text=risk_labels,
                textposition='outside',
                textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
                hovertemplate='<b>%{x}</b><br>Casos: %{y}<extra></extra>'
            )
        ])
        
        fig.update_layout(
            title={
                'text': '⚠️ Análise de Risco',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter', 'weight': 'bold'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title='',
                tickfont=dict(size=12, color=CARGLASS_GRAY, family='Inter')
            ),
            yaxis=dict(
                title='Número de Casos',
                titlefont=dict(size=13, color=CARGLASS_GRAY, family='Inter'),
                tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
            ),
            height=350,
            showlegend=False,
            plot_bgcolor='#FAFBFC',
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            margin=dict(l=60, r=40, t=70, b=50)
        )
        
        return fig
    return None

# Função para criar ranking de agentes
def create_agent_ranking(df, top_n=5):
    if 'CustomerAgent' in df.columns and 'NOTAS' in df.columns:
        agent_scores = df.groupby('CustomerAgent').agg({
            'NOTAS': 'mean',
            'IdAnalysis': 'count'
        }).round(1)
        
        agent_scores.columns = ['Score Médio', 'Total Ligações']
        agent_scores = agent_scores.sort_values('Score Médio', ascending=False).head(top_n)
        
        # Formatar nomes (primeiro e último nome)
        agent_names = []
        for name in agent_scores.index:
            parts = name.split()
            if len(parts) >= 2:
                agent_names.append(f"{parts[0]} {parts[-1]}")
            else:
                agent_names.append(parts[0])
        
        colors_agents = [CARGLASS_PURPLE if i == 0 else CARGLASS_LIGHT_PURPLE for i in range(len(agent_names))]
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=agent_scores['Score Médio'],
            y=agent_names,
            orientation='h',
            marker=dict(
                color=colors_agents,
                line=dict(color='white', width=2)
            ),
            text=[f"{score:.1f} pts<br>{calls} ligações" 
                  for score, calls in zip(agent_scores['Score Médio'], agent_scores['Total Ligações'])],
            textposition='outside',
            textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{y}</b><br>Score: %{x:.1f}<br>Ligações: %{text}<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': '👥 Top 5 Agentes por Score',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter', 'weight': 'bold'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title='Score Médio',
                range=[0, max(agent_scores['Score Médio']) * 1.2],
                titlefont=dict(size=13, color=CARGLASS_GRAY, family='Inter'),
                tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
            ),
            yaxis=dict(
                title='',
                tickfont=dict(size=12, color=CARGLASS_DARK_RED, family='Inter', weight='bold')
            ),
            height=350,
            showlegend=False,
            plot_bgcolor='#FAFBFC',
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            margin=dict(l=150, r=100, t=70, b=50)
        )
        
        return fig
    return None

# Função para criar gráfico de desfecho das ligações
def create_outcome_chart(df):
    if 'ClientOutcome' in df.columns:
        outcome_counts = df['ClientOutcome'].value_counts()
        
        # Mapear cores
        colors_map = {
            'POSITIVO': '#4ECDC4',
            'RESOLVIDO': '#4ECDC4',
            'RESOLVIDO CO': '#95E1D3',
            'PENDENTE': '#FFE66D',
            'INCOMPLETO': '#FF6B6B',
            'NEGATIVO': '#FF6B6B',
            'ATENDIMENTO ': '#A8DADC'
        }
        
        colors_list = [colors_map.get(x, CARGLASS_GRAY) for x in outcome_counts.index]
        
        # Limpar labels
        labels_clean = []
        for label in outcome_counts.index:
            if pd.isna(label):
                labels_clean.append('Não definido')
            else:
                labels_clean.append(str(label).strip().title())
        
        fig = go.Figure(data=[go.Pie(
            labels=labels_clean,
            values=outcome_counts.values,
            hole=.5,
            marker=dict(
                colors=colors_list,
                line=dict(color='white', width=3)
            ),
            textinfo='label+percent',
            textposition='outside',
            textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title={
                'text': '📞 Desfecho das Ligações',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter', 'weight': 'bold'},
                'x': 0.5,
                'xanchor': 'center'
            },
            height=350,
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.05,
                font=dict(size=10, family='Inter')
            ),
            margin=dict(l=40, r=150, t=70, b=40)
        )
        
        return fig
    return None

# Função para criar gráfico de timeline (CORRIGIDA)
def create_timeline_chart(df):
    if 'AnalysisDateTime' in df.columns and 'NOTAS' in df.columns:
        try:
            # Verificar se há dados suficientes
            if len(df) == 0:
                return None
            
            df_timeline = df.set_index('AnalysisDateTime').resample('D')['NOTAS'].agg(['mean', 'count']).reset_index()
            df_timeline.columns = ['Data', 'Score Médio', 'Quantidade']
            
            # Remover linhas com valores nulos
            df_timeline = df_timeline.dropna()
            
            if len(df_timeline) == 0:
                return None
            
            fig = go.Figure()
            
            # Adicionar linha de score médio
            fig.add_trace(go.Scatter(
                x=df_timeline['Data'],
                y=df_timeline['Score Médio'],
                mode='lines+markers',
                name='Score Médio',
                line=dict(color=CARGLASS_RED, width=3),
                marker=dict(size=8, color=CARGLASS_RED, line=dict(color='white', width=2)),
                yaxis='y',
                hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Score: %{y:.1f}<extra></extra>'
            ))
            
            # Adicionar barras de quantidade
            fig.add_trace(go.Bar(
                x=df_timeline['Data'],
                y=df_timeline['Quantidade'],
                name='Quantidade de Análises',
                marker=dict(color=CARGLASS_BLUE, line=dict(color='white', width=1)),
                opacity=0.4,
                yaxis='y2',
                hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Análises: %{y}<extra></extra>'
            ))
            
            # Layout com dois eixos Y
            fig.update_layout(
                title={
                    'text': '📈 Evolução Temporal do Score',
                    'font': {'size': 22, 'color': CARGLASS_DARK_RED, 'family': 'Inter', 'weight': 'bold'},
                    'x': 0.5,
                    'xanchor': 'center'
                },
                xaxis=dict(
                    title='Data',
                    titlefont=dict(size=14, color=CARGLASS_GRAY, family='Inter'),
                    tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
                ),
                yaxis=dict(
                    title='Score Médio',
                    titlefont=dict(color=CARGLASS_RED, size=13, family='Inter'),
                    tickfont=dict(color=CARGLASS_RED, size=11, family='Inter'),
                    side='left'
                ),
                yaxis2=dict(
                    title='Quantidade de Análises',
                    titlefont=dict(color=CARGLASS_BLUE, size=13, family='Inter'),
                    tickfont=dict(color=CARGLASS_BLUE, size=11, family='Inter'),
                    overlaying='y',
                    side='right'
                ),
                height=450,
                plot_bgcolor='#FAFBFC',
                paper_bgcolor='white',
                hovermode='x unified',
                font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1,
                    font=dict(size=11, family='Inter')
                ),
                margin=dict(l=60, r=60, t=80, b=60)
            )
            
            # Adicionar linha de meta
            fig.add_hline(
                y=70, 
                line_dash="dash", 
                line_color=CARGLASS_GREEN,
                line_width=2,
                annotation_text="Meta: 70", 
                annotation_position="left",
                annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
            )
            
            return fig
        except Exception as e:
            st.warning(f"Não foi possível criar o gráfico de evolução temporal: {str(e)}")
            return None
    return None

# Função para criar pontos de melhoria
def create_improvement_points(df):
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
    
    weak_questions = []
    for i in range(1, 13):
        q = f'Question{i}'
        if q in df.columns:
            performance = df[q].mean() * 100
            if performance < 70:
                weak_questions.append((q, performance))
    
    weak_questions.sort(key=lambda x: x[1])
    
    return [(question_names.get(q, q), perf) for q, perf in weak_questions[:3]]

# ============= SIDEBAR =============
with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 25px; background: white; border-radius: 15px; margin-bottom: 25px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);'>
        <h2 style='color: """ + CARGLASS_RED + """; margin: 0; font-size: 28px;'>🔴 Monitor AI</h2>
        <p style='color: """ + CARGLASS_GRAY + """; margin-top: 10px; font-size: 13px;'>Sistema de Análise Inteligente</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📁 Upload de Dados")
    uploaded_file = st.file_uploader(
        "Selecione o arquivo Excel",
        type=['xlsx', 'xls'],
        help="Faça upload do arquivo de análises de monitoria"
    )
    
    if uploaded_file:
        df = load_data(uploaded_file)
        
        if df is not None:
            st.success(f"✅ {len(df)} registros carregados")
            
            st.markdown("---")
            st.markdown("### 🔍 Filtros")
            
            # Filtro de data
            if 'AnalysisDateTime' in df.columns:
                min_date = df['AnalysisDateTime'].min().date()
                max_date = df['AnalysisDateTime'].max().date()
                
                date_range = st.date_input(
                    "Período de Análise",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date
                )
                
                if len(date_range) == 2:
                    df = df[(df['AnalysisDateTime'].dt.date >= date_range[0]) & 
                           (df['AnalysisDateTime'].dt.date <= date_range[1])]
            
            # Filtro de agente
            if 'CustomerAgent' in df.columns:
                agents = ['Todos'] + sorted(df['CustomerAgent'].unique().tolist())
                selected_agent = st.selectbox("Agente", agents)
                
                if selected_agent != 'Todos':
                    df = df[df['CustomerAgent'] == selected_agent]
            
            # Filtro de risco
            if 'ClientRisk' in df.columns:
                risks = ['Todos'] + sorted(df['ClientRisk'].unique().tolist())
                selected_risk = st.selectbox("Nível de Risco", risks)
                
                if selected_risk != 'Todos':
                    df = df[df['ClientRisk'] == selected_risk]
            
            st.markdown("---")
            st.markdown("### 💾 Exportar Dados")
            
            if st.button("📊 Gerar Relatório Excel", use_container_width=True):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Dados Filtrados', index=False)
                
                output.seek(0)
                st.download_button(
                    label="💾 Download Excel",
                    data=output,
                    file_name=f"monitoria_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    else:
        df = None
        st.info("👆 Carregue um arquivo para começar")

# ============= MAIN CONTENT =============

# Header com gradiente
st.markdown("""
<div class='header-gradient'>
    <h1>🔴 Monitor AI</h1>
    <p>Dashboard de Indicadores de Gestão - Análise de Atendimento</p>
</div>
""", unsafe_allow_html=True)

if df is not None and len(df) > 0:
    
    # ============= KPIs PRINCIPAIS =============
    col1, col2, col3, col4 = st.columns(4)
    
    total_analyses = len(df)
    avg_score = df['NOTAS'].mean() if 'NOTAS' in df.columns else 0
    avg_score_pct = (avg_score / 81) * 100  # Converter para percentual (81 é o máximo)
    
    low_risk_pct = (df['ClientRisk'] == 'BAIXO').sum() / len(df) * 100 if 'ClientRisk' in df.columns else 0
    
    satisfaction_pct = 0
    if 'Client' in df.columns:
        good_satisfaction = df[df['Client'].isin(['BOA', 'ALTA'])].shape[0]
        satisfaction_pct = (good_satisfaction / len(df)) * 100 if len(df) > 0 else 0
    
    # Calcular variação semanal
    if 'AnalysisDateTime' in df.columns:
        last_week = df[df['AnalysisDateTime'] >= (datetime.now() - timedelta(days=7))]
        week_count = len(last_week)
        week_delta = f"📈 +{week_count} esta semana"
    else:
        week_delta = ""
    
    with col1:
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Total de Análises</div>
            <div class='kpi-value'>{total_analyses:,}</div>
            <div class='kpi-delta'>{week_delta}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        delta_icon = "⚠️" if avg_score < 70 else "✅"
        delta_text = f"{delta_icon} {'Abaixo' if avg_score < 70 else 'Acima'} da meta (70)"
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Score Médio</div>
            <div class='kpi-value'>{avg_score:.1f}</div>
            <div class='kpi-delta'>{delta_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        risk_icon = "✅" if low_risk_pct >= 60 else "⚠️"
        risk_text = f"{risk_icon} {'Dentro' if low_risk_pct >= 60 else 'Abaixo'} do objetivo"
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Risco Baixo</div>
            <div class='kpi-value'>{low_risk_pct:.1f}%</div>
            <div class='kpi-delta'>{risk_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        sat_icon = "🎯" if satisfaction_pct >= 80 else "📊"
        sat_text = f"{sat_icon} {'Excelente' if satisfaction_pct >= 80 else 'Bom'} desempenho"
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Taxa Saudação</div>
            <div class='kpi-value'>{satisfaction_pct:.1f}%</div>
            <div class='kpi-delta'>{sat_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # ============= GRÁFICOS PRINCIPAIS =============
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        satisfaction_chart = create_satisfaction_donut(df)
        if satisfaction_chart:
            st.plotly_chart(satisfaction_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        performance_chart = create_performance_chart(df)
        if performance_chart:
            st.plotly_chart(performance_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    # ============= SEGUNDA LINHA DE GRÁFICOS =============
    col1, col2, col3, col4 = st.columns([1.2, 1.5, 1.2, 1.1])
    
    with col1:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        risk_chart = create_risk_analysis(df)
        if risk_chart:
            st.plotly_chart(risk_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        agent_ranking = create_agent_ranking(df)
        if agent_ranking:
            st.plotly_chart(agent_ranking, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        outcome_chart = create_outcome_chart(df)
        if outcome_chart:
            st.plotly_chart(outcome_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col4:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color: " + CARGLASS_DARK_RED + "; font-size: 18px; margin-bottom: 20px;'>🎯 Pontos de Melhoria</h3>", unsafe_allow_html=True)
        
        improvement_points = create_improvement_points(df)
        
        for q_name, perf in improvement_points:
            if perf < 50:
                color = CARGLASS_RED
                icon = "🔴"
            else:
                color = CARGLASS_ORANGE
                icon = "🟠"
            
            st.markdown(f"""
            <div style='margin-bottom: 15px; padding: 12px; background: #F8F9FA; border-radius: 8px; border-left: 4px solid {color};'>
                <div style='font-weight: 600; color: {CARGLASS_DARK_RED}; font-size: 13px;'>{icon} {q_name}</div>
                <div style='color: {color}; font-weight: bold; font-size: 16px; margin-top: 5px;'>{perf:.1f}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    # ============= GRÁFICO DE EVOLUÇÃO TEMPORAL =============
    st.markdown("<div class='content-card'>", unsafe_allow_html=True)
    timeline = create_timeline_chart(df)
    if timeline:
        st.plotly_chart(timeline, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    # ============= ANÁLISE DETALHADA POR AGENTE =============
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("## 📊 Análise Detalhada por Agente")
    
    tab1, tab2, tab3 = st.tabs(["📈 Performance Individual", "🎯 Comparativo", "📝 Detalhes"])
    
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
    
    with tab1:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        
        selected_agent = st.selectbox(
            "Selecione o Agente para Análise Detalhada",
            options=sorted(df['CustomerAgent'].unique()) if 'CustomerAgent' in df.columns else [],
            key='agent_detail'
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
            
            # Gráfico de performance por questão
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
                    marker=dict(
                        color=[CARGLASS_GREEN if p >= 70 else CARGLASS_ORANGE if p >= 50 else CARGLASS_RED 
                               for p in perf_df['Performance']],
                        line=dict(color='white', width=2)
                    ),
                    text=[f'{p:.1f}%' for p in perf_df['Performance']],
                    textposition='outside',
                    textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter', weight='bold')
                ))
                
                fig.update_layout(
                    title=f'Performance de {selected_agent}',
                    xaxis=dict(
                        title='Performance (%)',
                        range=[0, 110],
                        titlefont=dict(size=13, color=CARGLASS_GRAY, family='Inter'),
                        tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
                    ),
                    yaxis=dict(
                        tickfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter')
                    ),
                    height=450,
                    plot_bgcolor='#FAFBFC',
                    paper_bgcolor='white',
                    font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                    margin=dict(l=150, r=80, t=60, b=60)
                )
                
                fig.add_vline(
                    x=70, 
                    line_dash="dash", 
                    line_color=CARGLASS_GREEN,
                    line_width=2,
                    annotation_text="Meta", 
                    annotation_position="top",
                    annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with tab2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        
        if 'CustomerAgent' in df.columns and 'NOTAS' in df.columns:
            agent_comparison = df.groupby('CustomerAgent').agg({
                'NOTAS': 'mean',
                'IdAnalysis': 'count',
                'ClientRisk': lambda x: (x == 'BAIXO').sum() / len(x) * 100 if len(x) > 0 else 0
            }).round(1)
            
            agent_comparison.columns = ['Score Médio', 'Total Ligações', '% Risco Baixo']
            agent_comparison = agent_comparison.sort_values('Score Médio', ascending=False)
            
            # Gráfico de dispersão
            fig = go.Figure()
            
            fig.add_trace(go.Scatter(
                x=agent_comparison['Total Ligações'],
                y=agent_comparison['Score Médio'],
                mode='markers+text',
                marker=dict(
                    size=agent_comparison['% Risco Baixo'] * 0.8,
                    color=agent_comparison['Score Médio'],
                    colorscale=[[0, CARGLASS_RED], [0.5, CARGLASS_YELLOW], [1, CARGLASS_GREEN]],
                    showscale=True,
                    colorbar=dict(
                        title="Score<br>Médio",
                        titlefont=dict(size=11, family='Inter'),
                        tickfont=dict(size=10, family='Inter')
                    ),
                    line=dict(width=2, color='white')
                ),
                text=[name.split()[0] for name in agent_comparison.index],
                textposition='top center',
                textfont=dict(size=9, color=CARGLASS_DARK_RED, family='Inter'),
                hovertemplate='<b>%{text}</b><br>Ligações: %{x}<br>Score: %{y:.1f}<br>Risco Baixo: %{marker.size:.1f}%<extra></extra>'
            ))
            
            fig.update_layout(
                title='Análise Comparativa de Agentes',
                xaxis=dict(
                    title='Total de Ligações',
                    titlefont=dict(size=13, color=CARGLASS_GRAY, family='Inter'),
                    tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
                ),
                yaxis=dict(
                    title='Score Médio',
                    titlefont=dict(size=13, color=CARGLASS_GRAY, family='Inter'),
                    tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
                ),
                height=500,
                plot_bgcolor='#FAFBFC',
                paper_bgcolor='white',
                font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                margin=dict(l=60, r=100, t=60, b=60)
            )
            
            fig.add_hline(
                y=70, 
                line_dash="dash", 
                line_color=CARGLASS_GREEN,
                line_width=2,
                annotation_text="Meta Score", 
                annotation_position="right",
                annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Tabela de comparação
            st.dataframe(
                agent_comparison.style.background_gradient(subset=['Score Médio'], cmap='RdYlGn', vmin=0, vmax=100)
                                     .background_gradient(subset=['% Risco Baixo'], cmap='RdYlGn', vmin=0, vmax=100)
                                     .format({'Score Médio': '{:.1f}', '% Risco Baixo': '{:.1f}%'}),
                use_container_width=True,
                height=400
            )
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with tab3:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        st.markdown("### 📋 Dados Detalhados")
        
        display_columns = ['AnalysisDateTime', 'CustomerAgent', 'Client', 'ClientRisk', 'ClientOutcome', 'NOTAS']
        available_columns = [col for col in display_columns if col in df.columns]
        
        if available_columns:
            st.dataframe(
                df[available_columns].sort_values('AnalysisDateTime', ascending=False).head(100),
                use_container_width=True,
                hide_index=True,
                height=400
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
                
                st.markdown("<br>**3 Para Melhorar:**", unsafe_allow_html=True)
                for agent, score in worst_agents.items():
                    st.markdown(f"📈 {agent}: {score:.1f}")
        
        st.markdown("</div>", unsafe_allow_html=True)

else:
    # Tela de boas-vindas
    st.info("📁 Por favor, carregue um arquivo Excel na barra lateral para visualizar o dashboard")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class='content-card'>
            <h3 style='color: """ + CARGLASS_RED + """;'>📊 Análise Completa</h3>
            <p style='color: """ + CARGLASS_GRAY + """;'>Visualize métricas detalhadas de performance, satisfação e risco dos atendimentos</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class='content-card'>
            <h3 style='color: """ + CARGLASS_RED + """;'>👥 Performance Individual</h3>
            <p style='color: """ + CARGLASS_GRAY + """;'>Acompanhe o desempenho de cada agente com indicadores personalizados</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class='content-card'>
            <h3 style='color: """ + CARGLASS_RED + """;'>📈 Tendências</h3>
            <p style='color: """ + CARGLASS_GRAY + """;'>Identifique padrões e oportunidades de melhoria ao longo do tempo</p>
        </div>
        """, unsafe_allow_html=True)

# Botão de toggle da sidebar (JavaScript)
st.markdown("""
<script>
function toggleSidebar() {
    const sidebar = window.parent.document.querySelector('[data-testid="stSidebar"]');
    if (sidebar) {
        sidebar.style.display = sidebar.style.display === 'none' ? 'block' : 'none';
    }
}
</script>
<button class='sidebar-toggle' onclick='toggleSidebar()'>☰</button>
""", unsafe_allow_html=True)
