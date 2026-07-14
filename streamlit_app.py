import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.lib.colors import HexColor
import tempfile
import base64

META_GLOBAL = 85

def get_satisfaction_cluster(value):
    if pd.isna(value):
        return None
    value_upper = str(value).strip().upper()
    if value_upper in ['ALTA', 'ALTO', 'BOA', 'SATISFEITO', 'SATISFEITA']:
        return 'SATISFEITO'
    if value_upper in ['NEUTRA', 'NEUTRO', 'MÉDIA', 'MEDIA', 'MÉDIO', 'MEDIO', 'MODERADA', 'MODERADO']:
        return 'NEUTRO'
    if value_upper in ['BAIXA', 'BAIXO', 'INSATISFEITO', 'INSATISFEITA', 'INSATISFATÓRIO', 'INSATISFATORIA']:
        return 'INSATISFEITO'
    return None

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
CARGLASS_PURPLE = "#6B5B95"
CARGLASS_LIGHT_PURPLE = "#8B7AB8"
CARGLASS_GRAY = "#6C757D"
CARGLASS_LIGHT_GRAY = "#F8F9FA"
CARGLASS_GREEN = "#28A745"
CARGLASS_YELLOW = "#FFC107"
CARGLASS_ORANGE = "#FD7E14"

CHART_TEXT_COLOR = "#2D2D2D"
CHART_BG = "rgba(250,251,252,0.95)"
CHART_PAPER_BG = "rgba(255,255,255,0.95)"
CHART_GRID_COLOR = "rgba(200,200,200,0.4)"

custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background: #F5F7FA;
    }
    
    .header-gradient {
        background: linear-gradient(135deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
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
    
    .kpi-card-modern {
        background: linear-gradient(135deg, """ + CARGLASS_PURPLE + """ 0%, """ + CARGLASS_LIGHT_PURPLE + """ 100%);
        padding: 25px;
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
        font-size: 40px;
        font-weight: 700;
        color: white;
        margin: 8px 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .kpi-label {
        font-size: 12px;
        font-weight: 600;
        color: rgba(255,255,255,0.9);
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .kpi-delta {
        font-size: 12px;
        color: rgba(255,255,255,0.85);
        margin-top: 6px;
        display: flex;
        align-items: center;
        gap: 5px;
    }
    
    .kpi-card-green {
        background: linear-gradient(135deg, #1B8A3B 0%, #28A745 100%);
        padding: 25px;
        border-radius: 20px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
        margin-bottom: 20px;
    }
    
    .kpi-card-green::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -20%;
        width: 200px;
        height: 200px;
        background: rgba(255,255,255,0.1);
        border-radius: 50%;
    }
    
    .kpi-card-red {
        background: linear-gradient(135deg, #B00000 0%, #DC0A0A 100%);
        padding: 25px;
        border-radius: 20px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
        margin-bottom: 20px;
    }
    
    .kpi-card-red::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -20%;
        width: 200px;
        height: 200px;
        background: rgba(255,255,255,0.1);
        border-radius: 50%;
    }
    
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
    
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, """ + CARGLASS_RED + """ 0%, """ + CARGLASS_DARK_RED + """ 100%);
    }
    
    section[data-testid="stSidebar"] .stMarkdown {
        color: white;
    }
    
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
    
    .stSelectbox label, .stMultiSelect label, .stDateInput label {
        color: white !important;
        font-weight: 600 !important;
    }
    
    .js-plotly-plot {
        border-radius: 12px;
        overflow: hidden;
    }
    
    section[data-testid="stSidebar"] label {
        color: white !important;
    }
    
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stDateInput label,
    section[data-testid="stSidebar"] .stFileUploader label {
        color: white !important;
        font-weight: 600 !important;
    }
    
    section[data-testid="stSidebar"] p {
        color: white !important;
    }
    
    section[data-testid="stSidebar"] .stMarkdown p {
        color: white !important;
    }
    
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
    
    .risk-section-card {
        background: white;
        padding: 30px;
        border-radius: 16px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.08);
        margin-bottom: 20px;
        border: 1px solid rgba(0,0,0,0.05);
        border-top: 4px solid """ + CARGLASS_RED + """;
    }
    
    .risk-legend {
        background: #F8F9FA;
        padding: 15px 20px;
        border-radius: 10px;
        margin-bottom: 15px;
        border-left: 4px solid #4A90E2;
        font-size: 13px;
        color: #333;
        line-height: 1.6;
    }
    
    .tooltip-info {
        position: relative;
        display: inline-block;
        cursor: help;
        border-bottom: 1px dashed #999;
    }
</style>
"""

st.markdown(custom_css, unsafe_allow_html=True)

QUESTION_NAMES = {
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

QUESTION_DETAILS = {
    'Question1': '1 - Saudação e Atendimento Inicial (10 pts)',
    'Question2': '2 - Coleta de Dados Cadastrais (6 pts)',
    'Question3': '3 - Script LGPD (2 pts)',
    'Question4': '4 - Técnica do Eco (5 pts)',
    'Question5': '5 - Escuta Ativa (3 pts)',
    'Question6': '6 - Conhecimento do Produto (5 pts)',
    'Question7': '7 - Confirmação de Danos (10 pts)',
    'Question8': '8 - Seleção de Loja (10 pts)',
    'Question9': '9 - Comunicação Eficaz (5 pts)',
    'Question10': '10 - Conduta Acolhedora (4 pts)',
    'Question11': '11 - Script de Encerramento (15 pts)',
    'Question12': '12 - Pesquisa de Satisfação (6 pts)'
}

RISK_DESCRIPTIONS = {
    'BAIXO': '🟢 Risco Baixo: O atendimento foi realizado dentro dos padrões esperados, sem identificação de problemas que possam gerar reclamações, cancelamentos ou insatisfação significativa do cliente.',
    'MEDIO': '🟡 Risco Médio: Foram identificados pontos de atenção no atendimento que podem levar a insatisfação moderada. Recomenda-se acompanhamento e ação preventiva para evitar escalação.',
    'ALTO': '🔴 Risco Alto: O atendimento apresentou falhas críticas que podem resultar em reclamações formais, perda de cliente ou impacto negativo na imagem da empresa. Ação imediata é necessária.'
}

@st.cache_data
def load_data(file):
    try:
        xls = pd.ExcelFile(file)
        if 'Consulta1' in xls.sheet_names:
            df = pd.read_excel(file, sheet_name='Consulta1')
            
            if 'AnalysisDateTime' in df.columns:
                df['AnalysisDateTime'] = pd.to_datetime(df['AnalysisDateTime'])
            if 'CallDate' in df.columns:
                df['CallDate'] = pd.to_datetime(df['CallDate'])
            
            avaliacao_cols = ['Avaliação 100 pts', 'Avaliacao 100 pts', 'Avaliação100pts', 'Avaliacao100pts']
            for col in avaliacao_cols:
                if col in df.columns:
                    df['PERCENTUAL'] = pd.to_numeric(df[col], errors='coerce')
                    break
            
            if 'PERCENTUAL' not in df.columns and 'NOTAS' in df.columns:
                df['PERCENTUAL'] = (df['NOTAS'] / 81) * 100
            
            if 'PERCENTUAL' in df.columns:
                df = df[(df['PERCENTUAL'].notna()) & (df['PERCENTUAL'] >= 19.99)]
            
            if 'ClientRisk' in df.columns:
                df = df[(df['ClientRisk'] != 'INDETERMINADO') & (df['ClientRisk'].notna())]
            
            if 'Empresas' in df.columns:
                df = df[df['Empresas'].notna()]
            
            return df
        else:
            st.error("A planilha 'Consulta1' não foi encontrada no arquivo.")
            return None
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return None


def get_best_worst_questions(df):
    performances = {}
    for i in range(1, 13):
        q = f'Question{i}'
        if q in df.columns:
            perf = df[q].mean() * 100
            performances[QUESTION_NAMES.get(q, q)] = round(perf)
    
    if not performances:
        return None, None, None, None
    
    best_q = max(performances, key=performances.get)
    worst_q = min(performances, key=performances.get)
    return best_q, performances[best_q], worst_q, performances[worst_q]


def _compute_employee_metrics(df, employee_name):
    employee_df = df[df['CustomerAgent'] == employee_name].copy()
    
    score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    avg_score = employee_df[score_col].mean() if score_col == 'PERCENTUAL' else (employee_df[score_col].mean() / 81) * 100
    total_calls = len(employee_df)
    
    if 'ClientRisk' in employee_df.columns and len(employee_df) > 0:
        risk_baixo = (employee_df['ClientRisk'] == 'BAIXO').sum() / len(employee_df) * 100
        risk_alto = (employee_df['ClientRisk'] == 'ALTO').sum() / len(employee_df) * 100
    else:
        risk_baixo = 0
        risk_alto = 0
    
    if 'Client' in employee_df.columns and len(employee_df) > 0:
        employee_df['Client_Cluster'] = employee_df['Client'].apply(get_satisfaction_cluster)
        satisfaction = (employee_df['Client_Cluster'] == 'SATISFEITO').sum() / len(employee_df) * 100
        insatisfaction = (employee_df['Client_Cluster'] == 'INSATISFEITO').sum() / len(employee_df) * 100
    else:
        satisfaction = 0
        insatisfaction = 0
    
    criteria_performance = {}
    for i in range(1, 13):
        q = f'Question{i}'
        if q in employee_df.columns:
            perf = employee_df[q].mean() * 100
            criteria_performance[QUESTION_NAMES.get(q, q)] = perf
    
    return {
        'employee_df': employee_df,
        'score_col': score_col,
        'avg_score': avg_score,
        'total_calls': total_calls,
        'risk_baixo': risk_baixo,
        'risk_alto': risk_alto,
        'satisfaction': satisfaction,
        'insatisfaction': insatisfaction,
        'criteria_performance': criteria_performance
    }


def _build_summary_table(metrics):
    summary_data = [
        ['Métrica', 'Valor', 'Status'],
        ['Porcentagem de Acerto Média', f"{round(metrics['avg_score'])}%", '✓ Bom' if metrics['avg_score'] >= META_GLOBAL else '✗ Abaixo da Meta'],
        ['Total de Ligações Analisadas', str(metrics['total_calls']), '-'],
        ['Taxa de Risco Baixo', f"{round(metrics['risk_baixo'])}%", '✓ Bom' if metrics['risk_baixo'] >= 60 else '✗ Atenção'],
        ['Taxa de Satisfação do Cliente', f"{round(metrics['satisfaction'])}%", '✓ Bom' if metrics['satisfaction'] >= 70 else '✗ Atenção']
    ]
    
    summary_table = Table(summary_data, colWidths=[2.5*inch, 1.5*inch, 1.5*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor(CARGLASS_RED)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 9)
    ]))
    return summary_table


def _build_criteria_table(criteria_performance):
    criteria_data = [['Critério', 'Performance', 'Status']]
    for criterion, perf in sorted(criteria_performance.items(), key=lambda x: x[1], reverse=True):
        status = '✓ Excelente' if perf >= 90 else '✓ Bom' if perf >= META_GLOBAL else '✗ Precisa Melhorar'
        criteria_data.append([criterion, f'{round(perf)}%', status])
    
    criteria_table = Table(criteria_data, colWidths=[2.5*inch, 1.5*inch, 1.5*inch])
    criteria_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor(CARGLASS_RED)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 9)
    ]))
    return criteria_table


def _build_melhoria_text(criteria_performance):
    weak_points = [(k, v) for k, v in criteria_performance.items() if v < META_GLOBAL]
    weak_points.sort(key=lambda x: x[1])
    
    if weak_points:
        melhoria_text = "<b>Os seguintes pontos requerem atenção e desenvolvimento prioritário:</b><br/><br/>"
        for i, (criterion, perf) in enumerate(weak_points, 1):
            melhoria_text += f"<b>{i}. {criterion}</b> ({round(perf)}% de acerto)<br/>"
            if 'Saudação' in criterion:
                melhoria_text += """• Treinar abertura padronizada do atendimento com identificação clara<br/>
• Praticar tom de voz acolhedor e profissional<br/>
• Revisar script de saudação e aplicar consistentemente<br/><br/>"""
            elif 'Dados Cadastrais' in criterion:
                melhoria_text += """• Reforçar importância da coleta completa de informações<br/>
• Praticar técnicas de confirmação de dados<br/>
• Utilizar checklist mental durante o atendimento<br/><br/>"""
            elif 'LGPD' in criterion:
                melhoria_text += """• Treinamento obrigatório sobre Lei Geral de Proteção de Dados<br/>
• Incluir solicitação de consentimento em todas as ligações<br/>
• Revisar políticas de privacidade da empresa<br/><br/>"""
            elif 'Escuta Ativa' in criterion:
                melhoria_text += """• Praticar técnicas de parafraseamento<br/>
• Evitar interrupções durante a fala do cliente<br/>
• Demonstrar compreensão com confirmações verbais<br/><br/>"""
            elif 'Conhecimento' in criterion:
                melhoria_text += """• Intensificar estudo de produtos e serviços<br/>
• Participar de treinamentos técnicos regulares<br/>
• Consultar base de conhecimento antes de cada turno<br/><br/>"""
            else:
                melhoria_text += f"""• Revisar procedimentos padrão relacionados a {criterion}<br/>
• Buscar mentoria com colaboradores de alta performance<br/>
• Praticar através de role-playing e simulações<br/><br/>"""
    else:
        melhoria_text = f"""<b>Parabéns!</b> Todos os critérios estão acima da meta de {META_GLOBAL}%. 
        Foco agora deve ser em refinamento e excelência em todas as áreas.<br/><br/>"""
    return melhoria_text, weak_points


def _build_positivos_text(criteria_performance, satisfaction, risk_baixo):
    strong_points = [(k, v) for k, v in criteria_performance.items() if v >= META_GLOBAL]
    strong_points.sort(key=lambda x: x[1], reverse=True)
    
    if strong_points:
        positivos_text = "<b>O colaborador demonstra excelência e competência nas seguintes áreas:</b><br/><br/>"
        for i, (criterion, perf) in enumerate(strong_points[:5], 1):
            nivel = "Excelência" if perf >= 90 else "Bom"
            positivos_text += f"<b>{i}. {criterion}</b> - {round(perf)}% ({nivel})<br/>"
            if perf >= 90:
                positivos_text += f"""• Performance consistentemente acima das expectativas<br/>
• Pode servir como referência e mentor para outros colaboradores<br/>
• Manter este padrão e buscar oportunidades de compartilhar conhecimento<br/><br/>"""
            else:
                positivos_text += f"""• Atende aos padrões estabelecidos com consistência<br/>
• Continue desenvolvendo esta competência rumo à excelência<br/><br/>"""
        
        if satisfaction >= 70:
            positivos_text += f"""<b>Destaque Especial:</b> Taxa de satisfação do cliente de {round(satisfaction)}%, 
            demonstrando capacidade de gerar experiências positivas.<br/><br/>"""
        if risk_baixo >= 70:
            positivos_text += f"""<b>Gestão de Risco:</b> Excelente controle com {round(risk_baixo)}% de casos classificados como baixo risco, 
            evidenciando maturidade no tratamento de situações complexas.<br/><br/>"""
    else:
        positivos_text = """É importante reconhecer o esforço e dedicação do colaborador. 
        Mesmo em fase de desenvolvimento, há potencial a ser explorado com o treinamento adequado.<br/><br/>"""
    return positivos_text


def _build_pdi_text(avg_score, weak_points):
    pdi_text = f"""<b>Objetivo Geral:</b> Desenvolver competências para atingir e manter performance de excelência 
    (acima de {META_GLOBAL}% em todos os critérios).<br/><br/>"""
    
    pdi_text += "<b>Metas de Curto Prazo (30 dias):</b><br/>"
    if avg_score < META_GLOBAL:
        pdi_text += f"• Alcançar {round(avg_score + 10)}% de acerto médio através de treinamento intensivo<br/>"
        pdi_text += "• Participar de 4 sessões de coaching individual com o gestor<br/>"
        pdi_text += "• Realizar shadowing com colaborador de alta performance (mínimo 10 ligações)<br/>"
    else:
        pdi_text += f"• Manter performance acima de {META_GLOBAL}% com consistência<br/>"
        pdi_text += "• Identificar 2 áreas para aprimoramento rumo aos 95%<br/>"
        pdi_text += "• Participar como observador em treinamentos de novos colaboradores<br/>"
    
    if weak_points:
        pdi_text += f"• Focar no desenvolvimento prioritário de: {', '.join([wp[0] for wp in weak_points[:2]])}<br/>"
    
    pdi_text += "<br/><b>Ações de Médio Prazo (60-90 dias):</b><br/>"
    pdi_text += "• Certificação em técnicas avançadas de atendimento ao cliente<br/>"
    pdi_text += f"• Alcançar {META_GLOBAL}% de acerto médio em todos os critérios<br/>"
    pdi_text += "• Reduzir taxa de risco alto para menos de 5% dos atendimentos<br/>"
    pdi_text += "• Aumentar satisfação do cliente para acima de 80%<br/>"
    
    pdi_text += "<br/><b>Plano de Treinamento:</b><br/>"
    training_plan = []
    if weak_points:
        for criterion, _ in weak_points[:3]:
            if 'LGPD' in criterion:
                training_plan.append("• Curso online: Fundamentos da LGPD (4 horas)")
            elif 'Conhecimento' in criterion:
                training_plan.append("• Workshop: Catálogo de Produtos e Serviços (8 horas)")
            elif 'Comunicação' in criterion:
                training_plan.append("• Treinamento: Comunicação Eficaz no Atendimento (6 horas)")
            elif 'Escuta Ativa' in criterion:
                training_plan.append("• Workshop: Técnicas de Escuta Ativa (4 horas)")
            else:
                training_plan.append(f"• Módulo específico: {criterion} (4 horas)")
    training_plan.append("• Acompanhamento semanal com gestor (1 hora/semana)")
    training_plan.append("• Autoavaliação mensal de progresso")
    pdi_text += "<br/>".join(training_plan)
    
    pdi_text += "<br/><br/><b>Indicadores de Sucesso:</b><br/>"
    pdi_text += f"• Acerto médio acima de {META_GLOBAL}% em 3 meses<br/>"
    pdi_text += "• Zero casos de risco alto por mês<br/>"
    pdi_text += "• Satisfação do cliente acima de 85%<br/>"
    pdi_text += "• Feedback positivo do gestor em todas as revisões mensais<br/>"
    
    pdi_text += "<br/><b>Acompanhamento:</b><br/>"
    pdi_text += "• Reuniões de feedback: Semanais no primeiro mês, depois quinzenais<br/>"
    pdi_text += "• Revisão formal do PDI: Mensalmente<br/>"
    pdi_text += "• Análise de gravações: 2 ligações por semana com feedback detalhado<br/>"
    return pdi_text


def _get_pdf_styles():
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle', parent=styles['Heading1'],
        fontSize=24, textColor=HexColor(CARGLASS_RED),
        spaceAfter=30, alignment=TA_CENTER, fontName='Helvetica-Bold'
    )
    subtitle_style = ParagraphStyle(
        'CustomSubtitle', parent=styles['Heading2'],
        fontSize=16, textColor=HexColor(CARGLASS_DARK_RED),
        spaceAfter=15, spaceBefore=15, fontName='Helvetica-Bold'
    )
    section_style = ParagraphStyle(
        'SectionStyle', parent=styles['Heading3'],
        fontSize=13, textColor=HexColor(CARGLASS_RED),
        spaceAfter=10, spaceBefore=10, fontName='Helvetica-Bold'
    )
    normal_style = ParagraphStyle(
        'CustomNormal', parent=styles['Normal'],
        fontSize=10, textColor=colors.black,
        spaceAfter=8, leading=14, alignment=TA_LEFT
    )
    return title_style, subtitle_style, section_style, normal_style


def generate_manager_pdf(df, employee_name):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    title_style, subtitle_style, section_style, normal_style = _get_pdf_styles()
    
    elements.append(Paragraph(f"Relatório de Performance e Desenvolvimento - Análise Gestor", title_style))
    elements.append(Paragraph(f"Colaborador: {employee_name}", subtitle_style))
    elements.append(Paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y')}", normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    metrics = _compute_employee_metrics(df, employee_name)
    employee_df = metrics['employee_df']
    score_col = metrics['score_col']
    avg_score = metrics['avg_score']
    total_calls = metrics['total_calls']
    risk_baixo = metrics['risk_baixo']
    risk_alto = metrics['risk_alto']
    satisfaction = metrics['satisfaction']
    insatisfaction = metrics['insatisfaction']
    criteria_performance = metrics['criteria_performance']
    
    elements.append(Paragraph("Resumo Executivo de Performance", subtitle_style))
    elements.append(_build_summary_table(metrics))
    elements.append(Spacer(1, 0.3*inch))
    
    elements.append(Paragraph("1. Análise Qualitativa para Feedback do Gestor", subtitle_style))
    
    feedback_text = f"""
    <b>Orientações para o Gestor:</b><br/><br/>
    O colaborador {employee_name} apresenta uma performance {'acima' if avg_score >= META_GLOBAL else 'abaixo'} da meta estabelecida 
    de {META_GLOBAL}%, com um percentual médio de acerto de {round(avg_score)}% em {total_calls} ligações analisadas.<br/><br/>
    <b>Aspectos Comportamentais e Técnicos:</b><br/>
    """
    
    if avg_score >= 90:
        feedback_text += """• O colaborador demonstra excelência no atendimento e domínio das técnicas de comunicação.<br/>
    • Recomenda-se reconhecimento público e possível atuação como mentor para outros atendentes.<br/>
    • Utilize este colaborador como exemplo de boas práticas em treinamentos.<br/><br/>"""
    elif avg_score >= META_GLOBAL:
        feedback_text += """• O colaborador apresenta performance satisfatória, cumprindo os padrões estabelecidos.<br/>
    • Identificar oportunidades específicas de crescimento para alcançar o nível de excelência.<br/>
    • Feedback deve focar em refinamento e desenvolvimento de habilidades avançadas.<br/><br/>"""
    else:
        feedback_text += """• O colaborador necessita de atenção e acompanhamento próximo do gestor.<br/>
    • É fundamental estabelecer um plano de ação imediato com metas claras e alcançáveis.<br/>
    • Agendar sessões de feedback semanais para acompanhamento do progresso.<br/><br/>"""
    
    if risk_alto > 20:
        feedback_text += f"""<b>ATENÇÃO:</b> Taxa de risco alto em {round(risk_alto)}% dos atendimentos. 
    Priorizar treinamento em gestão de conflitos e técnicas de de-escalation.<br/><br/>"""
    
    if insatisfaction > 30:
        feedback_text += f"""<b>ATENÇÃO:</b> Taxa de insatisfação do cliente em {round(insatisfaction)}% dos casos. 
    Reforçar técnicas de empatia e resolução de problemas.<br/><br/>"""
    
    elements.append(Paragraph(feedback_text, normal_style))
    elements.append(Spacer(1, 0.2*inch))
    
    elements.append(Paragraph("2. Análise do Histórico Completo", subtitle_style))
    
    historico_text = f"""
    <b>Período Analisado:</b> {employee_df['AnalysisDateTime'].min().strftime('%d/%m/%Y') if 'AnalysisDateTime' in employee_df.columns and len(employee_df) > 0 else 'N/A'} 
    a {employee_df['AnalysisDateTime'].max().strftime('%d/%m/%Y') if 'AnalysisDateTime' in employee_df.columns and len(employee_df) > 0 else 'N/A'}<br/><br/>
    <b>Volume de Atendimentos:</b> O colaborador realizou {total_calls} atendimentos no período, 
    {'demonstrando consistência e volume adequado de trabalho' if total_calls >= 20 else 'com volume abaixo do esperado, sugerindo necessidade de aumento de produtividade'}.<br/><br/>
    <b>Padrões Identificados:</b><br/>
    """
    
    strong_areas = [k for k, v in criteria_performance.items() if v >= 90]
    good_areas = [k for k, v in criteria_performance.items() if META_GLOBAL <= v < 90]
    weak_areas = [k for k, v in criteria_performance.items() if v < META_GLOBAL]
    
    if strong_areas:
        historico_text += f"• <b>Áreas de Excelência:</b> {', '.join(strong_areas[:3])} - Demonstra domínio consistente.<br/>"
    if good_areas:
        historico_text += f"• <b>Áreas Satisfatórias:</b> {', '.join(good_areas[:3])} - Atende aos padrões estabelecidos.<br/>"
    if weak_areas:
        historico_text += f"• <b>Áreas Críticas:</b> {', '.join(weak_areas[:3])} - Requerem atenção imediata.<br/>"
    
    if 'AnalysisDateTime' in employee_df.columns and len(employee_df) >= 5:
        employee_df_sorted = employee_df.sort_values('AnalysisDateTime')
        first_half = employee_df_sorted.head(len(employee_df_sorted)//2)
        second_half = employee_df_sorted.tail(len(employee_df_sorted)//2)
        score_first = first_half[score_col].mean() if score_col == 'PERCENTUAL' else (first_half[score_col].mean() / 81) * 100
        score_second = second_half[score_col].mean() if score_col == 'PERCENTUAL' else (second_half[score_col].mean() / 81) * 100
        trend = score_second - score_first
        historico_text += f"<br/><b>Evolução Temporal:</b> "
        if trend > 5:
            historico_text += f"Tendência positiva detectada (+{round(trend, 1)}%). O colaborador está melhorando consistentemente.<br/>"
        elif trend < -5:
            historico_text += f"Tendência negativa detectada ({round(trend, 1)}%). Necessário investigar causas da queda de performance.<br/>"
        else:
            historico_text += f"Performance estável. Manter foco em consistência e buscar oportunidades de crescimento.<br/>"
    
    elements.append(Paragraph(historico_text, normal_style))
    elements.append(Spacer(1, 0.2*inch))
    
    elements.append(Paragraph("Performance Detalhada por Critério", section_style))
    elements.append(_build_criteria_table(criteria_performance))
    elements.append(PageBreak())
    
    elements.append(Paragraph("3. Pontos de Melhoria (Áreas Críticas)", subtitle_style))
    melhoria_text, weak_points = _build_melhoria_text(criteria_performance)
    elements.append(Paragraph(melhoria_text, normal_style))
    elements.append(Spacer(1, 0.2*inch))
    
    elements.append(Paragraph("4. Pontos Positivos (Forças Identificadas)", subtitle_style))
    positivos_text = _build_positivos_text(criteria_performance, satisfaction, risk_baixo)
    elements.append(Paragraph(positivos_text, normal_style))
    elements.append(PageBreak())
    
    elements.append(Paragraph("5. Plano de Desenvolvimento Individual (PDI)", subtitle_style))
    pdi_text = _build_pdi_text(avg_score, weak_points)
    elements.append(Paragraph(pdi_text, normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    if 'AnalysisDateTime' in employee_df.columns:
        elements.append(Paragraph("Histórico Recente de Atendimentos", section_style))
        recent_df = employee_df.sort_values('AnalysisDateTime', ascending=False).head(10)
        history_data = [['Data', 'Acerto (%)', 'Risco', 'Satisfação']]
        for _, row in recent_df.iterrows():
            date_str = row['AnalysisDateTime'].strftime('%d/%m/%Y')
            score = row[score_col] if score_col == 'PERCENTUAL' else (row[score_col]/81)*100
            risk = row.get('ClientRisk', 'N/A')
            client = row.get('Client', 'N/A')
            history_data.append([date_str, f'{round(score)}%', risk, client])
        
        history_table = Table(history_data, colWidths=[1.4*inch, 1.4*inch, 1.4*inch, 1.4*inch])
        history_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(CARGLASS_RED)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9)
        ]))
        elements.append(history_table)
    
    elements.append(Spacer(1, 0.5*inch))
    elements.append(Paragraph("_" * 80, normal_style))
    signature_data = [
        ['_____________________________', '_____________________________'],
        ['Assinatura do Colaborador', 'Assinatura do Gestor'],
        ['', ''],
        ['Data: ____/____/________', 'Data: ____/____/________']
    ]
    signature_table = Table(signature_data, colWidths=[3*inch, 3*inch])
    signature_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 8)
    ]))
    elements.append(signature_table)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer


def generate_employee_pdf(df, employee_name):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    title_style, subtitle_style, section_style, normal_style = _get_pdf_styles()
    
    elements.append(Paragraph(f"Relatório de Performance e Desenvolvimento", title_style))
    elements.append(Paragraph(f"Colaborador: {employee_name}", subtitle_style))
    elements.append(Paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y')}", normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    metrics = _compute_employee_metrics(df, employee_name)
    avg_score = metrics['avg_score']
    satisfaction = metrics['satisfaction']
    risk_baixo = metrics['risk_baixo']
    criteria_performance = metrics['criteria_performance']
    
    elements.append(Paragraph("Resumo Executivo de Performance", subtitle_style))
    elements.append(_build_summary_table(metrics))
    elements.append(Spacer(1, 0.3*inch))
    
    elements.append(Paragraph("1. Pontos de Melhoria (Áreas Críticas)", subtitle_style))
    melhoria_text, weak_points = _build_melhoria_text(criteria_performance)
    elements.append(Paragraph(melhoria_text, normal_style))
    elements.append(Spacer(1, 0.2*inch))
    
    elements.append(Paragraph("2. Pontos Positivos (Forças Identificadas)", subtitle_style))
    positivos_text = _build_positivos_text(criteria_performance, satisfaction, risk_baixo)
    elements.append(Paragraph(positivos_text, normal_style))
    elements.append(PageBreak())
    
    elements.append(Paragraph("3. Plano de Desenvolvimento Individual (PDI)", subtitle_style))
    pdi_text = _build_pdi_text(avg_score, weak_points)
    elements.append(Paragraph(pdi_text, normal_style))
    elements.append(Spacer(1, 0.5*inch))
    
    elements.append(Paragraph("_" * 80, normal_style))
    signature_data = [
        ['_____________________________', '_____________________________'],
        ['Assinatura do Colaborador', 'Assinatura do Gestor'],
        ['', ''],
        ['Data: ____/____/________', 'Data: ____/____/________']
    ]
    signature_table = Table(signature_data, colWidths=[3*inch, 3*inch])
    signature_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 8)
    ]))
    elements.append(signature_table)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer


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
    
    bar_colors = [CARGLASS_GREEN if p >= META_GLOBAL else CARGLASS_ORANGE if p >= 60 else CARGLASS_RED for p in performance]
    
    fig = go.Figure(data=[
        go.Bar(
            x=question_labels,
            y=performance,
            marker=dict(color=bar_colors, line=dict(color='white', width=2)),
            text=[f'{round(p)}%' for p in performance],
            textposition='outside',
            textfont=dict(size=12, color=CHART_TEXT_COLOR, family='Inter', weight='bold'),
            hovertemplate='<b>%{customdata}</b><br>Acerto: %{y:.0f}%<extra></extra>',
            customdata=[QUESTION_DETAILS.get(f'Question{i+1}', '') for i in range(12)]
        )
    ])
    
    fig.update_layout(
        title={
            'text': '✅ Performance do Checklist por Critério',
            'font': {'size': 20, 'color': CHART_TEXT_COLOR, 'family': 'Inter'},
            'x': 0.5, 'xanchor': 'center'
        },
        xaxis=dict(
            title=dict(text='Critérios de Avaliação', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
            tickfont=dict(size=12, color=CHART_TEXT_COLOR, family='Inter')
        ),
        yaxis=dict(
            range=[0, 110],
            title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
            tickfont=dict(size=12, color=CHART_TEXT_COLOR, family='Inter'),
            gridcolor=CHART_GRID_COLOR
        ),
        plot_bgcolor=CHART_BG,
        paper_bgcolor=CHART_PAPER_BG,
        height=450,
        showlegend=False,
        font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
        margin=dict(l=60, r=40, t=80, b=60)
    )
    
    fig.add_hline(
        y=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN, line_width=2,
        annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="right",
        annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
    )
    
    return fig


def create_risk_baixo_alto_chart(df):
    if 'ClientRisk' not in df.columns:
        return None
    
    risk_counts = df['ClientRisk'].value_counts()
    total_records = len(df)
    if len(risk_counts) == 0:
        return None
    
    risk_order = ['ALTO', 'MEDIO', 'BAIXO']
    risk_mapping = {
        'BAIXO': ('Risco Baixo', CARGLASS_GREEN),
        'MEDIO': ('Risco Médio', CARGLASS_YELLOW),
        'ALTO': ('Risco Alto', CARGLASS_RED)
    }
    
    labels = []
    values = []
    colors_list = []
    percentages = []
    hover_texts = []
    
    for risk in risk_order:
        if risk in risk_counts.index:
            count = risk_counts[risk]
            label, color = risk_mapping[risk]
            labels.append(label)
            values.append(count)
            colors_list.append(color)
            pct = (count / total_records * 100) if total_records > 0 else 0
            percentages.append(pct)
            hover_texts.append(RISK_DESCRIPTIONS.get(risk, ''))
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=labels, x=values, orientation='h',
        marker=dict(color=colors_list, line=dict(color='white', width=2)),
        text=[f"{val} ({pct:.1f}%)" for val, pct in zip(values, percentages)],
        textposition='outside',
        textfont=dict(size=14, color=CHART_TEXT_COLOR, family='Inter', weight='bold'),
        hovertemplate='<b>%{y}</b><br>Quantidade: %{x}<br>Percentual: %{customdata[0]:.1f}%<br><br>%{customdata[1]}<extra></extra>',
        customdata=list(zip(percentages, hover_texts))
    ))
    
    total_shown = sum(values)
    fig.update_layout(
        title={
            'text': f'⚠️ Distribuição de Risco (Total: {total_shown} registros)',
            'font': {'size': 18, 'color': CHART_TEXT_COLOR, 'family': 'Inter'},
            'x': 0.5, 'xanchor': 'center'
        },
        xaxis=dict(
            title=dict(text='Quantidade de Registros', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
            tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
            showgrid=True, gridcolor=CHART_GRID_COLOR
        ),
        yaxis=dict(tickfont=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
        height=350,
        paper_bgcolor=CHART_PAPER_BG,
        plot_bgcolor=CHART_BG,
        font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
        showlegend=False,
        margin=dict(l=120, r=150, t=80, b=60)
    )
    
    return fig


def create_agent_ranking(df, top_n=5):
    score_column = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    if 'CustomerAgent' not in df.columns or score_column not in df.columns:
        return None
    
    agent_scores = df.groupby('CustomerAgent').agg({
        score_column: 'mean',
        'IdAnalysis': 'count'
    }).round(1)
    
    if score_column == 'NOTAS':
        agent_scores[score_column] = (agent_scores[score_column] / 81) * 100
    
    agent_scores.columns = ['Porcentagem Média', 'Total Ligações']
    agent_scores = agent_scores.sort_values('Porcentagem Média', ascending=False).head(top_n)
    
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
        x=agent_scores['Porcentagem Média'], y=agent_names, orientation='h',
        marker=dict(color=colors_agents, line=dict(color='white', width=2)),
        text=[f"{round(score)}% ({calls} lig.)"
              for score, calls in zip(agent_scores['Porcentagem Média'], agent_scores['Total Ligações'])],
        textposition='outside',
        textfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter', weight='bold'),
        hovertemplate='<b>%{y}</b><br>Acerto: %{x:.0f}%<extra></extra>'
    ))
    
    fig.update_layout(
        title={
            'text': '👥 Top 5 Agentes por Porcentagem de Acerto',
            'font': {'size': 18, 'color': CHART_TEXT_COLOR, 'family': 'Inter'},
            'x': 0.5, 'xanchor': 'center'
        },
        xaxis=dict(
            range=[0, 110],
            title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
            tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
            gridcolor=CHART_GRID_COLOR
        ),
        yaxis=dict(title='', tickfont=dict(size=12, color=CHART_TEXT_COLOR, family='Inter', weight='bold')),
        height=350, showlegend=False,
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER_BG,
        font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
        margin=dict(l=150, r=120, t=70, b=50)
    )
    
    fig.add_vline(
        x=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN, line_width=2,
        annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="top",
        annotation_font=dict(size=11, color=CARGLASS_GREEN, family='Inter')
    )
    
    return fig


def create_bottom_performers(df, bottom_n=5):
    score_column = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    if 'CustomerAgent' not in df.columns or score_column not in df.columns:
        return None
    
    agent_scores = df.groupby('CustomerAgent').agg({
        score_column: 'mean',
        'IdAnalysis': 'count'
    }).round(1)
    
    if score_column == 'NOTAS':
        agent_scores[score_column] = (agent_scores[score_column] / 81) * 100
    
    agent_scores.columns = ['Porcentagem Média', 'Total Ligações']
    agent_scores = agent_scores[agent_scores['Total Ligações'] >= 10]
    agent_scores = agent_scores.sort_values('Porcentagem Média', ascending=True).head(bottom_n)
    
    agent_names = []
    for name in agent_scores.index:
        parts = name.split()
        if len(parts) >= 2:
            agent_names.append(f"{parts[0]} {parts[-1]}")
        else:
            agent_names.append(parts[0])
    
    colors_agents = [CARGLASS_RED if i == 0 else CARGLASS_ORANGE for i in range(len(agent_names))]
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=agent_scores['Porcentagem Média'], y=agent_names, orientation='h',
        marker=dict(color=colors_agents, line=dict(color='white', width=2)),
        text=[f"{round(score)}% ({calls} lig.)"
              for score, calls in zip(agent_scores['Porcentagem Média'], agent_scores['Total Ligações'])],
        textposition='outside',
        textfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
        hovertemplate='<b>%{y}</b><br>Acerto: %{x:.0f}%<extra></extra>'
    ))
    
    fig.update_layout(
        title={
            'text': '⚠️ Necessitam Treinamento',
            'font': {'size': 18, 'color': CHART_TEXT_COLOR, 'family': 'Inter'},
            'x': 0.5, 'xanchor': 'center'
        },
        xaxis=dict(
            title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
            range=[0, 110],
            tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
            gridcolor=CHART_GRID_COLOR
        ),
        yaxis=dict(title='', tickfont=dict(size=12, color=CHART_TEXT_COLOR, family='Inter')),
        height=350, showlegend=False,
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER_BG,
        font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
        margin=dict(l=150, r=120, t=70, b=50)
    )
    
    fig.add_vline(
        x=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN, line_width=2,
        annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="top",
        annotation_font=dict(size=11, color=CARGLASS_GREEN, family='Inter')
    )
    
    return fig


def create_timeline_chart(df):
    score_column = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    if 'AnalysisDateTime' not in df.columns or score_column not in df.columns:
        return None
    
    try:
        if len(df) == 0:
            return None
        
        df_timeline = df.set_index('AnalysisDateTime').resample('D')[score_column].agg(['mean', 'count']).reset_index()
        df_timeline.columns = ['Data', 'Porcentagem Média', 'Quantidade']
        
        if score_column == 'NOTAS':
            df_timeline['Porcentagem Média'] = (df_timeline['Porcentagem Média'] / 81) * 100
        
        df_timeline = df_timeline.dropna()
        df_timeline = df_timeline[df_timeline['Quantidade'] > 0]
        
        if len(df_timeline) == 0:
            return None
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=df_timeline['Data'],
            y=df_timeline['Porcentagem Média'],
            mode='lines+markers+text',
            name='Porcentagem de Acerto',
            line=dict(color=CARGLASS_RED, width=3),
            marker=dict(size=10, color=CARGLASS_RED, line=dict(color='white', width=2)),
            text=[f'{round(v)}%' for v in df_timeline['Porcentagem Média']],
            textposition='top center',
            textfont=dict(size=9, color=CHART_TEXT_COLOR, family='Inter', weight='bold'),
            yaxis='y',
            hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Acerto: %{y:.0f}%<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            x=df_timeline['Data'],
            y=df_timeline['Quantidade'],
            name='Quantidade de Análises',
            marker=dict(color=CARGLASS_BLUE, line=dict(color='white', width=1)),
            opacity=0.4,
            yaxis='y2',
            text=[str(int(q)) for q in df_timeline['Quantidade']],
            textposition='outside',
            textfont=dict(size=8, color=CARGLASS_BLUE, family='Inter'),
            hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Análises: %{y}<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': '📈 Evolução Temporal da Porcentagem de Acerto',
                'font': {'size': 20, 'color': CHART_TEXT_COLOR, 'family': 'Inter'},
                'x': 0.5, 'xanchor': 'center'
            },
            xaxis=dict(
                title=dict(text='Data', font=dict(size=14, color=CHART_TEXT_COLOR, family='Inter')),
                tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
                type='category',
                tickformat='%d/%m',
                tickvals=df_timeline['Data'],
                ticktext=[d.strftime('%d/%m') for d in df_timeline['Data']],
                tickangle=-45
            ),
            yaxis=dict(
                title=dict(text='Porcentagem de Acerto', font=dict(color=CARGLASS_RED, size=13, family='Inter')),
                tickfont=dict(color=CARGLASS_RED, size=11, family='Inter'),
                side='left', range=[0, 115],
                gridcolor=CHART_GRID_COLOR
            ),
            yaxis2=dict(
                title=dict(text='Quantidade de Análises', font=dict(color=CARGLASS_BLUE, size=13, family='Inter')),
                tickfont=dict(color=CARGLASS_BLUE, size=11, family='Inter'),
                overlaying='y', side='right'
            ),
            height=500,
            plot_bgcolor=CHART_BG,
            paper_bgcolor=CHART_PAPER_BG,
            hovermode='x unified',
            font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                font=dict(size=11, family='Inter', color=CHART_TEXT_COLOR)
            ),
            margin=dict(l=60, r=60, t=80, b=100)
        )
        
        fig.add_hline(
            y=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN, line_width=2,
            annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="left",
            annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
        )
        
        return fig
    except Exception as e:
        st.warning(f"Não foi possível criar o gráfico de evolução temporal: {str(e)}")
        return None


def create_improvement_points(df):
    weak_questions = []
    for i in range(1, 13):
        q = f'Question{i}'
        if q in df.columns:
            performance = df[q].mean() * 100
            if performance < META_GLOBAL:
                weak_questions.append((q, performance))
    weak_questions.sort(key=lambda x: x[1])
    return [(QUESTION_NAMES.get(q, q), perf) for q, perf in weak_questions[:3]]


def create_company_comparison(df):
    if 'Empresas' not in df.columns or 'PERCENTUAL' not in df.columns:
        return None, None
    
    company_stats = df.groupby('Empresas').agg({
        'PERCENTUAL': 'mean',
        'ClientRisk': lambda x: (x == 'BAIXO').sum() / len(x) * 100 if len(x) > 0 else 0
    }).round(1)
    
    company_stats['Total Análises'] = df.groupby('Empresas').size()
    company_stats.columns = ['Porcentagem Média', '% Risco Baixo', 'Total Análises']
    company_stats = company_stats[['Porcentagem Média', 'Total Análises', '% Risco Baixo']]
    company_stats = company_stats.sort_values('Porcentagem Média', ascending=False)
    
    total_in_chart = int(company_stats['Total Análises'].sum())
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Porcentagem de Acerto Média',
        x=company_stats.index,
        y=company_stats['Porcentagem Média'],
        marker_color=CARGLASS_RED,
        text=[f"{round(row['Porcentagem Média'])}%<br>({int(row['Total Análises'])} análises)"
              for _, row in company_stats.iterrows()],
        textposition='outside',
        textfont=dict(color=CHART_TEXT_COLOR, size=11, family='Inter'),
        hovertemplate='<b>%{x}</b><br>Acerto: %{y:.0f}%<br>Análises: %{customdata}<extra></extra>',
        customdata=company_stats['Total Análises']
    ))
    
    fig.update_layout(
        title={
            'text': f'📊 Comparativo de Performance por Cliente ({total_in_chart} análises)',
            'font': {'size': 18, 'color': CHART_TEXT_COLOR, 'family': 'Inter'},
            'x': 0.5, 'xanchor': 'center'
        },
        xaxis_title='Cliente',
        yaxis_title='Porcentagem de Acerto',
        xaxis=dict(tickfont=dict(color=CHART_TEXT_COLOR, size=11, family='Inter')),
        yaxis=dict(
            range=[0, max(company_stats['Porcentagem Média']) * 1.15],
            tickfont=dict(color=CHART_TEXT_COLOR, size=11, family='Inter'),
            gridcolor=CHART_GRID_COLOR
        ),
        height=400, showlegend=False,
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER_BG,
        font={'color': CHART_TEXT_COLOR, 'family': 'Inter'}
    )
    
    fig.add_hline(
        y=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN,
        annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="right",
        annotation_font=dict(color=CARGLASS_GREEN)
    )
    
    return fig, company_stats


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
            
            if 'Empresas' in df.columns:
                empresas = ['Todas'] + sorted(df['Empresas'].dropna().unique().tolist())
                selected_empresa = st.selectbox(
                    "🏢 Cliente", empresas,
                    help="Filtrar por cliente específico"
                )
                if selected_empresa != 'Todas':
                    df = df[df['Empresas'] == selected_empresa]
            
            if 'AnalysisDateTime' in df.columns:
                min_date = df['AnalysisDateTime'].min().date()
                max_date = df['AnalysisDateTime'].max().date()
                date_range = st.date_input(
                    "📅 Período de Análise",
                    value=(min_date, max_date),
                    min_value=min_date, max_value=max_date
                )
                if len(date_range) == 2:
                    df = df[(df['AnalysisDateTime'].dt.date >= date_range[0]) &
                           (df['AnalysisDateTime'].dt.date <= date_range[1])]
            
            if 'CustomerAgent' in df.columns:
                agents = ['Todos'] + sorted(df['CustomerAgent'].dropna().unique().tolist())
                selected_agent = st.selectbox("👤 Agente", agents)
                if selected_agent != 'Todos':
                    df = df[df['CustomerAgent'] == selected_agent]
            
            if 'ClientRisk' in df.columns:
                risks = ['Todos'] + sorted(df['ClientRisk'].dropna().unique().tolist())
                selected_risk = st.selectbox("⚠️ Nível de Risco", risks)
                if selected_risk != 'Todos':
                    df = df[df['ClientRisk'] == selected_risk]
            
            st.markdown("---")
            if 'PERCENTUAL' in df.columns:
                st.info("📊 Usando porcentagem de acerto (0-100%)")
            else:
                st.warning("⚠️ Coluna 'Avaliação 100 pts' não encontrada. Convertendo NOTAS para porcentagem.")
            
            st.markdown("---")
            st.markdown("### 📄 Relatório Individual")
            
            if 'CustomerAgent' in df.columns:
                agents_for_pdf = sorted(df['CustomerAgent'].dropna().unique().tolist())
                selected_agent_pdf = st.selectbox("Selecione o Colaborador", agents_for_pdf, key="pdf_agent")
                
                st.caption("Relatório do Colaborador: versão resumida, sem dados sensíveis de análise do gestor.")
                if st.button("📄 Gerar Relatório do Colaborador", use_container_width=True):
                    pdf_buffer = generate_employee_pdf(df, selected_agent_pdf)
                    st.download_button(
                        label="💾 Download PDF (Colaborador)", data=pdf_buffer,
                        file_name=f"relatorio_colaborador_{selected_agent_pdf.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf", use_container_width=True, key="dl_employee_pdf"
                    )
                
                st.caption("Análise Gestor: relatório completo com análise qualitativa e histórico, uso interno.")
                if st.button("📄 Gerar Análise Gestor", use_container_width=True):
                    manager_pdf_buffer = generate_manager_pdf(df, selected_agent_pdf)
                    st.download_button(
                        label="💾 Download PDF (Gestor)", data=manager_pdf_buffer,
                        file_name=f"analise_gestor_{selected_agent_pdf.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf", use_container_width=True, key="dl_manager_pdf"
                    )
            
            st.markdown("---")
            st.markdown("### 💾 Exportar Dados")
            if st.button("📊 Gerar Relatório Excel", use_container_width=True):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Dados Filtrados', index=False)
                output.seek(0)
                st.download_button(
                    label="💾 Download Excel", data=output,
                    file_name=f"monitoria_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    else:
        df = None
        st.info("👆 Carregue um arquivo para começar")


st.markdown("""
<div class='header-gradient'>
    <h1>🔴 Monitor AI</h1>
    <p>Dashboard de Indicadores de Gestão - Análise de Atendimento</p>
</div>
""", unsafe_allow_html=True)

if df is not None and len(df) > 0:
    
    total_analyses = len(df)
    
    if 'PERCENTUAL' in df.columns:
        avg_score = df['PERCENTUAL'].mean()
    else:
        avg_score = (df['NOTAS'].mean() / 81) * 100 if 'NOTAS' in df.columns else 0
    
    low_risk_pct = (df['ClientRisk'] == 'BAIXO').sum() / len(df) * 100 if 'ClientRisk' in df.columns else 0
    
    if 'CustomerAgent' in df.columns:
        avg_per_agent = round(total_analyses / df['CustomerAgent'].nunique(), 1)
    else:
        avg_per_agent = 0
    
    best_q, best_q_val, worst_q, worst_q_val = get_best_worst_questions(df)
    
    if 'AnalysisDateTime' in df.columns:
        last_week = df[df['AnalysisDateTime'] >= (datetime.now() - timedelta(days=7))]
        week_count = len(last_week)
        week_delta = f"📈 +{week_count} esta semana"
    else:
        week_delta = ""
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Total de Análises</div>
            <div class='kpi-value'>{total_analyses:,}</div>
            <div class='kpi-delta'>{week_delta}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        delta_icon = "⚠️" if avg_score < META_GLOBAL else "✅"
        delta_text = f"{delta_icon} {'Abaixo' if avg_score < META_GLOBAL else 'Acima'} da meta ({META_GLOBAL}%)"
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Porcentagem de Acerto</div>
            <div class='kpi-value'>{round(avg_score)}%</div>
            <div class='kpi-delta'>{delta_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        risk_icon = "✅" if low_risk_pct >= 60 else "⚠️"
        risk_text = f"{risk_icon} {'Dentro' if low_risk_pct >= 60 else 'Abaixo'} do objetivo"
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Risco Baixo</div>
            <div class='kpi-value'>{round(low_risk_pct)}%</div>
            <div class='kpi-delta'>{risk_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        if best_q:
            st.markdown(f"""
            <div class='kpi-card-green'>
                <div class='kpi-label'>Melhor Desempenho</div>
                <div class='kpi-value'>{best_q_val}%</div>
                <div class='kpi-delta'>🏆 {best_q}</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col5:
        if worst_q:
            st.markdown(f"""
            <div class='kpi-card-red'>
                <div class='kpi-label'>Pior Desempenho</div>
                <div class='kpi-value'>{worst_q_val}%</div>
                <div class='kpi-delta'>📉 {worst_q}</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style='display: flex; justify-content: center; margin-bottom: 20px;'>
        <div style='background: linear-gradient(135deg, #2C5AA0 0%, #4A90E2 100%); padding: 15px 40px; border-radius: 12px; 
            box-shadow: 0 4px 12px rgba(0,0,0,0.1); text-align: center;'>
            <span style='color: rgba(255,255,255,0.85); font-size: 12px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600;'>
                Média de Monitorias por Agente
            </span>
            <span style='color: white; font-size: 28px; font-weight: 700; margin-left: 15px; text-shadow: 1px 1px 2px rgba(0,0,0,0.2);'>
                {avg_per_agent}
            </span>
            <span style='color: rgba(255,255,255,0.8); font-size: 12px; margin-left: 8px;'>
                ({df['CustomerAgent'].nunique() if 'CustomerAgent' in df.columns else 0} agentes)
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if 'Empresas' in df.columns:
        st.markdown("## 🏢 Análise por Cliente")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            company_chart, company_stats = create_company_comparison(df)
            if company_chart:
                st.plotly_chart(company_chart, use_container_width=True)
        
        with col2:
            if company_stats is not None:
                st.markdown("### 📋 Estatísticas por Cliente")
                media_row = pd.DataFrame({
                    'Porcentagem Média': [company_stats['Porcentagem Média'].mean()],
                    'Total Análises': [company_stats['Total Análises'].sum()],
                    '% Risco Baixo': [company_stats['% Risco Baixo'].mean()]
                }, index=['MÉDIA GERAL'])
                company_stats_with_avg = pd.concat([company_stats, media_row])
                st.dataframe(
                    company_stats_with_avg.style.format({
                        'Porcentagem Média': '{:.0f}%',
                        '% Risco Baixo': '{:.0f}%'
                    }).apply(lambda x: ['background-color: #FFF3CD; font-weight: bold' if x.name == 'MÉDIA GERAL' else '' for _ in x], axis=1),
                    use_container_width=True
                )
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class='risk-legend'>
            <b>ℹ️ Legenda de Risco:</b><br/>
            🟢 <b>Baixo:</b> Atendimento dentro dos padrões, sem problemas identificados.<br/>
            🟡 <b>Médio:</b> Pontos de atenção que podem gerar insatisfação moderada.<br/>
            🔴 <b>Alto:</b> Falhas críticas com risco de reclamação formal ou perda de cliente.
        </div>
        """, unsafe_allow_html=True)
        risk_comparison_chart = create_risk_baixo_alto_chart(df)
        if risk_comparison_chart:
            st.plotly_chart(risk_comparison_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        performance_chart = create_performance_chart(df)
        if performance_chart:
            st.plotly_chart(performance_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1.5, 1.5, 1])
    
    with col1:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        agent_ranking = create_agent_ranking(df)
        if agent_ranking:
            st.plotly_chart(agent_ranking, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        bottom_chart = create_bottom_performers(df)
        if bottom_chart:
            st.plotly_chart(bottom_chart, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color: " + CARGLASS_DARK_RED + "; font-size: 18px; margin-bottom: 20px;'>🎯 Pontos de Melhoria</h3>", unsafe_allow_html=True)
        improvement_points = create_improvement_points(df)
        if improvement_points:
            for q_name, perf in improvement_points:
                if perf < 50:
                    color = CARGLASS_RED
                    icon = "🔴"
                else:
                    color = CARGLASS_ORANGE
                    icon = "🟠"
                st.markdown(f"""
                <div style='margin-bottom: 15px; padding: 12px; background: #F8F9FA; border-radius: 8px; border-left: 4px solid {color};'>
                    <div style='font-weight: 600; color: {CHART_TEXT_COLOR}; font-size: 13px;'>{icon} {q_name}</div>
                    <div style='color: {color}; font-weight: bold; font-size: 16px; margin-top: 5px;'>{round(perf)}%</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.success(f"✅ Todos os critérios estão acima da meta de {META_GLOBAL}%!")
        st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='content-card'>", unsafe_allow_html=True)
    timeline = create_timeline_chart(df)
    if timeline:
        st.plotly_chart(timeline, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("## 🚨 Análise de Casos de Risco Alto")
    
    st.markdown("""
    <div class='risk-section-card'>
        <div class='risk-legend'>
            <b>🔴 O que é Risco Alto?</b><br/>
            O risco alto indica que o atendimento apresentou falhas críticas que podem resultar em reclamações formais, 
            perda de cliente ou impacto negativo na imagem da empresa. Esta seção permite filtrar por cliente para 
            identificar padrões e agir preventivamente.
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if 'ClientRisk' in df.columns:
        high_risk_df_full = df[df['ClientRisk'] == 'ALTO'].copy()
        
        if not high_risk_df_full.empty:
            col_filter1, col_filter2, col_filter3 = st.columns([2, 2, 1])
            
            with col_filter1:
                if 'Empresas' in high_risk_df_full.columns:
                    risk_clientes = ['Todos os Clientes'] + sorted(high_risk_df_full['Empresas'].dropna().unique().tolist())
                    selected_risk_cliente = st.selectbox(
                        "🏢 Filtrar por Cliente",
                        risk_clientes,
                        key="risk_client_filter"
                    )
                else:
                    selected_risk_cliente = 'Todos os Clientes'
            
            with col_filter2:
                if 'CustomerAgent' in high_risk_df_full.columns:
                    risk_agents = ['Todos os Agentes'] + sorted(high_risk_df_full['CustomerAgent'].dropna().unique().tolist())
                    selected_risk_agent = st.selectbox(
                        "👤 Filtrar por Agente",
                        risk_agents,
                        key="risk_agent_filter"
                    )
                else:
                    selected_risk_agent = 'Todos os Agentes'
            
            filtered_risk = high_risk_df_full.copy()
            if selected_risk_cliente != 'Todos os Clientes' and 'Empresas' in filtered_risk.columns:
                filtered_risk = filtered_risk[filtered_risk['Empresas'] == selected_risk_cliente]
            if selected_risk_agent != 'Todos os Agentes' and 'CustomerAgent' in filtered_risk.columns:
                filtered_risk = filtered_risk[filtered_risk['CustomerAgent'] == selected_risk_agent]
            
            with col_filter3:
                st.markdown(f"""
                <div style='background: #DC0A0A; color: white; padding: 15px; border-radius: 10px; text-align: center; margin-top: 25px;'>
                    <div style='font-size: 24px; font-weight: 700;'>{len(filtered_risk)}</div>
                    <div style='font-size: 11px; text-transform: uppercase; letter-spacing: 1px;'>Casos de Risco Alto</div>
                </div>
                """, unsafe_allow_html=True)
            
            risk_display_cols = []
            col_rename = {}
            
            if 'AnalysisDateTime' in filtered_risk.columns:
                risk_display_cols.append('AnalysisDateTime')
                col_rename['AnalysisDateTime'] = 'Data'
            if 'Empresas' in filtered_risk.columns:
                risk_display_cols.append('Empresas')
                col_rename['Empresas'] = 'Cliente'
            if 'CustomerAgent' in filtered_risk.columns:
                risk_display_cols.append('CustomerAgent')
                col_rename['CustomerAgent'] = 'Agente'
            if 'Mp3FileName' in filtered_risk.columns:
                risk_display_cols.append('Mp3FileName')
                col_rename['Mp3FileName'] = 'Gravação (MP3)'
            if 'Justification' in filtered_risk.columns:
                risk_display_cols.append('Justification')
                col_rename['Justification'] = 'Justificativa'
            if 'PERCENTUAL' in filtered_risk.columns:
                risk_display_cols.append('PERCENTUAL')
                col_rename['PERCENTUAL'] = 'Acerto %'
            
            if risk_display_cols:
                risk_display = filtered_risk[risk_display_cols].copy()
                if 'AnalysisDateTime' in risk_display.columns:
                    risk_display = risk_display.sort_values('AnalysisDateTime', ascending=False)
                risk_display = risk_display.rename(columns=col_rename)
                if 'Acerto %' in risk_display.columns:
                    risk_display['Acerto %'] = risk_display['Acerto %'].round(0).astype(int).astype(str) + '%'
                st.dataframe(risk_display, use_container_width=True, hide_index=True, height=400)
            
            if selected_risk_cliente == 'Todos os Clientes' and 'Empresas' in high_risk_df_full.columns:
                st.markdown("#### 📊 Distribuição de Risco Alto por Cliente")
                risk_by_client = high_risk_df_full.groupby('Empresas').size().reset_index(name='Casos')
                risk_by_client = risk_by_client.sort_values('Casos', ascending=False)
                
                total_risk = risk_by_client['Casos'].sum()
                risk_by_client['% do Total'] = (risk_by_client['Casos'] / total_risk * 100).round(1)
                
                fig_risk_client = go.Figure(go.Bar(
                    x=risk_by_client['Empresas'],
                    y=risk_by_client['Casos'],
                    marker_color=CARGLASS_RED,
                    text=[f"{c} ({p}%)" for c, p in zip(risk_by_client['Casos'], risk_by_client['% do Total'])],
                    textposition='outside',
                    textfont=dict(color=CHART_TEXT_COLOR, size=11, family='Inter', weight='bold'),
                    hovertemplate='<b>%{x}</b><br>Casos: %{y}<br>%{text}<extra></extra>'
                ))
                fig_risk_client.update_layout(
                    xaxis=dict(title='Cliente', tickfont=dict(color=CHART_TEXT_COLOR, size=11), tickangle=-45),
                    yaxis=dict(title='Número de Casos', tickfont=dict(color=CHART_TEXT_COLOR, size=11), gridcolor=CHART_GRID_COLOR),
                    height=400,
                    plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER_BG,
                    font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
                    margin=dict(l=60, r=40, t=40, b=100)
                )
                st.plotly_chart(fig_risk_client, use_container_width=True)
        else:
            st.success("✅ Nenhum caso de risco alto identificado no período selecionado!")
    else:
        st.warning("⚠️ Coluna ClientRisk não encontrada no arquivo.")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("## 📊 Análise Detalhada por Agente")
    
    tab1, tab2, tab3 = st.tabs(["📈 Performance Individual", "🎯 Comparativo", "📝 Detalhes"])
    
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
                score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
                score_val = agent_df[score_col].mean() if score_col == 'PERCENTUAL' else (agent_df[score_col].mean() / 81) * 100
                st.metric("Acerto Médio", f"{round(score_val)}%")
            with col3:
                risk_baixo = (agent_df['ClientRisk'] == 'BAIXO').sum() / len(agent_df) * 100 if 'ClientRisk' in agent_df.columns else 0
                st.metric("Risco Baixo", f"{round(risk_baixo)}%")
            with col4:
                agent_df_copy = agent_df.copy()
                agent_df_copy['Client_Cluster'] = agent_df_copy['Client'].apply(get_satisfaction_cluster) if 'Client' in agent_df_copy.columns else None
                satisfaction = (agent_df_copy['Client_Cluster'] == 'SATISFEITO').sum() / len(agent_df_copy) * 100 if 'Client' in agent_df_copy.columns else 0
                st.metric("Satisfação", f"{round(satisfaction)}%")
            
            questions_performance = []
            for i in range(1, 13):
                q = f'Question{i}'
                if q in agent_df.columns:
                    perf = agent_df[q].mean() * 100
                    questions_performance.append({
                        'Critério': QUESTION_NAMES.get(q, q),
                        'Performance': round(perf)
                    })
            
            if questions_performance:
                perf_df = pd.DataFrame(questions_performance)
                
                fig = go.Figure(go.Bar(
                    x=perf_df['Performance'], y=perf_df['Critério'], orientation='h',
                    marker=dict(
                        color=[CARGLASS_GREEN if p >= META_GLOBAL else CARGLASS_ORANGE if p >= 60 else CARGLASS_RED
                               for p in perf_df['Performance']],
                        line=dict(color='white', width=2)
                    ),
                    text=[f'{p}%' for p in perf_df['Performance']],
                    textposition='outside',
                    textfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter', weight='bold')
                ))
                
                for color, label in [(CARGLASS_GREEN, f'≥ {META_GLOBAL}% (Meta)'), (CARGLASS_ORANGE, f'60-{META_GLOBAL-1}%'), (CARGLASS_RED, '< 60%')]:
                    fig.add_trace(go.Bar(
                        x=[None], y=[None], orientation='h',
                        marker=dict(color=color),
                        name=label, showlegend=True
                    ))
                
                fig.update_layout(
                    title=f'Performance de {selected_agent}',
                    xaxis=dict(
                        range=[0, 110],
                        title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
                        tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
                        gridcolor=CHART_GRID_COLOR
                    ),
                    yaxis=dict(tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter')),
                    height=450,
                    plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER_BG,
                    font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
                    margin=dict(l=150, r=80, t=60, b=60),
                    legend=dict(
                        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                        font=dict(size=11, family='Inter', color=CHART_TEXT_COLOR)
                    )
                )
                
                fig.add_vline(
                    x=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN, line_width=2,
                    annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="top",
                    annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with tab2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        
        score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
        
        if 'CustomerAgent' in df.columns and score_col in df.columns:
            agent_comparison = df.groupby('CustomerAgent').agg({
                score_col: 'mean',
                'IdAnalysis': 'count',
                'ClientRisk': lambda x: (x == 'BAIXO').sum() / len(x) * 100 if len(x) > 0 else 0
            }).round(1)
            
            if score_col == 'NOTAS':
                agent_comparison[score_col] = (agent_comparison[score_col] / 81) * 100
            
            agent_comparison.columns = ['Porcentagem Média', 'Total Ligações', '% Risco Baixo']
            agent_comparison['Porcentagem Média'] = agent_comparison['Porcentagem Média'].round(0).astype(int)
            agent_comparison['% Risco Baixo'] = agent_comparison['% Risco Baixo'].round(0).astype(int)
            agent_comparison = agent_comparison.sort_values('Porcentagem Média', ascending=False)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=agent_comparison['Total Ligações'],
                y=agent_comparison['Porcentagem Média'],
                mode='markers+text',
                marker=dict(
                    size=agent_comparison['% Risco Baixo'] * 0.8,
                    color=agent_comparison['Porcentagem Média'],
                    colorscale=[[0, CARGLASS_RED], [0.5, CARGLASS_YELLOW], [1, CARGLASS_GREEN]],
                    showscale=True,
                    colorbar=dict(
                        title=dict(text="Acerto (%)", font=dict(size=11, family='Inter', color=CHART_TEXT_COLOR)),
                        tickfont=dict(size=10, family='Inter', color=CHART_TEXT_COLOR)
                    ),
                    line=dict(width=2, color='white')
                ),
                text=[name.split()[0] for name in agent_comparison.index],
                textposition='top center',
                textfont=dict(size=9, color=CHART_TEXT_COLOR, family='Inter'),
                hovertemplate='<b>%{text}</b><br>Ligações: %{x}<br>Acerto: %{y}%<br>Risco Baixo: %{marker.size:.0f}%<extra></extra>'
            ))
            
            fig.add_annotation(
                x=0.01, y=0.99, xref='paper', yref='paper',
                text="Tamanho = % Risco Baixo | Cor = % Acerto",
                showarrow=False,
                font=dict(size=10, color=CHART_TEXT_COLOR, family='Inter'),
                bgcolor='rgba(255,255,255,0.8)',
                bordercolor=CARGLASS_GRAY,
                borderwidth=1, borderpad=4
            )
            
            fig.update_layout(
                title='Análise Comparativa de Agentes',
                xaxis=dict(
                    title=dict(text='Total de Ligações', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
                    tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
                    gridcolor=CHART_GRID_COLOR
                ),
                yaxis=dict(
                    title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CHART_TEXT_COLOR, family='Inter')),
                    tickfont=dict(size=11, color=CHART_TEXT_COLOR, family='Inter'),
                    range=[0, 110], gridcolor=CHART_GRID_COLOR
                ),
                height=500,
                plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER_BG,
                font={'color': CHART_TEXT_COLOR, 'family': 'Inter'},
                margin=dict(l=60, r=100, t=60, b=60)
            )
            
            fig.add_hline(
                y=META_GLOBAL, line_dash="dash", line_color=CARGLASS_GREEN, line_width=2,
                annotation_text=f"Meta: {META_GLOBAL}%", annotation_position="right",
                annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(
                agent_comparison.style.format({'Porcentagem Média': '{}%', '% Risco Baixo': '{}%'}),
                use_container_width=True, height=400
            )
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with tab3:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        st.markdown("### 📋 Dados Detalhados")
        
        display_columns = ['AnalysisDateTime', 'CustomerAgent', 'Client', 'ClientRisk', 'ClientOutcome']
        if 'Empresas' in df.columns:
            display_columns.insert(1, 'Empresas')
        if 'PERCENTUAL' in df.columns:
            display_columns.append('PERCENTUAL')
        elif 'NOTAS' in df.columns:
            display_columns.append('NOTAS')
        
        available_columns = [col for col in display_columns if col in df.columns]
        
        if available_columns:
            df_display = df[available_columns].sort_values('AnalysisDateTime', ascending=False).head(100).copy()
            column_rename = {
                'PERCENTUAL': 'Acerto %',
                'AnalysisDateTime': 'Data Análise',
                'CustomerAgent': 'Agente',
                'ClientRisk': 'Risco',
                'ClientOutcome': 'Resultado',
                'Empresas': 'Cliente'
            }
            df_display = df_display.rename(columns=column_rename)
            if 'Acerto %' in df_display.columns:
                df_display['Acerto %'] = df_display['Acerto %'].round(0).astype(int)
            st.dataframe(df_display, use_container_width=True, hide_index=True, height=400)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📊 Estatísticas Gerais")
            score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
            if score_col == 'NOTAS':
                avg_val = (df[score_col].mean() / 81) * 100
                med_val = (df[score_col].median() / 81) * 100
                std_val = (df[score_col].std() / 81) * 100
                min_val = (df[score_col].min() / 81) * 100
                max_val = (df[score_col].max() / 81) * 100
            else:
                avg_val = df[score_col].mean()
                med_val = df[score_col].median()
                std_val = df[score_col].std()
                min_val = df[score_col].min()
                max_val = df[score_col].max()
            
            stats_df = pd.DataFrame({
                'Métrica': ['Acerto Médio', 'Acerto Mediano', 'Desvio Padrão', 'Acerto Mínimo', 'Acerto Máximo'],
                'Valor': [f"{round(avg_val)}%", f"{round(med_val)}%", f"{round(std_val)}%", f"{round(min_val)}%", f"{round(max_val)}%"]
            })
            st.dataframe(stats_df, use_container_width=True, hide_index=True)
        
        with col2:
            st.markdown("### 🏆 Rankings")
            if 'CustomerAgent' in df.columns:
                score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
                if score_col == 'NOTAS':
                    agent_scores = df.groupby('CustomerAgent')[score_col].mean()
                    agent_scores = (agent_scores / 81) * 100
                else:
                    agent_scores = df.groupby('CustomerAgent')[score_col].mean()
                
                best_agents = agent_scores.sort_values(ascending=False).head(3)
                worst_agents = agent_scores.sort_values(ascending=True).head(3)
                
                st.markdown("**Top 3 Melhores:**")
                for i, (agent, score) in enumerate(best_agents.items(), 1):
                    medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉"
                    st.markdown(f"{medal} {agent}: {round(score)}%")
                
                st.markdown("<br>**3 Para Melhorar:**", unsafe_allow_html=True)
                for agent, score in worst_agents.items():
                    st.markdown(f"📈 {agent}: {round(score)}%")
        
        st.markdown("</div>", unsafe_allow_html=True)

else:
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
