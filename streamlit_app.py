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


def get_satisfaction_cluster(value):
    """Retorna o cluster de satisfação para um valor"""
    if pd.isna(value):
        return None
    
    value_upper = str(value).strip().upper()
    
    # Satisfeito (positivo)
    if value_upper in ['ALTA', 'ALTO', 'BOA', 'SATISFEITO', 'SATISFEITA']:
        return 'SATISFEITO'
    
    # Neutro
    if value_upper in ['NEUTRA', 'NEUTRO', 'MÉDIA', 'MEDIA', 'MÉDIO', 'MEDIO', 'MODERADA', 'MODERADO']:
        return 'NEUTRO'
    
    # Insatisfeito (negativo)
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
</style>
"""

st.markdown(custom_css, unsafe_allow_html=True)

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
            
            # Filtrar registros com PERCENTUAL vazio ou menor que 19.99
            if 'PERCENTUAL' in df.columns:
                df = df[(df['PERCENTUAL'].notna()) & (df['PERCENTUAL'] >= 19.99)]
            
            # Filtrar registros com ClientRisk Indeterminado ou vazio
            if 'ClientRisk' in df.columns:
                df = df[(df['ClientRisk'] != 'INDETERMINADO') & (df['ClientRisk'].notna())]
            
            # Filtrar registros com Empresas vazio (para consistência)
            if 'Empresas' in df.columns:
                df = df[df['Empresas'].notna()]
            
            return df
        else:
            st.error("A planilha 'Consulta1' não foi encontrada no arquivo.")
            return None
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return None

def generate_employee_pdf(df, employee_name):
    """
    Gera relatório PDF completo do colaborador com:
    1. Análise qualitativa para feedback do gestor
    2. Análise do histórico completo
    3. Pontos de melhoria
    4. Pontos positivos
    5. Plano de desenvolvimento individual
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Estilos customizados
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=HexColor(CARGLASS_RED),
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=HexColor(CARGLASS_DARK_RED),
        spaceAfter=15,
        spaceBefore=15,
        fontName='Helvetica-Bold'
    )
    
    section_style = ParagraphStyle(
        'SectionStyle',
        parent=styles['Heading3'],
        fontSize=13,
        textColor=HexColor(CARGLASS_RED),
        spaceAfter=10,
        spaceBefore=10,
        fontName='Helvetica-Bold'
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.black,
        spaceAfter=8,
        leading=14,
        alignment=TA_LEFT
    )
    
    # Cabeçalho
    elements.append(Paragraph(f"Relatório de Performance e Desenvolvimento", title_style))
    elements.append(Paragraph(f"Colaborador: {employee_name}", subtitle_style))
    elements.append(Paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y')}", normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    # Filtrar dados do colaborador
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
    
    # Resumo executivo
    elements.append(Paragraph("Resumo Executivo de Performance", subtitle_style))
    
    summary_data = [
        ['Métrica', 'Valor', 'Status'],
        ['Porcentagem de Acerto Média', f'{round(avg_score)}%', '✓ Bom' if avg_score >= 70 else '✗ Abaixo da Meta'],
        ['Total de Ligações Analisadas', str(total_calls), '-'],
        ['Taxa de Risco Baixo', f'{round(risk_baixo)}%', '✓ Bom' if risk_baixo >= 60 else '✗ Atenção'],
        ['Taxa de Satisfação do Cliente', f'{round(satisfaction)}%', '✓ Bom' if satisfaction >= 70 else '✗ Atenção']
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
    
    elements.append(summary_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # ========== 1. ANÁLISE QUALITATIVA PARA FEEDBACK DO GESTOR ==========
    elements.append(Paragraph("1. Análise Qualitativa para Feedback do Gestor", subtitle_style))
    
    feedback_text = f"""
    <b>Orientações para o Gestor:</b><br/><br/>
    
    O colaborador {employee_name} apresenta uma performance {'acima' if avg_score >= 70 else 'abaixo'} da meta estabelecida 
    de 70%, com um percentual médio de acerto de {round(avg_score)}% em {total_calls} ligações analisadas.<br/><br/>
    
    <b>Aspectos Comportamentais e Técnicos:</b><br/>
    """
    
    if avg_score >= 85:
        feedback_text += """• O colaborador demonstra excelência no atendimento e domínio das técnicas de comunicação.<br/>
    • Recomenda-se reconhecimento público e possível atuação como mentor para outros atendentes.<br/>
    • Utilize este colaborador como exemplo de boas práticas em treinamentos.<br/><br/>"""
    elif avg_score >= 70:
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
    
    # ========== 2. ANÁLISE DO HISTÓRICO COMPLETO ==========
    elements.append(Paragraph("2. Análise do Histórico Completo", subtitle_style))
    
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
    
    # Calcular performance por critério
    criteria_performance = {}
    for i in range(1, 13):
        q = f'Question{i}'
        if q in employee_df.columns:
            perf = employee_df[q].mean() * 100
            criteria_performance[question_names.get(q, q)] = perf
    
    # Análise textual do histórico
    historico_text = f"""
    <b>Período Analisado:</b> {employee_df['AnalysisDateTime'].min().strftime('%d/%m/%Y') if 'AnalysisDateTime' in employee_df.columns else 'N/A'} 
    a {employee_df['AnalysisDateTime'].max().strftime('%d/%m/%Y') if 'AnalysisDateTime' in employee_df.columns else 'N/A'}<br/><br/>
    
    <b>Volume de Atendimentos:</b> O colaborador realizou {total_calls} atendimentos no período, 
    {'demonstrando consistência e volume adequado de trabalho' if total_calls >= 20 else 'com volume abaixo do esperado, sugerindo necessidade de aumento de produtividade'}.<br/><br/>
    
    <b>Padrões Identificados:</b><br/>
    """
    
    # Identificar tendências
    strong_areas = [k for k, v in criteria_performance.items() if v >= 85]
    good_areas = [k for k, v in criteria_performance.items() if 70 <= v < 85]
    weak_areas = [k for k, v in criteria_performance.items() if v < 70]
    
    if strong_areas:
        historico_text += f"• <b>Áreas de Excelência:</b> {', '.join(strong_areas[:3])} - Demonstra domínio consistente.<br/>"
    if good_areas:
        historico_text += f"• <b>Áreas Satisfatórias:</b> {', '.join(good_areas[:3])} - Atende aos padrões estabelecidos.<br/>"
    if weak_areas:
        historico_text += f"• <b>Áreas Críticas:</b> {', '.join(weak_areas[:3])} - Requerem atenção imediata.<br/>"
    
    # Análise de evolução se houver dados temporais
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
    
    # Tabela de performance detalhada
    elements.append(Paragraph("Performance Detalhada por Critério", section_style))
    
    criteria_data = [['Critério', 'Performance', 'Status']]
    for criterion, perf in sorted(criteria_performance.items(), key=lambda x: x[1], reverse=True):
        status = '✓ Excelente' if perf >= 85 else '✓ Bom' if perf >= 70 else '✗ Precisa Melhorar'
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
    
    elements.append(criteria_table)
    elements.append(PageBreak())
    
    # ========== 3. PONTOS DE MELHORIA ==========
    elements.append(Paragraph("3. Pontos de Melhoria (Áreas Críticas)", subtitle_style))
    
    weak_points = [(k, v) for k, v in criteria_performance.items() if v < 70]
    weak_points.sort(key=lambda x: x[1])
    
    if weak_points:
        melhoria_text = "<b>Os seguintes pontos requerem atenção e desenvolvimento prioritário:</b><br/><br/>"
        
        for i, (criterion, perf) in enumerate(weak_points, 1):
            melhoria_text += f"<b>{i}. {criterion}</b> ({round(perf)}% de acerto)<br/>"
            
            # Recomendações específicas por critério
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
        melhoria_text = """<b>Parabéns!</b> Todos os critérios estão acima da meta de 70%. 
        Foco agora deve ser em refinamento e excelência em todas as áreas.<br/><br/>"""
    
    elements.append(Paragraph(melhoria_text, normal_style))
    elements.append(Spacer(1, 0.2*inch))
    
    # ========== 4. PONTOS POSITIVOS ==========
    elements.append(Paragraph("4. Pontos Positivos (Forças Identificadas)", subtitle_style))
    
    strong_points = [(k, v) for k, v in criteria_performance.items() if v >= 70]
    strong_points.sort(key=lambda x: x[1], reverse=True)
    
    if strong_points:
        positivos_text = "<b>O colaborador demonstra excelência e competência nas seguintes áreas:</b><br/><br/>"
        
        for i, (criterion, perf) in enumerate(strong_points[:5], 1):
            nivel = "Excelência" if perf >= 85 else "Bom"
            positivos_text += f"<b>{i}. {criterion}</b> - {round(perf)}% ({nivel})<br/>"
            
            if perf >= 85:
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
    
    elements.append(Paragraph(positivos_text, normal_style))
    elements.append(PageBreak())
    
    # ========== 5. PLANO DE DESENVOLVIMENTO ==========
    elements.append(Paragraph("5. Plano de Desenvolvimento Individual (PDI)", subtitle_style))
    
    pdi_text = """<b>Objetivo Geral:</b> Desenvolver competências para atingir e manter performance de excelência 
    (acima de 85% em todos os critérios).<br/><br/>"""
    
    # Metas de curto prazo (30 dias)
    pdi_text += "<b>Metas de Curto Prazo (30 dias):</b><br/>"
    
    if avg_score < 70:
        pdi_text += f"• Alcançar {round(avg_score + 10)}% de acerto médio através de treinamento intensivo<br/>"
        pdi_text += "• Participar de 4 sessões de coaching individual com o gestor<br/>"
        pdi_text += "• Realizar shadowing com colaborador de alta performance (mínimo 10 ligações)<br/>"
    else:
        pdi_text += "• Manter performance acima de 70% com consistência<br/>"
        pdi_text += "• Identificar 2 áreas para aprimoramento rumo aos 85%<br/>"
        pdi_text += "• Participar como observador em treinamentos de novos colaboradores<br/>"
    
    if weak_points:
        pdi_text += f"• Focar no desenvolvimento prioritário de: {', '.join([wp[0] for wp in weak_points[:2]])}<br/>"
    
    pdi_text += "<br/><b>Ações de Médio Prazo (60-90 dias):</b><br/>"
    pdi_text += "• Certificação em técnicas avançadas de atendimento ao cliente<br/>"
    pdi_text += "• Alcançar 85% de acerto médio em todos os critérios<br/>"
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
    pdi_text += "• Acerto médio acima de 85% em 3 meses<br/>"
    pdi_text += "• Zero casos de risco alto por mês<br/>"
    pdi_text += "• Satisfação do cliente acima de 85%<br/>"
    pdi_text += "• Feedback positivo do gestor em todas as revisões mensais<br/>"
    
    pdi_text += "<br/><b>Acompanhamento:</b><br/>"
    pdi_text += "• Reuniões de feedback: Semanais no primeiro mês, depois quinzenais<br/>"
    pdi_text += "• Revisão formal do PDI: Mensalmente<br/>"
    pdi_text += "• Análise de gravações: 2 ligações por semana com feedback detalhado<br/>"
    
    elements.append(Paragraph(pdi_text, normal_style))
    elements.append(Spacer(1, 0.3*inch))
    
    # Histórico recente
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
    
    # Assinaturas
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
    
    # Gerar PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

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
            text=[f'{round(p)}%' for p in performance],
            textposition='outside',
            textfont=dict(size=12, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{x}</b><br>Acerto: %{y:.0f}%<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title={
            'text': '✅ Performance do Checklist por Critério',
            'font': {'size': 22, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            'x': 0.5,
            'xanchor': 'center'
        },
        xaxis=dict(
            title=dict(text='Critérios de Avaliação', font=dict(size=14, color=CARGLASS_GRAY, family='Inter')),
            tickfont=dict(size=12, color=CARGLASS_GRAY, family='Inter')
        ),
        yaxis=dict(
            range=[0, 110],
            title=dict(text='Porcentagem de Acerto', font=dict(size=14, color=CARGLASS_GRAY, family='Inter')),
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

def create_satisfaction_donut(df):
    # Mapeamento de clusterização - agrupa categorias similares
    cluster_map = {
        # Satisfeito (verde)
        'ALTA': 'Satisfeito',
        'ALTO': 'Satisfeito',
        'BOA': 'Satisfeito',
        'SATISFEITO': 'Satisfeito',
        'SATISFEITA': 'Satisfeito',
        
        # Neutro (amarelo)
        'NEUTRA': 'Neutro',
        'NEUTRO': 'Neutro',
        'MÉDIA': 'Neutro',
        'MEDIA': 'Neutro',
        'MÉDIO': 'Neutro',
        'MEDIO': 'Neutro',
        'MODERADA': 'Neutro',
        'MODERADO': 'Neutro',
        
        # Insatisfeito (vermelho)
        'BAIXA': 'Insatisfeito',
        'BAIXO': 'Insatisfeito',
        'INSATISFEITO': 'Insatisfeito',
        'INSATISFEITA': 'Insatisfeito',
        'INSATISFATÓRIO': 'Insatisfeito',
        'INSATISFATORIA': 'Insatisfeito',
    }
    
    # Cores para os clusters
    color_map = {
        'Satisfeito': '#28A745',  # Verde
        'Neutro': '#FFC107',      # Amarelo
        'Insatisfeito': '#DC0A0A', # Vermelho
        'Indeterminado': '#6C757D' # Cinza
    }
    
    if 'Client' in df.columns:
        # Criar coluna de cluster
        df_copy = df.copy()
        df_copy['Client_Cluster'] = df_copy['Client'].apply(
            lambda x: cluster_map.get(str(x).strip().upper(), 'Indeterminado') if pd.notna(x) and len(str(x)) < 50 else 'Indeterminado'
        )
        
        # Remover categoria Indeterminado se houver
        df_copy = df_copy[df_copy['Client_Cluster'] != 'Indeterminado']
        
        # Contar clusters
        cluster_counts = df_copy['Client_Cluster'].value_counts()
        
        labels = cluster_counts.index.tolist()
        values = cluster_counts.values
        colors_list = [color_map.get(label, '#6C757D') for label in labels]
        
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
            textfont=dict(size=13, color=CARGLASS_DARK_RED, family='Inter'),
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title={
                'text': '😊 Distribuição de Satisfação do Cliente',
                'font': {'size': 22, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
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


def create_risk_baixo_alto_chart(df):
    """Cria gráfico de barras horizontais mostrando distribuição de Risco (Baixo, Médio e Alto)"""
    if 'ClientRisk' in df.columns:
        # Usar todos os riscos (Baixo, Médio e Alto)
        risk_counts = df['ClientRisk'].value_counts()
        
        # Total de registros
        total_records = len(df)
        
        if len(risk_counts) == 0:
            return None
        
        # Ordem inversa para exibir: Alto no topo, Baixo embaixo
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
        
        for risk in risk_order:
            if risk in risk_counts.index:
                count = risk_counts[risk]
                label, color = risk_mapping[risk]
                labels.append(label)
                values.append(count)
                colors_list.append(color)
                pct = (count / total_records * 100) if total_records > 0 else 0
                percentages.append(pct)
        
        # Criar gráfico de barras horizontais
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            y=labels,
            x=values,
            orientation='h',
            marker=dict(
                color=colors_list,
                line=dict(color='white', width=2)
            ),
            text=[f"{val} ({pct:.1f}%)" for val, pct in zip(values, percentages)],
            textposition='outside',
            textfont=dict(size=14, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{y}</b><br>Quantidade: %{x}<br>Percentual: %{customdata:.1f}%<extra></extra>',
            customdata=percentages
        ))
        
        # Calcular total mostrado no gráfico
        total_shown = sum(values)
        
        fig.update_layout(
            title={
                'text': f'⚠️ Distribuição de Risco (Total: {total_shown} registros)',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title=dict(text='Quantidade de Registros', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
                tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter'),
                showgrid=True,
                gridcolor='#E8E8E8'
            ),
            yaxis=dict(
                tickfont=dict(size=13, color=CARGLASS_DARK_RED, family='Inter')
            ),
            height=350,
            paper_bgcolor='white',
            plot_bgcolor='#FAFBFC',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            showlegend=False,
            margin=dict(l=120, r=150, t=80, b=60)
        )
        
        return fig
    return None



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
                risk_labels.append(f'{risk.capitalize()}\n{count} casos ({round(pct)}%)')
                
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
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title='',
                tickfont=dict(size=12, color=CARGLASS_GRAY, family='Inter')
            ),
            yaxis=dict(
                title=dict(text='Número de Casos', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
                tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter'),
                range=[0, max(risk_data) * 1.25]  # Adicionar espaço para o texto acima das barras
            ),
            height=450,
            showlegend=False,
            plot_bgcolor='#FAFBFC',
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            margin=dict(l=60, r=40, t=70, b=50)
        )
        
        return fig
    return None

def create_agent_ranking(df, top_n=5):
    score_column = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    
    if 'CustomerAgent' in df.columns and score_column in df.columns:
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
            x=agent_scores['Porcentagem Média'],
            y=agent_names,
            orientation='h',
            marker=dict(
                color=colors_agents,
                line=dict(color='white', width=2)
            ),
            text=[f"{round(score)}%<br>{calls} ligações" 
                  for score, calls in zip(agent_scores['Porcentagem Média'], agent_scores['Total Ligações'])],
            textposition='outside',
            textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter', weight='bold'),
            hovertemplate='<b>%{y}</b><br>Acerto: %{round(x)}%<br>Ligações: %{text}<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': '👥 Top 5 Agentes por Porcentagem de Acerto',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                range=[0, 110],
                title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
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

def create_bottom_performers(df, bottom_n=5):
    score_column = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    
    if 'CustomerAgent' in df.columns and score_column in df.columns:
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
            x=agent_scores['Porcentagem Média'],
            y=agent_names,
            orientation='h',
            marker=dict(
                color=colors_agents,
                line=dict(color='white', width=2)
            ),
            text=[f"{round(score)}%<br>{calls} ligações" 
                  for score, calls in zip(agent_scores['Porcentagem Média'], agent_scores['Total Ligações'])],
            textposition='outside',
            textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter'),
            hovertemplate='<b>%{y}</b><br>Acerto: %{round(x)}%<br>Ligações: %{text}<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': '⚠️ Necessitam Treinamento',
                'font': {'size': 20, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
                range=[0, 110],
                tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
            ),
            yaxis=dict(
                title=dict(text='', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
                tickfont=dict(size=12, color=CARGLASS_DARK_RED, family='Inter')
            ),
            height=350,
            showlegend=False,
            plot_bgcolor='#FAFBFC',
            paper_bgcolor='white',
            font={'color': CARGLASS_DARK_RED, 'family': 'Inter'},
            margin=dict(l=150, r=100, t=70, b=50)
        )
        
        fig.add_vline(
            x=70, 
            line_dash="dash", 
            line_color=CARGLASS_GREEN,
            line_width=2,
            annotation_text="Meta: 70%", 
            annotation_position="top",
            annotation_font=dict(size=11, color=CARGLASS_GREEN, family='Inter')
        )
        
        return fig
    return None

def create_timeline_chart(df):
    score_column = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
    
    if 'AnalysisDateTime' in df.columns and score_column in df.columns:
        try:
            if len(df) == 0:
                return None
            
            df_timeline = df.set_index('AnalysisDateTime').resample('D')[score_column].agg(['mean', 'count']).reset_index()
            df_timeline.columns = ['Data', 'Porcentagem Média', 'Quantidade']
            
            if score_column == 'NOTAS':
                df_timeline['Porcentagem Média'] = (df_timeline['Porcentagem Média'] / 81) * 100
            
            df_timeline = df_timeline.dropna()
            
            if len(df_timeline) == 0:
                return None
            
            fig = go.Figure()
            
            fig.add_trace(go.Scatter(
                x=df_timeline['Data'],
                y=df_timeline['Porcentagem Média'],
                mode='lines+markers',
                name='Porcentagem de Acerto',
                line=dict(color=CARGLASS_RED, width=3),
                marker=dict(size=8, color=CARGLASS_RED, line=dict(color='white', width=2)),
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
                hovertemplate='<b>Data: %{x|%d/%m/%Y}</b><br>Análises: %{y}<extra></extra>'
            ))
            
            fig.update_layout(
                title={
                    'text': '📈 Evolução Temporal da Porcentagem de Acerto',
                    'font': {'size': 22, 'color': CARGLASS_DARK_RED, 'family': 'Inter'},
                    'x': 0.5,
                    'xanchor': 'center'
                },
                xaxis=dict(
                    title=dict(text='Data', font=dict(size=14, color=CARGLASS_GRAY, family='Inter')),
                    tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
                ),
                yaxis=dict(
                    title=dict(text='Porcentagem de Acerto', font=dict(color=CARGLASS_RED, size=13, family='Inter')),
                    tickfont=dict(color=CARGLASS_RED, size=11, family='Inter'),
                    side='left',
                    range=[0, 110]
                ),
                yaxis2=dict(
                    title=dict(text='Quantidade de Análises', font=dict(color=CARGLASS_BLUE, size=13, family='Inter')),
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
            
            fig.add_hline(
                y=70, 
                line_dash="dash", 
                line_color=CARGLASS_GREEN,
                line_width=2,
                annotation_text="Meta: 70%", 
                annotation_position="left",
                annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
            )
            
            return fig
        except Exception as e:
            st.warning(f"Não foi possível criar o gráfico de evolução temporal: {str(e)}")
            return None
    return None

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

def create_company_comparison(df):
    if 'Empresas' in df.columns and 'PERCENTUAL' in df.columns:
        # DataFrame já vem filtrado do load_data
        company_stats = df.groupby('Empresas').agg({
            'PERCENTUAL': 'mean',
            'ClientRisk': lambda x: (x == 'BAIXO').sum() / len(x) * 100 if len(x) > 0 else 0
        }).round(1)
        
        # Adicionar contagem usando size() que conta TODOS os registros (não ignora NaN)
        company_stats['Total Análises'] = df.groupby('Empresas').size()
        
        company_stats.columns = ['Porcentagem Média', '% Risco Baixo', 'Total Análises']
        # Reordenar colunas
        company_stats = company_stats[['Porcentagem Média', 'Total Análises', '% Risco Baixo']]
        company_stats = company_stats.sort_values('Porcentagem Média', ascending=False)
        
        # Total de análises no gráfico
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
            hovertemplate='<b>%{x}</b><br>Acerto: %{y:.0f}%<br>Análises: %{customdata}<extra></extra>',
            customdata=company_stats['Total Análises']
        ))
        
        fig.update_layout(
            title=f'📊 Comparativo de Performance por Empresa<br><sub style="font-size:12px;">({total_in_chart} análises)</sub>',
            xaxis_title='Empresa',
            yaxis_title='Porcentagem de Acerto',
            yaxis_range=[0, max(company_stats['Porcentagem Média']) * 1.1],
            height=400,
            showlegend=False,
            plot_bgcolor='#FAFBFC',
            paper_bgcolor='white'
        )
        
        fig.add_hline(y=70, line_dash="dash", line_color=CARGLASS_GREEN, 
                     annotation_text="Meta: 70%", annotation_position="right")
        
        return fig, company_stats
    return None, None

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
                    "🏢 Empresa",
                    empresas,
                    help="Filtrar por empresa específica"
                )
                
                if selected_empresa != 'Todas':
                    df = df[df['Empresas'] == selected_empresa]
            
            if 'AnalysisDateTime' in df.columns:
                min_date = df['AnalysisDateTime'].min().date()
                max_date = df['AnalysisDateTime'].max().date()
                
                date_range = st.date_input(
                    "📅 Período de Análise",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date
                )
                
                if len(date_range) == 2:
                    df = df[(df['AnalysisDateTime'].dt.date >= date_range[0]) & 
                           (df['AnalysisDateTime'].dt.date <= date_range[1])]
            
            if 'CustomerAgent' in df.columns:
                agents = ['Todos'] + sorted(df['CustomerAgent'].dropna().unique().tolist())
                selected_agent = st.selectbox(
                    "👤 Agente",
                    agents
                )
                
                if selected_agent != 'Todos':
                    df = df[df['CustomerAgent'] == selected_agent]
            
            if 'ClientRisk' in df.columns:
                risks = ['Todos'] + sorted(df['ClientRisk'].dropna().unique().tolist())
                selected_risk = st.selectbox(
                    "⚠️ Nível de Risco",
                    risks
                )
                
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
                selected_agent_pdf = st.selectbox(
                    "Selecione o Colaborador",
                    agents_for_pdf,
                    key="pdf_agent"
                )
                
                if st.button("📄 Gerar Relatório PDF", use_container_width=True):
                    pdf_buffer = generate_employee_pdf(df, selected_agent_pdf)
                    st.download_button(
                        label="💾 Download PDF",
                        data=pdf_buffer,
                        file_name=f"relatorio_{selected_agent_pdf.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            
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

st.markdown("""
<div class='header-gradient'>
    <h1>🔴 Monitor AI</h1>
    <p>Dashboard de Indicadores de Gestão - Análise de Atendimento</p>
</div>
""", unsafe_allow_html=True)

if df is not None and len(df) > 0:
    
    col1, col2, col3, col4 = st.columns(4)
    
    # Total de análises - usar len(df) direto após filtros
    # A tabela mostra a soma dos registros por empresa, que deve ser igual a len(df)
    total_analyses = len(df)
    
    if 'PERCENTUAL' in df.columns:
        avg_score = df['PERCENTUAL'].mean()
    else:
        avg_score = (df['NOTAS'].mean() / 81) * 100 if 'NOTAS' in df.columns else 0
    
    low_risk_pct = (df['ClientRisk'] == 'BAIXO').sum() / len(df) * 100 if 'ClientRisk' in df.columns else 0
    
    # Taxa de Saudação (Question1) - mesmo cálculo do gráfico
    saudacao_pct = 0
    if 'Question1' in df.columns and len(df) > 0:
        # Garantir que estamos usando o mesmo cálculo do gráfico de performance
        saudacao_pct = df['Question1'].mean() * 100
    else:
        # Se Question1 não existir, tentar calcular de outra forma
        saudacao_pct = 0
    
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
        delta_text = f"{delta_icon} {'Abaixo' if avg_score < 70 else 'Acima'} da meta (70%)"
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
        sat_icon = "🎯" if saudacao_pct >= 80 else "📊"
        sat_text = f"{sat_icon} {'Excelente' if saudacao_pct >= 80 else 'Bom'} desempenho"
        st.markdown(f"""
        <div class='kpi-card-modern'>
            <div class='kpi-label'>Taxa Saudação</div>
            <div class='kpi-value'>{round(saudacao_pct)}%</div>
            <div class='kpi-delta'>{sat_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if 'Empresas' in df.columns:
        st.markdown("## 🏢 Análise por Empresa")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            company_chart, company_stats = create_company_comparison(df)
            if company_chart:
                st.plotly_chart(company_chart, use_container_width=True)
        
        with col2:
            if company_stats is not None:
                st.markdown("### 📋 Estatísticas por Empresa")
                
                # Adicionar linha de média
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
        # Gráfico de Risco Baixo vs Alto (removido Satisfação)
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
    
    col1, col2, col3, col4 = st.columns([1.2, 1.5, 1.2, 1.1])
    
    with col1:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        # Tabela de casos de Risco Alto (gráfico removido)
        st.markdown("<h3 style='color: #DC0A0A; margin-bottom: 15px;'>🚨 Casos de Risco Alto</h3>", unsafe_allow_html=True)
        
        if 'ClientRisk' in df.columns:
            required_cols = ['Mp3FileName', 'Justification', 'CustomerAgent']
            if all(col in df.columns for col in required_cols):
                high_risk_df = df[df['ClientRisk'] == 'ALTO'][required_cols].copy()
                if not high_risk_df.empty:
                    high_risk_df.columns = ['Gravação (MP3)', 'Justificativa', 'Agente']
                    st.dataframe(high_risk_df, use_container_width=True, hide_index=True, height=400)
                else:
                    st.success("✅ Nenhum caso de risco alto identificado!")
            else:
                st.warning("⚠️ Colunas necessárias não encontradas no arquivo.")
        else:
            st.warning("⚠️ Coluna ClientRisk não encontrada no arquivo.")
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        agent_ranking = create_agent_ranking(df)
        if agent_ranking:
            st.plotly_chart(agent_ranking, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown("<div class='content-card'>", unsafe_allow_html=True)
        bottom_chart = create_bottom_performers(df)
        if bottom_chart:
            st.plotly_chart(bottom_chart, use_container_width=True)
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
                <div style='color: {color}; font-weight: bold; font-size: 16px; margin-top: 5px;'>{round(perf)}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='content-card'>", unsafe_allow_html=True)
    timeline = create_timeline_chart(df)
    if timeline:
        st.plotly_chart(timeline, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
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
                score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
                if score_col == 'NOTAS':
                    score_val = (agent_df[score_col].mean() / 81) * 100
                else:
                    score_val = agent_df[score_col].mean()
                st.metric("Acerto Médio", f"{round(score_val)}%")
            with col3:
                risk_baixo = (agent_df['ClientRisk'] == 'BAIXO').sum() / len(agent_df) * 100 if 'ClientRisk' in agent_df.columns else 0
                st.metric("Risco Baixo", f"{round(risk_baixo)}%")
            with col4:
                agent_df['Client_Cluster'] = agent_df['Client'].apply(get_satisfaction_cluster) if 'Client' in agent_df.columns else None
                satisfaction = (agent_df['Client_Cluster'] == 'SATISFEITO').sum() / len(agent_df) * 100 if 'Client' in agent_df.columns else 0
                st.metric("Satisfação", f"{round(satisfaction)}%")
            
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
                    text=[f'{round(p)}%' for p in perf_df['Performance']],
                    textposition='outside',
                    textfont=dict(size=11, color=CARGLASS_DARK_RED, family='Inter', weight='bold')
                ))
                
                fig.update_layout(
                    title=f'Performance de {selected_agent}',
                    xaxis=dict(
                        range=[0, 110],
                        title=dict(text='Porcentagem de Acerto', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
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
        
        score_col = 'PERCENTUAL' if 'PERCENTUAL' in df.columns else 'NOTAS'
        
        if 'CustomerAgent' in df.columns and score_col in df.columns:
            # Não filtrar risco aqui - precisamos de todas as ligações para cálculo correto
            agent_comparison = df.groupby('CustomerAgent').agg({
                score_col: 'mean',
                'IdAnalysis': 'count',
                'ClientRisk': lambda x: (x == 'BAIXO').sum() / len(x) * 100 if len(x) > 0 else 0
            }).round(1)
            
            if score_col == 'NOTAS':
                agent_comparison[score_col] = (agent_comparison[score_col] / 81) * 100
            
            agent_comparison.columns = ['Porcentagem Média', 'Total Ligações', '% Risco Baixo']
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
                        title=dict(text="Acerto<br>Médio (%)", font=dict(size=11, family='Inter')),
                        tickfont=dict(size=10, family='Inter')
                    ),
                    line=dict(width=2, color='white')
                ),
                text=[name.split()[0] for name in agent_comparison.index],
                textposition='top center',
                textfont=dict(size=9, color=CARGLASS_DARK_RED, family='Inter'),
                hovertemplate='<b>%{text}</b><br>Ligações: %{x}<br>Acerto: %{y:.0f}%<br>Risco Baixo: %{marker.size:.1f}%<extra></extra>'
            ))
            
            fig.update_layout(
                title='Análise Comparativa de Agentes',
                xaxis=dict(
                    title=dict(text='Total de Ligações', font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
                    tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter')
                ),
                yaxis=dict(
                    title=dict(text='Porcentagem de Acerto', 
                              font=dict(size=13, color=CARGLASS_GRAY, family='Inter')),
                    tickfont=dict(size=11, color=CARGLASS_GRAY, family='Inter'),
                    range=[0, 110]
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
                annotation_text="Meta: 70%", 
                annotation_position="right",
                annotation_font=dict(size=12, color=CARGLASS_GREEN, family='Inter')
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(
                agent_comparison.style.format({'Porcentagem Média': '{:.0f}%', '% Risco Baixo': '{:.0f}%'}),
                use_container_width=True,
                height=400
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
            
            # Renomear colunas para exibição
            column_rename = {
                'PERCENTUAL': 'Percentual %',
                'AnalysisDateTime': 'Data Análise',
                'CustomerAgent': 'Agente',
                'ClientRisk': 'Risco',
                'ClientOutcome': 'Resultado'
            }
            df_display = df_display.rename(columns=column_rename)
            
            st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                height=400
            )
        
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
                'Valor': [
                    f"{round(avg_val)}%",
                    f"{round(med_val)}%",
                    f"{round(std_val)}%",
                    f"{round(min_val)}%",
                    f"{round(max_val)}%"
                ]
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
