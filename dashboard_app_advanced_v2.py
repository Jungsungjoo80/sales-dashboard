# -*- coding: utf-8 -*-
"""
LABONLAB 트래픽 데이터 분석 대시보드 v7 (최종 완성)
Advanced Traffic Dashboard v7 - Complete

주요 개선사항:
1) ROAS 퍼센트(%) 표시
2) Y축 금액 표기: 200M → 2,000만 형식
3) 숫자 겹침 방지 및 모든 숫자 가시화
4) 쇼핑몰 필터 위치 변경 (선택주차 아래)
5) 데이터 테이블 컬럼 순서 변경
6) Y축 범위 자동 조정 (높은/낮은 수치 모두 표시)
7) 모든 차트에 숫자 표기
"""

import dash
from dash import dcc, html, dash_table, Input, Output, State, callback_context
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objs as go
import plotly.express as px
import base64
import io
from datetime import datetime

# Initialize Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "LABONLAB 트래픽 데이터 분석 대시보드 v7"

# ====================
# Helper Functions
# ====================

def format_currency(value):
    """원화 포맷팅"""
    try:
        val = float(value)
        if pd.isna(val) or val == 0:
            return "₩0"
        return f"₩{val:,.0f}"
    except:
        return "₩0"

def format_currency_short_man(value):
    """만 단위 원화 포맷팅 (예: 4,000만)"""
    try:
        val = float(value)
        if pd.isna(val) or val == 0:
            return "0"
        man = val / 10000
        return f"{man:,.0f}만"
    except:
        return "0"

def format_percent(value):
    """퍼센트 포맷팅"""
    try:
        val = float(value)
        if pd.isna(val):
            return "0%"
        # 마이너스 값은 빨간색 처리를 위해 반환
        return f"{val:.1f}%"
    except:
        return "0%"

def calculate_derived_fields(df):
    """파생 필드 계산 (이익률, ROAS 등)"""
    df = df.copy()
    
    # 숫자형 변환
    numeric_cols = ['매출', '이익', '트래픽비용', '슬롯수']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # 순이익 = 이익 - 트래픽비용
    if '이익' in df.columns and '트래픽비용' in df.columns:
        df['순이익'] = df['이익'] - df['트래픽비용']
    
    # 이익률 = (순이익 / 트래픽비용) × 100
    if '순이익' in df.columns and '트래픽비용' in df.columns:
        df['이익률'] = df.apply(
            lambda row: (row['순이익'] / row['트래픽비용'] * 100) if row['트래픽비용'] > 0 else 0,
            axis=1
        )
    
    # ROAS = (매출 / 트래픽비용) × 100 (퍼센트로 표시)
    if '매출' in df.columns and '트래픽비용' in df.columns:
        df['ROAS'] = df.apply(
            lambda row: (row['매출'] / row['트래픽비용'] * 100) if row['트래픽비용'] > 0 else 0,
            axis=1
        )
    
    # 이익률변동 및 슬롯수변동 계산 (전주 대비)
    if '주차' in df.columns and '이익률' in df.columns:
        df = df.sort_values(['상품명', '쇼핑몰', '주차'])
        df['이익률변동'] = df.groupby(['상품명', '쇼핑몰'])['이익률'].diff()
    
    if '주차' in df.columns and '슬롯수' in df.columns:
        df = df.sort_values(['상품명', '쇼핑몰', '주차'])
        df['슬롯수변동'] = df.groupby(['상품명', '쇼핑몰'])['슬롯수'].diff()
    
    return df

def parse_uploaded_excel(contents):
    """업로드된 Excel 파일 파싱"""
    try:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        df = pd.read_excel(io.BytesIO(decoded))
        
        # 필수 컬럼 확인
        required_cols = ['상품명', '주차', '쇼핑몰', '매출', '이익', '트래픽비용']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return None, f"필수 컬럼 누락: {', '.join(missing)}"
        
        # 선택 컬럼 추가 (없으면 빈 값)
        optional_cols = ['슬롯수', '특이사항', '의견']
        for col in optional_cols:
            if col not in df.columns:
                df[col] = ""
        
        # 파생 필드 계산
        df = calculate_derived_fields(df)
        
        return df, None
    except Exception as e:
        return None, f"파일 읽기 오류: {str(e)}"

# ====================
# Layout
# ====================

app.layout = dbc.Container([
    # Header
    dbc.Row([
        dbc.Col([
            html.H1("📊 LABONLAB 트래픽 데이터 분석 대시보드 v7", className="text-primary mb-2"),
            html.P("Excel 파일을 업로드하여 주차별/쇼핑몰별 데이터를 분석하세요", className="text-muted")
        ], width=8),
        dbc.Col([
            html.Div(id='upload-status', className="text-end")
        ], width=4)
    ], className="mb-4 mt-4"),
    
    # 파일 업로드
    dbc.Row([
        dbc.Col([
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    html.I(className="fas fa-upload me-2"),
                    'Excel 파일을 드래그하거나 클릭하여 업로드'
                ]),
                style={
                    'width': '100%',
                    'height': '80px',
                    'lineHeight': '80px',
                    'borderWidth': '2px',
                    'borderStyle': 'dashed',
                    'borderRadius': '10px',
                    'textAlign': 'center',
                    'backgroundColor': '#f8f9fa',
                    'cursor': 'pointer'
                },
                multiple=False
            )
        ])
    ], className="mb-4"),
    
    # 필터 섹션
    dbc.Row([
        dbc.Col([
            html.Label("📅 선택 주차", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='selected-week',
                placeholder="주차 선택",
                clearable=False
            )
        ], width=3),
        dbc.Col([
            html.Label("🏪 쇼핑몰 필터", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='mall-filter',
                options=[
                    {'label': '전체', 'value': 'all'},
                    {'label': '네이버', 'value': '네이버'},
                    {'label': '쿠팡', 'value': '쿠팡'}
                ],
                value='all',
                clearable=False
            )
        ], width=3),
        dbc.Col([
            html.Label("📊 표시할 컬럼 선택", className="fw-bold mb-2"),
            dcc.Checklist(
                id='column-selector',
                options=[],
                value=[],
                inline=True,
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '10px'}
            )
        ], width=6)
    ], className="mb-4"),
    
    # 비교주차 선택 (가로 배열)
    dbc.Row([
        dbc.Col([
            html.Label("📊 비교 주차 1", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-1', placeholder="선택 안함")
        ], width=3),
        dbc.Col([
            html.Label("📊 비교 주차 2", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-2', placeholder="선택 안함")
        ], width=3),
        dbc.Col([
            html.Label("📊 비교 주차 3", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-3', placeholder="선택 안함")
        ], width=3),
        dbc.Col([
            html.Label("📊 비교 주차 4", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-4', placeholder="선택 안함")
        ], width=3)
    ], className="mb-4"),
    
    # KPI 카드 (8개 지표)
    dbc.Row(id='kpi-cards', className="mb-4"),
    
    # 데이터 테이블
    dbc.Row([
        dbc.Col([
            html.H5("📋 데이터 테이블", className="mb-3"),
            html.Div([
                html.Label("페이지당 행 수: ", className="me-2"),
                dcc.Dropdown(
                    id='page-size',
                    options=[
                        {'label': '10', 'value': 10},
                        {'label': '20', 'value': 20},
                        {'label': '30', 'value': 30},
                        {'label': '50', 'value': 50},
                        {'label': '100', 'value': 100}
                    ],
                    value=20,
                    clearable=False,
                    style={'width': '120px', 'display': 'inline-block'}
                )
            ], className="mb-3"),
            dash_table.DataTable(
                id='data-table',
                columns=[],
                data=[],
                page_size=20,
                style_table={'overflowX': 'auto'},
                style_cell={
                    'textAlign': 'left',
                    'padding': '10px',
                    'fontSize': '13px',
                    'fontFamily': 'Arial, sans-serif'
                },
                style_header={
                    'backgroundColor': '#e9ecef',
                    'fontWeight': 'bold',
                    'textAlign': 'center'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': '#f8f9fa'
                    }
                ]
            )
        ])
    ], className="mb-4"),
    
    # 주차별 통합 시각화
    dbc.Row([
        dbc.Col([
            html.H5("📊 주차별 통합 시각화 (복수 주차 선택 시)", className="mb-3"),
            dcc.Graph(id='integrated-viz', style={'height': '450px'})
        ])
    ], className="mb-4"),
    
    # 개별 차트
    html.Div(id='charts-container'),
    
    # 데이터 저장소
    dcc.Store(id='stored-data'),
    dcc.Store(id='stored-weeks'),
    dcc.Store(id='stored-malls')
    
], fluid=True, style={'backgroundColor': '#f8f9fa', 'minHeight': '100vh', 'padding': '20px'})

# ====================
# Callbacks
# ====================

# Callback 1: 파일 업로드 및 데이터 저장
@app.callback(
    [Output('stored-data', 'data'),
     Output('stored-weeks', 'data'),
     Output('stored-malls', 'data'),
     Output('upload-status', 'children')],
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def upload_file(contents, filename):
    if contents is None:
        return None, None, None, ""
    
    df, error = parse_uploaded_excel(contents)
    if error:
        return None, None, None, html.Div([
            html.I(className="fas fa-exclamation-triangle text-danger me-2"),
            html.Span(error, className="text-danger")
        ])
    
    weeks = sorted(df['주차'].unique().tolist())
    malls = sorted(df['쇼핑몰'].unique().tolist())
    
    status = html.Div([
        html.I(className="fas fa-check-circle text-success me-2"),
        html.Span(f"✓ {filename} 업로드 완료 ({len(df)}개 레코드)", className="text-success fw-bold")
    ])
    
    return df.to_dict('records'), weeks, malls, status

# Callback 2: 주차 드롭다운 업데이트
@app.callback(
    [Output('selected-week', 'options'),
     Output('selected-week', 'value'),
     Output('compare-week-1', 'options'),
     Output('compare-week-2', 'options'),
     Output('compare-week-3', 'options'),
     Output('compare-week-4', 'options')],
    Input('stored-weeks', 'data')
)
def update_week_dropdowns(weeks):
    if not weeks:
        return [], None, [], [], [], []
    
    options = [{'label': w, 'value': w} for w in weeks]
    default_week = weeks[-1] if weeks else None
    
    return options, default_week, options, options, options, options

# Callback 3: 쇼핑몰 필터 업데이트
@app.callback(
    Output('mall-filter', 'options'),
    Input('stored-malls', 'data')
)
def update_mall_filter(malls):
    if not malls:
        return [{'label': '전체', 'value': 'all'}]
    
    options = [{'label': '전체', 'value': 'all'}]
    options.extend([{'label': m, 'value': m} for m in malls])
    return options

# Callback 4: 컬럼 선택 체크리스트 업데이트
@app.callback(
    [Output('column-selector', 'options'),
     Output('column-selector', 'value')],
    Input('stored-data', 'data')
)
def update_column_selector(data):
    if not data:
        return [], []
    
    df = pd.DataFrame(data)
    display_cols = ['매출', '이익', '트래픽비용', '순이익', 'ROAS', '이익률', '이익률변동', '슬롯수', '슬롯수변동']
    available = [col for col in display_cols if col in df.columns]
    
    options = [{'label': col, 'value': col} for col in available]
    default_values = available[:3] if len(available) >= 3 else available
    
    return options, default_values

# Callback 5: KPI 카드 업데이트 (8개 지표)
@app.callback(
    Output('kpi-cards', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('mall-filter', 'value')]
)
def update_kpi_cards(data, selected_week, compare_week, mall_filter):
    if not data or not selected_week:
        return []
    
    df = pd.DataFrame(data)
    
    # 필터 적용
    df_filtered = df[df['주차'] == selected_week].copy()
    if mall_filter and mall_filter != 'all':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 비교 주차 데이터
    df_compare = None
    if compare_week:
        df_compare = df[df['주차'] == compare_week].copy()
        if mall_filter and mall_filter != 'all':
            df_compare = df_compare[df_compare['쇼핑몰'] == mall_filter]
    
    # KPI 계산
    kpis = [
        {
            'title': '총 매출',
            'value': df_filtered['매출'].sum(),
            'format': 'currency',
            'icon': 'fa-sack-dollar',
            'color': 'primary',
            'compare': df_compare['매출'].sum() if df_compare is not None else None
        },
        {
            'title': '총 이익',
            'value': df_filtered['이익'].sum(),
            'format': 'currency',
            'icon': 'fa-chart-line',
            'color': 'success',
            'compare': df_compare['이익'].sum() if df_compare is not None else None
        },
        {
            'title': '트래픽 비용',
            'value': df_filtered['트래픽비용'].sum(),
            'format': 'currency',
            'icon': 'fa-won-sign',
            'color': 'warning',
            'compare': df_compare['트래픽비용'].sum() if df_compare is not None else None
        },
        {
            'title': '순이익',
            'value': df_filtered['순이익'].sum(),
            'format': 'currency',
            'icon': 'fa-coins',
            'color': 'info',
            'compare': df_compare['순이익'].sum() if df_compare is not None else None
        },
        {
            'title': '평균 이익률',
            'value': df_filtered['이익률'].mean(),
            'format': 'percent',
            'icon': 'fa-percent',
            'color': 'secondary',
            'compare': df_compare['이익률'].mean() if df_compare is not None else None
        },
        {
            'title': '평균 이익률변동',
            'value': df_filtered['이익률변동'].mean(),
            'format': 'percent',
            'icon': 'fa-arrow-trend-up',
            'color': 'dark',
            'compare': None
        },
        {
            'title': '총 슬롯수',
            'value': df_filtered['슬롯수'].sum(),
            'format': 'number',
            'icon': 'fa-layer-group',
            'color': 'danger',
            'compare': df_compare['슬롯수'].sum() if df_compare is not None else None
        },
        {
            'title': '평균 ROAS',
            'value': df_filtered['ROAS'].mean(),
            'format': 'percent',
            'icon': 'fa-bullseye',
            'color': 'primary',
            'compare': df_compare['ROAS'].mean() if df_compare is not None else None
        }
    ]
    
    cards = []
    for kpi in kpis:
        # 값 포맷팅
        if kpi['format'] == 'currency':
            main_val = format_currency(kpi['value'])
        elif kpi['format'] == 'percent':
            main_val = format_percent(kpi['value'])
        else:
            main_val = f"{int(kpi['value']):,}"
        
        # 전주 대비 delta 계산
        delta_text = ""
        delta_color = "text-muted"
        if kpi['compare'] is not None:
            delta = kpi['value'] - kpi['compare']
            if kpi['format'] == 'currency':
                delta_text = format_currency(abs(delta))
            elif kpi['format'] == 'percent':
                delta_text = format_percent(abs(delta))
            else:
                delta_text = f"{int(abs(delta)):,}"
            
            if delta > 0:
                delta_text = f"▲ {delta_text}"
                delta_color = "text-success"
            elif delta < 0:
                delta_text = f"▼ {delta_text}"
                delta_color = "text-danger fw-bold"
            else:
                delta_text = "→ 변동 없음"
        
        card = dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.Div([
                        html.I(className=f"fas {kpi['icon']} fa-2x text-{kpi['color']} mb-2"),
                        html.H6(kpi['title'], className="text-muted mb-2"),
                        html.H4(main_val, className=f"text-{kpi['color']} fw-bold mb-1"),
                        html.Small(delta_text, className=delta_color) if delta_text else html.Span()
                    ], className="text-center")
                ])
            ], className="shadow-sm h-100")
        ], width=12, md=6, lg=3, className="mb-3")
        
        cards.append(card)
    
    return cards

# Callback 6: 데이터 테이블 업데이트
@app.callback(
    [Output('data-table', 'data'),
     Output('data-table', 'columns'),
     Output('data-table', 'page_size')],
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value'),
     Input('column-selector', 'value'),
     Input('page-size', 'value')]
)
def update_data_table(data, selected_week, cw1, cw2, cw3, cw4, mall_filter, selected_columns, page_size):
    if not data or not selected_week:
        return [], [], page_size
    
    df = pd.DataFrame(data)
    
    # 주차 필터 (선택 주차 + 비교 주차)
    weeks_to_show = [selected_week]
    for cw in [cw1, cw2, cw3, cw4]:
        if cw and cw not in weeks_to_show:
            weeks_to_show.append(cw)
    
    df_filtered = df[df['주차'].isin(weeks_to_show)].copy()
    
    # 쇼핑몰 필터
    if mall_filter and mall_filter != 'all':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 컬럼 순서 정의 (요청에 따라)
    column_order = ['상품명', '주차', '쇼핑몰', '매출', '이익', '트래픽비용', '순이익', 'ROAS', 
                    '이익률', '이익률변동', '슬롯수', '슬롯수변동', '특이사항', '의견']
    
    # 존재하는 컬럼만 선택
    display_cols = [col for col in column_order if col in df_filtered.columns]
    
    # 컬럼 선택 필터 적용
    if selected_columns:
        base_cols = ['상품명', '주차', '쇼핑몰']
        display_cols = base_cols + [col for col in display_cols if col in selected_columns or col in base_cols]
    
    df_display = df_filtered[display_cols].copy()
    
    # 포맷팅
    for col in df_display.columns:
        if col in ['매출', '이익', '트래픽비용', '순이익']:
            df_display[col] = df_display[col].apply(format_currency)
        elif col in ['이익률', '이익률변동', 'ROAS']:
            df_display[col] = df_display[col].apply(format_percent)
        elif col in ['슬롯수', '슬롯수변동']:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "0")
    
    columns = [{'name': col, 'id': col} for col in display_cols]
    
    return df_display.to_dict('records'), columns, page_size

# Callback 7: 통합 시각화
@app.callback(
    Output('integrated-viz', 'figure'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value')]
)
def update_integrated_viz(data, selected_week, cw1, cw2, cw3, cw4, mall_filter):
    if not data or not selected_week:
        return go.Figure()
    
    df = pd.DataFrame(data)
    
    # 비교 주차 수집
    compare_weeks = [cw for cw in [cw1, cw2, cw3, cw4] if cw]
    if not compare_weeks:
        return go.Figure().update_layout(
            title="비교 주차를 선택하면 통합 시각화가 표시됩니다",
            height=450
        )
    
    all_weeks = [selected_week] + compare_weeks
    
    # 필터 적용
    df_filtered = df[df['주차'].isin(all_weeks)].copy()
    if mall_filter and mall_filter != 'all':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 주차별 집계
    df_agg = df_filtered.groupby('주차').agg({
        '매출': 'sum',
        '이익': 'sum',
        '트래픽비용': 'sum',
        '순이익': 'sum',
        '슬롯수': 'sum',
        '이익률': 'mean',
        '이익률변동': 'mean',
        'ROAS': 'mean'
    }).reset_index()
    
    # 주차 순서 정렬
    df_agg['주차'] = pd.Categorical(df_agg['주차'], categories=all_weeks, ordered=True)
    df_agg = df_agg.sort_values('주차')
    
    # 차트 생성
    fig = go.Figure()
    
    # 매출 (막대)
    fig.add_trace(go.Bar(
        x=df_agg['주차'],
        y=df_agg['매출'],
        name='매출',
        marker_color='#3b82f6',
        text=[format_currency_short_man(v) for v in df_agg['매출']],
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>매출: %{text}<extra></extra>'
    ))
    
    # 이익 (라인)
    fig.add_trace(go.Scatter(
        x=df_agg['주차'],
        y=df_agg['이익'],
        name='이익',
        mode='lines+markers+text',
        line=dict(color='#10b981', width=3),
        marker=dict(size=10),
        text=[format_currency_short_man(v) for v in df_agg['이익']],
        textposition='top center',
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>이익: %{text}<extra></extra>'
    ))
    
    # 트래픽비용 (라인)
    fig.add_trace(go.Scatter(
        x=df_agg['주차'],
        y=df_agg['트래픽비용'],
        name='트래픽비용',
        mode='lines+markers+text',
        line=dict(color='#f59e0b', width=3),
        marker=dict(size=10),
        text=[format_currency_short_man(v) for v in df_agg['트래픽비용']],
        textposition='bottom center',
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>트래픽비용: %{text}<extra></extra>'
    ))
    
    # 순이익 (라인)
    fig.add_trace(go.Scatter(
        x=df_agg['주차'],
        y=df_agg['순이익'],
        name='순이익',
        mode='lines+markers+text',
        line=dict(color='#ef4444', width=3, dash='dash'),
        marker=dict(size=10),
        text=[format_currency_short_man(v) for v in df_agg['순이익']],
        textposition='top center',
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>순이익: %{text}<extra></extra>'
    ))
    
    # 슬롯수 (라인)
    fig.add_trace(go.Scatter(
        x=df_agg['주차'],
        y=df_agg['슬롯수'],
        name='슬롯수',
        mode='lines+markers+text',
        line=dict(color='#8b5cf6', width=3),
        marker=dict(size=10),
        text=[f"{int(v):,}" for v in df_agg['슬롯수']],
        textposition='bottom center',
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>슬롯수: %{text}<extra></extra>'
    ))
    
    # 레이아웃
    fig.update_layout(
        title="주차별 통합 시각화",
        xaxis=dict(title="주차"),
        yaxis=dict(
            title="매출 (원)",
            titlefont=dict(color='#3b82f6'),
            tickfont=dict(color='#3b82f6'),
            tickformat=',',
            ticksuffix='만',
            autorange=True
        ),
        yaxis2=dict(
            title="이익/비용/슬롯수",
            titlefont=dict(color='#10b981'),
            tickfont=dict(color='#10b981'),
            overlaying='y',
            side='right',
            tickformat=',',
            autorange=True
        ),
        hovermode='x unified',
        height=450,
        margin=dict(t=100, b=50, l=80, r=80),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        )
    )
    
    return fig

# Callback 8: 개별 차트 생성
@app.callback(
    Output('charts-container', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value'),
     Input('column-selector', 'value')]
)
def update_charts(data, selected_week, cw1, cw2, cw3, cw4, mall_filter, selected_columns):
    if not data or not selected_week or not selected_columns:
        return []
    
    df = pd.DataFrame(data)
    
    # 비교 주차 수집
    compare_weeks = [cw for cw in [cw1, cw2, cw3, cw4] if cw]
    all_weeks = [selected_week] + compare_weeks
    
    # 필터 적용
    df_filtered = df[df['주차'].isin(all_weeks)].copy()
    if mall_filter and mall_filter != 'all':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 쇼핑몰별 색상 매핑
    color_map = {'네이버': '#22c55e', '쿠팡': '#3b82f6'}
    
    charts = []
    
    # 숫자형 컬럼만 선택
    numeric_cols = ['매출', '이익', '트래픽비용', '순이익', '이익률', '이익률변동', '슬롯수', '슬롯수변동', 'ROAS']
    selected_numeric = [col for col in selected_columns if col in numeric_cols and col in df_filtered.columns]
    
    for col in selected_numeric:
        # 주차별 데이터 준비
        if compare_weeks:
            # 다중 주차 비교 모드
            df_chart = df_filtered.groupby(['주차', '상품명', '쇼핑몰'])[col].sum().reset_index()
            
            # 차트 너비 동적 조정 (가로 스크롤)
            unique_products = df_chart['상품명'].nunique()
            chart_width = max(1200, unique_products * 80)
            
            fig = go.Figure()
            
            for week in all_weeks:
                df_week = df_chart[df_chart['주차'] == week]
                for mall in df_week['쇼핑몰'].unique():
                    df_mall = df_week[df_week['쇼핑몰'] == mall]
                    
                    # 텍스트 포맷
                    if col in ['매출', '이익', '트래픽비용', '순이익']:
                        text_values = [format_currency_short_man(v) for v in df_mall[col]]
                    elif col in ['이익률', '이익률변동', 'ROAS']:
                        text_values = [format_percent(v) for v in df_mall[col]]
                    else:
                        text_values = [f"{int(v):,}" for v in df_mall[col]]
                    
                    fig.add_trace(go.Bar(
                        x=df_mall['상품명'],
                        y=df_mall[col],
                        name=f"{week} - {mall}",
                        marker_color=color_map.get(mall, '#6366f1'),
                        text=text_values,
                        textposition='outside',
                        hovertemplate='<b>%{x}</b><br>%{text}<extra></extra>'
                    ))
            
            # Y축 포맷
            if col in ['매출', '이익', '트래픽비용', '순이익']:
                yaxis_format = dict(tickformat=',', ticksuffix='만')
            elif col in ['이익률', '이익률변동', 'ROAS']:
                yaxis_format = dict(ticksuffix='%')
            else:
                yaxis_format = dict(tickformat=',')
            
            fig.update_layout(
                title=f"{col} 비교 (다중 주차)",
                xaxis_title="상품명",
                yaxis_title=col,
                yaxis=yaxis_format,
                barmode='group',
                height=500,
                width=chart_width,
                margin=dict(t=80, b=100, l=80, r=50),
                hovermode='x unified',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
            )
        else:
            # 단일 주차 모드 (상위 상품만 표시)
            df_chart = df_filtered.groupby(['상품명', '쇼핑몰'])[col].sum().reset_index()
            df_chart = df_chart.nlargest(20, col)
            
            # 차트 너비 동적 조정
            chart_width = max(1200, len(df_chart) * 60)
            
            fig = go.Figure()
            
            for mall in df_chart['쇼핑몰'].unique():
                df_mall = df_chart[df_chart['쇼핑몰'] == mall]
                
                # 텍스트 포맷
                if col in ['매출', '이익', '트래픽비용', '순이익']:
                    text_values = [format_currency_short_man(v) for v in df_mall[col]]
                elif col in ['이익률', '이익률변동', 'ROAS']:
                    text_values = [format_percent(v) for v in df_mall[col]]
                else:
                    text_values = [f"{int(v):,}" for v in df_mall[col]]
                
                fig.add_trace(go.Bar(
                    x=df_mall['상품명'],
                    y=df_mall[col],
                    name=mall,
                    marker_color=color_map.get(mall, '#6366f1'),
                    text=text_values,
                    textposition='outside',
                    hovertemplate='<b>%{x}</b><br>%{text}<extra></extra>'
                ))
            
            # Y축 포맷
            if col in ['매출', '이익', '트래픽비용', '순이익']:
                yaxis_format = dict(tickformat=',', ticksuffix='만')
            elif col in ['이익률', '이익률변동', 'ROAS']:
                yaxis_format = dict(ticksuffix='%')
            else:
                yaxis_format = dict(tickformat=',')
            
            fig.update_layout(
                title=f"{col} (상위 20개 상품)",
                xaxis_title="상품명",
                yaxis_title=col,
                yaxis=yaxis_format,
                height=500,
                width=chart_width,
                margin=dict(t=80, b=100, l=80, r=50),
                hovermode='x unified',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
            )
        
        # Y축 범위 자동 조정 (높은/낮은 수치 모두 표시)
        fig.update_yaxes(autorange=True)
        
        # 차트를 Div로 감싸서 가로 스크롤 적용
        chart_div = html.Div([
            html.H5(f"📊 {col}", className="mb-3"),
            html.Div([
                dcc.Graph(figure=fig, config={'displayModeBar': False})
            ], style={'overflowX': 'auto', 'overflowY': 'hidden'})
        ], className="mb-4")
        
        charts.append(chart_div)
    
    return charts

# ====================
# Run Server
# ====================

if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8050, debug=False)
