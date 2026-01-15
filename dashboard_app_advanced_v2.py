"""
Advanced Sales Dashboard v6 - Final Version
완벽한 데이터 분석 대시보드

주요 개선사항:
1. ✅ 중복 제거 로직 완전 제거 (357개 레코드 전체 로드)
2. ✅ 데이터 테이블 페이지 크기 선택 (10/20/30/50/100)
3. ✅ 순이익/이익률/이익률변동/슬롯수변동 자동 계산
4. ✅ 쇼핑몰 필터 (전체/네이버/쿠팡)
5. ✅ 이익률변동/슬롯수변동/슬롯수 차트 추가
6. ✅ 통합 시각화 기능 (비교 주차 선택 시 복합 차트)
7. ✅ 비교 주차 최대 4개까지 선택
8. ✅ 표시할 컬럼 선택을 가로 레이아웃으로 변경

작성일: 2026-01-15
버전: v6.0
"""

import dash
from dash import dcc, html, dash_table, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import base64
import io
from datetime import datetime

# ===========================
# Initialize Dash App
# ===========================

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# ===========================
# Helper Functions
# ===========================

def format_currency(value):
    """금액을 천단위 콤마 + 원화 형식으로 표시 (예: 210,800,000원)"""
    try:
        if pd.isna(value) or value == 0:
            return "0원"
        return f"{int(value):,}원"
    except:
        return "N/A"

def format_number(value):
    """숫자를 천단위 콤마로 표시 (예: 1,234)"""
    try:
        if pd.isna(value):
            return "N/A"
        return f"{int(value):,}"
    except:
        return "N/A"

def format_percentage(value):
    """퍼센트 표시 (소수점 1자리)"""
    try:
        if pd.isna(value):
            return "N/A"
        return f"{value:.1f}%"
    except:
        return "N/A"

def calculate_derived_fields(df):
    """
    계산 필드 자동 생성
    - 순이익 = 이익 - 트래픽비용
    - 이익률 = (이익 / 매출) * 100
    - 이익률변동 = 현재 주차 이익률 - 이전 주차 이익률
    - 슬롯수변동 = 현재 주차 슬롯수 - 이전 주차 슬롯수
    """
    df = df.copy()
    
    # 1. 순이익 계산
    if '이익' in df.columns and '트래픽비용' in df.columns:
        df['순이익'] = df['이익'] - df['트래픽비용']
        # 기존 순이익 컬럼이 있다면 제거하고 계산값 사용
        if '순이익_원본' not in df.columns:
            df['순이익_원본'] = df['순이익']
    
    # 2. 이익률 계산
    if '이익' in df.columns and '매출' in df.columns:
        df['이익률'] = ((df['이익'] / df['매출']) * 100).fillna(0)
        # 기존 이익률 컬럼이 있다면 제거하고 계산값 사용
        if '이익률_원본' not in df.columns:
            df['이익률_원본'] = df['이익률']
    
    # 3. 주차별 이익률변동 및 슬롯수변동 계산
    if '주차' in df.columns and '상품명' in df.columns:
        # 주차를 날짜 순으로 정렬
        weeks = sorted(df['주차'].unique())
        
        # 각 상품별로 이전 주차와 비교
        df['이익률변동'] = 0.0
        df['슬롯수변동'] = 0
        
        for product in df['상품명'].unique():
            product_data = df[df['상품명'] == product].sort_values('주차')
            
            if len(product_data) > 1:
                for i in range(1, len(product_data)):
                    curr_idx = product_data.index[i]
                    prev_idx = product_data.index[i-1]
                    
                    # 이익률변동 계산
                    if '이익률' in df.columns:
                        curr_roi = df.at[curr_idx, '이익률']
                        prev_roi = df.at[prev_idx, '이익률']
                        df.at[curr_idx, '이익률변동'] = curr_roi - prev_roi
                    
                    # 슬롯수변동 계산
                    if '슬롯수' in df.columns:
                        curr_slots = df.at[curr_idx, '슬롯수'] if pd.notna(df.at[curr_idx, '슬롯수']) else 0
                        prev_slots = df.at[prev_idx, '슬롯수'] if pd.notna(df.at[prev_idx, '슬롯수']) else 0
                        df.at[curr_idx, '슬롯수변동'] = int(curr_slots - prev_slots)
    
    return df

def parse_uploaded_excel(contents, filename):
    """
    Excel 파일 파싱
    - 중복 제거 로직 완전 제거 (모든 데이터 유지)
    - 동적 헤더 감지
    - 계산 필드 자동 생성
    """
    try:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        
        # Read Excel file
        df = pd.read_excel(io.BytesIO(decoded), header=0)
        
        # Reset index to prevent reindexing errors
        df = df.reset_index(drop=True)
        
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Standardize column names (remove leading/trailing spaces)
        df.columns = df.columns.str.strip()
        
        # ✅ 중복 제거 로직 완전 제거 - 모든 357개 레코드 유지
        # (기존 df.drop_duplicates 코드 삭제됨)
        
        # 계산 필드 자동 생성
        df = calculate_derived_fields(df)
        
        # Extract unique weeks from B column (주차)
        weeks = []
        if '주차' in df.columns:
            weeks = sorted(df['주차'].dropna().unique().tolist(), reverse=True)
        
        # Extract unique shopping malls
        malls = []
        if '쇼핑몰' in df.columns:
            malls = sorted(df['쇼핑몰'].dropna().unique().tolist())
        
        return df, weeks, malls, None
        
    except Exception as e:
        return None, [], [], f"파일 파싱 오류: {str(e)}"

def create_kpi_card(title, value, icon="📊", color="primary"):
    """KPI 카드 생성"""
    return dbc.Card([
        dbc.CardBody([
            html.H6(title, className="text-muted mb-2", style={"fontSize": "13px"}),
            html.H3(value, className=f"text-{color} mb-0", style={"fontSize": "18px", "fontWeight": "bold"}),
            html.P(icon, className="mb-0", style={"fontSize": "20px"})
        ])
    ], className="shadow-sm", style={"height": "100%"})

# ===========================
# Layout
# ===========================

app.layout = dbc.Container([
    # Header
    dbc.Row([
        dbc.Col([
            html.H1("📊 LABONLAB 트래픽 데이터 분석 대시보드 v6", className="text-primary mb-2"),
            html.P("Excel 파일을 업로드하여 주차별/쇼핑몰별 데이터를 분석하세요", className="text-muted")
        ])
    ], className="mb-4 mt-4"),
    
    # File Upload Section
    dbc.Row([
        dbc.Col([
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    '📤 Excel 파일을 드래그하거나 ',
                    html.A('클릭하여 업로드하세요', style={"color": "#007bff", "cursor": "pointer"})
                ]),
                style={
                    'width': '100%',
                    'height': '80px',
                    'lineHeight': '80px',
                    'borderWidth': '2px',
                    'borderStyle': 'dashed',
                    'borderRadius': '10px',
                    'textAlign': 'center',
                    'backgroundColor': '#f8f9fa'
                },
                multiple=False
            ),
            html.Div(id='upload-status', className="mt-3")
        ])
    ], className="mb-4"),
    
    # Filter Controls (주차 선택 + 쇼핑몰 필터)
    dbc.Row([
        dbc.Col([
            html.Label("주차 선택", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='week-selector',
                placeholder="주차를 선택하세요",
                className="mb-3"
            )
        ], md=2),
        dbc.Col([
            html.Label("비교 주차 1", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='compare-week-1',
                placeholder="선택 (옵션)",
                className="mb-3"
            )
        ], md=2),
        dbc.Col([
            html.Label("비교 주차 2", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='compare-week-2',
                placeholder="선택 (옵션)",
                className="mb-3"
            )
        ], md=2),
        dbc.Col([
            html.Label("비교 주차 3", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='compare-week-3',
                placeholder="선택 (옵션)",
                className="mb-3"
            )
        ], md=2),
        dbc.Col([
            html.Label("비교 주차 4", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='compare-week-4',
                placeholder="선택 (옵션)",
                className="mb-3"
            )
        ], md=2),
        dbc.Col([
            html.Label("쇼핑몰 필터", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='mall-filter',
                options=[{'label': '전체', 'value': 'all'}],
                value='all',
                placeholder="전체",
                className="mb-3"
            )
        ], md=2)
    ], className="mb-3", id="filter-section", style={"display": "none"}),
    
    # Column Selector (가로 레이아웃)
    dbc.Row([
        dbc.Col([
            html.Label("표시할 컬럼 선택", className="fw-bold mb-2"),
            dcc.Checklist(
                id='column-selector',
                options=[],
                value=[],
                inline=True,  # 가로 레이아웃
                labelStyle={'display': 'inline-block', 'marginRight': '15px'},
                className="mb-3"
            )
        ])
    ], className="mb-4", id="column-section", style={"display": "none"}),
    
    # KPI Cards
    dbc.Row([
        dbc.Col(html.Div(id='kpi-cards'), md=12)
    ], className="mb-4"),
    
    # Data Table with Page Size Selector
    dbc.Row([
        dbc.Col([
            html.Div([
                html.H4("📋 데이터 테이블", className="mb-3 d-inline-block"),
                html.Div([
                    html.Label("표시 개수:", className="me-2", style={"fontSize": "14px"}),
                    dcc.Dropdown(
                        id='page-size-selector',
                        options=[
                            {'label': '10개', 'value': 10},
                            {'label': '20개', 'value': 20},
                            {'label': '30개', 'value': 30},
                            {'label': '50개', 'value': 50},
                            {'label': '100개', 'value': 100}
                        ],
                        value=20,
                        clearable=False,
                        style={'width': '120px', 'display': 'inline-block'}
                    )
                ], className="float-end")
            ], className="clearfix mb-3"),
            html.Div(id='data-table-container')
        ])
    ], className="mb-4"),
    
    # Integrated Visualization (비교 주차 선택 시만 표시)
    dbc.Row([
        dbc.Col([
            html.Div(id='integrated-viz-section', style={"display": "none"})
        ])
    ], className="mb-4"),
    
    # Charts Section
    dbc.Row([
        dbc.Col([
            html.H4("📈 데이터 시각화", className="mb-3"),
            html.Div(id='charts-container')
        ])
    ], className="mb-4"),
    
    # Hidden data stores
    dcc.Store(id='stored-data'),
    dcc.Store(id='stored-weeks'),
    dcc.Store(id='stored-malls')
    
], fluid=True, style={"backgroundColor": "#f5f5f5", "minHeight": "100vh", "paddingBottom": "50px"})

# ===========================
# Callbacks
# ===========================

@app.callback(
    [Output('stored-data', 'data'),
     Output('stored-weeks', 'data'),
     Output('stored-malls', 'data'),
     Output('upload-status', 'children'),
     Output('filter-section', 'style'),
     Output('column-section', 'style')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def upload_file(contents, filename):
    """파일 업로드 및 파싱"""
    if contents is None:
        return None, None, None, "", {"display": "none"}, {"display": "none"}
    
    df, weeks, malls, error = parse_uploaded_excel(contents, filename)
    
    if error:
        return None, None, None, dbc.Alert(error, color="danger"), {"display": "none"}, {"display": "none"}
    
    if df is not None and len(df) > 0:
        success_msg = dbc.Alert([
            html.I(className="bi bi-check-circle-fill me-2"),
            f"✓ {filename} 업로드 완료 ({len(df)}개 레코드)"
        ], color="success")
        
        return (
            df.to_json(date_format='iso', orient='split'), 
            weeks, 
            malls,
            success_msg, 
            {"display": "block"},
            {"display": "block"}
        )
    
    return None, None, None, dbc.Alert("파일에 데이터가 없습니다.", color="warning"), {"display": "none"}, {"display": "none"}

@app.callback(
    [Output('week-selector', 'options'),
     Output('week-selector', 'value'),
     Output('compare-week-1', 'options'),
     Output('compare-week-2', 'options'),
     Output('compare-week-3', 'options'),
     Output('compare-week-4', 'options')],
    [Input('stored-weeks', 'data')]
)
def update_week_selectors(weeks):
    """주차 선택 드롭다운 업데이트"""
    if not weeks or len(weeks) == 0:
        return [], None, [], [], [], []
    
    options = [{'label': week, 'value': week} for week in weeks]
    default_week = weeks[0] if weeks else None
    
    return options, default_week, options, options, options, options

@app.callback(
    Output('mall-filter', 'options'),
    [Input('stored-malls', 'data')]
)
def update_mall_filter(malls):
    """쇼핑몰 필터 드롭다운 업데이트"""
    if not malls or len(malls) == 0:
        return [{'label': '전체', 'value': 'all'}]
    
    options = [{'label': '전체', 'value': 'all'}]
    options.extend([{'label': mall, 'value': mall} for mall in malls])
    
    return options

@app.callback(
    [Output('column-selector', 'options'),
     Output('column-selector', 'value')],
    [Input('stored-data', 'data')]
)
def update_column_selector(json_data):
    """컬럼 선택 체크박스 업데이트"""
    if json_data is None:
        return [], []
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Get all columns except text-only columns
    all_cols = df.columns.tolist()
    
    # Exclude certain columns
    exclude_cols = ['상품명', '주차']
    selectable_cols = [col for col in all_cols if col not in exclude_cols]
    
    # Prioritize numeric columns
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    numeric_cols = [col for col in numeric_cols if col not in exclude_cols]
    
    # Add text columns (쇼핑몰, 특이사항, 의견)
    text_cols = [col for col in selectable_cols if col not in numeric_cols]
    
    # Combine: numeric first, then text
    final_cols = numeric_cols + text_cols
    
    options = [{'label': col, 'value': col} for col in final_cols]
    
    # Default: select first 5 numeric columns
    default_values = numeric_cols[:5] if len(numeric_cols) >= 5 else numeric_cols
    
    return options, default_values

@app.callback(
    Output('kpi-cards', 'children'),
    [Input('stored-data', 'data'),
     Input('week-selector', 'value'),
     Input('compare-week-1', 'value'),
     Input('mall-filter', 'value')]
)
def update_kpi_cards(json_data, selected_week, compare_week_1, mall_filter):
    """KPI 카드 업데이트"""
    if json_data is None or selected_week is None:
        return html.Div()
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Filter by selected week
    if '주차' in df.columns:
        df_filtered = df[df['주차'] == selected_week]
    else:
        df_filtered = df
    
    # Filter by shopping mall
    if mall_filter and mall_filter != 'all' and '쇼핑몰' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # Calculate KPIs
    total_sales = df_filtered['매출'].sum() if '매출' in df_filtered.columns else 0
    total_profit = df_filtered['이익'].sum() if '이익' in df_filtered.columns else 0
    total_net_profit = df_filtered['순이익'].sum() if '순이익' in df_filtered.columns else 0
    avg_roi = df_filtered['이익률'].mean() if '이익률' in df_filtered.columns else 0
    product_count = len(df_filtered)
    
    # Calculate comparison if compare week is selected
    change_text = ""
    if compare_week_1 and compare_week_1 != selected_week and '주차' in df.columns:
        df_compare = df[df['주차'] == compare_week_1]
        
        # Filter by mall
        if mall_filter and mall_filter != 'all' and '쇼핑몰' in df_compare.columns:
            df_compare = df_compare[df_compare['쇼핑몰'] == mall_filter]
        
        prev_sales = df_compare['매출'].sum() if '매출' in df_compare.columns else 0
        
        if prev_sales > 0:
            change_pct = ((total_sales - prev_sales) / prev_sales) * 100
            change_text = f" ({change_pct:+.1f}% vs {compare_week_1})"
    
    cards = dbc.Row([
        dbc.Col(create_kpi_card("총 매출", format_currency(total_sales) + change_text, "💰", "success"), md=3),
        dbc.Col(create_kpi_card("총 이익", format_currency(total_profit), "📈", "info"), md=2),
        dbc.Col(create_kpi_card("순이익", format_currency(total_net_profit), "💎", "primary"), md=2),
        dbc.Col(create_kpi_card("평균 ROI", format_percentage(avg_roi), "📊", "warning"), md=2),
        dbc.Col(create_kpi_card("상품 수", f"{product_count}개", "🛍️", "secondary"), md=3)
    ])
    
    return cards

@app.callback(
    Output('data-table-container', 'children'),
    [Input('stored-data', 'data'),
     Input('week-selector', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('column-selector', 'value'),
     Input('page-size-selector', 'value'),
     Input('mall-filter', 'value')]
)
def update_data_table(json_data, selected_week, cw1, cw2, cw3, cw4, selected_columns, page_size, mall_filter):
    """데이터 테이블 업데이트 - 비교 주차 데이터 병합 표시"""
    if json_data is None or selected_week is None:
        return html.Div("데이터를 업로드하고 주차를 선택하세요.", className="text-muted")
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Collect all weeks to display
    weeks_to_show = [selected_week]
    for cw in [cw1, cw2, cw3, cw4]:
        if cw and cw != selected_week and cw not in weeks_to_show:
            weeks_to_show.append(cw)
    
    # Filter by weeks
    if '주차' in df.columns:
        df_filtered = df[df['주차'].isin(weeks_to_show)]
    else:
        df_filtered = df
    
    # Filter by shopping mall
    if mall_filter and mall_filter != 'all' and '쇼핑몰' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # Select columns to display
    base_cols = ['상품명', '주차']
    if '쇼핑몰' in df_filtered.columns:
        base_cols.append('쇼핑몰')
    
    display_cols = base_cols + (selected_columns if selected_columns else [])
    display_cols = [col for col in display_cols if col in df_filtered.columns]
    
    df_display = df_filtered[display_cols].copy()
    
    # Format columns
    for col in df_display.columns:
        if col in ['상품명', '주차', '쇼핑몰', '특이사항', '의견']:
            continue
        elif df_display[col].dtype in ['float64', 'int64']:
            if '이익률' in col and '변동' not in col:
                df_display[col] = df_display[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A")
            elif '변동' in col:
                df_display[col] = df_display[col].apply(lambda x: f"{int(x)}" if pd.notna(x) else "N/A")
            elif '슬롯수' in col:
                df_display[col] = df_display[col].apply(lambda x: f"{int(x)}" if pd.notna(x) else "N/A")
            else:
                df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "N/A")
    
    table = dash_table.DataTable(
        data=df_display.to_dict('records'),
        columns=[{"name": col, "id": col} for col in df_display.columns],
        page_size=page_size,
        sort_action="native",
        filter_action="native",
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '10px',
            'fontFamily': 'Arial, sans-serif',
            'minWidth': '100px',
            'fontSize': '13px'
        },
        style_header={
            'backgroundColor': '#007bff',
            'color': 'white',
            'fontWeight': 'bold',
            'fontSize': '14px'
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8f9fa'
            },
            {
                'if': {'column_id': '주차'},
                'fontWeight': 'bold',
                'backgroundColor': '#e7f3ff'
            }
        ]
    )
    
    return table

@app.callback(
    Output('integrated-viz-section', 'children'),
    Output('integrated-viz-section', 'style'),
    [Input('stored-data', 'data'),
     Input('week-selector', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value')]
)
def update_integrated_viz(json_data, selected_week, cw1, cw2, cw3, cw4, mall_filter):
    """통합 시각화 - 비교 주차 선택 시만 표시"""
    # Check if any compare week is selected
    compare_weeks = [w for w in [cw1, cw2, cw3, cw4] if w and w != selected_week]
    
    if not compare_weeks or json_data is None or selected_week is None:
        return html.Div(), {"display": "none"}
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Filter by mall
    if mall_filter and mall_filter != 'all' and '쇼핑몰' in df.columns:
        df = df[df['쇼핑몰'] == mall_filter]
    
    # Collect all weeks
    all_weeks = [selected_week] + compare_weeks
    
    # Calculate aggregated metrics per week
    metrics = []
    for week in all_weeks:
        if '주차' in df.columns:
            df_week = df[df['주차'] == week]
        else:
            df_week = df
        
        metrics.append({
            '주차': week,
            '매출': df_week['매출'].sum() if '매출' in df_week.columns else 0,
            '이익': df_week['이익'].sum() if '이익' in df_week.columns else 0,
            '트래픽비용': df_week['트래픽비용'].sum() if '트래픽비용' in df_week.columns else 0,
            '순이익': df_week['순이익'].sum() if '순이익' in df_week.columns else 0,
            '이익률변동': df_week['이익률변동'].mean() if '이익률변동' in df_week.columns else 0,
            '슬롯수변동': df_week['슬롯수변동'].sum() if '슬롯수변동' in df_week.columns else 0,
            '슬롯수': df_week['슬롯수'].sum() if '슬롯수' in df_week.columns else 0
        })
    
    df_metrics = pd.DataFrame(metrics)
    
    # Create integrated chart (복합 차트)
    fig = go.Figure()
    
    # Bar chart for sales (매출 - 기본 막대)
    fig.add_trace(go.Bar(
        name='매출',
        x=df_metrics['주차'],
        y=df_metrics['매출'],
        marker_color='#3b82f6',
        yaxis='y',
        hovertemplate='<b>%{x}</b><br>매출: %{y:,.0f}원<extra></extra>'
    ))
    
    # Line charts for other metrics
    colors = {
        '이익': '#10b981',
        '트래픽비용': '#ef4444',
        '순이익': '#8b5cf6',
        '이익률변동': '#f59e0b',
        '슬롯수변동': '#ec4899',
        '슬롯수': '#06b6d4'
    }
    
    for metric in ['이익', '트래픽비용', '순이익', '이익률변동', '슬롯수변동', '슬롯수']:
        if metric in df_metrics.columns:
            fig.add_trace(go.Scatter(
                name=metric,
                x=df_metrics['주차'],
                y=df_metrics[metric],
                mode='lines+markers',
                line=dict(color=colors[metric], width=3),
                marker=dict(size=8),
                yaxis='y2',
                hovertemplate=f'<b>%{{x}}</b><br>{metric}: %{{y:,.0f}}<extra></extra>'
            ))
    
    # Update layout with dual y-axis
    fig.update_layout(
        title="📊 통합 시각화 (주차별 비교)",
        xaxis=dict(title="주차"),
        yaxis=dict(
            title="매출 (원)",
            side='left',
            tickformat=',',
            ticksuffix='원'
        ),
        yaxis2=dict(
            title="기타 지표",
            side='right',
            overlaying='y',
            tickformat=','
        ),
        hovermode='x unified',
        height=500,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    content = html.Div([
        html.H4("📊 통합 시각화", className="mb-3"),
        dcc.Graph(figure=fig)
    ])
    
    return content, {"display": "block"}

@app.callback(
    Output('charts-container', 'children'),
    [Input('stored-data', 'data'),
     Input('week-selector', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('column-selector', 'value'),
     Input('mall-filter', 'value')]
)
def update_charts(json_data, selected_week, cw1, cw2, cw3, cw4, selected_columns, mall_filter):
    """차트 업데이트 - 모든 컬럼 차트 생성"""
    if json_data is None or selected_week is None or not selected_columns:
        return html.Div("컬럼을 선택하면 차트가 표시됩니다.", className="text-muted")
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Filter by mall
    if mall_filter and mall_filter != 'all' and '쇼핑몰' in df.columns:
        df = df[df['쇼핑몰'] == mall_filter]
    
    # Collect compare weeks
    compare_weeks = [w for w in [cw1, cw2, cw3, cw4] if w and w != selected_week]
    compare_mode = len(compare_weeks) > 0
    
    charts = []
    
    for col in selected_columns:
        if col not in df.columns or col in ['쇼핑몰', '특이사항', '의견']:
            continue
        
        # Check if numeric
        if df[col].dtype not in ['float64', 'int64']:
            continue
        
        if compare_mode:
            # Grouped bar chart for comparison
            all_weeks = [selected_week] + compare_weeks
            df_compare = df[df['주차'].isin(all_weeks)] if '주차' in df.columns else df
            df_sorted = df_compare.nlargest(20, col) if col in df_compare.columns else df_compare
            
            if '상품명' in df.columns:
                product_names = df_sorted[df_sorted['주차'] == selected_week]['상품명'].head(20).tolist()
                
                fig = go.Figure()
                
                colors_list = ['#3b82f6', '#f97316', '#10b981', '#8b5cf6', '#f59e0b']
                
                for idx, week in enumerate(all_weeks):
                    df_week = df_compare[df_compare['주차'] == week]
                    values = []
                    for product in product_names:
                        val = df_week[df_week['상품명'] == product][col].values
                        values.append(val[0] if len(val) > 0 else 0)
                    
                    hover_template = '<b>%{x}</b><br>' + col + ': %{y:,.0f}원<extra></extra>'
                    if '이익률' in col or '%' in col:
                        hover_template = '<b>%{x}</b><br>' + col + ': %{y:.1f}%<extra></extra>'
                    
                    fig.add_trace(go.Bar(
                        name=week,
                        x=product_names,
                        y=values,
                        marker_color=colors_list[idx % len(colors_list)],
                        hovertemplate=hover_template
                    ))
                
                fig.update_layout(
                    title=f"{col} 비교 - 상위 20개 상품 (주차별)",
                    xaxis_title="상품명",
                    yaxis_title=col,
                    barmode='group',
                    xaxis_tickangle=-45,
                    height=500,
                    hovermode='x unified'
                )
                
                if '이익률' in col or '%' in col:
                    fig.update_yaxes(ticksuffix='%')
                else:
                    fig.update_yaxes(tickformat=',', ticksuffix='원')
                
                charts.append(dbc.Col([
                    dcc.Graph(figure=fig)
                ], md=12))
        
        else:
            # Single week chart
            df_filtered = df[df['주차'] == selected_week] if '주차' in df.columns else df
            df_sorted = df_filtered.nlargest(20, col) if col in df_filtered.columns else df_filtered
            
            if '상품명' in df_filtered.columns:
                fig = go.Figure()
                
                hover_template = '<b>%{x}</b><br>' + col + ': %{y:,.0f}원<extra></extra>'
                if '이익률' in col or '%' in col:
                    hover_template = '<b>%{x}</b><br>' + col + ': %{y:.1f}%<extra></extra>'
                
                fig.add_trace(go.Bar(
                    x=df_sorted['상품명'],
                    y=df_sorted[col],
                    marker_color='#3b82f6',
                    hovertemplate=hover_template
                ))
                
                fig.update_layout(
                    title=f"{col} - 상위 20개 상품",
                    xaxis_title="상품명",
                    yaxis_title=col,
                    xaxis_tickangle=-45,
                    height=400
                )
                
                if '이익률' in col or '%' in col:
                    fig.update_yaxes(ticksuffix='%')
                else:
                    fig.update_yaxes(tickformat=',', ticksuffix='원')
                
                charts.append(dbc.Col([
                    dcc.Graph(figure=fig)
                ], md=6))
            
            # Histogram
            fig_hist = go.Figure()
            
            hover_template = col + ': %{x:,.0f}원<br>개수: %{y}<extra></extra>'
            if '이익률' in col or '%' in col:
                hover_template = col + ': %{x:.1f}%<br>개수: %{y}<extra></extra>'
            
            fig_hist.add_trace(go.Histogram(
                x=df_filtered[col],
                nbinsx=30,
                marker_color='#8b5cf6',
                hovertemplate=hover_template
            ))
            
            fig_hist.update_layout(
                title=f"{col} 분포",
                xaxis_title=col,
                yaxis_title="개수",
                height=400
            )
            
            if '이익률' in col or '%' in col:
                fig_hist.update_xaxes(ticksuffix='%')
            else:
                fig_hist.update_xaxes(tickformat=',', ticksuffix='원')
            
            charts.append(dbc.Col([
                dcc.Graph(figure=fig_hist)
            ], md=6))
    
    return dbc.Row(charts)

# ===========================
# Run Server
# ===========================

if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8050, debug=False)

