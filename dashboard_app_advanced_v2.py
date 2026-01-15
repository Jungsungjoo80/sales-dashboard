# ========================================
# Advanced Traffic Dashboard v7 - Complete
# ========================================
# 주요 개선사항:
# 1. 데이터 시각화: 매출 분포 제거, 전체 상품 표시, 네이버 녹색/쿠팡 파란색
# 2. 그래프 위 숫자: 만 단위 표기 (예: 4,000만), 겹침 방지
# 3. 비교주차 레이아웃: 1~4 주차 가로 배열
# 4. 상품 필터 기능 제거 (비교주차 시 상품 수 유지)
# 5. 비교주차 시각화: 가로 스크롤 지원
# 6. KPI 카드 확장: 8개 지표 모두 표시
# 7. 이익률 계산식: 순이익 / 트래픽비용
# 8. 이익률변동: % 표기 및 전주 대비 계산
# 9. ROAS 추가: 매출 / 트래픽비용
# 10. 마이너스 값: 빨간색 볼드 처리
# ========================================

import dash
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import base64
import io

# ========================================
# Dash App 초기화
# ========================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "LABONLAB 트래픽 데이터 분석 대시보드 v7"

# ========================================
# 헬퍼 함수들
# ========================================

def format_currency(value):
    """숫자를 원화 형식으로 포맷 (예: 10,000,000원)"""
    try:
        return f"{int(value):,}원"
    except:
        return "0원"

def format_currency_short_man(value):
    """숫자를 만 단위로 포맷 (예: 4,000만)"""
    try:
        man_value = int(value) / 10000
        return f"{man_value:,.0f}만"
    except:
        return "0만"

def format_percent(value):
    """숫자를 퍼센트로 포맷"""
    try:
        return f"{float(value):.1f}%"
    except:
        return "0.0%"

def calculate_derived_fields(df):
    """
    순이익, 이익률, 이익률변동, 슬롯수변동, ROAS 자동 계산
    - 이익률 = (순이익 / 트래픽비용) * 100
    - 이익률변동 = 전주 대비 이익률 변동(%)
    - ROAS = 매출 / 트래픽비용
    """
    # 숫자 컬럼 강제 변환
    numeric_cols = ['매출', '이익', '트래픽비용', '슬롯수']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # 순이익 계산
    if '이익' in df.columns and '트래픽비용' in df.columns:
        df['순이익'] = df['이익'] - df['트래픽비용']
    else:
        df['순이익'] = 0
    
    # 이익률 계산 (순이익 / 트래픽비용 * 100)
    if '순이익' in df.columns and '트래픽비용' in df.columns:
        df['이익률'] = df.apply(
            lambda row: (row['순이익'] / row['트래픽비용'] * 100) if row['트래픽비용'] > 0 else 0,
            axis=1
        )
    else:
        df['이익률'] = 0
    
    # ROAS 계산 (매출 / 트래픽비용)
    if 'ROAS' not in df.columns:  # 엑셀에 없으면 자동 계산
        if '매출' in df.columns and '트래픽비용' in df.columns:
            df['ROAS'] = df.apply(
                lambda row: (row['매출'] / row['트래픽비용']) if row['트래픽비용'] > 0 else 0,
                axis=1
            )
        else:
            df['ROAS'] = 0
    
    # 이익률변동 계산 (전주 대비 % 변화)
    if '상품명' in df.columns and '주차' in df.columns and '이익률' in df.columns:
        df = df.sort_values(['상품명', '주차'])
        df['이익률변동'] = df.groupby('상품명')['이익률'].pct_change() * 100
        df['이익률변동'] = df['이익률변동'].fillna(0)
    else:
        df['이익률변동'] = 0
    
    # 슬롯수변동 계산 (전주 대비 차이)
    if '상품명' in df.columns and '주차' in df.columns and '슬롯수' in df.columns:
        df = df.sort_values(['상품명', '주차'])
        df['슬롯수변동'] = df.groupby('상품명')['슬롯수'].diff().fillna(0)
    else:
        df['슬롯수변동'] = 0
    
    return df

def parse_uploaded_excel(contents, filename):
    """
    업로드된 Excel 파일 파싱
    - 중복 제거 로직 완전 제거 (모든 데이터 유지)
    - 계산 필드 자동 생성
    """
    try:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        
        # Excel 읽기
        df = pd.read_excel(io.BytesIO(decoded), header=0)
        
        # 완전히 빈 행만 제거
        df = df.dropna(how='all')
        
        # 계산 필드 생성
        df = calculate_derived_fields(df)
        
        # 주차/쇼핑몰 목록 자동 추출
        weeks = []
        malls = []
        
        if '주차' in df.columns:
            weeks = sorted(df['주차'].dropna().unique().tolist(), reverse=True)
        
        if '쇼핑몰' in df.columns:
            malls = ['전체'] + sorted(df['쇼핑몰'].dropna().unique().tolist())
        
        record_count = len(df)
        
        return df, weeks, malls, f"✓ {filename} 업로드 완료 ({record_count}개 레코드)"
    
    except Exception as e:
        return None, [], [], f"✗ 파일 파싱 오류: {str(e)}"

# ========================================
# 레이아웃
# ========================================

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
                    html.I(className="fas fa-cloud-upload-alt me-2"),
                    'LABONLAB 트래픽 데이터 Excel 파일을 드래그하거나 클릭하여 업로드'
                ]),
                style={
                    'width': '100%',
                    'height': '60px',
                    'lineHeight': '60px',
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
    
    # 필터 영역 (주차 선택 - 가로 배열)
    dbc.Row([
        dbc.Col([
            html.Label("📅 선택 주차", className="fw-bold mb-2"),
            dcc.Dropdown(id='selected-week', placeholder="주차 선택", className="mb-2")
        ], width=2),
        dbc.Col([
            html.Label("📊 비교주차 1", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-1', placeholder="비교주차 1", className="mb-2")
        ], width=2),
        dbc.Col([
            html.Label("📊 비교주차 2", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-2', placeholder="비교주차 2", className="mb-2")
        ], width=2),
        dbc.Col([
            html.Label("📊 비교주차 3", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-3', placeholder="비교주차 3", className="mb-2")
        ], width=2),
        dbc.Col([
            html.Label("📊 비교주차 4", className="fw-bold mb-2"),
            dcc.Dropdown(id='compare-week-4', placeholder="비교주차 4", className="mb-2")
        ], width=2),
        dbc.Col([
            html.Label("🏪 쇼핑몰 필터", className="fw-bold mb-2"),
            dcc.Dropdown(id='mall-filter', placeholder="쇼핑몰 선택", className="mb-2")
        ], width=2)
    ], className="mb-3"),
    
    # 컬럼 표시 선택 (가로 배열)
    dbc.Row([
        dbc.Col([
            html.Label("📋 표시할 컬럼 선택", className="fw-bold mb-2"),
            dcc.Checklist(
                id='column-selector',
                options=[],
                value=[],
                inline=True,
                className="mb-3",
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '15px'}
            )
        ])
    ], className="mb-3"),
    
    # KPI 카드 영역 (확장: 8개 지표)
    html.Div(id='kpi-cards', className="mb-4"),
    
    # 데이터 테이블
    dbc.Row([
        dbc.Col([
            html.H5("📊 데이터 테이블", className="mb-3"),
            html.Div([
                html.Label("페이지당 표시:", className="me-2"),
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
                    style={'width': '150px', 'display': 'inline-block'}
                )
            ], className="mb-3"),
            html.Div(id='data-table-container')
        ])
    ], className="mb-4"),
    
    # 통합 시각화 (비교주차 선택 시)
    dbc.Row([
        dbc.Col([
            html.Div(id='integrated-viz-container')
        ])
    ], className="mb-4"),
    
    # 차트 컨테이너
    html.Div(id='chart-container'),
    
    # 숨겨진 저장소
    dcc.Store(id='stored-data'),
    dcc.Store(id='stored-weeks'),
    dcc.Store(id='stored-malls')
    
], fluid=True, className="p-4")

# ========================================
# Callbacks
# ========================================

# Callback 1: 파일 업로드
@app.callback(
    [Output('stored-data', 'data'),
     Output('stored-weeks', 'data'),
     Output('stored-malls', 'data'),
     Output('upload-status', 'children')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def upload_file(contents, filename):
    """Excel 파일 업로드 및 파싱"""
    if contents is None:
        return None, [], [], ""
    
    df, weeks, malls, status_msg = parse_uploaded_excel(contents, filename)
    
    if df is not None:
        return df.to_dict('records'), weeks, malls, html.Div([
            html.I(className="fas fa-check-circle text-success me-2"),
            html.Span(status_msg, className="text-success fw-bold")
        ])
    else:
        return None, [], [], html.Div([
            html.I(className="fas fa-exclamation-circle text-danger me-2"),
            html.Span(status_msg, className="text-danger fw-bold")
        ])

# Callback 2: 주차 선택 드롭다운 업데이트
@app.callback(
    [Output('selected-week', 'options'),
     Output('compare-week-1', 'options'),
     Output('compare-week-2', 'options'),
     Output('compare-week-3', 'options'),
     Output('compare-week-4', 'options')],
    [Input('stored-weeks', 'data')]
)
def update_week_selectors(weeks):
    """주차 드롭다운 옵션 업데이트"""
    if not weeks:
        return [], [], [], [], []
    
    options = [{'label': week, 'value': week} for week in weeks]
    return options, options, options, options, options

# Callback 3: 쇼핑몰 필터 업데이트
@app.callback(
    Output('mall-filter', 'options'),
    [Input('stored-malls', 'data')]
)
def update_mall_filter(malls):
    """쇼핑몰 필터 옵션 업데이트"""
    if not malls:
        return []
    return [{'label': mall, 'value': mall} for mall in malls]

# Callback 4: 컬럼 선택 체크리스트 업데이트
@app.callback(
    [Output('column-selector', 'options'),
     Output('column-selector', 'value')],
    [Input('stored-data', 'data')]
)
def update_column_selector(data):
    """컬럼 선택 체크리스트 업데이트"""
    if not data:
        return [], []
    
    df = pd.DataFrame(data)
    
    # 숫자 컬럼 우선 정렬 (ROAS 포함)
    numeric_cols = ['매출', '이익', '트래픽비용', '순이익', '이익률', '이익률변동', 
                    '슬롯수', '슬롯수변동', 'ROAS']
    text_cols = ['상품명', '주차', '쇼핑몰', '특이사항', '의견']
    
    available_cols = [col for col in numeric_cols + text_cols if col in df.columns]
    
    options = [{'label': col, 'value': col} for col in available_cols]
    default_values = [col for col in ['상품명', '주차', '쇼핑몰', '매출', '이익', '순이익', '이익률', 'ROAS'] 
                      if col in available_cols]
    
    return options, default_values

# Callback 5: KPI 카드 업데이트 (8개 지표 확장)
@app.callback(
    Output('kpi-cards', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('mall-filter', 'value')]
)
def update_kpi_cards(data, selected_week, compare_week, mall_filter):
    """KPI 카드 업데이트 - 8개 지표 모두 표시"""
    if not data or not selected_week:
        return html.Div("데이터를 업로드하고 주차를 선택하세요.", className="text-muted text-center p-4")
    
    df = pd.DataFrame(data)
    
    # 필터 적용
    df_filtered = df[df['주차'] == selected_week].copy()
    if mall_filter and mall_filter != '전체':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 비교 주차 데이터
    compare_data = None
    if compare_week:
        compare_data = df[df['주차'] == compare_week].copy()
        if mall_filter and mall_filter != '전체':
            compare_data = compare_data[compare_data['쇼핑몰'] == mall_filter]
    
    # KPI 계산
    kpis = [
        {'label': '총 매출', 'value': df_filtered['매출'].sum(), 'icon': 'fa-won-sign', 'color': 'primary', 'format': 'currency'},
        {'label': '총 이익', 'value': df_filtered['이익'].sum(), 'icon': 'fa-chart-line', 'color': 'success', 'format': 'currency'},
        {'label': '트래픽 비용', 'value': df_filtered['트래픽비용'].sum(), 'icon': 'fa-bullhorn', 'color': 'warning', 'format': 'currency'},
        {'label': '순이익', 'value': df_filtered['순이익'].sum(), 'icon': 'fa-sack-dollar', 'color': 'info', 'format': 'currency'},
        {'label': '평균 이익률', 'value': df_filtered['이익률'].mean(), 'icon': 'fa-percent', 'color': 'secondary', 'format': 'percent'},
        {'label': '평균 이익률변동', 'value': df_filtered['이익률변동'].mean(), 'icon': 'fa-arrow-trend-up', 'color': 'danger', 'format': 'percent'},
        {'label': '총 슬롯수', 'value': df_filtered['슬롯수'].sum(), 'icon': 'fa-list-ol', 'color': 'dark', 'format': 'number'},
        {'label': '평균 ROAS', 'value': df_filtered['ROAS'].mean(), 'icon': 'fa-chart-pie', 'color': 'success', 'format': 'decimal'}
    ]
    
    cards = []
    for kpi in kpis:
        # 전주 대비 계산
        delta_text = ""
        delta_class = ""
        if compare_data is not None:
            if kpi['label'] == '총 매출':
                compare_val = compare_data['매출'].sum()
            elif kpi['label'] == '총 이익':
                compare_val = compare_data['이익'].sum()
            elif kpi['label'] == '트래픽 비용':
                compare_val = compare_data['트래픽비용'].sum()
            elif kpi['label'] == '순이익':
                compare_val = compare_data['순이익'].sum()
            elif kpi['label'] == '평균 이익률':
                compare_val = compare_data['이익률'].mean()
            elif kpi['label'] == '평균 이익률변동':
                compare_val = compare_data['이익률변동'].mean()
            elif kpi['label'] == '총 슬롯수':
                compare_val = compare_data['슬롯수'].sum()
            else:  # ROAS
                compare_val = compare_data['ROAS'].mean()
            
            if compare_val != 0:
                delta = ((kpi['value'] - compare_val) / compare_val) * 100
                delta_text = f"{delta:+.1f}%"
                # 마이너스 값은 빨간색 볼드
                delta_class = "text-success" if delta > 0 else "text-danger fw-bold"
        
        # 포맷팅 (마이너스 값 빨간색 볼드 처리)
        if kpi['format'] == 'currency':
            if kpi['value'] < 0:
                display_val = html.Span(format_currency(kpi['value']), className="text-danger fw-bold")
            else:
                display_val = format_currency(kpi['value'])
        elif kpi['format'] == 'percent':
            if kpi['value'] < 0:
                display_val = html.Span(format_percent(kpi['value']), className="text-danger fw-bold")
            else:
                display_val = format_percent(kpi['value'])
        elif kpi['format'] == 'decimal':
            if kpi['value'] < 0:
                display_val = html.Span(f"{kpi['value']:.2f}", className="text-danger fw-bold")
            else:
                display_val = f"{kpi['value']:.2f}"
        else:
            display_val = f"{int(kpi['value']):,}"
        
        card = dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.Div([
                        html.I(className=f"fas {kpi['icon']} fa-2x text-{kpi['color']} mb-2"),
                        html.H6(kpi['label'], className="text-muted mb-1"),
                        html.H4(display_val, className=f"text-{kpi['color']} mb-0"),
                        html.Small(delta_text, className=delta_class) if delta_text else html.Span()
                    ])
                ])
            ], className="shadow-sm h-100")
        ], width=3)
        cards.append(card)
    
    return dbc.Row(cards, className="mb-4")

# Callback 6: 데이터 테이블 업데이트
@app.callback(
    Output('data-table-container', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value'),
     Input('column-selector', 'value'),
     Input('page-size-selector', 'value')]
)
def update_data_table(data, selected_week, cw1, cw2, cw3, cw4, mall_filter, 
                      selected_columns, page_size):
    """데이터 테이블 업데이트 - 비교주차 병합 표시"""
    if not data or not selected_week:
        return html.Div("데이터를 업로드하고 주차를 선택하세요.", className="text-muted")
    
    df = pd.DataFrame(data)
    
    # 선택 주차 + 비교 주차 병합
    compare_weeks = [w for w in [cw1, cw2, cw3, cw4] if w]
    all_weeks = [selected_week] + compare_weeks
    
    df_filtered = df[df['주차'].isin(all_weeks)].copy()
    
    # 쇼핑몰 필터
    if mall_filter and mall_filter != '전체':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 컬럼 필터
    if selected_columns:
        available_cols = [col for col in selected_columns if col in df_filtered.columns]
        df_filtered = df_filtered[available_cols]
    
    # 포맷팅 (마이너스 값 빨간색 볼드)
    for col in df_filtered.columns:
        if col in ['매출', '이익', '트래픽비용', '순이익']:
            df_filtered[col] = df_filtered[col].apply(lambda x: format_currency(x) if pd.notna(x) else "")
        elif col in ['이익률', '이익률변동']:
            df_filtered[col] = df_filtered[col].apply(lambda x: format_percent(x) if pd.notna(x) else "")
        elif col == 'ROAS':
            df_filtered[col] = df_filtered[col].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "")
        elif col == '슬롯수':
            df_filtered[col] = df_filtered[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        elif col == '슬롯수변동':
            df_filtered[col] = df_filtered[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    
    # 데이터 테이블 생성 (마이너스 값 스타일링)
    table = dash_table.DataTable(
        data=df_filtered.to_dict('records'),
        columns=[{'name': col, 'id': col} for col in df_filtered.columns],
        page_size=page_size,
        page_action='native',
        sort_action='native',
        filter_action='native',
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '10px',
            'fontSize': '13px',
            'fontFamily': 'Noto Sans KR'
        },
        style_header={
            'backgroundColor': '#f1f5f9',
            'fontWeight': 'bold',
            'color': '#1e293b'
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8fafc'
            },
            # 마이너스 값 빨간색 볼드
            {
                'if': {
                    'filter_query': '{이익률} contains "-"',
                    'column_id': '이익률'
                },
                'color': '#dc2626',
                'fontWeight': 'bold'
            },
            {
                'if': {
                    'filter_query': '{이익률변동} contains "-"',
                    'column_id': '이익률변동'
                },
                'color': '#dc2626',
                'fontWeight': 'bold'
            },
            {
                'if': {
                    'filter_query': '{순이익} contains "-"',
                    'column_id': '순이익'
                },
                'color': '#dc2626',
                'fontWeight': 'bold'
            }
        ]
    )
    
    return table

# Callback 7: 통합 시각화 (비교주차 선택 시)
@app.callback(
    Output('integrated-viz-container', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value')]
)
def update_integrated_viz(data, selected_week, cw1, cw2, cw3, cw4, mall_filter):
    """통합 시각화 - 비교주차 선택 시 복합 차트 표시"""
    if not data or not selected_week:
        return html.Div()
    
    compare_weeks = [w for w in [cw1, cw2, cw3, cw4] if w]
    if not compare_weeks:
        return html.Div()
    
    df = pd.DataFrame(data)
    all_weeks = [selected_week] + compare_weeks
    
    # 필터 적용
    df_filtered = df[df['주차'].isin(all_weeks)].copy()
    if mall_filter and mall_filter != '전체':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    # 주차별 집계
    agg_data = df_filtered.groupby('주차').agg({
        '매출': 'sum',
        '이익': 'sum',
        '트래픽비용': 'sum',
        '순이익': 'sum',
        '이익률': 'mean',
        '이익률변동': 'mean',
        '슬롯수': 'sum',
        '슬롯수변동': 'sum'
    }).reset_index()
    
    # 복합 차트 생성 (매출은 바, 나머지는 선)
    fig = go.Figure()
    
    # 매출 바 차트
    fig.add_trace(go.Bar(
        x=agg_data['주차'],
        y=agg_data['매출'],
        name='매출',
        marker_color='#3b82f6',
        text=[format_currency_short_man(v) for v in agg_data['매출']],
        textposition='outside',
        yaxis='y1'
    ))
    
    # 나머지 지표들 선 차트
    metrics = [
        ('이익', '#10b981', 'y2'),
        ('트래픽비용', '#f59e0b', 'y2'),
        ('순이익', '#8b5cf6', 'y2'),
        ('슬롯수', '#ef4444', 'y3')
    ]
    
    for metric, color, yaxis in metrics:
        fig.add_trace(go.Scatter(
            x=agg_data['주차'],
            y=agg_data[metric],
            name=metric,
            mode='lines+markers+text',
            line=dict(color=color, width=3),
            marker=dict(size=8),
            text=[format_currency_short_man(v) if metric != '슬롯수' else f"{int(v):,}" for v in agg_data[metric]],
            textposition='top center',
            yaxis=yaxis
        ))
    
    # 레이아웃 (듀얼 Y축)
    fig.update_layout(
        title="📊 주차별 통합 시각화 (복합 차트)",
        xaxis=dict(title="주차"),
        yaxis=dict(title="매출 (원)", side='left'),
        yaxis2=dict(title="이익/비용 (원)", overlaying='y', side='right'),
        yaxis3=dict(title="슬롯수", overlaying='y', side='right', anchor='free', position=0.95),
        hovermode='x unified',
        height=450,
        margin=dict(l=50, r=150, t=80, b=50),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        plot_bgcolor='#f8fafc'
    )
    
    return html.Div([
        html.H5("📈 주차별 통합 시각화", className="mb-3"),
        dcc.Graph(figure=fig, config={'displayModeBar': False})
    ], className="mb-4")

# Callback 8: 차트 생성
@app.callback(
    Output('chart-container', 'children'),
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
    """차트 생성 - 전체 상품 표시, 네이버 녹색/쿠팡 파란색, 만 단위 표기, 가로 스크롤, 겹침 방지"""
    if not data or not selected_week or not selected_columns:
        return html.Div()
    
    df = pd.DataFrame(data)
    compare_weeks = [w for w in [cw1, cw2, cw3, cw4] if w]
    
    # 필터 적용
    df_filtered = df[df['주차'] == selected_week].copy()
    if mall_filter and mall_filter != '전체':
        df_filtered = df_filtered[df_filtered['쇼핑몰'] == mall_filter]
    
    charts = []
    
    # 숫자 컬럼만 차트로 표시
    numeric_cols = ['매출', '이익', '트래픽비용', '순이익', '이익률', '이익률변동', 
                    '슬롯수', '슬롯수변동', 'ROAS']
    chart_cols = [col for col in selected_columns if col in numeric_cols]
    
    # 쇼핑몰별 색상 매핑
    color_map = {'네이버': '#22c55e', '쿠팡': '#3b82f6'}  # 네이버 녹색, 쿠팡 파란색
    
    for col in chart_cols:
        if col not in df_filtered.columns:
            continue
        
        # 비교 모드 (여러 주차 비교)
        if compare_weeks:
            all_weeks = [selected_week] + compare_weeks
            df_compare = df[df['주차'].isin(all_weeks)].copy()
            
            if mall_filter and mall_filter != '전체':
                df_compare = df_compare[df_compare['쇼핑몰'] == mall_filter]
            
            # 전체 상품 표시 (상위 20개 제한 제거)
            products = df_compare['상품명'].unique().tolist()
            
            fig = go.Figure()
            
            for week_idx, week in enumerate(all_weeks):
                week_data = df_compare[df_compare['주차'] == week]
                
                # 쇼핑몰별로 그룹화
                for mall in week_data['쇼핑몰'].unique():
                    mall_data = week_data[week_data['쇼핑몰'] == mall]
                    bar_color = color_map.get(mall, '#94a3b8')
                    
                    # 텍스트 위치 조정 (겹침 방지)
                    text_values = []
                    for v in mall_data[col]:
                        if col in ['매출', '이익', '트래픽비용', '순이익']:
                            text_values.append(format_currency_short_man(v))
                        elif col in ['이익률', '이익률변동']:
                            text_values.append(format_percent(v))
                        elif col == 'ROAS':
                            text_values.append(f"{v:.2f}")
                        else:
                            text_values.append(f"{int(v):,}")
                    
                    fig.add_trace(go.Bar(
                        x=mall_data['상품명'],
                        y=mall_data[col],
                        name=f"{week} - {mall}",
                        marker_color=bar_color,
                        text=text_values,
                        textposition='outside',
                        textangle=0,
                        textfont=dict(size=9),  # 글씨 크기 축소
                        hovertemplate='<b>%{x}</b><br>' + col + ': %{y:,.0f}<extra></extra>'
                    ))
            
            # 가로 스크롤 지원 (제품 수에 따라 너비 자동 조정)
            chart_width = max(1200, len(products) * 100)
            
            fig.update_layout(
                title=f"📊 {col} 비교 (전체 상품, 주차별)",
                xaxis_title="상품명",
                yaxis_title=col,
                barmode='group',
                height=600,
                width=chart_width,
                hovermode='x unified',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                plot_bgcolor='#f8fafc',
                xaxis=dict(tickangle=-45, tickfont=dict(size=10)),
                uniformtext=dict(mode='hide', minsize=8)  # 텍스트 겹침 방지
            )
            
        else:
            # 단일 주차 모드
            df_sorted = df_filtered.sort_values(col, ascending=False)  # 전체 상품 표시
            
            # 쇼핑몰별 색상 매핑
            colors = [color_map.get(mall, '#94a3b8') for mall in df_sorted['쇼핑몰']]
            
            # 텍스트 값 생성
            text_values = []
            for v in df_sorted[col]:
                if col in ['매출', '이익', '트래픽비용', '순이익']:
                    text_values.append(format_currency_short_man(v))
                elif col in ['이익률', '이익률변동']:
                    text_values.append(format_percent(v))
                elif col == 'ROAS':
                    text_values.append(f"{v:.2f}")
                else:
                    text_values.append(f"{int(v):,}")
            
            fig = go.Figure(go.Bar(
                x=df_sorted['상품명'],
                y=df_sorted[col],
                marker_color=colors,
                text=text_values,
                textposition='outside',
                textangle=0,
                textfont=dict(size=9),
                hovertemplate='<b>%{x}</b><br>' + col + ': %{y:,.0f}<extra></extra>'
            ))
            
            # 가로 스크롤 지원
            chart_width = max(1200, len(df_sorted) * 60)
            
            fig.update_layout(
                title=f"📊 {col} (전체 상품, {selected_week})",
                xaxis_title="상품명",
                yaxis_title=col,
                height=600,
                width=chart_width,
                hovermode='x unified',
                plot_bgcolor='#f8fafc',
                xaxis=dict(tickangle=-45, tickfont=dict(size=10)),
                uniformtext=dict(mode='hide', minsize=8)
            )
        
        # 차트를 가로 스크롤 가능한 컨테이너에 담기
        chart_div = html.Div([
            html.H5(f"📈 {col}", className="mb-3"),
            html.Div([
                dcc.Graph(figure=fig, config={'displayModeBar': False})
            ], style={'overflowX': 'scroll', 'width': '100%'})
        ], className="mb-4")
        
        charts.append(chart_div)
    
    return html.Div(charts)

# ========================================
# 서버 실행
# ========================================
if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8050, debug=False)
