# -*- coding: utf-8 -*-
"""
LABONLAB 트래픽 데이터 분석 대시보드 v7 (베이스라인 기반 수정본)
요청 반영:
1) 숫자 잘림/겹침 해결(공통 margin/cliponaxis/uniformtext/라벨 정책)
2) 상품 필터 추가(쇼핑몰 필터처럼) + 모든 시각화/테이블 연동
3) 회의록 인사이트 반영(상품 단위 관리 강화 → 상품 필터 + 라벨 정책 UI)

주의:
- "모든 막대에 무조건 숫자 표시"는 데이터가 많을 경우 겹침이 구조적으로 발생합니다.
  따라서 라벨 정책 토글(전체/상위N/없음)을 추가했습니다.
"""

import dash
from dash import dcc, html, dash_table, Input, Output, State
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objs as go
import base64
import io

# ====================
# App Init
# ====================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "LABONLAB 트래픽 데이터 분석 대시보드 v7"

# ====================
# Helper Functions
# ====================

def format_currency(value):
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
    try:
        val = float(value)
        if pd.isna(val):
            return "0%"
        return f"{val:.1f}%"
    except:
        return "0%"

def calculate_derived_fields(df):
    df = df.copy()

    numeric_cols = ['매출', '이익', '트래픽비용', '슬롯수']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    if '이익' in df.columns and '트래픽비용' in df.columns:
        df['순이익'] = df['이익'] - df['트래픽비용']

    if '순이익' in df.columns and '트래픽비용' in df.columns:
        df['이익률'] = df.apply(
            lambda row: (row['순이익'] / row['트래픽비용'] * 100) if row['트래픽비용'] > 0 else 0,
            axis=1
        )

    # ROAS = (매출 / 트래픽비용) * 100
    if '매출' in df.columns and '트래픽비용' in df.columns:
        df['ROAS'] = df.apply(
            lambda row: (row['매출'] / row['트래픽비용'] * 100) if row['트래픽비용'] > 0 else 0,
            axis=1
        )

    # 전주 대비 변화
    if '주차' in df.columns and '이익률' in df.columns:
        df = df.sort_values(['상품명', '쇼핑몰', '주차'])
        df['이익률변동'] = df.groupby(['상품명', '쇼핑몰'])['이익률'].diff()

    if '주차' in df.columns and '슬롯수' in df.columns:
        df = df.sort_values(['상품명', '쇼핑몰', '주차'])
        df['슬롯수변동'] = df.groupby(['상품명', '쇼핑몰'])['슬롯수'].diff()

    return df

def parse_uploaded_excel(contents):
    try:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        df = pd.read_excel(io.BytesIO(decoded))

        required_cols = ['상품명', '주차', '쇼핑몰', '매출', '이익', '트래픽비용']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return None, f"필수 컬럼 누락: {', '.join(missing)}"

        optional_cols = ['슬롯수', '특이사항', '의견']
        for col in optional_cols:
            if col not in df.columns:
                df[col] = ""

        df = calculate_derived_fields(df)
        return df, None
    except Exception as e:
        return None, f"파일 읽기 오류: {str(e)}"

# ---------- 공통 Figure 스타일(숫자 잘림/겹침 방지 기본값) ----------
def apply_common_figure_style(fig, height=500):
    """
    공통 레이아웃/겹침 방지 전략:
    - 상단 마진 확장(잘림 방지)
    - automargin 활성화
    - uniformtext: 너무 작은 텍스트는 자동 숨김
    """
    fig.update_layout(
        height=height,
        margin=dict(t=160, b=120, l=80, r=60),
        hovermode='x unified',
        uniformtext_minsize=10,
        uniformtext_mode='hide',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    fig.update_xaxes(automargin=True, tickangle=35)
    fig.update_yaxes(automargin=True)
    return fig

def get_label_text_and_position(df_vals, col, label_mode, top_n=5):
    """
    label_mode:
      - 'all': 모든 값 text 표시 (겹침 가능성 매우 높음)
      - 'top': 상위 N개만 text 표시(권장)
      - 'none': text 표시 안함(hover만)
    """
    # text values
    if col in ['매출', '이익', '트래픽비용', '순이익']:
        texts = [format_currency_short_man(v) for v in df_vals]
    elif col in ['이익률', '이익률변동', 'ROAS']:
        texts = [format_percent(v) for v in df_vals]
    else:
        texts = [f"{int(v):,}" if pd.notna(v) else "0" for v in df_vals]

    if label_mode == 'none':
        return texts, 'none', None

    if label_mode == 'all':
        # 전부 표시하되, 공통 스타일 + cliponaxis로 잘림 방지
        return texts, 'outside', None

    # label_mode == 'top': 상위 N개만 표시
    # df_vals 기준 상위 N index
    s = pd.Series(df_vals).fillna(0)
    top_idx = set(s.nlargest(min(top_n, len(s))).index.tolist())
    # Plotly bar는 trace-level textposition만 가능 → "보이는 텍스트만 넣고 나머지는 ''"로 처리
    masked_texts = [texts[i] if i in top_idx else "" for i in range(len(texts))]
    return masked_texts, 'outside', top_idx

# ====================
# Layout
# ====================

app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.H1("📊 LABONLAB 트래픽 데이터 분석 대시보드 v7", className="text-primary mb-2"),
            html.P("Excel 파일을 업로드하여 주차별/쇼핑몰별 데이터를 분석하세요", className="text-muted")
        ], width=8),
        dbc.Col([html.Div(id='upload-status', className="text-end")], width=4)
    ], className="mb-4 mt-4"),

    dbc.Row([
        dbc.Col([
            dcc.Upload(
                id='upload-data',
                children=html.Div(['Excel 파일을 드래그하거나 클릭하여 업로드']),
                style={
                    'width': '100%', 'height': '80px', 'lineHeight': '80px',
                    'borderWidth': '2px', 'borderStyle': 'dashed', 'borderRadius': '10px',
                    'textAlign': 'center', 'backgroundColor': '#f8f9fa', 'cursor': 'pointer'
                },
                multiple=False
            )
        ])
    ], className="mb-4"),

    # ---- Filters Row 1 ----
    dbc.Row([
        dbc.Col([
            html.Label("📅 선택 주차", className="fw-bold mb-2"),
            dcc.Dropdown(id='selected-week', placeholder="주차 선택", clearable=False)
        ], width=3),

        dbc.Col([
            html.Label("🏪 쇼핑몰 필터", className="fw-bold mb-2"),
            dcc.Dropdown(id='mall-filter', value='all', clearable=False)
        ], width=3),

        dbc.Col([
            html.Label("🧴 상품 필터(복수 선택)", className="fw-bold mb-2"),
            dcc.Dropdown(
                id='product-filter',
                placeholder="전체(선택 안 함) / 검색 후 복수 선택",
                multi=True,
                clearable=True
            )
        ], width=6),
    ], className="mb-3"),

    # ---- Filters Row 2 ----
    dbc.Row([
        dbc.Col([
            html.Label("📊 표시할 컬럼 선택", className="fw-bold mb-2"),
            dcc.Checklist(
                id='column-selector',
                options=[],
                value=[],
                inline=True,
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '10px'}
            )
        ], width=8),

        dbc.Col([
            html.Label("🏷️ 라벨 표시 방식(겹침 방지)", className="fw-bold mb-2"),
            dcc.RadioItems(
                id='label-mode',
                options=[
                    {'label': '상위 5개만(권장)', 'value': 'top'},
                    {'label': '전체 표시(겹침 가능)', 'value': 'all'},
                    {'label': '없음(hover만)', 'value': 'none'},
                ],
                value='top',
                inline=True,
                style={'display': 'flex', 'gap': '14px', 'flexWrap': 'wrap'}
            )
        ], width=4),
    ], className="mb-4"),

    # Compare weeks
    dbc.Row([
        dbc.Col([html.Label("📊 비교 주차 1", className="fw-bold mb-2"),
                 dcc.Dropdown(id='compare-week-1', placeholder="선택 안함")], width=3),
        dbc.Col([html.Label("📊 비교 주차 2", className="fw-bold mb-2"),
                 dcc.Dropdown(id='compare-week-2', placeholder="선택 안함")], width=3),
        dbc.Col([html.Label("📊 비교 주차 3", className="fw-bold mb-2"),
                 dcc.Dropdown(id='compare-week-3', placeholder="선택 안함")], width=3),
        dbc.Col([html.Label("📊 비교 주차 4", className="fw-bold mb-2"),
                 dcc.Dropdown(id='compare-week-4', placeholder="선택 안함")], width=3),
    ], className="mb-4"),

    dbc.Row(id='kpi-cards', className="mb-4"),

    # Data Table
    dbc.Row([
        dbc.Col([
            html.H5("📋 데이터 테이블", className="mb-3"),
            html.Div([
                html.Label("페이지당 행 수: ", className="me-2"),
                dcc.Dropdown(
                    id='page-size',
                    options=[{'label': str(v), 'value': v} for v in [10,20,30,50,100]],
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
                style_cell={'textAlign': 'left','padding': '10px','fontSize': '13px','fontFamily': 'Arial, sans-serif'},
                style_header={'backgroundColor': '#e9ecef','fontWeight': 'bold','textAlign': 'center'},
                style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#f8f9fa'}]
            )
        ])
    ], className="mb-4"),

    dbc.Row([
        dbc.Col([
            html.H5("📊 주차별 통합 시각화 (복수 주차 선택 시)", className="mb-3"),
            dcc.Graph(id='integrated-viz', style={'height': '450px'})
        ])
    ], className="mb-4"),

    html.Div(id='charts-container'),

    dcc.Store(id='stored-data'),
    dcc.Store(id='stored-weeks'),
    dcc.Store(id='stored-malls'),
    dcc.Store(id='stored-products'),
], fluid=True, style={'backgroundColor': '#f8f9fa', 'minHeight': '100vh', 'padding': '20px'})


# ====================
# Callbacks
# ====================

@app.callback(
    [Output('stored-data', 'data'),
     Output('stored-weeks', 'data'),
     Output('stored-malls', 'data'),
     Output('stored-products', 'data'),
     Output('upload-status', 'children')],
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def upload_file(contents, filename):
    if contents is None:
        return None, None, None, None, ""

    df, error = parse_uploaded_excel(contents)
    if error:
        return None, None, None, None, html.Div([
            html.Span(error, className="text-danger")
        ])

    weeks = sorted(df['주차'].unique().tolist())
    malls = sorted(df['쇼핑몰'].unique().tolist())
    products = sorted(df['상품명'].unique().tolist())

    status = html.Div([
        html.Span(f"✓ {filename} 업로드 완료 ({len(df)}개 레코드)", className="text-success fw-bold")
    ])

    return df.to_dict('records'), weeks, malls, products, status


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
    default_week = weeks[-1]
    return options, default_week, options, options, options, options


@app.callback(
    Output('mall-filter', 'options'),
    Input('stored-malls', 'data')
)
def update_mall_filter(malls):
    if not malls:
        return [{'label': '전체', 'value': 'all'}]
    opts = [{'label': '전체', 'value': 'all'}]
    opts += [{'label': m, 'value': m} for m in malls]
    return opts


@app.callback(
    [Output('product-filter', 'options'),
     Output('product-filter', 'value')],
    Input('stored-products', 'data'),
)
def update_product_filter(products):
    if not products:
        return [], []
    # value=[] => 전체 의미
    opts = [{'label': p, 'value': p} for p in products]
    return opts, []


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


def apply_filters(df, weeks_to_show, mall_filter, product_filter):
    dff = df[df['주차'].isin(weeks_to_show)].copy()

    if mall_filter and mall_filter != 'all':
        dff = dff[dff['쇼핑몰'] == mall_filter]

    # product_filter: [] 또는 None이면 전체
    if product_filter:
        dff = dff[dff['상품명'].isin(product_filter)]

    return dff


@app.callback(
    Output('kpi-cards', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('mall-filter', 'value'),
     Input('product-filter', 'value')]
)
def update_kpi_cards(data, selected_week, compare_week, mall_filter, product_filter):
    if not data or not selected_week:
        return []

    df = pd.DataFrame(data)

    df_filtered = apply_filters(df, [selected_week], mall_filter, product_filter)

    df_compare = None
    if compare_week:
        df_compare = apply_filters(df, [compare_week], mall_filter, product_filter)

    kpis = [
        {'title': '총 매출','value': df_filtered['매출'].sum(),'format': 'currency','color': 'primary',
         'compare': df_compare['매출'].sum() if df_compare is not None else None},
        {'title': '총 이익','value': df_filtered['이익'].sum(),'format': 'currency','color': 'success',
         'compare': df_compare['이익'].sum() if df_compare is not None else None},
        {'title': '트래픽 비용','value': df_filtered['트래픽비용'].sum(),'format': 'currency','color': 'warning',
         'compare': df_compare['트래픽비용'].sum() if df_compare is not None else None},
        {'title': '순이익','value': df_filtered['순이익'].sum(),'format': 'currency','color': 'info',
         'compare': df_compare['순이익'].sum() if df_compare is not None else None},
        {'title': '평균 이익률','value': df_filtered['이익률'].mean(),'format': 'percent','color': 'secondary',
         'compare': df_compare['이익률'].mean() if df_compare is not None else None},
        {'title': '평균 이익률변동','value': df_filtered['이익률변동'].mean(),'format': 'percent','color': 'dark',
         'compare': None},
        {'title': '총 슬롯수','value': df_filtered['슬롯수'].sum(),'format': 'number','color': 'danger',
         'compare': df_compare['슬롯수'].sum() if df_compare is not None else None},
        {'title': '평균 ROAS','value': df_filtered['ROAS'].mean(),'format': 'percent','color': 'primary',
         'compare': df_compare['ROAS'].mean() if df_compare is not None else None},
    ]

    cards = []
    for k in kpis:
        if k['format'] == 'currency':
            main_val = format_currency(k['value'])
        elif k['format'] == 'percent':
            main_val = format_percent(k['value'])
        else:
            main_val = f"{int(k['value']):,}" if pd.notna(k['value']) else "0"

        delta_text = ""
        delta_class = "text-muted"
        if k['compare'] is not None and pd.notna(k['compare']):
            delta = k['value'] - k['compare']
            if k['format'] == 'currency':
                dt = format_currency(abs(delta))
            elif k['format'] == 'percent':
                dt = format_percent(abs(delta))
            else:
                dt = f"{int(abs(delta)):,}"

            if delta > 0:
                delta_text = f"▲ {dt}"
                delta_class = "text-success"
            elif delta < 0:
                delta_text = f"▼ {dt}"
                delta_class = "text-danger fw-bold"
            else:
                delta_text = "→ 변동 없음"

        cards.append(
            dbc.Col(
                dbc.Card(
                    dbc.CardBody([
                        html.H6(k['title'], className="text-muted mb-2"),
                        html.H4(main_val, className=f"text-{k['color']} fw-bold mb-1"),
                        html.Small(delta_text, className=delta_class) if delta_text else html.Span()
                    ]),
                    className="shadow-sm h-100"
                ),
                width=12, md=6, lg=3, className="mb-3"
            )
        )

    return cards


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
     Input('product-filter', 'value'),
     Input('column-selector', 'value'),
     Input('page-size', 'value')]
)
def update_data_table(data, selected_week, cw1, cw2, cw3, cw4, mall_filter, product_filter, selected_columns, page_size):
    if not data or not selected_week:
        return [], [], page_size

    df = pd.DataFrame(data)

    weeks_to_show = [selected_week]
    for cw in [cw1, cw2, cw3, cw4]:
        if cw and cw not in weeks_to_show:
            weeks_to_show.append(cw)

    df_filtered = apply_filters(df, weeks_to_show, mall_filter, product_filter)

    column_order = ['상품명', '주차', '쇼핑몰', '매출', '이익', '트래픽비용', '순이익', 'ROAS',
                    '이익률', '이익률변동', '슬롯수', '슬롯수변동', '특이사항', '의견']

    display_cols = [col for col in column_order if col in df_filtered.columns]

    if selected_columns:
        base_cols = ['상품명', '주차', '쇼핑몰']
        display_cols = base_cols + [c for c in display_cols if c in selected_columns or c in base_cols]

    df_display = df_filtered[display_cols].copy()

    for col in df_display.columns:
        if col in ['매출', '이익', '트래픽비용', '순이익']:
            df_display[col] = df_display[col].apply(format_currency)
        elif col in ['이익률', '이익률변동', 'ROAS']:
            df_display[col] = df_display[col].apply(format_percent)
        elif col in ['슬롯수', '슬롯수변동']:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "0")

    columns = [{'name': col, 'id': col} for col in display_cols]
    return df_display.to_dict('records'), columns, page_size


# -------- 통합 시각화: column-selector 반영 + 상품 필터 반영 + 라벨 정책 반영 --------
@app.callback(
    Output('integrated-viz', 'figure'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value'),
     Input('product-filter', 'value'),
     Input('column-selector', 'value'),
     Input('label-mode', 'value')]
)
def update_integrated_viz(data, selected_week, cw1, cw2, cw3, cw4, mall_filter, product_filter, selected_columns, label_mode):
    if not data or not selected_week:
        return go.Figure()

    df = pd.DataFrame(data)

    compare_weeks = [cw for cw in [cw1, cw2, cw3, cw4] if cw]
    if not compare_weeks:
        fig = go.Figure()
        fig.update_layout(title="비교 주차를 선택하면 통합 시각화가 표시됩니다", height=450)
        return fig

    all_weeks = [selected_week] + compare_weeks
    df_filtered = apply_filters(df, all_weeks, mall_filter, product_filter)

    # 컬럼 선택이 비어있으면 기본값(안전)
    if not selected_columns:
        default_cols = ['매출', '이익', '트래픽비용', '순이익', 'ROAS', '이익률', '슬롯수']
        selected_columns = [c for c in default_cols if c in df_filtered.columns]
        if not selected_columns:
            selected_columns = ['매출'] if '매출' in df_filtered.columns else []

    # 집계 규칙
    agg_dict = {}
    for c in selected_columns:
        if c in ['매출', '이익', '트래픽비용', '순이익', '슬롯수', '슬롯수변동']:
            agg_dict[c] = 'sum'
        else:
            agg_dict[c] = 'mean'

    if not agg_dict:
        fig = go.Figure()
        fig.update_layout(title="표시할 수 있는 컬럼이 없습니다(데이터 컬럼 확인 필요)", height=450)
        return fig

    df_agg = df_filtered.groupby('주차').agg(agg_dict).reset_index()
    df_agg['주차'] = pd.Categorical(df_agg['주차'], categories=all_weeks, ordered=True)
    df_agg = df_agg.sort_values('주차')

    fig = go.Figure()

    col_styles = {
        '매출': {'color': '#3b82f6', 'type': 'bar', 'yaxis': 'y'},
        '이익': {'color': '#10b981', 'type': 'line', 'yaxis': 'y2'},
        '트래픽비용': {'color': '#f59e0b', 'type': 'line', 'yaxis': 'y2'},
        '순이익': {'color': '#ef4444', 'type': 'line', 'yaxis': 'y2'},
        '슬롯수': {'color': '#8b5cf6', 'type': 'line', 'yaxis': 'y2'},
        'ROAS': {'color': '#ec4899', 'type': 'line', 'yaxis': 'y2'},
        '이익률': {'color': '#06b6d4', 'type': 'line', 'yaxis': 'y2'},
        '이익률변동': {'color': '#84cc16', 'type': 'line', 'yaxis': 'y2'},
        '슬롯수변동': {'color': '#f97316', 'type': 'line', 'yaxis': 'y2'}
    }

    # 겹침 방지를 위해 trace별 기본 textposition을 다르게
    text_positions = ['top center', 'bottom center', 'middle left', 'middle right', 'top left', 'top right', 'bottom left', 'bottom right']

    for idx, col in enumerate(selected_columns):
        if col not in df_agg.columns:
            continue

        style = col_styles.get(col, {'color': '#6366f1', 'type': 'line', 'yaxis': 'y2'})

        # 라벨 정책 적용(통합시각화는 데이터포인트가 적으므로 all도 비교적 안전)
        texts, _, _ = get_label_text_and_position(df_agg[col].tolist(), col, label_mode, top_n=5)
        text_pos = text_positions[idx % len(text_positions)]

        if style['type'] == 'bar':
            fig.add_trace(go.Bar(
                x=df_agg['주차'],
                y=df_agg[col],
                name=col,
                marker_color=style['color'],
                text=texts,
                textposition=text_pos,
                cliponaxis=False,
                yaxis=style['yaxis'],
                hovertemplate=f'<b>%{{x}}</b><br>{col}: %{{text}}<extra></extra>'
            ))
        else:
            fig.add_trace(go.Scatter(
                x=df_agg['주차'],
                y=df_agg[col],
                name=col,
                mode='lines+markers+text',
                line=dict(color=style['color'], width=3),
                marker=dict(size=10),
                text=texts,
                textposition=text_pos,
                cliponaxis=False,
                yaxis=style['yaxis'],
                hovertemplate=f'<b>%{{x}}</b><br>{col}: %{{text}}<extra></extra>'
            ))

    # y축 포맷(금액/퍼센트 혼합이므로 기본은 tickformat 분리 최소화)
    fig.update_layout(
        title="주차별 통합 시각화",
        xaxis=dict(title="주차"),
        yaxis=dict(
            title="매출(등 금액 지표)",
            ticksuffix='만',
            autorange=True
        ),
        yaxis2=dict(
            title="기타 지표",
            overlaying='y',
            side='right',
            autorange=True
        )
    )
    fig = apply_common_figure_style(fig, height=450)
    return fig


# -------- 개별 차트: 상품 필터 + 라벨 정책 + 공통 겹침 방지 --------
@app.callback(
    Output('charts-container', 'children'),
    [Input('stored-data', 'data'),
     Input('selected-week', 'value'),
     Input('compare-week-1', 'value'),
     Input('compare-week-2', 'value'),
     Input('compare-week-3', 'value'),
     Input('compare-week-4', 'value'),
     Input('mall-filter', 'value'),
     Input('product-filter', 'value'),
     Input('column-selector', 'value'),
     Input('label-mode', 'value')]
)
def update_charts(data, selected_week, cw1, cw2, cw3, cw4, mall_filter, product_filter, selected_columns, label_mode):
    if not data or not selected_week or not selected_columns:
        return []

    df = pd.DataFrame(data)

    compare_weeks = [cw for cw in [cw1, cw2, cw3, cw4] if cw]
    all_weeks = [selected_week] + compare_weeks

    df_filtered = apply_filters(df, all_weeks, mall_filter, product_filter)

    color_map = {'네이버': '#22c55e', '쿠팡': '#3b82f6'}
    numeric_cols = ['매출', '이익', '트래픽비용', '순이익', '이익률', '이익률변동', '슬롯수', '슬롯수변동', 'ROAS']
    selected_numeric = [c for c in selected_columns if c in numeric_cols and c in df_filtered.columns]

    charts = []

    for col in selected_numeric:
        if compare_weeks:
            # 다중 주차 비교 모드
            df_chart = df_filtered.groupby(['주차', '상품명', '쇼핑몰'])[col].sum().reset_index()

            unique_products = df_chart['상품명'].nunique()
            chart_width = max(1200, unique_products * 80)

            fig = go.Figure()

            for week in all_weeks:
                df_week = df_chart[df_chart['주차'] == week]
                for mall in df_week['쇼핑몰'].unique():
                    df_mall = df_week[df_week['쇼핑몰'] == mall].copy()

                    yvals = df_mall[col].tolist()
                    texts, textpos, _ = get_label_text_and_position(yvals, col, label_mode, top_n=5)

                    fig.add_trace(go.Bar(
                        x=df_mall['상품명'],
                        y=df_mall[col],
                        name=f"{week} - {mall}",
                        marker_color=color_map.get(mall, '#6366f1'),
                        text=texts,
                        textposition=textpos,
                        cliponaxis=False,
                        hovertemplate='<b>%{x}</b><br>%{text}<extra></extra>'
                    ))

            # y축 suffix
            if col in ['매출', '이익', '트래픽비용', '순이익']:
                yaxis_format = dict(ticksuffix='만')
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
                width=chart_width
            )
            fig.update_yaxes(autorange=True)
            fig = apply_common_figure_style(fig, height=520)

        else:
            # 단일 주차 모드(상위 20개)
            df_chart = df_filtered.groupby(['상품명', '쇼핑몰'])[col].sum().reset_index()
            df_chart = df_chart.nlargest(20, col)

            chart_width = max(1200, len(df_chart) * 60)
            fig = go.Figure()

            for mall in df_chart['쇼핑몰'].unique():
                df_mall = df_chart[df_chart['쇼핑몰'] == mall].copy()
                yvals = df_mall[col].tolist()
                texts, textpos, _ = get_label_text_and_position(yvals, col, label_mode, top_n=5)

                fig.add_trace(go.Bar(
                    x=df_mall['상품명'],
                    y=df_mall[col],
                    name=mall,
                    marker_color=color_map.get(mall, '#6366f1'),
                    text=texts,
                    textposition=textpos,
                    cliponaxis=False,
                    hovertemplate='<b>%{x}</b><br>%{text}<extra></extra>'
                ))

            if col in ['매출', '이익', '트래픽비용', '순이익']:
                yaxis_format = dict(ticksuffix='만')
            elif col in ['이익률', '이익률변동', 'ROAS']:
                yaxis_format = dict(ticksuffix='%')
            else:
                yaxis_format = dict(tickformat=',')

            fig.update_layout(
                title=f"{col} (상위 20개 상품)",
                xaxis_title="상품명",
                yaxis_title=col,
                yaxis=yaxis_format,
                width=chart_width
            )
            fig.update_yaxes(autorange=True)
            fig = apply_common_figure_style(fig, height=520)

        chart_div = html.Div([
            html.H5(f"📊 {col}", className="mb-3"),
            html.Div([dcc.Graph(figure=fig, config={'displayModeBar': False})],
                     style={'overflowX': 'auto', 'overflowY': 'hidden'})
        ], className="mb-4")

        charts.append(chart_div)

    return charts


if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8050, debug=False)
