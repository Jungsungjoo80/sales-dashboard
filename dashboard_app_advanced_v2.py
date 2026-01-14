"""
Advanced Sales Dashboard v5
개선사항:
- 모든 금액 표시를 천단위 콤마 + 원화 형식으로 변경
- 차트 hover 시 원화 표시
- 누락된 컬럼 추가 (쇼핑몰, 슬롯수, 특이사항, 의견)
- 슬롯수변동을 숫자로 표시
- 비교 주차 데이터 테이블 병합
- 비교 주차 차트 동시 표시 (grouped bar)
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

# Initialize Dash app
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

def parse_uploaded_excel(contents, filename):
    """
    Excel 파일 파싱
    - B열에서 주차 정보 추출
    - 동적 헤더 감지
    - 중복 제거 및 데이터 정제
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
        
        # Remove duplicate rows based on 상품명 and 주차
        if '상품명' in df.columns and '주차' in df.columns:
            df = df.drop_duplicates(subset=['상품명', '주차'], keep='first')
        
        # Extract unique weeks from B column (주차)
        weeks = []
        if '주차' in df.columns:
            weeks = sorted(df['주차'].dropna().unique().tolist(), reverse=True)
        
        return df, weeks, None
        
    except Exception as e:
        return None, [], f"파일 파싱 오류: {str(e)}"

def create_kpi_card(title, value, icon="📊", color="primary"):
    """KPI 카드 생성"""
    return dbc.Card([
        dbc.CardBody([
            html.H6(title, className="text-muted mb-2"),
            html.H3(value, className=f"text-{color} mb-0", style={"fontSize": "20px"}),
            html.P(icon, className="mb-0", style={"fontSize": "24px"})
        ])
    ], className="shadow-sm")

# ===========================
# Layout
# ===========================

app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.H1("📊 LABONLAB 트래픽 데이터 분석 대시보드", className="text-primary mb-3"),
            html.P("Excel 파일을 업로드하여 주차별 데이터를 분석하세요", className="text-muted")
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
    
    # Filter Controls
    dbc.Row([
        dbc.Col([
            html.Label("주차 선택", className="fw-bold"),
            dcc.Dropdown(
                id='week-selector',
                placeholder="주차를 선택하세요",
                className="mb-3"
            )
        ], md=4),
        dbc.Col([
            html.Label("비교 주차 선택 (선택사항)", className="fw-bold"),
            dcc.Dropdown(
                id='compare-week-selector',
                placeholder="비교할 주차를 선택하세요",
                className="mb-3"
            )
        ], md=4),
        dbc.Col([
            html.Label("표시할 컬럼 선택", className="fw-bold"),
            dcc.Checklist(
                id='column-selector',
                options=[],
                value=[],
                inline=False,
                className="mb-3"
            )
        ], md=4)
    ], className="mb-4", id="filter-section", style={"display": "none"}),
    
    # KPI Cards
    dbc.Row([
        dbc.Col(html.Div(id='kpi-cards'), md=12)
    ], className="mb-4"),
    
    # Data Table
    dbc.Row([
        dbc.Col([
            html.H4("📋 데이터 테이블", className="mb-3"),
            html.Div(id='data-table-container')
        ])
    ], className="mb-4"),
    
    # Charts Section
    dbc.Row([
        dbc.Col([
            html.H4("📈 데이터 시각화", className="mb-3"),
            html.Div(id='charts-container')
        ])
    ], className="mb-4"),
    
    # Hidden data store
    dcc.Store(id='stored-data'),
    dcc.Store(id='stored-weeks')
    
], fluid=True, style={"backgroundColor": "#f5f5f5", "minHeight": "100vh"})

# ===========================
# Callbacks
# ===========================

@app.callback(
    [Output('stored-data', 'data'),
     Output('stored-weeks', 'data'),
     Output('upload-status', 'children'),
     Output('filter-section', 'style')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def upload_file(contents, filename):
    if contents is None:
        return None, None, "", {"display": "none"}
    
    df, weeks, error = parse_uploaded_excel(contents, filename)
    
    if error:
        return None, None, dbc.Alert(error, color="danger"), {"display": "none"}
    
    if df is not None and len(df) > 0:
        success_msg = dbc.Alert([
            html.I(className="bi bi-check-circle-fill me-2"),
            f"✓ {filename} 업로드 완료 ({len(df)}개 레코드)"
        ], color="success")
        
        return df.to_json(date_format='iso', orient='split'), weeks, success_msg, {"display": "block"}
    
    return None, None, dbc.Alert("파일에 데이터가 없습니다.", color="warning"), {"display": "none"}

@app.callback(
    [Output('week-selector', 'options'),
     Output('week-selector', 'value'),
     Output('compare-week-selector', 'options')],
    [Input('stored-weeks', 'data')]
)
def update_week_selectors(weeks):
    if not weeks or len(weeks) == 0:
        return [], None, []
    
    options = [{'label': week, 'value': week} for week in weeks]
    default_week = weeks[0] if weeks else None
    
    return options, default_week, options

@app.callback(
    [Output('column-selector', 'options'),
     Output('column-selector', 'value')],
    [Input('stored-data', 'data')]
)
def update_column_selector(json_data):
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
     Input('compare-week-selector', 'value')]
)
def update_kpi_cards(json_data, selected_week, compare_week):
    if json_data is None or selected_week is None:
        return html.Div()
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Filter by selected week
    if '주차' in df.columns:
        df_filtered = df[df['주차'] == selected_week]
    else:
        df_filtered = df
    
    # Calculate KPIs
    total_sales = df_filtered['매출'].sum() if '매출' in df_filtered.columns else 0
    total_profit = df_filtered['이익'].sum() if '이익' in df_filtered.columns else 0
    avg_roi = df_filtered['이익률'].mean() if '이익률' in df_filtered.columns else 0
    product_count = len(df_filtered)
    
    # Calculate comparison if compare week is selected
    change_text = ""
    if compare_week and compare_week != selected_week and '주차' in df.columns:
        df_compare = df[df['주차'] == compare_week]
        prev_sales = df_compare['매출'].sum() if '매출' in df_compare.columns else 0
        
        if prev_sales > 0:
            change_pct = ((total_sales - prev_sales) / prev_sales) * 100
            change_text = f" ({change_pct:+.1f}% vs {compare_week})"
    
    cards = dbc.Row([
        dbc.Col(create_kpi_card("총 매출", format_currency(total_sales) + change_text, "💰", "success"), md=3),
        dbc.Col(create_kpi_card("총 이익", format_currency(total_profit), "📈", "info"), md=3),
        dbc.Col(create_kpi_card("평균 ROI", format_percentage(avg_roi), "📊", "warning"), md=3),
        dbc.Col(create_kpi_card("상품 수", f"{product_count}개", "🛍️", "primary"), md=3)
    ])
    
    return cards

@app.callback(
    Output('data-table-container', 'children'),
    [Input('stored-data', 'data'),
     Input('week-selector', 'value'),
     Input('compare-week-selector', 'value'),
     Input('column-selector', 'value')]
)
def update_data_table(json_data, selected_week, compare_week, selected_columns):
    if json_data is None or selected_week is None:
        return html.Div("데이터를 업로드하고 주차를 선택하세요.", className="text-muted")
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Filter by weeks
    if '주차' in df.columns:
        if compare_week and compare_week != selected_week:
            # Show both weeks
            df_filtered = df[df['주차'].isin([selected_week, compare_week])]
        else:
            # Show only selected week
            df_filtered = df[df['주차'] == selected_week]
    else:
        df_filtered = df
    
    # Select columns to display
    base_cols = ['상품명', '주차']
    display_cols = base_cols + (selected_columns if selected_columns else [])
    display_cols = [col for col in display_cols if col in df_filtered.columns]
    
    df_display = df_filtered[display_cols].copy()
    
    # Format columns
    for col in df_display.columns:
        if col in ['상품명', '주차', '쇼핑몰', '특이사항', '의견']:
            # Text columns - keep as is
            continue
        elif df_display[col].dtype in ['float64', 'int64']:
            if '이익률' in col or 'ROI' in col:
                # Percentage columns
                df_display[col] = df_display[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A")
            elif '슬롯수변동' in col:
                # 슬롯수변동: 숫자로 표시 (% 아님)
                df_display[col] = df_display[col].apply(lambda x: f"{int(x)}" if pd.notna(x) else "N/A")
            elif '슬롯수' in col:
                # 슬롯수: 정수로 표시
                df_display[col] = df_display[col].apply(lambda x: f"{int(x)}" if pd.notna(x) else "N/A")
            else:
                # 금액 컬럼: 천단위 콤마
                df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "N/A")
    
    table = dash_table.DataTable(
        data=df_display.to_dict('records'),
        columns=[{"name": col, "id": col} for col in df_display.columns],
        page_size=20,
        sort_action="native",
        filter_action="native",
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '10px',
            'fontFamily': 'Arial, sans-serif',
            'minWidth': '100px'
        },
        style_header={
            'backgroundColor': '#007bff',
            'color': 'white',
            'fontWeight': 'bold'
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
    Output('charts-container', 'children'),
    [Input('stored-data', 'data'),
     Input('week-selector', 'value'),
     Input('compare-week-selector', 'value'),
     Input('column-selector', 'value')]
)
def update_charts(json_data, selected_week, compare_week, selected_columns):
    if json_data is None or selected_week is None or not selected_columns:
        return html.Div("컬럼을 선택하면 차트가 표시됩니다.", className="text-muted")
    
    df = pd.read_json(io.StringIO(json_data), orient='split')
    
    # Check if compare mode
    compare_mode = compare_week and compare_week != selected_week and '주차' in df.columns
    
    charts = []
    
    for col in selected_columns:
        if col not in df.columns:
            continue
        
        # Skip text columns for charts
        if col in ['쇼핑몰', '특이사항', '의견']:
            continue
        
        # Check if numeric
        if df[col].dtype not in ['float64', 'int64']:
            continue
        
        if compare_mode:
            # Grouped bar chart for comparison
            df_selected = df[df['주차'] == selected_week].nlargest(20, col) if col in df.columns else df[df['주차'] == selected_week]
            df_compare = df[df['주차'] == compare_week]
            
            # Merge data
            if '상품명' in df.columns:
                product_names = df_selected['상품명'].tolist()
                
                fig = go.Figure()
                
                # Selected week
                fig.add_trace(go.Bar(
                    name=f'{selected_week}',
                    x=product_names,
                    y=df_selected[col].tolist(),
                    marker_color='#3b82f6',
                    hovertemplate='<b>%{x}</b><br>' + col + ': %{y:,.0f}원<extra></extra>' if '이익률' not in col and 'ROI' not in col and '%' not in col else '<b>%{x}</b><br>' + col + ': %{y:.1f}%<extra></extra>'
                ))
                
                # Compare week
                compare_values = []
                for product in product_names:
                    val = df_compare[df_compare['상품명'] == product][col].values
                    compare_values.append(val[0] if len(val) > 0 else 0)
                
                fig.add_trace(go.Bar(
                    name=f'{compare_week}',
                    x=product_names,
                    y=compare_values,
                    marker_color='#f97316',
                    hovertemplate='<b>%{x}</b><br>' + col + ': %{y:,.0f}원<extra></extra>' if '이익률' not in col and 'ROI' not in col and '%' not in col else '<b>%{x}</b><br>' + col + ': %{y:.1f}%<extra></extra>'
                ))
                
                fig.update_layout(
                    title=f"{col} 비교 - 상위 20개 상품 ({selected_week} vs {compare_week})",
                    xaxis_title="상품명",
                    yaxis_title=col,
                    barmode='group',
                    xaxis_tickangle=-45,
                    height=500,
                    hovermode='x unified'
                )
                
                # Format y-axis
                if '이익률' in col or 'ROI' in col or '%' in col:
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
                
                fig.add_trace(go.Bar(
                    x=df_sorted['상품명'],
                    y=df_sorted[col],
                    marker_color='#3b82f6',
                    hovertemplate='<b>%{x}</b><br>' + col + ': %{y:,.0f}원<extra></extra>' if '이익률' not in col and 'ROI' not in col and '%' not in col else '<b>%{x}</b><br>' + col + ': %{y:.1f}%<extra></extra>'
                ))
                
                fig.update_layout(
                    title=f"{col} - 상위 20개 상품",
                    xaxis_title="상품명",
                    yaxis_title=col,
                    xaxis_tickangle=-45,
                    height=400
                )
                
                # Format y-axis
                if '이익률' in col or 'ROI' in col or '%' in col:
                    fig.update_yaxes(ticksuffix='%')
                else:
                    fig.update_yaxes(tickformat=',', ticksuffix='원')
                
                charts.append(dbc.Col([
                    dcc.Graph(figure=fig)
                ], md=6))
            
            # Histogram
            fig_hist = go.Figure()
            
            fig_hist.add_trace(go.Histogram(
                x=df_filtered[col],
                nbinsx=30,
                marker_color='#8b5cf6',
                hovertemplate=col + ': %{x:,.0f}원<br>개수: %{y}<extra></extra>' if '이익률' not in col and 'ROI' not in col and '%' not in col else col + ': %{x:.1f}%<br>개수: %{y}<extra></extra>'
            ))
            
            fig_hist.update_layout(
                title=f"{col} 분포",
                xaxis_title=col,
                yaxis_title="개수",
                height=400
            )
            
            # Format x-axis
            if '이익률' in col or 'ROI' in col or '%' in col:
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
