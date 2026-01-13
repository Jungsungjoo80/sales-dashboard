"""
트래픽 비용 및 이익률 분석 대시보드
Excel 파일 업로드 기능 포함
"""

import dash
from dash import dcc, html, dash_table, Input, Output, State
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import base64
import io
from datetime import datetime

# 앱 초기화
app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server

# 브랜드 분류 함수
def classify_brand(product_name):
    """상품명에서 브랜드 추출"""
    if pd.isna(product_name):
        return '기타'

    product_name_lower = str(product_name).lower()

    if '랩온랩' in product_name_lower or 'wraponwrap' in product_name_lower:
        return '랩온랩'
    elif '네이쳐리브' in product_name_lower or 'natureliv' in product_name_lower:
        return '네이쳐리브'
    elif '닥터하르' in product_name_lower or 'dr.haru' in product_name_lower or 'dr haru' in product_name_lower:
        return '닥터하르'
    else:
        return '기타'

# Excel 파일 파싱 함수
def parse_excel_file(contents, filename):
    """업로드된 Excel 파일을 파싱하여 DataFrame 반환"""
    try:
        # Base64 디코딩
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)

        # Excel 파일 읽기
        excel_file = pd.ExcelFile(io.BytesIO(decoded))

        all_data = []

        # 모든 시트 읽기
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)

            # 컬럼명 정규화
            df.columns = df.columns.str.strip()

            # 필요한 컬럼 확인
            required_columns = ['상품명', '매출', '이익', '이익률']

            # 컬럼 매핑 (다양한 컬럼명 대응)
            column_mapping = {
                '상품명': ['상품명', '제품명', 'Product Name', 'product_name'],
                '매출': ['매출', '매출액', 'Sales', 'sales_amount', '매출금액'],
                '이익': ['이익', '이익금액', 'Profit', 'profit_amount', '순이익'],
                '이익률': ['이익률', 'Profit Rate', 'profit_rate', '이익율', '수익률']
            }

            # 컬럼명 변환
            for target_col, possible_cols in column_mapping.items():
                for col in df.columns:
                    if col in possible_cols:
                        df = df.rename(columns={col: target_col})
                        break

            # 필요한 컬럼이 있는지 확인
            if all(col in df.columns for col in required_columns):
                # 주차 정보 추출
                df['week_name'] = sheet_name

                # 주차 번호 추출 (예: "12월 1주차" -> 1)
                week_number = 0
                if '주차' in sheet_name:
                    try:
                        week_number = int(''.join(filter(str.isdigit, sheet_name.split('주차')[0])))
                    except:
                        week_number = len(all_data) + 1

                df['week_number'] = week_number

                # 브랜드 분류
                df['brand'] = df['상품명'].apply(classify_brand)

                # 숫자형 데이터 정리
                for col in ['매출', '이익', '이익률']:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')

                # 결측치 제거
                df = df.dropna(subset=['상품명', '매출'])

                all_data.append(df[['week_name', 'week_number', 'brand', '상품명', '매출', '이익', '이익률']])

        if not all_data:
            return None

        # 모든 데이터 병합
        final_df = pd.concat(all_data, ignore_index=True)

        # 컬럼명 영문으로 변경 (내부 처리용)
        final_df = final_df.rename(columns={
            '상품명': 'product_name',
            '매출': 'sales_amount',
            '이익': 'profit_amount',
            '이익률': 'profit_rate'
        })

        return final_df

    except Exception as e:
        print(f"파일 파싱 오류: {e}")
        return None

# 앱 레이아웃
app.layout = html.Div([
    # 헤더
    html.Div([
        html.H1('🚀 트래픽 분석 대시보드', 
                style={'textAlign': 'center', 'color': '#2c3e50', 'marginBottom': '10px'}),
        html.P('Excel 파일을 업로드하여 실시간 데이터 분석',
               style={'textAlign': 'center', 'color': '#7f8c8d', 'fontSize': '16px'})
    ], style={'backgroundColor': '#ecf0f1', 'padding': '20px', 'borderRadius': '10px', 'marginBottom': '20px'}),

    # 파일 업로드 영역
    html.Div([
        dcc.Upload(
            id='upload-data',
            children=html.Div([
                html.I(className='fas fa-cloud-upload-alt', style={'fontSize': '48px', 'color': '#3498db'}),
                html.Br(),
                html.Br(),
                '드래그 앤 드롭 또는 ',
                html.A('파일 선택', style={'color': '#3498db', 'fontWeight': 'bold', 'cursor': 'pointer'}),
                html.Br(),
                html.Span('Excel 파일 (.xlsx, .xls)', style={'fontSize': '12px', 'color': '#95a5a6'})
            ]),
            style={
                'width': '100%',
                'height': '120px',
                'lineHeight': '120px',
                'borderWidth': '2px',
                'borderStyle': 'dashed',
                'borderRadius': '10px',
                'borderColor': '#3498db',
                'textAlign': 'center',
                'backgroundColor': '#f8f9fa',
                'cursor': 'pointer',
                'transition': 'all 0.3s'
            },
            multiple=False
        ),
        html.Div(id='upload-status', style={'marginTop': '10px', 'textAlign': 'center'})
    ], style={'marginBottom': '30px'}),

    # 데이터 저장소 (숨김)
    dcc.Store(id='stored-data'),

    # 대시보드 콘텐츠 (초기에는 숨김)
    html.Div(id='dashboard-content', style={'display': 'none'}, children=[
        # 필터 영역
        html.Div([
            html.Div([
                html.Label('📅 주차 선택:', style={'fontWeight': 'bold', 'marginBottom': '5px'}),
                dcc.Dropdown(
                    id='week-filter',
                    placeholder='전체',
                    multi=True,
                    style={'width': '100%'}
                )
            ], style={'flex': '1', 'marginRight': '15px'}),

            html.Div([
                html.Label('🎨 브랜드 선택:', style={'fontWeight': 'bold', 'marginBottom': '5px'}),
                dcc.Dropdown(
                    id='brand-filter',
                    placeholder='전체',
                    multi=True,
                    style={'width': '100%'}
                )
            ], style={'flex': '1', 'marginRight': '15px'}),

            html.Div([
                html.Label('🔍 상품명 검색:', style={'fontWeight': 'bold', 'marginBottom': '5px'}),
                dcc.Input(
                    id='product-search',
                    type='text',
                    placeholder='상품명 입력...',
                    style={'width': '100%', 'padding': '8px', 'borderRadius': '5px', 'border': '1px solid #ddd'}
                )
            ], style={'flex': '1'})
        ], style={'display': 'flex', 'marginBottom': '30px', 'padding': '20px', 
                  'backgroundColor': '#f8f9fa', 'borderRadius': '10px'}),

        # KPI 카드
        html.Div(id='kpi-cards', style={'display': 'flex', 'gap': '20px', 'marginBottom': '30px'}),

        # 차트 영역
        html.Div([
            # 주간 매출 추이
            html.Div([
                dcc.Graph(id='sales-trend-chart')
            ], style={'flex': '2', 'marginRight': '15px'}),

            # 브랜드별 매출
            html.Div([
                dcc.Graph(id='brand-sales-chart')
            ], style={'flex': '1'})
        ], style={'display': 'flex', 'marginBottom': '30px'}),

        # 이익률 분포
        html.Div([
            dcc.Graph(id='profit-rate-chart')
        ], style={'marginBottom': '30px'}),

        # 데이터 테이블
        html.Div([
            html.H3('📊 상세 데이터', style={'color': '#2c3e50', 'marginBottom': '15px'}),
            dash_table.DataTable(
                id='data-table',
                style_table={'overflowX': 'auto'},
                style_cell={
                    'textAlign': 'left',
                    'padding': '10px',
                    'fontSize': '14px'
                },
                style_header={
                    'backgroundColor': '#3498db',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'textAlign': 'center'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': '#f8f9fa'
                    }
                ],
                page_size=20,
                sort_action='native',
                filter_action='native'
            )
        ])
    ])
], style={'maxWidth': '1400px', 'margin': '0 auto', 'padding': '30px', 'fontFamily': 'Arial, sans-serif'})

# 파일 업로드 및 데이터 저장 콜백
@app.callback(
    [Output('stored-data', 'data'),
     Output('upload-status', 'children'),
     Output('dashboard-content', 'style')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_data(contents, filename):
    if contents is None:
        return None, '', {'display': 'none'}

    # 파일 파싱
    df = parse_excel_file(contents, filename)

    if df is None:
        return None, html.Div('❌ 파일을 읽을 수 없습니다. Excel 형식을 확인해주세요.', 
                              style={'color': 'red', 'fontWeight': 'bold'}), {'display': 'none'}

    # 데이터를 JSON으로 변환하여 저장
    data_json = df.to_json(date_format='iso', orient='split')

    success_msg = html.Div([
        html.Span('✅ ', style={'fontSize': '20px'}),
        html.Span(f'파일 업로드 성공: {filename}', style={'fontWeight': 'bold', 'color': '#27ae60'}),
        html.Br(),
        html.Span(f'총 {len(df)}개 레코드 로드됨', style={'fontSize': '12px', 'color': '#7f8c8d'})
    ])

    return data_json, success_msg, {'display': 'block'}

# 필터 옵션 업데이트
@app.callback(
    [Output('week-filter', 'options'),
     Output('brand-filter', 'options')],
    [Input('stored-data', 'data')]
)
def update_filter_options(data_json):
    if data_json is None:
        return [], []

    df = pd.read_json(data_json, orient='split')

    # 주차 옵션
    weeks = sorted(df['week_name'].unique())
    week_options = [{'label': week, 'value': week} for week in weeks]

    # 브랜드 옵션
    brands = sorted(df['brand'].unique())
    brand_options = [{'label': brand, 'value': brand} for brand in brands]

    return week_options, brand_options

# KPI 및 차트 업데이트
@app.callback(
    [Output('kpi-cards', 'children'),
     Output('sales-trend-chart', 'figure'),
     Output('brand-sales-chart', 'figure'),
     Output('profit-rate-chart', 'figure'),
     Output('data-table', 'data'),
     Output('data-table', 'columns')],
    [Input('stored-data', 'data'),
     Input('week-filter', 'value'),
     Input('brand-filter', 'value'),
     Input('product-search', 'value')]
)
def update_dashboard(data_json, selected_weeks, selected_brands, search_text):
    if data_json is None:
        return [], {}, {}, {}, [], []

    df = pd.read_json(data_json, orient='split')

    # 필터 적용
    filtered_df = df.copy()

    if selected_weeks:
        filtered_df = filtered_df[filtered_df['week_name'].isin(selected_weeks)]

    if selected_brands:
        filtered_df = filtered_df[filtered_df['brand'].isin(selected_brands)]

    if search_text:
        filtered_df = filtered_df[filtered_df['product_name'].str.contains(search_text, case=False, na=False)]

    # KPI 계산
    total_sales = filtered_df['sales_amount'].sum()
    total_profit = filtered_df['profit_amount'].sum()
    avg_profit_rate = filtered_df['profit_rate'].mean()
    product_count = filtered_df['product_name'].nunique()

    # KPI 카드 생성
    kpi_cards = [
        html.Div([
            html.H4('총 매출', style={'color': '#7f8c8d', 'marginBottom': '10px'}),
            html.H2(f'₩{total_sales:,.1f}M', style={'color': '#3498db', 'margin': '0'})
        ], style={'flex': '1', 'padding': '20px', 'backgroundColor': '#ecf0f1', 
                  'borderRadius': '10px', 'textAlign': 'center'}),

        html.Div([
            html.H4('총 이익', style={'color': '#7f8c8d', 'marginBottom': '10px'}),
            html.H2(f'₩{total_profit:,.1f}M', style={'color': '#2ecc71', 'margin': '0'})
        ], style={'flex': '1', 'padding': '20px', 'backgroundColor': '#ecf0f1', 
                  'borderRadius': '10px', 'textAlign': 'center'}),

        html.Div([
            html.H4('평균 이익률', style={'color': '#7f8c8d', 'marginBottom': '10px'}),
            html.H2(f'{avg_profit_rate:.1f}%', style={'color': '#e74c3c', 'margin': '0'})
        ], style={'flex': '1', 'padding': '20px', 'backgroundColor': '#ecf0f1', 
                  'borderRadius': '10px', 'textAlign': 'center'}),

        html.Div([
            html.H4('상품 수', style={'color': '#7f8c8d', 'marginBottom': '10px'}),
            html.H2(f'{product_count}개', style={'color': '#9b59b6', 'margin': '0'})
        ], style={'flex': '1', 'padding': '20px', 'backgroundColor': '#ecf0f1', 
                  'borderRadius': '10px', 'textAlign': 'center'})
    ]

    # 차트 1: 주간 매출 추이
    weekly_sales = filtered_df.groupby('week_name')['sales_amount'].sum().reset_index()
    weekly_sales = weekly_sales.sort_values('week_name')

    sales_trend_fig = go.Figure()
    sales_trend_fig.add_trace(go.Scatter(
        x=weekly_sales['week_name'],
        y=weekly_sales['sales_amount'],
        mode='lines+markers',
        name='매출',
        line=dict(color='#3498db', width=3),
        marker=dict(size=10)
    ))
    sales_trend_fig.update_layout(
        title='주간 매출 추이',
        xaxis_title='주차',
        yaxis_title='매출 (M원)',
        hovermode='x unified',
        plot_bgcolor='white'
    )

    # 차트 2: 브랜드별 매출
    brand_sales = filtered_df.groupby('brand')['sales_amount'].sum().reset_index()
    brand_sales = brand_sales.sort_values('sales_amount', ascending=True)

    brand_sales_fig = go.Figure()
    brand_sales_fig.add_trace(go.Bar(
        x=brand_sales['sales_amount'],
        y=brand_sales['brand'],
        orientation='h',
        marker=dict(color=['#e74c3c', '#3498db', '#2ecc71', '#f39c12'])
    ))
    brand_sales_fig.update_layout(
        title='브랜드별 매출',
        xaxis_title='매출 (M원)',
        yaxis_title='브랜드',
        plot_bgcolor='white'
    )

    # 차트 3: 이익률 분포
    profit_rate_fig = go.Figure()
    profit_rate_fig.add_trace(go.Histogram(
        x=filtered_df['profit_rate'],
        nbinsx=20,
        marker=dict(color='#9b59b6')
    ))
    profit_rate_fig.update_layout(
        title='이익률 분포',
        xaxis_title='이익률 (%)',
        yaxis_title='상품 수',
        plot_bgcolor='white'
    )

    # 데이터 테이블
    table_df = filtered_df[['brand', 'product_name', 'week_name', 'sales_amount', 
                             'profit_amount', 'profit_rate']].copy()
    table_df = table_df.sort_values('sales_amount', ascending=False)

    # 컬럼명 한글로 변경
    table_df = table_df.rename(columns={
        'brand': '브랜드',
        'product_name': '상품명',
        'week_name': '주차',
        'sales_amount': '매출 (M원)',
        'profit_amount': '이익 (M원)',
        'profit_rate': '이익률 (%)'
    })

    # 숫자 포맷팅
    table_df['매출 (M원)'] = table_df['매출 (M원)'].round(2)
    table_df['이익 (M원)'] = table_df['이익 (M원)'].round(2)
    table_df['이익률 (%)'] = table_df['이익률 (%)'].round(1)

    columns = [{'name': col, 'id': col} for col in table_df.columns]
    data = table_df.to_dict('records')

    return kpi_cards, sales_trend_fig, brand_sales_fig, profit_rate_fig, data, columns

if __name__ == '__main__':
    app.run_server(debug=False, host='0.0.0.0', port=8050)
