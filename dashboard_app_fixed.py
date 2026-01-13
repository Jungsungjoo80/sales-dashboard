"""
Tableau 스타일 인터랙티브 대시보드 (Standalone Version)
트래픽 비용 및 이익률 데이터 시각화

실행 방법:
1. 로컬 실행:
   python3 dashboard_app.py

2. 브라우저 접속:
   http://localhost:8050

파일 구조:
dashboard_app.py (이 파일)
traffic_data.db (데이터베이스)

작성자: GenSpark AI
날짜: 2026-01-13
"""

import dash
from dash import dcc, html, dash_table, Input, Output, State, callback
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
import sqlite3
from datetime import datetime
import os
import sys

# ==================== 데이터 처리 클래스 ====================
class DataProcessor:
    """트래픽 데이터 처리 및 DB 관리 클래스"""

    def __init__(self, db_path: str = 'traffic_data.db'):
        """초기화"""
        self.db_path = db_path

        # 데이터베이스 파일 존재 확인
        if not os.path.exists(self.db_path):
            print(f"⚠️  경고: {self.db_path} 파일을 찾을 수 없습니다.")
            print(f"현재 디렉토리: {os.getcwd()}")
            print(f"디렉토리 파일 목록: {os.listdir('.')}")

    def get_connection(self):
        """데이터베이스 연결 반환"""
        return sqlite3.connect(self.db_path)

    def get_all_data(self) -> pd.DataFrame:
        """모든 데이터 조회"""
        conn = self.get_connection()
        query = """
        SELECT * FROM traffic_data
        ORDER BY week_number, brand, product_name
        """
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df

    def get_weeks(self) -> list:
        """주차 목록 조회"""
        conn = self.get_connection()
        query = "SELECT DISTINCT week_name FROM traffic_data ORDER BY week_number"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df['week_name'].tolist()

    def get_brands(self) -> list:
        """브랜드 목록 조회"""
        conn = self.get_connection()
        query = "SELECT DISTINCT brand FROM traffic_data ORDER BY brand"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df['brand'].tolist()

    def get_products(self, brand: str = None) -> list:
        """상품 목록 조회"""
        conn = self.get_connection()
        if brand:
            query = f"SELECT DISTINCT product_name FROM traffic_data WHERE brand = '{brand}' ORDER BY product_name"
        else:
            query = "SELECT DISTINCT product_name FROM traffic_data ORDER BY product_name"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df['product_name'].tolist()

    def get_summary_by_week(self, week_name: str = None) -> pd.DataFrame:
        """주차별 요약 데이터"""
        conn = self.get_connection()
        if week_name:
            query = f"""
            SELECT 
                week_name,
                brand,
                SUM(sales_amount) as total_sales,
                SUM(profit_amount) as total_profit,
                AVG(profit_rate) as avg_profit_rate,
                COUNT(*) as product_count
            FROM traffic_data
            WHERE week_name = '{week_name}'
            GROUP BY week_name, brand
            ORDER BY total_sales DESC
            """
        else:
            query = """
            SELECT 
                week_name,
                brand,
                SUM(sales_amount) as total_sales,
                SUM(profit_amount) as total_profit,
                AVG(profit_rate) as avg_profit_rate,
                COUNT(*) as product_count
            FROM traffic_data
            GROUP BY week_name, brand
            ORDER BY week_name, total_sales DESC
            """
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df

# ==================== Dash 앱 초기화 ====================
app = dash.Dash(__name__)
app.title = "트래픽 분석 대시보드"

# 데이터 프로세서 초기화
processor = DataProcessor(db_path='traffic_data.db')

# 전역 스타일 설정
colors = {
    'background': '#f8f9fa',
    'card_bg': '#ffffff',
    'text': '#212529',
    'primary': '#0066cc',
    'success': '#28a745',
    'warning': '#ffc107',
    'danger': '#dc3545',
    'info': '#17a2b8'
}

# ==================== 레이아웃 ====================
app.layout = html.Div([
    # 헤더
    html.Div([
        html.H1("🎯 트래픽 분석 대시보드", 
                style={'color': colors['primary'], 'textAlign': 'center', 'margin': '20px'}),
        html.P("Tableau 스타일 인터랙티브 데이터 시각화 시스템",
               style={'textAlign': 'center', 'color': colors['text'], 'marginBottom': '30px'}),
    ], style={'backgroundColor': colors['card_bg'], 'padding': '20px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'}),

    # 필터 패널
    html.Div([
        html.Div([
            html.Label("📅 주차 선택:", style={'fontWeight': 'bold', 'marginBottom': '5px'}),
            dcc.Dropdown(
                id='week-dropdown',
                options=[{'label': week, 'value': week} for week in processor.get_weeks()],
                value=processor.get_weeks()[-1] if processor.get_weeks() else None,
                placeholder="주차를 선택하세요",
                style={'width': '100%'}
            ),
        ], style={'width': '30%', 'display': 'inline-block', 'padding': '10px'}),

        html.Div([
            html.Label("🏷️ 브랜드 선택:", style={'fontWeight': 'bold', 'marginBottom': '5px'}),
            dcc.Dropdown(
                id='brand-dropdown',
                options=[{'label': '전체', 'value': 'ALL'}] + 
                        [{'label': brand, 'value': brand} for brand in processor.get_brands()],
                value='ALL',
                placeholder="브랜드를 선택하세요",
                style={'width': '100%'}
            ),
        ], style={'width': '30%', 'display': 'inline-block', 'padding': '10px'}),

        html.Div([
            html.Label("📦 상품 선택:", style={'fontWeight': 'bold', 'marginBottom': '5px'}),
            dcc.Dropdown(
                id='product-dropdown',
                options=[{'label': '전체', 'value': 'ALL'}],
                value='ALL',
                placeholder="상품을 선택하세요",
                style={'width': '100%'}
            ),
        ], style={'width': '30%', 'display': 'inline-block', 'padding': '10px'}),
    ], style={'backgroundColor': colors['card_bg'], 'padding': '10px', 'marginTop': '20px', 
              'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'}),

    # KPI 카드
    html.Div(id='kpi-cards', style={'marginTop': '20px'}),

    # 차트 영역
    html.Div([
        # 매출 추이 차트
        html.Div([
            dcc.Graph(id='sales-trend-chart')
        ], style={'width': '100%', 'display': 'inline-block', 'padding': '10px'}),

        # 브랜드별 매출 비교
        html.Div([
            dcc.Graph(id='brand-comparison-chart')
        ], style={'width': '50%', 'display': 'inline-block', 'padding': '10px'}),

        # 이익률 분포
        html.Div([
            dcc.Graph(id='profit-distribution-chart')
        ], style={'width': '50%', 'display': 'inline-block', 'padding': '10px'}),
    ], style={'marginTop': '20px'}),

    # 데이터 테이블
    html.Div([
        html.H3("📊 상세 데이터", style={'color': colors['primary'], 'marginBottom': '15px'}),
        html.Div(id='data-table')
    ], style={'backgroundColor': colors['card_bg'], 'padding': '20px', 'marginTop': '20px',
              'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'}),

], style={'backgroundColor': colors['background'], 'padding': '20px', 'minHeight': '100vh'})

# ==================== 콜백 함수 ====================

@app.callback(
    Output('product-dropdown', 'options'),
    Input('brand-dropdown', 'value')
)
def update_product_dropdown(selected_brand):
    """브랜드 선택에 따라 상품 목록 업데이트"""
    if selected_brand == 'ALL' or not selected_brand:
        products = processor.get_products()
    else:
        products = processor.get_products(brand=selected_brand)

    return [{'label': '전체', 'value': 'ALL'}] + [{'label': p, 'value': p} for p in products]

@app.callback(
    [Output('kpi-cards', 'children'),
     Output('sales-trend-chart', 'figure'),
     Output('brand-comparison-chart', 'figure'),
     Output('profit-distribution-chart', 'figure'),
     Output('data-table', 'children')],
    [Input('week-dropdown', 'value'),
     Input('brand-dropdown', 'value'),
     Input('product-dropdown', 'value')]
)
def update_dashboard(selected_week, selected_brand, selected_product):
    """대시보드 업데이트"""

    # 데이터 로드
    df = processor.get_all_data()

    if df.empty:
        empty_fig = go.Figure()
        empty_fig.add_annotation(text="데이터가 없습니다", showarrow=False, font=dict(size=20))
        return html.Div("데이터가 없습니다"), empty_fig, empty_fig, empty_fig, html.Div("데이터가 없습니다")

    # 필터 적용
    filtered_df = df.copy()
    if selected_week:
        filtered_df = filtered_df[filtered_df['week_name'] == selected_week]
    if selected_brand != 'ALL':
        filtered_df = filtered_df[filtered_df['brand'] == selected_brand]
    if selected_product != 'ALL':
        filtered_df = filtered_df[filtered_df['product_name'] == selected_product]

    # KPI 계산
    total_sales = filtered_df['sales_amount'].sum()
    total_profit = filtered_df['profit_amount'].sum()
    avg_profit_rate = filtered_df['profit_rate'].mean()
    product_count = filtered_df['product_name'].nunique()

    # KPI 카드
    kpi_cards = html.Div([
        html.Div([
            html.H4("💰 총 매출", style={'color': colors['text'], 'marginBottom': '10px'}),
            html.H2(f"{total_sales:,.0f}원", style={'color': colors['primary'], 'margin': '0'}),
        ], style={'width': '23%', 'display': 'inline-block', 'padding': '20px', 
                  'backgroundColor': colors['card_bg'], 'margin': '1%',
                  'boxShadow': '0 2px 4px rgba(0,0,0,0.1)', 'borderRadius': '5px'}),

        html.Div([
            html.H4("📈 총 이익", style={'color': colors['text'], 'marginBottom': '10px'}),
            html.H2(f"{total_profit:,.0f}원", style={'color': colors['success'], 'margin': '0'}),
        ], style={'width': '23%', 'display': 'inline-block', 'padding': '20px',
                  'backgroundColor': colors['card_bg'], 'margin': '1%',
                  'boxShadow': '0 2px 4px rgba(0,0,0,0.1)', 'borderRadius': '5px'}),

        html.Div([
            html.H4("📊 평균 이익률", style={'color': colors['text'], 'marginBottom': '10px'}),
            html.H2(f"{avg_profit_rate:.1f}%", style={'color': colors['warning'], 'margin': '0'}),
        ], style={'width': '23%', 'display': 'inline-block', 'padding': '20px',
                  'backgroundColor': colors['card_bg'], 'margin': '1%',
                  'boxShadow': '0 2px 4px rgba(0,0,0,0.1)', 'borderRadius': '5px'}),

        html.Div([
            html.H4("🛍️ 상품 수", style={'color': colors['text'], 'marginBottom': '10px'}),
            html.H2(f"{product_count}개", style={'color': colors['info'], 'margin': '0'}),
        ], style={'width': '23%', 'display': 'inline-block', 'padding': '20px',
                  'backgroundColor': colors['card_bg'], 'margin': '1%',
                  'boxShadow': '0 2px 4px rgba(0,0,0,0.1)', 'borderRadius': '5px'}),
    ])

    # 1. 매출 추이 차트
    sales_trend_df = df.groupby('week_name')['sales_amount'].sum().reset_index()
    sales_trend_fig = go.Figure()
    sales_trend_fig.add_trace(go.Scatter(
        x=sales_trend_df['week_name'],
        y=sales_trend_df['sales_amount'],
        mode='lines+markers',
        name='매출액',
        line=dict(color=colors['primary'], width=3),
        marker=dict(size=10)
    ))
    sales_trend_fig.update_layout(
        title='📈 주차별 매출 추이',
        xaxis_title='주차',
        yaxis_title='매출액 (원)',
        hovermode='x unified',
        plot_bgcolor='white',
        height=400
    )

    # 2. 브랜드별 매출 비교
    brand_sales_df = filtered_df.groupby('brand')['sales_amount'].sum().reset_index()
    brand_sales_df = brand_sales_df.sort_values('sales_amount', ascending=True)
    brand_comparison_fig = go.Figure()
    brand_comparison_fig.add_trace(go.Bar(
        y=brand_sales_df['brand'],
        x=brand_sales_df['sales_amount'],
        orientation='h',
        marker=dict(color=colors['primary'])
    ))
    brand_comparison_fig.update_layout(
        title='🏷️ 브랜드별 매출 비교',
        xaxis_title='매출액 (원)',
        yaxis_title='브랜드',
        height=400,
        plot_bgcolor='white'
    )

    # 3. 이익률 분포
    profit_dist_fig = go.Figure()
    profit_dist_fig.add_trace(go.Histogram(
        x=filtered_df['profit_rate'],
        nbinsx=30,
        marker=dict(color=colors['success']),
        name='이익률 분포'
    ))
    profit_dist_fig.update_layout(
        title='📊 이익률 분포',
        xaxis_title='이익률 (%)',
        yaxis_title='상품 수',
        height=400,
        plot_bgcolor='white'
    )

    # 4. 데이터 테이블
    table_df = filtered_df[['week_name', 'brand', 'product_name', 'sales_amount', 'profit_amount', 'profit_rate']].copy()
    table_df = table_df.sort_values('sales_amount', ascending=False).head(50)

    data_table = dash_table.DataTable(
        data=table_df.to_dict('records'),
        columns=[
            {'name': '주차', 'id': 'week_name'},
            {'name': '브랜드', 'id': 'brand'},
            {'name': '상품명', 'id': 'product_name'},
            {'name': '매출액', 'id': 'sales_amount', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
            {'name': '이익액', 'id': 'profit_amount', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
            {'name': '이익률(%)', 'id': 'profit_rate', 'type': 'numeric', 'format': {'specifier': '.1f'}},
        ],
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '10px',
            'fontFamily': 'Arial, sans-serif'
        },
        style_header={
            'backgroundColor': colors['primary'],
            'color': 'white',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8f9fa'
            }
        ],
        page_size=20
    )

    return kpi_cards, sales_trend_fig, brand_comparison_fig, profit_dist_fig, data_table

# ==================== 서버 실행 ====================
if __name__ == '__main__':
    print("=" * 60)
    print("🎯 트래픽 분석 대시보드 시작")
    print("=" * 60)
    print(f"📂 현재 디렉토리: {os.getcwd()}")
    print(f"📊 데이터베이스: {processor.db_path}")
    print(f"🌐 브라우저에서 접속: http://localhost:8050")
    print("=" * 60)
    print("⏹️  종료하려면 Ctrl+C를 누르세요")
    print("=" * 60)

    app.run_server(debug=True, host='0.0.0.0', port=8050)
