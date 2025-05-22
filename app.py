import dash
from dash import dcc, html, Input, Output, State
import pandas as pd
import requests
import plotly.express as px
import dash_bootstrap_components as dbc
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.SANDSTONE])
server = app.server
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>台股財報指標查詢系統</title>
        {%favicon%}
        {%css%}
        <link href="https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic&display=swap" rel="stylesheet">
        <style>
            body, html, .card, .card-body, .form-label, label, h1, h2, h3, h4, h5, h6, div, span, p, input, select {
                font-family: 'Zen Maru Gothic', sans-serif !important;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''


FINMIND_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNS0wNS0xNyAyMTozMjowOSIsInVzZXJfaWQiOiJzdGVsbGE0NyIsImlwIjoiMzYuMjMwLjUzLjI0OSJ9.dSiERmcl7OJS86hvIQskcOVOoW6_Crh2Pv6CFhoyf5M"


app.title = "財報獲利指標查詢"

def generate_all_quarters(start_year=1990, end_year=2025): 
    quarters = []
    for year in range(start_year, end_year + 1):
        for q in range(1, 5):
            quarters.append(f"{year}Q{q}")
    return quarters

def generate_quarter_range(start, end):
    start_year, start_q = int(start[:4]), int(start[-1])
    end_year, end_q = int(end[:4]), int(end[-1])
    result = []
    for y in range(start_year, end_year + 1):
        for q in range(1, 5):
            if (y == start_year and q < start_q) or (y == end_year and q > end_q):
                continue
            result.append(f"{y}Q{q}")
    return result


# 先用函式產生完整季度列表
all_quarters = generate_all_quarters(1990, 2025)

# 逆序排序，最新季度排在前面
all_quarters_desc = sorted(all_quarters, reverse=True)

app.layout = dbc.Container([

    # 1️⃣ 頁首區塊
    html.Div([
        html.H2("📊 台股財報指標查詢系統", className="text-center fw-bold mb-0"),
    ], style={
        "backgroundColor": "#264653",
        "color": "white",
        "padding": "20px",
        "borderRadius": "10px",
        "marginBottom": "30px"
    }),

    # 2️⃣ 股票代碼輸入
    dbc.Row([
        dbc.Col([
            html.Label("輸入股票代碼（用逗號分隔，如：2330,2317）"),
            dcc.Input(id="stock-input", type="text", className="form-control"),
        ], width=12)
    ], className="mb-3"),

    # 3️⃣ 季度選擇
    dbc.Row([
        dbc.Col([
            html.Label("起始季度", className="me-2"),
            dcc.Dropdown(
                id="start-quarter",
                options=[{"label": q, "value": q} for q in all_quarters_desc],
                clearable=False,
                className="form-select",
                style={"width": "150px"}
            ),
        ],width="auto", className="d-flex align-items-center"),

        dbc.Col([
            html.Label("結束季度", className="me-2"),
            dcc.Dropdown(
                id="end-quarter",
                options=[{"label": q, "value": q} for q in all_quarters_desc],
                clearable=False,
                className="form-select",
                style={"width": "150px"}
            ),
        ], width="auto",className="d-flex align-items-center"),
    ], className="justify-content-center mb-3"),


    # 4️⃣ 查詢與下載按鈕
    dbc.Row([
        dbc.Col(
            dbc.Button("查詢", id="submit-btn", color="primary", className="me-3", style={"width": "150px"}),
            width="auto"
        ),
        dbc.Col(
            dbc.Button("下載 Excel", id="download-btn", color="success", style={"width": "150px"}),
            width="auto"
        )
    ], justify="center", className="mb-4"),

    # 5️⃣ 指標勾選區
    html.Div([
        html.Small("請勾選要產生趨勢圖的指標", className="text-muted d-block text-center mb-3"),

        dbc.Row([
            # ➤ 獲利能力指標
            dbc.Col([
                html.H4("獲利能力指標", className="text-info fw-bold text-center"),
                dcc.Checklist(
                    id="selected-profitability-indicators",
                    options=[{"label": i, "value": i} for i in [
                        "每股盈餘(EPS)", "營業毛利率(GrossProfitMargin)", "營業利益率(OperatingMargin)",
                        "稅前淨利率(PreTaxProfitMargin)", "稅後淨利率(NetProfitMargin)",
                        "資產報酬率(ROA)", "股東權益報酬率(ROE)"
                    ]],
                    value=["每股盈餘(EPS)"],
                    labelStyle={"display": "flex", "alignItems": "center", "gap": "20px", "marginBottom": "10px"}
                ),
            ], width=3),

            # ➤ 償債能力指標
            dbc.Col([
                html.H4("償債能力指標", className="text-info fw-bold text-center"),
                dcc.Checklist(
                    id="selected-solvency-indicators",
                    options=[{"label": i, "value": i} for i in [
                        "現金比率(CashRatio)", "流動比率(CurrentRatio)",
                        "利息保障倍數(InterestCoverageRatio)", "現金流量比(OperatingCashFlowRatio)"
                    ]],
                    value=[],
                    labelStyle={"display": "flex", "alignItems": "center", "gap": "20px", "marginBottom": "10px"}
                ),
            ], width=3),

            # ➤ 成長能力指標
            dbc.Col([
                html.H4("獲利年成長率", className="text-info fw-bold text-center"),
                dcc.Checklist(
                    id="selected-growth-indicators",
                    options=[{"label": i, "value": i} for i in [
                        "營收年成長率(Revenue YoY)", "毛利年成長率(GrossProfit YoY)",
                        "營業利益年成長率(OperatingIncome YoY)", "稅前淨利年成長率(PreTaxIncome YoY)",
                        "稅後淨利年成長率(IncomeAfterTaxes YoY)", "每股盈餘年成長率(EPS YoY)"
                    ]],
                    value=[],
                    labelStyle={"display": "flex", "alignItems": "center", "gap": "20px", "marginBottom": "10px"}
                ),
            ], width=3),

            # ➤ 經營能力指標
            dbc.Col([
                html.H4("經營能力指標", className="text-info fw-bold text-center"),
                dcc.Checklist(
                    id="selected-efficiency-indicators",
                    options=[{"label": i, "value": i} for i in [
                        "營業成本率(CostMargin)", "營業費用率(ExpenseMargin)",
                        "存貨週轉率(InventoryTurnover)", "平均售貨日數(DaysSalesOutstanding)",
                        "總資產週轉率(TotalAssetTurnover)"
                    ]],
                    value=[],
                    labelStyle={"display": "flex", "alignItems": "center", "gap": "20px", "marginBottom": "10px"}
                ),
            ], width=3),
        ], style={"minHeight": "400px"}),  # 四欄等高

    ], className="mb-4"),

    # 6️⃣ 下載元件（只放一次）
    dcc.Download(id="download-excel"),

    # 7️⃣ 結果與圖表區
    dbc.Row([
        dbc.Col(html.Div(id="result-table"), width=12, className="mt-4"),
        dbc.Col(dcc.Graph(id="trend-graph"), width=12),
        dbc.Col(html.Div(id="indicator-graphs"), width=12, className="mt-4"),
    ])

], fluid=True)




def fetch_finmind_data(dataset, stock_id, token):
    url = "https://api.finmindtrade.com/api/v4/data"
    params = {
        "dataset": dataset,
        "data_id": stock_id,
        "token": token,
        "start_date": "1990-01-01",
        "end_date": "2025-12-31"
    }
    r = requests.get(url, params=params)
    if r.status_code != 200:
        return pd.DataFrame()
    res = r.json()
    if "data" not in res or not res["data"]:
        return pd.DataFrame()
    return pd.DataFrame(res["data"])

def convert_quarter_to_date(q):
    """
    將 2023Q1 → 2023-03-31 格式（財報發佈的代表日）
    """
    year = int(q[:4])
    quarter = int(q[-1])
    if quarter == 1:
        return f"{year}-03-31"
    elif quarter == 2:
        return f"{year}-06-30"
    elif quarter == 3:
        return f"{year}-09-30"
    elif quarter == 4:
        return f"{year}-12-31"



def get_financial_indicators(stock_id, quarters, token):
    income_df = fetch_finmind_data("TaiwanStockFinancialStatements", stock_id, token)
    balance_df = fetch_finmind_data("TaiwanStockBalanceSheet", stock_id, token)
    cashflow_df = fetch_finmind_data("TaiwanStockCashFlowsStatement", stock_id, token)

    if income_df.empty and balance_df.empty and cashflow_df.empty:
        return pd.DataFrame()

    income_items = ["EPS", "GrossProfit", "OperatingIncome", "PreTaxIncome", "IncomeAfterTaxes", "Revenue","CostOfGoodsSold", "OperatingExpenses"]
    balance_items = ["TotalAssets", "EquityAttributableToOwnersOfParent", "CurrentAssets", "CurrentLiabilities", "CashAndCashEquivalents", "Inventories"]
    cashflow_items = ["NetCashInflowFromOperatingActivities","NetIncomeBeforeTax","PayTheInterest"]

    income_df = income_df[income_df["type"].isin(income_items)]
    balance_df = balance_df[balance_df["type"].isin(balance_items)]
    cashflow_df = cashflow_df[cashflow_df["type"].isin(cashflow_items)]

    income_pivot = income_df.pivot_table(index="date", columns="type", values="value").reset_index()
    balance_pivot = balance_df.pivot_table(index="date", columns="type", values="value").reset_index()
    cashflow_pivot = cashflow_df.pivot_table(index="date", columns="type", values="value").reset_index()

    df = pd.merge(income_pivot, balance_pivot, on="date", how="outer")
    df = pd.merge(df, cashflow_pivot, on="date", how="outer")

    # 補齊缺的欄位
    required_columns = income_items + balance_items + cashflow_items
    for col in required_columns:
        if col not in df.columns:
            df[col] = pd.NA

    # 計算原本的獲利指標
    df["營業毛利率(GrossProfitMargin)"] = (df["GrossProfit"] / df["Revenue"]) * 100
    df["營業利益率(OperatingMargin)"] = (df["OperatingIncome"] / df["Revenue"]) * 100
    df["稅前淨利率(PreTaxProfitMargin)"] = (df["PreTaxIncome"] / df["Revenue"]) * 100
    df["稅後淨利率(NetProfitMargin)"] = (df["IncomeAfterTaxes"] / df["Revenue"]) * 100
    df["資產報酬率(ROA)"] = (df["IncomeAfterTaxes"] / df["TotalAssets"]) * 100
    df["股東權益報酬率(ROE)"] = (df["IncomeAfterTaxes"] / df["EquityAttributableToOwnersOfParent"]) * 100
    df.rename(columns={"EPS": "每股盈餘(EPS)"}, inplace=True)

    # 新增償債能力指標
    df["現金比率(CashRatio)"] = (df["CashAndCashEquivalents"] / df["CurrentLiabilities"]) * 100
    df["流動比率(CurrentRatio)"] = (df["CurrentAssets"] / df["CurrentLiabilities"]) * 100
    df["利息保障倍數(InterestCoverageRatio)"] = (df["NetIncomeBeforeTax"] + df["PayTheInterest"]) / df["PayTheInterest"]
    df["現金流量比(OperatingCashFlowRatio)"] = (df["NetCashInflowFromOperatingActivities"] / df["CurrentLiabilities"]) * 100

     # 加入季度欄位
    df["季度"] = pd.to_datetime(df["date"]).apply(lambda x: f"{x.year}Q{((x.month - 1) // 3 + 1)}")
    df["股票代碼"] = stock_id
    # ==== 計算年成長率（YoY）指標 ====
    df = df.sort_values("季度")  # 確保季度順序正確
    
    # 計算經營能力指標
    df = df.sort_values("date")  # 先照時間排序，確保 shift 準確

    # 上一季的存貨與總資產
    df["上一季存貨"] = df["Inventories"].shift(1)
    df["上一季總資產"] = df["TotalAssets"].shift(1)

    df["營業成本率(CostMargin)"] = (df["CostOfGoodsSold"] / df["Revenue"]) * 100
    df["營業費用率(ExpenseMargin)"] = (df["OperatingExpenses"] / df["Revenue"]) * 100

    df["存貨週轉率(InventoryTurnover)"] = (df["CostOfGoodsSold"] / ((df["上一季存貨"] + df["Inventories"]) / 2)) * 4
    df["平均售貨日數(DaysSalesOutstanding)"] = 365 / df["存貨週轉率(InventoryTurnover)"]
    df["總資產週轉率(TotalAssetTurnover)"] = (df["Revenue"] / ((df["上一季總資產"] + df["TotalAssets"]) / 2)) * 4


    def compute_yoy(series):
        return (series - series.shift(4)) / series.shift(4) * 100

    df["營收年成長率(Revenue YoY)"] = compute_yoy(df["Revenue"])
    df["毛利年成長率(GrossProfit YoY)"] = compute_yoy(df["GrossProfit"])
    df["營業利益年成長率(OperatingIncome YoY)"] = compute_yoy(df["OperatingIncome"])
    df["稅前淨利年成長率(PreTaxIncome YoY)"] = compute_yoy(df["PreTaxIncome"])
    df["稅後淨利年成長率(IncomeAfterTaxes YoY)"] = compute_yoy(df["IncomeAfterTaxes"])
    df["每股盈餘年成長率(EPS YoY)"] = compute_yoy(df["每股盈餘(EPS)"])

    df = df.drop_duplicates()

    # === 新增：年成長率計算（過去四季 YoY） ===
    df_yoy = df[["季度", "Revenue", "GrossProfit", "OperatingIncome", "PreTaxIncome", "IncomeAfterTaxes", "每股盈餘(EPS)"]].copy()
    df_yoy.set_index("季度", inplace=True)

    for col in ["Revenue", "GrossProfit", "OperatingIncome", "PreTaxIncome", "IncomeAfterTaxes", "每股盈餘(EPS)"]:
        yoy_col = f"{col}年成長率"
        df[yoy_col] = df[col].pct_change(periods=4) * 100  # 四季前為基準

    # 過濾季度
    df = df[df["季度"].isin(set(quarters))]

    columns_order = [
    "股票代碼", "date", "季度",

    # 獲利能力指標
    "每股盈餘(EPS)",
    "營業毛利率(GrossProfitMargin)",
    "營業利益率(OperatingMargin)",
    "稅前淨利率(PreTaxProfitMargin)",
    "稅後淨利率(NetProfitMargin)",
    "資產報酬率(ROA)",
    "股東權益報酬率(ROE)",

    # 償債能力指標
    "現金比率(CashRatio)",
    "流動比率(CurrentRatio)",
    "利息保障倍數(InterestCoverageRatio)",
    "現金流量比(OperatingCashFlowRatio)",

    # 年成長率指標
    "營收年成長率(Revenue YoY)",
    "毛利年成長率(GrossProfit YoY)",
    "營業利益年成長率(OperatingIncome YoY)",
    "稅前淨利年成長率(PreTaxIncome YoY)",
    "稅後淨利年成長率(IncomeAfterTaxes YoY)",
    "每股盈餘年成長率(EPS YoY)",
        
    # 經營能力指標
    "營業成本率(CostMargin)",
    "營業費用率(ExpenseMargin)", 
    "存貨週轉率(InventoryTurnover)", 
    "平均售貨日數(DaysSalesOutstanding)", 
    "總資產週轉率(TotalAssetTurnover)" 
    ]


    return df[columns_order]




@app.callback(
    Output("indicator-graphs", "children"),
    Input("submit-btn", "n_clicks"),
    State("stock-input", "value"),
    State("start-quarter", "value"),
    State("end-quarter", "value"),
    State("selected-profitability-indicators", "value"),
    State("selected-solvency-indicators", "value"),
    State("selected-growth-indicators", "value"),
    State("selected-efficiency-indicators", "value"),
    prevent_initial_call=True
)
def update_graphs(n_clicks, stock_input, start_quarter, end_quarter,
                  profitability_indicators, solvency_indicators, growth_indicators, efficiency_indicators):

    if n_clicks == 0:
        return []
    if not stock_input or not start_quarter or not end_quarter:
        return dbc.Alert("請輸入股票代碼與選擇季度區間", color="danger", className="mt-3")

    selected_indicators = profitability_indicators + solvency_indicators + growth_indicators + efficiency_indicators
    stock_ids = [s.strip() for s in stock_input.split(",") if s.strip()]
    quarters = generate_quarter_range(start_quarter, end_quarter)
    all_dfs = []
    no_data_stocks = []

    for sid in stock_ids:
        try:
            df = get_financial_indicators(sid, quarters, FINMIND_TOKEN)
            if df.empty:
                no_data_stocks.append(sid)
                continue
            indicator_cols = ["每股盈餘(EPS)", "營業毛利率(GrossProfitMargin)", "營業利益率(OperatingMargin)",
                              "稅前淨利率(PreTaxProfitMargin)", "稅後淨利率(NetProfitMargin)", "資產報酬率(ROA)", "股東權益報酬率(ROE)"]
            if df[indicator_cols].dropna(how='all').empty:
                no_data_stocks.append(sid)
                continue
            all_dfs.append(df)
        except Exception as e:
            no_data_stocks.append(sid)
            continue

    if not all_dfs and not no_data_stocks:
        return dbc.Alert("查無任何資料", color="warning", className="mt-3")

    content = []

    if no_data_stocks:
        content.append(
            dbc.Alert(
                [html.H5("以下股票查無資料：", className="mb-2"),
                 html.Ul([html.Li(sid) for sid in no_data_stocks])],
                color="danger",
                className="mt-3"
            )
        )

    # 篩選有效 DataFrame
    filtered_dfs = [df for df in all_dfs if not df.empty and df.dropna(how='all').shape[1] > 0 and df.dropna(how='all').shape[0] > 0]
    if not filtered_dfs:
        return dbc.Alert("查無有效資料", color="danger", className="mt-3")

    combined_df = pd.concat(filtered_dfs).sort_values(["股票代碼", "季度"])

    # 格式化數字欄位
    num_cols = combined_df.columns.drop(["股票代碼", "季度", "date"], errors='ignore')
    for col in num_cols:
        combined_df[col] = combined_df[col].map(lambda x: f"{float(x):.2f}" if pd.notnull(x) and x != "" else "無資料")

    indicator_groups = {
        "獲利能力指標": [
            "每股盈餘(EPS)", "營業毛利率(GrossProfitMargin)", "營業利益率(OperatingMargin)",
            "稅前淨利率(PreTaxProfitMargin)", "稅後淨利率(NetProfitMargin)",
            "資產報酬率(ROA)", "股東權益報酬率(ROE)"
        ],
        "償債能力指標": [
            "現金比率(CashRatio)", "流動比率(CurrentRatio)",
            "利息保障倍數(InterestCoverageRatio)", "現金流量比(OperatingCashFlowRatio)"
        ],
        "獲利年成長率指標": [
            "營收年成長率(Revenue YoY)", "毛利年成長率(GrossProfit YoY)",
            "營業利益年成長率(OperatingIncome YoY)", "稅前淨利年成長率(PreTaxIncome YoY)",
            "稅後淨利年成長率(IncomeAfterTaxes YoY)", "每股盈餘年成長率(EPS YoY)"
        ],
        "經營能力指標": [
            "營業成本率(CostMargin)", "營業費用率(ExpenseMargin)",
            "存貨週轉率(InventoryTurnover)", "平均售貨日數(DaysSalesOutstanding)",
            "總資產週轉率(TotalAssetTurnover)"
        ]
    }

    for sid, group_df in combined_df.groupby("股票代碼"):
        content.append(html.H4(f"股票代碼：{sid}", className="mt-4 text-primary"))

        for section_title, columns in indicator_groups.items():
            available_cols = [c for c in columns if c in group_df.columns]
            if not available_cols:
                continue

            cols_order = ["季度"] + available_cols

            # 先計算欄寬（固定總寬度除以欄數）
            col_width = f"{round(1000 / len(cols_order))}px"

            # 表頭
            table_header = [
                html.Th(col, style={
                    'fontWeight': 'bold',
                    'minWidth': '120px',
                    'maxWidth': '200px',
                    'textAlign': 'center',
                    'whiteSpace': 'nowrap',         # ❗不換行
                    'overflow': 'hidden',
                    'textOverflow': 'ellipsis',
                    'fontSize': '14px'              # ❗縮小字體
                }) for col in cols_order
            ]

            table_rows = [
                html.Tr([
                    html.Td(group_df.iloc[i][col], style={
                        'minWidth': '120px',
                        'maxWidth': '200px',
                        'textAlign': 'center',
                        'whiteSpace': 'nowrap',     # ❗不換行
                        'overflow': 'hidden',
                        'textOverflow': 'ellipsis',
                        'fontSize': '14px'          # ❗縮小字體
                    }) for col in cols_order
                ]) for i in range(len(group_df))
            ]

            table = dbc.Table(
                [html.Thead(html.Tr(table_header)), html.Tbody(table_rows)],
                bordered=True,
                striped=True,
                hover=True,
                responsive=True,
                className="mb-4",
                style={"tableLayout": "fixed", "width": "100%"}
            )

            card = dbc.Card([
                dbc.CardHeader(html.H5(section_title)),
                dbc.CardBody(
                    html.Div(table, style={"overflowX": "auto"}),  # 💡 表格可橫向滾動
                )
            ], className="mb-4")

            content.append(card)


    # 繪製趨勢圖
    charts = []
    for indicator in selected_indicators:
        temp_df = combined_df[["股票代碼", "季度", indicator]].copy()
        temp_df[indicator] = pd.to_numeric(temp_df[indicator], errors="coerce")

        no_data_stocks_for_indicator = [
            sid for sid in stock_ids if temp_df[temp_df["股票代碼"] == sid][indicator].dropna().empty
        ]

        fig = px.line(
            temp_df.dropna(subset=[indicator]),
            x="季度",
            y=indicator,
            color="股票代碼",
            markers=True,
            title=f"{indicator} 趨勢圖"
        )
        fig.update_layout(legend_title_text="股票代碼")

        if temp_df[indicator].dropna().empty:
            fig = px.line(title=f"{indicator} 趨勢圖 - 無資料")

        if no_data_stocks_for_indicator:
            msg = "，".join(no_data_stocks_for_indicator)
            charts.append(
                dbc.Alert(f"{indicator} 無資料股票：{msg}", color="warning", className="mb-2")
            )
        charts.append(dcc.Graph(figure=fig))

    content.extend(charts)

    return content

import io
import xlsxwriter
import plotly.graph_objects as go
import plotly.io as pio
import base64

@app.callback(
    Output("download-excel", "data"),
    Input("download-btn", "n_clicks"),
    State("stock-input", "value"),
    State("start-quarter", "value"),
    State("end-quarter", "value"),
    State("selected-profitability-indicators", "value"),
    State("selected-solvency-indicators", "value"),
    State("selected-growth-indicators", "value"),
    State("selected-efficiency-indicators", "value"),
    prevent_initial_call=True
)
def generate_excel(n_clicks, stock_input, start_q, end_q, selected_profit, selected_solv, selected_growth, selected_eff):
    import pandas as pd
    import io
    import plotly.graph_objects as go
    import plotly.io as pio
    import xlsxwriter

    stock_ids = [s.strip() for s in stock_input.split(",")]
    quarters = generate_quarter_range(start_q, end_q)

    all_indicators = {
        "獲利能力指標": [
            "每股盈餘(EPS)", "營業毛利率(GrossProfitMargin)", "營業利益率(OperatingMargin)",
            "稅前淨利率(PreTaxProfitMargin)", "稅後淨利率(NetProfitMargin)",
            "資產報酬率(ROA)", "股東權益報酬率(ROE)"
        ],
        "償債能力指標": [
            "現金比率(CashRatio)", "流動比率(CurrentRatio)",
            "利息保障倍數(InterestCoverageRatio)", "現金流量比(OperatingCashFlowRatio)"
        ],
        "獲利年成長率指標": [
            "營收年成長率(Revenue YoY)", "毛利年成長率(GrossProfit YoY)",
            "營業利益年成長率(OperatingIncome YoY)", "稅前淨利年成長率(PreTaxIncome YoY)",
            "稅後淨利年成長率(IncomeAfterTaxes YoY)", "每股盈餘年成長率(EPS YoY)"
        ],
        "經營能力指標": [
            "營業成本率(CostMargin)", "營業費用率(ExpenseMargin)",
            "存貨週轉率(InventoryTurnover)", "平均售貨日數(DaysSalesOutstanding)",
            "總資產週轉率(TotalAssetTurnover)"
        ]
    }

    selected_indicators = selected_profit + selected_solv + selected_growth + selected_eff

    # 擷取資料（你自己寫的get_financial_indicators和FINMIND_TOKEN）
    combined_df = pd.DataFrame()
    for stock_id in stock_ids:
        df = get_financial_indicators(stock_id, quarters, FINMIND_TOKEN)
        if not df.empty:
            combined_df = pd.concat([combined_df, df], ignore_index=True)

    if combined_df.empty:
        return dcc.send_string("查無資料，請確認股票代碼與季度區間是否正確。", filename="查無資料.txt")

    # 新增Excel writer跟workbook
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book     
    used_sheet_names = set()

    # --- 第一頁：財報資料，分股票及指標類別整理 ---
    sheet_name = "財報資料"
    used_sheet_names.add(sheet_name)
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Excel 起始列（0-based）
    start_row = 0
    start_col = 0

    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    stock_format = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1})
    indicator_format = workbook.add_format({'italic': True, 'bg_color': '#E2EFDA', 'border': 1})
    cell_format = workbook.add_format({'border': 1})

    for stock_id in stock_ids:
        # 股票名稱標題列
        worksheet.merge_range(start_row, start_col, start_row, start_col + 10, f"股票代碼：{stock_id}", stock_format)
        start_row += 1

        # 先寫入「季度」欄標題
        worksheet.write(start_row, start_col, "季度", header_format)

        col_idx = start_col + 1
        # 寫入所有指標類別的指標名稱（欄位標題）
        for category, indicators in all_indicators.items():
            # 過濾掉不在 combined_df 或不屬於用戶選擇的指標（選擇性更強）
            filtered_inds = [ind for ind in indicators if ind in combined_df.columns and (ind in selected_indicators or True)]
            if not filtered_inds:
                continue

            # 先寫指標類別合併欄位
            merge_start = col_idx
            merge_end = col_idx + len(filtered_inds) - 1
            worksheet.merge_range(start_row, merge_start, start_row, merge_end, category, indicator_format)

            # 下一列寫指標名稱
            for i, ind in enumerate(filtered_inds):
                worksheet.write(start_row + 1, col_idx + i, ind, header_format)

            col_idx += len(filtered_inds)

        # 寫入股票的資料列，從第三列開始（季度與指標資料）
        start_row += 2

        # 篩選該股票的資料
        df_stock = combined_df[combined_df["股票代碼"] == stock_id].copy()
        df_stock = df_stock.sort_values("季度")

        for row_i, (_, row) in enumerate(df_stock.iterrows()):
            worksheet.write(start_row + row_i, start_col, row["季度"], cell_format)
            col_idx = start_col + 1

            for category, indicators in all_indicators.items():
                filtered_inds = [ind for ind in indicators if ind in combined_df.columns and (ind in selected_indicators or True)]
                for ind in filtered_inds:
                    val = row.get(ind, None)
                    worksheet.write(start_row + row_i, col_idx, val, cell_format)
                    col_idx += 1

        # 跟下一個股票區塊空一行
        start_row += len(df_stock) + 2

    # --- 各指標分頁：繪圖與插入圖片 ---
    for indicator in selected_indicators:
        if indicator not in combined_df.columns:
            continue

        base_name = indicator[:28]
        sheet_name = base_name
        i = 1
        while sheet_name in used_sheet_names:
            sheet_name = f"{base_name}_{i}"
            i += 1
        used_sheet_names.add(sheet_name)

        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet

        fig = go.Figure()
        for stock_id in stock_ids:
            df_plot = combined_df[combined_df["股票代碼"] == stock_id]
            if indicator not in df_plot.columns:
                continue
            df_plot = df_plot.sort_values("季度")
            fig.add_trace(go.Scatter(
                x=df_plot["季度"],
                y=df_plot[indicator],
                mode="lines+markers",
                name=stock_id
            ))
        fig.update_layout(title=indicator, xaxis_title="季度", yaxis_title=indicator)

        img_bytes = io.BytesIO()
        try:
            pio.write_image(fig, img_bytes, format="png", width=800, height=500)
            img_bytes.seek(0)
            worksheet.insert_image("B2", f"{indicator}.png", {"image_data": img_bytes})
        except Exception as e:
            worksheet.write("A1", f"圖表生成失敗：{str(e)}")

    writer.close()
    output.seek(0)
    return dcc.send_bytes(output.read(), filename="財報報表.xlsx")



if __name__ == "__main__":
    app.run_server(debug=True)

