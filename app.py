# Deploy with: streamlit run app.py
import io
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from streamlit_option_menu import option_menu

APP_TITLE = "服務費管理系統"

st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")

# -----------------------------
# Style
# -----------------------------
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;700&display=swap');

        html, body, [data-testid="stAppViewContainer"] {
            font-family: 'Inter', 'Noto Sans TC', sans-serif;
            background-color: #f8fafc;
        }

        .main { padding-top: 0; }

        [data-testid="stSidebar"] { background-color: #1e293b !important; }
        [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] p, [data-testid="stSidebar"] label {
            color: #f8fafc !important;
        }

        .content-card {
            background: #ffffff;
            border-radius: 14px;
            padding: 24px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            margin-bottom: 20px;
        }

        .card-title {
            font-size: 18px;
            font-weight: 700;
            color: #334155;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .card-title::before {
            content: '';
            width: 4px;
            height: 16px;
            background: #3b82f6;
            border-radius: 2px;
        }

        .kpi-card {
            background: #ffffff;
            border-radius: 14px;
            padding: 18px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            height: 100%;
        }

        .kpi-total { border-top: 4px solid #3b82f6; }
        .kpi-icon { font-size: 22px; margin-bottom: 10px; }
        .kpi-label { font-size: 14px; color: #64748b; font-weight: 600; margin-bottom: 4px; }
        .kpi-value { font-size: 28px; font-weight: 700; color: #0f172a; }
        .kpi-sub { font-size: 13px; color: #94a3b8; margin-top: 4px; }

        .detail-section {
            background: #ffffff;
            border-radius: 14px;
            padding: 20px;
            border: 1px solid #e2e8f0;
            height: 100%;
        }

        .detail-header {
            font-size: 14px;
            color: #94a3b8;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 14px;
            padding-bottom: 8px;
            border-bottom: 2px solid #f1f5f9;
        }

        .detail-row {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            padding: 10px 0;
            border-bottom: 1px solid #f8fafc;
        }

        .detail-label { color: #475569; }
        .detail-value { color: #1e293b; font-weight: 600; }

        .summary-box {
            background: #f8fafc;
            border-radius: 10px;
            padding: 16px;
            margin-top: 12px;
        }

        .summary-total {
            font-size: 28px;
            font-weight: 700;
            color: #3b82f6;
        }

        .stDataFrame td { text-align: center; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Helpers
# -----------------------------
NUMERIC_COLS = [
    "遠通臨停佣金",
    "遠通月租佣金",
    "遠通中獎通知",
    "APS繳費機(非現金)",
    "APS繳費機(現金)",
    "APS繳費機(非代收)",
    "uTagGO易付",
    "uTagGO停車",
    "遠創中獎通知",
    "充電費",
    "月租代收應收服務費",
    "臨停應收佣金",
    "月租應收佣金",
]


def clean_numeric(x):
    if pd.isna(x):
        return 0
    if isinstance(x, str):
        return pd.to_numeric(x.replace(",", "").strip(), errors="coerce")
    return x


def get_quarter(month_str):
    try:
        m = int(str(month_str).strip())
        if m <= 3:
            return "Q1"
        if m <= 6:
            return "Q2"
        if m <= 9:
            return "Q3"
        return "Q4"
    except Exception:
        return "N/A"


def money(x):
    try:
        if pd.isna(x):
            return "$0"
        return f"${float(x):,.0f}"
    except Exception:
        return "$0"


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_excel_any(file_bytes: bytes) -> pd.DataFrame:
    # Try all sheets and pick the first sheet that contains data rows.
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    last_df = None
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
            df = normalize_columns(df)
            # drop fully empty rows
            df = df.dropna(how="all")
            last_df = df
            if not df.empty:
                return df
        except Exception:
            continue
    if last_df is not None:
        return last_df
    return pd.DataFrame()


def load_data_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    df = read_excel_any(file_bytes)
    return prepare_dataframe(df)


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    df = df.dropna(how="all").copy()

    if df.empty:
        return df

    # 數字欄位清理
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = df[col].apply(clean_numeric).fillna(0)

    # 必要欄位
    required = ["年月", "業者"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位：{', '.join(missing)}")

    # 有些檔案把資料當作第一列文字時，先排除純表頭列
    # （如果年月只有單一值但業者也存在，仍可處理）
    df["年月"] = df["年月"].astype(str).str.strip()
    df["業者"] = df["業者"].astype(str).str.strip()

    # 可能的欄位別名
    yt_parking = df["臨停應收佣金"] if "臨停應收佣金" in df.columns else df.get("遠通臨停佣金", 0)
    yt_monthly = df["月租應收佣金"] if "月租應收佣金" in df.columns else df.get("遠通月租佣金", 0)

    df["遠通臨停"] = yt_parking.fillna(0) + df.get("遠通中獎通知", 0).fillna(0)
    df["遠通月租"] = yt_monthly.fillna(0)

    fechuang_parking_cols = [
        "uTagGO易付",
        "uTagGO停車",
        "遠創中獎通知",
        "充電費",
        "APS繳費機(非現金)",
        "APS繳費機(現金)",
        "APS繳費機(非代收)",
    ]
    for col in fechuang_parking_cols + ["月租代收應收服務費"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = df[col].apply(clean_numeric).fillna(0)

    df["遠創臨停"] = df[fechuang_parking_cols].sum(axis=1)
    df["遠創月租"] = df["月租代收應收服務費"]

    # 若 Excel 已有總服務費就保留，否則重算
    if "總服務費" in df.columns:
        df["總服務費"] = df["總服務費"].apply(clean_numeric).fillna(0)
        if df["總服務費"].eq(0).all():
            df["總服務費"] = df["遠通臨停"] + df["遠通月租"] + df["遠創臨停"] + df["遠創月租"]
    else:
        df["總服務費"] = df["遠通臨停"] + df["遠通月租"] + df["遠創臨停"] + df["遠創月租"]

    # 年月欄位推導
    ym = df["年月"].astype(str).str.replace("-", "/", regex=False).str.split("/", n=1, expand=True)
    df["年份"] = ym[0].fillna("").astype(str).str.strip()
    df["月份"] = ym[1].fillna("").astype(str).str.strip() if ym.shape[1] > 1 else ""
    if not isinstance(df["月份"], pd.Series):
        df["月份"] = ""
    df["季度"] = df["月份"].apply(get_quarter)
    df["年季"] = df["年份"] + " " + df["季度"]

    # 名稱修正
    df["業者"] = (
        df["業者"]
        .astype(str)
        .str.strip()
        .str.replace("?亭", "俥亭", regex=False)
        .str.replace("？亭", "俥亭", regex=False)
        .str.replace("俥亭停車", "俥亭", regex=False)
    return df

def find_local_workbook() -> Optional[Path]:
    candidates = [
        Path("0505_merged.xlsx"),
        Path("0505.xlsx"),
        Path(__file__).with_name("0505_merged.xlsx"),
        Path(__file__).with_name("0505.xlsx"),
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def get_df_from_sources(uploaded_file) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    # 優先使用上傳檔案；沒有上傳才找本機/Repo 內 Excel
    if uploaded_file is not None:
        try:
            df = load_data_from_bytes(uploaded_file.getvalue())
            return df, uploaded_file.name
        except Exception as e:
            return None, f"讀取上傳 Excel 失敗：{e}"

    local = find_local_workbook()
    if local is not None:
        try:
            df = load_data_from_bytes(local.read_bytes())
            return df, local.name
        except Exception as e:
            return None, f"讀取本機 Excel 失敗：{e}"

    return None, None


# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.markdown("<h2 style='color:white; font-size:1.05rem; margin-bottom:1.25rem;'>📊 服務費管理系統</h2>", unsafe_allow_html=True)

    selected = option_menu(
        menu_title=None,
        options=["儀表板", "業者查詢", "資料匯入"],
        icons=["grid-1x2", "search", "cloud-arrow-up"],
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": "transparent"},
            "icon": {"color": "#3b82f6", "font-size": "1.2rem"},
            "nav-link": {
                "color": "#94a3b8",
                "font-size": "1rem",
                "text-align": "left",
                "margin": "0.2rem 0",
                "--hover-color": "#334155",
            },
            "nav-link-selected": {"background-color": "#3b82f6", "color": "#fff"},
        },
    )

    st.markdown("<hr style='border-color: #334155;'>", unsafe_allow_html=True)
    time_unit = st.radio("時間維度", ["按月", "按季", "按年"], index=0)

    st.markdown("---")
    st.caption("資料來源")

    st.markdown("### 上傳 Excel")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=["xlsx", "xls"], label_visibility="collapsed")

    if st.button("重新讀取資料", use_container_width=True):
        st.rerun()

# -----------------------------
# Load data
# -----------------------------
df, data_msg = get_df_from_sources(uploaded_file)

if df is None:
    st.title(APP_TITLE)
    if data_msg:
        st.error(data_msg)
    else:
        st.info("請在左側上傳 Excel 檔案。")
    st.stop()

st.sidebar.write(f"目前資料：{uploaded_file.name if uploaded_file is not None else (find_local_workbook().name if find_local_workbook() else '尚未載入')}")

if df.empty:
    st.title(APP_TITLE)
    st.warning("這個 Excel 只有欄位標題，沒有資料列，所以無法顯示儀表板。請上傳有資料的版本。")
    st.stop()

# -----------------------------
# Filter logic
# -----------------------------
if time_unit == "按月":
    time_options = sorted([x for x in df["年月"].dropna().astype(str).unique().tolist() if x and x.lower() != "nan"], reverse=True)
    selected_time = st.sidebar.selectbox("選擇時間範圍", ["全部"] + list(time_options)) if time_options else "全部"
    trend_col = "年月"
elif time_unit == "按季":
    time_options = sorted([x for x in df["年季"].dropna().astype(str).unique().tolist() if x and x.lower() != "nan"], reverse=True)
    selected_time = st.sidebar.selectbox("選擇時間範圍", ["全部"] + list(time_options)) if time_options else "全部"
    trend_col = "年季"
else:
    time_options = sorted([x for x in df["年份"].dropna().astype(str).unique().tolist() if x and x.lower() != "nan"], reverse=True)
    selected_time = st.sidebar.selectbox("選擇時間範圍", ["全部"] + list(time_options)) if time_options else "全部"
    trend_col = "年份"

filtered_df = df.copy()
if selected_time != "全部":
    filtered_df = filtered_df[filtered_df[trend_col].astype(str) == str(selected_time)]

# -----------------------------
# Pages
# -----------------------------
if selected == "儀表板":
    st.markdown("<h2 style='color:#0f172a; font-size:2rem; margin-bottom:0.2rem;'>儀表板</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='color:#64748b; font-size:1rem; margin-bottom:1.2rem;'>時間範圍：{selected_time if selected_time != '全部' else '歷史全紀錄'}</p>",
        unsafe_allow_html=True,
    )

    if filtered_df.empty:
        st.warning("目前篩選條件下沒有資料。")
        st.stop()

    k1, k2, k3, k4, k5 = st.columns(5)
    sums = filtered_df[["遠通臨停", "遠通月租", "遠創臨停", "遠創月租", "總服務費"]].sum(numeric_only=True)
    v1, v2, v3, v4, v_total = [float(sums.get(c, 0) or 0) for c in ["遠通臨停", "遠通月租", "遠創臨停", "遠創月租", "總服務費"]]

    with k1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon">🚗</div><div class="kpi-label">遠通臨停服務費</div><div class="kpi-value">{money(v1)}</div><div class="kpi-sub">臨停應收 + 遠通中獎</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon">📅</div><div class="kpi-label">遠通月租服務費</div><div class="kpi-value">{money(v2)}</div><div class="kpi-sub">月租應收佣金</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon">📱</div><div class="kpi-label">遠創臨停服務費</div><div class="kpi-value">{money(v3)}</div><div class="kpi-sub">uTagGO + APS + 充電</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon">🏠</div><div class="kpi-label">遠創月租服務費</div><div class="kpi-value">{money(v4)}</div><div class="kpi-sub">月租代收應收服務費</div></div>', unsafe_allow_html=True)
    with k5:
        st.markdown(f'<div class="kpi-card kpi-total"><div class="kpi-icon">💰</div><div class="kpi-label">總服務費</div><div class="kpi-value">{money(v_total)}</div><div class="kpi-sub">以上四項加總</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-top:2rem;'></div>", unsafe_allow_html=True)

    c1, c2 = st.columns([2, 1])
    with c1:
        st.markdown('<div class="content-card"><div class="card-title">歷史趨勢分析 (遠通 vs 遠創)</div>', unsafe_allow_html=True)
        t_df = (
            df.groupby(trend_col, dropna=False)
            .agg({"遠通臨停": "sum", "遠通月租": "sum", "遠創臨停": "sum", "遠創月租": "sum"})
            .reset_index()
            .sort_values(trend_col)
        )
        t_df["遠通合計"] = t_df["遠通臨停"] + t_df["遠通月租"]
        t_df["遠創合計"] = t_df["遠創臨停"] + t_df["遠創月租"]

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(x=t_df[trend_col], y=t_df["遠通合計"], name="遠通合計", line=dict(width=3), fill="tozeroy", fillcolor="rgba(59,130,246,0.05)"))
        fig_trend.add_trace(go.Scatter(x=t_df[trend_col], y=t_df["遠創合計"], name="遠創合計", line=dict(width=3), fill="tozeroy", fillcolor="rgba(16,185,129,0.05)"))
        fig_trend.update_layout(height=320, margin=dict(l=0, r=0, t=0, b=0), hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_trend, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="content-card"><div class="card-title">服務費組成佔比</div>', unsafe_allow_html=True)
        pie_df = pd.DataFrame({"類別": ["遠通臨停", "遠通月租", "遠創臨停", "遠創月租"], "金額": [v1, v2, v3, v4]})
        fig_pie = px.pie(pie_df, values="金額", names="類別", hole=0.6)
        fig_pie.update_layout(height=320, margin=dict(l=0, r=0, t=0, b=0), showlegend=True)
        st.plotly_chart(fig_pie, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="content-card"><div class="card-title">前 10 大貢獻業者排行榜</div>', unsafe_allow_html=True)
    top_sites = filtered_df.groupby("業者", dropna=False)["總服務費"].sum().reset_index().sort_values("總服務費", ascending=False).head(10)
    top_sites.insert(0, "排名", range(1, len(top_sites) + 1))
    top_sites["佔比"] = (top_sites["總服務費"] / v_total * 100).round(1).astype(str) + "%" if v_total else "0.0%"
    st.dataframe(top_sites.style.format({"總服務費": "{:,.0f}"}), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

elif selected == "業者查詢":
    st.markdown("<h2 style='color:#0f172a; font-size:2rem; margin-bottom:0.2rem;'>業者明細查詢</h2>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#64748b; font-size:1rem; margin-bottom:1.2rem;'>時間範圍：{selected_time if selected_time != '全部' else '歷史全紀錄'}</p>", unsafe_allow_html=True)

    search_query = st.text_input("🔍 輸入關鍵字搜尋業者 (例如：俥亭、萊爾富)", "", help="輸入業者名稱關鍵字進行搜尋")

    all_sites = sorted((filtered_df if selected_time != "全部" else df)["業者"].dropna().astype(str).unique().tolist())
    if search_query:
        all_sites = [s for s in all_sites if search_query.lower() in s.lower()]

    if not all_sites:
        st.warning("找不到符合關鍵字的業者。")
    else:
        target_site = st.selectbox("請選擇業者名稱", all_sites, help="從列表中選擇一個業者以查看詳細資訊")
        site_data = filtered_df[filtered_df["業者"].astype(str) == str(target_site)].sum(numeric_only=True)

        st.markdown(f"<h3 style='font-size:1.5rem;'>{target_site}</h3>", unsafe_allow_html=True)
        col2, col3 = st.columns(2)

        with col2:
            st.markdown(
                f"""
                <div class="detail-section">
                    <div class="detail-header">遠創系列</div>
                    <div class="detail-row"><span class="detail-label">APS 繳費機 (非現金)</span><span class="detail-value">{money(site_data.get('APS繳費機(非現金)', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">APS 繳費機 (現金)</span><span class="detail-value">{money(site_data.get('APS繳費機(現金)', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">APS 繳費機 (非代收)</span><span class="detail-value">{money(site_data.get('APS繳費機(非代收)', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">uTagGO 易付</span><span class="detail-value">{money(site_data.get('uTagGO易付', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">uTagGO 停車</span><span class="detail-value">{money(site_data.get('uTagGO停車', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">遠創中獎通知</span><span class="detail-value">{money(site_data.get('遠創中獎通知', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">充電費</span><span class="detail-value">{money(site_data.get('充電費', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">月租代收應收服務費</span><span class="detail-value">{money(site_data.get('月租代收應收服務費', 0))}</span></div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with col3:
            st.markdown(
                f"""
                <div class="detail-section">
                    <div class="detail-header">分類小計</div>
                    <div class="detail-row"><span class="detail-label">遠通臨停</span><span class="detail-value">{money(site_data.get('遠通臨停', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">遠通月租</span><span class="detail-value">{money(site_data.get('遠通月租', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">遠創臨停</span><span class="detail-value">{money(site_data.get('遠創臨停', 0))}</span></div>
                    <div class="detail-row"><span class="detail-label">遠創月租</span><span class="detail-value">{money(site_data.get('遠創月租', 0))}</span></div>
                    <div class="summary-box">
                        <div style="font-size:12px; color:#64748b; font-weight:600;">總服務費</div>
                        <div class="summary-total">{money(site_data.get('總服務費', 0))}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

else:
    st.markdown("<h2 style='color:#0f172a; font-size:2rem; margin-bottom:0.2rem;'>資料匯入</h2>", unsafe_allow_html=True)
    st.info("這個版本支援直接上傳 Excel，不需要把檔案放在固定路徑。")
    st.write(f"目前資料列數：{len(df):,}")

