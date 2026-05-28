import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="SKU 价格毛利分析", layout="wide")

# ── 页眉：Logo + 标题 + 制作人信息 ───────────────────────────────────────────
header_left, header_mid, header_right = st.columns([1, 3, 2])

with header_left:
    try:
        st.image("carzone_logo.png", width=110)
    except Exception:
        pass

with header_mid:
    st.markdown("## SKU 价格毛利分析")
    st.caption("底盘件价格体系优化 · 核算价 × 毛利率模拟工具")

with header_right:
    st.markdown(
        """
        <div style='text-align:right; line-height:1.7; font-size:13px; color:#888; padding-top:8px'>
            <b>南京新康众 · 供应链底盘组</b><br>
            制作人：李宇凡<br>
            <span style='color:#e74c3c; font-size:12px'>⚠ 仅供底盘组内部使用，注意保护敏感数据安全</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ── 数据加载 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("数据源")
    uploaded = st.file_uploader("上传 SKU 价格导出文件（.xlsx）", type=["xlsx"])

@st.cache_data
def load_data(source):
    df = pd.read_excel(source)
    cols = list(df.columns)
    if len(cols) < 11:
        raise ValueError(f"列数不足：该文件只有 {len(cols)} 列，需要至少 11 列的 SKU 价格导出格式。")
    renamed = df.rename(columns={
        cols[1]:  "SKU编码",
        cols[4]:  "品类",
        cols[5]:  "粗称",
        cols[6]:  "品牌",
        cols[9]:  "采购成本",
        cols[10]: "核算价",
    })[["SKU编码", "品类", "粗称", "品牌", "采购成本", "核算价"]]
    renamed["采购成本"] = pd.to_numeric(renamed["采购成本"], errors="coerce")
    renamed["核算价"]   = pd.to_numeric(renamed["核算价"],   errors="coerce")
    return renamed

if uploaded:
    try:
        raw = load_data(uploaded)
    except Exception as e:
        st.error(f"❌ 文件格式不兼容：{e}")
        st.info("请上传系统导出的「SKU价格导出」标准格式文件，法雷奥/菲罗多销售报表等其他格式暂不支持。")
        st.stop()
else:
    st.info("👈 请在左侧上传 SKU 价格导出文件（.xlsx）开始分析")
    st.stop()

df = raw[raw["采购成本"].notna() & raw["核算价"].notna()].copy()
df = df[(df["采购成本"] > 0) & (df["核算价"] > 0)].reset_index(drop=True)
df["当前毛利率"] = (df["核算价"] - df["采购成本"]) / df["核算价"]

# ── 侧边栏 ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.header("筛选")
    brands = ["全部"] + sorted(df["品牌"].dropna().unique().tolist())
    cats   = ["全部"] + sorted(df["品类"].dropna().unique().tolist())
    sel_brand = st.selectbox("品牌", brands)
    sel_cat   = st.selectbox("品类", cats)

    st.divider()
    st.header("调价参数")
    adj_pct   = st.slider("调价幅度（%）", -30, 30, 0, 1, help="正数=涨价，负数=降价")
    threshold = st.slider("毛利率基准线（%）", 5, 25, 10, 1)

# ── 筛选 + 计算 ───────────────────────────────────────────────────────────────
view = df.copy()
if sel_brand != "全部":
    view = view[view["品牌"] == sel_brand]
if sel_cat != "全部":
    view = view[view["品类"] == sel_cat]

view = view.copy()
view["调整后核算价"] = view["核算价"] * (1 + adj_pct / 100)
view["调整后毛利率"] = (view["调整后核算价"] - view["采购成本"]) / view["调整后核算价"]
view["单位利润变化"]  = view["调整后核算价"] - view["核算价"]

thr        = threshold / 100
n_total    = len(view)
n_red      = int((view["调整后毛利率"] < 0.08).sum())
n_yellow   = int(((view["调整后毛利率"] >= 0.08) & (view["调整后毛利率"] < thr)).sum())
n_green    = int((view["调整后毛利率"] >= thr).sum())
n_below    = n_red + n_yellow
avg_before = view["当前毛利率"].mean()
avg_after  = view["调整后毛利率"].mean()

# ── 概览指标卡 ────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("SKU 总数", f"{n_total:,}")
c2.metric("低于基准线", f"{n_below:,}",
          f"{n_below/n_total*100:.1f}%" if n_total else "—",
          delta_color="inverse")
c3.metric("当前均值毛利率", f"{avg_before:.1%}")
c4.metric("调整后均值毛利率", f"{avg_after:.1%}",
          f"{avg_after - avg_before:+.1%}",
          delta_color="normal")

st.divider()

# ── 可视化卡片区 ──────────────────────────────────────────────────────────────
v1, v2, v3 = st.columns([1, 1.6, 1.4])

# 卡片1：SKU 健康状态环形图
with v1:
    st.markdown("#### SKU 健康分布")
    donut = go.Figure(go.Pie(
        labels=["健康（≥基准线）", f"预警（8%~{threshold}%）", "危险（<8%）"],
        values=[n_green, n_yellow, n_red],
        hole=0.62,
        marker_colors=["#2ecc71", "#f39c12", "#e74c3c"],
        textinfo="percent+value",
        hovertemplate="%{label}<br>%{value} 个 SKU<br>占比 %{percent}<extra></extra>",
    ))
    donut.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        height=260,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, font_size=11),
        annotations=[dict(
            text=f"<b>{n_green}</b><br>健康",
            x=0.5, y=0.5, font_size=15, showarrow=False
        )],
    )
    st.plotly_chart(donut, use_container_width=True, config={"displayModeBar": False})

# 卡片2：毛利率分布直方图（调价前 vs 调价后叠加）
with v2:
    st.markdown("#### 毛利率分布（调价前 vs 调价后）")
    hist = go.Figure()
    hist.add_trace(go.Histogram(
        x=view["当前毛利率"] * 100,
        name="调价前",
        nbinsx=40,
        marker_color="rgba(52,152,219,0.55)",
        hovertemplate="毛利率 %{x:.1f}%<br>SKU 数 %{y}<extra>调价前</extra>",
    ))
    hist.add_trace(go.Histogram(
        x=view["调整后毛利率"] * 100,
        name="调价后",
        nbinsx=40,
        marker_color="rgba(46,204,113,0.55)",
        hovertemplate="毛利率 %{x:.1f}%<br>SKU 数 %{y}<extra>调价后</extra>",
    ))
    hist.add_vline(x=threshold, line_dash="dash", line_color="#e74c3c", line_width=1.5,
                   annotation_text=f"基准线 {threshold}%", annotation_position="top right",
                   annotation_font_color="#e74c3c")
    hist.update_layout(
        barmode="overlay",
        margin=dict(t=10, b=30, l=40, r=10),
        height=260,
        xaxis_title="毛利率（%）",
        yaxis_title="SKU 数",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(hist, use_container_width=True, config={"displayModeBar": False})

# 卡片3：品类调价前后毛利率对比柱状图
with v3:
    st.markdown("#### 品类毛利率对比")
    cat_summary = (
        view.groupby("品类")
        .agg(调价前=("当前毛利率", "mean"), 调价后=("调整后毛利率", "mean"))
        .reset_index()
        .sort_values("调价后", ascending=True)
    )
    bar = go.Figure()
    bar.add_trace(go.Bar(
        y=cat_summary["品类"],
        x=cat_summary["调价前"] * 100,
        name="调价前",
        orientation="h",
        marker_color="rgba(52,152,219,0.7)",
        hovertemplate="%{y}<br>调价前 %{x:.1f}%<extra></extra>",
    ))
    bar.add_trace(go.Bar(
        y=cat_summary["品类"],
        x=cat_summary["调价后"] * 100,
        name="调价后",
        orientation="h",
        marker_color="rgba(46,204,113,0.7)",
        hovertemplate="%{y}<br>调价后 %{x:.1f}%<extra></extra>",
    ))
    bar.add_vline(x=threshold, line_dash="dash", line_color="#e74c3c", line_width=1.5)
    bar.update_layout(
        barmode="group",
        margin=dict(t=10, b=30, l=10, r=10),
        height=260,
        xaxis_title="均值毛利率（%）",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(bar, use_container_width=True, config={"displayModeBar": False})

st.divider()

# ── 品类汇总表 ────────────────────────────────────────────────────────────────
st.subheader("品类汇总")
summary = (
    view.groupby("品类", sort=False)
    .agg(
        SKU数=("SKU编码", "count"),
        均值成本=("采购成本", "mean"),
        均值核算价=("核算价", "mean"),
        当前均值毛利率=("当前毛利率", "mean"),
        调整后均值毛利率=("调整后毛利率", "mean"),
        低于基准=("调整后毛利率", lambda x: (x < thr).sum()),
    )
    .reset_index()
)

def style_margin(val):
    if isinstance(val, float):
        if val < 0.08:
            return "color: #c0392b; font-weight:bold"
        elif val < thr:
            return "color: #e67e22; font-weight:bold"
        else:
            return "color: #27ae60"
    return ""

fmt = {"均值成本": "{:.2f}", "均值核算价": "{:.2f}",
       "当前均值毛利率": "{:.1%}", "调整后均值毛利率": "{:.1%}"}
st.dataframe(
    summary.style.map(style_margin, subset=["当前均值毛利率", "调整后均值毛利率"]).format(fmt),
    use_container_width=True, hide_index=True
)

# ── SKU 明细 ─────────────────────────────────────────────────────────────────
st.subheader("SKU 明细")

col_filter, col_page = st.columns([2, 1])
with col_filter:
    show_only_below = st.checkbox("只显示低于基准线的 SKU", value=False)
with col_page:
    page_size = st.selectbox("每页显示", [50, 100, 200], index=1)

detail = view if not show_only_below else view[view["调整后毛利率"] < thr]
detail = detail[["SKU编码", "品类", "品牌", "采购成本", "核算价", "当前毛利率",
                  "调整后核算价", "调整后毛利率", "单位利润变化"]].reset_index(drop=True)

total_pages = max(1, (len(detail) - 1) // page_size + 1)
page = st.number_input(f"页码（共 {total_pages} 页，{len(detail)} 条）",
                       min_value=1, max_value=total_pages, value=1, step=1) - 1
page_df = detail.iloc[page * page_size : (page + 1) * page_size].copy()

fmt2 = {"采购成本": "{:.2f}", "核算价": "{:.2f}", "当前毛利率": "{:.1%}",
        "调整后核算价": "{:.2f}", "调整后毛利率": "{:.1%}", "单位利润变化": "{:+.2f}"}

def row_color(row):
    m = row["调整后毛利率"]
    bg = "#fff0f0" if m < 0.08 else ("#fffbea" if m < thr else "#f0fff4")
    return [f"background-color: {bg}"] * len(row)

st.dataframe(
    page_df.style
    .apply(row_color, axis=1)
    .map(style_margin, subset=["当前毛利率", "调整后毛利率"])
    .format(fmt2),
    use_container_width=True, height=420, hide_index=True
)

# ── 导出 ─────────────────────────────────────────────────────────────────────
st.divider()
export_df = detail[["SKU编码", "品类", "品牌", "采购成本", "核算价",
                     "当前毛利率", "调整后核算价", "调整后毛利率", "单位利润变化"]].copy()
export_df["当前毛利率"]   = export_df["当前毛利率"].map("{:.1%}".format)
export_df["调整后毛利率"] = export_df["调整后毛利率"].map("{:.1%}".format)

st.download_button(
    "导出当前视图为 Excel",
    data=export_df.to_csv(index=False, encoding="utf-8-sig").encode(),
    file_name=f"价格毛利分析_调价{adj_pct:+d}pct.csv",
    mime="text/csv",
)
